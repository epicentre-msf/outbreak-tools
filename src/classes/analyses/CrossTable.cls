VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Create and format cross-tables for analysis"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Create and format cross-tables for analysis")
'@IgnoreModule UnrecognizedAnnotation, IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ICrossTable

'@class CrossTable
'@description
'The CrossTable class builds and formats structured analysis output tables
'on an Excel worksheet. It supports six table scopes: univariate, bivariate,
'time series, spatial, spatio-temporal, and global summary. Each cross-table
'occupies a rectangular region with row categories, column categories, named
'ranges, and visual formatting. Consumers interact through the ICrossTable
'interface, and AnalysisOutput orchestrates the build sequence by calling
'Build followed by Format for each table specification.
'@depends ITableSpecs, ILinelistSpecs, ITranslationObject, ILLVariables, LLVariables, ILLFormat, IHiddenNames, HiddenNames, BetterArray
'@version 1.1 (2026-02-11)


Private Type TCrossTable
    specs As ITableSpecs
    outputSh As Worksheet
    lData As ILinelistSpecs
    endRw As Long
    endCol As Long
    nbCols As Long
    tsHeader As String                           'Header for time series (used in go-to section)
    deFactoNewSect As Boolean                    'Cached IsDeFactoNewSection result
    deFactoNewSectCached As Boolean              'Whether the cache has been computed
End Type

Private this As TCrossTable
Private Const STANDARD_START_COL As Long = 3
Private Const CLASS_NAME As String = "CrossTable"
Private Const NB_ROWS_TIME_SERIES As Long = 56
Private allRangeNames As BetterArray

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a CrossTable instance from table specifications.
'@details
'Factory method that creates and returns a new ICrossTable instance from
'the given table specifications. A cross-table is a structured analysis
'output (univariate, bivariate, time series, spatial, spatio-temporal, or
'global summary) that is built on an output worksheet and contains row
'categories, column categories, named ranges, and formatting. This method
'validates all three required dependencies, then injects them into a
'fresh CrossTable instance via the PredeclaredId/New pattern and returns
'the interface reference.
'@param specs ITableSpecs. An ITableSpecs object representing one row from the analysis setup sheet.
'@param outputWksh Worksheet. The Excel Worksheet where this cross-table will physically write its headers, row labels, column labels, data placeholders, and named ranges.
'@param lData ILinelistSpecs. The linelist specifications object that provides the translation object
'for locale-aware labels and is passed to ITableSpecs.RowCategories / ColumnCategories for
'category resolution at build time.
'@return ICrossTable. A fully initialized ICrossTable ready for Build or individual step calls.
'@throws InvalidArgument When specs is Nothing.
'@throws InvalidArgument When outputWksh is Nothing.
'@throws InvalidArgument When lData is Nothing.
'@export
Public Function Create(ByVal specs As ITableSpecs, _
                       ByVal outputWksh As Worksheet, _
                       ByRef lData As ILinelistSpecs) As ICrossTable

    If specs Is Nothing Then ThrowError InvalidArgument, "Table specifications not defined"
    If outputWksh Is Nothing Then ThrowError InvalidArgument, "Output worksheet not defined"
    If lData Is Nothing Then ThrowError InvalidArgument, "Linelist specifications not defined"

    With New CrossTable
        Set .Specifications = specs
        Set .Wksh = outputWksh
        Set .LinelistData = lData
        Set Create = .Self
    End With
End Function

'@section Internal Properties
'===============================================================================
'@description Properties used during factory instantiation and internal wiring.

'@label:self
'@prop-title Current object instance.
'@details
'Convenience accessor so consuming code can fluently retrieve the interface
'reference from the predeclared Create method.
'@return ICrossTable. The current instance cast to the interface.
Public Property Get Self() As ICrossTable
    Set Self = Me
End Property

'@label:wksh
'@prop-title Output worksheet where the table is built.
'@details
'Returns the Worksheet reference where this cross-table writes its
'headers, row labels, column labels, and data cells.
'@return Worksheet. The output worksheet.
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.outputSh
End Property

'@label:wksh-set
'@prop-title Assign the output worksheet.
'@param sh Worksheet. The target worksheet to assign.
Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.outputSh = sh
End Property

'@label:specifications
'@prop-title Table specifications used to build this cross-table.
'@details
'Returns the ITableSpecs instance that defines the configuration for
'this cross-table. Provides access to table scope, variable names,
'category lists, and validation flags.
'@return ITableSpecs. The backing table specification.
'@export
Public Property Get Specifications() As ITableSpecs
    Set Specifications = this.specs
End Property

'@label:specifications-set
'@prop-title Assign the table specifications.
'@param specs ITableSpecs. The table specification to assign.
Public Property Set Specifications(ByVal specs As ITableSpecs)
    Set this.specs = specs
End Property

'@label:linelist-data
'@prop-title Linelist specifications providing translations and categories.
'@details
'Returns the ILinelistSpecs instance stored at construction time. Used to
'extract the translation object and to pass as the lData parameter when
'calling specs.RowCategories(lData) and specs.ColumnCategories(lData).
'@return ILinelistSpecs. The linelist specifications.
Public Property Get LinelistData() As ILinelistSpecs
    Set LinelistData = this.lData
End Property

'@label:linelist-data-set
'@prop-title Assign the linelist specifications.
'@param lData ILinelistSpecs. The linelist specifications to assign.
Public Property Set LinelistData(ByVal lData As ILinelistSpecs)
    Set this.lData = lData
End Property

'@label:translations
'@prop-title Translation object for locale-aware labels.
'@details
'Returns the ITranslationObject extracted from the stored linelist
'specifications. Resolves translated strings such as "Total", "Missing",
'"Percent", and date-unit names.
'@return ITranslationObject. The translation object.
Public Property Get Translations() As ITranslationObject
    Set Translations = this.lData.TransObject()
End Property

'@label:number-of-columns
'@prop-title Number of data columns in the table.
'@details
'Returns the count of data columns excluding the row-label column.
'For bivariate and time series tables this equals the number of column
'categories plus optional total and missing columns.
'@return Long. The data column count.
Private Property Get NumberOfColumns() As Long
    NumberOfColumns = this.nbCols
End Property

'@label:number-of-columns-set
'@prop-title Assign the number of data columns.
'@param col Long. The data column count to store.
Private Property Let NumberOfColumns(ByVal col As Long)
    this.nbCols = col
End Property

'@label:named-ranges-list
'@prop-title List of named range names created by this table.
'@details
'Returns a cloned BetterArray containing all the named range names that
'were accumulated during the build process. Used by AnaTabIds to register
'the ranges for export tracking.
'@return BetterArray. The collection of named range names.
Private Property Get NamedRangesList() As BetterArray
    Set NamedRangesList = allRangeNames.Clone()
End Property

'@section Position Properties
'===============================================================================
'@description Properties that compute or retrieve the table position on the worksheet.

'@label:range-exists
'@sub-title Check whether a named range exists on the output worksheet.
'@details
'Checks whether a named range with the given name is defined and
'reachable on the current output worksheet. This is used extensively
'throughout the class to decide whether structural infrastructure
'(section headers, start rows, end-table markers, column setups, etc.)
'has already been created by a previous table in the same section, or
'whether it needs to be created for the first time. The check uses
'On Error Resume Next to silently swallow the error that Excel raises
'when a name does not exist, then inspects whether the resulting Range
'reference is Nothing.
'@param rngName String. The string name of the named range to look up, e.g. "STARTROW_T3" or "ENDTABLE_T5".
'@return Boolean. True if the named range exists and resolves to a valid Range on the output worksheet; False otherwise.
'@remarks This does not distinguish between workbook-scoped and worksheet-scoped names.
'If a workbook-level name exists but does not refer to a range on this sheet,
'the lookup may still succeed. All names in this class are created via
'Cell.Name which produces workbook-scoped names.
Private Function RangeExists(ByVal rngName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = Wksh().Range(rngName)
    On Error GoTo 0
    RangeExists = (Not rng Is Nothing)
End Function

'@label:start-row
'@prop-title Starting row of the table on the output worksheet.
'@details
'Computes or retrieves the starting row for this cross-table on the
'output worksheet. The logic has three branches depending on whether the
'named range STARTROW_{tableId} already exists (reuse), the table is a
'new section or non-temporal type (compute fresh), or the table is a
'non-new-section temporal table (inherit from previous table). For new
'sections and non-temporal tables the start row is computed by finding
'the last used row in column C via End(xlUp), then adding vertical
'padding that varies by table scope. For non-new-section temporal tables
'the start row is inherited from the Previous cross-table so that all
'temporal tables in a section are laid out side-by-side. Bug #3 fallback:
'if the previous table was never built, the code computes a fresh start
'row as if this were a new section.
'@return Long. The 1-based row number on the output worksheet where this table begins.
'@remarks This property is called many times during a single Build cycle (by
'StartColumn, AddRows, AddColumns, AddHeader, NameRanges, Format). After the
'first call creates the STARTROW_ named range, subsequent calls are cheap
'lookups.
Private Property Get StartRow() As Long
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim strtRow As Long
    Dim tabId As String

    Set specs = Specifications()
    tabId = specs.TableId
    Set sh = Wksh()

    With sh
        If Not RangeExists("STARTROW_" & tabId) Then
            If specs.IsNewSection() Or _
               (specs.TableScope <> ScopeTimeSeries And specs.TableScope <> ScopeSpatioTemporal) Then
                ' New section or non-temporal: compute from last used row
                strtRow = .Cells(.Rows.Count, STANDARD_START_COL).End(xlUp).Row + 1
                strtRow = IIf(strtRow = 2, 6, strtRow)

                strtRow = Switch(specs.TableScope = ScopeSpatial, strtRow + 6, _
                                 specs.TableScope = ScopeGlobalSummary, strtRow, _
                                 True, strtRow + 4)

                .Cells(strtRow, STANDARD_START_COL).Name = "STARTROW_" & tabId

            ElseIf Not specs.IsNewSection() Then
                ' Non-new-section temporal: use previous table's start row
                ' Bug #3 fix: check that previous table was actually built
                Dim prevTabId As String
                On Error Resume Next
                prevTabId = specs.Previous.TableId
                On Error GoTo 0

                If prevTabId <> vbNullString And RangeExists("ENDTABLE_" & prevTabId) Then
                    strtRow = Previous.StartRow
                Else
                    ' Previous table was never built (invalid); treat as de-facto new section
                    strtRow = .Cells(.Rows.Count, STANDARD_START_COL).End(xlUp).Row + 1
                    strtRow = IIf(strtRow = 2, 6, strtRow)
                    strtRow = strtRow + 4
                    .Cells(strtRow, STANDARD_START_COL).Name = "STARTROW_" & tabId
                End If
            End If
            StartRow = strtRow
        Else
            StartRow = .Range("STARTROW_" & tabId).Row
        End If
    End With
End Property

'@label:previous
'@prop-title Previous cross-table in the same section.
'@details
'Creates and returns an ICrossTable instance representing the table
'that immediately precedes this one within the same analysis section.
'This is used by temporal table scopes (time series and spatio-temporal)
'where multiple tables share the same row structure and are laid out
'side-by-side in columns. The previous table's start row and end column
'are needed to position the current table correctly. The method delegates
'to the Create factory, passing the Previous spec from the current
'table's ITableSpecs, along with the same output worksheet and
'translation object.
'@return ICrossTable. The preceding cross-table, or Nothing when this is the first table.
'@remarks Each call creates a new CrossTable instance; the result is not cached. For
'temporal tables, this property is typically called once during StartRow
'resolution and should not be called in tight loops.
Private Property Get Previous() As ICrossTable
    Set Previous = Create(Specifications.Previous, Wksh(), LinelistData())
End Property

'@label:start-column
'@prop-title Starting column of the table on the output worksheet.
'@details
'Computes or retrieves the starting column for this cross-table on the
'output worksheet. For most table scopes and for the first table in a
'temporal section, the start column is the constant STANDARD_START_COL
'(column C, index 3). For non-new-section temporal tables (time series
'and spatio-temporal), the start column is determined by finding the
'last used column in the row immediately below StartRow via
'End(xlToLeft). This produces the side-by-side horizontal layout where
'each successive temporal table in a section is placed to the right of
'the previous one. The computed column is persisted as the named range
'STARTCOL_{tableId} so that subsequent calls are idempotent.
'@return Long. The 1-based column number where this table's row category labels begin.
'@remarks The STARTCOL_ named range is stored at (StartRow - 1, strtCol + 1), meaning
'the actual named range cell is one column to the right and one row above the
'logical start column. When reading back, the code subtracts 1 from the
'named range's column to recover the true start column.
Private Property Get StartColumn() As Long
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim strtCol As Long
    Dim tabId As String
    Dim tablType As Byte

    Set specs = Specifications
    tabId = specs.TableId
    tablType = specs.TableScope
    Set sh = Wksh

    With sh
        If Not RangeExists("STARTCOL_" & tabId) Then
            If (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And _
                (Not specs.IsNewSection()) Then
                strtCol = .Cells(StartRow + 1, .Columns.Count).End(xlToLeft).Column
            Else
                strtCol = STANDARD_START_COL
            End If
            StartColumn = strtCol
            .Cells(StartRow - 1, strtCol + 1).Name = "STARTCOL_" & tabId
        Else
            StartColumn = .Range("STARTCOL_" & tabId).Column - 1
        End If
    End With
End Property

'@label:end-row-set
'@prop-title Assign the ending row index.
'@param rw Long. The 1-based ending row index.
Private Property Let EndRow(ByVal rw As Long)
    this.endRw = rw
End Property

'@label:end-column-set
'@prop-title Assign the ending column index.
'@param col Long. The 1-based ending column index.
Private Property Let EndColumn(ByVal col As Long)
    this.endCol = col
End Property

'@label:end-row
'@prop-title Last row used by the table on the worksheet.
'@details
'Returns the 1-based row index of the bottommost cell occupied by this
'cross-table. Used to compute the start row for the next stacked table.
'@return Long. The ending row index.
Private Property Get EndRow() As Long
    EndRow = this.endRw
End Property

'@label:end-column
'@prop-title Last column used by the table on the worksheet.
'@details
'Returns the 1-based column index of the rightmost cell occupied by
'this cross-table. Used by sibling tables and formatting routines to
'avoid overlapping output.
'@return Long. The ending column index.
Private Property Get EndColumn() As Long
    EndColumn = this.endCol
End Property

'@section Range Access
'===============================================================================
'@description Properties that return Range objects for structural regions of the table.

'@label:header-range
'@prop-title Header range of the table on the output worksheet.
'@details
'Returns the Range object spanning the column header row of this
'cross-table. The range starts at (StartRow, StartColumn + 1) and
'extends to (StartRow, EndColumn), covering all column category cells
'but excluding the row-label column at StartColumn. This range is used
'by NameRanges to create the COLUMN_CATEGORIES_ named range, by Format
'to apply header styling, and by ColumnRange to search for a specific
'column value.
'@return Range. The header row range excluding the leftmost row-label cell.
'@remarks EndColumn must be set (by AddColumns/BivariateColumns) before this
'property is called. If EndColumn is 0, the returned range will be invalid.
'@export
Public Property Get HeaderRange() As Range
    Dim sh As Worksheet
    Set sh = Wksh()
    With sh
        Set HeaderRange = .Range(.Cells(StartRow, StartColumn + 1), .Cells(StartRow, EndColumn))
    End With
End Property

'@label:time-series-header
'@prop-title Header label for time series tables.
'@details
'Returns the composite header text for a time-series or spatio-temporal
'table. This text is a human-readable string like
'"Cases -- Notification date -- Region" built from the summary label,
'the time variable label, and optionally the column variable label,
'separated by horizontal-line Unicode characters (U+2500). It is stored
'in the STARTCOL_ named-range cell on the worksheet and used by the
'GoTo navigation dropdown so the user can jump to this table's section.
'@return String. The cached header text string, or an empty string if not yet set.
'@remarks This value is populated during AddColumns for temporal tables only.
'Non-temporal table scopes never set this property.
Private Property Get TimeSeriesHeader() As String
    TimeSeriesHeader = this.tsHeader
End Property

'@label:time-series-header-set
'@prop-title Store the composite header text for temporal tables.
'@param headerText String. The full header string to cache.
Private Property Let TimeSeriesHeader(ByVal headerText As String)
    this.tsHeader = headerText
End Property

'@label:rows-categories-range
'@prop-title Row categories range with optional filtering.
'@details
'Returns a Range covering the row category cells of this cross-table in
'column C (STANDARD_START_COL). The range can optionally include or
'exclude the header label row, and can be trimmed to only the "pure"
'category rows (excluding Total and Missing footer rows). The start row
'varies by table scope: for bivariate, time series, spatial, and
'spatio-temporal tables, the categories begin at StartRow + 1; for
'univariate tables, they begin at StartRow itself. If includeHeaders is
'False, an additional row is skipped. The end row also varies based on
'the onlyCategories flag, trimming footer rows for time series or tables
'with HasTotal.
'@param includeHeaders Optional Boolean. When True (default), includes the row variable label row. Defaults to True.
'@param onlyCategories Optional Boolean. When True, excludes Total and Missing footer rows. Defaults to False.
'@return Range. A single-column Range in column C spanning the requested row categories.
'@throws ErrorUnexpectedState When the table scope is ScopeGlobalSummary.
'@remarks Raises an error for ScopeGlobalSummary because global summary tables do not
'have row categories.
'@export
Public Property Get RowsCategoriesRange(Optional ByVal includeHeaders As Boolean = True, _
                                         Optional ByVal onlyCategories As Boolean = False) As Range
    Dim startRw As Long
    Dim endRw As Long
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim tablType As Byte

    Set specs = Specifications
    tablType = specs.TableScope
    If tablType = ScopeGlobalSummary Then ThrowError ErrorUnexpectedState, "Global summary has no row categories"

    Set sh = Wksh()

    Select Case tablType
    Case ScopeBivariate, ScopeTimeSeries, ScopeSpatial, ScopeSpatioTemporal
        startRw = StartRow + 1
    Case ScopeUnivariate
        startRw = StartRow
    End Select

    If Not includeHeaders Then startRw = startRw + 1

    endRw = Switch( _
        (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And onlyCategories, Me.EndRow - 2, _
        (specs.HasTotal Or tablType = ScopeSpatial) And onlyCategories, Me.EndRow - 1, _
        True, Me.EndRow)

    With sh
        Set RowsCategoriesRange = .Range(.Cells(startRw, STANDARD_START_COL), _
                                          .Cells(endRw, STANDARD_START_COL))
    End With
End Property

'@label:value-exists
'@sub-title Check whether a value exists in a given range.
'@details
'Checks whether a specific string value exists anywhere within the given
'Range by performing an exact, case-sensitive, whole-cell match using
'Excel's Range.Find method. This is used internally by RowRange and
'ColumnRange to verify that a category label is present before
'attempting to locate its row or column position.
'@param rng Range. The Range to search within, typically RowsCategoriesRange or HeaderRange.
'@param searchValue String. The exact string to look for (case-sensitive, whole-cell match).
'@return Boolean. True if at least one cell in rng contains exactly searchValue; False otherwise.
'@remarks Uses xlWhole and MatchCase:=True, so "total" will NOT match "Total". This is
'important because translated labels may have specific capitalization.
Private Function ValueExists(ByVal rng As Range, ByVal searchValue As String) As Boolean
    ValueExists = (Not rng.Find(What:=searchValue, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
End Function

'@label:row-range
'@prop-title Row data range for a specific row value.
'@details
'Finds a row in the cross-table by its category label value and returns
'a horizontal Range spanning that entire row from the start column to the
'end column. The search is performed within RowsCategoriesRange using an
'exact case-sensitive match via ValueExists/Find. If the value is found,
'the returned range includes either the full row from StartColumn (with
'the row-label cell) or from StartColumn + 1 (data cells only), depending
'on the includeHeaders parameter. If the value is not found, Nothing is
'returned.
'@param rowVal String. The exact string value to search for in the row categories column.
'@param includeHeaders Optional Boolean. When True (default), the returned range starts at StartColumn. Defaults to True.
'@return Range. A single-row Range, or Nothing if rowVal is not found.
'@remarks The search uses RowsCategoriesRange() with default parameters, so it searches
'ALL row category cells including footer rows.
'@export
Public Property Get RowRange(ByVal rowVal As String, _
                              Optional ByVal includeHeaders As Boolean = True) As Range
    Dim rng As Range
    Dim valueRow As Long
    Dim sh As Worksheet
    Dim startCol As Long

    Set rng = RowsCategoriesRange()
    Set sh = Wksh()

    If ValueExists(rng, rowVal) Then
        valueRow = rng.Find(What:=rowVal, LookAt:=xlWhole, MatchCase:=True).Row
        If includeHeaders Then
            startCol = StartColumn
        Else
            startCol = StartColumn + 1
        End If
        With sh
            Set RowRange = .Range(.Cells(valueRow, startCol), .Cells(valueRow, EndColumn))
        End With
    End If
End Property

'@label:column-range
'@prop-title Column data range for a specific column value.
'@details
'Finds a column in the cross-table by its header label value and returns
'a vertical Range spanning that entire column from the first data row to
'the end row. The search is performed within HeaderRange using an exact
'case-sensitive match via ValueExists/Find. The start row of the
'returned range varies by table scope. The end row can be trimmed to
'exclude footer rows when onlyCategories is True. The includeHeaders
'parameter extends the range one row upward to include the column header
'cell. Returns Nothing if colVal is not found.
'@param colVal String. The exact string value to search for in the header row.
'@param onlyCategories Optional Boolean. When True, excludes Total and Missing footer rows. Defaults to False.
'@param includeHeaders Optional Boolean. When True, includes the column-label row. Defaults to False.
'@return Range. A single-column Range, or Nothing if colVal is not found in the header range.
'@remarks Callers must check for Nothing before using the result.
'@export
Public Property Get ColumnRange(ByVal colVal As String, _
                                 Optional ByVal onlyCategories As Boolean = False, _
                                 Optional ByVal includeHeaders As Boolean = False) As Range
    Dim rng As Range
    Dim valueColumn As Long
    Dim startRw As Long
    Dim endRw As Long
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim tablType As Byte

    Set rng = HeaderRange()
    Set sh = Wksh()
    Set specs = Specifications
    tablType = specs.TableScope

    startRw = Switch(tablType = ScopeBivariate, StartRow + 2, _
                     (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal), StartRow + 2, _
                     tablType = ScopeSpatial, StartRow + 2, _
                     True, StartRow + 1)

    endRw = Switch( _
        (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And onlyCategories, Me.EndRow - 2, _
        (specs.HasTotal Or tablType = ScopeSpatial) And onlyCategories, Me.EndRow - 1, _
        True, Me.EndRow)

    If includeHeaders Then startRw = startRw - 1

    If ValueExists(rng, colVal) Then
        valueColumn = rng.Find(What:=colVal, LookAt:=xlWhole, MatchCase:=True).Column
        With sh
            Set ColumnRange = .Range(.Cells(startRw, valueColumn), .Cells(endRw, valueColumn))
        End With
    End If
End Property

'@section AddRows
'===============================================================================
'@description Populate row categories on the output worksheet.

'@label:addrows
'@sub-title Add row headers and row category labels to the worksheet.
'@details
'Populates the row category labels in the leftmost column
'(STANDARD_START_COL) of this cross-table on the output worksheet. The
'row content depends entirely on the table scope: ScopeGlobalSummary writes
'a single label cell; ScopeUnivariate writes the row variable's main label
'then all row categories with optional Missing and Total; ScopeBivariate
'is similar but offset by one extra row; ScopeTimeSeries and
'ScopeSpatioTemporal create a fixed-size grid of NB_ROWS_TIME_SERIES (56)
'rows only for new sections; ScopeSpatial writes the geo-level label then
'reserves blank rows for dynamic geo data plus Missing. After writing,
'this sub sets EndRow for subsequent steps.
'@remarks For temporal tables, AddRows is only called once per section (the first table
'or the de-facto first table). Subsequent tables in the same temporal section
'reuse the row structure from the first table and only add columns.
'@export
Private Sub AddRows()
    Dim startRw As Long
    Dim specs As ITableSpecs
    Dim sh As Worksheet
    Dim vars As ILLVariables
    Dim rowVar As String
    Dim rowsData As BetterArray
    Dim trans As ITranslationObject
    Dim transMissing As String
    Dim transTotal As String
    Dim counter As Long
    Dim spatType As String

    startRw = StartRow
    Set sh = Wksh()
    Set specs = Specifications()
    Set trans = Translations()
    Set vars = LLVariables.Create(specs.Dictionary)

    transMissing = trans.TranslatedValue("MSG_NA")
    transTotal = trans.TranslatedValue("MSG_Total")
    Set rowsData = specs.RowCategories(LinelistData())

    With sh
        Select Case specs.TableScope()

        Case ScopeGlobalSummary
            .Cells(startRw, StartColumn).Value = specs.Value("label")

        Case ScopeUnivariate
            If specs.Value("missing") = "yes" Then rowsData.Push transMissing
            rowsData.Push transTotal
            rowsData.ToExcelRange .Cells(startRw + 1, StartColumn)
            rowVar = specs.Value("row")
            .Cells(startRw, StartColumn).Value = vars.Value(colName:="main label", varName:=rowVar)
            EndRow = startRw + rowsData.Length

        Case ScopeBivariate
            If specs.Value("missing") = "row" Or specs.Value("missing") = "all" Then
                rowsData.Push transMissing
            End If
            rowsData.Push transTotal
            rowsData.ToExcelRange .Cells(startRw + 2, StartColumn)
            rowVar = specs.Value("row")
            .Cells(startRw + 1, StartColumn).Value = vars.Value(colName:="main label", varName:=rowVar)
            EndRow = startRw + 1 + rowsData.Length

        Case ScopeTimeSeries, ScopeSpatioTemporal
            ' Add rows only on new sections (or de-facto new sections for Bug #3)
            If specs.IsNewSection() Or IsDeFactoNewSection() Then
                rowsData.Clear
                rowsData.Push trans.TranslatedValue("MSG_Period"), vbNullString
                For counter = 1 To 13
                    rowsData.Push vbNullString, vbNullString, vbNullString, vbNullString
                Next
                ' Total and Missing rows are ALWAYS present in time series
                rowsData.Push transTotal, transMissing
                rowsData.ToExcelRange .Cells(startRw + 1, StartColumn)
            End If
            EndRow = startRw + NB_ROWS_TIME_SERIES

        Case ScopeSpatial
            rowVar = specs.Value("row")
            rowsData.Clear
            spatType = specs.SpatialTableScopes()

            If spatType = "geo" Then
                rowsData.Push vars.Value(colName:="sub section", varName:="adm1_" & rowVar)
            ElseIf spatType = "hf" Then
                rowsData.Push vars.Value(colName:="main label", varName:="hf_" & rowVar)
            End If

            For counter = 1 To CLng(specs.Value("geo max"))
                rowsData.Push vbNullString
            Next

            rowsData.Push transMissing
            rowsData.ToExcelRange .Cells(startRw + 1, StartColumn)
            EndRow = startRw + rowsData.Length

        End Select
    End With
End Sub

'@section Helpers
'===============================================================================
'@description Private helper methods that support the public API.

'@label:percent-label
'@sub-title Build a display label for percentage columns.
'@details
'Builds a display label for percentage columns by appending a directional
'Unicode arrow to the translated percentage string. The arrow indicates
'the direction of the percentage computation: a horizontal double arrow
'(U+2194) for row percentages and a vertical double arrow (U+2195) for
'column percentages. When the percentage type is "all" or any other
'value, no arrow is appended.
'@param percentVal String. The translated percentage label text.
'@param percentType Optional String. The percentage direction: "row", "column", or "all". Defaults to "all".
'@return String. The formatted percentage label string with or without a directional arrow.
Private Function PercentLabel(ByVal percentVal As String, _
                               Optional ByVal percentType As String = "all") As String
    Select Case percentType
    Case "row"
        PercentLabel = percentVal & " " & ChrW(8596)
    Case "column"
        PercentLabel = percentVal & " " & ChrW(8597)
    Case Else
        PercentLabel = percentVal
    End Select
End Function

'@label:bivariate-columns
'@sub-title Write column headers for bivariate, temporal, and spatial tables.
'@details
'Shared column-building logic used by bivariate, time series, spatial,
'and spatio-temporal table scopes. This sub writes the column category
'headers and column-label sub-headers to the output worksheet, starting
'at StartColumn + 1. It handles two layout scenarios: when column
'categories exist, they are written to the header row with optional
'percentage interleaving; when column categories are empty, a single
'summary column is created. This sub also sets NumberOfColumns and
'EndColumn.
'@param startRw Long. The starting row for column headers.
'@param specs ITableSpecs. The ITableSpecs providing column categories and flags.
'@param trans ITranslationObject. The ITranslationObject providing translated strings.
'@param sh Worksheet. The output Worksheet where the column headers are written.
'@remarks NumberOfColumns counts only the "pure" data columns, not including Missing
'or Total. However, EndColumn includes all columns written.
Private Sub BivariateColumns(ByVal startRw As Long, ByVal specs As ITableSpecs, _
                              ByVal trans As ITranslationObject, ByVal sh As Worksheet)
    Dim colLab As String
    Dim colData As BetterArray
    Dim newColData As BetterArray
    Dim labData As BetterArray
    Dim percentageValue As String
    Dim transMissing As String
    Dim transTotal As String
    Dim transPercent As String
    Dim counter As Long

    Set colData = New BetterArray
    Set newColData = New BetterArray
    Set labData = New BetterArray
    colData.LowerBound = 1
    labData.LowerBound = 1
    newColData.LowerBound = 1

    percentageValue = specs.Value("percentage")
    transMissing = trans.TranslatedValue("MSG_NA")
    transTotal = trans.TranslatedValue("MSG_Total")
    transPercent = trans.TranslatedValue("MSG_Percent")

    With sh
        colLab = specs.Value("label")
        If specs.HasMissing Then labData.Push transMissing
        If specs.HasTotal Then labData.Push transTotal
        Set colData = specs.ColumnCategories(LinelistData())

        If colData.Length > 0 Then
            NumberOfColumns = colData.Length

            colData.CONCAT labData.Items
            labData.Clear

            If specs.HasPercentage Then
                counter = 1
                Do While counter <= colData.Length
                    newColData.Push colData.Item(counter), vbNullString
                    labData.Push colLab, PercentLabel(transPercent, percentageValue)
                    counter = counter + 1
                Loop
            Else
                For counter = 1 To colData.Length
                    labData.Push colLab
                Next
                Set newColData = colData.Clone()
            End If

            newColData.ToExcelRange .Cells(startRw, StartColumn + 1), TransposeValues:=True
            labData.ToExcelRange .Cells(startRw + 1, StartColumn + 1), TransposeValues:=True
        Else
            ' No column data — just a summary function
            labData.Push colLab
            NumberOfColumns = 1
            labData.ToExcelRange .Cells(startRw, StartColumn + 1), TransposeValues:=True
            ' Placeholder for End(xlToLeft) detection in next table's StartColumn
            .Cells(startRw + 1, StartColumn + 1).Value = " "
        End If

        EndColumn = StartColumn + labData.Length
    End With
End Sub

'@section AddColumns
'===============================================================================
'@description Populate column headers on the output worksheet.

'@label:addcolumns
'@sub-title Add column headers and column category labels to the worksheet.
'@details
'Writes column headers and category labels to the output worksheet. The
'column structure depends on the table scope: ScopeGlobalSummary creates two
'fixed columns; ScopeUnivariate creates one data column plus an optional
'percentage column; ScopeBivariate delegates to BivariateColumns then writes
'the column variable label; ScopeTimeSeries and ScopeSpatioTemporal delegate
'to BivariateColumns then construct a composite TimeSeriesHeader string;
'ScopeSpatial delegates entirely to BivariateColumns. After writing, this
'sub sets EndColumn and NumberOfColumns.
'@remarks For global summary tables, the two-column header is created only once and
'shared by all subsequent global summary rows. The COLGS_SET named range acts
'as a guard to prevent duplicate creation.
'@export
Private Sub AddColumns()
    Dim startRw As Long
    Dim specs As ITableSpecs
    Dim sh As Worksheet
    Dim vars As ILLVariables
    Dim trans As ITranslationObject
    Dim colData As BetterArray
    Dim colVarLab As String
    Dim sumLab As String
    Dim timeVarLab As String
    Dim tabId As String
    Dim tablType As Byte
    Dim spType As String
    Dim percentageValue As String
    Dim transPercent As String

    startRw = StartRow
    Set sh = Wksh()
    Set specs = Specifications()
    Set trans = Translations()
    Set vars = LLVariables.Create(specs.Dictionary)
    tabId = specs.TableId()
    tablType = specs.TableScope()
    Set colData = New BetterArray
    colData.LowerBound = 1
    spType = specs.SpatialTableScopes()
    percentageValue = specs.Value("percentage")
    transPercent = trans.TranslatedValue("MSG_Percent")

    With sh
        Select Case tablType

        Case ScopeGlobalSummary
            If Not RangeExists("COLGS_SET") Then
                .Cells(startRw - 1, StartColumn + 1).Value = trans.TranslatedValue("MSG_AllData")
                .Cells(startRw - 1, StartColumn + 2).Value = trans.TranslatedValue("MSG_FilteredData")
                .Cells(startRw - 1, StartColumn + 1).Name = "COLGS_SET"
            End If
            EndColumn = StartColumn + 2
            NumberOfColumns = 2

        Case ScopeUnivariate
            Dim colLab As String
            colLab = specs.Value("label")
            colData.Push colLab
            If percentageValue = "yes" Then colData.Push transPercent
            colData.ToExcelRange .Cells(startRw, StartColumn + 1), TransposeValues:=True
            EndColumn = StartColumn + colData.Length
            NumberOfColumns = 1

        Case ScopeBivariate
            BivariateColumns startRw, specs, trans, sh
            .Range("STARTCOL_" & tabId).Value = vars.Value( _
                colName:="main label", varName:=specs.Value("column"))

        Case ScopeTimeSeries, ScopeSpatioTemporal
            BivariateColumns startRw, specs, trans, sh

            If tablType = ScopeTimeSeries Then
                colVarLab = vars.Value(colName:="main label", varName:=specs.Value("column"))
            ElseIf spType = "geo" Then
                colVarLab = vars.Value(colName:="sub section", varName:="adm1_" & specs.Value("column"))
            ElseIf spType = "hf" Then
                colVarLab = vars.Value(colName:="main label", varName:="hf_" & specs.Value("column"))
            End If

            timeVarLab = vars.Value(colName:="main label", varName:=specs.Value("time"))
            sumLab = specs.Value("label")

            If specs.Value("column") = vbNullString Then
                TimeSeriesHeader = sumLab & " " & ChrW(9472) & " " & timeVarLab
            Else
                TimeSeriesHeader = sumLab & " " & ChrW(9472) & " " & timeVarLab & _
                                   " " & ChrW(9472) & " " & colVarLab
            End If

            .Range("STARTCOL_" & tabId).Value = TimeSeriesHeader

        Case ScopeSpatial
            BivariateColumns startRw, specs, trans, sh

        End Select
    End With
End Sub

'@section AddHeader
'===============================================================================
'@description Create section headers, titles, and date/geo controls.

'@label:addheader
'@sub-title Add the table title and header to the worksheet.
'@details
'Inserts the structural header infrastructure above the data area of this
'cross-table: section titles, table titles, date/time input controls for
'temporal tables, and geo-level selection controls for spatial tables.
'This is the first step called by Build and may insert rows into the
'worksheet (shifting existing content down) to make room for the header.
'Behavior varies by table scope: univariate and bivariate write a title
'and optional section header; global summary writes a shared label once;
'temporal tables insert date controls and optional geo-level inputs;
'spatial tables create admin-unit or health-facility dropdowns. This sub
'also initializes the allRangeNames BetterArray.
'@remarks Row insertions shift the entire sheet down, which is why AddHeader must run
'before AddRows, AddColumns, and NameRanges. For temporal tables, the date
'controls are unlocked (Locked = False) so the user can interactively change
'dates at runtime.
'@export
Private Sub AddHeader()
    Dim startRw As Long
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim counter As Long
    Dim trans As ITranslationObject
    Dim tabId As String
    Dim rng As Range
    Dim cellRng As Range
    Dim nGeo As Long
    Dim tablType As Byte
    Dim nbShift As Long
    Dim timeUnitName As String
    Dim sptGeoRngTag As String

    Set sh = Wksh()
    Set specs = Specifications
    Set trans = Translations()

    Set allRangeNames = New BetterArray

    startRw = StartRow
    tabId = specs.TableId
    tablType = specs.TableScope
    sptGeoRngTag = "INPUTSPTGEO_"
    If specs.Value("spatial type") = "hf" Then sptGeoRngTag = "INPUTSPTHF_"

    With sh
        Select Case tablType

        Case ScopeUnivariate, ScopeBivariate
            .Cells(startRw - 2, STANDARD_START_COL).Value = specs.Value("title")
            .Cells(startRw - 2, STANDARD_START_COL).Name = "TITLE_" & tabId
            allRangeNames.Push "TITLE_" & tabId

            If specs.IsNewSection() And (Not RangeExists("SECTION_" & tabId)) Then
                For counter = 1 To 3
                    .Rows(startRw - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next
                .Range("TITLE_" & tabId).Cells(-1, 1) = specs.Value("section")
                .Range("TITLE_" & tabId).Cells(-1, 1).Name = "SECTION_" & tabId
                allRangeNames.Push "SECTION_" & tabId
            End If

        Case ScopeGlobalSummary
            If Not RangeExists("ROWGS_SET") Then
                For counter = 1 To 3
                    .Rows(startRw - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next
                .Cells(startRw - 2, STANDARD_START_COL).Value = trans.TranslatedValue("MSG_GlobalSummary")
                .Cells(startRw - 2, STANDARD_START_COL).Name = "ROWGS_SET"
                allRangeNames.Push "ROWGS_SET"
            End If

        Case ScopeTimeSeries, ScopeSpatioTemporal
            ' Bug #3 fix: treat as new section if previous table was never built
            If (specs.IsNewSection() Or IsDeFactoNewSection()) And _
               (Not RangeExists("SECTION_" & tabId)) Then

                If tablType = ScopeSpatioTemporal Then
                    nGeo = 5
                    On Error Resume Next
                    nGeo = CLng(specs.Value("n geo"))
                    On Error GoTo 0
                    nbShift = nGeo + 14
                Else
                    nbShift = 7
                End If

                For counter = 1 To nbShift
                    .Rows(startRw - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next
                .Cells(startRw - 2, StartColumn).Value = specs.Value("section")
                .Cells(startRw - 2, StartColumn).Name = "SECTION_" & tabId
                allRangeNames.Push "SECTION_" & tabId

                Set rng = .Range("SECTION_" & tabId)

                ' Spatio-temporal geo-level selection
                If tablType = ScopeSpatioTemporal Then
                    Set rng = rng.Offset(4)
                    rng.Value = trans.TranslatedValue("MSG_InputGeoLevels")
                    rng.Name = "INPUTSPTGEO_LABEL_" & tabId
                    allRangeNames.Push "INPUTSPTGEO_LABEL_" & tabId
                    rng.Cells(1, 2).Value = nGeo
                    rng.Cells(1, 2).Name = "INPUTSPTGEO_NGEO_" & tabId

                    Set cellRng = rng.Offset(2)
                    For counter = 1 To nGeo
                        cellRng.Cells(counter, 1).Name = sptGeoRngTag & counter & "_" & tabId
                        cellRng.Cells(counter, 2).Value = 1
                    Next

                    Set rng = .Range(cellRng.Cells(1, 1), cellRng.Cells(nGeo, 1))
                    rng.Name = "INPUTSPTGEO_USER_" & tabId
                    allRangeNames.Push "INPUTSPTGEO_USER_" & tabId
                    rng.Locked = False
                    rng.Offset(, 1).Name = "INPUTPREVADM_USER_" & tabId
                    Set rng = cellRng.Offset(nGeo + 3)
                End If

                ' Date controls (start date, time unit, end date)
                rng.Cells(6, 1).Name = "INFO_ANA_PERIOD_" & tabId
                rng.Cells(3, 3).Value = trans.TranslatedValue("MSG_StartDate")
                rng.Cells(3, 3).Name = "START_DATE_LABEL_" & tabId
                rng.Cells(4, 3).Locked = False
                rng.Cells(4, 3).Name = "USER_START_DATE_" & tabId
                rng.Cells(5, 3).Name = "INFO_START_DATE_" & tabId
                rng.Cells(6, 3).Name = "START_DATE_" & tabId

                rng.Cells(3, 5).Value = trans.TranslatedValue("MSG_TimeUnit")
                rng.Cells(3, 5).Name = "TIME_UNIT_LABEL_" & tabId
                rng.Cells(4, 5).Name = "TIME_UNIT_" & tabId
                rng.Cells(4, 5).Locked = False

                rng.Cells(3, 7).Value = trans.TranslatedValue("MSG_EndDate")
                rng.Cells(3, 7).Name = "END_DATE_LABEL_" & tabId
                rng.Cells(4, 7).Name = "USER_END_DATE_" & tabId
                rng.Cells(4, 7).Locked = False
                rng.Cells(5, 7).Name = "INFO_END_DATE_" & tabId
                rng.Cells(6, 7).Name = "END_DATE_" & tabId

                rng.Cells(4, 9).Name = "VALIDATION_MIN_DATE_" & tabId
                rng.Cells(4, 11).Name = "VALIDATION_MAX_DATE_" & tabId
                rng.Cells(4, 13).Name = "MIN_MIN_DATE_" & tabId
                rng.Cells(4, 15).Name = "MAX_MAX_DATE_" & tabId

                allRangeNames.Push "INFO_ANA_PERIOD_" & tabId, _
                    "START_DATE_LABEL_" & tabId, "USER_START_DATE_" & tabId, _
                    "INFO_START_DATE_" & tabId, "START_DATE_" & tabId, _
                    "END_DATE_LABEL_" & tabId, "USER_END_DATE_" & tabId, _
                    "INFO_END_DATE_" & tabId, "END_DATE_" & tabId

                timeUnitName = "TIME_UNIT_LIST"
                If tablType = ScopeSpatioTemporal Then timeUnitName = "SP" & timeUnitName

                Set cellRng = .Cells(1, 26)
                If Not RangeExists(timeUnitName) Then
                    cellRng.Cells(1, 1).Value = trans.TranslatedValue("MSG_Day")
                    cellRng.Cells(2, 1).Value = trans.TranslatedValue("MSG_Week")
                    cellRng.Cells(3, 1).Value = trans.TranslatedValue("MSG_Month")
                    cellRng.Cells(4, 1).Value = trans.TranslatedValue("MSG_Quarter")
                    cellRng.Cells(5, 1).Value = trans.TranslatedValue("MSG_Year")
                    .Range(.Cells(1, cellRng.Column), .Cells(5, cellRng.Column)).Name = timeUnitName
                    .Range(timeUnitName).Font.Color = vbWhite
                    .Range(timeUnitName).Interior.Color = vbWhite
                    .Range(timeUnitName).Locked = True
                End If

                .Range("TIME_UNIT_" & tabId).Value = cellRng.Cells(2, 1).Value
            End If

        Case ScopeSpatial
            If Not RangeExists("ADM_UNIT_LIST") Then
                Set cellRng = .Cells(1, 26)
                cellRng.Cells(1, 1).Formula = "= RNG_ADM1NAME"
                cellRng.Cells(2, 1).Formula = "= RNG_ADM2NAME"
                cellRng.Cells(3, 1).Formula = "= RNG_ADM3NAME"
                cellRng.Cells(4, 1).Formula = "= RNG_ADM4NAME"
                .Range(.Cells(1, cellRng.Column), .Cells(4, cellRng.Column)).Name = "ADM_UNIT_LIST"
                .Range("ADM_UNIT_LIST").Font.Color = vbWhite
                .Range("ADM_UNIT_LIST").Interior.Color = vbWhite
                .Range("ADM_UNIT_LIST").Locked = True
            End If

            If Not RangeExists("POPULATION_FACTOR_LIST") Then
                Set cellRng = .Cells(1, 27)
                cellRng.Cells(1, 1).Value = trans.TranslatedValue("MSG_NoDevide")
                cellRng.Cells(2, 1).Value = trans.TranslatedValue("MSG_Devide")
                .Range(.Cells(1, cellRng.Column), .Cells(2, cellRng.Column)).Name = "POPULATION_FACTOR_LIST"
                .Range("POPULATION_FACTOR_LIST").Font.Color = vbWhite
                .Range("POPULATION_FACTOR_LIST").Interior.Color = vbWhite
                .Range("POPULATION_FACTOR_LIST").Locked = True
            End If

            If (Not RangeExists("ADM_DROPDOWN_" & tabId)) And _
               (Not RangeExists("HF_DROPDOWN_" & tabId)) Then

                If specs.SpatialTableScopes = "geo" Then
                    .Cells(startRw - 2, STANDARD_START_COL).Name = "ADM_DROPDOWN_" & tabId
                    .Cells(startRw - 2, STANDARD_START_COL).Value = trans.TranslatedValue("MSG_SelectAdmin")

                    Set rng = .Range("ADM_DROPDOWN_" & tabId)
                    rng.Cells(1, 2).Value = "adm1"
                    rng.Cells(1, 2).Name = "PREVIOUS_ADM_" & tabId
                    rng.Cells(1, 3).Value = trans.TranslatedValue("MSG_SelectPOPFACT")
                    rng.Cells(1, 3).Name = "DEVIDEPOP_" & tabId
                    rng.Cells(0, 4).Value = trans.TranslatedValue("MSG_MultiplyBy")
                    rng.Cells(0, 4).Name = "POPFACTLABEL_" & tabId
                    rng.Cells(1, 4).Value = 100
                    rng.Cells(1, 4).Name = "POPFACT_" & tabId
                    rng.Cells(1, 5).Name = "POPPREVFACT_" & tabId
                    rng.Cells(1, 5).Value = "0"

                    allRangeNames.Push "DEVIDEPOP_" & tabId, "POPFACT_" & tabId, _
                                        "ADM_DROPDOWN_" & tabId

                ElseIf specs.SpatialTableScopes = "hf" Then
                    .Cells(startRw - 2, STANDARD_START_COL).Name = "HF_DROPDOWN_" & tabId
                    Set rng = .Range("HF_DROPDOWN_" & tabId)
                    rng.Cells(1, 1).Value = trans.TranslatedValue("MSG_HF")
                    allRangeNames.Push "HF_DROPDOWN_" & tabId
                End If

                If specs.IsNewSection() Then
                    For counter = 1 To 3
                        .Rows(rng.Row - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    Next
                    rng.Cells(-3, 1).Value = specs.Value("section")
                    rng.Cells(-3, 1).Name = "SECTION_" & tabId
                    allRangeNames.Push "SECTION_" & tabId
                End If

                .Rows(rng.Row - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                rng.Cells(-1, 1).Value = specs.Value("title")
                rng.Cells(-1, 1).Name = "TITLE_" & tabId
                allRangeNames.Push "TITLE_" & tabId
            End If

        End Select
    End With
End Sub

'@section NameRanges
'===============================================================================
'@description Create all structural named ranges for the table.

'@label:name-ranges
'@sub-title Create all named ranges for the table data regions.
'@details
'Creates all structural named ranges for this cross-table after the data
'area (rows and columns) has been populated. These named ranges serve as
'stable references for formula generation, VBA-driven data updates, and
'export tracking. The method creates ENDTABLE_, ROW_CATEGORIES_,
'time-period markers, per-column VALUES_COL_ and LABEL_COL_ ranges,
'Missing/Total/Percent row and column ranges, intersection cells,
'INTERIOR_VALUES_, OUTER_VALUES_ (spatial only), and COLUMN_CATEGORIES_.
'All created range names are accumulated in the allRangeNames BetterArray
'for retrieval via NamedRangesList.
'@throws ErrorUnexpectedState When NumberOfColumns is 0, indicating AddColumns was not called.
'@remarks Exits immediately for ScopeGlobalSummary (which has no row categories or
'column structure beyond the two fixed columns).
'@export
Private Sub NameRanges()
    Dim specs As ITableSpecs
    Dim sh As Worksheet
    Dim nbCols As Long
    Dim catRng As Range
    Dim rng As Range
    Dim nbColCounter As Long
    Dim counter As Long
    Dim trans As ITranslationObject
    Dim labCol As Long
    Dim cData As BetterArray
    Dim tData As BetterArray
    Dim rw As Long
    Dim col As Long
    Dim addBegin As String
    Dim addEnd As String
    Dim tabId As String
    Dim cellRng As Range
    Dim tablType As Byte
    Dim isNewSect As Boolean

    Set sh = Wksh()
    Set trans = Translations()
    Set specs = Specifications
    tabId = specs.TableId
    tablType = specs.TableScope
    ' Bug #3 fix: treat as new section if previous table was never built
    isNewSect = specs.IsNewSection() Or IsDeFactoNewSection()

    If tablType = ScopeGlobalSummary Then Exit Sub

    With sh
        If Not RangeExists("ENDTABLE_" & tabId) Then
            .Cells(EndRow + 1, StartColumn).Name = "ENDTABLE_" & tabId
        End If

        nbCols = NumberOfColumns
        If nbCols = 0 Then ThrowError ErrorUnexpectedState, _
                                      "Table " & tabId & " has 0 columns"

        ' Row categories — only for new sections or non-temporal tables
        If isNewSect Or (tablType <> ScopeTimeSeries And tablType <> ScopeSpatioTemporal) Then
            Set catRng = RowsCategoriesRange(onlyCategories:=True, includeHeaders:=False)
            catRng.Name = "ROW_CATEGORIES_" & tabId
            catRng.Cells(0, 1).Name = "LABEL_ROW_CATEGORIES_" & tabId
            allRangeNames.Push "LABEL_ROW_CATEGORIES_" & tabId
            allRangeNames.Push "ROW_CATEGORIES_" & tabId

            If tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal Then
                Set rng = catRng.Offset(, -1)
                rng.Name = "END_TIME_PERIOD_" & tabId
                Set rng = rng.Offset(, -1)
                rng.Cells(1, 1).Name = "FIRST_VALUE_START_TIME_" & tabId
                Set rng = .Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, 1))
                rng.Name = "START_TIME_PERIOD_" & tabId
            End If
        End If

        ' Non-new-section temporal: derive catRng from ENDTABLE position
        ' Bug #3 fix: only enter this block if section infrastructure actually exists
        If (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And Not isNewSect Then
            Dim prevTabId As String
            prevTabId = specs.Previous.TableId
            If Not RangeExists("ENDTABLE_" & prevTabId) Then
                ThrowError ErrorUnexpectedState, _
                           "Cannot append time series: no previous table found for " & prevTabId
            End If
            Set cellRng = .Range("ENDTABLE_" & specs.TableId)
            Set catRng = .Range(cellRng.Offset(1 - NB_ROWS_TIME_SERIES), cellRng.Offset(-3))
        End If

        ' Name all value columns
        counter = StartColumn
        nbColCounter = 1
        Set rng = catRng.Offset(, 1)

        Do While (counter <= EndColumn) And (nbColCounter <= nbCols)
            rng.Name = "VALUES_COL_" & nbColCounter & "_" & tabId
            allRangeNames.Push "VALUES_COL_" & nbColCounter & "_" & tabId

            Select Case tablType
            Case ScopeBivariate, ScopeTimeSeries, ScopeSpatial, ScopeSpatioTemporal
                labCol = -1
            Case ScopeUnivariate
                labCol = 0
            End Select

            rng.Cells(labCol, 1).Name = "LABEL_COL_" & nbColCounter & "_" & tabId
            allRangeNames.Push "LABEL_COL_" & nbColCounter & "_" & tabId

            Set rng = rng.Offset(, 1)

            If specs.HasPercentage Then
                rng.Name = "PERC_COL_" & nbColCounter & "_" & tabId
                allRangeNames.Push "PERC_COL_" & nbColCounter & "_" & tabId

                If tablType = ScopeUnivariate Then
                    rng.Cells(labCol, 1).Name = "PERC_LABEL_COL_" & tabId
                    allRangeNames.Push "PERC_LABEL_COL_" & tabId
                End If

                Set rng = rng.Offset(, 1)
                counter = counter + 1
            End If

            counter = counter + 1
            nbColCounter = nbColCounter + 1
        Loop

        ' Name Missing/Total/Percent rows and columns
        Set cData = New BetterArray
        Set tData = New BetterArray
        cData.Push "MISSING", "TOTAL", "PERC"
        tData.Push "MSG_NA", "MSG_Total", "MSG_Percent"

        For counter = cData.LowerBound To cData.UpperBound
            ' Row ranges
            Set rng = RowRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=True)
            If Not rng Is Nothing Then
                rng.Name = cData.Item(counter) & "_ROW_" & tabId
                allRangeNames.Push cData.Item(counter) & "_ROW_" & tabId
                Set rng = RowRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=False)
                rng.Name = cData.Item(counter) & "_ROW_VALUES_" & tabId
                allRangeNames.Push cData.Item(counter) & "_ROW_VALUES_" & tabId
            End If

            ' Column ranges (not for univariate)
            If tablType <> ScopeUnivariate Then
                Set rng = ColumnRange(trans.TranslatedValue(tData.Item(counter)), _
                                      includeHeaders:=True, onlyCategories:=True)
                If Not rng Is Nothing Then
                    rng.Name = cData.Item(counter) & "_COL_" & tabId
                    allRangeNames.Push cData.Item(counter) & "_COL_" & tabId
                    rng.Cells(0, 1).Name = cData.Item(counter) & "_LABEL_COL_" & tabId
                    Set rng = ColumnRange(trans.TranslatedValue(tData.Item(counter)), _
                                          includeHeaders:=False, onlyCategories:=True)
                    rng.Name = cData.Item(counter) & "_COL_VALUES_" & tabId
                    allRangeNames.Push cData.Item(counter) & "_COL_VALUES_" & tabId

                    If specs.HasPercentage Then
                        Set rng = rng.Offset(, 1)
                        rng.Name = cData.Item(counter) & "_PERC_VALUES_" & tabId
                        allRangeNames.Push cData.Item(counter) & "_PERC_VALUES_" & tabId
                        .Range(cData.Item(counter) & "_COL_" & tabId).Offset(, 1).Name = _
                            cData.Item(counter) & "_PERC_" & tabId
                    End If
                End If
            End If
        Next

        ' Intersection cells (total/missing combinations)
        If RangeExists("TOTAL_COL_" & tabId) Then
            If RangeExists("TOTAL_ROW_" & tabId) Then
                rw = .Range("TOTAL_ROW_" & tabId).Row
                col = .Range("TOTAL_COL_" & tabId).Column
                .Cells(rw, col).Name = "TOTAL_TOTAL_" & tabId
                allRangeNames.Push "TOTAL_TOTAL_" & tabId
            End If
            If RangeExists("MISSING_ROW_" & tabId) Then
                rw = .Range("MISSING_ROW_" & tabId).Row
                col = .Range("TOTAL_COL_" & tabId).Column
                .Cells(rw, col).Name = "MISSING_TOTAL_" & tabId
                allRangeNames.Push "MISSING_TOTAL_" & tabId
                If RangeExists("MISSING_COL_" & tabId) Then
                    col = .Range("MISSING_COL_" & tabId).Column
                    .Cells(rw, col).Name = "MISSING_MISSING_" & tabId
                    allRangeNames.Push "MISSING_MISSING_" & tabId
                End If
            End If
        End If

        If RangeExists("TOTAL_ROW_" & tabId) Then
            If RangeExists("MISSING_COL_" & tabId) Then
                col = .Range("MISSING_COL_" & tabId).Column
                rw = .Range("TOTAL_ROW_" & tabId).Row
                .Cells(rw, col).Name = "TOTAL_MISSING_" & tabId
                allRangeNames.Push "TOTAL_MISSING_" & tabId
            End If
        End If

        ' Interior values (all data cells)
        addBegin = Split(.Range("VALUES_COL_1_" & tabId).Address, ":")(0)
        addEnd = IIf(specs.HasPercentage, "PERC_COL_", "VALUES_COL_")
        addEnd = Split(.Range(addEnd & nbCols & "_" & tabId).Address, ":")(1)
        .Range(addBegin & ":" & addEnd).Name = "INTERIOR_VALUES_" & tabId
        allRangeNames.Push "INTERIOR_VALUES_" & tabId

        ' Spatial outer values for population computation
        If tablType = ScopeSpatial Then
            Set rng = .Range("INTERIOR_VALUES_" & tabId)
            Set rng = .Range(rng.Cells(1, -1), rng.Cells(rng.Rows.Count + 1, rng.Columns.Count + 4))
            rng.Name = "OUTER_VALUES_" & tabId
            allRangeNames.Push "OUTER_VALUES_" & tabId
        End If

        ' Column categories header
        Set rng = HeaderRange()
        rng.Name = "COLUMN_CATEGORIES_" & tabId
        allRangeNames.Push "COLUMN_CATEGORIES_" & tabId
    End With
End Sub

'@section Bug #3 Helper
'===============================================================================
'@description Detect de-facto new sections when the previous table was never built.

'@label:is-de-facto-new-section
'@sub-title Determine whether this table is a de-facto new section.
'@details
'Determines whether this table should be treated as a de-facto new
'section even though its ITableSpecs.IsNewSection returns False. This
'situation arises from Bug #3: when the previous table in the same section
'was invalid and was therefore skipped during the build process, the
'section infrastructure was never created. The logic checks whether the
'previous table's ENDTABLE_ named range exists; if not, this table must
'act as the section anchor. The result is cached after the first
'computation since IsDeFactoNewSection is called multiple times during
'a single Build cycle.
'@return Boolean. True if this table must assume the section-anchor role; False otherwise.
'@remarks For tables where specs.IsNewSection is True, this function always returns
'False. The typical usage pattern is: "If specs.IsNewSection() Or IsDeFactoNewSection() Then ...".
Private Function IsDeFactoNewSection() As Boolean

    If this.deFactoNewSectCached Then
        IsDeFactoNewSection = this.deFactoNewSect
        Exit Function
    End If

    Dim specs As ITableSpecs
    Dim result As Boolean
    Set specs = Specifications

    ' Only relevant for non-new-section temporal tables
    If specs.IsNewSection() Then
        result = False
    ElseIf specs.TableScope <> ScopeTimeSeries And specs.TableScope <> ScopeSpatioTemporal Then
        result = False
    Else
        Dim prevTabId As String
        On Error Resume Next
        prevTabId = specs.Previous.TableId
        On Error GoTo 0

        If prevTabId = vbNullString Then
            result = True
        Else
            ' If the previous table's ENDTABLE range doesn't exist, it was never built
            result = Not RangeExists("ENDTABLE_" & prevTabId)
        End If
    End If

    this.deFactoNewSect = result
    this.deFactoNewSectCached = True
    IsDeFactoNewSection = result
End Function

'@section Format
'===============================================================================
'@description Apply visual formatting to all structural regions of the cross-table.

'@label:format
'@sub-title Apply formatting to the table using design specifications.
'@details
'Applies comprehensive visual formatting to every part of the cross-table
'using the design format object (ILLFormat). This is the final step after
'Build completes the structural layout. Formatting is applied in order:
'section header, table title, spatial dropdown, column headers, row and
'column categories, interior values, Total/Missing rows and columns,
'single cells, info cells, hidden cells, percentage columns, whole table
'borders, and end-table gap. Bug #1 fix: for temporal tables where
'HasTotal is True but TotalRequested is False, hides the Total column
'entirely since it is needed internally for percentage computation but
'should be invisible to the user.
'@param desFormat ILLFormat. The formatting specification providing ApplyFormat for each scope.
'@remarks Format must be called after Build because it relies on named ranges being
'present. Some operations are idempotent, guarded by HiddenNames like
'ROWGS_FORMATSET and MERGED_.
'@export
Private Sub Format(ByVal desFormat As ILLFormat)
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim rng As Range
    Dim rngTable As BetterArray
    Dim counter As Long
    Dim rngName As String
    Dim tabId As String
    Dim designFormat As ILLFormat
    Dim wholeTableStart As Long
    Dim hasPercent As Boolean
    Dim newSection As Boolean
    Dim tablType As Byte

    Set rngTable = New BetterArray
    Set specs = Specifications
    Set sh = Wksh()
    Set designFormat = desFormat

    tabId = specs.TableId
    tablType = specs.TableScope
    hasPercent = specs.HasPercentage
    newSection = specs.IsNewSection() Or IsDeFactoNewSection()

    With sh
        Dim hn As IHiddenNames
        Set hn = HiddenNames.Create(sh)

        ' Sections ==============================================================
        If newSection Then
            Set rng = .Range("SECTION_" & tabId)
            designFormat.ApplyFormat rng, AnalysisSection
        ElseIf RangeExists("ROWGS_SET") And (Not hn.HasName("ROWGS_FORMATSET")) Then
            Set rng = .Range("ROWGS_SET")
            designFormat.ApplyFormat rng, scope:=AnalysisSection
            hn.EnsureName "ROWGS_FORMATSET", True, "Boolean"
        End If

        ' Title (GlobalSummary and TimeSeries don't have titles)
        If (tablType <> ScopeGlobalSummary) And (tablType <> ScopeTimeSeries) And _
           (tablType <> ScopeSpatioTemporal) Then
            Set rng = .Range("TITLE_" & tabId)
            designFormat.ApplyFormat rng, AnalysisTableTitle
        End If

        ' Spatial dropdown formatting
        If tablType = ScopeSpatial Then
            If specs.SpatialTableScopes = "geo" Then
                Set rng = .Range("ADM_DROPDOWN_" & tabId)
                designFormat.ApplyFormat rng, scope:=LinelistSelectionDropdown
                Set rng = .Range("DEVIDEPOP_" & tabId)
                designFormat.ApplyFormat rng, scope:=LinelistSelectionDropdown
            ElseIf specs.SpatialTableScopes = "hf" Then
                Set rng = .Range("HF_DROPDOWN_" & tabId)
                designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
            End If
            Set rng = .Range("ROW_CATEGORIES_" & tabId).Offset(, -1)
            designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
            Set rng = rng.Offset(, -1)
            designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
        End If

        ' Headers ===============================================================
        Select Case tablType
        Case ScopeGlobalSummary
            If Not hn.HasName("COLGS_FORMATSET") Then
                Set rng = .Range("COLGS_SET")
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(1, 2))
                designFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
                hn.EnsureName "COLGS_FORMATSET", True, "Boolean"
            End If
        Case ScopeUnivariate
            Set rng = .Range("COLUMN_CATEGORIES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
            Set rng = .Range("LABEL_ROW_CATEGORIES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelRow
        Case ScopeBivariate, ScopeSpatial
            Set rng = .Range("LABEL_ROW_CATEGORIES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelRow
            Set rng = .Range("STARTCOL_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelColumn
        Case ScopeTimeSeries, ScopeSpatioTemporal
            If newSection Then
                Set rng = .Range("LABEL_ROW_CATEGORIES_" & tabId)
                designFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelRow
            End If
            Set rng = .Range("STARTCOL_" & tabId)
            Set rng = .Range(.Cells(rng.Row, rng.Column), .Cells(rng.Row, EndColumn()))
            designFormat.ApplyFormat rng, scope:=AnalysisTimeSeriesHeader
        End Select

        ' Categories rows and columns ===========================================
        Select Case tablType
        Case ScopeGlobalSummary
            Set rng = .Range("STARTROW_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisGlobalSummaryRow
        Case ScopeUnivariate
            Set rng = .Range("ROW_CATEGORIES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisCategoriesUnivariateRow
        Case ScopeBivariate, ScopeTimeSeries, ScopeSpatial, ScopeSpatioTemporal
            If ((tablType <> ScopeTimeSeries) And (tablType <> ScopeSpatioTemporal)) Or newSection Then
                Set rng = .Range("ROW_CATEGORIES_" & tabId)
                designFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateRow
                If tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal Then
                    Set rng = rng.Offset(, -1)
                    Set rng = .Range(rng.Cells(1, 0), rng.Cells(rng.Rows.Count + 1, 1))
                    designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
                    rngTable.Clear
                    rngTable.Push "START_DATE_", "END_DATE_", "VALIDATION_MIN_DATE_", _
                                  "VALIDATION_MAX_DATE_", "MIN_MIN_DATE_", "MAX_MAX_DATE_"
                    For counter = rngTable.LowerBound To rngTable.UpperBound
                        rngName = rngTable.Item(counter) & tabId
                        Set rng = .Range(rngName)
                        designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
                        designFormat.ApplyFormat rng, scope:=AnalysisDateCell
                    Next
                End If
            End If

            If RangeExists("INPUTSPTGEO_USER_" & tabId) Then
                Set rng = .Range("INPUTSPTGEO_USER_" & tabId)
                designFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateRow
            End If

            ' Column label merging
            Set rng = .Range("LABEL_COL_1_" & tabId)
            If Not hn.HasName("MERGED_" & tabId) Then
                Do While rng.Column <= EndColumn()
                    If hasPercent Then
                        Set rng = rng.Offset(, 2)
                        .Range(rng.Cells(1, -1), rng.Cells(1, 0)).Merge
                    Else
                        Set rng = rng.Offset(, 1)
                    End If
                Loop
                hn.EnsureName "MERGED_" & tabId, True, "Boolean"
            End If

            Set rng = .Range("COLUMN_CATEGORIES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateColumn
            Set rng = rng.Offset(1)
            designFormat.ApplyFormat rng, scope:=AnalysisFirstRow
        End Select

        ' Interior values =======================================================
        If tablType <> ScopeGlobalSummary Then
            Set rng = .Range("INTERIOR_VALUES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisInteriorTable
        End If

        ' Unique cells — Total/Missing rows and columns ==========================
        rngTable.Clear
        rngTable.Push "TOTAL_ROW_VALUES_", "TOTAL_COL_VALUES_", _
                      "MISSING_ROW_VALUES_", "MISSING_COL_VALUES_", _
                      "TOTAL_PERC_VALUES_", "MISSING_PERC_VALUES_"

        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If (InStr(1, rngName, "ROW") > 0) And RangeExists(rngName) Then
                Set rng = .Range(rngName)
                If (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And _
                   (Not newSection) Then
                    designFormat.ApplyFormat rng, scope:=AnalysisMissingRowWithoutHeader
                Else
                    designFormat.ApplyFormat rng, scope:=AnalysisMissingRow
                End If
            ElseIf RangeExists(rngName) Then
                Set rng = .Range(rngName)
                designFormat.ApplyFormat rng, scope:=AnalysisMissingColumn
            End If
        Next

        ' Single cells
        rngTable.Clear
        rngTable.Push "TOTAL_MISSING_", "TOTAL_TOTAL_", _
                      "USER_START_DATE_", "USER_END_DATE_", "TIME_UNIT_", _
                      "MISSING_TOTAL_", "MISSING_MISSING_", "START_DATE_LABEL_", _
                      "TIME_UNIT_LABEL_", "END_DATE_LABEL_", "POPFACT_", "POPFACTLABEL_", _
                      "INPUTSPTGEO_LABEL_"

        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If RangeExists(rngName) Then
                Set rng = .Range(rngName)
                If (InStr(1, rngName, "LABEL") > 0) Or (InStr(1, rngName, "POP") > 0) Then
                    designFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
                ElseIf (InStr(1, rngName, "USER") > 0) Or (rngName = "TIME_UNIT_" & tabId) Then
                    designFormat.ApplyFormat rng, scope:=AnalysisOneCellCenter
                Else
                    designFormat.ApplyFormat rng, scope:=AnalysisOneCell
                End If
            End If
        Next

        ' Info cells
        rngTable.Clear
        rngTable.Push "INFO_START_DATE_", "INFO_END_DATE_", "INFO_ANA_PERIOD_"
        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If RangeExists(rngName) Then
                Set rng = .Range(rngName)
                designFormat.ApplyFormat rng, scope:=AnalysisInfoCell
                If InStr(1, rngName, "INFO_ANA_") > 0 Then
                    designFormat.ApplyFormat rng, scope:=AnalysisInfoCellLeft
                End If
            End If
        Next

        ' Hidden cells
        rngTable.Clear
        rngTable.Push "PREVIOUS_ADM_", "POPPREVFACT_", "POPFACT_", "POPFACTLABEL_", _
                      "INPUTSPTGEO_NGEO_", "INPUTPREVADM_USER_"
        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If RangeExists(rngName) Then
                Set rng = .Range(rngName)
                designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
            End If
        Next

        ' Percentage columns ====================================================
        If hasPercent Then
            Set rng = .Range("PERC_COL_1_" & tabId)
            If (specs.TableScope <> ScopeSpatial) Or specs.HasTotal Then
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count + 1, 1))
            End If
            Do While rng.Column <= EndColumn()
                designFormat.ApplyFormat rng, scope:=AnalysisPercent
                Set rng = rng.Offset(, 2)
            Loop
            rngTable.Clear
            rngTable.Push "MISSING_MISSING_", "TOTAL_TOTAL_", "MISSING_TOTAL_"
            For counter = rngTable.LowerBound To rngTable.UpperBound
                rngName = rngTable.Item(counter) & tabId
                If RangeExists(rngName) Then
                    Set rng = .Range(rngName).Offset(, 1)
                    designFormat.ApplyFormat rng, scope:=AnalysisOneCellCenter
                End If
            Next
        End If

        ' Whole table ===========================================================
        wholeTableStart = StartColumn
        If (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And _
           (Not newSection) Then
            wholeTableStart = StartColumn + 1
        End If

        If (tablType <> ScopeUnivariate) And (tablType <> ScopeGlobalSummary) Then
            Set rng = .Range(.Cells(StartRow, wholeTableStart), .Cells(EndRow, EndColumn))
            designFormat.ApplyFormat rng, scope:=AnalysisWholeTable
        ElseIf tablType = ScopeUnivariate Then
            Set rng = .Range(.Cells(StartRow + 1, wholeTableStart), .Cells(EndRow, EndColumn))
            designFormat.ApplyFormat rng, scope:=AnalysisWholeTable
        End If

        ' End table gap for graphs
        If RangeExists("ENDTABLE_" & tabId) And _
           (tablType <> ScopeTimeSeries) And _
           (tablType <> ScopeSpatioTemporal) Then
            Set rng = .Range("ENDTABLE_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisEndTable
        End If

        ' Bug #1 fix: Hide Total COLUMN when not explicitly requested by user
        ' For time series/spatio-temporal tables, HasTotal can be True because
        ' percentage="row"/"column" needs the total column for computation.
        ' When TotalRequested=False, hide the column to keep it invisible.
        If (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And _
           specs.HasTotal And (Not specs.TotalRequested) Then
            If RangeExists("TOTAL_COL_" & tabId) Then
                .Range("TOTAL_COL_" & tabId).EntireColumn.Hidden = True
                ' Also hide the adjacent percentage column for Total if it exists
                If hasPercent And RangeExists("TOTAL_PERC_VALUES_" & tabId) Then
                    .Range("TOTAL_PERC_VALUES_" & tabId).EntireColumn.Hidden = True
                End If
            End If
        End If

    End With
End Sub

'@section Build
'===============================================================================
'@description Orchestrate the complete table construction sequence.

'@label:build
'@sub-title Build the complete cross-table on the worksheet.
'@details
'Orchestrates the complete construction of this cross-table by calling
'the four build steps in the correct order: AddHeader (inserts rows),
'AddRows (populates row categories and sets EndRow), AddColumns
'(populates column headers and sets EndColumn/NumberOfColumns), and
'NameRanges (creates all named ranges). The order is critical because
'AddHeader may insert rows that shift content down. After Build
'completes, the table is structurally complete and ready for Format
'and formula generation.
'@remarks Build does NOT call Format. Formatting is applied separately by the caller
'(typically the analysis builder) after all tables in a section have been built.
'@export
Private Sub Build()
    AddHeader
    AddRows
    AddColumns
    NameRanges
End Sub

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError-based exception.
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME,
'providing a consistent stack trace across all methods in this class.
'@param errNumber Long. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the ICrossTable contract.
'See the corresponding Public members above for full documentation.

Private Sub ICrossTable_AddRows()
    AddRows
End Sub

Private Sub ICrossTable_AddColumns()
    AddColumns
End Sub

Private Sub ICrossTable_AddHeader()
    AddHeader
End Sub

Private Sub ICrossTable_NameRanges()
    NameRanges
End Sub

Private Sub ICrossTable_Format(ByVal designFormat As ILLFormat)
    Format designFormat
End Sub

Private Sub ICrossTable_Build()
    Build
End Sub

Private Property Get ICrossTable_HeaderRange() As Range
    Set ICrossTable_HeaderRange = HeaderRange
End Property

Private Property Get ICrossTable_TimeSeriesHeader() As String
    ICrossTable_TimeSeriesHeader = TimeSeriesHeader()
End Property

Private Property Get ICrossTable_RowsCategoriesRange(Optional ByVal includeHeaders As Boolean = True, _
                                                      Optional ByVal onlyCategories As Boolean = False) As Range
    Set ICrossTable_RowsCategoriesRange = RowsCategoriesRange(includeHeaders, onlyCategories)
End Property

Private Property Get ICrossTable_ColumnRange(ByVal colVal As String, _
                                              Optional ByVal onlyCategories As Boolean = False, _
                                              Optional ByVal includeHeaders As Boolean = False) As Range
    Set ICrossTable_ColumnRange = ColumnRange(colVal, onlyCategories, includeHeaders)
End Property

Private Property Get ICrossTable_RowRange(ByVal rowVal As String, _
                                           Optional ByVal includeHeaders As Boolean = True) As Range
    Set ICrossTable_RowRange = RowRange(rowVal, includeHeaders)
End Property

Private Property Get ICrossTable_StartRow() As Long
    ICrossTable_StartRow = StartRow
End Property

Private Property Get ICrossTable_Previous() As ICrossTable
    Set ICrossTable_Previous = Previous
End Property

Private Property Get ICrossTable_Specifications() As ITableSpecs
    Set ICrossTable_Specifications = Specifications
End Property

Private Property Get ICrossTable_Wksh() As Worksheet
    Set ICrossTable_Wksh = Wksh()
End Property

Private Property Get ICrossTable_NumberOfColumns() As Long
    ICrossTable_NumberOfColumns = NumberOfColumns()
End Property

Private Property Get ICrossTable_EndColumn() As Long
    ICrossTable_EndColumn = EndColumn()
End Property

Private Property Get ICrossTable_EndRow() As Long
    ICrossTable_EndRow = EndRow()
End Property

Private Property Get ICrossTable_NamedRangesList() As BetterArray
    Set ICrossTable_NamedRangesList = NamedRangesList()
End Property
