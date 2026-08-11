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

'@class CrossTable
'@description
'Builds and formats one analysis output table on a worksheet. Six table scopes
'are handled: univariate, bivariate, time series, spatial, spatio-temporal and
'global summary. Build lays the table out and Format gives it its look.
'StartRow, EndRow, EndColumn, NumberOfColumns, HeaderRange, RowRange,
'ColumnRange, RowsCategoriesRange and TimeSeriesHeader answer where the table
'sits and what it holds. AnalysisOutput calls Build and then Format for each
'table specification.
'
'THE LINELIST DATA PARAMETER IS TYPED Object
'-------------------------------------------------------------------------------
'Create takes lData as an Object. The class asks it for TransObject and hands
'it on to TableSpecs.RowCategories and ColumnCategories, which type the same
'parameter as Object and dispatch on TypeName. AnalysisOutput passes a
'LinelistSpecs, and the Object type keeps the LinelistSpecs closure outside the
'compile dependencies of this class.
'
'THE FOUR BUILD PHASES RUN IN ONE ORDER
'-------------------------------------------------------------------------------
'AddHeader inserts rows and shifts the sheet, AddRows sets EndRow, AddColumns
'sets EndColumn and NumberOfColumns, and NameRanges reads all three. Build runs
'them in that order and records how far it has gone. NameRanges and Format
'refuse a table whose earlier phases have not run. Between AddHeader and
'AddRows the layout is fixed: the start row is held and the start column is
'written onto the sheet.
'
'EACH OF THE THREE LARGE PHASES IS A DISPATCHER
'-------------------------------------------------------------------------------
'AddHeader, NameRanges and Format each carry one Select Case over the table
'scope and hand the work to a small private routine, the way VarWriter does.
'
'WHAT IS RESOLVED ONCE PER TABLE
'-------------------------------------------------------------------------------
'The table identifier, the table scope, the range-name builder, the spatial
'type, the translation provider and the variable reader are read once and held
'for the life of the instance.
'@depends AnalysisRanges, TableSpecs, TranslationObject, LLVariables, LLFormat, HiddenNames, BetterArray, Checking
'@version 1.2 (2026-07-31)


Private Type TCrossTable
    specs As TableSpecs
    outputSh As Worksheet
    lData As Object
    endRw As Long
    endCol As Long
    nbCols As Long
    tsHeader As String                           'Header for time series (used in go-to section)
    hasChecks As Boolean
    sealed As Boolean
    tabId As String                              'Cached Specifications.TableId
    tabType As Byte                              'Cached Specifications.TableScope
    tabNames As AnalysisRanges                   'Range-name builder bound to tabId
    spatialType As String                        'Cached Specifications.SpatialTableScopes
    spatialTypeRead As Boolean
    trans As TranslationObject                   'Cached translation provider
    vars As LLVariables                          'Cached variable reader
    prevTable As CrossTable                      'The table this one continues, kept
    startRw As Long                              'Cached StartRow, valid after AddHeader
    startRwRead As Boolean
    phase As Byte                                'How far Build has gone
End Type

Private this As TCrossTable
Private Const STANDARD_START_COL As Long = 3
Private Const CLASS_NAME As String = "CrossTable"
Private Const NB_ROWS_TIME_SERIES As Long = 56

' The four build phases, in the order Build runs them. AddHeader inserts rows,
' AddRows sets EndRow, AddColumns sets EndColumn and NumberOfColumns, and
' NameRanges needs all three, so the order is load-bearing. RequirePhase is what
' says so to the compiler-less reader and to the caller that gets it wrong.
Private Const PHASE_HEADER As Byte = 1
Private Const PHASE_ROWS As Byte = 2
Private Const PHASE_COLUMNS As Byte = 3
Private Const PHASE_NAMES As Byte = 4

' The prefix lists Format walks. Each one was a BetterArray cleared and refilled
' per table, and each list is the same for every table of a scope, so they are
' written once here and split on use.
Private Const TEMPORAL_DATE_CELLS As String = _
    "START_DATE_|END_DATE_|VALIDATION_MIN_DATE_|" & _
    "VALIDATION_MAX_DATE_|MIN_MIN_DATE_|MAX_MAX_DATE_"
Private Const BAND_CELLS As String = _
    "TOTAL_ROW_VALUES_|TOTAL_COL_VALUES_|MISSING_ROW_VALUES_|" & _
    "MISSING_COL_VALUES_|TOTAL_PERC_VALUES_|MISSING_PERC_VALUES_"
Private Const SINGLE_CELLS As String = _
    "TOTAL_MISSING_|TOTAL_TOTAL_|USER_START_DATE_|USER_END_DATE_|TIME_UNIT_|" & _
    "MISSING_TOTAL_|MISSING_MISSING_|START_DATE_LABEL_|TIME_UNIT_LABEL_|" & _
    "END_DATE_LABEL_|POPFACT_|POPFACTLABEL_|INPUTSPTGEO_LABEL_"
Private Const INFO_CELLS As String = _
    "INFO_START_DATE_|INFO_END_DATE_|INFO_ANA_PERIOD_"
Private Const HIDDEN_CELLS As String = _
    "PREVIOUS_ADM_|POPPREVFACT_|POPFACT_|POPFACTLABEL_|" & _
    "INPUTSPTGEO_NGEO_|INPUTPREVADM_USER_"
Private Const PERCENT_CORNER_CELLS As String = _
    "MISSING_MISSING_|TOTAL_TOTAL_|MISSING_TOTAL_"

Private internalChecks As Checking

' Supplies the unique key every Checking.Add needs. It counts within one
' instance, which is all uniqueness requires: each instance owns its own
' Checking, and AnalysisOutput harvests them per table.
Private checkCounter As Long

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a CrossTable instance from table specifications.
'@details
'Factory method that creates and returns a new CrossTable instance from
'the given table specifications. A cross-table is a structured analysis
'output (univariate, bivariate, time series, spatial, spatio-temporal, or
'global summary) that is built on an output worksheet and contains row
'categories, column categories, named ranges, and formatting. This method
'validates all three required dependencies, then injects them into a
'fresh CrossTable instance via the PredeclaredId/New pattern and returns
'the interface reference.
'@param specs TableSpecs. An TableSpecs object representing one row from the analysis setup sheet.
'@param outputWksh Worksheet. The Excel Worksheet where this cross-table will physically write its headers, row labels, column labels, data placeholders, and named ranges.
'@param lData Object. The linelist specifications object that provides the translation object
'for locale-aware labels and is passed to TableSpecs.RowCategories / ColumnCategories for
'category resolution at build time. Production passes a LinelistSpecs. The type is Object so
'that this class carries no compile dependency on it; the two members it needs, TransObject
'and Categories, are resolved at run time.
'@param previousTable Optional CrossTable. The table this one continues, when the
'caller already holds it. A caller that walks the setup rows in order builds every
'ancestor anyway, so handing the last one in saves this instance rebuilding it.
'@return CrossTable. A fully initialized CrossTable ready for Build or individual step calls.
'@throws InvalidArgument When specs is Nothing.
'@throws InvalidArgument When outputWksh is Nothing.
'@throws InvalidArgument When lData is Nothing.
'@export
Public Function Create(ByVal specs As TableSpecs, _
                       ByVal outputWksh As Worksheet, _
                       ByVal lData As Object, _
                       Optional ByVal previousTable As CrossTable) As CrossTable

    If specs Is Nothing Then ThrowError InvalidArgument, "Table specifications not defined"
    If outputWksh Is Nothing Then ThrowError InvalidArgument, "Output worksheet not defined"
    If lData Is Nothing Then ThrowError InvalidArgument, "Linelist specifications not defined"

    Dim instance As CrossTable
    Set instance = New CrossTable
    Set instance.Specifications = specs
    Set instance.Wksh = outputWksh
    Set instance.LinelistData = lData
    If Not previousTable Is Nothing Then Set instance.PreviousTable = previousTable
    instance.Seal
    Set Create = instance
End Function

'@section Internal Properties
'===============================================================================
'@description Properties used during factory instantiation and internal wiring.

'@label:seal
'@sub-title Seal the instance against further setup writes.
'@details
'Also resolves the two facts every phase of the build reads first -- the table
'identifier and the table scope -- and binds the range-name builder to the
'identifier. Neither can change while this instance lives, so each is read once
'here instead of once per phase.
Public Sub Seal()
    this.sealed = True
    this.tabId = this.specs.TableId
    this.tabType = this.specs.TableScope
    Set this.tabNames = AnalysisRanges.Create(this.tabId)
End Sub

'@label:previous-table-set
'@prop-title Assign the table this one continues.
'@param tabl CrossTable. The preceding cross-table.
Public Property Set PreviousTable(ByVal tabl As CrossTable)
    GuardNotSealed "PreviousTable"
    Set this.prevTable = tabl
End Property

'@label:guard-not-sealed
'@sub-title Guard a setup setter against post-creation writes.
'@param propName String. The property name for the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

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
    GuardNotSealed "Wksh"
    Set this.outputSh = sh
End Property

'@label:specifications
'@prop-title Table specifications used to build this cross-table.
'@details
'Returns the TableSpecs instance that defines the configuration for
'this cross-table. Provides access to table scope, variable names,
'category lists, and validation flags.
'@return TableSpecs. The backing table specification.
'@export
Public Property Get Specifications() As TableSpecs
    Set Specifications = this.specs
End Property

'@label:specifications-set
'@prop-title Assign the table specifications.
'@param specs TableSpecs. The table specification to assign.
Public Property Set Specifications(ByVal specs As TableSpecs)
    GuardNotSealed "Specifications"
    Set this.specs = specs
End Property

'@label:linelist-data
'@prop-title Linelist specifications providing translations and categories.
'@details
'Returns the linelist specifications object stored at construction time. Used
'to extract the translation object and to pass as the lData parameter when
'calling specs.RowCategories(lData) and specs.ColumnCategories(lData).
'@return Object. The linelist specifications.
Public Property Get LinelistData() As Object
    Set LinelistData = this.lData
End Property

'@label:linelist-data-set
'@prop-title Assign the linelist specifications.
'@param lData Object. The linelist specifications to assign.
Public Property Set LinelistData(ByVal lData As Object)
    GuardNotSealed "LinelistData"
    Set this.lData = lData
End Property

'@label:translations
'@prop-title Translation object for locale-aware labels.
'@details
'Returns the TranslationObject extracted from the stored linelist
'specifications. Resolves translated strings such as "Total", "Missing",
'"Percent", and date-unit names.
'@return TranslationObject. The translation object.
Public Property Get Translations() As TranslationObject
    If this.trans Is Nothing Then Set this.trans = this.lData.TransObject()
    Set Translations = this.trans
End Property

'@label:number-of-columns
'@prop-title Number of data columns in the table.
'@details
'Returns the count of data columns excluding the row-label column.
'For bivariate and time series tables this equals the number of column
'categories plus optional total and missing columns.
'@return Long. The data column count.
Public Property Get NumberOfColumns() As Long
    NumberOfColumns = this.nbCols
End Property

'@label:number-of-columns-set
'@prop-title Assign the number of data columns.
'@param col Long. The data column count to store.
Private Property Let NumberOfColumns(ByVal col As Long)
    this.nbCols = col
End Property

'@section Cached Facts
'===============================================================================
'@description
'Every phase of a build opened by re-deriving the same handful of facts, and
'none of them can change while one instance lives. They are resolved once and
'held here, the way VarWriter holds its collaborators.

'@label:table-id
'@prop-title Identifier of the table being built.
'@return String. The table identifier, such as "TS_tab3".
Private Property Get TableId() As String
    TableId = this.tabId
End Property

'@label:table-scope
'@prop-title Analysis scope of the table being built.
'@return Byte. A member of AnalysisTableScope.
Private Property Get TableScope() As Byte
    TableScope = this.tabType
End Property

'@label:range-names
'@prop-title Range-name builder bound to this table's identifier.
'@return AnalysisRanges. The builder.
Private Property Get RangeNames() As AnalysisRanges
    Set RangeNames = this.tabNames
End Property

'@label:spatial-type
'@prop-title Whether the spatial variable is geographic or a health facility.
'@details
'Resolved on first use rather than at creation, because the answer costs a
'dictionary probe and the four non-spatial scopes never ask for it. An empty
'answer is a real answer here, so a separate flag records that the probe ran.
'@return String. "geo", "hf", or an empty string.
Private Property Get SpatialType() As String
    If Not this.spatialTypeRead Then
        this.spatialType = this.specs.SpatialTableScopes()
        this.spatialTypeRead = True
    End If
    SpatialType = this.spatialType
End Property

'@label:variables
'@prop-title Variable reader over the dictionary of this build.
'@details
'AddRows and AddColumns each built one of these, and TableSpecs builds a third
'for every category lookup. One per table is enough.
'@return LLVariables. The variable reader.
Private Property Get Variables() As LLVariables
    If this.vars Is Nothing Then Set this.vars = LLVariables.Create(this.specs.Dictionary)
    Set Variables = this.vars
End Property

'@label:value-of
'@prop-title One setup column of the row this table is built from.
'@details
'Every read of the setup row goes through here. TableSpecs.Value scans its
'cached header names, so the cost is one linear scan per call, and putting the
'twenty-odd call sites behind one name is what makes a per-table memo a single
'edit later.
'@param colName String. The setup column to read, matched loosely.
'@return String. The cell value, or an empty string.
Private Property Get ValueOf(ByVal colName As String) As String
    ValueOf = this.specs.Value(colName)
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
'@remarks Every name this class writes comes from Cell.Name, which produces a
'WORKBOOK-scoped name, and Wksh().Range(name) resolves one of those whichever
'sheet owns it. So the question "does this name exist" and the question this
'class actually asks -- "did I already build that on MY sheet" -- were the same
'question only while no two analysis sheets used a name in common. The answer is
'now filtered on the owning worksheet. The names without a table id are where it
'mattered: ROWGS_SET, COLGS_SET, TIME_UNIT_LIST, ADM_UNIT_LIST and
'POPULATION_FACTOR_LIST are each written on one sheet and read from four.
Private Function RangeExists(ByVal rngName As String) As Boolean
    RangeExists = (Not NamedRange(rngName) Is Nothing)
End Function

'@label:named-range
'@sub-title Resolve a named range on the output worksheet, or Nothing.
'@details
'The answer to "does this name exist" and the answer to "give me the range"
'were two separate name resolutions at every call site of RangeExists that went
'on to read the range. Format alone did that about fifty times per table. This
'answers both questions with one resolution.
'@param rngName String. The name to look up.
'@return Range. The range, or Nothing when the name is absent or owned by
'another worksheet.
'@remarks The owning sheet is compared by NAME. Wksh() hands back the stored
'reference while Range.Worksheet builds a fresh wrapper over the same sheet, and
'Is compares wrapper identity, so it can answer False for the very sheet it was
'handed. An Is answering False for the right sheet makes this function answer
'Nothing everywhere, and every "did I already build this" guard then rebuilds:
'AddHeader re-inserts its section rows on each call and StartRow recomputes from
'End(xlUp) against a sheet that just grew. The table walks down the sheet and the
'run never finishes.
Private Function NamedRange(ByVal rngName As String) As Range
    Dim rng As Range
    Dim sh As Worksheet

    Set sh = Wksh()

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then Exit Function
    If StrComp(rng.Worksheet.Name, sh.Name, vbTextCompare) <> 0 Then Exit Function

    Set NamedRange = rng
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
'StartColumn, AddRows, AddColumns, AddHeader, NameRanges, Format). Build resolves
'it once, straight after AddHeader, and every later read is a field read.
'
'WHY THE CACHE IS FILLED AFTER AddHeader AND NOT AT CREATION
'-------------------------------------------------------------------------------
'AddHeader inserts rows above the table, which moves it down the sheet. Excel
'moves the STARTROW_ name with it; a Long held from before the insert would be
'stale. Nothing inserts rows above a table once its own header is written, so
'from that point on the value is fixed for the life of the instance.
Public Property Get StartRow() As Long
    If this.startRwRead Then
        StartRow = this.startRw
    Else
        StartRow = ComputeStartRow()
    End If
End Property

'@label:cache-start-row
'@sub-title Resolve the start row once and hold it.
Private Sub CacheStartRow()
    this.startRw = ComputeStartRow()
    this.startRwRead = True
End Sub

'@label:compute-start-row
'@sub-title Read or compute the start row from the worksheet.
'@return Long. The 1-based row this table begins at.
'@throws ErrorUnexpectedState When no branch could resolve a row.
Private Function ComputeStartRow() As Long
    Dim sh As Worksheet
    Dim specs As TableSpecs
    Dim strtRow As Long
    Dim tabId As String
    Dim startRowName As String
    Dim rng As Range

    Set specs = Specifications()
    tabId = TableId()
    startRowName = "STARTROW_" & tabId
    Set sh = Wksh()

    Set rng = NamedRange(startRowName)
    If Not rng Is Nothing Then
        ComputeStartRow = rng.Row
        Exit Function
    End If

    With sh
        If specs.IsNewSection() Or _
           (TableScope() <> ScopeTimeSeries And TableScope() <> ScopeSpatioTemporal) Then
            ' New section or non-temporal: compute from last used row
            strtRow = .Cells(.Rows.Count, STANDARD_START_COL).End(xlUp).Row + 1
            strtRow = IIf(strtRow = 2, 6, strtRow)

            strtRow = Switch(TableScope() = ScopeSpatial, strtRow + 6, _
                             TableScope() = ScopeGlobalSummary, strtRow, _
                             True, strtRow + 4)

            .Cells(strtRow, STANDARD_START_COL).Name = startRowName

        Else
            ' Non-new-section temporal: use previous table's start row.
            ' Previous answers Nothing when the section holds no valid
            ' table above this row, so the condition is read directly.
            ' The previous table also has to have been built.
            Dim prevSpecs As TableSpecs
            Dim prevTabId As String

            Set prevSpecs = specs.Previous
            If Not prevSpecs Is Nothing Then prevTabId = prevSpecs.TableId

            If prevTabId <> vbNullString And RangeExists("ENDTABLE_" & prevTabId) Then
                strtRow = Previous.StartRow
                ' The inherited branch is the one that pays for the ancestor
                ' walk, and it was the one branch that never wrote the answer
                ' down. Naming the cell here makes every later read of this
                ' table's start row a lookup, the way the two fresh branches
                ' already are.
                .Cells(strtRow, STANDARD_START_COL).Name = startRowName
            Else
                ' Previous table was never built (invalid); treat as de-facto new section
                strtRow = .Cells(.Rows.Count, STANDARD_START_COL).End(xlUp).Row + 1
                strtRow = IIf(strtRow = 2, 6, strtRow)
                strtRow = strtRow + 4
                .Cells(strtRow, STANDARD_START_COL).Name = startRowName
            End If
        End If

        ' The two branches above are exhaustive today, so this cannot fire.
        ' It is here because the second one used to be written as an ElseIf
        ' repeating the first one's negation: a scope added to
        ' AnalysisTableScope would have fallen past both arms, left strtRow
        ' at zero, and handed every consumer row 0 to build from. That surfaces
        ' as Cells(0, ...) raising 1004 several frames away from the cause.
        If strtRow < 1 Then _
            ThrowError ErrorUnexpectedState, _
                       "Start row could not be resolved for " & tabId

        ComputeStartRow = strtRow
    End With
End Function

'@label:previous
'@prop-title Previous cross-table in the same section.
'@details
'Creates and returns an CrossTable instance representing the table
'that immediately precedes this one within the same analysis section.
'This is used by temporal table scopes (time series and spatio-temporal)
'where multiple tables share the same row structure and are laid out
'side-by-side in columns. The previous table's start row and end column
'are needed to position the current table correctly. The method delegates
'to the Create factory, passing the Previous spec from the current
'table's TableSpecs, along with the same output worksheet and
'translation object.
'@return CrossTable. The preceding cross-table, or Nothing when this is the first table.
'@remarks The instance is built once and held.
'
'WHY THIS PROPERTY USED TO COST O(n) PER CALL
'-------------------------------------------------------------------------------
'It built a fresh CrossTable every time it was read, and StartRow reads
'Previous.StartRow, which reads Previous again. For a temporal section of n
'tables, resolving the last table's start row constructed O(n) cross-tables and
'walked O(n) specification chains, and StartRow is read from six places. Holding
'the instance collapses that to one construction per table.
'
'AND WHY A HANDED-IN ANCESTOR IS CHECKED BEFORE IT IS USED
'-------------------------------------------------------------------------------
'AnalysisOutput walks the setup rows in order and keeps the last table it built,
'so it can hand the ancestor in and save this instance building it again. The two
'can disagree: a row that passes validation and then fails inside Build is still
'the nearest valid row above, so the caller's last table is one further up. The
'identifiers are compared and a mismatch falls back to building the real one.
Public Property Get Previous() As CrossTable
    Dim prevSpecs As TableSpecs

    ' Create rejects a Nothing specification, so the first table of a section
    ' is answered here instead of raising InvalidArgument out of the factory.
    Set prevSpecs = Specifications.Previous
    If prevSpecs Is Nothing Then
        Set this.prevTable = Nothing
        Exit Property
    End If

    If Not this.prevTable Is Nothing Then
        If StrComp(this.prevTable.Specifications.TableId, prevSpecs.TableId, _
                   vbBinaryCompare) = 0 Then
            Set Previous = this.prevTable
            Exit Property
        End If
    End If

    Set this.prevTable = Create(prevSpecs, Wksh(), LinelistData())
    Set Previous = this.prevTable
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
'
'THIS PROPERTY WRITES NOTHING
'-------------------------------------------------------------------------------
'It used to name STARTCOL_ on the worksheet, so reading it from a consumer turned
'into writing names onto tables the caller had not built -- and, through StartRow,
'into walking the whole ancestor chain. EnsureStartColumn is the assignment half
'and Build calls it once, straight after AddHeader. A read before that answers
'what the assignment would store.
Private Property Get StartColumn() As Long
    Dim rng As Range

    Set rng = NamedRange("STARTCOL_" & TableId())

    If Not rng Is Nothing Then
        StartColumn = rng.Column - 1
    Else
        StartColumn = ComputeStartColumn()
    End If
End Property

'@label:compute-start-column
'@sub-title Work out which column this table starts in.
'@details
'A temporal table continuing a section sits to the right of the one before it,
'which is what the End(xlToLeft) scan finds. Every other table, and the table
'that opens a temporal section, starts at the standard column.
'@return Long. The 1-based start column.
Private Function ComputeStartColumn() As Long
    Dim sh As Worksheet

    Set sh = Wksh()

    If ((TableScope() = ScopeTimeSeries) Or (TableScope() = ScopeSpatioTemporal)) And _
       (Not Specifications.IsNewSection()) Then
        ComputeStartColumn = sh.Cells(StartRow + 1, sh.Columns.Count).End(xlToLeft).Column
    Else
        ComputeStartColumn = STANDARD_START_COL
    End If
End Function

'@label:ensure-start-column
'@sub-title Fix this table's start column on the worksheet.
'@details
'Writes STARTCOL_<tableId> once. It has to run before AddColumns, because the
'End(xlToLeft) scan behind a continuation table's column answers something else
'as soon as that table's own columns are on the sheet, and because AddColumns
'writes the column label into the named cell. Build calls it straight after
'AddHeader; AddRows and AddColumns call it too, so a caller driving the phases
'one at a time gets the same sheet.
Private Sub EnsureStartColumn()
    Dim sh As Worksheet
    Dim strtCol As Long

    If RangeExists("STARTCOL_" & TableId()) Then Exit Sub

    Set sh = Wksh()
    strtCol = ComputeStartColumn()
    sh.Cells(StartRow - 1, strtCol + 1).Name = "STARTCOL_" & TableId()
End Sub

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
Public Property Get EndRow() As Long
    EndRow = this.endRw
End Property

'@label:end-column
'@prop-title Last column used by the table on the worksheet.
'@details
'Returns the 1-based column index of the rightmost cell occupied by
'this cross-table. Used by sibling tables and formatting routines to
'avoid overlapping output.
'
'THE CONTRACT, PINNED HERE BECAUSE IT IS DECIDED HERE
'-------------------------------------------------------------------------------
'EndColumn is INCLUSIVE and it is the last column WRITTEN. The total and missing
'columns count toward it. NumberOfColumns stays outside it and stops at the data
'columns. Every assignment below obeys this -- StartColumn + labData.Length for the
'bivariate and temporal arms, StartColumn + 2 for the global summary pair,
'StartColumn + colData.Length for the spatial arm -- because labData and colData
'already carry whatever total and missing labels the table asked for.
'
'So a reader walking the value columns wants `< EndColumn` only if it means to
'stop before the totals, and `<= EndColumn` if it means to include them. The
'formula writer currently uses `< tabEndColumn` for bivariate and
'`<= tabEndColumn` for time series and spatial, which cannot be right for all
'three against one definition. The loops live in CrossTableFormula and are
'fixed there. Stating the contract here is the half that belongs to this class.
'@return Long. The ending column index, inclusive of total and missing columns.
Public Property Get EndColumn() As Long
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
Public Property Get TimeSeriesHeader() As String
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
    Dim specs As TableSpecs
    Dim tablType As Byte

    Set specs = Specifications
    tablType = TableScope()
    If tablType = ScopeGlobalSummary Then ThrowError ErrorUnexpectedState, "Global summary has no row categories"

    Set sh = Wksh()

    Select Case tablType
    Case ScopeBivariate, ScopeTimeSeries, ScopeSpatial, ScopeSpatioTemporal
        startRw = StartRow + 1
    Case ScopeUnivariate
        startRw = StartRow
    Case Else
        ' StartRow carries the same guard and says why. Without it a scope
        ' outside the five leaves startRw at zero and .Cells(0, 3) raises 1004
        ' several frames from the cause. ScopeTimeSeriesGraph is a real member
        ' of AnalysisTableScope, so the arm is reachable.
        ThrowError ErrorUnexpectedState, _
                   "Row categories were asked of table " & TableId() & _
                   ", whose scope this class does not lay out"
    End Select

    If Not includeHeaders Then startRw = startRw + 1

    endRw = Switch( _
        (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And onlyCategories, EndRow - 2, _
        (specs.HasTotal Or tablType = ScopeSpatial) And onlyCategories, EndRow - 1, _
        True, EndRow)

    ' STANDARD_START_COL is the right column for every scope, and this is the
    ' place the reason is written down. The audit's item 47 asked for
    ' StartColumn here, on the reading that a temporal table continuing a
    ' section starts further right. It does -- and its row labels do not.
    ' AddRows writes the Period, Total and Missing labels of a temporal section
    ' ONCE, from the table that opens the section, into column 3, and every
    ' later table of that section shares them. RowRange is the only caller this
    ' property has on the continuation path, and what it looks for is exactly
    ' those shared labels. Reading them in the continuation's own column finds
    ' empty cells, so TOTAL_ROW_ and MISSING_ROW_ would stop being created and
    ' every total formula of a continuation table would go with them.
    '
    ' The other path agrees: NameRanges derives a continuation table's category
    ' range from its own ENDTABLE_ marker, which sits in the continuation's
    ' column, and the value columns hang off that. So the two paths use two
    ' columns on purpose.
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
    Dim startCol As Long
    Dim sh As Worksheet

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
    Dim specs As TableSpecs
    Dim tablType As Byte

    Set rng = HeaderRange()
    Set sh = Wksh()
    Set specs = Specifications
    tablType = TableScope()

    startRw = Switch(tablType = ScopeBivariate, StartRow + 2, _
                     (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal), StartRow + 2, _
                     tablType = ScopeSpatial, StartRow + 2, _
                     True, StartRow + 1)

    endRw = Switch( _
        (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And onlyCategories, EndRow - 2, _
        (specs.HasTotal Or tablType = ScopeSpatial) And onlyCategories, EndRow - 1, _
        True, EndRow)

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
Public Sub AddRows()
    Dim startRw As Long
    Dim startCol As Long
    Dim specs As TableSpecs
    Dim sh As Worksheet
    Dim vars As LLVariables
    Dim rowVar As String
    Dim rowsData As BetterArray
    Dim trans As TranslationObject
    Dim transMissing As String
    Dim transTotal As String
    Dim counter As Long
    Dim spatType As String

    EnsureStartColumn
    startRw = StartRow
    startCol = StartColumn
    Set sh = Wksh()
    Set specs = Specifications()
    Set trans = Translations()
    Set vars = Variables()

    transMissing = trans.TranslatedValue("MSG_NA")
    transTotal = trans.TranslatedValue("MSG_Total")

    ' The categories are read inside the two arms that use them, and that is a
    ' correctness fix rather than a saving.
    '
    ' THE ROW VARIABLE OF A TEMPORAL OR SPATIAL TABLE HAS NO CATEGORIES
    ' ---------------------------------------------------------------------------
    ' This used to read them for every scope, above the dispatch, and the
    ' temporal and spatial arms then threw the answer away. A temporal table's
    ' row variable is a DATE -- ValidTable requires it -- and a date exposes no
    ' selectable categories, so LinelistSpecs.Categories raises ElementNotFound
    ' for it and TableSpecs passes that on. Every temporal table therefore died
    ' inside Build, landed on the per-table handler in AnalysisOutput, and was
    ' skipped with one line in the report. The header of the section was already
    ' on the sheet by then, because AddHeader runs first, so what a user saw was
    ' a section banner with no table under it.
    '
    ' The two arms that do read the categories are the two whose row variable
    ' ValidTable already requires to be a choice.
    Set rowsData = New BetterArray
    rowsData.LowerBound = 1

    With sh
        Select Case TableScope()

        Case ScopeGlobalSummary
            .Cells(startRw, startCol).Value = ValueOf("label")

        Case ScopeUnivariate
            Set rowsData = specs.RowCategories(LinelistData())
            If ValueOf("missing") = "yes" Then rowsData.Push transMissing
            rowsData.Push transTotal
            rowsData.ToExcelRange .Cells(startRw + 1, startCol)
            rowVar = ValueOf("row")
            .Cells(startRw, startCol).Value = vars.Value(colName:="main label", varName:=rowVar)
            EndRow = startRw + rowsData.Length

        Case ScopeBivariate
            Set rowsData = specs.RowCategories(LinelistData())
            If ValueOf("missing") = "row" Or ValueOf("missing") = "all" Then
                rowsData.Push transMissing
            End If
            rowsData.Push transTotal
            rowsData.ToExcelRange .Cells(startRw + 2, startCol)
            rowVar = ValueOf("row")
            .Cells(startRw + 1, startCol).Value = vars.Value(colName:="main label", varName:=rowVar)
            EndRow = startRw + 1 + rowsData.Length

        Case ScopeTimeSeries, ScopeSpatioTemporal
            ' Add rows only on new sections (or de-facto new sections for Bug #3)
            If specs.IsNewSection() Then WriteTemporalRowGrid startRw, startCol
            EndRow = startRw + NB_ROWS_TIME_SERIES

        Case ScopeSpatial
            rowVar = ValueOf("row")
            spatType = SpatialType()

            If spatType = "geo" Then
                rowsData.Push vars.Value(colName:="sub section", varName:="adm1_" & rowVar)
            ElseIf spatType = "hf" Then
                rowsData.Push vars.Value(colName:="main label", varName:="hf_" & rowVar)
            End If

            ' GeoCount carries the clamp and the default, so an empty
            ' "N geo max" cell no longer reaches CLng("").
            For counter = 1 To specs.GeoCount()
                rowsData.Push vbNullString
            Next

            rowsData.Push transMissing
            rowsData.ToExcelRange .Cells(startRw + 1, startCol)
            EndRow = startRw + rowsData.Length

        End Select
    End With

    MarkPhase PHASE_ROWS
End Sub

'@label:write-temporal-row-grid
'@sub-title Write the fixed row block a temporal section shares.
'@details
'The block is NB_ROWS_TIME_SERIES rows tall: a Period label at the top, blank
'rows for the periods themselves, then Total and Missing. It used to be built by
'pushing 52 empty strings one at a time into a BetterArray, in a loop whose
'bounds had to be kept in step with the constant by hand. The size comes from the
'constant now and the block reaches the sheet in one crossing.
'
'THE TOTAL AND MISSING LABELS ARE ALWAYS WRITTEN
'-------------------------------------------------------------------------------
'RowsCategoriesRange trims the two footer rows BY COUNT, so dropping a label
'would move every category range of the section up by one row and take every
'total formula with it. Whether the user SEES the Total row is decided in Format,
'which hides it when no table of the section asked for one -- issue #338.
'@param startRw Long. The row this table starts at.
'@param startCol Long. The column the shared labels go in.
Private Sub WriteTemporalRowGrid(ByVal startRw As Long, ByVal startCol As Long)
    Dim gridValues() As Variant
    Dim counter As Long
    Dim trans As TranslationObject
    Dim sh As Worksheet

    Set trans = Translations()
    Set sh = Wksh()

    ReDim gridValues(1 To NB_ROWS_TIME_SERIES, 1 To 1)
    For counter = 1 To NB_ROWS_TIME_SERIES
        gridValues(counter, 1) = vbNullString
    Next

    gridValues(1, 1) = trans.TranslatedValue("MSG_Period")
    gridValues(NB_ROWS_TIME_SERIES - 1, 1) = trans.TranslatedValue("MSG_Total")
    gridValues(NB_ROWS_TIME_SERIES, 1) = trans.TranslatedValue("MSG_NA")

    sh.Range(sh.Cells(startRw + 1, startCol), _
             sh.Cells(startRw + NB_ROWS_TIME_SERIES, startCol)).Value = gridValues
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
'@param specs TableSpecs. The TableSpecs providing column categories and flags.
'@param trans TranslationObject. The TranslationObject providing translated strings.
'@param sh Worksheet. The output Worksheet where the column headers are written.
'@remarks NumberOfColumns counts only the "pure" data columns, not including Missing
'or Total. However, EndColumn includes all columns written.
Private Sub BivariateColumns(ByVal startRw As Long, ByVal specs As TableSpecs, _
                              ByVal trans As TranslationObject, ByVal sh As Worksheet)
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

    percentageValue = ValueOf("percentage")
    transMissing = trans.TranslatedValue("MSG_NA")
    transTotal = trans.TranslatedValue("MSG_Total")
    transPercent = trans.TranslatedValue("MSG_Percent")

    With sh
        colLab = ValueOf("label")
        If specs.HasMissing Then labData.Push transMissing
        If specs.HasTotal Then labData.Push transTotal
        Set colData = specs.ColumnCategories(LinelistData())

        If colData.Length > 0 Then
            NumberOfColumns = colData.Length

            colData.CONCAT labData.Items
            labData.Clear

            If specs.HasPercentage Then
                ' The column list arrives from the linelist data object, which
                ' this class holds as an Object and cannot type-check, so the
                ' walk reads the bounds the supplier set rather than assuming
                ' the list starts at one. The sibling branch below is already
                ' bound-agnostic because Clone carries the bounds over.
                For counter = colData.LowerBound To colData.UpperBound
                    newColData.Push colData.Item(counter), vbNullString
                    labData.Push colLab, PercentLabel(transPercent, percentageValue)
                Next
            Else
                For counter = colData.LowerBound To colData.UpperBound
                    labData.Push colLab
                Next
                Set newColData = colData.Clone()
            End If

            newColData.ToExcelRange .Cells(startRw, StartColumn + 1), TransposeValues:=True
            labData.ToExcelRange .Cells(startRw + 1, StartColumn + 1), TransposeValues:=True
        Else
            ' No column data -- just a summary function. The Missing and Total
            ' labels collected above head columns cut from the column
            ' categories, and a table grouped by nothing has none of those, so
            ' they are dropped here rather than written as column headers. The
            ' arm above drops them the same way once they have been folded into
            ' the category list.
            labData.Clear
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
Public Sub AddColumns()
    Dim startRw As Long
    Dim specs As TableSpecs
    Dim sh As Worksheet
    Dim vars As LLVariables
    Dim trans As TranslationObject
    Dim colData As BetterArray
    Dim colVarLab As String
    Dim sumLab As String
    Dim timeVarLab As String
    Dim tabId As String
    Dim tablType As Byte
    Dim spType As String
    Dim percentageValue As String
    Dim transPercent As String

    EnsureStartColumn
    startRw = StartRow
    Set sh = Wksh()
    Set specs = Specifications()
    Set trans = Translations()
    Set vars = Variables()
    tabId = TableId()
    tablType = TableScope()
    Set colData = New BetterArray
    colData.LowerBound = 1
    percentageValue = ValueOf("percentage")
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
            colLab = ValueOf("label")
            colData.Push colLab
            If percentageValue = "yes" Then colData.Push transPercent
            colData.ToExcelRange .Cells(startRw, StartColumn + 1), TransposeValues:=True
            EndColumn = StartColumn + colData.Length
            NumberOfColumns = 1

        Case ScopeBivariate
            BivariateColumns startRw, specs, trans, sh
            .Range("STARTCOL_" & tabId).Value = vars.Value( _
                colName:="main label", varName:=ValueOf("column"))

        Case ScopeTimeSeries, ScopeSpatioTemporal
            BivariateColumns startRw, specs, trans, sh

            spType = SpatialType()

            If tablType = ScopeTimeSeries Then
                colVarLab = vars.Value(colName:="main label", varName:=ValueOf("column"))
            ElseIf spType = "geo" Then
                colVarLab = vars.Value(colName:="sub section", varName:="adm1_" & ValueOf("column"))
            ElseIf spType = "hf" Then
                colVarLab = vars.Value(colName:="main label", varName:="hf_" & ValueOf("column"))
            End If

            timeVarLab = vars.Value(colName:="main label", varName:=ValueOf("time"))
            sumLab = ValueOf("label")

            If ValueOf("column") = vbNullString Then
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

    MarkPhase PHASE_COLUMNS
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
'spatial tables create admin-unit or health-facility dropdowns.
'@remarks Row insertions shift the entire sheet down, which is why AddHeader must run
'before AddRows, AddColumns, and NameRanges. For temporal tables, the date
'controls are unlocked (Locked = False) so the user can interactively change
'dates at runtime.
'@export
Public Sub AddHeader()
    Select Case TableScope()
    Case ScopeUnivariate, ScopeBivariate
        AddTitleAndSectionHeader
    Case ScopeGlobalSummary
        AddGlobalSummaryHeader
    Case ScopeTimeSeries, ScopeSpatioTemporal
        AddTemporalHeader
    Case ScopeSpatial
        AddSpatialHeader
    End Select

    MarkPhase PHASE_HEADER
End Sub

'@label:add-title-and-section-header
'@sub-title Write the title of a univariate or bivariate table.
'@details
'A table that opens a section also gets three rows inserted above it for the
'section label. The insert shifts the sheet down and Excel carries the names
'with it, which is why the local start row is read before the insert and the
'section cell is addressed through the title name afterwards.
Private Sub AddTitleAndSectionHeader()
    Dim sh As Worksheet
    Dim startRw As Long
    Dim tabId As String
    Dim counter As Long
    Dim sectionRng As Range

    Set sh = Wksh()
    startRw = StartRow
    tabId = TableId()

    With sh
        .Cells(startRw - 2, STANDARD_START_COL).Value = ValueOf("title")
        .Cells(startRw - 2, STANDARD_START_COL).Name = "TITLE_" & tabId

        If Specifications.IsNewSection() And (Not RangeExists("SECTION_" & tabId)) Then
            For counter = 1 To 3
                .Rows(startRw - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Next
            Set sectionRng = .Range("TITLE_" & tabId).Cells(-1, 1)
            sectionRng.Value = ValueOf("section")
            sectionRng.Name = "SECTION_" & tabId
        End If
    End With
End Sub

'@label:add-global-summary-header
'@sub-title Write the shared global summary label, once per sheet.
'@details
'Every global summary row of a sheet shares one label, so ROWGS_SET is both the
'name of the cell and the guard that stops the second row writing it again.
Private Sub AddGlobalSummaryHeader()
    Dim sh As Worksheet
    Dim startRw As Long
    Dim counter As Long

    If RangeExists("ROWGS_SET") Then Exit Sub

    Set sh = Wksh()
    startRw = StartRow

    With sh
        For counter = 1 To 3
            .Rows(startRw - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next
        .Cells(startRw - 2, STANDARD_START_COL).Value = Translations().TranslatedValue("MSG_GlobalSummary")
        .Cells(startRw - 2, STANDARD_START_COL).Name = "ROWGS_SET"
    End With
End Sub

'@label:add-temporal-header
'@sub-title Open a temporal section and write the controls it carries.
'@details
'Runs for the table that opens the section only. A spatio-temporal section needs
'room for its geographic input rows on top of the date controls, which is what
'the taller insert is for.
Private Sub AddTemporalHeader()
    Dim sh As Worksheet
    Dim startRw As Long
    Dim tabId As String
    Dim counter As Long
    Dim nGeo As Long
    Dim nbShift As Long
    Dim anchorRng As Range

    If Not Specifications.IsNewSection() Then Exit Sub
    If RangeExists("SECTION_" & TableId()) Then Exit Sub

    Set sh = Wksh()
    startRw = StartRow
    tabId = TableId()

    With sh
        If TableScope() = ScopeSpatioTemporal Then
            ' GeoCount carries the default of 5 and clamps to 1..20, so
            ' an empty cell no longer raises out of CLng.
            nGeo = Specifications.GeoCount()
            nbShift = nGeo + 14
        Else
            nbShift = 7
        End If

        For counter = 1 To nbShift
            .Rows(startRw - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next

        .Cells(startRw - 2, StartColumn).Value = ValueOf("section")
        .Cells(startRw - 2, StartColumn).Name = "SECTION_" & tabId

        Set anchorRng = .Range("SECTION_" & tabId)
    End With

    If TableScope() = ScopeSpatioTemporal Then _
        Set anchorRng = AddSpatioTemporalGeoInputs(anchorRng, nGeo)

    AddTemporalDateControls anchorRng
End Sub

'@label:add-spatio-temporal-geo-inputs
'@sub-title Write the geographic input rows of a spatio-temporal section.
'@details
'One input cell per geographic unit, each named so the formula writer can point
'at it. The tag comes from the dictionary probe, which is what the formula
'writer uses to reference these same cells. It used to come from the "spatial
'type" column of this row, and that column is never filled: the setup workbook
'validates it on the spatio-temporal specification table and the analysis table
'only ever receives the geo dropdown. So the read returned empty, every table
'was tagged geographic, and a health facility table's label formulas pointed at
'INPUTSPTHF_ names that were never created.
'@param sectionRng Range. The section anchor cell.
'@param nGeo Long. How many geographic units the table shows.
'@return Range. The anchor the date controls hang from.
Private Function AddSpatioTemporalGeoInputs(ByVal sectionRng As Range, _
                                            ByVal nGeo As Long) As Range
    Dim rng As Range
    Dim cellRng As Range
    Dim counter As Long
    Dim tabId As String
    Dim sh As Worksheet

    Set sh = Wksh()
    tabId = TableId()

    Set rng = sectionRng.Offset(4)
    rng.Value = Translations().TranslatedValue("MSG_InputGeoLevels")
    rng.Name = "INPUTSPTGEO_LABEL_" & tabId
    rng.Cells(1, 2).Value = nGeo
    rng.Cells(1, 2).Name = "INPUTSPTGEO_NGEO_" & tabId

    Set cellRng = rng.Offset(2)
    For counter = 1 To nGeo
        cellRng.Cells(counter, 1).Name = RangeNames().SpatialInput(counter, SpatialType())
        cellRng.Cells(counter, 2).Value = 1
    Next

    Set rng = sh.Range(cellRng.Cells(1, 1), cellRng.Cells(nGeo, 1))
    rng.Name = "INPUTSPTGEO_USER_" & tabId
    rng.Locked = False
    rng.Offset(, 1).Name = "INPUTPREVADM_USER_" & tabId

    Set AddSpatioTemporalGeoInputs = cellRng.Offset(nGeo + 3)
End Function

'@label:add-temporal-date-controls
'@sub-title Write the start date, time unit and end date controls of a section.
'@param anchorRng Range. The cell the control block is measured from.
Private Sub AddTemporalDateControls(ByVal anchorRng As Range)
    Dim tabId As String
    Dim trans As TranslationObject

    tabId = TableId()
    Set trans = Translations()

    With anchorRng
        .Cells(6, 1).Name = "INFO_ANA_PERIOD_" & tabId
        .Cells(3, 3).Value = trans.TranslatedValue("MSG_StartDate")
        .Cells(3, 3).Name = "START_DATE_LABEL_" & tabId
        .Cells(4, 3).Locked = False
        .Cells(4, 3).Name = "USER_START_DATE_" & tabId
        .Cells(5, 3).Name = "INFO_START_DATE_" & tabId
        .Cells(6, 3).Name = "START_DATE_" & tabId

        .Cells(3, 5).Value = trans.TranslatedValue("MSG_TimeUnit")
        .Cells(3, 5).Name = "TIME_UNIT_LABEL_" & tabId
        .Cells(4, 5).Name = "TIME_UNIT_" & tabId
        .Cells(4, 5).Locked = False

        .Cells(3, 7).Value = trans.TranslatedValue("MSG_EndDate")
        .Cells(3, 7).Name = "END_DATE_LABEL_" & tabId
        .Cells(4, 7).Name = "USER_END_DATE_" & tabId
        .Cells(4, 7).Locked = False
        .Cells(5, 7).Name = "INFO_END_DATE_" & tabId
        .Cells(6, 7).Name = "END_DATE_" & tabId

        .Cells(4, 9).Name = "VALIDATION_MIN_DATE_" & tabId
        .Cells(4, 11).Name = "VALIDATION_MAX_DATE_" & tabId
        .Cells(4, 13).Name = "MIN_MIN_DATE_" & tabId
        .Cells(4, 15).Name = "MAX_MAX_DATE_" & tabId
    End With

    'The time units are one workbook-level dropdown that Linelist.Prepare adds,
    'so a table only has to say which of them it starts on. Each analysis sheet
    'used to build five cells of its own in column Z under one of two names.
    Wksh().Range("TIME_UNIT_" & tabId).Value = trans.TranslatedValue("MSG_Week")
End Sub

'@label:add-spatial-header
'@sub-title Write the dropdowns and the title of a spatial table.
'@details
'Two of the three lists are shared by every spatial table of the sheet and are
'built once. The third is this table's own dropdown, geographic or health
'facility, and the title and section labels are inserted above it.
'@throws ErrorUnexpectedState When the table names neither an adm1_ nor an hf_ variable.
Private Sub AddSpatialHeader()
    Dim sh As Worksheet
    Dim startRw As Long
    Dim tabId As String
    Dim counter As Long
    Dim rng As Range
    Dim cellRng As Range
    Dim listRng As Range
    Dim trans As TranslationObject
    Dim spatType As String

    Set sh = Wksh()
    Set trans = Translations()
    startRw = StartRow
    tabId = TableId()

    With sh
        If Not RangeExists("ADM_UNIT_LIST") Then
            Set cellRng = .Cells(1, 26)
            cellRng.Cells(1, 1).Formula = "= RNG_ADM1NAME"
            cellRng.Cells(2, 1).Formula = "= RNG_ADM2NAME"
            cellRng.Cells(3, 1).Formula = "= RNG_ADM3NAME"
            cellRng.Cells(4, 1).Formula = "= RNG_ADM4NAME"
            Set listRng = .Range(.Cells(1, cellRng.Column), .Cells(4, cellRng.Column))
            listRng.Name = "ADM_UNIT_LIST"
            listRng.Font.Color = vbWhite
            listRng.Interior.Color = vbWhite
            listRng.Locked = True
        End If

        If Not RangeExists("POPULATION_FACTOR_LIST") Then
            Set cellRng = .Cells(1, 27)
            cellRng.Cells(1, 1).Value = trans.TranslatedValue("MSG_NoDevide")
            cellRng.Cells(2, 1).Value = trans.TranslatedValue("MSG_Devide")
            Set listRng = .Range(.Cells(1, cellRng.Column), .Cells(2, cellRng.Column))
            listRng.Name = "POPULATION_FACTOR_LIST"
            listRng.Font.Color = vbWhite
            listRng.Interior.Color = vbWhite
            listRng.Locked = True
        End If

        If RangeExists("ADM_DROPDOWN_" & tabId) Then Exit Sub
        If RangeExists("HF_DROPDOWN_" & tabId) Then Exit Sub

        ' SpatialTableScopes has three outcomes and the two arms below
        ' handle two of them. rng is assigned inside those arms and then
        ' dereferenced unconditionally underneath, so the third outcome
        ' reaches .Rows(rng.Row - 1) with rng still Nothing and raises 91
        ' well away from the cause. TableSpecs already requires one of
        ' the two prefixes before a spatial table is judged valid, so an
        ' empty answer here means that invariant broke. Saying so costs
        ' one line and names the table.
        spatType = SpatialType()
        If LenB(spatType) = 0 Then _
            ThrowError ErrorUnexpectedState, _
                       "Spatial table " & tabId & _
                       " is missing both an adm1_ and an hf_ variable"

        If spatType = "geo" Then
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

        ElseIf spatType = "hf" Then
            .Cells(startRw - 2, STANDARD_START_COL).Name = "HF_DROPDOWN_" & tabId
            Set rng = .Range("HF_DROPDOWN_" & tabId)
            rng.Cells(1, 1).Value = trans.TranslatedValue("MSG_HF")
        End If

        If Specifications.IsNewSection() Then
            For counter = 1 To 3
                .Rows(rng.Row - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Next
            rng.Cells(-3, 1).Value = ValueOf("section")
            rng.Cells(-3, 1).Name = "SECTION_" & tabId
        End If

        .Rows(rng.Row - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        rng.Cells(-1, 1).Value = ValueOf("title")
        rng.Cells(-1, 1).Name = "TITLE_" & tabId
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
'Every name is created on the worksheet itself, which is where the export
'reads them back from.
'@throws ErrorUnexpectedState When NumberOfColumns is 0, indicating AddColumns was not called.
'@remarks Exits immediately for ScopeGlobalSummary (which has no row categories or
'column structure beyond the two fixed columns).
'@export
Public Sub NameRanges()
    Dim catRng As Range

    RequirePhase PHASE_COLUMNS, "NameRanges"
    MarkPhase PHASE_NAMES

    If TableScope() = ScopeGlobalSummary Then Exit Sub

    NameEndTable

    If NumberOfColumns() = 0 Then _
        ThrowError ErrorUnexpectedState, "Table " & TableId() & " has 0 columns"

    Set catRng = NameCategoryRows()
    NameValueColumns catRng
    NameTotalsAndMissing
    NameIntersections
    NameInteriorValues
    NameColumnCategories
End Sub

'@label:name-end-table
'@sub-title Mark the row under the table, which is what the next one measures from.
Private Sub NameEndTable()
    Dim sh As Worksheet

    If RangeExists("ENDTABLE_" & TableId()) Then Exit Sub

    Set sh = Wksh()
    sh.Cells(EndRow + 1, StartColumn).Name = "ENDTABLE_" & TableId()
End Sub

'@label:name-category-rows
'@sub-title Name the row category block and hand it back.
'@details
'A table that opens a section, and every non-temporal table, owns its category
'block and names it. A temporal table continuing a section shares the labels the
'anchor wrote, so it derives the block from its own ENDTABLE_ marker instead and
'names nothing.
'@return Range. The category block the value columns hang off.
'@throws ErrorUnexpectedState When a continuation table finds no table above it.
Private Function NameCategoryRows() As Range
    Dim catRng As Range
    Dim rng As Range
    Dim cellRng As Range
    Dim sh As Worksheet
    Dim tabId As String
    Dim tablType As Byte
    Dim prevSpecs As TableSpecs
    Dim prevTabId As String

    Set sh = Wksh()
    tabId = TableId()
    tablType = TableScope()

    With sh
        If Specifications.IsNewSection() Or _
           (tablType <> ScopeTimeSeries And tablType <> ScopeSpatioTemporal) Then

            Set catRng = RowsCategoriesRange(onlyCategories:=True, includeHeaders:=False)
            catRng.Name = "ROW_CATEGORIES_" & tabId
            catRng.Cells(0, 1).Name = "LABEL_ROW_CATEGORIES_" & tabId

            If tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal Then
                Set rng = catRng.Offset(, -1)
                rng.Name = "END_TIME_PERIOD_" & tabId
                Set rng = rng.Offset(, -1)
                rng.Cells(1, 1).Name = "FIRST_VALUE_START_TIME_" & tabId
                Set rng = .Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, 1))
                rng.Name = "START_TIME_PERIOD_" & tabId
            End If

        Else
            Set prevSpecs = Specifications.Previous
            If Not prevSpecs Is Nothing Then prevTabId = prevSpecs.TableId

            If Not RangeExists("ENDTABLE_" & prevTabId) Then
                ThrowError ErrorUnexpectedState, _
                           "Cannot append time series: no previous table found for " & prevTabId
            End If

            Set cellRng = .Range("ENDTABLE_" & tabId)
            Set catRng = .Range(cellRng.Offset(1 - NB_ROWS_TIME_SERIES), cellRng.Offset(-3))
        End If
    End With

    Set NameCategoryRows = catRng
End Function

'@label:name-value-columns
'@sub-title Name every value column and its label cell.
'@details
'The label row sits one row above the values for every scope but the univariate
'one, where it sits on the same row. That answer is the same for every column of
'a table, so it is read once above the loop.
'@param catRng Range. The category block the first value column sits beside.
'@throws ErrorUnexpectedState When the table scope has no value column layout.
Private Sub NameValueColumns(ByVal catRng As Range)
    Dim rng As Range
    Dim counter As Long
    Dim nbColCounter As Long
    Dim nbCols As Long
    Dim labCol As Long
    Dim tabId As String
    Dim tablType As Byte
    Dim hasPercent As Boolean

    tabId = TableId()
    tablType = TableScope()
    nbCols = NumberOfColumns()
    hasPercent = Specifications.HasPercentage

    Select Case tablType
    Case ScopeBivariate, ScopeTimeSeries, ScopeSpatial, ScopeSpatioTemporal
        labCol = -1
    Case ScopeUnivariate
        labCol = 0
    Case Else
        ' Same family as the guard in RowsCategoriesRange. A scope outside the
        ' five leaves labCol at zero, which silently names the univariate label
        ' row on a table that has no such row.
        ThrowError ErrorUnexpectedState, _
                   "Value columns were asked of table " & tabId & _
                   ", whose scope this class does not lay out"
    End Select

    counter = StartColumn
    nbColCounter = 1
    Set rng = catRng.Offset(, 1)

    Do While (counter <= EndColumn) And (nbColCounter <= nbCols)
        rng.Name = "VALUES_COL_" & nbColCounter & "_" & tabId
        rng.Cells(labCol, 1).Name = "LABEL_COL_" & nbColCounter & "_" & tabId

        Set rng = rng.Offset(, 1)

        If hasPercent Then
            rng.Name = "PERC_COL_" & nbColCounter & "_" & tabId

            If tablType = ScopeUnivariate Then
                rng.Cells(labCol, 1).Name = "PERC_LABEL_COL_" & tabId
            End If

            Set rng = rng.Offset(, 1)
            counter = counter + 1
        End If

        counter = counter + 1
        nbColCounter = nbColCounter + 1
    Loop
End Sub

'@label:name-totals-and-missing
'@sub-title Name the Total, Missing and Percent rows and columns.
'@details
'Each of the three is found by its translated label, so a table that carries
'none of them names none of them. The range just named is kept and offset for
'the percentage twin instead of being resolved back by name.
Private Sub NameTotalsAndMissing()
    Dim rng As Range
    Dim colRng As Range
    Dim counter As Long
    Dim tabId As String
    Dim tablType As Byte
    Dim hasPercent As Boolean
    Dim trans As TranslationObject
    Dim prefixes As Variant
    Dim tags As Variant
    Dim label As String

    tabId = TableId()
    tablType = TableScope()
    hasPercent = Specifications.HasPercentage
    Set trans = Translations()

    prefixes = Split("MISSING|TOTAL|PERC", "|")
    tags = Split("MSG_NA|MSG_Total|MSG_Percent", "|")

    For counter = LBound(prefixes) To UBound(prefixes)
        label = trans.TranslatedValue(CStr(tags(counter)))

        ' Row ranges
        Set rng = RowRange(label, includeHeaders:=True)
        If Not rng Is Nothing Then
            rng.Name = prefixes(counter) & "_ROW_" & tabId
            Set rng = RowRange(label, includeHeaders:=False)
            rng.Name = prefixes(counter) & "_ROW_VALUES_" & tabId
        End If

        ' Column ranges (not for univariate)
        If tablType <> ScopeUnivariate Then
            Set colRng = ColumnRange(label, includeHeaders:=True, onlyCategories:=True)
            If Not colRng Is Nothing Then
                colRng.Name = prefixes(counter) & "_COL_" & tabId
                colRng.Cells(0, 1).Name = prefixes(counter) & "_LABEL_COL_" & tabId

                Set rng = ColumnRange(label, includeHeaders:=False, onlyCategories:=True)
                rng.Name = prefixes(counter) & "_COL_VALUES_" & tabId

                If hasPercent Then
                    rng.Offset(, 1).Name = prefixes(counter) & "_PERC_VALUES_" & tabId
                    colRng.Offset(, 1).Name = prefixes(counter) & "_PERC_" & tabId
                End If
            End If
        End If
    Next
End Sub

'@label:name-intersections
'@sub-title Name the cells where a total or missing row meets such a column.
Private Sub NameIntersections()
    Dim sh As Worksheet
    Dim tabId As String
    Dim totalCol As Range
    Dim missingCol As Range
    Dim totalRow As Range
    Dim missingRow As Range

    Set sh = Wksh()
    tabId = TableId()

    Set totalCol = NamedRange("TOTAL_COL_" & tabId)
    Set missingCol = NamedRange("MISSING_COL_" & tabId)
    Set totalRow = NamedRange("TOTAL_ROW_" & tabId)
    Set missingRow = NamedRange("MISSING_ROW_" & tabId)

    With sh
        If Not totalCol Is Nothing Then
            If Not totalRow Is Nothing Then _
                .Cells(totalRow.Row, totalCol.Column).Name = "TOTAL_TOTAL_" & tabId

            If Not missingRow Is Nothing Then
                .Cells(missingRow.Row, totalCol.Column).Name = "MISSING_TOTAL_" & tabId
                If Not missingCol Is Nothing Then _
                    .Cells(missingRow.Row, missingCol.Column).Name = "MISSING_MISSING_" & tabId
            End If
        End If

        If (Not totalRow Is Nothing) And (Not missingCol Is Nothing) Then _
            .Cells(totalRow.Row, missingCol.Column).Name = "TOTAL_MISSING_" & tabId
    End With
End Sub

'@label:name-interior-values
'@sub-title Name the block of data cells, and the outer block a spatial table adds.
'@details
'The two end columns are read as ranges and the block is built from them.
'Splitting an address on a colon reads the second half of a one-cell address
'that has no colon in it, and a value column is one cell tall whenever the table
'has one category row: one category and no missing row for a univariate table,
'or a geo count of one for a spatial one. This range is required -- Format and
'CrossTableFormula both read it, and twenty-odd names are already on the sheet
'by the time it is built, so a raise here left the table half-named.
Private Sub NameInteriorValues()
    Dim sh As Worksheet
    Dim tabId As String
    Dim firstValRng As Range
    Dim lastValRng As Range
    Dim interiorRng As Range
    Dim rng As Range

    Set sh = Wksh()
    tabId = TableId()

    With sh
        Set firstValRng = .Range("VALUES_COL_1_" & tabId)
        Set lastValRng = .Range(IIf(Specifications.HasPercentage, "PERC_COL_", "VALUES_COL_") & _
                                NumberOfColumns() & "_" & tabId)
        Set interiorRng = .Range(firstValRng.Cells(1, 1), _
                                 lastValRng.Cells(lastValRng.Rows.Count, 1))
        interiorRng.Name = "INTERIOR_VALUES_" & tabId

        ' Spatial outer values for population computation
        If TableScope() = ScopeSpatial Then
            Set rng = .Range(interiorRng.Cells(1, -1), _
                             interiorRng.Cells(interiorRng.Rows.Count + 1, _
                                               interiorRng.Columns.Count + 4))
            rng.Name = "OUTER_VALUES_" & tabId
        End If
    End With
End Sub

'@label:name-column-categories
'@sub-title Name the header row of the table.
Private Sub NameColumnCategories()
    HeaderRange().Name = "COLUMN_CATEGORIES_" & TableId()
End Sub

'@section Format
'===============================================================================
'@description Apply visual formatting to all structural regions of the cross-table.

'@label:format
'@sub-title Apply formatting to the table using design specifications.
'@details
'Applies comprehensive visual formatting to every part of the cross-table
'using the design format object (LLFormat). This is the final step after
'Build completes the structural layout. Formatting is applied in order:
'section header, table title, spatial dropdown, column headers, row and
'column categories, interior values, Total/Missing rows and columns,
'single cells, info cells, hidden cells, percentage columns, whole table
'borders, and end-table gap. Bug #1 fix: for temporal tables where
'HasTotal is True but TotalRequested is False, hides the Total column
'entirely since it is needed internally for percentage computation but
'should be invisible to the user.
'@param desFormat LLFormat. The formatting specification providing ApplyFormat for each scope.
'@remarks Format must be called after Build because it relies on named ranges being
'present. Some operations are idempotent, guarded by HiddenNames like
'ROWGS_FORMATSET and MERGED_.
'@export
Public Sub Format(ByVal desFormat As LLFormat)
    Dim hn As HiddenNames

    RequirePhase PHASE_NAMES, "Format"
    Set hn = HiddenNames.Create(Wksh())

    FormatSectionAndTitle desFormat, hn
    FormatSpatialDropdown desFormat
    FormatHeaders desFormat, hn
    FormatCategories desFormat, hn
    FormatInteriorValues desFormat
    FormatBandsAndCells desFormat
    FormatPercentageColumns desFormat
    FormatWholeTable desFormat
    HideUnrequestedTotals
End Sub

'@label:format-section-and-title
'@sub-title Format the section label and the table title.
'@details
'The guard on the second arm is the table scope. ROWGS_SET, the name it used to
'read, belongs to the global summary and is written on one sheet, but this arm
'is reached by any table that does not start a section, on any of the four
'analysis sheets. Asking whether the name resolved meant the time series builder
'formatted a range on the normal sheet and then stamped ROWGS_FORMATSET, a
'worksheet-scoped flag, onto its own -- so the flag described a sheet it was not
'written on. Tightening RangeExists closes the same hole from the other side;
'this states the intent.
'@param desFormat LLFormat. The formatting provider.
'@param hn HiddenNames. The worksheet's hidden name store.
Private Sub FormatSectionAndTitle(ByVal desFormat As LLFormat, ByVal hn As HiddenNames)
    Dim sh As Worksheet
    Dim tabId As String
    Dim tablType As Byte

    Set sh = Wksh()
    tabId = TableId()
    tablType = TableScope()

    If Specifications.IsNewSection() Then
        desFormat.ApplyFormat sh.Range("SECTION_" & tabId), AnalysisSection
    ElseIf (tablType = ScopeGlobalSummary) And (Not hn.HasName("ROWGS_FORMATSET")) Then
        desFormat.ApplyFormat sh.Range("ROWGS_SET"), scope:=AnalysisSection
        hn.EnsureName "ROWGS_FORMATSET", True, HiddenNameTypeBoolean
    End If

    ' Global summary and the two temporal scopes carry no title of their own.
    If (tablType <> ScopeGlobalSummary) And (tablType <> ScopeTimeSeries) And _
       (tablType <> ScopeSpatioTemporal) Then
        desFormat.ApplyFormat sh.Range("TITLE_" & tabId), AnalysisTableTitle
    End If
End Sub

'@label:format-spatial-dropdown
'@sub-title Format the geo or health facility dropdown of a spatial table.
'@param desFormat LLFormat. The formatting provider.
Private Sub FormatSpatialDropdown(ByVal desFormat As LLFormat)
    Dim sh As Worksheet
    Dim tabId As String
    Dim rng As Range
    Dim spatType As String

    If TableScope() <> ScopeSpatial Then Exit Sub

    Set sh = Wksh()
    tabId = TableId()
    spatType = SpatialType()

    If spatType = "geo" Then
        desFormat.ApplyFormat sh.Range("ADM_DROPDOWN_" & tabId), scope:=LinelistSelectionDropdown
        desFormat.ApplyFormat sh.Range("DEVIDEPOP_" & tabId), scope:=LinelistSelectionDropdown
    ElseIf spatType = "hf" Then
        desFormat.ApplyFormat sh.Range("HF_DROPDOWN_" & tabId), scope:=LinelistHiddenCell
    End If

    Set rng = sh.Range("ROW_CATEGORIES_" & tabId).Offset(, -1)
    desFormat.ApplyFormat rng, scope:=LinelistHiddenCell
    Set rng = rng.Offset(, -1)
    desFormat.ApplyFormat rng, scope:=LinelistHiddenCell
End Sub

'@label:format-headers
'@sub-title Format the header row of the table.
'@param desFormat LLFormat. The formatting provider.
'@param hn HiddenNames. The worksheet's hidden name store.
Private Sub FormatHeaders(ByVal desFormat As LLFormat, ByVal hn As HiddenNames)
    Dim sh As Worksheet
    Dim rng As Range
    Dim tabId As String

    Set sh = Wksh()
    tabId = TableId()

    With sh
        Select Case TableScope()
        Case ScopeGlobalSummary
            If Not hn.HasName("COLGS_FORMATSET") Then
                Set rng = .Range("COLGS_SET")
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(1, 2))
                desFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
                hn.EnsureName "COLGS_FORMATSET", True, HiddenNameTypeBoolean
            End If
        Case ScopeUnivariate
            desFormat.ApplyFormat .Range("COLUMN_CATEGORIES_" & tabId), _
                                  scope:=AnalysisUnivariateHeader
            desFormat.ApplyFormat .Range("LABEL_ROW_CATEGORIES_" & tabId), _
                                  scope:=AnalysisCategoriesLabelRow
        Case ScopeBivariate, ScopeSpatial
            desFormat.ApplyFormat .Range("LABEL_ROW_CATEGORIES_" & tabId), _
                                  scope:=AnalysisCategoriesLabelRow
            desFormat.ApplyFormat .Range("STARTCOL_" & tabId), _
                                  scope:=AnalysisCategoriesLabelColumn
        Case ScopeTimeSeries, ScopeSpatioTemporal
            If Specifications.IsNewSection() Then
                desFormat.ApplyFormat .Range("LABEL_ROW_CATEGORIES_" & tabId), _
                                      scope:=AnalysisCategoriesLabelRow
            End If
            Set rng = .Range("STARTCOL_" & tabId)
            Set rng = .Range(.Cells(rng.Row, rng.Column), .Cells(rng.Row, EndColumn()))
            desFormat.ApplyFormat rng, scope:=AnalysisTimeSeriesHeader
        End Select
    End With
End Sub

'@label:format-categories
'@sub-title Format the category rows and columns, and merge the column labels.
'@param desFormat LLFormat. The formatting provider.
'@param hn HiddenNames. The worksheet's hidden name store.
Private Sub FormatCategories(ByVal desFormat As LLFormat, ByVal hn As HiddenNames)
    Dim sh As Worksheet
    Dim rng As Range
    Dim tabId As String
    Dim tablType As Byte
    Dim newSection As Boolean
    Dim hasPercent As Boolean
    Dim counter As Long
    Dim dateCells As Variant

    Set sh = Wksh()
    tabId = TableId()
    tablType = TableScope()
    newSection = Specifications.IsNewSection()
    hasPercent = Specifications.HasPercentage

    With sh
        Select Case tablType
        Case ScopeGlobalSummary
            desFormat.ApplyFormat .Range("STARTROW_" & tabId), scope:=AnalysisGlobalSummaryRow
            Exit Sub
        Case ScopeUnivariate
            desFormat.ApplyFormat .Range("ROW_CATEGORIES_" & tabId), _
                                  scope:=AnalysisCategoriesUnivariateRow
            Exit Sub
        Case ScopeBivariate, ScopeTimeSeries, ScopeSpatial, ScopeSpatioTemporal
            ' Handled below.
        Case Else
            Exit Sub
        End Select

        If ((tablType <> ScopeTimeSeries) And (tablType <> ScopeSpatioTemporal)) Or newSection Then
            Set rng = .Range("ROW_CATEGORIES_" & tabId)
            desFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateRow

            If tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal Then
                Set rng = rng.Offset(, -1)
                Set rng = .Range(rng.Cells(1, 0), rng.Cells(rng.Rows.Count + 1, 1))
                desFormat.ApplyFormat rng, scope:=LinelistHiddenCell

                dateCells = Split(TEMPORAL_DATE_CELLS, "|")
                For counter = LBound(dateCells) To UBound(dateCells)
                    Set rng = .Range(dateCells(counter) & tabId)
                    desFormat.ApplyFormat rng, scope:=LinelistHiddenCell
                    desFormat.ApplyFormat rng, scope:=AnalysisDateCell
                Next
            End If
        End If

        Set rng = NamedRange("INPUTSPTGEO_USER_" & tabId)
        If Not rng Is Nothing Then _
            desFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateRow

        ' Column label merging
        If Not hn.HasName("MERGED_" & tabId) Then
            Set rng = .Range("LABEL_COL_1_" & tabId)
            Do While rng.Column <= EndColumn()
                If hasPercent Then
                    Set rng = rng.Offset(, 2)
                    .Range(rng.Cells(1, -1), rng.Cells(1, 0)).Merge
                Else
                    Set rng = rng.Offset(, 1)
                End If
            Loop
            hn.EnsureName "MERGED_" & tabId, True, HiddenNameTypeBoolean
        End If

        Set rng = .Range("COLUMN_CATEGORIES_" & tabId)
        desFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateColumn
        desFormat.ApplyFormat rng.Offset(1), scope:=AnalysisFirstRow
    End With
End Sub

'@label:format-interior-values
'@sub-title Format the block of data cells.
'@param desFormat LLFormat. The formatting provider.
Private Sub FormatInteriorValues(ByVal desFormat As LLFormat)
    If TableScope() = ScopeGlobalSummary Then Exit Sub
    desFormat.ApplyFormat Wksh().Range("INTERIOR_VALUES_" & TableId()), _
                          scope:=AnalysisInteriorTable
End Sub

'@label:format-bands-and-cells
'@sub-title Format the total and missing bands, the single cells, the info cells
'and the hidden cells.
'@details
'Four walks over four constant prefix lists. Each entry used to cost two name
'resolutions, one to ask whether it existed and one to read it, and about half
'the entries of any list do not exist for the scope in hand. NamedRange answers
'both questions at once.
'@param desFormat LLFormat. The formatting provider.
Private Sub FormatBandsAndCells(ByVal desFormat As LLFormat)
    Dim sh As Worksheet
    Dim rng As Range
    Dim counter As Long
    Dim rngName As String
    Dim tabId As String
    Dim tablType As Byte
    Dim newSection As Boolean
    Dim prefixes As Variant

    Set sh = Wksh()
    tabId = TableId()
    tablType = TableScope()
    newSection = Specifications.IsNewSection()

    ' Total and missing bands
    prefixes = Split(BAND_CELLS, "|")
    For counter = LBound(prefixes) To UBound(prefixes)
        rngName = prefixes(counter) & tabId
        Set rng = NamedRange(rngName)
        If Not rng Is Nothing Then
            If InStr(1, rngName, "ROW") > 0 Then
                If (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And _
                   (Not newSection) Then
                    desFormat.ApplyFormat rng, scope:=AnalysisMissingRowWithoutHeader
                Else
                    desFormat.ApplyFormat rng, scope:=AnalysisMissingRow
                End If
            Else
                desFormat.ApplyFormat rng, scope:=AnalysisMissingColumn
            End If
        End If
    Next

    ' Single cells
    prefixes = Split(SINGLE_CELLS, "|")
    For counter = LBound(prefixes) To UBound(prefixes)
        rngName = prefixes(counter) & tabId
        Set rng = NamedRange(rngName)
        If Not rng Is Nothing Then
            If (InStr(1, rngName, "LABEL") > 0) Or (InStr(1, rngName, "POP") > 0) Then
                desFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
            ElseIf (InStr(1, rngName, "USER") > 0) Or (rngName = "TIME_UNIT_" & tabId) Then
                desFormat.ApplyFormat rng, scope:=AnalysisOneCellCenter
            Else
                desFormat.ApplyFormat rng, scope:=AnalysisOneCell
            End If
        End If
    Next

    ' Info cells
    prefixes = Split(INFO_CELLS, "|")
    For counter = LBound(prefixes) To UBound(prefixes)
        rngName = prefixes(counter) & tabId
        Set rng = NamedRange(rngName)
        If Not rng Is Nothing Then
            desFormat.ApplyFormat rng, scope:=AnalysisInfoCell
            If InStr(1, rngName, "INFO_ANA_") > 0 Then
                desFormat.ApplyFormat rng, scope:=AnalysisInfoCellLeft
            End If
        End If
    Next

    ' Hidden cells
    prefixes = Split(HIDDEN_CELLS, "|")
    For counter = LBound(prefixes) To UBound(prefixes)
        Set rng = NamedRange(prefixes(counter) & tabId)
        If Not rng Is Nothing Then desFormat.ApplyFormat rng, scope:=LinelistHiddenCell
    Next
End Sub

'@label:format-percentage-columns
'@sub-title Format every percentage column and the corner cells beside them.
'@param desFormat LLFormat. The formatting provider.
Private Sub FormatPercentageColumns(ByVal desFormat As LLFormat)
    Dim sh As Worksheet
    Dim rng As Range
    Dim counter As Long
    Dim tabId As String
    Dim prefixes As Variant

    If Not Specifications.HasPercentage Then Exit Sub

    Set sh = Wksh()
    tabId = TableId()

    With sh
        Set rng = .Range("PERC_COL_1_" & tabId)
        If (TableScope() <> ScopeSpatial) Or Specifications.HasTotal Then
            Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count + 1, 1))
        End If

        Do While rng.Column <= EndColumn()
            desFormat.ApplyFormat rng, scope:=AnalysisPercent
            Set rng = rng.Offset(, 2)
        Loop
    End With

    prefixes = Split(PERCENT_CORNER_CELLS, "|")
    For counter = LBound(prefixes) To UBound(prefixes)
        Set rng = NamedRange(prefixes(counter) & tabId)
        If Not rng Is Nothing Then _
            desFormat.ApplyFormat rng.Offset(, 1), scope:=AnalysisOneCellCenter
    Next
End Sub

'@label:format-whole-table
'@sub-title Draw the outline of the table and the gap under it.
'@param desFormat LLFormat. The formatting provider.
Private Sub FormatWholeTable(ByVal desFormat As LLFormat)
    Dim sh As Worksheet
    Dim rng As Range
    Dim tabId As String
    Dim tablType As Byte
    Dim wholeTableStart As Long

    Set sh = Wksh()
    tabId = TableId()
    tablType = TableScope()

    wholeTableStart = StartColumn
    If (tablType = ScopeTimeSeries Or tablType = ScopeSpatioTemporal) And _
       (Not Specifications.IsNewSection()) Then
        wholeTableStart = StartColumn + 1
    End If

    With sh
        If (tablType <> ScopeUnivariate) And (tablType <> ScopeGlobalSummary) Then
            Set rng = .Range(.Cells(StartRow, wholeTableStart), .Cells(EndRow, EndColumn))
            desFormat.ApplyFormat rng, scope:=AnalysisWholeTable
        ElseIf tablType = ScopeUnivariate Then
            Set rng = .Range(.Cells(StartRow + 1, wholeTableStart), .Cells(EndRow, EndColumn))
            desFormat.ApplyFormat rng, scope:=AnalysisWholeTable
        End If
    End With

    ' End table gap for graphs
    If (tablType = ScopeTimeSeries) Or (tablType = ScopeSpatioTemporal) Then Exit Sub

    Set rng = NamedRange("ENDTABLE_" & tabId)
    If Not rng Is Nothing Then desFormat.ApplyFormat rng, scope:=AnalysisEndTable
End Sub

'@label:hide-unrequested-totals
'@sub-title Hide the total row and column of a temporal table that asked for neither.
'@details
'For time series and spatio-temporal tables HasTotal can be True while the user
'asked for nothing: percentage="row" needs the total column as its denominator,
'so the column has to exist and the user still has no reason to see it.
'
'THE ROW IS THE OTHER HALF OF THE SAME ANSWER -- ISSUE #338
'-------------------------------------------------------------------------------
'AddRows writes the Total and Missing labels of a temporal section
'unconditionally, because the row grid is a fixed NB_ROWS_TIME_SERIES block with
'exactly two footer rows and RowsCategoriesRange trims them BY COUNT. So the row
'cannot be dropped and it is hidden instead, the way the column is. Without this
'a table with "Add total" left empty still showed a Total row, which is what the
'field reported.
'
'The rows belong to the section rather than to one table, so the anchor -- the
'table that wrote them -- is what hides them, and any later table of the same
'section that does ask for a total shows them again. Nothing hides a row a table
'has asked for, whatever order the section is built in. Every table of the
'section names the same sheet row through its own TOTAL_ROW_<tableId>, so the
'name is always this table's.
Private Sub HideUnrequestedTotals()
    Dim tabId As String
    Dim tablType As Byte
    Dim totalCol As Range
    Dim totalPerc As Range
    Dim totalRow As Range

    tablType = TableScope()
    If (tablType <> ScopeTimeSeries) And (tablType <> ScopeSpatioTemporal) Then Exit Sub

    tabId = TableId()

    If Specifications.HasTotal And (Not Specifications.TotalRequested) Then
        Set totalCol = NamedRange("TOTAL_COL_" & tabId)
        If Not totalCol Is Nothing Then
            totalCol.EntireColumn.Hidden = True
            If Specifications.HasPercentage Then
                Set totalPerc = NamedRange("TOTAL_PERC_VALUES_" & tabId)
                If Not totalPerc Is Nothing Then totalPerc.EntireColumn.Hidden = True
            End If
        End If
    End If

    Set totalRow = NamedRange("TOTAL_ROW_" & tabId)
    If totalRow Is Nothing Then Exit Sub

    If Specifications.TotalRequested() Then
        totalRow.EntireRow.Hidden = False
    ElseIf Specifications.IsNewSection() Then
        totalRow.EntireRow.Hidden = True
    End If
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
Public Sub Build()
    AddHeader

    ' The layout is fixed here, between the phase that moves the sheet and the
    ' three that read it. AddHeader inserts rows, so the start row is only stable
    ' once it has returned; and the start column has to be on the sheet before
    ' AddColumns, because the scan that finds a continuation table's column
    ' answers something else as soon as that table's own columns are written.
    CacheStartRow
    EnsureStartColumn

    AddRows
    AddColumns
    NameRanges
    LogInfo "Table " & TableId() & " built"
End Sub

'@label:mark-phase
'@sub-title Record that a build phase has run.
'@param reached Byte. The phase that has just finished.
Private Sub MarkPhase(ByVal reached As Byte)
    If this.phase < reached Then this.phase = reached
End Sub

'@label:require-phase
'@sub-title Refuse a phase whose inputs have not been written yet.
'@details
'The four phases have to run in order and nothing said so. AddHeader inserts
'rows and shifts the sheet, AddRows sets EndRow, AddColumns sets EndColumn and
'NumberOfColumns, and NameRanges needs all three. Calling them out of order used
'to give a table built at row zero or a name over an empty block, several frames
'from the mistake.
'@param completed Byte. The phase that must already have run.
'@param caller String. The member refusing, for the message.
'@throws ErrorUnexpectedState When the phase has not been reached.
Private Sub RequirePhase(ByVal completed As Byte, ByVal caller As String)
    If this.phase >= completed Then Exit Sub

    ThrowError ErrorUnexpectedState, _
               caller & " was called on table " & TableId() & _
               " before the build reached phase " & completed
End Sub

'@section Checkings
'===============================================================================
'@description Diagnostic logging for cross-table building.

'Adds a trace entry to the internal Checking instance, initialising it
'lazily on first use.
'@param message String. The diagnostic message text.
'@param scope Optional Byte. The checking scope (default checkingSuccess).
Private Sub LogInfo(ByVal message As String, _
                    Optional ByVal scope As Byte = checkingSuccess)
    If internalChecks Is Nothing Then
        Set internalChecks = Checking.Create("CrossTable")
    End If

    ' Checking.Add takes (key, label, scope). Passing (message, scope) made the
    ' message the key and coerced the Byte scope into the label, so every check
    ' this class raised was filed as a note carrying a bare number where its text
    ' belonged, and the scope the caller asked for was thrown away. Keys must
    ' also be unique or Add raises, and a message used as a key is only unique
    ' while every message differs -- a second call with a fixed message would have
    ' taken its table down from inside the guarded window.
    '
    ' The key names the class and the table as well as the counter. AnalysisOutput
    ' pours the entries of this class and of CrossTableFormula, for every table of
    ' the sheet, into one report, and a bare counter made the second Append raise
    ' on a key the first had already taken.
    checkCounter = checkCounter + 1
    internalChecks.Add "CrossTable-" & TableId() & "-" & CStr(checkCounter), _
                       message, scope
    this.hasChecks = True
End Sub

Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

Public Property Get CheckingValues() As Checking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = internalChecks
End Property


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

