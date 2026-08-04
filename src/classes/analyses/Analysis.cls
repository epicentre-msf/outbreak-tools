VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Setup analysis configuration worksheet manager"
Option Explicit


'@PredeclaredId
'@Folder("analyses")
'@ModuleDescription("Setup analysis configuration worksheet manager")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  UseMeaningfulName, HungarianNotation

'@class Analysis
'@description
'Manages the Analysis configuration worksheet of a setup workbook, and the copy
'of that worksheet the designer places inside a generated linelist. AddRows,
'InsertRows, DeleteRows, RemoveRows and ManageRows work on the rows, Sort
'orders the tables, Import and Export move the sheet between workbooks, and
'Translate rewrites the analysis labels. Each instance is bound to one
'worksheet and hands its table work to CustomTable. Entries filed along the way
'leave through HasCheckings and CheckingValues.
'
'WHERE THE ANALYSIS OUTPUT LIVES
'-------------------------------------------------------------------------------
'The tables a field user reads at run time are built by AnalysisOutput,
'CrossTable and TableSpecs. This class handles the configuration sheet. A
'linelist instance of this class receives one call, Translate, from
'LinelistSpecs.
'
'THE NINE REQUIRED TABLES
'-------------------------------------------------------------------------------
'Create checks that the worksheet carries the nine ListObjects an analysis
'definition is written in, and every member reads its rows from one of them.
'
'TABLE NAMES ARE MATCHED WITHOUT REGARD TO CASE
'-------------------------------------------------------------------------------
'Excel matches sheet and table names without regard to case, so this class does
'the same everywhere. The required-table list is walked with vbTextCompare and
'the table settings are dispatched on a lower-cased name. The shipped workbook
'writes "Tab_Global_Summary" where the list says "Tab_global_summary".
'
'THE TEN INSTRUCTION STRINGS ARE HELD TWICE
'-------------------------------------------------------------------------------
'ManageRows matches the ten instruction strings the user picks from the
'dropdown in cell RNG_SelectTable. SetupPreparation holds the same ten strings,
'because it is the class that fills that dropdown. Change one list and change
'the other.
'@depends CustomTable, Checking, TranslationObject, BetterArray
'@version 2026-07-29


Private Type TAnalysis
    hostsheet As Worksheet
    checkCounter As Long
    internalChecks As Checking
    requiredTablesCaches As BetterArray
End Type

Private this As TAnalysis

'ListObject names used in analysis worksheets
Private Const GLOBALSUMMARY As String = "Tab_Global_Summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"
Private Const SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const SPATIOTEMPSPECS As String = "Tab_SpatioTemporal_Specs"
Private Const CLASSNAME As String = "Analysis"

'Named range holding the row-management instruction on the setup Analysis sheet
Private Const INSTRUCTIONRANGE As String = "RNG_SelectTable"

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create an Analysis instance bound to a worksheet
'@details
'Factory method on the predeclared instance. Validates that the worksheet
'is not Nothing and binds a new Analysis to it. Callers interact with the
'returned Analysis instance. The required-tables cache is filled lazily on
'the instance that needs it.
'@param hostsheet Worksheet. The worksheet hosting the analysis tables.
'@return Analysis. A fully initialised analysis controller.
'@throws ProjectError.ObjectNotInitialized When hostsheet is Nothing.
'@depends CustomTable, Checking
'@export
Public Function Create(ByVal hostsheet As Worksheet) As Analysis
    If hostsheet Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, _
                                         "Worksheet object not defined"
    'You can create analysis on a empty worksheet, without tables.
    Dim instance As Analysis
    Set instance = New Analysis
    Set instance.Wksh = hostsheet
    Set Create = instance
End Function

'@label:wksh
'@prop-title Backing analysis worksheet
'@details
'Returns the worksheet that stores the analysis tables manipulated by
'this class. Set during construction via the Create factory.
'@return Worksheet. The bound analysis worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.hostsheet
End Property

'@label:wksh-set
'@prop-title Assign the backing worksheet
'@details
'Stores the worksheet reference used by all subsequent operations.
'Called internally by the Create factory.
'@param hostsheet Worksheet. The worksheet to bind.
Public Property Set Wksh(ByVal hostsheet As Worksheet)
    Set this.hostsheet = hostsheet
End Property

'@section RowManagement
'===============================================================================
'@description Add, insert, delete, and remove rows in analysis ListObjects.

'@label:addrows
'@sub-title Add rows based on worksheet header selection
'@details
'Delegates to ManageRows in append mode to add rows to each target
'analysis table based on the instruction held in the RNG_SelectTable cell
'of the worksheet.
'@export
Public Sub AddRows()
    ManageRows
End Sub

'@label:insertrows
'@sub-title Insert rows based on the supplied selection
'@details
'Determines which analysis table hosts the selection by calling
'ResolveSelectionTable, retrieves the table settings (prefix, id column,
'minimum count), and delegates to CustomTable.InsertRowsAt to mirror the
'selection height inside that table. Worksheet row insertion is enabled
'by default to protect stacked ListObjects.
'@param targetCell Range. Anchor describing the rows to mirror.
'@param insertShift Optional Boolean. When True, worksheet rows are
'inserted ahead of the resize. Defaults to True.
'@export
Public Sub InsertRows(ByVal targetCell As Range, _
                       Optional ByVal insertShift As Boolean = True)
    Dim targetList As ListObject
    Dim targetTable As CustomTable
    Dim prefixName As String
    Dim idColumn As String
    Dim minimumCount As Long
    Dim shouldIncludeIds As Boolean

    Set targetList = ResolveSelectionTable(targetCell, "insert")
    If targetList Is Nothing Then Exit Sub

    GetTableSettings targetList.Name, prefixName, idColumn, minimumCount
    shouldIncludeIds = (LenB(idColumn) <> 0)

    Set targetTable = CustomTable.Create(targetList, idColumn, prefixName)
    targetTable.InsertRowsAt targetCell:=targetCell, _
                             insertShift:=insertShift, _
                             includeIds:=shouldIncludeIds
End Sub

'@label:deleterows
'@sub-title Delete rows intersecting the selection
'@details
'Locates the analysis table hosting the selection via ResolveSelectionTable,
'retrieves its settings, and delegates to CustomTable.DeleteRowsAt.
'Identifier columns are cleared for every table that configures one, and
'forceShift controls whether worksheet rows are shifted up after deletion.
'@param targetCell Range. Anchor describing the rows to delete.
'@param forceShift Optional Boolean. When True, worksheet rows are shifted
'up after deletion. Defaults to True.
'@export
Public Sub DeleteRows(ByVal targetCell As Range, _
                       Optional ByVal forceShift As Boolean = True)
    Dim targetList As ListObject
    Dim targetTable As CustomTable
    Dim prefixName As String
    Dim idColumn As String
    Dim minimumCount As Long
    Dim shouldIncludeIds As Boolean

    Set targetList = ResolveSelectionTable(targetCell, "delete")
    If targetList Is Nothing Then Exit Sub

    GetTableSettings targetList.Name, prefixName, idColumn, minimumCount
    shouldIncludeIds = (LenB(idColumn) <> 0)

    Set targetTable = CustomTable.Create(targetList, idColumn, prefixName)
    targetTable.DeleteRowsAt targetCell:=targetCell, _
                             includeIds:=shouldIncludeIds, _
                             forceShift:=forceShift
End Sub

'@label:removerows
'@sub-title Remove rows while preserving minimum entries
'@details
'Calls ManageRows in deletion mode to trim each analysis table down to
'its required minimum row count.
'@export
Public Sub RemoveRows()
    ManageRows True
End Sub

'@label:managerows
'@sub-title Central row-management dispatcher
'@details
'Reads the instruction the user picked in the RNG_SelectTable cell of the
'analysis worksheet and determines which ListObject(s) to target. When del
'is False, rows are appended via CustomTable.AddRows, which renumbers the
'identifier column itself. When del is True, rows are removed via
'CustomTable.RemoveRows down to the required minimum. Falls back to
'processing all known tables when the instruction text does not match a
'specific table name.
'@param del Optional Boolean. When True, rows are removed; otherwise rows
'are added. Defaults to False.
Public Sub ManageRows(Optional ByVal del As Boolean = False)
    Dim listNames As BetterArray
    Dim selectionText As String
    Dim loname As String
    Dim prefName As String
    Dim idCol As String
    Dim requiredCount As Long
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim tabl As CustomTable
    Dim idx As Long
    Dim headerValue As Variant

    Set sh = Wksh()
    'Capture the user instruction displayed in the worksheet header.
    headerValue = InstructionCell(sh).Value
    If IsError(headerValue) Then
        selectionText = vbNullString
    Else
        selectionText = LCase$(CStr(headerValue))
        selectionText = Trim$(Replace(selectionText, "add or remove rows of", vbNullString))
    End If

    Set listNames = New BetterArray

    'SetupPreparation fills the dropdown these ten strings come from and holds
    'its own copy of them. Changing one list means changing the other.
    Select Case selectionText
        Case "global summary": listNames.Push GLOBALSUMMARY
        Case "univariate analysis": listNames.Push UNIVARIATE
        Case "bivariate analysis": listNames.Push BIVARIATE
        Case "time series analysis": listNames.Push TIMESERIES
        Case "labels for time series graphs": listNames.Push TITLEGRAPHTIMESERIES
        Case "graph on time series": listNames.Push GRAPHTIMESERIES
        Case "spatial analysis": listNames.Push SPATIAL
        Case "spatio-temporal specifications": listNames.Push SPATIOTEMPSPECS
        Case "spatio-temporal analysis": listNames.Push SPATIOTEMPORAL
        Case Else
            listNames.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, _
                           TITLEGRAPHTIMESERIES, GRAPHTIMESERIES, SPATIAL, _
                           SPATIOTEMPSPECS, SPATIOTEMPORAL
    End Select

    For idx = listNames.LowerBound To listNames.UpperBound
        loname = CStr(listNames.Item(idx))
        GetTableSettings loname, prefName, idCol, requiredCount

        'A targeted table can be absent from the worksheet, so the lookup is
        'the test.
        On Error Resume Next
            Set lo = sh.ListObjects(loname)
        On Error GoTo 0

        If lo Is Nothing Then
            LogInfo "ListObject " & loname & " not found during row management", checkingWarning
            GoTo NextTable
        End If

        Set tabl = CustomTable.Create(lo, idCol, prefName)

        If del Then
            tabl.RemoveRows totalCount:=requiredCount, forceShift:=True
        Else
            'AddRows renumbers the identifier column on the way out, so the
            'column is written once.
            tabl.AddRows insertShift:=True
        End If

NextTable:
        Set lo = Nothing
        Set tabl = Nothing
    Next idx
End Sub

'@label:instruction-cell
'@sub-title Resolve the cell carrying the row-management instruction
'@details
'The rest of the setup reaches this cell through the RNG_SelectTable named
'range, so this class asks for the name first and falls back to A1 when the
'workbook carries no such name. A name that resolves to another worksheet is
'refused, because a workbook-scoped name answers from any sheet.
'@param sh Worksheet. The analysis worksheet.
'@return Range. The single cell holding the instruction.
Private Function InstructionCell(ByVal sh As Worksheet) As Range
    Dim namedCell As Range

    'A workbook without the named range is a normal case here, so the lookup
    'is the test.
    On Error Resume Next
        Set namedCell = sh.Range(INSTRUCTIONRANGE)
    On Error GoTo 0

    If Not namedCell Is Nothing Then
        If namedCell.Worksheet Is sh Then
            Set InstructionCell = namedCell.Cells(1, 1)
            Exit Function
        End If
    End If

    Set InstructionCell = sh.Cells(1, 1)
End Function

'@label:resolve-selection-table
'@sub-title Resolve the ListObject that hosts a worksheet selection
'@details
'Validates that the target cell belongs to the analysis worksheet and
'intersects a ListObject. Logs a warning and returns Nothing when any
'validation check fails. Used by InsertRows and DeleteRows to identify
'the target table before delegating to CustomTable.
'@param targetCell Range. The user selection to resolve.
'@param actionLabel String. Human-readable action name for log messages
'(e.g. "insert", "delete").
'@return ListObject. The hosting ListObject, or Nothing when not found.
Private Function ResolveSelectionTable(ByVal targetCell As Range, ByVal actionLabel As String) As ListObject
    'Ensure the provided selection belongs to the analysis worksheet and intersects a listobject.
    Dim hostsheet As Worksheet
    Dim selectionSheet As Worksheet
    Dim normalizedAction As String

    normalizedAction = LCase$(actionLabel)
    Set hostsheet = Wksh()

    If targetCell Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the target selection is missing", checkingWarning
        Exit Function
    End If

    'A selection detached from any worksheet is a normal case here, so the
    'read is the test.
    On Error Resume Next
        Set selectionSheet = targetCell.Worksheet
    On Error GoTo 0

    If selectionSheet Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the worksheet could not be resolved", checkingWarning
        Exit Function
    End If

    If Not (selectionSheet Is hostsheet) Then
        LogInfo "Unable to " & normalizedAction & " rows because the selection belongs to " & selectionSheet.Name, checkingWarning
        Exit Function
    End If

    'A selection outside every table answers Nothing, so the read is the test.
    On Error Resume Next
        Set ResolveSelectionTable = targetCell.ListObject
    On Error GoTo 0

    If ResolveSelectionTable Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the selection is outside analysis tables", checkingWarning
    End If
End Function

'@label:table-settings
'@sub-title Derive prefix, identifier column, and minimum row counts
'@details
'Populates the three ByRef outputs based on the ListObject name. Only
'TIMESERIES, TITLEGRAPHTIMESERIES, GRAPHTIMESERIES, and SPATIOTEMPORAL
'have non-default settings; all other tables use empty strings and zero.
'The name is lower-cased before it is matched, because a table name read
'back from a worksheet can carry any spelling of the same name.
'@param loName String. Name of the target ListObject.
'@param prefix ByRef String. Populated with the identifier prefix.
'@param idColumn ByRef String. Populated with the identifier column name.
'@param minimumCount ByRef Long. Populated with the minimum row count.
Private Sub GetTableSettings(ByVal loName As String, _
                             ByRef prefix As String, _
                             ByRef idColumn As String, _
                             ByRef minimumCount As Long)
    prefix = vbNullString
    idColumn = vbNullString
    minimumCount = 0

    Select Case LCase$(loName)
        Case LCase$(TIMESERIES)
            prefix = "Series"
            idColumn = "series id"
            minimumCount = 2
        Case LCase$(TITLEGRAPHTIMESERIES)
            prefix = "Graph"
            idColumn = "graph id"
            minimumCount = 1
        Case LCase$(GRAPHTIMESERIES)
            minimumCount = 7
        Case LCase$(SPATIOTEMPORAL)
            minimumCount = 3
    End Select
End Sub

'@section DataExchange
'===============================================================================
'@description Import, export, translate, sort, and clean operations.

'@label:clean
'@sub-title Clean the analysis tables on the worksheet
'@details
'Iterates through the ListObjects on the analysis worksheet and delegates
'to CustomTable.Clean for each one. Tables listed in keepTables are left
'alone: Import passes the tables it is about to rewrite, and
'CustomTable.Import clears each of those itself.
'@param keepTables Optional Collection. Tables to skip, keyed on the
'lower-cased table name.
Private Sub Clean(Optional ByVal keepTables As Collection = Nothing)
    Dim hostLo As ListObject
    Dim hostTab As CustomTable
    Dim hostsheet As Worksheet
    Dim prefName As String
    Dim idCol As String
    Dim minimumCount As Long

    Set hostsheet = Wksh()
    For Each hostLo In hostsheet.ListObjects
        If TableNamed(keepTables, hostLo.Name) Is Nothing Then
            'The identifier column is handed over so CustomTable protects it
            'rather than clearing it as ordinary data.
            GetTableSettings hostLo.Name, prefName, idCol, minimumCount
            Set hostTab = CustomTable.Create(hostLo, idCol, prefName)
            hostTab.Clean
        End If
    Next
End Sub


'@label:import
'@sub-title Import analysis tables from an external worksheet
'@details
'Copies each ListObject and its associated headers from the source worksheet
'into the managed worksheet. Tables the source does not carry are cleared
'first; the rest are rewritten by CustomTable.Import with strict column
'matching. Source columns that found no home in the host table are reported,
'and so are required tables missing from the source. The worksheet is sorted
'at the end, which is also where the host worksheet is checked for missing
'tables.
'@param sourcesheet Worksheet. The worksheet providing the source analysis tables.
'@export
Public Sub Import(ByVal sourcesheet As Worksheet)

    Dim hostsheet As Worksheet
    Dim srcLo As ListObject
    Dim hostLo As ListObject
    Dim srcTab As CustomTable
    Dim hostTab As CustomTable
    Dim validSources As Collection
    Dim counter As Long
    Dim tabName As String
    Dim prefName As String
    Dim idCol As String
    Dim minimumCount As Long

    If sourcesheet Is Nothing Then
        LogInfo "Analysis imports aborted: source worksheet not provided", checkingError
        Exit Sub
    End If


    EnsureRequiredTables

    Set hostsheet = Wksh()

    'Gather the source tables this worksheet accepts, keyed on the lower-cased
    'name so the clean pass below can skip them.
    Set validSources = New Collection
    For Each srcLo In sourcesheet.ListObjects
        If IsValidListObject(srcLo) Then validSources.Add srcLo, LCase$(srcLo.Name)
    Next

    'Clear only the host tables the source has nothing to write into.
    Clean validSources

    For Each srcLo In validSources

        Set hostLo = hostsheet.ListObjects(srcLo.Name)
        GetTableSettings hostLo.Name, prefName, idCol, minimumCount

        Set srcTab = CustomTable.Create(srcLo)
        Set hostTab = CustomTable.Create(hostLo, idCol, prefName)

        hostTab.Import srcTab, insertShift:=True, strictColumnSearch:=True
        LogDroppedColumns hostTab, hostLo.Name
    Next

    'Look for unfound (missing listObject)
    For counter = this.requiredTablesCaches.LowerBound To this.requiredTablesCaches.UpperBound
        Set srcLo = Nothing
        'A required table can be absent from the source, so the lookup is the test.
        On Error Resume Next
            tabName = this.requiredTablesCaches.Item(counter)
            Set srcLo = sourcesheet.ListObjects(tabName)
        On Error GoTo 0

        If srcLo Is Nothing Then
            LogInfo "Table " & tabName & " not found in imported worksheet", checkingWarning
        End If
    Next

    Sort
End Sub


'@label:sort
'@sub-title Sort analysis ListObjects and enforce minimum rows
'@details
'Resolves the nine required tables in one pass, reports the ones the
'worksheet does not carry, then sorts each table by its primary key column
'(Section or Graph ID), applies secondary sorts where needed (Table order,
'Graph order), and trims trailing rows via CustomTable.RemoveRows. The
'identifier column of each table is handed over, so rows that move are
'renumbered to match their new position.
'@export
Public Sub Sort()
    Dim hostTables As Collection
    Dim Lo As ListObject
    Dim csTab As CustomTable
    Dim colName As String
    Dim counter As Long
    Dim loName As String
    Dim rowCount As Long
    Dim prefName As String
    Dim idCol As String
    Dim minimumCount As Long

    Set hostTables = ResolveAnalysisTables()
    EnsureAnalysisTables hostTables

    For counter = this.requiredTablesCaches.LowerBound To this.requiredTablesCaches.UpperBound
        loName = CStr(this.requiredTablesCaches.Item(counter))
        Set Lo = TableNamed(hostTables, loName)

        If Not (Lo Is Nothing) Then
            GetTableSettings loName, prefName, idCol, minimumCount
            Set csTab = CustomTable.Create(Lo, idCol, prefName)

            'Row count is the number of elements in the row that triggers deletion of the
            'entire row of the listObject
            rowCount = Switch(loName = TIMESERIES, 2, loName = GRAPHTIMESERIES, 4, True, 0)


            colName = IIf(((loName = GRAPHTIMESERIES) Or (loName = TITLEGRAPHTIMESERIES)), "Graph ID", "Section")

            csTab.Sort colName, directSort:=False
            If loName = TIMESERIES Then csTab.Sort "Table order", directSort:=True, strictSearch:=False
            If ((loName = GRAPHTIMESERIES) Or (loName = TITLEGRAPHTIMESERIES)) Then
                csTab.Sort "Graph order", directSort:=True, strictSearch:=False
            End If

            'Resize the listObject after sorting, which renumbers the ids
            csTab.RemoveRows totalCount:=rowCount, forceShift:=True
        End If

        Set Lo = Nothing
    Next
End Sub


'@label:export
'@sub-title Export the analysis worksheet to a workbook
'@details
'Duplicates the managed worksheet into the provided workbook. If a sheet
'with the same name already exists it is cleared and reused; otherwise a
'new sheet is created. The requested visibility is applied either way. Each
'ListObject is copied by value and the header rows above it are transferred
'with font formatting, up to four rows or up to the top of the sheet,
'whichever comes first.
'@param sourcewb Workbook. The workbook receiving the exported worksheet.
'@param Hide Optional Long. Excel visibility constant applied to the
'exported sheet. Defaults to xlSheetHidden.
'@export
Public Sub Export(ByVal sourcewb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)

    Dim hostsheet As Worksheet
    Dim hostLo As ListObject
    Dim sourcesheet As Worksheet
    Dim sourceRange As Range
    Dim counter As Long
    Dim maxOffset As Long
    Dim sourceLo As ListObject
    Dim tempRange As Range
    Dim hostHeader As Range
    Dim sourceHeader As Range


    If sourcewb Is Nothing Then
        LogInfo "Export aborted: destination workbook not provided", checkingWarning
        Exit Sub
    End If

    Set hostsheet = Wksh()

    'The export workbook may already carry the sheet, so the lookup is the test.
    On Error Resume Next
    Set sourcesheet = sourcewb.Worksheets(hostsheet.Name)
    On Error GoTo 0

    If sourcesheet Is Nothing Then
        Set sourcesheet = sourcewb.Worksheets.Add(after:=sourcewb.Worksheets(sourcewb.Worksheets.Count))
        sourcesheet.Name = hostsheet.Name
    Else
        LogInfo sourcesheet.Name & " already exists in the export workbook, will be cleared", checkingWarning
        sourcesheet.Cells.Clear
    End If

    'Every caller states the visibility it wants, so a reused sheet gets it too.
    sourcesheet.Visible = Hide


    For Each hostLo In hostsheet.ListObjects

        Set sourceRange = sourcesheet.Range(hostLo.Range.Address)
        sourceRange.Value = hostLo.Range.Value
        Set sourceLo = sourcesheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                   Source:=sourceRange, _
                                                   XlListObjectHasHeaders:=xlYes)
        sourceLo.Name = hostLo.Name
        sourceLo.TableStyle = hostLo.TableStyle

            'Stop at the top of the sheet: a table whose header sits on row 3
            'has only two rows above it.
            maxOffset = hostLo.HeaderRowRange.Row - 1
            If maxOffset > 4 Then maxOffset = 4

            For counter = 0 To maxOffset
                Set hostHeader = hostLo.HeaderRowRange.Offset(-counter)
                Set sourceHeader = sourceLo.HeaderRowRange.Offset(-counter)
                Set tempRange = hostHeader.Cells(1, 1)
                With tempRange.Font
                    sourceHeader.Font.Color = .Color
                    sourceHeader.Font.Size = .Size
                    sourceHeader.Font.Name = .Name
                    sourceHeader.Font.Bold = .Bold
                End With
                sourceHeader.Value = hostHeader.Value
            Next
    Next

    sourcesheet.UsedRange.EntireColumn.AutoFit
End Sub

'@label:translate
'@sub-title Translate analysis labels using the provided engine
'@details
'Iterates through all ListObjects on the analysis worksheet and translates
'known textual columns (Section, Table Title, Graph Title, Summary label,
'Choices, Label, Section (select)) through the translation object. The
'Summary function column is translated with formula awareness enabled.
'Each header row is read once and matched in memory, so a column a table
'does not carry costs nothing. Logs a note on completion and a warning when
'the translator is Nothing.
'@param TransObject TranslationObject. The translation engine to use.
'@export
Public Sub Translate(ByVal TransObject As TranslationObject)
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim csTab As CustomTable
    Dim columnsToTranslate As BetterArray
    Dim idx As Long
    Dim colName As String
    Dim colRng As Range
    Dim tableHeaders As Variant

    If TransObject Is Nothing Then
        LogInfo "Translation aborted: translation object not provided", checkingWarning
        Exit Sub
    End If

    Set sh = Wksh()
    Set columnsToTranslate = New BetterArray
    columnsToTranslate.Push "Section", "Table Title", "Graph Title", "Summary label", _
                            "Choices", "Label", "Section (select)"

    For Each lo In sh.ListObjects
        Set csTab = CustomTable.Create(lo)
        tableHeaders = TableHeaderNames(csTab)

        For idx = columnsToTranslate.LowerBound To columnsToTranslate.UpperBound
            colName = CStr(columnsToTranslate.Item(idx))
            If HasHeader(tableHeaders, colName) Then
                Set colRng = csTab.DataRange(colName, strictSearch:=True)
                If Not colRng Is Nothing Then TransObject.TranslateRange colRng
            End If
        Next idx

        If HasHeader(tableHeaders, "Summary function") Then
            Set colRng = csTab.DataRange("Summary function", strictSearch:=True)
            If Not colRng Is Nothing Then TransObject.TranslateRange colRng, containsFormula:=True
        End If
    Next lo

    LogInfo "Translation completed", checkingNote
End Sub

'@label:table-header-names
'@sub-title Read one table's header row into an array
'@details
'One worksheet crossing per table. A single-column table answers a scalar,
'so the value is wrapped into a one-element array and every caller reads
'the same shape.
'@param csTab CustomTable. The table whose header row is wanted.
'@return Variant. A zero-based array of header texts.
Private Function TableHeaderNames(ByVal csTab As CustomTable) As Variant
    Dim headerRng As Range
    Dim rawValues As Variant
    Dim headerTexts() As String
    Dim counter As Long

    Set headerRng = csTab.HeaderRange()
    If headerRng Is Nothing Then
        TableHeaderNames = Array()
        Exit Function
    End If

    rawValues = headerRng.Value

    If Not IsArray(rawValues) Then
        ReDim headerTexts(0 To 0)
        headerTexts(0) = CStr(rawValues)
        TableHeaderNames = headerTexts
        Exit Function
    End If

    ReDim headerTexts(0 To UBound(rawValues, 2) - LBound(rawValues, 2))
    For counter = LBound(rawValues, 2) To UBound(rawValues, 2)
        headerTexts(counter - LBound(rawValues, 2)) = CStr(rawValues(LBound(rawValues, 1), counter))
    Next counter

    TableHeaderNames = headerTexts
End Function

'@label:has-header
'@sub-title Test a header name against one table's header row
'@details
'Whole-string match without regard to case, the same rule
'CustomTable.DataRange applies when strictSearch is True.
'@param headerList Variant. Header texts read by TableHeaderNames.
'@param colName String. Column name to look for.
'@return Boolean. True when the table carries the column.
Private Function HasHeader(ByVal headerList As Variant, ByVal colName As String) As Boolean
    Dim counter As Long

    If Not IsArray(headerList) Then Exit Function
    If UBound(headerList) < LBound(headerList) Then Exit Function

    For counter = LBound(headerList) To UBound(headerList)
        If StrComp(CStr(headerList(counter)), colName, vbTextCompare) = 0 Then
            HasHeader = True
            Exit Function
        End If
    Next counter
End Function



'@section Checking
'===============================================================================
'@description Logging, validation, and error-raising support.

'@label:log-info
'@sub-title Log an informational or warning message
'@details
'Adds a trace entry to the internal Checking instance, initialising it
'on first use with a name derived from the class and worksheet. Each call
'increments the check counter.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level from CheckingScope. Defaults
'to checkingNote.
Private Sub LogInfo(ByVal label As String, Optional ByVal scope As Byte = checkingNote)
    'The trace name is built once, on the entry that creates the store.
    If this.internalChecks Is Nothing Then
        Set this.internalChecks = Checking.Create(CLASSNAME & " " & Wksh().Name & " trace")
    End If

    this.internalChecks.Add CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1
End Sub

'@label:log-dropped-columns
'@sub-title Report the source columns an import could not place
'@details
'CustomTable records every source column that found no matching header in
'the target table. Without this, a column renamed between two setup versions
'is dropped in silence.
'@param hostTab CustomTable. The table that has just been imported into.
'@param loName String. Name of that table, used in the message.
Private Sub LogDroppedColumns(ByVal hostTab As CustomTable, ByVal loName As String)
    Dim droppedColumns As BetterArray
    Dim columnList As String
    Dim idx As Long

    If Not hostTab.HasColumnsNotImported Then Exit Sub

    Set droppedColumns = hostTab.ImportColumnsNotFound
    If droppedColumns Is Nothing Then Exit Sub

    For idx = droppedColumns.LowerBound To droppedColumns.UpperBound
        If LenB(columnList) > 0 Then columnList = columnList & ", "
        columnList = columnList & CStr(droppedColumns.Item(idx))
    Next idx

    If LenB(columnList) = 0 Then Exit Sub

    LogInfo "Columns of " & loName & " left out of the import: " & columnList, checkingWarning
End Sub

'@label:hascheckings
'@prop-title Whether trace entries were recorded
'@details
'Provides a quick way for callers to know if diagnostics are available.
'Returns True when at least one trace entry has been logged via LogInfo.
'@return Boolean. True when trace entries exist.
Public Property Get HasCheckings() As Boolean
    HasCheckings = Not (this.internalChecks Is Nothing)
End Property

'@label:checkingvalues
'@prop-title Retrieve collected trace entries
'@details
'Returns the underlying Checking object when traces are available.
'Returns Nothing when no traces have been logged.
'@return Checking. The Checking instance, or Nothing.
Public Property Get CheckingValues() As Checking
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = this.internalChecks
End Property


'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASSNAME,
'providing a consistent stack trace across all methods in this class.
'@param errNumber Long. The ProjectError code to raise.
'@param errorMessage String. Human-readable description of the failure.
'@throws ProjectError. Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, CLASSNAME, "Error message : " & Chr$(34) & errorMessage & Chr$(34)
End Sub

'@section Helpers
'===============================================================================
'@description Private helper methods supporting the public API.

'@label:ensure-required-tables
'@sub-title Populate the required-tables cache
'@details
'Lazily initialises the requiredTablesCaches BetterArray with the names
'of all nine ListObjects expected on an analysis worksheet. Subsequent
'calls are no-ops once the cache is populated. The push order decides the
'order Import, Sort and EnsureAnalysisTables work in, so it is fixed.
Private Sub EnsureRequiredTables()
    If this.requiredTablesCaches Is Nothing Then
        Set this.requiredTablesCaches = New BetterArray
        this.requiredTablesCaches.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, GRAPHTIMESERIES, _
                TITLEGRAPHTIMESERIES, SPATIAL, SPATIOTEMPORAL, SPATIOTEMPSPECS
    End If
End Sub

'@label:is-required-table
'@sub-title Test a table name against the required-tables cache
'@details
'The comparison ignores case, because Excel does. A binary comparison here
'once made Import skip the whole global summary section without a word.
'@param loName String. Table name to look for.
'@return Boolean. True when the name is one of the nine required tables.
Private Function IsRequiredTable(ByVal loName As String) As Boolean
    Dim idx As Long

    EnsureRequiredTables

    For idx = this.requiredTablesCaches.LowerBound To this.requiredTablesCaches.UpperBound
        If StrComp(CStr(this.requiredTablesCaches.Item(idx)), loName, vbTextCompare) = 0 Then
            IsRequiredTable = True
            Exit Function
        End If
    Next idx
End Function

'@label:is-valid-list-object
'@sub-title Check if a ListObject is valid for this analysis
'@details
'Returns True when the ListObject name appears in the required-tables
'cache and the corresponding ListObject exists on the host worksheet.
'Used by Import to skip unrecognised tables from the source worksheet.
'@param Lo ListObject. The ListObject to validate.
'@return Boolean. True when the ListObject is valid for this analysis.
Private Function IsValidListObject(ByVal Lo As ListObject) As Boolean

    If Lo Is Nothing Then Exit Function
    If Not IsRequiredTable(Lo.Name) Then Exit Function

    Dim hostsheet As Worksheet
    Set hostsheet = Wksh()
    'The host worksheet may not carry the table, so the lookup is the test.
    On Error Resume Next
    IsValidListObject = Not (hostsheet.ListObjects(Lo.Name) Is Nothing)
    On Error GoTo 0
End Function

'@label:resolve-analysis-tables
'@sub-title Resolve the required tables of the host worksheet in one pass
'@details
'Nine lookups, once. The result is keyed on the lower-cased table name, so
'callers read a table by name without crossing to Excel again. A table the
'worksheet does not carry is left out of the collection.
'@return Collection. The ListObjects found, keyed on the lower-cased name.
Private Function ResolveAnalysisTables() As Collection
    Dim tables As Collection
    Dim sh As Worksheet
    Dim idx As Long
    Dim loName As String
    Dim Lo As ListObject

    EnsureRequiredTables

    Set tables = New Collection
    Set sh = Wksh()

    For idx = this.requiredTablesCaches.LowerBound To this.requiredTablesCaches.UpperBound
        loName = CStr(this.requiredTablesCaches.Item(idx))
        Set Lo = Nothing
        'A required table can be absent from the worksheet, so the lookup is
        'the test.
        On Error Resume Next
        Set Lo = sh.ListObjects(loName)
        On Error GoTo 0

        If Not Lo Is Nothing Then tables.Add Lo, LCase$(loName)
    Next idx

    Set ResolveAnalysisTables = tables
End Function

'@label:table-named
'@sub-title Read one table out of a resolved collection
'@details
'A Collection cannot be asked whether a key exists, so the lookup itself is
'the test and a miss answers Nothing.
'@param tables Collection. Tables keyed on the lower-cased name.
'@param loName String. Table name to read.
'@return ListObject. The table, or Nothing when the collection has none.
Private Function TableNamed(ByVal tables As Collection, ByVal loName As String) As ListObject
    If tables Is Nothing Then Exit Function

    On Error Resume Next
    Set TableNamed = tables(LCase$(loName))
    On Error GoTo 0
End Function

'@label:checktables
'@sub-title Report the expected analysis tables the worksheet does not carry
'@details
'Walks the required-tables cache and logs a warning for every name absent
'from the resolved collection.
'@param hostTables Collection. Tables resolved by ResolveAnalysisTables.
Private Sub EnsureAnalysisTables(ByVal hostTables As Collection)
    Dim idx As Long
    Dim loName As String

    EnsureRequiredTables

    For idx = this.requiredTablesCaches.LowerBound To this.requiredTablesCaches.UpperBound
        loName = CStr(this.requiredTablesCaches.Item(idx))
        'Log a warning when a required listobject is absent from the worksheet.
        If TableNamed(hostTables, loName) Is Nothing Then
            LogInfo "Missing listobject " & loName & " in analysis", checkingWarning
        End If
    Next idx
End Sub
