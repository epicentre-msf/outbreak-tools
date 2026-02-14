VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Core analysis worksheet manager"
Option Explicit


'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Core analysis worksheet manager")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  UseMeaningfulName, HungarianNotation

Implements IAnalysis

'@class Analysis
'@description
'Manages an analysis worksheet within an outbreak linelist workbook. Provides
'row manipulation (add, insert, delete, remove), data exchange (import, export,
'sort, clean), translation of analysis labels, and validation of required
'ListObjects. Each instance is bound to a single worksheet and delegates table
'operations to CustomTable. Trace diagnostics are recorded via IChecking and
'can be retrieved through HasCheckings and CheckingValues.
'@depends CustomTable, Checking, ITranslationObject, BetterArray
'@version 2025-01-15
'@author Yves Amevoin

Private Type TAnalysis
    hostsheet As Worksheet
    hasChecks As Boolean
End Type

Private this As TAnalysis

'Checking support
Private checkCounter As Long
Private internalChecks As IChecking
Private requiredTablesCaches As BetterArray

'ListObject names used in analysis worksheets
Private Const GLOBALSUMMARY As String = "Tab_global_summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"
Private Const SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const SPATIOTEMPSPECS As String = "Tab_SpatioTemporal_Specs"
Private Const CLASSNAME As String = "Analysis"

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create an Analysis instance bound to a worksheet
'@details
'Factory method on the predeclared instance. Validates that the worksheet
'is not Nothing, binds a new Analysis to it, and populates the required-
'tables cache on the predeclared instance. Callers interact with the
'returned IAnalysis interface.
'@param hostsheet Worksheet. The worksheet hosting the analysis tables.
'@return IAnalysis. A fully initialised analysis controller.
'@throws ProjectError.ObjectNotInitialized When hostsheet is Nothing.
'@depends CustomTable, Checking
'@export
Public Function Create(ByVal hostsheet As Worksheet) As IAnalysis
    If hostsheet Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, _
                                         "Worksheet object not defined"
    'You can create analysis on a empty worksheet, without tables.
    With New Analysis
        Set .Wksh = hostsheet
        Set Create = .Self
    End With

    EnsureRequiredTables
End Function

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor so consuming code can fluently retrieve the interface
'reference from the predeclared Create method.
'@return IAnalysis. The current instance cast to the interface.
Public Property Get Self() As IAnalysis
    Set Self = Me
End Property

'@label:wksh
'@prop-title Backing analysis worksheet
'@details
'Returns the worksheet that stores the analysis tables manipulated by
'this class. Set during construction via the Create factory.
'@return Worksheet. The bound analysis worksheet.
Private Property Get Wksh() As Worksheet
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
'analysis table based on the instruction in cell A1 of the worksheet.
'@export
Private Sub AddRows()
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
Private Sub InsertRows(ByVal targetCell As Range, _
                       Optional ByVal insertShift As Boolean = True)
    Dim targetList As ListObject
    Dim targetTable As ICustomTable
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
'retrieves its settings, and delegates to CustomTable.DeleteRowsAt. The
'includeIds flag controls whether identifier columns are cleared, and
'forceShift controls whether worksheet rows are shifted up after deletion.
'@param targetCell Range. Anchor describing the rows to delete.
'@param includeIds Optional Boolean. When True, identifiers are cleared
'along with data. Defaults to True.
'@param forceShift Optional Boolean. When True, worksheet rows are shifted
'up after deletion. Defaults to True.
'@export
Private Sub DeleteRows(ByVal targetCell As Range, _
                       Optional ByVal includeIds As Boolean = True, _
                       Optional ByVal forceShift As Boolean = True)
    Dim targetList As ListObject
    Dim targetTable As ICustomTable
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
Private Sub RemoveRows()
    ManageRows True
End Sub

'@label:managerows
'@sub-title Central row-management dispatcher
'@details
'Reads the header instruction from cell A1 of the analysis worksheet and
'determines which ListObject(s) to target. When del is False, rows are
'appended via CustomTable.AddRows and identifiers are regenerated. When
'del is True, rows are removed via CustomTable.RemoveRows down to the
'required minimum. Falls back to processing all known tables when the
'header text does not match a specific table name.
'@param del Optional Boolean. When True, rows are removed; otherwise rows
'are added. Defaults to False.
Private Sub ManageRows(Optional ByVal del As Boolean = False)
    Dim listNames As BetterArray
    Dim selectionText As String
    Dim loname As String
    Dim prefName As String
    Dim idCol As String
    Dim requiredCount As Long
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim tabl As ICustomTable
    Dim idx As Long
    Dim headerValue As Variant

    Set sh = Wksh()
    'Capture the user instruction displayed in the worksheet header.
    headerValue = sh.Cells(1, 1).Value
    If IsError(headerValue) Then
        selectionText = vbNullString
    Else
        selectionText = LCase$(CStr(headerValue))
        selectionText = Trim$(Replace(selectionText, "add or remove rows of", vbNullString))
    End If

    Set listNames = New BetterArray

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
            tabl.AddRows insertShift:=True
            If idCol <> vbNullString Then tabl.AddIds
        End If

NextTable:
        Set lo = Nothing
        Set tabl = Nothing
    Next idx
End Sub

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
    Dim hostSheet As Worksheet
    Dim selectionSheet As Worksheet
    Dim normalizedAction As String

    normalizedAction = LCase$(actionLabel)
    Set hostSheet = Wksh()

    If targetCell Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the target selection is missing", checkingWarning
        Exit Function
    End If

    On Error Resume Next
        Set selectionSheet = targetCell.Worksheet
    On Error GoTo 0

    If selectionSheet Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the worksheet could not be resolved", checkingWarning
        Exit Function
    End If

    If Not (selectionSheet Is hostSheet) Then
        LogInfo "Unable to " & normalizedAction & " rows because the selection belongs to " & selectionSheet.Name, checkingWarning
        Exit Function
    End If

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

    Select Case loName
        Case TIMESERIES
            prefix = "Series"
            idColumn = "series id"
            minimumCount = 2
        Case TITLEGRAPHTIMESERIES
            prefix = "Graph"
            idColumn = "graph id"
            minimumCount = 1
        Case GRAPHTIMESERIES
            minimumCount = 7
        Case SPATIOTEMPORAL
            minimumCount = 3
    End Select
End Sub

'@section DataExchange
'===============================================================================
'@description Import, export, translate, sort, and clean operations.

'@label:clean
'@sub-title Clean all analysis tables on the worksheet
'@details
'Iterates through all ListObjects on the analysis worksheet and delegates
'to CustomTable.Clean for each one. Ensures the required-tables cache is
'populated before processing.
'@export
Private Sub Clean()
    Dim hostLo As ListObject
    Dim hostTab As ICustomTable
    Dim hostSheet As Worksheet
    Dim counter As Long
    EnsureRequiredTables
    Set hostSheet = Wksh()
    For Each hostLo in hostSheet.ListObjects
        Set hostTab = CustomTable.Create(hostLo)
        hostTab.Clean
    Next
End Sub


'@label:import
'@sub-title Import analysis tables from an external worksheet
'@details
'Copies each ListObject and its associated headers from the source worksheet
'into the managed worksheet. The method first cleans existing tables, then
'imports each valid ListObject via CustomTable.Import with strict column
'matching. Missing source tables are logged as warnings. After import the
'method validates all required tables and sorts them.
'@param sourcesheet Worksheet. The worksheet providing the source analysis tables.
'@export
Private Sub Import(ByVal sourcesheet As Worksheet)

    Dim hostsheet As Worksheet
    Dim srcLo As ListObject
    Dim hostLo As ListObject
    Dim srcTab As ICustomTable
    Dim hostTab As ICustomTable
    Dim counter As Long
    Dim tabName As String

    If sourcesheet Is Nothing Then
        LogInfo "Analysis imports aborted: source worksheet not provided", checkingError
        Exit Sub
    End If


    EnsureRequiredTables

    'Clean actual elements before imports
    Clean

    Set hostsheet = Wksh()

    For Each srcLo In sourcesheet.ListObjects

        If Not IsValidListObject(srcLo) Then  GoTo Continue

        Set srcTab = CustomTable.Create(srcLo)
        Set hostLo = hostsheet.ListObjects(srcLo.Name)
        Set hostTab = CustomTable.Create(hostLo)

        hostTab.Import srcTab, insertShift:=True, strictColumnSearch:=True
    Continue:
    Next

    'Look for unfound (missing listObject)
    For counter = requiredTablesCaches.LowerBound To requiredTablesCaches.UpperBound
        Set srcLo = Nothing
        On Error Resume Next
            tabName = requiredTablesCaches.Item(counter)
            Set srcLo = sourcesheet.ListObjects(tabName)
        On Error GoTo 0

        If srcLo Is Nothing Then
            LogInfo "Table " & tabName & " not found in imported worksheet", checkingWarning
        End If
    Next

    EnsureAnalysisTables
    Sort
End Sub


'@label:sort
'@sub-title Sort analysis ListObjects and enforce minimum rows
'@details
'Iterates through the required-tables cache, sorts each table by its
'primary key column (Section or Graph ID), applies secondary sorts where
'needed (Table order, Graph order), and trims trailing rows via
'CustomTable.RemoveRows to enforce structural rules.
'@export
Private Sub Sort()
    Dim sh As Worksheet
    Dim Lo As listObject
    Dim csTab As ICustomTable
    Dim colName As String
    Dim counter As Long
    Dim loName As String
    Dim rowCount As Long

    Set sh = Wksh()

    EnsureAnalysisTables


    For counter = requiredTablesCaches.LowerBound To requiredTablesCaches.UpperBound
        loName = requiredTablesCaches.Items(counter)

        On Error Resume Next
        Set Lo = sh.ListObjects(loName)
        On Error GoTo 0

        If Not (Lo Is Nothing) Then
            Set csTab = CustomTable.Create(Lo)

            'Row count is the number of elements in the row that triggers deletion of the
            'entire row of the listObject
            rowCount = Switch(loName = TIMESERIES, 2, loName = GRAPHTIMESERIES, 4, True, 0)


            colName = IIF(((loName = GRAPHTIMESERIES) Or (loName = TITLEGRAPHTIMESERIES)), "Graph ID", "Section")

            csTab.Sort colName, directSort:=False
            If loName = TIMESERIES Then csTab.Sort "Table order", directSort:=True, strictSearch:=False
            If ((loName = GRAPHTIMESERIES) Or (loName = TITLEGRAPHTIMESERIES)) Then
                csTab.Sort "Graph order", directSort:=True, strictSearch:=False
            End If

            'Resize the listObject after sorting
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
'new sheet is created. Each ListObject is copied by value and its header
'rows (up to 3 rows above) are transferred with font formatting.
'@param sourceWb Workbook. The workbook receiving the exported worksheet.
'@param Hide Optional Long. Excel visibility constant applied to the
'exported sheet. Defaults to xlSheetHidden.
'@export
Private Sub Export(ByVal sourceWb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)

    Dim hostsheet As Worksheet
    Dim hostLo As ListObject
    Dim sourcesheet As Worksheet
    Dim sourceRange As Range
    Dim counter As Long
    Dim sourceLo As ListObject
    Dim tempRange As Range


    If sourcewb Is Nothing Then
        LogInfo "Export aborted: destination workbook not provided", checkingWarning
        Exit Sub
    End If

    'Sort the actual analysis worksheet before exports

    Set hostsheet = Wksh()

    On Error Resume Next
    Set sourcesheet = sourceWb.Worksheets(hostsheet.Name)
    On Error GoTo 0

    If sourcesheet Is Nothing Then
        Set sourcesheet = sourcewb.Worksheets.Add(after:=sourcewb.Worksheets(sourcewb.Worksheets.Count))
        sourcesheet.Name = hostsheet.Name
        sourcesheet.Visible = Hide
    Else
        LogInfo sourcesheet.Name & " already exists in the export workbook, will be cleared", checkingWarning
        sourcesheet.Cells.Clear
    End If


    For Each hostLo In hostsheet.ListObjects

        Set sourceRange = sourcesheet.Range(hostLo.Range.Address)
        sourceRange.Value = hostLo.Range.Value
        sourcesheet.ListObjects.Add(xlSrcRange, sourceRange, , xlYes).Name = hostLo.Name
        Set sourceLo = sourcesheet.ListObjects(hostLo.Name)
        sourceLo.TableStyle = hostLo.TableStyle
            For counter = 0 To 4
                Set tempRange = hostLo.HeaderRowRange.Offset(-counter).Cells(1, 1)
                sourceLo.HeaderRowRange.Offset(-counter).Font.Color = tempRange.Font.Color
                sourceLo.HeaderRowRange.Offset(-counter).Font.Size = tempRange.Font.Size
                sourceLo.HeaderRowRange.Offset(-counter).Font.Name = tempRange.Font.Name
                sourceLo.HeaderRowRange.Offset(-counter).Font.Bold = tempRange.Font.Bold
                sourceLo.HeaderRowRange.Offset(-counter).Value = hostLo.HeaderRowRange.Offset(-counter).Value
            Next
    Next

    sourceSheet.Cells.EntireColumn.AutoFit
End Sub

'@label:translate
'@sub-title Translate analysis labels using the provided engine
'@details
'Iterates through all ListObjects on the analysis worksheet and translates
'known textual columns (Section, Table Title, Graph Title, Summary label,
'Choices, Label, Section (select)) through the translation object. The
'Summary function column is translated with formula awareness enabled.
'Logs a note on completion and a warning when the translator is Nothing.
'@param TransObject ITranslationObject. The translation engine to use.
'@export
Private Sub Translate(ByVal TransObject As ITranslationObject)
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim csTab As ICustomTable
    Dim columnsToTranslate As BetterArray
    Dim idx As Long
    Dim colName As String
    Dim colRng As Range

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
        For idx = columnsToTranslate.LowerBound To columnsToTranslate.UpperBound
            colName = CStr(columnsToTranslate.Item(idx))
            Set colRng = csTab.DataRange(colName, strictSearch:=True)
            If Not colRng Is Nothing Then TransObject.TranslateRange colRng
        Next idx

        Set colRng = csTab.DataRange("Summary function", strictSearch:=True)
        If Not colRng Is Nothing Then TransObject.TranslateRange colRng, containsFormula:=True
    Next lo

    LogInfo "Translation completed", checkingNote
End Sub



'@section Checking
'===============================================================================
'@description Logging, validation, and error-raising support.

'@label:log-info
'@sub-title Log an informational or warning message
'@details
'Adds a trace entry to the internal IChecking instance, initialising it
'on first use with a name derived from the class and worksheet. Each call
'increments the check counter and sets the hasChecks flag.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level from CheckingLogType. Defaults
'to checkingNote.
Private Sub LogInfo(ByVal label As String, Optional ByVal scope As Byte = checkingNote)
    Dim checkName As String

    checkName = CLASSNAME & " " & Wksh().Name & " trace"
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)

    internalChecks.Add CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1
    this.hasChecks = True
End Sub

'@label:hascheckings
'@prop-title Whether trace entries were recorded
'@details
'Provides a quick way for callers to know if diagnostics are available.
'Returns True when at least one trace entry has been logged via LogInfo.
'@return Boolean. True when trace entries exist.
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checkingvalues
'@prop-title Retrieve collected trace entries
'@details
'Returns the underlying IChecking object when traces are available.
'Returns Nothing when no traces have been logged.
'@return Object. The IChecking instance, or Nothing.
Private Property Get CheckingValues() As Object
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = internalChecks
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

'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the IAnalysis contract.
'See the corresponding Private/Public members above for full documentation.

Private Property Get IAnalysis_Wksh() As Worksheet
    Set IAnalysis_Wksh = Wksh()
End Property


Private Sub IAnalysis_AddRows()
    AddRows
End Sub

Private Sub IAnalysis_InsertRows(ByVal targetCell As Range, _
                                 Optional ByVal insertShift As Boolean = True)
    InsertRows targetCell, insertShift
End Sub

Private Sub IAnalysis_DeleteRows(ByVal targetCell As Range, _
                                 Optional ByVal includeIds As Boolean = True, _
                                 Optional ByVal forceShift As Boolean = True)
    DeleteRows targetCell, includeIds, forceShift
End Sub

Private Sub IAnalysis_RemoveRows()
    RemoveRows
End Sub

Private Sub IAnalysis_ManageRows(Optional ByVal del As Boolean = False)
    ManageRows del
End Sub

Private Sub IAnalysis_Import(ByVal sourcesheet As Worksheet)
    Import sourcesheet
End Sub

Private Sub IAnalysis_Export(ByVal sourcewb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    Export sourcewb, Hide
End Sub

Private Sub IAnalysis_Translate(ByVal translator As ITranslationObject)
    Translate translator
End Sub

Private Sub IAnalysis_Sort()
    Sort
End Sub

Private Sub IAnalysis_Clean()
    Clean
End Sub

Private Sub IAnalysis_EnsureAnalysisTables()
    EnsureAnalysisTables True
End Sub

Private Property Get IAnalysis_HasCheckings() As Boolean
    IAnalysis_HasCheckings = HasCheckings()
End Property

Private Property Get IAnalysis_CheckingValues() As Object
    Set IAnalysis_CheckingValues = CheckingValues()
End Property

'@section Helpers
'===============================================================================
'@description Private helper methods supporting the public API.

'@label:ensure-required-tables
'@sub-title Populate the required-tables cache
'@details
'Lazily initialises the requiredTablesCaches BetterArray with the names
'of all nine ListObjects expected on an analysis worksheet. Subsequent
'calls are no-ops once the cache is populated.
Private Sub EnsureRequiredTables()
    If requiredTablesCaches Is Nothing Then
        Set requiredTablesCaches = New BetterArray
        requiredTablesCaches.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, GRAPHTIMESERIES, _
                TITLEGRAPHTIMESERIES, SPATIAL, SPATIOTEMPORAL, SPATIOTEMPSPECS
    End If
End Sub

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
    EnsureRequiredTables
    If Not requiredTablesCaches.Includes(Lo.Name) Then Exit Function

    Dim hostsheet As Worksheet
    Set hostsheet = Wksh()
    On Error Resume Next
    IsValidListObject = Not (hostsheet.ListObjects(Lo.Name) Is Nothing)
    On Error GoTo 0
End Function

'@label:checktables
'@sub-title Ensure all expected analysis tables are present
'@details
'Iterates through the required-tables cache and checks that each
'ListObject exists on the analysis worksheet. When raiseError is True,
'missing tables raise a ProjectError.ElementNotFound; otherwise a warning
'is logged via LogInfo.
'@param raiseError Optional Boolean. When True, raises an error for
'missing tables instead of logging. Defaults to False.
'@throws ProjectError.ElementNotFound When raiseError is True and a
'required ListObject is missing.
Private Sub EnsureAnalysisTables(Optional ByVal raiseError As Boolean = False)
    Dim idx As Long
    Dim loName As String
    Dim Lo As ListObject

    EnsureRequiredTables

    For idx = requiredTablesCaches.LowerBound To requiredTablesCaches.UpperBound
        loName = CStr(requiredTablesCaches.Item(idx))
        On Error Resume Next
        Set Lo = Wksh().ListObjects(loName)
        On Error GoTo 0
        'Log a warning when a required listobject is absent from the worksheet.
        If  Lo Is Nothing Then
            If raiseError Then
                ThrowError ProjectError.ElementNotFound, "Missing listobject " & loName & " in analysis"
            Else
                LogInfo "Missing listobject " & loName & " in analysis", checkingWarning
            End If
        End If
        Set Lo = Nothing
    Next idx
End Sub
