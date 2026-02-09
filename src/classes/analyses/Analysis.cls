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
'@prop-title Instantiate analysis controller
'@details
'Creates a predeclared analysis manager bound to the supplied worksheet so
'callers can immediately work with the analysis without additional setup.
'@param hostsheet Worksheet hosting the analysis tables
'@return IAnalysis
'@depends Checking, CustomTable
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

'@prop-title Current object instance
'@details
'Expose the current instance so callers can fluently retrieve the interface reference.
'@return IAnalysis
Public Property Get Self() As IAnalysis
    Set Self = Me
End Property

'@label:wksh
'@prop-title Backing worksheet
'@details
'Return the worksheet that stores the analysis tables manipulated by this class.
'@return Worksheet
Private Property Get Wksh() As Worksheet
    Set Wksh = this.hostsheet
End Property

'@label:wksh-set
'@sub-title Store the backing worksheet
'@param hostsheet Worksheet used to drive analysis automation.
Public Property Set Wksh(ByVal hostsheet As Worksheet)
    Set this.hostsheet = hostsheet
End Property

'@section RowManagement
'===============================================================================

'@label:addrows
'@sub-title Add rows based on worksheet header selection
'@details
'Delegates to the row manager to append the required rows to each target table.
'@export
Private Sub AddRows()
    ManageRows
End Sub

'@label:insertrows
'@sub-title Insert rows based on the supplied selection
'@details
'Determines which analysis table hosts the selection and mirrors the selection
'height inside that table. Worksheet row insertion is enabled by default to
'protect stacked listobjects.
'@param targetCell Range describing the selection to mirror.
'@param insertShift Optional Boolean. When True worksheet rows are inserted ahead of the resize.
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

'@label:DeleteRows
'@sub-title Delete rows intersecting the selection
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
'Calls the row manager in deletion mode to trim each table to its required length.
'@export
Private Sub RemoveRows()
    ManageRows True
End Sub

'@label:managerows
'@details
'Central routine that inspects the worksheet header instruction and adjusts
'analysis tables accordingly, either by adding or deleting rows.
'@param del Boolean. When True rows are removed; otherwise rows are added.
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

'@label:tablesettings
'@details
'Derives prefix, identifier column and minimum row counts for supported analysis tables.
'@param loName String identifying the targeted listobject.
'@param prefix String returned prefix applied to generated identifiers.
'@param idColumn String returned identifier column name.
'@param minimumCount Long returned minimum number of rows required.
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

'@section ImportExport
'===============================================================================

'@label:Clean
'@sub-title Clean the analysis tables in the current worksheet
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
'Copies each listobject (and associated headers) from the source worksheet into
'the managed worksheet, preserving cell formatting before validating the result.
'@param sourceSheet Worksheet providing the source analysis tables.
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
'@sub-title Sort analysis listobjects and enforce minimum rows
'@details
'Iterates through the known analysis tables, sorts them with the correct
'priorities and trims trailing rows so every table meets its structural rules.
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
'Duplicates the managed worksheet into the provided workbook so external
'processes can consume the analysis tables.
'@param sourceWb Workbook receiving the exported worksheet.
'@export
Private Sub Export(ByVal sourceWb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)

    Dim hostsheet As Worksheet
    Dim hostLo As ListObject
    Dim sourcesheet As Worksheet
    Dim sourceRange As Range
    Dim counter As Long
    Dim sourceLo As ListObject
   

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
        On Error Resume Next
            For counter = 0 To 3
                sourceLo.HeaderRowRange.Offset(-counter).Value = hostLo.HeaderRowRange.Offset(-counter).Value
                sourceLo.HeaderRowRange.Offset(-counter).Font.Color = hostLo.HeaderRowRange.Offset(-counter).Font.Color
                sourceLo.HeaderRowRange.Offset(-counter).Font.Size = hostLo.HeaderRowRange.Offset(-counter).Font.Size
                sourceLo.HeaderRowRange.Offset(-counter).Font.Name = hostLo.HeaderRowRange.Offset(-counter).Font.Name
                sourceLo.HeaderRowRange.Offset(-counter).Font.Bold = hostLo.HeaderRowRange.Offset(-counter).Font.Bold
            Next
            sourceSheet.Cells.EntireColumn.AutoFit
        On Error GoTo 0
    Next
End Sub

'@section Translation
'===============================================================================

'@label:translate
'@sub-title Translate analysis labels using the provided engine
'@details
'Pushes known textual columns through the translation object so the analysis
'worksheet reflects the currently selected language.
'@param TransObject ITranslationObject responsible for translating ranges.
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



'@label:checkings
'@section Checkings and Errors
'===============================================================================

'@label:loginfo
'@details
'Adds a trace entry to the internal checking object, initialising it when needed.
'@param label String message describing the event.
'@param scope Byte indicating CheckingScope severity. Defaults to note level.
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
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checkingvalues
'@prop-title Retrieve collected trace entries
'@details
'Returns the underlying checking object when traces are available.
Private Property Get CheckingValues() As Object
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = internalChecks
End Property


'@label:throwerror
'@details
'Raises a strongly typed project error so callers can react consistently.
'@param errNumber ProjectError enumeration value describing the failure.
'@param errorMessage String detailing the problem context.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, CLASSNAME, "Error message : " & Chr$(34) & errorMessage & Chr$(34)
End Sub

'@section Interface
'===============================================================================

'@jump:wksh
Private Property Get IAnalysis_Wksh() As Worksheet
    Set IAnalysis_Wksh = Wksh()
End Property


'@jump:AddRows
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

'@jump:RemoveRows
Private Sub IAnalysis_RemoveRows()
    RemoveRows
End Sub

Private Sub IAnalysis_ManageRows(Optional ByVal del As Boolean = False)
    ManageRows del
End Sub

'@jump:Import
Private Sub IAnalysis_Import(ByVal sourcesheet As Worksheet)
    Import sourcesheet
End Sub

'@jump:Export
Private Sub IAnalysis_Export(ByVal sourcewb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    Export sourcewb, Hide
End Sub

'@jump:Translate
Private Sub IAnalysis_Translate(ByVal translator As ITranslationObject)
    Translate translator
End Sub

'@jump:Sort
Private Sub IAnalysis_Sort()
    Sort
End Sub

'@jump:Clean
Private Sub IAnalysis_Clean()
    Clean
End Sub

Private Sub IAnalysis_EnsureAnalysisTables()
    EnsureAnalysisTables True
End Sub

'@jump:HasCheckings
Private Property Get IAnalysis_HasCheckings() As Boolean
    IAnalysis_HasCheckings = HasCheckings()
End Property

'@jump:CheckingValues
Private Property Get IAnalysis_CheckingValues() As Object
    Set IAnalysis_CheckingValues = CheckingValues()
End Property

'@section Helpers
'===============================================================================

Private Sub EnsureRequiredTables()
    If requiredTablesCaches Is Nothing Then
        Set requiredTablesCaches = New BetterArray
        requiredTablesCaches.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, GRAPHTIMESERIES, _
                TITLEGRAPHTIMESERIES, SPATIAL, SPATIOTEMPORAL, SPATIOTEMPSPECS
    End If
End Sub

'Check if a listObject is valid for the analysis
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
'@details
'Ensures all expected analysis tables exist and records warnings for any that
'are missing so downstream processes can alert users.
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
