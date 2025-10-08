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
    sht As Worksheet
    hasChecks As Boolean
End Type

Private this As TAnalysis

'Checking support
Private checkCounter As Long
Private internalChecks As IChecking

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
'@param anaWksh Worksheet hosting the analysis tables
'@return IAnalysis
'@depends Checking, CustomTable
'@export
Public Function Create(ByVal anaWksh As Worksheet) As IAnalysis
    If anaWksh Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, _
                                         "Worksheet object not defined"

    With New Analysis
        Set .Wksh = anaWksh
        Set Create = .Self
    End With
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
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sht
End Property

'@label:wksh-set
'@sub-title Store the backing worksheet
'@param anaWksh Worksheet used to drive analysis automation.
Public Property Set Wksh(ByVal anaWksh As Worksheet)
    Set this.sht = anaWksh
End Property

'@section RowManagement
'===============================================================================

'@label:addrows
'@sub-title Add rows based on worksheet header selection
'@details
'Delegates to the row manager to append the required rows to each target table.
'@export
Public Sub AddRows()
    ManageRows
End Sub

'@label:removerows
'@sub-title Remove rows while preserving minimum entries
'@details
'Calls the row manager in deletion mode to trim each table to its required length.
'@export
Public Sub RemoveRows()
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

    Set sh = Wksh()
    'Capture the user instruction displayed in the worksheet header.
    selectionText = LCase$(Trim$(Replace(sh.Cells(1, 1).Value, "Add or remove rows of", vbNullString)))

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
            tabl.RemoveRows totalCount:=requiredCount
        Else
            tabl.AddRows
            If idCol <> vbNullString Then tabl.AddIds
        End If
NextTable:
        Set lo = Nothing
        Set tabl = Nothing
    Next idx
End Sub

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

'@label:import
'@sub-title Import analysis tables from an external worksheet
'@details
'Copies each listobject (and associated headers) from the source worksheet into
'the managed worksheet, preserving cell formatting before validating the result.
'@param fromWksh Worksheet providing the source analysis tables.
'@export
Public Sub Import(ByVal fromWksh As Worksheet)
    Dim targetSh As Worksheet
    Dim srcLo As ListObject
    Dim targetLo As ListObject
    Dim srcRng As Range
    Dim targetRng As Range

    If fromWksh Is Nothing Then
        LogInfo "Analysis imports aborted: source worksheet not provided", checkingError
        Exit Sub
    End If

    Set targetSh = Wksh()
    targetSh.Cells.Clear

    For Each srcLo In fromWksh.ListObjects
        Set srcRng = srcLo.Range
        Set targetRng = targetSh.Range(srcRng.Address)
        targetRng.Value = srcRng.Value
        targetRng.Interior.Color = srcRng.Interior.Color

        targetRng.Cells(-1, 1).Value = srcRng.Cells(-1, 1).Value
        targetRng.Cells(-1, 1).Font.Size = srcRng.Cells(-1, 1).Font.Size
        targetRng.Cells(-1, 1).Font.color = srcRng.Cells(-1, 1).Font.color
        targetRng.Cells(-1, 1).Font.Bold = True

        Set targetLo = targetSh.ListObjects.Add(xlSrcRange, source:=targetRng, _
                                              XlListObjectHasHeaders:=xlYes)
        targetLo.Name = srcLo.Name
        targetLo.TableStyle = srcLo.TableStyle
    Next

    CheckTablesRequirements

    LogInfo "Import completed", checkingNote

    targetSh.Cells(2, 3).Value = "analysis imported - " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
    targetSh.Cells(2, 3).Font.Color = RGB(176, 196, 222)

    'Sort the analysis worksheet
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
    Dim tabl As BetterArray 'Will contain the list of listObjects to sort
    Dim counter As Long
    Dim loName As String
    Dim rowCount As Long

    Set tabl = New BetterArray
    tabl.LowerBound = 1

    tabl.Push UNIVARIATE, BIVARIATE, TIMESERIES, SPATIAL, GRAPHTIMESERIES, _ 
               TITLEGRAPHTIMESERIES, SPATIOTEMPORAL, SPATIOTEMPSPECS

    Set sh = Wksh()

    For counter = 1 To tabl.Length

        loName = tabl.Items(counter)
        Set Lo = sh.ListObjects(loName)
        Set csTab = CustomTable.Create(Lo)

        'Row count is the number of elements in the row that triggers deletion of the
        'entire row of the listObject
        rowCount = Switch(loName = TIMESERIES, 2, loName = GRAPHTIMESERIES, 4, True, 0)
        colName = Switch( _ 
            loName = GRAPHTIMESERIES Or loName = TITLEGRAPHTIMESERIES, "Graph ID", _ 
            loName = SPATIOTEMPORAL, "Section (select)", _ 
            True, "Section" _ 
        )

        csTab.Sort colName, directSort:=True
        If loName = TIMESERIES Then csTab.Sort "Table order", directSort:=False

        'Resize the listObject
        csTab.RemoveRows totalCount:=rowCount
    Next
End Sub



'@label:export
'@sub-title Export the analysis worksheet to a workbook
'@details
'Duplicates the managed worksheet into the provided workbook so external
'processes can consume the analysis tables.
'@param toWkb Workbook receiving the exported worksheet.
'@export
Public Sub Export(ByVal toWkb As Workbook)
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet

    If toWkb Is Nothing Then
        LogInfo "Export aborted: destination workbook not provided", checkingWarning
        Exit Sub
    End If

    Set sourceSheet = Wksh()
    sourceSheet.Copy After:=toWkb.Worksheets(toWkb.Worksheets.Count)
    Set targetSheet = toWkb.Worksheets(toWkb.Worksheets.Count)
    targetSheet.Name = sourceSheet.Name
    LogInfo "Export completed to workbook " & toWkb.Name, checkingNote
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
Public Sub Translate(ByVal TransObject As ITranslationObject)
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

'@section Validation
'===============================================================================

'@label:checktables
'@details
'Ensures all expected analysis tables exist and records warnings for any that
'are missing so downstream processes can alert users.
Private Sub CheckTablesRequirements()
    Dim required As BetterArray
    Dim idx As Long
    Dim loName As String

    Set required = New BetterArray
    required.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, GRAPHTIMESERIES, _
                  TITLEGRAPHTIMESERIES, SPATIAL, SPATIOTEMPORAL, SPATIOTEMPSPECS

    For idx = required.LowerBound To required.UpperBound
        loName = CStr(required.Item(idx))
        On Error Resume Next
        'Log a warning when a required listobject is absent from the worksheet.
        If Wksh().ListObjects(loName) Is Nothing Then LogInfo "Missing listobject " & loName, checkingWarning
        On Error GoTo 0
    Next idx
End Sub

'@label:checkings
'@section Checkings
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
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checkingvalues
'@prop-title Retrieve collected trace entries
'@details
'Returns the underlying checking object when traces are available.
Public Property Get CheckingValues() As Object
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = internalChecks
End Property

'@section Errors
'===============================================================================

'@label:throwerror
'@details
'Raises a strongly typed project error so callers can react consistently.
'@param errNumber ProjectError enumeration value describing the failure.
'@param errorMessage String detailing the problem context.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, "Analysis", "Error message : " & Chr(34) & errorMessage
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

'@jump:RemoveRows
Private Sub IAnalysis_RemoveRows()
    RemoveRows
End Sub

'@jump:Import
Private Sub IAnalysis_Import(ByVal fromWksh As Worksheet)
    Import fromWksh
End Sub

'@jump:Export
Private Sub IAnalysis_Export(ByVal toWkb As Workbook)
    Export toWkb
End Sub

'@jump:Translate
Private Sub IAnalysis_Translate(ByVal TransObject As ITranslationObject)
    Translate TransObject
End Sub

'@jump:Sort
Private Sub IAnalysis_Sort()
    Sort
End Sub

'@jump:HasCheckings
Private Property Get IAnalysis_HasCheckings() As Boolean
    IAnalysis_HasCheckings = HasCheckings()
End Property

'@jump:CheckingValues
Private Property Get IAnalysis_CheckingValues() As Object
    Set IAnalysis_CheckingValues = CheckingValues()
End Property
