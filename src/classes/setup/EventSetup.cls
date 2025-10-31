VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Workbook-level setup event coordinator"
Option Explicit

'@Folder("Setup")
'@ModuleDescription("Encapsulates workbook level setup events and caches for high-frequency handlers")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends BetterArray, HiddenNames, Passwords, UpdatedValues, DropdownLists, LLdictionary, LLVariables, CustomTable, ChoiceFormula, RibbonDev, ProjectError

Implements IEventSetup

'@section State
'===============================================================================
Private Type TEventSetup
    hostBook        As Workbook          'Workbook owning the setup sheets
    updater         As IUpdatedValues    'Registry tracker for watched columns
    dropdowns       As IDropdownLists    'Cached dropdown manager
    ribbon          As IRibbonUI         'Cached ribbon instance when available
    tsTable         As ICustomTable      'Analysis time-series table wrapper
    graphTable      As ICustomTable      'Analysis graph table wrapper
    spatioTable     As ICustomTable      'Analysis spatio-temporal table wrapper
    spatioSpecsTable As ICustomTable     'Spatio-temporal specification table wrapper
    variables       As ILLVariables      'Dictionary variable helper
    dictionary      As ILLdictionary     'Cached dictionary reference
End Type

Private this As TEventSetup

'@section Constants
'===============================================================================
Private Const SHEET_UPDATED          As String = "__updated"
Private Const SHEET_DICTIONARY       As String = "Dictionary"
Private Const SHEET_CHOICES          As String = "Choices"
Private Const SHEET_ANALYSIS         As String = "Analysis"
Private Const SHEET_DROPDOWN         As String = "__variables"
Private Const SHEET_PASSWORDS        As String = "__pass"
Private Const COUNTER_NAME           As String = "_SetupTranslationsCounter"

Private Const LIST_TS_DATA           As String = "Tab_TimeSeries_Analysis"
Private Const LIST_GRAPH_TS          As String = "Tab_Graph_TimeSeries"
Private Const LIST_SPATIO_TEMPORAL   As String = "Tab_SpatioTemporal_Analysis"
Private Const LIST_SPATIO_TEMPORAL_SPECS As String = "Tab_SpatioTemporal_Specs"

Private Const DICT_START_ROW         As Long = 5
Private Const DICT_START_COLUMN      As Long = 1
Private Const CHOICES_START_ROW      As Long = 4
Private Const CHOICES_START_COLUMN   As Long = 1

Private Const DROPDOWN_PREFIX        As String = "dropdown_"

'@section Factory helpers
'===============================================================================
Public Function Create(ByVal hostBook As Workbook) As IEventSetup
    Dim instance As EventSetup

    Set instance = New EventSetup
    instance.Configure hostBook
    Set Create = instance.Self
End Function

Public Property Get Self() As IEventSetup
    Set Self = Me
End Property

'@section IEventSetup implementation
'===============================================================================
Private Sub IEventSetup_Configure(ByVal hostBook As Workbook)
    If (hostBook Is Nothing) Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook is required to configure EventSetup"
    End If

    Set this.hostBook = hostBook
    ResetInternalCaches
End Sub

Private Sub IEventSetup_OnWorkbookOpen()
    EnsureUpdatedValues
    If Not (this.updater Is Nothing) Then
        On Error Resume Next
            this.updater.SwitchTagsToYes
        On Error GoTo 0
    End If

    ResetTranslationCounter

    On Error Resume Next
        Application.FormatStaleValues = False
    On Error GoTo 0
End Sub

Private Sub IEventSetup_OnSheetActivate(ByVal sh As Worksheet)
    If (sh Is Nothing Then) Exit Sub

    EnsureRibbon
    If Not (this.ribbon Is Nothing) Then
        this.ribbon.InvalidateControl "btnDelLoRows"
        this.ribbon.InvalidateControl "btnDelLoCol"
        this.ribbon.InvalidateControl "btnSort"
    End If

    If StrComp(sh.Name, SHEET_ANALYSIS, vbTextCompare) = 0 Then
        UpdateAnalysisDropdowns True
    End If
End Sub

Private Sub IEventSetup_OnSheetChange(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    EnsureUpdatedValues
    If Not (this.updater Is Nothing) Then
        this.updater.CheckUpdate sh, target
    End If

    If StrComp(sh.Name, SHEET_ANALYSIS, vbTextCompare) <> 0 Then Exit Sub

    EnsureAnalysisTable this.tsTable, LIST_TS_DATA, "title"
    EnsureAnalysisTable this.graphTable, LIST_GRAPH_TS, "series title"
    EnsureAnalysisTable this.spatioTable, LIST_SPATIO_TEMPORAL, "section"

    If target.ListObject Is Nothing Then Exit Sub

    CalculateAnalysis

    If target.Rows.Count > 1 Then Exit Sub
    If target.Columns.Count > 1 Then Exit Sub

    AddChoicesDropdown target
    AddGeoDropdown target
End Sub

Private Sub IEventSetup_UpdateAnalysisDropdowns(Optional ByVal forceUpdate As Boolean = False)
    EnsureUpdatedValues
    If (this.updater Is Nothing) Then Exit Sub

    EnsureDictionary
    If this.dictionary Is Nothing Then Exit Sub

    EnsureDropdownValues
    If this.drop Is Nothing Then Exit Sub

    If forceUpdate Or _
       this.updater.IsUpdated("control_details") Or _
       this.updater.IsUpdated("variable_name") Then

        On Error Resume Next
            this.drop.Update this.dictionary.GeoVars, "__geo_vars"
            this.drop.Update this.dictionary.ChoicesVars, "__choice_vars"
            this.drop.Update this.dictionary.SpecialVars("hf"), "__hfonly_vars"
            this.drop.Update this.dictionary.SpecialVars("geo"), "__geoonly_vars"
        On Error GoTo 0
    End If

    If forceUpdate Or _
       this.updater.IsUpdated("variable_type") Or _
       this.updater.IsUpdated("variable_name") Then

        On Error Resume Next
            this.drop.Update this.dictionary.TimeVars, "__time_vars"
        On Error GoTo 0
    End If

    EnsureVars
    If Not this.variables Is Nothing Then
        On Error Resume Next
            this.variables.InvalidateCaches
        On Error GoTo 0
    End If
End Sub

Private Sub IEventSetup_ResetCaches()
    ResetInternalCaches
End Sub

Private Sub IEventSetup_RecalculateAnalysis()
    CalculateAnalysis
End Sub

Private Sub IEventSetup_ResetTranslationCounter()
    ResetTranslationCounter
End Sub

Private Function IEventSetup_BuildTimeSeriesHeader(ByVal timeVar As String, _
                                                   ByVal groupVar As String, _
                                                   ByVal sumLabel As String) As String
    IEventSetup_BuildTimeSeriesHeader = BuildTimeSeriesHeaderInternal(timeVar, groupVar, sumLabel)
End Function

Private Function IEventSetup_AnalysisGraphValue(ByVal graphTitle As String, _
                                                Optional ByVal graphCol As String = "Graph ID") As String
    IEventSetup_AnalysisGraphValue = AnalysisGraphValueInternal(graphTitle, graphCol)
End Function

Private Function IEventSetup_AnalysisTimeSeriesValue(ByVal seriesTitle As String, _
                                                     Optional ByVal tsColumn As String = "Series ID") As String
    IEventSetup_AnalysisTimeSeriesValue = AnalysisTimeSeriesValueInternal(seriesTitle, tsColumn)
End Function

Private Function IEventSetup_SpatioTemporalSpecValue(ByVal sectionName As String, _
                                                     Optional ByVal specColumn As String = "N geo max") As String
    IEventSetup_SpatioTemporalSpecValue = SpatioTemporalSpecValueInternal(sectionName, specColumn)
End Function

'@section Internal helpers
'===============================================================================
Private Sub ResetInternalCaches()
    Set this.updater = Nothing
    Set this.drop = Nothing
    Set this.ribbon = Nothing
    Set this.tsTable = Nothing
    Set this.graphTable = Nothing
    Set this.spatioTable = Nothing
    Set this.spatioSpecsTable = Nothing
    Set this.variables = Nothing
    Set this.dictionary = Nothing
End Sub

Private Sub EnsureUpdatedValues()
    If Not (this.updater Is Nothing) Then Exit Sub

    Dim registrySheet As Worksheet

    registrySheet = Nothing
    Set registrySheet = SheetOrNothing(SHEET_UPDATED)
    If registrySheet Is Nothing Then Exit Sub

    On Error Resume Next
        Set this.updater = UpdatedValues.Create(registrySheet)
    On Error GoTo 0
End Sub

Private Sub EnsureDropdownValues()
    If Not (this.drop Is Nothing) Then Exit Sub

    Dim dropSheet As Worksheet

    Set dropSheet = SheetOrNothing(SHEET_DROPDOWN)
    If (dropSheet Is Nothing) Then Exit Sub

    On Error Resume Next
        Set this.drop = DropdownLists.Create(dropSheet, DROPDOWN_PREFIX)
    On Error GoTo 0
End Sub

Private Sub EnsureRibbon()
    If Not (this.ribbon Is Nothing) Then Exit Sub
    On Error Resume Next
        Set this.ribbon = RibbonDev.ActualRibbon()
    On Error GoTo 0
End Sub

Private Sub EnsureDictionary()
    If Not (this.dictionary Is Nothing) Then Exit Sub

    Dim dictSheet As Worksheet

    Set dictSheet = SheetOrNothing(SHEET_DICTIONARY)
    If dictSheet Is Nothing Then Exit Sub

    On Error Resume Next
        Set this.dictionary = LLdictionary.Create(dictSheet, DICT_START_ROW, DICT_START_COLUMN)
    On Error GoTo 0
End Sub

Private Sub EnsureVars()
    If Not (this.variables Is Nothing) Then Exit Sub

    EnsureDictionary
    If this.dictionary Is Nothing Then Exit Sub

    On Error Resume Next
        Set this.variables = LLVariables.Create(this.dictionary)
    On Error GoTo 0
End Sub

'Ensure the cached analysis tables are created lazily and reused.
Private Function EnsureAnalysisTable(ByRef tableRef As ICustomTable, _
                                     ByVal listName As String, _
                                     ByVal keyColumn As String) As Boolean
    If Not (tableRef Is Nothing) Then
        EnsureAnalysisTable = True
        Exit Function
    End If

    Dim lo As ListObject
    Set lo = AnalysisListObject(listName)
    If lo Is Nothing Then Exit Function

    On Error Resume Next
        Set tableRef = CustomTable.Create(lo, keyColumn)
    On Error GoTo 0

    EnsureAnalysisTable = Not (tableRef Is Nothing)
End Function

'Calculate the associated listobject once the cache is primed.
Private Sub EnsureTableCalculated(ByRef tableRef As ICustomTable, _
                                  ByVal listName As String, _
                                  ByVal keyColumn As String)
    If Not EnsureAnalysisTable(tableRef, listName, keyColumn) Then Exit Sub

    On Error Resume Next
        tableRef.DataRange.Calculate
    On Error GoTo 0
End Sub

Private Function AnalysisListObject(ByVal loName As String) As ListObject
    Dim analysisSheet As Worksheet
    Dim lo As ListObject

    Set analysisSheet = SheetOrNothing(SHEET_ANALYSIS)
    If analysisSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set lo = analysisSheet.ListObjects(loName)
    On Error GoTo 0

    Set AnalysisListObject = lo
End Function

Private Function SheetOrNothing(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    If this.hostBook Is Nothing Then Exit Function

    On Error Resume Next
        Set sh = this.hostBook.Worksheets(sheetName)
    On Error GoTo 0

    Set SheetOrNothing = sh
End Function

Private Function PasswordManager() As IPasswords
    Dim passSheet As Worksheet

    Set passSheet = SheetOrNothing(SHEET_PASSWORDS)
    If passSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set PasswordManager = Passwords.Create(passSheet)
    On Error GoTo 0
End Function

'Temporarily unlock the provided sheet and hand back the password helper.
Private Function BeginSheetUpdate(ByVal sheetName As String) As IPasswords
    Dim pass As IPasswords

    Set pass = PasswordManager()
    On Error Resume Next
        If Not pass Is Nothing Then pass.UnProtect sheetName
    On Error GoTo 0

    Set BeginSheetUpdate = pass
End Function

'Restore sheet protection previously removed via BeginSheetUpdate.
Private Sub EndSheetUpdate(ByVal pass As IPasswords, ByVal sheetName As String)
    On Error Resume Next
        If Not pass Is Nothing Then pass.Protect sheetName
    On Error GoTo 0
End Sub

Private Sub CalculateAnalysis()
    EnsureTableCalculated this.tsTable, LIST_TS_DATA, "title"
    EnsureTableCalculated this.graphTable, LIST_GRAPH_TS, "series title"
    EnsureTableCalculated this.spatioTable, LIST_SPATIO_TEMPORAL, "section"
    EnsureTableCalculated this.spatioSpecsTable, LIST_SPATIO_TEMPORAL_SPECS, "section"
End Sub

Private Sub ResetTranslationCounter()
    Dim registrySheet As Worksheet
    Dim namesManager As IHiddenNames

    Set registrySheet = SheetOrNothing(SHEET_UPDATED)
    If registrySheet Is Nothing Then Exit Sub

    On Error Resume Next
        Set namesManager = HiddenNames.Create(registrySheet)
        namesManager.EnsureName COUNTER_NAME, CLng(0), HiddenNameTypeLong
        namesManager.SetValue COUNTER_NAME, CLng(0)
    On Error GoTo 0
End Sub

Private Function BuildTimeSeriesHeaderInternal(ByVal timeVar As String, _
                                               ByVal groupVar As String, _
                                               ByVal sumLabel As String) As String
    Const SEP As String = " " & ChrW(9472) & " "
    Dim header As String
    Dim timeLabel As String
    Dim groupLabel As String

    header = sumLabel
    timeLabel = VariableLabel(timeVar)
    If LenB(timeLabel) > 0 Then
        header = header & SEP & timeLabel
    End If

    If LenB(groupVar) > 0 Then
        groupLabel = VariableLabel(groupVar)
        If LenB(groupLabel) > 0 Then
            header = header & SEP & groupLabel
        End If
    End If

    BuildTimeSeriesHeaderInternal = header
End Function

Private Function AnalysisGraphValueInternal(ByVal graphTitle As String, _
                                             ByVal graphCol As String) As String
    AnalysisGraphValueInternal = AnalysisTableValue(this.graphTable, _
                                                    LIST_GRAPH_TS, _
                                                    "series title", _
                                                    graphCol, _
                                                    graphTitle)
End Function

Private Function AnalysisTimeSeriesValueInternal(ByVal seriesTitle As String, _
                                                  ByVal tsColumn As String) As String
    AnalysisTimeSeriesValueInternal = AnalysisTableValue(this.tsTable, _
                                                         LIST_TS_DATA, _
                                                         "title", _
                                                         tsColumn, _
                                                         seriesTitle)
End Function

Private Function SpatioTemporalSpecValueInternal(ByVal sectionName As String, _
                                                 ByVal specColumn As String) As String
    SpatioTemporalSpecValueInternal = AnalysisTableValue(this.spatioSpecsTable, _
                                                         LIST_SPATIO_TEMPORAL_SPECS, _
                                                         "section", _
                                                         specColumn, _
                                                         sectionName)
End Function

Private Function VariableLabel(ByVal varName As String) As String
    Dim label As String

    If LenB(varName) = 0 Then Exit Function

    EnsureVars
    If this.variables Is Nothing Then
        VariableLabel = varName
        Exit Function
    End If

    On Error Resume Next
        label = this.variables.Value("Main Label", varName)
    On Error GoTo 0

    If LenB(label) = 0 Then label = varName
    VariableLabel = label
End Function

Private Function SafeTableValue(ByVal tableRef As ICustomTable, _
                                ByVal columnName As String, _
                                ByVal keyName As String) As String
    Dim value As Variant

    If tableRef Is Nothing Then Exit Function

    On Error Resume Next
        value = tableRef.Value(colName:=columnName, keyName:=keyName)
    On Error GoTo 0

    If IsError(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    SafeTableValue = CStr(value)
End Function

'Shared lookup helper for analysis listobjects.
Private Function AnalysisTableValue(ByRef tableRef As ICustomTable, _
                                    ByVal listName As String, _
                                    ByVal keyColumn As String, _
                                    ByVal valueColumn As String, _
                                    ByVal keyValue As String) As String
    If Not EnsureAnalysisTable(tableRef, listName, keyColumn) Then Exit Function
    AnalysisTableValue = SafeTableValue(tableRef, valueColumn, keyValue)
End Function

Private Sub AddGeoDropdown(ByVal target As Range)
    Dim geoCell As Range
    Dim spatialType As String
    Dim pass As IPasswords
    Const ANALYSIS_SHEET As String = SHEET_ANALYSIS

    If Not EnsureAnalysisTable(this.spatioTable, LIST_SPATIO_TEMPORAL, "section") Then Exit Sub
    EnsureDropdownValues
    If this.drop Is Nothing Then Exit Sub

    If this.spatioTable.CellRange("section", target.Row) Is Nothing Then Exit Sub

    Set geoCell = this.spatioTable.CellRange("geo", target.Row)
    If geoCell Is Nothing Then Exit Sub

    spatialType = CStr(this.spatioTable.Value("spatial type", target.Value))

    Set pass = BeginSheetUpdate(ANALYSIS_SHEET)
    On Error GoTo Cleanup
        geoCell.ClearContents

        Select Case LCase$(spatialType)
            Case "hf"
                this.drop.SetValidation geoCell, "__hfonly_vars"
            Case "geo"
                this.drop.SetValidation geoCell, "__geoonly_vars"
        End Select
Cleanup:
    EndSheetUpdate pass, ANALYSIS_SHEET
End Sub

Private Sub AddChoicesDropdown(ByVal target As Range)
    Dim seriesRange As Range
    Dim columnVariable As String
    Dim seriesValue As String
    Dim dropArray As BetterArray
    Dim choiceName As String
    Dim listName As String
    Dim choiceCell As Range
    Dim valueCell As Range
    Dim pass As IPasswords

    If Not EnsureAnalysisTable(this.graphTable, LIST_GRAPH_TS, "series title") Then Exit Sub
    If Not EnsureAnalysisTable(this.tsTable, LIST_TS_DATA, "title") Then Exit Sub
    EnsureDropdownValues

    If this.graphTable Is Nothing Then Exit Sub
    If this.tsTable Is Nothing Then Exit Sub
    If this.drop Is Nothing Then Exit Sub

    Set seriesRange = this.graphTable.DataRange("series title")
    If seriesRange Is Nothing Then Exit Sub
    If Intersect(target, seriesRange) Is Nothing Then Exit Sub

    seriesValue = CStr(target.Value)
    On Error Resume Next
        columnVariable = CStr(this.graphTable.Value("column", seriesValue))
    On Error GoTo 0
    If columnVariable = vbNullString Then Exit Sub
    dropArray = GetCategories(columnVariable, seriesValue, choiceName)

    If dropArray Is Nothing Then Exit Sub
    If dropArray.Length = 0 Then Exit Sub

    listName = choiceName & "__"

    If this.drop.Exists(listName) Then
        this.drop.Update dropArray, listName
    Else
        this.drop.Add dropArray, listName
    End If

    Set choiceCell = this.graphTable.CellRange("choice", target.Row)
    Set valueCell = this.graphTable.CellRange("values or percentages", target.Row)

    Set pass = BeginSheetUpdate(SHEET_ANALYSIS)
    On Error GoTo Cleanup
        If Not choiceCell Is Nothing Then
            choiceCell.ClearContents
            this.drop.SetValidation choiceCell, listName, ignoreBlank:=False
            FormatLockCell choiceCell, False
        End If

        If Not valueCell Is Nothing Then
            this.drop.SetValidation valueCell, "__perc_val"
            FormatLockCell valueCell, False
        End If
Cleanup:
    EndSheetUpdate pass, SHEET_ANALYSIS
End Sub

Private Function GetCategories(ByVal columnVariable As String, _
                               ByVal seriesValue As String, _
                               ByRef choiceName As String) As BetterArray
    Dim categoryList As BetterArray
    Dim controlDetails As String
    Dim choiceResolver As Object
    Dim addTotalFlag As String

    Set categoryList = New BetterArray

    EnsureVars
    If this.variables Is Nothing Then GoTo ExitFunction

    controlDetails = Trim$(UCase$(this.variables.Value("Control Details", columnVariable)))
    If controlDetails = vbNullString Then GoTo ExitFunction

    If Left$(controlDetails, Len("CHOICE_FORMULA")) = "CHOICE_FORMULA" Then
        Set choiceResolver = ChoiceFormula.Create(controlDetails)
        Set categoryList = choiceResolver.Categories()
        choiceName = choiceResolver.choiceName()
    Else
        choiceName = controlDetails
        Set categoryList = ChoiceCategories(controlDetails)
    End If

    If (Not this.tsTable Is Nothing) Then
        On Error Resume Next
            addTotalFlag = CStr(this.tsTable.Value("add total", seriesValue))
            If StrComp(addTotalFlag, "yes", vbTextCompare) = 0 Then
                categoryList.Push "Total"
            End If
        On Error GoTo 0
    End If

ExitFunction:
    Set GetCategories = categoryList
End Function

Private Function ChoiceCategories(ByVal listName As String) As BetterArray
    Dim choiceSheet As Worksheet
    Dim choiceObj As ILLChoices
    Dim emptyList As BetterArray

    Set emptyList = New BetterArray
    Set choiceSheet = SheetOrNothing(SHEET_CHOICES)
    If choiceSheet Is Nothing Then
        Set ChoiceCategories = emptyList
        Exit Function
    End If

    choiceObj = Nothing
    On Error Resume Next
        Set choiceObj = LLChoices.Create(choiceSheet, CHOICES_START_ROW, CHOICES_START_COLUMN)
    On Error GoTo 0
    If choiceObj Is Nothing Then
        Set ChoiceCategories = emptyList
    Else
        Set ChoiceCategories = choiceObj.Categories(listName)
    End If
End Function

Private Sub FormatLockCell(ByVal cellRng As Range, Optional ByVal locked As Boolean = True)
    If cellRng Is Nothing Then Exit Sub

    cellRng.Font.Color = IIf(locked, RGB(51, 142, 202), vbBlack)
    cellRng.Font.Italic = locked
    cellRng.Locked = locked
End Sub
