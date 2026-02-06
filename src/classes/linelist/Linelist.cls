VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Linelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist workbook facade: creates, populates and saves the output workbook"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Linelist workbook facade: creates, populates and saves the output workbook")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

Implements ILinelist

Private Const CLASS_NAME As String = "Linelist"
Private Const PRINT_PREFIX As String = "print_"
Private Const CRF_PREFIX As String = "crf_"
Private Const SHEET_NAME_MAX_LENGTH As Long = 30
Private Const WORKBOOK_MODULE_NAME As String = "LinelistWorkbook"

Private Type TLinelist
    outWkb As Workbook
    specs As ILinelistSpecs
    tempRepos As ITemporaryRepos
    firstSheet As String
End Type

Private this As TLinelist

'@section Factory
'===============================================================================

Public Function Create(ByVal specs As ILinelistSpecs) As ILinelist
    If specs Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist specifications are required"
    End If

    Dim instance As Linelist
    Set instance = New Linelist

    Set instance.BinSpecs = specs
    Set Create = instance
End Function

'@section Bindings
'===============================================================================

Friend Property Set BinSpecs(ByVal value As ILinelistSpecs)
    Set this.specs = value
End Property

'@section Public properties
'===============================================================================

Public Property Get LinelistData() As ILinelistSpecs
    Set LinelistData = this.specs
End Property

Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.specs.Dictionary
End Property

Public Property Get Wksh(ByVal sheetName As String) As Worksheet
    ResolveWksh sheetName, 1, Wksh
End Property

Public Property Get PrintWksh(ByVal sheetName As String) As Worksheet
    ResolveWksh sheetName, 2, PrintWksh
End Property

Public Property Get CRFWksh(ByVal sheetName As String) As Worksheet
    ResolveWksh sheetName, 3, CRFWksh
End Property

Public Property Get Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
    Dim sh As Worksheet
    Dim listSheetName As String
    Dim trans As ITranslationObject
    Dim headerPrefix As String

    Set trans = this.specs.TransObject

    Select Case scope
        Case 1
            listSheetName = this.specs.TemporarySheetName(TEMP_SHEET_LIST)
            headerPrefix = "dropdown_"
        Case 2
            listSheetName = trans.TranslatedValue("LLSHEET_CustomChoice")
            headerPrefix = vbNullString
    End Select

    Set sh = Wksh(listSheetName)
    Set Dropdown = DropdownLists.Create(sh, headerPrefix)
End Property

Public Property Get Pivots() As ICustomPivotTables
    Dim sh As Worksheet
    Dim trans As ITranslationObject

    Set trans = this.specs.TransObject
    Set sh = Wksh(trans.TranslatedValue("LLSHEET_CustomPivotTable"))
    Set Pivots = CustomPivotTables.Create(sh)
End Property

Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim sh As Worksheet

    On Error Resume Next
    Set sh = OutputWkb.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = (Not sh Is Nothing)
End Function

Public Sub AddOutputSheet(ByVal sheetName As String, _
                          Optional ByVal Hide As Long = xlSheetVeryHidden, _
                          Optional ByVal sheetScope As Byte = 1)
    Dim actSheetName As String
    Dim outWkb As Workbook

    Select Case sheetScope
        Case 1
            actSheetName = sheetName
        Case 2
            actSheetName = PRINT_PREFIX & sheetName
        Case 3
            actSheetName = CRF_PREFIX & sheetName
    End Select

    actSheetName = Left$(actSheetName, SHEET_NAME_MAX_LENGTH)

    If SheetExists(actSheetName) Then
        ThrowError ProjectError.InvalidArgument, _
                   "The sheet " & actSheetName & " already exists in the output workbook"
    End If

    Set outWkb = OutputWkb

    With outWkb
        .Worksheets.Add.Name = actSheetName
        .Worksheets(actSheetName).Visible = Hide
    End With
End Sub

'@section Lifecycle
'===============================================================================

Public Sub Prepare()
    Dim dict As ILLdictionary
    Dim trans As ITranslationObject
    Dim sheetsList As BetterArray
    Dim counter As Long
    Dim sheetName As String
    Dim outWkb As Workbook
    Dim prevAlert As Boolean
    Dim mainObj As IMain
    Dim codeTransfer As ICodeTransfer

    Set mainObj = this.specs.MainObject
    Set this.tempRepos = TemporaryRepos.Create(mainObj.ValueOf("lldir"))
    this.tempRepos.EnsureReady

    Set outWkb = OutputWkb
    Set trans = this.specs.TransObject
    Set dict = this.specs.Dictionary
    Set sheetsList = New BetterArray

    'Handle instruction worksheet for template workbooks
    If mainObj.HasTemplate() And (mainObj.ValueOf("instruct") = "add") Then
        outWkb.Worksheets(trans.TranslatedValue("INSTSHEETNAME")).Visible = xlSheetVisible
    End If

    'Add temporary sheets (6 sheets — TemporarySheetUpdates removed, flags are now workbook-level hidden names)
    sheetsList.Push this.specs.TemporarySheetName(TEMP_SHEET_SINGLE), _
                    this.specs.TemporarySheetName(TEMP_SHEET_LIST), _
                    this.specs.TemporarySheetName(TEMP_SHEET_ANALYSIS), _
                    this.specs.TemporarySheetName(TEMP_SHEET_IMPORT), _
                    this.specs.TemporarySheetName(TEMP_SHEET_SPATIAL), _
                    this.specs.TemporarySheetName(TEMP_SHEET_SHOWHIDE)

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVeryHidden
    Next

    'Prepare temporary sheets with hidden names and ListObjects
    PrepareTemporarySheets

    'Export specifications to the workbook
    this.specs.Export outWkb

    'Add analysis/custom sheets (reversed order to maintain dictionary order)
    sheetsList.Clear
    sheetsList.Push "LLSHEET_Analysis", "LLSHEET_TemporalAnalysis", _
                    "LLSHEET_SpatialAnalysis", "LLSHEET_SpatioTemporalAnalysis", _
                    "LLSHEET_CustomChoice", "LLSHEET_CustomPivotTable"
    sheetsList.Reverse

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = trans.TranslatedValue(sheetsList.Item(counter))
        AddOutputSheet sheetName, xlSheetVisible
    Next

    'Add data entry sheets from dictionary
    sheetsList.Clear
    Set sheetsList = dict.UniqueValues("sheet name")
    sheetsList.Reverse

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVisible
    Next

    'Move instruction sheet before the last data entry sheet
    With outWkb
        .Worksheets(trans.TranslatedValue("INSTSHEETNAME")).Move _
            Before:=.Worksheets(sheetsList.Item(counter - 1))
    End With

    'Add admin sheet if not using template
    If Not mainObj.HasTemplate() Then
        AddOutputSheet trans.TranslatedValue("LLSHEET_Admin"), xlSheetVisible
        PrepareAdmin
    End If

    'Delete the initial default sheet
    prevAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    outWkb.Worksheets(this.firstSheet).Delete
    Application.DisplayAlerts = prevAlert

    'Transfer code via CodeTransfer
    Set codeTransfer = CodeTransfer.Create(this.specs.Workbook, outWkb, this.tempRepos)
    TransferAllCode codeTransfer, mainObj
End Sub

Public Sub SaveLL()
    Dim outWkb As Workbook
    Dim passwd As String
    Dim outPath As String
    Dim mainObj As IMain
    Dim passObj As IPasswords
    Dim msgTrads As ITranslationObject

    Set outWkb = OutputWkb
    Set mainObj = this.specs.MainObject
    Set passObj = this.specs.Password
    Set msgTrads = this.specs.TransObject

    passwd = mainObj.ValueOf("llpassword")
    outPath = mainObj.ValueOf("lldir") & Application.PathSeparator & _
              mainObj.ValueOf("llname") & ".xlsb"

    'Select an appropriate visible sheet before saving
    On Error Resume Next
    outWkb.Worksheets(msgTrads.TranslatedValue("LLSHEET_Admin")).Activate
    outWkb.Worksheets(msgTrads.TranslatedValue("INSTSHEETNAME")).Activate
    On Error GoTo 0

    'Protect the workbook structure
    passObj.Protect outWkb

    outWkb.SaveAs fileName:=outPath, fileFormat:=xlExcel12, Password:=passwd, _
                   ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    outWkb.Close

    If Not this.tempRepos Is Nothing Then this.tempRepos.Reset
End Sub

Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString)
    Dim outWkb As Workbook
    Dim answer As Long

    answer = MsgBox("An internal error occurred with specified information: " & Chr(10) & _
                    textMessage & Chr(10) & _
                    "Do you want to see the uncompleted linelist?", _
                    vbExclamation + vbYesNo, "ERROR")

    Set outWkb = OutputWkb

    If answer = vbYes Then
        outWkb.Application.Visible = True
    Else
        outWkb.Close savechanges:=False
    End If
End Sub

'@section Private — Workbook management
'===============================================================================

'Lazy-create or return the output workbook.
Private Property Get OutputWkb() As Workbook
    Dim mainObj As IMain
    Dim outDir As String
    Dim sep As String

    If this.outWkb Is Nothing Then
        Set mainObj = this.specs.MainObject
        sep = Application.PathSeparator
        outDir = mainObj.ValueOf("lldir") & sep & "OBTApp_"

        BusyApp

        If mainObj.HasTemplate() Then
            Set this.outWkb = Workbooks.Open(mainObj.ValueOf("temppath"))
            this.outWkb.SaveAs fileName:=outDir & sep & "__temp.xlsb", fileFormat:=xlExcel12
        Else
            Set this.outWkb = Workbooks.Add
        End If

        this.firstSheet = this.outWkb.Worksheets(1).Name
    End If

    Set OutputWkb = this.outWkb
End Property

'Suppress UI updates for batch operations.
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'Resolve a worksheet from the output workbook with an optional scope prefix.
Private Sub ResolveWksh(ByVal sheetName As String, ByVal sheetScope As Byte, _
                        ByRef result As Worksheet)
    Dim actSheetName As String

    Select Case sheetScope
        Case 1
            actSheetName = sheetName
        Case 2
            actSheetName = PRINT_PREFIX & sheetName
        Case 3
            actSheetName = CRF_PREFIX & sheetName
    End Select

    If Not SheetExists(actSheetName) Then
        ThrowError ProjectError.InvalidArgument, _
                   "The sheet " & actSheetName & " does not exist in the output workbook"
    End If

    Set result = OutputWkb.Worksheets(actSheetName)
End Sub

'@section Private — Prepare helpers
'===============================================================================

'Store workbook-level flags via HiddenNames and set up ListObjects on temporary sheets.
Private Sub PrepareTemporarySheets()
    Dim outWkb As Workbook
    Dim sh As Worksheet
    Dim rng As Range
    Dim cellRng As Range
    Dim trads As ITranslationObject
    Dim sheetName As String
    Dim counter As Long
    Dim namesArray As BetterArray
    Dim actName As String
    Dim wkbNames As IHiddenNames
    Dim shNames As IHiddenNames

    Set outWkb = OutputWkb
    Set trads = this.specs.TransObject
    Set namesArray = New BetterArray

    'Workbook-level hidden names for global flags (replaces TemporarySheetUpdates)
    Set wkbNames = HiddenNames.Create(outWkb)
    wkbNames.EnsureName "RNG_UpdateListAuto", vbNullString, HiddenNameTypeString
    wkbNames.EnsureName "RNG_ImportDone", vbNullString, HiddenNameTypeString
    wkbNames.EnsureName "RNG_GeoUpdateManual", vbNullString, HiddenNameTypeString
    wkbNames.EnsureName "RNG_ShowAllOptionals", vbNullString, HiddenNameTypeString
    wkbNames.EnsureName "RNG_ExportInit", vbNullString, HiddenNameTypeString
    wkbNames.EnsureName "RNG_NumberExports", this.specs.Value("numberofexports"), HiddenNameTypeString
    wkbNames.EnsureName "RNG_EpiWeekStart", this.specs.Value("epiweekstart"), HiddenNameTypeString

    'Import report sheet — create ListObjects
    sheetName = this.specs.TemporarySheetName(TEMP_SHEET_IMPORT)
    Set sh = outWkb.Worksheets(sheetName)

    With sh
        .Cells(1, 1).Value = "sheetNotImported"
        .Cells(1, 3).Value = "sheetNotTouched"
        .Cells(1, 5).Value = "variableNotImported"
        .Cells(1, 6).Value = "correspondingSheet"
        .Cells(1, 8).Value = "variableNotFound"
        .Cells(1, 9).Value = "correspondingSheet"

        Set rng = .Range(.Cells(1, 1), .Cells(2, 1))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_sheetsNotImp"

        Set rng = .Range(.Cells(1, 3), .Cells(2, 3))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_sheetsNotTouch"

        Set rng = .Range(.Cells(1, 5), .Cells(2, 6))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_varsNotImp"

        Set rng = .Range(.Cells(1, 8), .Cells(2, 9))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_varsNotTouch"
    End With

    'Analysis table IDs sheet — create ListObjects + worksheet-level hidden names
    sheetName = this.specs.TemporarySheetName(TEMP_SHEET_ANALYSIS)
    Set sh = outWkb.Worksheets(sheetName)
    Set cellRng = sh.Cells(1, 1)

    namesArray.Clear
    namesArray.Push "uba", "ts", "sp", "sptemp"

    For counter = namesArray.LowerBound To namesArray.UpperBound
        actName = namesArray.Item(counter)

        'Table IDs ListObject
        cellRng.Value = "tabId"
        cellRng.Cells(1, 2).Value = "tabRange"
        cellRng.Cells(1, 3).Value = "exportTable"

        Set rng = sh.Range(cellRng, cellRng.Cells(2, 3))
        sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "tab_ids_" & actName

        'Graph IDs ListObject
        Set cellRng = cellRng.Cells(1, 5)
        cellRng.Value = "tabId"
        cellRng.Cells(1, 2).Value = "graphId"
        cellRng.Cells(1, 3).Value = "SeriesName"
        cellRng.Cells(1, 4).Value = "SeriesType"
        cellRng.Cells(1, 5).Value = "SeriesPos"
        cellRng.Cells(1, 6).Value = "SeriesLabel"
        cellRng.Cells(1, 7).Value = "SeriesColumnLabel"
        cellRng.Cells(1, 8).Value = "hardCodeLabels"
        cellRng.Cells(1, 9).Value = "outRangeAddress"
        cellRng.Cells(1, 10).Value = "prefix"
        cellRng.Cells(1, 11).Value = "prefixOnly"

        Set rng = sh.Range(cellRng, cellRng.Cells(2, 11))
        sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "graph_ids_" & actName

        'Graph formats ListObject
        Set cellRng = cellRng.Cells(1, 13)
        cellRng.Value = "tabId"
        cellRng.Cells(1, 2).Value = "graphId"
        cellRng.Cells(1, 3).Value = "catTitle"
        cellRng.Cells(1, 4).Value = "valuesTitle"
        cellRng.Cells(1, 5).Value = "hardCodeLabels"
        cellRng.Cells(1, 6).Value = "heightFactor"
        cellRng.Cells(1, 7).Value = "plotTitle"

        Set rng = sh.Range(cellRng, cellRng.Cells(2, 7))
        sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "graph_formats_" & actName

        Set cellRng = cellRng.Cells(1, 10)
    Next

    'Worksheet-level hidden names for analysis sheet name references
    Set shNames = HiddenNames.Create(sh)
    shNames.EnsureName "RNG_SheetUAName", _
                       trads.TranslatedValue("LLSHEET_Analysis"), HiddenNameTypeString
    shNames.EnsureName "RNG_SheetTSName", _
                       trads.TranslatedValue("LLSHEET_TemporalAnalysis"), HiddenNameTypeString
    shNames.EnsureName "RNG_SheetSPName", _
                       trads.TranslatedValue("LLSHEET_SpatialAnalysis"), HiddenNameTypeString
    shNames.EnsureName "RNG_SheetSPTempName", _
                       trads.TranslatedValue("LLSHEET_SpatioTemporalAnalysis"), HiddenNameTypeString
End Sub

'Format the admin sheet and add Advanced/Export buttons.
Private Sub PrepareAdmin()
    Const ADVANCED_COMMAND As String = "ClickAdvanced"
    Const EXPORT_COMMAND As String = "ClickExport"
    Const SHP_ADVANCED As String = "SHP_Advanced"
    Const SHP_EXPORT As String = "SHP_Export"

    Dim sh As Worksheet
    Dim trads As ITranslationObject
    Dim design As ILLFormat
    Dim pass As IPasswords
    Dim but As IButtons
    Dim rng As Range

    Set design = this.specs.DesignFormat
    Set trads = this.specs.TransObject
    Set pass = this.specs.Password

    Set sh = Wksh(trads.TranslatedValue("LLSHEET_Admin"))
    design.ApplyFormat sh, scope:=AllLinelistSheet

    'Switch to shapes translation scope
    Set trads = this.specs.TransObject(TranslationOfShapes)

    'Advanced button
    Set rng = sh.Cells(2, 3)
    Set but = Buttons.Create(outrng:=rng, codeName:=SHP_ADVANCED, scope:=ButtonScopeLarge)
    but.Add actionCommand:=ADVANCED_COMMAND, shapeLabel:=trads.TranslatedValue(SHP_ADVANCED)
    but.Format design

    'Export button
    Set rng = sh.Cells(2, 5)
    Set but = Buttons.Create(outrng:=rng, codeName:=SHP_EXPORT, scope:=ButtonScopeLarge)
    but.Add actionCommand:=EXPORT_COMMAND, shapeLabel:=trads.TranslatedValue(SHP_EXPORT)
    but.Format design

    BusyApp
    sh.Activate
    BusyApp

    pass.Protect "_active"
End Sub

'Transfer all VBA code components to the output workbook.
Private Sub TransferAllCode(ByVal codeTransfer As ICodeTransfer, _
                            ByVal mainObj As IMain)
    Dim codesList As BetterArray
    Dim counter As Long
    Dim codeName As String

    Set codesList = New BetterArray

    'Transfer classes (42 classes)
    codesList.Push "BetterArray", "DataSheet", "IDataSheet", "LLExport", _
                   "ILLExport", "LLChoices", "ILLChoices", _
                   "LLdictionary", "ILLdictionary", "LLGeo", _
                   "ILLGeo", "Passwords", "IPasswords", _
                   "LLSheets", "ILLSheets", "LLTranslations", _
                   "ILLTranslations", "OSFiles", "IOSFiles", _
                   "TranslationObject", "ITranslationObject", "LLVariables", _
                   "ILLVariables", "ILLSpatial", "LLSpatial", _
                   "DropdownLists", "IDropdownLists", "LLShowHide", _
                   "ILLShowHide", "CustomTable", "ICustomTable", _
                   "OutputSpecs", "IOutputSpecs", _
                   "ImpRep", "IImpRep", "ImpSpecs", "IImpSpecs", _
                   "UpVal", "IUpVal", "IGraphs", "Graphs", "AnaTabIds", _
                   "IAnaTabIds", "ExportButtons", "IExportButtons"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferClass codeName
    Next

    'Transfer modules (6 modules)
    codesList.Clear
    codesList.Push "LinelistCustomFunctions", _
                   "LinelistGeo", "LinelistEvents", _
                   "EventsSpatialAnalysis", _
                   "EventsLinelistButtons", "EventsGlobalAnalysis", _
                   "LLChange"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferModule codeName
    Next

    'Copy ribbon module text (for template workbooks)
    codeTransfer.CopyModuleText "EventsLinelistRibbon"

    'Transfer forms (9 forms)
    codesList.Clear
    codesList.Push "F_Advanced", "F_Export", "F_ExportMig", _
                   "F_Geo", "F_ImportRep", _
                   "F_ShowHideLL", "F_ShowHidePrint", "F_ShowVarLabels", _
                   "F_EpiWeek"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferForm codeName
    Next

    'Transfer workbook code
    codeTransfer.TransferWorkbookCode WORKBOOK_MODULE_NAME

    'Clean up temp files
    If Not this.tempRepos Is Nothing Then this.tempRepos.Reset
End Sub

'@section Private — Error handling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, CLASS_NAME, errorMessage
End Sub

'@section ILinelist implementation
'===============================================================================

Private Property Get ILinelist_Wksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_Wksh = Wksh(sheetName)
End Property

Private Property Get ILinelist_PrintWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_PrintWksh = PrintWksh(sheetName)
End Property

Private Property Get ILinelist_CRFWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_CRFWksh = CRFWksh(sheetName)
End Property

Private Property Get ILinelist_Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
    Set ILinelist_Dropdown = Dropdown(scope)
End Property

Private Property Get ILinelist_Pivots() As ICustomPivotTables
    Set ILinelist_Pivots = Pivots
End Property

Private Property Get ILinelist_LinelistData() As ILinelistSpecs
    Set ILinelist_LinelistData = LinelistData
End Property

Private Property Get ILinelist_Dictionary() As ILLdictionary
    Set ILinelist_Dictionary = Dictionary
End Property

Private Sub ILinelist_Prepare()
    Prepare
End Sub

Private Function ILinelist_SheetExists(ByVal sheetName As String) As Boolean
    ILinelist_SheetExists = SheetExists(sheetName)
End Function

Private Sub ILinelist_AddOutputSheet(ByVal sheetName As String, _
                                     Optional ByVal Hide As Long = xlSheetVeryHidden, _
                                     Optional ByVal sheetScope As Byte = 1)
    AddOutputSheet sheetName, Hide, sheetScope
End Sub

Private Sub ILinelist_SaveLL()
    SaveLL
End Sub

Private Sub ILinelist_ErrorManage(Optional ByVal textMessage As String = vbNullString)
    ErrorManage textMessage
End Sub
