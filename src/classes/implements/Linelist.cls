VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Linelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Final linelist workbook"

'@Folder("Linelist Building")
'@IgnoreModule UseMeaningfulName, HungarianNotation
'@ModuleDescription("Final linelist workbook")
'@PredeclaredId

Option Explicit
Implements ILinelist

Private Enum CodeScope
    codeScopeClass = 1
    codeScopeModule = 2
End Enum

Private Type TLinelist
    outWkb As Workbook
    Data As ILinelistSpecs
    firstSheet As String                         'stock the name of the first sheet (depending on the os language)
End Type

Private this As TLinelist
Private Const CLASSNAME  As String = "Linelist"
Private Const PRINTPREFIX As String = "print_"
Private Const CRFPREFIX As String = "crf_"

Public Function Create(lData As ILinelistSpecs) As ILinelist

    If lData Is Nothing Then ThrowError "Linelist specifications unfound"
    With New Linelist
        Set .LinelistData = lData
        Set Create = .Self()
    End With
End Function

Public Property Get Self() As ILinelist
    Set Self = Me
End Property

Public Property Get LinelistData() As ILinelistSpecs
    Set LinelistData = this.Data
End Property

Public Property Set LinelistData(ByVal lData As ILinelistSpecs)
    Set this.Data = lData
End Property

'Dictionary
Private Property Get Dictionary() As ILLdictionary
    Dim lData As ILinelistSpecs
    Set lData = LinelistData()
    Set Dictionary = lData.Dictionary()
End Property

'Speed up before a work
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'Workbook
Private Property Get Wkb() As Workbook
    Dim mainobj As IMain
    Dim lData As ILinelistSpecs
    Dim outDir As String
    Dim sep As String

    Set lData = LinelistData()
    Set mainobj = lData.MainObject()
    sep = Application.PathSeparator
    outDir = mainobj.ValueOf("lldir") & sep & "LinelistApp_"

    If this.outWkb Is Nothing Then
       'Check if we could move to ribbon menu or use custom workbook
       If mainobj.HasTemplate() Then
            BusyApp
            Set this.outWkb = Workbooks.Open(mainobj.ValueOf("temppath"))
            this.outWkb.SaveAs fileName:=outDir & sep & "__temp.xlsb", fileFormat:=xlExcel12
        Else
            BusyApp
            Set this.outWkb = Workbooks.Add
        End If
        this.firstSheet = this.outWkb.Worksheets(1).Name
    End If
    Set Wkb = this.outWkb
End Property

'Test if a worksheet Exists
Private Function sheetExists(ByVal sheetName As String) As Boolean
    Dim sh As Worksheet
    Dim outWkb As Workbook

    Set outWkb = Wkb()
    On Error Resume Next
    Set sh = outWkb.Worksheets(sheetName)
    On Error GoTo 0
    sheetExists = (Not sh Is Nothing)
End Function

'Get the dropdown object for all the choices in a worksheet
'Scope values represent the differents dropdown worksheets:

'1- For choices_manual
'2- For choices_custom

Private Property Get Dropdown(Optional ByVal scope As Byte = 1) _
                              As IDropdownLists
    Dim sh As Worksheet
    Dim listSheetName As String
    Dim lData As ILinelistSpecs
    Dim drop As IDropdownLists
    Dim trans As ITranslation
    Dim headerPrefix As String

    Set lData = LinelistData()
    Set trans = lData.TransObject()
    
    Select Case scope
    Case 1

        listSheetName = lData.TemporarySheetName(TemporarySheetList)
        headerPrefix = "dropdown_"
    Case 2
        listSheetName = trans.TranslatedValue("LLSHEET_CustomChoice")
        headerPrefix = vbNullString
    End Select

    Set sh = Wksh(listSheetName)

    Set drop = DropdownLists.Create(sh, headerPrefix)
    Set Dropdown = drop
End Property

'Get the pivot tables Object
Private Property Get Pivots() As ICustomPivotTables
    Dim sh As Worksheet
    Dim sheetName As String
    Dim lData As ILinelistSpecs
    Dim trans As ITranslation
    Set lData = LinelistData()
    Set trans = lData.TransObject()
    sheetName = trans.TranslatedValue("LLSHEET_CustomPivotTable")
    Set sh = Wksh(sheetName)
    Set Pivots = CustomPivotTables.Create(sh)
End Property

'sheetScope can take 3 values: 1 for normal worksheet, 2 for printed worksheet,
'3 for crf worksheet

'Add worksheets to workbook
Private Sub AddOutputSheet(ByVal sheetName As String, _
                          Optional ByVal Hide As Long = xlSheetVeryHidden, _
                          Optional ByVal sheetScope As Byte = 1)

    'Maximum Length for sheetName in excel is 31. If
    'The sheet Length is > 31, I truncate the sheetName to firsts 30
    'Characters.
    Const SHEETNAMEMAXLENGTH As Integer = 30

    Dim outWkb As Workbook
    Dim actSheetName As String

    Set outWkb = Wkb()
    
    Select Case sheetScope
    Case 1
     actSheetName = sheetName
    Case 2
     actSheetName = PRINTPREFIX & sheetName
    Case 3
     actSheetName = CRFPREFIX & sheetName
    End Select
   
    'Truncate sheetnames to 30 characters
    actSheetName = Left(actSheetName, SHEETNAMEMAXLENGTH)
    If sheetExists(actSheetName) Then ThrowError "The sheet " & actSheetName & " already exists in the output workbook"
    With outWkb
        .Worksheets.Add.Name = actSheetName
        .Worksheets(actSheetName).Visible = Hide
    End With
End Sub

'Get a Worksheet in the output workbook
'Sheet scope can take 3 values:
' 1 for normal sheetName,
' 2 for printed worksheet
' 3 for crf worksheet
Private Property Get Wksh(ByVal sheetName As String, Optional ByVal sheetScope As Byte = 1) As Worksheet
    Dim sh As Worksheet
    Dim outWkb As Workbook
    Dim actSheetName As String 'actual sheet name either with the printed prefix, or not

    Select Case sheetScope
    Case 1
     actSheetName = sheetName
    Case 2
     actSheetName = PRINTPREFIX & sheetName
    Case 3
     actSheetName = CRFPREFIX & sheetName
    End Select

    If Not sheetExists(actSheetName) Then ThrowError "The sheet " & sheetName & "does not exists in current workbook"
    Set outWkb = Wkb()
    Set sh = outWkb.Worksheets(actSheetName)
    Set Wksh = sh
End Property


Private Sub TransferCode(moduleName As String, scope As CodeScope)
    Dim codeObject As Object                     'Where to stock to code in the workbook
    Dim componentObject As Object
    Dim outWkb As Workbook
    Dim inpWkb As Workbook
    Dim mainobj As IMain
    Dim sep As String
    Dim outDir As String
    Dim outPath As String

    Dim lData As ILinelistSpecs

    Set outWkb = Wkb()                           'Output workbook
    Set lData = LinelistData()
    Set inpWkb = lData.specswkb()
    Set mainobj = lData.MainObject()
    sep = Application.PathSeparator
    outDir = mainobj.ValueOf("lldir") & sep & "LinelistApp_"

    'get all the values within the actual module to transfer
    Select Case scope
    Case codeScopeModule
        outPath = outDir & sep & "CopieMod.bas"
    Case codeScopeClass
        outPath = outDir & sep & "CopieCls.cls"
    End Select

    'I need to import/export classes to keep their attribute. (self instanciation, etc.)
    On Error Resume Next
    Kill outPath
    On Error GoTo 0

    Set codeObject = inpWkb.VBProject.VBComponents(moduleName)
    codeObject.Export outPath

    Set componentObject = outWkb.VBProject.VBComponents
    componentObject.Import outPath

    On Error Resume Next
    Kill outPath
    On Error GoTo 0
End Sub

'This should be same module Name, in the workbooks
'Copy codes from a module in the designer, to a module in the template file
Private Sub CopyCodeModule(ByVal moduleName As String)
    Dim lData As ILinelistSpecs
    Dim mainobj As IMain
    Dim codeContent As String
    Dim codeMod As Object
    Dim vbProj As Object
    Dim impWkb As Workbook
    Dim outWkb As Workbook

    Set lData = LinelistData()
    Set mainobj = lData.MainObject()

    'Only on template files
    If Not mainobj.HasTemplate() Then Exit Sub

    Set impWkb = lData.specswkb()
    Set outWkb = Wkb()
    Set vbProj = impWkb.VBProject

    'Extract the code from the actual vbProject
    With vbProj
        With .VBComponents(moduleName).CodeModule
            codeContent = .Lines(1, .CountOfLines)
        End With
    End With

    'Export codeModule on export workbook
    Set vbProj = outWkb.VBProject
    Set codeMod = vbProj.VBComponents(moduleName).CodeModule
    'Adding to the export codeModule
    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub

'Transfert code to the workbook
Private Sub TransferCodeWkb()
    Const LINELISTWORKBOOKMODULENAME As String = "LinelistWorkbook"

    Dim lData As ILinelistSpecs
    Dim codeContent As String                    'a string to contain code to add
    Dim vbProj As Object                         'component, project and modules
    Dim vbComp As Object
    Dim codeMod As Object
    Dim inpWkb As Workbook
    Dim outWkb As Workbook

    Set lData = LinelistData()
    Set inpWkb = lData.specswkb()
    Set outWkb = Wkb()

    'save the code module in the string sNouvCode
    With inpWkb.VBProject.VBComponents(LINELISTWORKBOOKMODULENAME).CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With

    With outWkb
        Set vbProj = .VBProject
        Set vbComp = vbProj.VBComponents(.codeName)
        Set codeMod = vbComp.CodeModule
    End With

    'Adding the code
    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub

'Prepare temporary folder for the linelist creation
Public Sub TemporaryFolder(Optional Create As Boolean = True)
    Dim lData As ILinelistSpecs
    Dim mainobj As IMain
    Dim outPath As String
    Dim sep As String                            'Separator

    Set lData = LinelistData()
    Set mainobj = lData.MainObject()
    outPath = mainobj.ValueOf("lldir")
    sep = Application.PathSeparator

    'required temporary folder for analysis
    On Error Resume Next
    Kill outPath & sep & "LinelistApp_" & sep & "*.frm"
    Kill outPath & sep & "LinelistApp_" & sep & "*.frx"
    Kill outPath & sep & "LinelistApp_" & sep & "__temp.xlsb"
    RmDir outPath & sep & "LinelistApp_"
    If Create Then MkDir outPath & sep & "LinelistApp_" 'create a folder for sending all the data from designer
    On Error GoTo 0
End Sub

'Transfert a form to the output workbook
Private Sub TransferForm(formName As String)
    Dim outDir As String
    Dim lData As ILinelistSpecs
    Dim outWkb As Workbook
    Dim inpWkb As Workbook
    Dim mainobj As IMain
    Dim sep As String                            'Application path separator
    Dim formPath As String

    Set lData = LinelistData()
    Set outWkb = Wkb()
    Set inpWkb = lData.specswkb()
    Set mainobj = lData.MainObject()
    sep = Application.PathSeparator
    outDir = mainobj.ValueOf("lldir") & sep & "LinelistApp_"

    'Output dir is where to write the form for import / export
    On Error Resume Next
    Kill outDir & sep & "CopieUsf.frm"
    Kill outDir & sep & "CopieUsf.frx"
    On Error GoTo 0

    formPath = outDir & sep & "CopieUsf.frm"

    'Copy the frame to the output directory
    inpWkb.VBProject.VBComponents(formName).Export formPath
    outWkb.VBProject.VBComponents.Import formPath

    DoEvents

    On Error Resume Next
    Kill outDir & sep & "CopieUsf.frm"
    Kill outDir & sep & "CopieUsf.frx"
    On Error GoTo 0
End Sub

'On some temporary sheets, add some range names
Private Sub PrepareTemporarySheets()
    Dim sh As Worksheet
    Dim outWkb As Workbook
    Dim rng As Range
    Dim cellRng As Range
    Dim lData As ILinelistSpecs
    Dim trads As ITranslation
    Dim sheetName As String
    Dim counter As Long
    Dim namesArray As BetterArray
    Dim actName As String

    Set outWkb = Wkb()
    Set lData = LinelistData()
    Set namesArray = New BetterArray
    Set trads = lData.TransObject()

    'On Update sheets ----------------------------------------------------------
    
    sheetName = lData.TemporarySheetName(TemporarySheetUpdates)
    Set sh = outWkb.Worksheets(sheetName)

    'Range to check if:
    '- listAuto has been updated
    '- an import has been done
    '- the geobase has been updated manually by the user
    '.. You can add here some ranges for further analysis

    namesArray.Clear
    namesArray.Push "RNG_UpdateListAuto", "RNG_ImportDone", "RNG_GeoUpdateManual", _ 
                    "RNG_ShowAllOptionals", "RNG_ExportInit", "RNG_NumberExports", _
                    "RNG_EpiWeekStart"
    Set rng = sh.Cells(1, 1)
    
    For counter = namesArray.LowerBound To namesArray.UpperBound
        rng.Name = namesArray.Item(counter)
        Set rng = rng.Offset(2)
    Next

    sh.Range("RNG_NumberExports").Value = lData.NumberOfExports()
    sh.Range("RNG_EpiWeekStart").Value = -1

    'On Import report sheets ---------------------------------------------------

    'Create required tables for import report
    sheetName = lData.TemporarySheetName(TemporarySheetImport)
    Set sh = outWkb.Worksheets(sheetName)
    With sh
        .Cells(1, 1).Value = "sheetNotImported"
        .Cells(1, 3).Value = "sheetNotTouched"

        .Cells(1, 5).Value = "variableNotImported"
        .Cells(1, 6).Value = "correspondingSheet"

        .Cells(1, 8).Value = "variableNotFound"
        .Cells(1, 9).Value = "correspondingSheet"

        'Add the corresponding listObjects

        'sheetNotImported ListObject
        Set rng = .Range(.Cells(1, 1), .Cells(2, 1))
        
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_sheetsNotImp"

        'sheetsNot touched listObject
        Set rng = .Range(.Cells(1, 3), .Cells(2, 3))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_sheetsNotTouch"

        'Variable not imported listObject
        Set rng = .Range(.Cells(1, 5), .Cells(2, 6))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_varsNotImp"

        'variable not touched listObject
        Set rng = .Range(.Cells(1, 8), .Cells(2, 9))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_varsNotTouch"
    End With

    'Analysis Table IDs sheet --------------------------------------------------

    sheetName = lData.TemporarySheetName(TemporarySheetAnalysis)
    Set sh = outWkb.Worksheets(sheetName)
    Set cellRng = sh.Cells(1, 1)
    namesArray.Clear
    namesArray.Push "uba", "ts", _
                     "sp", "sptemp"
    For counter = namesArray.LowerBound To namesArray.UpperBound
        actName = namesArray.Item(counter)
        
        'Add headers for table IDs
        cellRng.Value = "tabId"
        cellRng.Cells(1, 2).Value = "tabRange"
        cellRng.Cells(1, 3).Value = "exportTable"

        'Add ListObject for table IDs
        Set rng = sh.Range(cellRng, cellRng.Cells(2, 3))
        sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "tab_ids_" & actName

        'Add ListObject for the graph Ids
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

        'Add ListObject for graphs label and Name
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

    'Add ranges for each of the analysis worksheets names (avoid using lltrads)
    'UA and bivariate sheet, TS sheet, spatial sheet and spatio-temporal sheet
    namesArray.Clear
    namesArray.Push "RNG_SheetUAName", "RNG_SheetTSName", _
                    "RNG_SheetSPName", "RNG_SheetSPTempName"
    Set cellRng = cellRng.Offset(, 2)
    For counter = namesArray.LowerBound To namesArray.UpperBound
        cellRng.Name = namesArray.Item(counter)
        Set cellRng = cellRng.Offset(1)
    Next

    'write the sheet names for analysis, time series, spatial and spatio-temporal
    sh.Range("RNG_SheetUAName").Value = trads.TranslatedValue("LLSHEET_Analysis")
    sh.Range("RNG_SheetTSName").Value = trads.TranslatedValue("LLSHEET_TemporalAnalysis")
    sh.Range("RNG_SheetSPName").Value = trads.TranslatedValue("LLSHEET_SpatialAnalysis")
    sh.Range("RNG_SheetSPTempName").Value = trads.TranslatedValue("LLSHEET_SpatioTemporalAnalysis")
End Sub

'Prepare the Linelist by adding required worksheets
Private Sub Prepare()

    Dim dict As ILLdictionary
    Dim trans As ITranslation
    Dim sheetsList As BetterArray
    Dim lData As ILinelistSpecs
    Dim counter As Long
    Dim sheetName As String
    Dim codeName As String
    Dim outWkb As Workbook
    Dim prevAlert As Boolean
    Dim codesList As BetterArray
    Dim mainobj As IMain

    'Create temporary folder for preparation steps
    TemporaryFolder Create:=True

    Set outWkb = Wkb()
    Set lData = LinelistData()
    Set sheetsList = New BetterArray
    Set trans = lData.TransObject()
    Set dict = lData.Dictionary()
    Set mainobj = lData.MainObject()

    'Select/unhide the instruction worksheet
    If mainObj.HasTemplate() And (mainObj.ValueOf("instruct") = "add") Then
        outWkb.Worksheets(trans.TranslatedValue("INSTSHEETNAME")).Visible = xlSheetVisible
    End If

    'temporary sheet Names
    sheetsList.Push lData.TemporarySheetName(TemporarySingle), _
        lData.TemporarySheetName(TemporarySheetList), _
        lData.TemporarySheetName(TemporarySheetUpdates), _
        lData.TemporarySheetName(TemporarySheetAnalysis), _
        lData.TemporarySheetName(TemporarySheetImport), _
        lData.TemporarySheetName(TemporarySpatial), _
        lData.TemporarySheetName(TemporaryShowHide)

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVeryHidden
    Next

    DoEvents

    'Prepare the temporary sheets by adding required ranges
    PrepareTemporarySheets

    'Export differents parts to the workbook (dictionary, choices, translation, geo)
    lData.Export outWkb

    DoEvents

    'Add the sheets for Bivariate Analysis, Time series and spatial,
    'Spatio-temporal Analysis As well as custom sheets (custom choices)

    sheetsList.Clear
    sheetsList.Push "LLSHEET_Analysis", "LLSHEET_TemporalAnalysis", _
                    "LLSHEET_SpatialAnalysis", "LLSHEET_SpatioTemporalAnalysis", _
                    "LLSHEET_CustomChoice", "LLSHEET_CustomPivotTable"

    'Reverse the order to keep the same sheet order in the linelist.
    sheetsList.Reverse

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = trans.TranslatedValue(sheetsList.Item(counter))
        AddOutputSheet sheetName, xlSheetVisible
    Next

    sheetsList.Clear
    
    'Unique sheet Names List(HList or VList)
    Set sheetsList = dict.UniqueValues("sheet name")
    
    'Important to reverse to keep the worksheets in the same order as
    'in the dictionary
    sheetsList.Reverse
    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVisible
    Next

    With outWkb
        .Worksheets(trans.TranslatedValue("INSTSHEETNAME")).Move _
         Before:=.Worksheets(sheetsList.Item(counter - 1))
    End With

    'Add admin sheet only if the user chooses buttons instead of ribbon linelist
    If Not mainobj.HasTemplate() Then
        'Admin sheet Name
        AddOutputSheet trans.TranslatedValue("LLSHEET_Admin"), xlSheetVisible
        'Prepare the LinelistAdmin sheet
        PrepareAdmin
    End If

    'Delete the first worksheet with name sheet1
    prevAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    outWkb.Worksheets(this.firstSheet).Delete
    Application.DisplayAlerts = prevAlert


    'Send the required classes to the linelist
    Set codesList = New BetterArray
    codesList.Push "BetterArray", "DataSheet", "IDataSheet", "LLExport", _
                    "ILLExport", "LLchoice", "ILLchoice", _
                    "LLdictionary", "ILLdictionary", "LLGeo", _
                    "ILLGeo", "LLPasswords", "ILLPasswords", _
                    "LLSheets", "ILLSheets", "LLTranslations", _
                     "ILLTranslations", "OSFiles", "IOSFiles", _
                    "Translation", "ITranslation", "LLVariables", _
                    "ILLVariables", "ILLSpatial", "LLSpatial", _
                    "DropdownLists", "IDropdownLists", "LLShowHide", _
                     "ILLShowHide", "CustomTable", "ICustomTable", _
                    "OutputSpecs", "IOutputSpecs", _
                    "ImpRep", "IImpRep", "ImpSpecs", "IImpSpecs", _
                    "UpVal", "IUpVal", "IGraphs", "Graphs", "AnaTabIds", _ 
                    "IAnaTabIds", "ExportButtons", "IExportButtons"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        TransferCode codeName, codeScopeClass
    Next

    DoEvents

    'Send the modules to the linelist
    codesList.Clear
    codesList.Push  "LinelistCustomFunctions", _
                    "LinelistGeo", "LinelistEvents", _
                    "EventsSpatialAnalysis", _
                    "EventsLinelistButtons", "EventsGlobalAnalysis"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        TransferCode codeName, codeScopeModule
    Next

    'Copy codes text from designer to ribbon (if the main object hastemplate)
    CopyCodeModule "EventsLinelistRibbon"

    'Transfer the forms
    codesList.Clear
    codesList.Push "F_Advanced", "F_Export", "F_ExportMig", _
                     "F_Geo", "F_ImportRep", _
                     "F_ShowHideLL", "F_ShowHidePrint", "F_ShowVarLabels", _
                     "F_EpiWeek"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        TransferForm codeName
    Next

    'Transfert workbook close codes to the workbook
    TransferCodeWkb
    TemporaryFolder Create:=False
End Sub

Private Sub PrepareAdmin()
    Const ADVANCEDCOMMAND As String = "ClickAdvanced"
    Const EXPORTCOMMAND As String = "ClickExport"
    Const SHPADVANCE As String = "SHP_Advanced"
    Const SHPEXPORT As String = "SHP_Export"

    Dim sh As Worksheet
    Dim trads As ITranslation
    Dim design As ILLFormat
    Dim pass As ILLPasswords
    Dim lData As LinelistSpecs
    Dim but As IButtons
    Dim rng As Range

    Set lData = LinelistData()
    Set design = lData.DesignFormat()
    Set trads = lData.TransObject()
    Set pass = lData.Password()

    Set sh = Wksh(trads.TranslatedValue("LLSHEET_Admin"))
    'Format the admin worksheet
    design.ApplyFormat sh, scope:=AllLinelistSheet

    'Change the translation to the Shapes one
    Set trads = lData.TransObject(TranslationOfShapes)

    'Advance button
    Set rng = sh.Cells(2, 3)
    Set but = Buttons.Create(outrng:=rng, codeName:=SHPADVANCE, scope:=ButtonScopeLarge)
    but.Add actionCommand:=ADVANCEDCOMMAND, shapeLabel:=trads.TranslatedValue(SHPADVANCE)
    but.Format design

    'Import button
    Set rng = sh.Cells(2, 5)
    Set but = Buttons.Create(outrng:=rng, codeName:=SHPEXPORT, scope:=ButtonScopeLarge)
    but.Add actionCommand:=EXPORTCOMMAND, shapeLabel:=trads.TranslatedValue(SHPEXPORT)
    but.Format design

    BusyApp
    sh.Activate
    BusyApp

    pass.Protect "_active"
End Sub

'Manage errors when building the linelist:
'- Asks if the user wants the uncomplete linelist.
'- If not, close the workbook and leave the application.

Private Sub ErrorManage(Optional ByVal textMessage As String = vbNullString)
    Dim outWkb As Workbook 'Output workbook
    Dim Quit As Integer

    'Ask the user if he wants to see (Or not) the uncompleted linelist
    Quit = MsgBox("An internal Error occured with specified informations: " & Chr(10) & _ 
                  textMessage & Chr(10) & _
                  "do you want to see the uncompleted linelist?", _
                  vbExclamation + vbYesNo, "ERROR")

    If Quit = vbYes Then
        Set outWkb = Wkb()
        'Return the uncomplete workbook
        outWkb.Application.Visible = True
    Else
        Set outWkb = Wkb()
        outWkb.Close savechanges:=False
    End If
End Sub

Private Sub SaveLL()
    Dim outWkb As Workbook
    Dim passwd As String
    Dim outPath As String
    Dim lData As ILinelistSpecs
    Dim mainobj As IMain
    Dim passobj As ILLPasswords
    Dim msgTrads As ITranslation

    Set lData = LinelistData()
    Set outWkb = Wkb()
    Set mainobj = lData.MainObject()
    Set passobj = lData.Password()
    Set msgTrads = lData.TransObject()

    passwd = mainobj.ValueOf("llpassword")
    outPath = mainobj.ValueOf("lldir") & Application.PathSeparator & mainobj.ValueOf("llname") & ".xlsb"

    'Select the admin sheet before saving
    On Error Resume Next
    outWkb.Worksheets(msgTrads.TranslatedValue("LLSHEET_Admin")).Activate
    outWkb.Worksheets(msgTrads.TranslatedValue("INSTSHEETNAME")).Activate
    On Error GoTo 0
    

    'Protect the workbook structure from eventual deletion/renaming of a worksheet
    passobj.ProtectWkb outWkb

    outWkb.SaveAs fileName:=outPath, fileFormat:=xlExcel12, Password:=passwd, _
                  ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    outWkb.Close
    TemporaryFolder Create:=False
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Properties
Private Property Get ILinelist_Wksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_Wksh = Wksh(sheetName)
End Property

Private Property Get ILinelist_PrintWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_PrintWksh = Wksh(sheetName, sheetScope:=2)
End Property

Private Property Get ILinelist_CRFWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_CRFWksh = Wksh(sheetName, sheetScope:=3)
End Property

Private Property Get ILinelist_LinelistData() As ILinelistSpecs
    Set ILinelist_LinelistData = LinelistData()
End Property

Private Property Get ILinelist_Dictionary() As ILLdictionary
    Set ILinelist_Dictionary = Dictionary()
End Property

Private Property Get ILinelist_Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
    Set ILinelist_Dropdown = Dropdown(scope)
End Property

'PivotTables Object
Private Property Get ILinelist_Pivots() As ICustomPivotTables
    Set ILinelist_Pivots = Pivots()
End Property

'Exposed methods
Private Sub ILinelist_Prepare()
    Prepare
End Sub

Private Function ILinelist_SheetExists(ByVal sheetName As String) As Boolean
    ILinelist_SheetExists = sheetExists(sheetName)
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
