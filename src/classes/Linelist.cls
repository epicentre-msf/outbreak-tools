VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Linelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("Linelist")
'@ModuleDescription("Final linelist workbook")
'@IgnoreModule

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
Public Property Get Dictionary() As ILLdictionary
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
Public Property Get Wkb() As Workbook
    Dim xlsapp As Object

    If this.outWkb Is Nothing Then
        BusyApp
        Set this.outWkb = Workbooks.Add()
        BusyApp
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
Private Property Get Dropdown() As IDropdownLists
    Dim sh As Worksheet
    Dim listSheetName As String
    Dim lData As ILinelistSpecs
    Dim drop As IDropdownLists

    Set lData = LinelistData()
    listSheetName = lData.TemporarySheetName(TemporarySheetList)
    Set sh = Wksh(listSheetName)

    Set drop = DropdownLists.Create(sh)
    Set Dropdown = drop
End Property

'Add worksheets to workbook
Private Sub AddOutputSheet(ByVal sheetName As String, _
                          Optional ByVal Hide As Long = xlSheetVeryHidden, _
                          Optional ByVal usePrinted As Boolean = False)
    Dim outWkb As Workbook
    Dim actSheetName As String

    Set outWkb = Wkb()
    actSheetName = IIF(usePrinted, PRINTPREFIX & sheetName, sheetName)

    If sheetExists(actSheetName) Then ThrowError "The sheet " & actSheetName & " already exists in the output workbook"
    With outWkb
        .Worksheets.Add.Name = actSheetName
        .Worksheets(actSheetName).Visible = Hide
    End With
End Sub

'Get a Worksheet in the output workbook
Private Property Get Wksh(ByVal sheetName As String, Optional ByVal usePrinted As Boolean = False) As Worksheet
    Dim sh As Worksheet
    Dim outWkb As Workbook
    Dim actSheetName As String 'actual sheet name either with the printed prefix, or not

    'Test if the sheet Exists
    actSheetName = IIF(usePrinted, PRINTPREFIX & sheetName, sheetName)
    If Not sheetExists(actSheetName) Then ThrowError "The sheet " & sheetName & "does not exists in current workbook"
    Set outWkb = Wkb()
    Set sh = outWkb.Worksheets(actSheetName)
    Set Wksh = sh
End Property


Private Sub TransferCode(moduleName As String, scope As CodeScope)
    Dim codeObject As Object                     'Where to stock to code in the workbook
    Dim codeContent As String
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
    codeObject.export outPath
    Set codeObject = outWkb.VBProject.VBComponents
    codeObject.Import outPath

    On Error Resume Next
    Kill outPath
    On Error GoTo 0
End Sub

'Transfert code to the workbook
Private Sub TransferCodeWkb()
    Const LINELISTWORKBOOKMODULENAME = "LinelistWorkbook"

    Dim lData As ILinelistSpecs
    Dim codeContent As String                    'a string to contain code to add
    Dim vbProj As Object                         'component, project and modules
    Dim vbComp As Object
    Dim codeMod As Object
    Dim sheetName As String
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
    RmDir outPath & sep & "LinelistApp_"
    If Create Then MkDir outPath & sep & "LinelistApp_" 'create a folder for sending all the data from designer
    On Error GoTo 0
End Sub

'Transfert a form to the output workbook
Public Sub TransferForm(formName As String)
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
    inpWkb.VBProject.VBComponents(formName).export formPath
    outWkb.VBProject.VBComponents.Import formPath

    DoEvents

    On Error Resume Next
    Kill outDir & sep & "CopieUsf.frm"
    Kill outDir & sep & "CopieUsf.frx"
    On Error GoTo 0
End Sub

'Prepare the Linelist by adding required worksheets
Private Sub Prepare()
    Const FILTERPREFIX As String = "f"

    Dim dict As ILLdictionary
    Dim trans As ITranslation
    Dim sheetsList As BetterArray
    Dim lData As ILinelistSpecs
    Dim counter As Long
    Dim sheetName As String
    Dim codeName As String
    Dim outWkb As Workbook
    Dim prevAlert As Boolean
    Dim classesList As BetterArray

    Set outWkb = Wkb()

    Set lData = LinelistData()
    Set sheetsList = New BetterArray
    Set trans = lData.TransObject()
    Set dict = lData.Dictionary()

    'temporary sheet Names
    sheetsList.Push lData.TemporarySheetName(TemporarySingle), lData.TemporarySheetName(TemporarySheetList), _
        lData.TemporarySheetName(TemporarySheetDebug), _
        lData.TemporarySheetName(TemporarySheetAnalysis), lData.TemporarySheetName(TemporarySheetImport), _
        lData.TemporarySheetName(TemporarySpatial)

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVeryHidden
    Next

    DoEvents

    'Export differents parts to the workbook (dictionary, choices, translation, geo)
    lData.export outWkb

    DoEvents

    'Unique sheet Names (HList or VList)
    'Add the sheets for Bivariate Analysis, Time series and spatial Analysis
    AddOutputSheet trans.TranslatedValue("LLSHEET_SpatialAnalysis"), xlsheetvisible
    AddOutputSheet trans.TranslatedValue("LLSHEET_TemporalAnalysis"), xlsheetvisible
    AddOutputSheet trans.TranslatedValue("LLSHEET_Analysis"), xlsheetvisible

    sheetsList.Clear
    Set sheetsList = dict.UniqueValues("sheet name")
    sheetsList.Reverse

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlsheetvisible
    Next

    'Admin sheet Name
    AddOutputSheet trans.TranslatedValue("LLSHEET_Admin"), xlsheetvisible

    'Delete the first worksheet with name sheet1
    prevAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    outWkb.Worksheets(this.firstSheet).Delete
    Application.DisplayAlerts = prevAlert

    'Send the User forms and classes to the linelist
    TemporaryFolder Create:=True

    'Send the required classes to the linelist
    Set classesList = New BetterArray
    classesList.Push "BetterArray", "DataSheet", "IDataSheet", "LLExport", "ILLExport", "LLchoice", "ILLchoice", _
                     "LLdictionary", "ILLdictionary", "LLGeo", "ILLGeo", "LLPasswords", "ILLPasswords", _
                     "LLSheets", "ILLSheets", "LLTranslations", "ILLTranslations", "OSFiles", "IOSFiles", _
                     "Translation", "ITranslation", "LLVariables", "ILLVariables", "ILLSpatial", "LLSpatial", _
                     "DropdownLists", "IDropdownLists"

    For counter = classesList.LowerBound To classesList.UpperBound
        codeName = classesList.Item(counter)
        TransferCode codeName, codeScopeClass
    Next
    DoEvents

    'Send the modules to the linelist
    classesList.Clear
    classesList.Push "Helpers", "LinelistConstants", "LinelistCustomFunctions", _
                     "LinelistExport", "LinelistGeo", "LinelistMigration", "LinelistShowHide", _
                     "LinelistTranslation", "LinelistEvents", "LinelistDictionary"

    For counter = classesList.LowerBound To classesList.UpperBound
        codeName = classesList.Item(counter)
        TransferCode codeName, codeScopeModule
    Next

    classesList.Clear
    classesList.Push "F_Advanced", "F_Export", "F_ExportMig", "F_Geo", "F_ImportMig", "F_ImportRep", _
                     "F_NomVisible"

    For counter = classesList.LowerBound To classesList.UpperBound
        codeName = classesList.Item(counter)
        TransferForm codeName
    Next

    'Transfert workbook close codes to the workbook
    TransferCodeWkb
    TemporaryFolder Create:=False
    'Prepare the LinelistAdmin sheet
    PrepareAdmin
End Sub

Private Sub PrepareAdmin()
    Const ADVANCEDCOMMAND As String = "ClicCmdAdvanced"
    Const EXPORTCOMMAND As String = "ClicCmdExport"
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
    Set but = Buttons.Create(outRng:=rng, codeName:=SHPADVANCE, scope:=ButtonScopeLarge)
    but.Add actionCommand:=ADVANCEDCOMMAND, shapeLabel:=trads.TranslatedValue(SHPADVANCE)
    but.Format design

    'Import button
    Set rng = sh.Cells(2, 5)
    Set but = Buttons.Create(outRng:=rng, codeName:=SHPEXPORT, scope:=ButtonScopeLarge)
    but.Add actionCommand:=EXPORTCOMMAND, shapeLabel:=trads.TranslatedValue(SHPEXPORT)
    but.Format design

    BusyApp
    sh.Activate
    BusyApp

    pass.Protect "_active"
End Sub

'Ask question on Error to the user:
Private Sub ErrorManage()
    Dim outWkb As Workbook 'Output workbook
    Dim Quit As Integer

    'Ask the user if he wants to see (Or not) the uncompleted linelist
    Quit = MsgBox("An internal Error occured, do you want to see the uncompleted linelist?", _
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
    outWkb.Worksheets(msgTrads.TranslatedValue("LLSHEET_Admin")).Activate

    'Protect the workbook structure from eventual deletion/renaming of a worksheet
    passobj.ProtectWkb outWkb

    outWkb.SaveAs FileName:=outPath, fileformat:=xlExcel12, Password:=passwd, _
                  ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    outWkb.Close
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Properties
Public Property Get ILinelist_Wksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_Wksh = Wksh(sheetName)
End Property

Public Property Get ILinelist_PrintWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_PrintWksh = Wksh(sheetName, usePrinted:=True)
End Property

Public Property Get ILinelist_LinelistData() As ILinelistSpecs
    Set ILinelist_LinelistData = LinelistData()
End Property

Public Property Get ILinelist_Dictionary() As ILLdictionary
    Set ILinelist_Dictionary = Dictionary()
End Property

Public Property Get ILinelist_Dropdown() As IDropdownLists
    Set ILinelist_Dropdown = Dropdown()
End Property


'Exposed methods
Public Sub ILinelist_Prepare()
    Prepare
End Sub

Public Function ILinelist_SheetExists(ByVal sheetName As String) As Boolean
    ILinelist_SheetExists = sheetExists(sheetName)
End Function

Public Sub ILinelist_AddOutputSheet(ByVal sheetName As String, _
                                    Optional ByVal Hide As Long = xlSheetVeryHidden, _
                                    Optional ByVal usePrinted As Boolean = False)
    AddOutputSheet sheetName, Hide, usePrinted
End Sub

Public Sub ILinelist_SaveLL()
    SaveLL
End Sub

Public Sub ILinelist_ErrorManage()
    ErrorManage
End Sub
