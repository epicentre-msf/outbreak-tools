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
    outwkb As Workbook
    Data As ILinelistSpecs
    firstSheet As String 'stock the name of the first sheet (depending on the os language)
End Type

Private this As TLinelist
Private Const CLASSNAME  As String = "Linelist"

Public Function Create(lData As ILinelistSpecs) As ILinelist

    If lData Is Nothing Then ThrowError "Linelist specifications unfound"
    With New Linelist
        Set .linelistData = lData
        Set Create = .Self()
    End With

End Function

Public Property Get Self() As ILinelist
    Set Self = Me
End Property

Public Property Get linelistData() As ILinelistSpecs
    Set linelistData = this.Data
End Property

Public Property Set linelistData(ByVal lData As ILinelistSpecs)
    Set this.Data = lData
End Property

'Dictionary
Public Property Get Dictionary() As ILLdictionary
    Dim lData As ILinelistSpecs
    Set lData = linelistData()
    Set Dictionary = lData.Dictionary()
End Property

'Workbook
Public Property Get Wkb() As Workbook
    If this.outwkb Is Nothing Then
        Set this.outwkb = Workbooks.Add()
        this.firstSheet = this.outwkb.Worksheets(1).Name
    End If
    Set Wkb = this.outwkb
End Property

'Test if a worksheet Exists
Private Function sheetExists(ByVal sheetName As String) As Boolean
    Dim sh As Worksheet
    Dim outwkb As Workbook

    Set outwkb = Wkb()
    On Error Resume Next
    Set sh = outwkb.Worksheets(sheetName)
    On Error GoTo 0
    sheetExists = (Not sh Is Nothing)
End Function

'Get the dropdown object for all the choices in a worksheet
Private Property Get Dropdown() As IDropdownLists
    Dim sh As Worksheet
    Dim listSheetName As String
    Dim lData As ILinelistSpecs
    Dim drop As IDropdownLists

    Set lData = linelistData()
    listSheetName = lData.TemporarySheetName(TemporarySheetList)
    Set sh = Wksh(listSheetName)

    Set drop = DropdownLists.Create(sh)
    Set Dropdown = drop
End Property

'Add worksheets to workbook
Private Sub AddOutputSheet(ByVal sheetName As String, Optional ByVal Hide As Long = xlSheetVeryHidden)
    Dim outwkb As Workbook
    Set outwkb = Wkb()

    If sheetExists(sheetName) Then ThrowError "The sheet " & sheetName & "already exists in the output workbook"
    With outwkb
        .Worksheets.Add.Name = sheetName
        .Worksheets(sheetName).Visible = Hide
    End With
End Sub

'Get a Worksheet in the output workbook
Private Property Get Wksh(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet
    Dim outwkb As Workbook

    'Test if the sheet Exists
    If Not sheetExists(sheetName) Then ThrowError "The sheet " & sheetName & "does not exists in current workbook"
    Set outwkb = Wkb()
    Set sh = outwkb.Worksheets(sheetName)
    Set Wksh = sh
End Property


Private Sub TransferCode(moduleName As String, scope As CodeScope)
    
    Dim codeObject As Object                         'Where to stock to code in the workbook
    Dim codeContent As String
    Dim outwkb As Workbook
    Dim inpWkb As Workbook
    Dim lData As ILinelistSpecs

    Set outwkb = Wkb() 'Output workbook
    Set lData = linelistData()
    Set inpWkb = lData.specswkb()

    'get all the values within the actual module to transfer
    With inpWkb.VBProject.VBComponents(moduleName).CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With

    Select Case scope
    Case codeScopeClass
        Set codeObject = outwkb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    Case codeScopeModule
        Set codeObject = outwkb.VBProject.VBComponents.Add(vbext_ct_ClassModule)
    End Select

    'keep the name and add the codes
    codeObject.Name = moduleName
    With outwkb.VBProject.VBComponents(moduleName).CodeModule
        .DeleteLines 1, .CountOfLines
         DoEvents
        .AddFromString codeContent
    End With
End Sub


'Prepare temporary folder for the linelist creation
Public Sub TemporaryFolder(Optional Create As Boolean = True)
    Dim lData As ILinelistSpecs
    Dim mainobj As IMain
    Dim outPath As String
    Dim sep As String 'Separator

    Set lData = linelistData()
    Set mainobj = lData.MainObject()
    outPath = mainobj.OutputPath()
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
    Dim outwkb As Workbook
    Dim inpWkb As Workbook
    Dim mainobj As IMain
    Dim sep As String 'Application path separator
    Dim formPath As String

    Set lData = linelistData()
    Set outwkb = Wkb()
    Set inpWkb = lData.specswkb()
    Set manObj = lData.MainObject()
    sep = Application.PathSeparator
    outDir = mainobj.OutputPath() & sep & "LinelistApp_"

    'Output dir is where to write the form for import / export
    On Error Resume Next
    Kill outDir & sep & "CopieUsf.frm"
    Kill outDir & sep & "CopieUsf.frx"
    On Error GoTo 0
    
    formPath = outDir & sep & "CopieUsf.frm"

    'Copy the frame to the output directory
    inpWkb.VBProject.VBComponents(formName).Export formPath
    outwkb.VBProject.VBComponents.Import formPath

    DoEvents

    On Error Resume Next
    Kill outDir & sep & "CopieUsf.frm"
    Kill outDir & sep & "CopieUsf.frx"
    On Error GoTo 0
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
    Dim outwkb As Workbook
    Dim prevAlert As Boolean
    Dim classesList As BetterArray
    
    Set outwkb = Wkb()
    
    Set lData = linelistData()
    Set sheetsList = New BetterArray
    Set trans = lData.TransObject()
    Set dict = lData.Dictionary()


    'temporary sheet Names
    sheetsList.Push lData.TemporarySheetName(TemporarySingle), lData.TemporarySheetName(TemporarySheetList), _
        lData.TemporarySheetName(TemporarySheetDebug), lData.TemporarySheetName(TemporarySheetMetadata), _
        lData.TemporarySheetName(TemporarySheetAnalysis), lData.TemporarySheetName(TemporarySheetImport)

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVeryHidden
    Next
    
    'Unique sheet Names (HList or VList)
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
    outwkb.Worksheets(this.firstSheet).Delete
    Application.DisplayAlerts = prevAlert

    'Send the required classes to the linelist
    Set classesList = New BetterArray
    classesList.Push "BetterArray", "DataSheet", "IDataSheet", "LLExport", "ILLExport", "LLchoice", "ILLchoice", _
                     "LLdictionary", "ILLdictionary", "LLGeo", "ILLGeo", "LLPasswords", "ILLPasswords", _
                     "LLSheets", "ILLSheets", "LLTranslations", "ILLTranslations", "OSFiles", "IOSFiles", _
                     "Translation", "ITranslation", "LLVariables", "ILLVariables"

    For counter = classesList.LowerBound To classesList.UpperBound
        codeName = classesList.Item(counter)
        TransferCode codeName, codeScopeClass
    Next

    'Send the modules to the linelist
    classesList.Clear
    classesList.Push "Helpers", "LinelistConstants", "LinelistChange", "LinelistCustomFunctions", _
                     "LinelistExport", "LinelistGeo", "LinelistMigration", "LinelistShowHide", _
                     "LinelistTranslation", "EstheticConstants", "LinelistEvents"

    For counter = classesList.LowerBound To classesList.UpperBound
        codeName = classesList.Item(counter)
        TransferCode codeName, codeScopeModule
    Next
End Sub

Private Sub Build()

    'prepare the temporary folder
    TemporaryFolder Create:=True

    'Prepare the workbook
    Prepare



    TemporaryFolder Create:=False
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Properties
Public Property Get ILinelist_Wksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_Wksh = Wksh(sheetName)
End Property

Public Property Get ILinelist_LinelistData() As ILinelistSpecs
    Set ILinelist_LinelistData = linelistData()
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

Public Sub ILinelist_Build()
    Build
End Sub


