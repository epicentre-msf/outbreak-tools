VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Build All Vlist variables")
'@IgnoreModule

Option Explicit

Implements IVlist

Private Type TVlist
    shName As String
    ll As ILinelist
End Type

Private this As TVlist
Private Const CLASSNAME As String = "Vlist"
Private Const GOTOSECCODE As String = "go_to_section"

Public Function Create(ByVal sheetName As String, ll As ILinelist) As IVlist

    CheckRequirements sheetName, ll

    With New Vlist
        .Name = sheetName
        Set .LinelistObject = ll
        Set Create = .Self
    End With

End Function

Public Property Get Self() As IVlist
    Set Self = Me
End Property

Public Property Let Name(ByVal sheetName As String)
    this.shName = sheetName
End Property

Public Property Get Name() As String
    Name = this.shName
End Property

Public Property Get LinelistObject() As ILinelist
    Set LinelistObject = this.ll
End Property

Public Property Set LinelistObject(ByVal ll As ILinelist)
    Set this.ll = ll
End Property

Private Property Get Dictionary() As ILLdictionary
    Dim ll As ILinelist
    Set ll = LinelistObject()
    Set Dictionary = ll.Dictionary()
End Property

Private Property Get Wksh() As Worksheet
    Dim ll As ILinelist
    Set ll = LinelistObject()
    Set Wksh = ll.Wksh(Name)
End Property

'Next sheet of a Vlist
Public Property Get NextSheet() As String
    Dim dict As ILLdictionary
    Dim llshs As ILLSheets
    Dim strtrw As Long
    Dim endrw As Long
    Dim sheetName As String
    Dim sheetRng As Range

    sheetName = Name()
    Set dict = Dictionary()
    Set sheetRng = dict.DataRange("sheet name", includeHeaders:=True)
    Set llshs = LLSheets.Create(dict)
    strtrw = llshs.RowIndex(sheetName)
    endrw = llshs.NumberOfVars(sheetName) + strtrw
  
    'Since datarange does includes the colname, we keep endrw without adding 1
    NextSheet = sheetRng.Cells(endrw, 1)
End Property

'Transfert codes to a worksheet

Private Sub TransferCodeWksh()

    Const LINELISTCHANGEMODULENAME = "VListChange"

    Dim sh As Worksheet                          'Actual worksheet
    Dim inpWkb As Workbook                       'Input workbook
    Dim ll As ILinelist
    Dim lData As ILinelistSpecs

    Dim codeContent As String                    'a string to contain code to add
    Dim vbProj As Object                         'component, project and modules
    Dim vbComp As Object
    Dim codeMod As Object
    Dim sheetName As String

    Set ll = LinelistObject()
    Set lData = ll.linelistData()

    Set inpWkb = lData.specswkb()
    Set sh = Wksh()
    sheetName = Name()
    

    'save the code module in the string sNouvCode
    With inpWkb.VBProject.VBComponents(LINELISTCHANGEMODULENAME).CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With

    With sh.Parent
        Set vbProj = .VBProject
        Set vbComp = vbProj.VBComponents(.sheets(sheetName).codeName)
        Set codeMod = vbComp.CodeModule
    End With

    'Adding the code
    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub

Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

Public Sub Build()

    Dim sh As Worksheet
    Dim llshs As ILLSheets
    Dim dict As ILLdictionary
    Dim ll As ILinelist
    Dim lData As ILinelistSpecs
    Dim but As IButtons
    Dim trads As ITranslation
    Dim rng As Range
    Dim design As ILLFormat
    Dim sheetName As String
    Dim hasGeo As Boolean
    Dim sec As IVListSections
    Dim strt As Long                             'Start for the section
    Dim drop As IDropdownLists
    Dim pass As ILLPasswords
    Dim sectionsList As BetterArray              'The list of all sections in a worksheet (for a dropdown)
    Dim tableName As String
    Dim goToSection As String
    Dim gotoSectionListName As String
    Dim shcolStart As Long
    Dim shcolEnd As Long
    Dim shrowStart As Long
    Dim shrowEnd As Long
    Dim Lo As ListObject
    

    Set sh = Wksh()
    Set dict = Dictionary()
    Set llshs = LLSheets.Create(dict)
    Set ll = LinelistObject()
    Set lData = ll.linelistData()
    Set design = lData.DesignFormat()
    Set drop = ll.Dropdown()
    Set pass = lData.Password()
    sheetName = Name()
    

    'Format the linelist and prepare it
    design.ApplyFormat sh, scope:=AllLinelistSheet

    'Split the rows and columns
    On Error Resume Next
    sh.Activate
    BusyApp
    With ActiveWindow
        .SplitRow = 1
        .FreezePanes = True
    End With
    On Error GoTo 0

    'Now Build each Section
    Set trads = lData.TransObject(TranslationOfMessages)
    goToSection = trads.TranslatedValue("MSG_GoToSec")

    Set sectionsList = New BetterArray
    strt = llshs.RowIndex(sheetName) - 1         '- 1 because the row index includes the header row of the dictionary
    
    Set sec = VListSections.Create(strt, ll)
    sec.Build
    sec.Format
    If sec.Name <> vbNullString Then sectionsList.Push goToSection & ": " & sec.Name

    Do While sec.HasNext()
        Set sec = sec.NextSection
        sec.Build
        sec.Format
        If sec.Name <> vbNullString Then sectionsList.Push goToSection & ": " & sec.Name
    Loop

    'Add dropdown on first cell for GoTo section
    Set rng = sh.Cells(1, 1)
    tableName = llshs.SheetInfo(sheetName, 2)
    gotoSectionListName = tableName & "_" & GOTOSECCODE 'Name of the range goTo section
    rng.Name = gotoSectionListName
    rng.Value = goToSection
    drop.Add sectionsList, gotoSectionListName
    drop.SetValidation rng, gotoSectionListName
    design.ApplyFormat rng, scope:=LinelistSelectionDropdown
    rng.Locked = False
    
    'Add the sheet type on vlist
    Set rng = sh.Cells(1, 3)
    rng.Value = "VList"
    design.ApplyFormat rng, scope:=LinelistHiddenCell
    
    'Add the table Name
    Set rng = sh.Cells(1, 4)
    rng.Value = tableName
    design.ApplyFormat rng, scope:=LinelistHiddenCell
    

    'Transfert code to the worksheet
    TransferCodeWksh
    
    'Protect the worksheet
    'pass.Protect "_active"
End Sub

'Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

Private Sub CheckRequirements(ByVal sheetName As String, ByVal ll As ILinelist)
    Dim llshs As ILLSheets
    Dim dict As ILLdictionary
    
    If ll Is Nothing Then ThrowError "Please provide a linelist object"
    Set dict = ll.Dictionary()
    Set llshs = LLSheets.Create(dict)
    If Not llshs.Contains(sheetName) Then ThrowError "The provided sheetname " & sheetName _
                                   & " is not present in the dictionary"
End Sub

'Exposed methods
Public Sub IVlist_Build()
    Build
End Sub

Public Property Get IVlist_NextSheet() As String
    IVlist_NextSheet = NextSheet()
End Property


