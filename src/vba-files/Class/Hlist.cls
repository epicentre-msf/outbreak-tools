VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("All HList variables")
'@IgnoreModule

Option Explicit

Implements IHlist

Private Type THList
    shName As String
    ll As ILinelist
End Type

Private this As THList
Private Const CLASSNAME As String = "HList"
Private Const SHPGEO As String = "SHP_GeoApps"
Private Const SHPADDROWS As String = "SHP_Add200L"
Private Const SHPREMOVEFILT As String = "SHP_ClearFilters"
Private Const SHPSHOWHIDE As String = "SHP_ShowHide"
Private Const GOTOSECCODE As String = "go_to_section"


Public Function Create(ByVal sheetName As String, ll As ILinelist) As IHlist

    CheckRequirements sheetName, ll

    With New Hlist
        .Name = sheetName
        Set .LinelistObject = ll
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IHlist
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

'Next sheet of a HList
Public Property Get NextSheet() As String
  Dim dict As ILLdictionary
  Dim llshs As ILLSheets
  Dim strtrw As Long
  Dim endrw As Long
  Dim sheetName As String
  Dim sheetRng As Range

  sheetName = Name()
  Set dict = Dictionary()
  Set sheetRng = dict.DataRange("sheet name")
  Set llshs = LLSheets.Create(dict)
  strtrw = llshs.RowIndex(sheetName)
  endrw = llshs.NumberOfVars(sheetName) + strtrw - 1
  
  'Since datarange does not includes the colname, we keep endrw without adding 1
   NextSheet = sheetRng.Cells(endrw, 1)
End Property

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
    Dim sec As IHListSections
    Dim strt As Long 'Start for the section
    Dim drop As IDropdownLists
    Dim sectionsList As BetterArray 'The list of all sections in a worksheet (for a dropdown)
    Dim tableName As String
    Dim gotoSection As String
    Dim gotoSectionListName As String

    Set sh = Wksh()
    Set dict = Dictionary()
    Set llshs = LLSheets.Create(dict)
    Set ll = LinelistObject()
    Set lData = ll.linelistData()
    Set design = lData.DesignFormat()
    Set drop = ll.Dropdown()
    sheetName = Name()

    'Format the linelist and prepare it
    design.ApplyFormat sh, scope:=AllLinelistSheet

    'Split the rows and columns
    On Error Resume Next
        sh.Activate
        With ActiveWindow
            .SplitRow = llshs.DataBounds(sheetName, 1)
            .SplitColumn = 2
            .FreezePanes = True
        End With
    On Error GoTo 0

    'Add the buttons

    'Show Hide button
    Set rng = sh.Cells(1, 2)
    Set trads = lData.TransObject(TranslationOfShapes)
    Set but = Buttons.Create(rng, trads, codeName:=SHPSHOWHIDE, scope:=ButtonScopeSmall)
    but.Add
    but.Format design

    'Add Rows button
    Set rng = sh.Cells(2, 2)
    Set but = Buttons.Create(rng, trads, codeName:=SHPADDROWS, scope:=ButtonScopeSmall)
    but.Add
    but.Format design
    
    'Remove Filters
    Set rng = sh.Cells(2, 1)
    Set but = Buttons.Create(rng, trads, codeName:=SHPREMOVEFILT, scope:=ButtonScopeSmall)
    but.Add
    but.Format design

    'Geo If neededed
    hasGeo = llshs.ContainsControl(sheetName, "geo")
    If hasGeo Then
        Set rng = sh.Cells(3, 1)
        Set but = Buttons.Create(rng, trads, codeName:=SHPGEO, scope:=ButtonScopeGeo)
        but.Add
        but.Format design
    End If

    'Now Build each Section
    Set trads = lData.TransObject(TranslationOfMessages)
    gotoSection = trads.TranslatedValue("MSG_GoToSec")

    Set sectionsList = New BetterArray
    strt = llshs.RowIndex(sheetName) - 1 '- 1 because the row index includes the header row of the dictionary
    Set sec = HListSections.Create(strt, ll)
    sec.Build
    sec.Format
    If sec.Name <> vbNullString Then sectionsList.Push gotoSection & ": " & sec.Name

    Do While sec.HasNext()
        Set sec = sec.NextSection
        sec.Build
        sec.Format
        If sec.Name <> vbNullString Then sectionsList.Push gotoSection & ": " & sec.Name
    Loop

    'Add dropdown on first cell for GoTo section
    Set rng = sh.Cells(1, 1)
    tableName = llshs.SheetInfo(sheetName, 2)
    gotoSectionListName = tableName & "_" & GOTOSECCODE 'Name of the range goTo section
    rng.Name = gotoSectionListName
    rng.Value = gotoSection
    drop.Add sectionsList, gotoSectionListName
    drop.SetValidation rng, gotoSectionListName
    design.ApplyFormat rng, scope:=LinelistSelectionDropdown


    'Create the whole table



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
Public Sub IHList_Build()
    Build
End Sub

