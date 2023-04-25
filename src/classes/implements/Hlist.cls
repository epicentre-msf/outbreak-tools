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
    Set sheetRng = dict.DataRange("sheet name", includeHeaders:=True)
    Set llshs = LLSheets.Create(dict)
    strtrw = llshs.RowIndex(sheetName)
    endrw = llshs.NumberOfVars(sheetName) + strtrw

    'Since datarange does includes the colname, we keep endrw without adding 1
    NextSheet = sheetRng.Cells(endrw, 1)
End Property

'Transfert codes to a worksheet

Private Sub TransferCodeWksh()

    Const LINELISTCHANGEMODULENAME = "HListChange"

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
    Set lData = ll.LinelistData()

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

'Busy app
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

Public Sub Build()

    Const ADDROWSCOMMAND As String = "ClicCmdAddRows"
    Const SHOWHIDECOMMAND As String = "ClicCmdShowHide"
    Const REMOVEFILTERCOMMAND As String = "ClearAllFilters"
    Const GEOCOMMAND As String = "ClicCmdGeoApp"
    Const FILTEREDPREFIX  As String = "f"
    Const PRINTPREFIX As String = "pr"


    Dim sh As Worksheet
    Dim filtsh As Worksheet
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
    Dim printsh As Worksheet

    Set sh = Wksh()
    Set dict = Dictionary()
    Set llshs = LLSheets.Create(dict)
    Set ll = LinelistObject()
    Set lData = ll.LinelistData()
    Set design = lData.DesignFormat()
    Set drop = ll.Dropdown()
    Set pass = lData.Password()
    sheetName = Name()

    'Add Printed Worksheet before building
    ll.AddOutputSheet sheetName, xlSheetHidden, usePrinted:=True
    Set printsh = ll.Printwksh(sheetName)

    'Format the linelist and prepare it
    design.ApplyFormat sh, scope:=AllLinelistSheet
    design.ApplyFormat printsh, scope:=HListPrintedSheet

    'Change the View of the printed worksheet
    On Error Resume Next
        printsh.Activate
        BusyApp
        ActiveWindow.View = xlPageLayoutView
    On Error GoTo 0

    'Split the rows and columns
    On Error Resume Next
    sh.Activate
    BusyApp
    With ActiveWindow
        .SplitRow = llshs.DataBounds(sheetName, 1)
        .SplitColumn = 2
        .FreezePanes = True
    End With
    On Error GoTo 0

    'Add the buttons on printed and non printed worksheet

    'Show Hide button
    Set rng = sh.Cells(1, 2)
    Set trads = lData.TransObject(TranslationOfShapes)
    Set but = Buttons.Create(outRng:=rng, codeName:=SHPSHOWHIDE, scope:=ButtonScopeSmall)
    but.Add actionCommand:=SHOWHIDECOMMAND, shapeLabel:=trads.TranslatedValue(SHPSHOWHIDE)
    but.Format design

    'Add Rows button
    Set rng = sh.Cells(2, 2)
    Set but = Buttons.Create(outRng:=rng, codeName:=SHPADDROWS, scope:=ButtonScopeSmall)
    but.Add actionCommand:=ADDROWSCOMMAND, shapeLabel:=trads.TranslatedValue(SHPADDROWS)
    but.Format design

    'Remove Filters
    Set rng = sh.Cells(2, 1)
    Set but = Buttons.Create(outRng:=rng, codeName:=SHPREMOVEFILT, scope:=ButtonScopeSmall)
    but.Add actionCommand:=REMOVEFILTERCOMMAND, shapeLabel:=trads.TranslatedValue(SHPREMOVEFILT)
    but.Format design

    'Geo If neededed
    hasGeo = llshs.ContainsVarsOf(sheetName, "geo1")
    If hasGeo Then
        Set rng = sh.Cells(3, 1)
        Set but = Buttons.Create(outRng:=rng, codeName:=SHPGEO, scope:=ButtonScopeGeo)
        but.Add actionCommand:=GEOCOMMAND, shapeLabel:=trads.TranslatedValue(SHPGEO)
        but.Format design
    End If

    'Show Hide Button on print worksheet
    Set rng = printsh.Cells(1, 2)
    Set but = Buttons.Create(outRng:=rng, codeName:=SHPSHOWHIDE, scope:=ButtonScopeSmall)
    but.Add actionCommand:=SHOWHIDECOMMAND, shapeLabel:=trads.TranslatedValue(SHPSHOWHIDE)
    but.Format design

    'Now Build each Section (including sections of printed worksheet)
    Set trads = lData.TransObject(TranslationOfMessages)
    goToSection = trads.TranslatedValue("MSG_GoToSec")

    Set sectionsList = New BetterArray
    strt = llshs.RowIndex(sheetName) - 1         '- 1 because the row index includes the header row of the dictionary
    Set sec = HListSections.Create(strt, ll)
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
    tableName = llshs.sheetInfo(sheetName, 2)
    gotoSectionListName = tableName & "_" & GOTOSECCODE 'Name of the range goTo section

    rng.Name = gotoSectionListName
    rng.Value = goToSection
    drop.Add sectionsList, gotoSectionListName
    drop.SetValidation rng, gotoSectionListName
    design.ApplyFormat rng, scope:=LinelistSelectionDropdown
    rng.Locked = False

    'Add the sheet type on HList sheet
    Set rng = sh.Cells(1, 3)
    rng.Value = "HList"
    design.ApplyFormat rng, scope:=LinelistHiddenCell

    'Add the table name
    Set rng = sh.Cells(1, 4)
    rng.Value = tableName
    design.ApplyFormat rng, scope:=LinelistHiddenCell


    'Create the whole table of the HList
    With sh
        shrowStart = llshs.DataBounds(sheetName, 1)
        shrowEnd = shrowStart + 1
        shcolStart = llshs.DataBounds(sheetName, 3)
        shcolEnd = llshs.DataBounds(sheetName, 4)

        Set rng = .Range(.Cells(shrowStart, shcolStart), .Cells(shrowEnd, shcolEnd))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = tableName

        'Define the listobject before resizing it to keep the format of the cells
        Set Lo = .ListObjects(tableName)
        shrowEnd = llshs.DataBounds(sheetName, 2)

        Set rng = .Range(.Cells(shrowStart, shcolStart), .Cells(shrowEnd, shcolEnd))
        .Calculate
        Lo.Resize rng

        design.ApplyFormat Lo, scope:=HListTable

        'Format the header range
        Set rng = Lo.HeaderRowRange
        design.ApplyFormat rng, scope:=HListHeader
    End With


    'Create the whole table of the HListPrinted
    With printsh
        shrowEnd = shrowStart + 1
        Set rng = .Range(.Cells(shrowStart, shcolStart), .Cells(shrowEnd, shcolEnd))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = PRINTPREFIX & tableName

         'Define the listobject before resizing it to keep the format of the cells
         Set Lo = .ListObjects(PRINTPREFIX & tableName)
         shrowEnd = llshs.DataBounds(sheetName, 2)

         Set rng = .Range(.Cells(shrowStart, shcolStart), .Cells(shrowEnd, shcolEnd))
         Lo.Resize rng
         design.ApplyFormat Lo, scope:=HListTable

         'Format the header range
         Set rng = Lo.HeaderRowRange
         design.ApplyFormat rng, scope:= HListPrintedHeader
    End With

    'FILTERED WORKSHEET --------------------------------------------------------

    'Define the same listObject to the filtered worksheet
    ll.AddOutputSheet FILTEREDPREFIX & sheetName, xlSheetVeryHidden
    Set filtsh = ll.Wksh(FILTEREDPREFIX & sheetName)
    With filtsh
        Set rng = .Range(Lo.Range.Address)
        rng.Value = Lo.Range.Value
        'Define listObject on the filtered sheet
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = FILTEREDPREFIX & tableName
    End With

    'Add the name of the filtered worksheet to the HList
    Set rng = sh.Cells(1, 5)
    rng.Value = filtsh.Name
    design.ApplyFormat rng, scope:=LinelistHiddenCell

    'Transfert code to the worksheet
    TransferCodeWksh

    'Add editing range for sorting
    'sh.Protection.AllowEditRanges.Add title:="EditAll" & tableName, Range:=Range(tableName & "[#All]")

    'Protect the worksheet
    pass.Protect "_active"
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

Public Property Get IHList_NextSheet() As String
    IHList_NextSheet = NextSheet()
End Property


