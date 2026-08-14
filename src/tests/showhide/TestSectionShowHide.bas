Attribute VB_Name = "TestSectionShowHide"
Attribute VB_Description = "Tests for SectionShowHide"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for SectionShowHide")

Option Explicit

'@description
'Drives SectionShowHide, the class the section button and the sections form both
'work through. It binds three classes that already have suites of their own -
'SectionMap, ShowHide and ShowHideLayout - so what is tested here is the join:
'which cell names a section, what a section's state reads as, and what hiding
'one does to the sheet.
'
'WHAT THE FIXTURE BUILDS
'-------------------------------------------------------------------------------
'The dictionary fixture is written and prepared once. Each test then takes a
'scratch worksheet, builds an entry list over hlist2D-sheet1, and writes a
'section map on that worksheet from the column indices the entry list reports.
'The columns are read back rather than written down, so a change to the fixture
'dictionary moves the sections with it.
'
'Two blocks are recorded on every sheet:
'
'  free      the two optional variables, which follow the user
'  fixed     the mandatory variable alone, which never moves
'
'A CELL ON THE TITLE ROW AND NOWHERE ELSE
'-------------------------------------------------------------------------------
'The rule the class exists to hold: a section is named by a cell on row 5 of a
'data entry sheet, and by a cell in column 2 of a vertical one. Reading a data
'row as a section is what let a second press of the button collapse the section
'the user never pointed at.
'@depends SectionShowHide, SectionMap, ShowHide, ShowHideLayout, LLdictionary, CustomTest

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private Dict As LLdictionary
Private ScratchSheets As Collection
Private ScratchTaken As Long
Private SetupError As Long
Private SetupMessage As String

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "SectionShowHide"
Private Const DICTIONARY_SHEET As String = "DictionaryFixture"

'How many scratch worksheets the pool holds. No test here asks for more than one.
Private Const SCRATCH_POOL_SIZE As Long = 2

'How far the geometry reset of a pool sheet reaches, in rows and in columns
Private Const SCRATCH_RESET_SPAN As Long = 300

Private Const HLIST_SHEET As String = "hlist2D-sheet1"
Private Const VLIST_SHEET As String = "vlist1D-sheet1"

'The line SectionBuilder writes the titles on. Repeated here so a test failure
'says which number the class disagreed with.
Private Const HLIST_TITLE_ROW As Long = 5
Private Const VLIST_TITLE_COLUMN As Long = 2

Private Const FREE_SECTION As String = "Free section"
Private Const FIXED_SECTION As String = "Fixed section"


'@section Lifecycle
'===============================================================================

'@sub-title Build the assertion harness, the fixture workbook and the dictionary.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    Dim counter As Long

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSectionShowHide"

    SetupError = 0
    SetupMessage = vbNullString

    'An error escaping this routine is a modal dialog, and a modal costs the
    'whole run. The reason is captured and every test reports it as its own.
    On Error Resume Next
        Set FixtureWorkbook = NewWorkbook()
        DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, FixtureWorkbook
        Set Dict = LLdictionary.Create(FixtureWorkbook.Worksheets(DICTIONARY_SHEET), 1, 1)
        Dict.Prepare

        Set ScratchSheets = New Collection
        For counter = 1 To SCRATCH_POOL_SIZE
            ScratchSheets.Add FixtureWorkbook.Worksheets.Add
        Next counter

        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print the results and drop the fixture workbook.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set ScratchSheets = Nothing
    Set Dict = Nothing
    Set FixtureWorkbook = Nothing
    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    ScratchTaken = 0
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    ScratchTaken = 0
End Sub


'@section What names a section
'===============================================================================

'@TestMethod("SectionShowHide")
Public Sub TestCreateRefusesAMissingPiece()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRefusesAMissingPiece"
    If Not FixtureReady("TestCreateRefusesAMissingPiece") Then Exit Sub
    On Error GoTo TestFail

    Dim raised As Long

    On Error Resume Next
        SectionShowHide.Create Nothing, Nothing, Nothing
        raised = Err.Number
    On Error GoTo TestFail

    Assert.IsTrue raised <> 0, _
                  "A toggle with no section map refuses to be built"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRefusesAMissingPiece", Err.Number, Err.Description
End Sub

'@TestMethod("SectionShowHide")
Public Sub TestOnlyTheTitleRowNamesASectionOnAnHList()
    CustomTestSetTitles Assert, TESTMODULE, "TestOnlyTheTitleRowNamesASectionOnAnHList"
    If Not FixtureReady("TestOnlyTheTitleRowNamesASectionOnAnHList") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionShowHide
    Dim sh As Worksheet
    Dim freeStart As Long

    Set sut = HListToggle(sh, freeStart)

    Assert.AreEqual CLng(1), sut.SectionAtCell(sh.Cells(HLIST_TITLE_ROW, freeStart)), _
                    "A cell on the title row names the section it stands over"
    Assert.AreEqual CLng(0), sut.SectionAtCell(sh.Cells(HLIST_TITLE_ROW + 4, freeStart)), _
                    "The same column on a data row names no section"
    Assert.AreEqual CLng(0), sut.SectionAtCell(Nothing), _
                    "No selection names no section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestOnlyTheTitleRowNamesASectionOnAnHList", Err.Number, Err.Description
End Sub

'@TestMethod("SectionShowHide")
Public Sub TestOnlyColumnTwoNamesASectionOnAVList()
    CustomTestSetTitles Assert, TESTMODULE, "TestOnlyColumnTwoNamesASectionOnAVList"
    If Not FixtureReady("TestOnlyColumnTwoNamesASectionOnAVList") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionShowHide
    Dim entries As ShowHide
    Dim secMap As SectionMap
    Dim sh As Worksheet
    Dim firstRow As Long

    Set sh = ScratchSheet()
    Set entries = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    firstRow = FirstFreePosition(entries)

    Set secMap = SectionMap.Create(sh)
    secMap.Clear
    secMap.Add FREE_SECTION, firstRow, firstRow

    Set sut = SectionShowHide.Create(secMap, entries, _
                                     ShowHideLayout.Create(sh, ShowHideLayerVList))

    Assert.AreEqual CLng(VLIST_TITLE_COLUMN), sut.TitleLine, _
                    "A vertical sheet carries its titles in column 2"
    Assert.AreEqual CLng(1), sut.SectionAtCell(sh.Cells(firstRow, VLIST_TITLE_COLUMN)), _
                    "A cell in column 2 names the section on that row"
    Assert.AreEqual CLng(0), sut.SectionAtCell(sh.Cells(firstRow, VLIST_TITLE_COLUMN + 1)), _
                    "The column beside it names no section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestOnlyColumnTwoNamesASectionOnAVList", Err.Number, Err.Description
End Sub

'@TestMethod("SectionShowHide")
Public Sub TestTitleCellSitsOverTheStartOfTheSection()
    CustomTestSetTitles Assert, TESTMODULE, "TestTitleCellSitsOverTheStartOfTheSection"
    If Not FixtureReady("TestTitleCellSitsOverTheStartOfTheSection") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionShowHide
    Dim sh As Worksheet
    Dim freeStart As Long
    Dim titleRng As Range

    Set sut = HListToggle(sh, freeStart)
    Set titleRng = sut.TitleCell(1)

    Assert.IsFalse titleRng Is Nothing, "The first section has a title cell"
    Assert.AreEqual CLng(HLIST_TITLE_ROW), titleRng.Row, _
                    "The title cell sits on the title row"
    Assert.AreEqual freeStart, titleRng.Column, _
                    "And over the first column of the section"
    Assert.IsTrue sut.TitleCell(99) Is Nothing, _
                  "An index outside the map has no title cell"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTitleCellSitsOverTheStartOfTheSection", Err.Number, Err.Description
End Sub


'@section What a section does
'===============================================================================

'@TestMethod("SectionShowHide")
Public Sub TestHidingASectionCollapsesEveryColumnOfIt()
    CustomTestSetTitles Assert, TESTMODULE, "TestHidingASectionCollapsesEveryColumnOfIt"
    If Not FixtureReady("TestHidingASectionCollapsesEveryColumnOfIt") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionShowHide
    Dim sh As Worksheet
    Dim freeStart As Long

    Set sut = HListToggle(sh, freeStart)

    Assert.IsFalse sut.IsHidden(1), "The section starts out shown"

    sut.SetHidden 1, True

    Assert.IsTrue sh.Columns(freeStart).Hidden, _
                  "The first column of the section is hidden on the sheet"
    Assert.IsTrue sut.IsHidden(1), "And the section reads as hidden"

    sut.SetHidden 1, False

    Assert.IsFalse sh.Columns(freeStart).Hidden, _
                   "Showing the section brings the column back"
    Assert.IsFalse sut.IsHidden(1), "And the section reads as shown"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHidingASectionCollapsesEveryColumnOfIt", Err.Number, Err.Description
End Sub

'@TestMethod("SectionShowHide")
Public Sub TestASectionOfMandatoryVariablesRefusesTheUser()
    CustomTestSetTitles Assert, TESTMODULE, "TestASectionOfMandatoryVariablesRefusesTheUser"
    If Not FixtureReady("TestASectionOfMandatoryVariablesRefusesTheUser") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionShowHide
    Dim sh As Worksheet
    Dim freeStart As Long
    Dim fixedStart As Long

    Set sut = HListToggle(sh, freeStart, fixedStart)

    Assert.IsTrue sut.CanChange(1), "The free section follows the user"
    Assert.IsFalse sut.CanChange(2), _
                   "A section holding one mandatory variable cannot be collapsed"

    sut.SetHidden 2, True
    Assert.IsFalse sh.Columns(fixedStart).Hidden, _
                   "And the mandatory column stays visible when it is asked to hide"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASectionOfMandatoryVariablesRefusesTheUser", Err.Number, Err.Description
End Sub

'@TestMethod("SectionShowHide")
Public Sub TestCanChangeIsFalseOutsideTheMap()
    CustomTestSetTitles Assert, TESTMODULE, "TestCanChangeIsFalseOutsideTheMap"
    If Not FixtureReady("TestCanChangeIsFalseOutsideTheMap") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionShowHide
    Dim sh As Worksheet
    Dim freeStart As Long

    Set sut = HListToggle(sh, freeStart)

    Assert.AreEqual CLng(2), sut.Count, "The fixture writes two sections"
    Assert.IsFalse sut.CanChange(0), "Index 0 names no section"
    Assert.IsFalse sut.CanChange(3), "Index 3 is past the end of the map"
    Assert.AreEqual CLng(0), sut.SetHidden(3, True), _
                    "And setting it changes nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCanChangeIsFalseOutsideTheMap", Err.Number, Err.Description
End Sub

'@TestMethod("SectionShowHide")
Public Sub TestSetAllHiddenCollapsesEverySectionTheUserOwns()
    CustomTestSetTitles Assert, TESTMODULE, "TestSetAllHiddenCollapsesEverySectionTheUserOwns"
    If Not FixtureReady("TestSetAllHiddenCollapsesEverySectionTheUserOwns") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionShowHide
    Dim sh As Worksheet
    Dim freeStart As Long
    Dim fixedStart As Long

    Set sut = HListToggle(sh, freeStart, fixedStart)

    sut.SetAllHidden True

    Assert.IsTrue sh.Columns(freeStart).Hidden, "The free section collapses"
    Assert.IsFalse sh.Columns(fixedStart).Hidden, _
                   "The mandatory column stays where it is"

    sut.SetAllHidden False
    Assert.IsFalse sh.Columns(freeStart).Hidden, "And they all come back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetAllHiddenCollapsesEverySectionTheUserOwns", Err.Number, Err.Description
End Sub

'@TestMethod("SectionShowHide")
Public Sub TestSectionNameIsCarriedBack()
    CustomTestSetTitles Assert, TESTMODULE, "TestSectionNameIsCarriedBack"
    If Not FixtureReady("TestSectionNameIsCarriedBack") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionShowHide
    Dim sh As Worksheet
    Dim freeStart As Long

    Set sut = HListToggle(sh, freeStart)

    Assert.AreEqual FREE_SECTION, sut.SectionNameAt(1), _
                    "The title of the first section reads back"
    Assert.AreEqual FIXED_SECTION, sut.SectionNameAt(2), _
                    "And the title of the second"
    Assert.IsFalse sut.Wksh Is Nothing, "The toggle knows its worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSectionNameIsCarriedBack", Err.Number, Err.Description
End Sub


'@section Fixture helpers
'===============================================================================

'@fun-title Report a broken fixture as the running test's own failure.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is usable.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If SetupError = 0 And Not Dict Is Nothing Then
        FixtureReady = True
        Exit Function
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The fixture could not be built - " & SetupMessage
End Function

'@fun-title Hand the running test the next worksheet of the pool.
'@details
'Nothing here is deleted. Worksheet.Delete is unreliable on macOS Excel and
'hangs the headless run, so a pool sheet is hard-cleared instead.
'@return Worksheet. An empty worksheet of the fixture workbook.
Private Function ScratchSheet() As Worksheet
    Dim sh As Worksheet

    ScratchTaken = ScratchTaken + 1
    If ScratchTaken > ScratchSheets.Count Then
        Err.Raise vbObjectError + 3101, "TestSectionShowHide", _
                  "The scratch pool holds " & CStr(ScratchSheets.Count) & _
                  " sheets and this test asked for " & CStr(ScratchTaken)
    End If

    Set sh = ScratchSheets.Item(ScratchTaken)
    ResetScratchSheet sh
    Set ScratchSheet = sh
End Function

'@sub-title Put one pool sheet back to empty.
'@param sh Worksheet. The sheet to clear.
Private Sub ResetScratchSheet(ByVal sh As Worksheet)
    Dim span As Range

    On Error Resume Next
        Do While sh.ListObjects.Count > 0
            sh.ListObjects(1).Delete
        Loop

        sh.Cells.Clear

        Set span = sh.Range(sh.Cells(1, 1), _
                            sh.Cells(SCRATCH_RESET_SPAN, SCRATCH_RESET_SPAN))
        span.EntireColumn.Hidden = False
        span.EntireRow.Hidden = False
    On Error GoTo 0
End Sub

'@fun-title Build the toggle every HList test works from.
'@details
'Two sections are written on a fresh scratch sheet, both derived from what the
'entry list reports rather than from column numbers written down here:
'
'  section 1   the run of optional variables, which the user owns
'  section 2   the mandatory variable alone, which never moves
'@param sh Worksheet. Set to the scratch sheet the toggle writes to.
'@param freeStart Long. Set to the first column of the free section.
'@return SectionShowHide. The toggle.
Private Function HListToggle(ByRef sh As Worksheet, _
                             ByRef freeStart As Long, _
                             Optional ByRef fixedStart As Long) As SectionShowHide
    Dim entries As ShowHide
    Dim secMap As SectionMap
    Dim freeEnd As Long
    Dim swap As Long

    Set sh = ScratchSheet()
    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET)

    freeStart = entries.PositionIndex(entries.IndexOf("opt_hid_h2"))
    freeEnd = entries.PositionIndex(entries.IndexOf("opt_vis_h2"))
    fixedStart = entries.PositionIndex(entries.IndexOf("mand_h2"))

    If freeEnd < freeStart Then
        swap = freeStart
        freeStart = freeEnd
        freeEnd = swap
    End If

    Set secMap = SectionMap.Create(sh)
    secMap.Clear
    secMap.Add FREE_SECTION, freeStart, freeEnd
    secMap.Add FIXED_SECTION, fixedStart, fixedStart

    Set HListToggle = SectionShowHide.Create(secMap, entries, _
                                             ShowHideLayout.Create(sh, ShowHideLayerHList))
End Function

'@fun-title The first position of the free entry list, for the VList test.
'@param entries ShowHide. The entry list to read.
'@return Long. The row the first free entry sits on.
Private Function FirstFreePosition(ByVal entries As ShowHide) As Long
    Dim idx As Long

    For idx = 1 To entries.EntryCount
        If entries.IsFree(idx) And entries.PositionIndex(idx) > 0 Then
            FirstFreePosition = entries.PositionIndex(idx)
            Exit Function
        End If
    Next idx
End Function
