Attribute VB_Name = "TestSectionMap"
Attribute VB_Description = "Tests for SectionMap class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for SectionMap class")

Option Explicit

'@description
'Drives SectionMap, which records where each section of a worksheet begins and
'ends and reads those boundaries back.
'
'THE SHEETS ARE A POOL AND NOTHING IS DELETED
'-------------------------------------------------------------------------------
'The same shape TestShowHide uses. Worksheet.Delete is unreliable on macOS
'Excel, so the pool sheets are made once in ModuleInitialize and hard-cleared as
'each test takes one. The reset drops the sheet's defined names as well as its
'cells, because the names ARE the map: a test reading a sheet a previous test
'wrote would read that test's sections.
'
'WHAT THE TESTS ARE WATCHING FOR
'-------------------------------------------------------------------------------
'That a block survives being written and read back through the hidden names,
'and that two blocks sharing one section title stay apart.
'
'A prepared dictionary hands SectionBuilder one run per title, because
'LLdictionary.Prepare sorts by main section before it derives `column index`.
'The blocks below are built by hand, so the same-title case is reached here and
'the class is held to keeping the two apart whatever it is handed.
'@depends SectionMap, HiddenNames, BetterArray, CustomTest

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private ScratchSheets As Collection
Private ScratchTaken As Long
Private SetupError As Long
Private SetupMessage As String

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "SectionMap"

'How many scratch worksheets the pool holds. The test that takes the most asks
'for one, but the pool costs nothing and a second is kept for headroom.
Private Const SCRATCH_POOL_SIZE As Long = 2


'@section Lifecycle
'===============================================================================

'@sub-title Build the assertion harness, the fixture workbook and the sheet pool.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    Dim counter As Long

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSectionMap"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set FixtureWorkbook = NewWorkbook()

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


'@section Factory
'===============================================================================

'@TestMethod("SectionMap")
Public Sub TestCreateRaisesWithoutAWorksheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRaisesWithoutAWorksheet"
    If Not FixtureReady("TestCreateRaisesWithoutAWorksheet") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As SectionMap
    Set sut = SectionMap.Create(Nothing)

    Assert.LogFailure "Create should raise when the worksheet is Nothing."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), CLng(Err.Number), _
                    "A missing worksheet should raise ObjectNotInitialized - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub

'@TestMethod("SectionMap")
Public Sub TestAFreshSheetCarriesNoSections()
    CustomTestSetTitles Assert, TESTMODULE, "TestAFreshSheetCarriesNoSections"
    If Not FixtureReady("TestAFreshSheetCarriesNoSections") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionMap

    Set sut = SectionMap.Create(ScratchSheet())

    Assert.AreEqual CLng(0), sut.Count, _
                    "A sheet that was never built carries no sections"
    Assert.AreEqual CLng(0), sut.IndexAtPosition(3), _
                    "And no position falls in a block"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFreshSheetCarriesNoSections", Err.Number, Err.Description
End Sub


'@section Recording a block
'===============================================================================

'@TestMethod("SectionMap")
Public Sub TestAddRecordsOneBlock()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddRecordsOneBlock"
    If Not FixtureReady("TestAddRecordsOneBlock") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionMap
    Dim filedAt As Long

    Set sut = SectionMap.Create(ScratchSheet())
    filedAt = sut.Add("Demographics", 2, 7)

    Assert.AreEqual CLng(1), filedAt, "The first block is filed at index 1"
    Assert.AreEqual CLng(1), sut.Count, "And the map now holds one block"
    Assert.AreEqual "Demographics", sut.SectionNameAt(1), _
                    "The block carries the title it was given"
    Assert.AreEqual CLng(2), sut.StartAt(1), "And its first position"
    Assert.AreEqual CLng(7), sut.EndAt(1), "And its last position"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddRecordsOneBlock", Err.Number, Err.Description
End Sub

'@TestMethod("SectionMap")
Public Sub TestAddRejectsAStartBeforeTheSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddRejectsAStartBeforeTheSheet"
    If Not FixtureReady("TestAddRejectsAStartBeforeTheSheet") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As SectionMap

    Set sut = SectionMap.Create(ScratchSheet())
    sut.Add "Nowhere", 0, 4

    Assert.LogFailure "Add should raise for a section starting before position 1."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), CLng(Err.Number), _
                    "A start below 1 should raise InvalidArgument - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub

'@TestMethod("SectionMap")
Public Sub TestAddRejectsAnEndBeforeItsStart()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddRejectsAnEndBeforeItsStart"
    If Not FixtureReady("TestAddRejectsAnEndBeforeItsStart") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As SectionMap

    Set sut = SectionMap.Create(ScratchSheet())
    sut.Add "Backwards", 9, 4

    Assert.LogFailure "Add should raise for a section ending before it starts."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), CLng(Err.Number), _
                    "An end before the start should raise InvalidArgument - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub

'@TestMethod("SectionMap")
Public Sub TestAnIndexOutsideTheMapIsRefused()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnIndexOutsideTheMapIsRefused"
    If Not FixtureReady("TestAnIndexOutsideTheMapIsRefused") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As SectionMap
    Dim readBack As Long

    Set sut = SectionMap.Create(ScratchSheet())
    sut.Add "Only one", 1, 3
    Assert.AreEqual CLng(1), sut.StartAt(1), "The one block reads back"

    readBack = sut.StartAt(2)

    Assert.LogFailure "StartAt should raise for an index outside the map, " & _
                      "and it answered " & CStr(readBack) & " instead."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), CLng(Err.Number), _
                    "An index outside the map should raise InvalidArgument - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub


'@section What the sheet keeps
'===============================================================================

'@TestMethod("SectionMap")
Public Sub TestTheMapSurvivesANewReader()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheMapSurvivesANewReader"
    If Not FixtureReady("TestTheMapSurvivesANewReader") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim writer As SectionMap
    Dim reader As SectionMap

    Set sh = ScratchSheet()

    Set writer = SectionMap.Create(sh)
    writer.Add "Demographics", 2, 7
    writer.Add "Laboratory", 8, 11

    'A second reader of the same sheet stands for the next session: the map
    'lives in the sheet's hidden names, not in the instance that wrote it.
    Set reader = SectionMap.Create(sh)

    Assert.AreEqual CLng(2), reader.Count, _
                    "Both blocks are still on the sheet"
    Assert.AreEqual "Demographics", reader.SectionNameAt(1), _
                    "The first block kept its title"
    Assert.AreEqual CLng(8), reader.StartAt(2), _
                    "And the second its first position"
    Assert.AreEqual CLng(11), reader.EndAt(2), _
                    "And its last"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheMapSurvivesANewReader", Err.Number, Err.Description
End Sub

'@TestMethod("SectionMap")
Public Sub TestAHandedInStoreReadsTheSameBlocks()
    CustomTestSetTitles Assert, TESTMODULE, "TestAHandedInStoreReadsTheSameBlocks"
    If Not FixtureReady("TestAHandedInStoreReadsTheSameBlocks") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim writer As SectionMap
    Dim store As HiddenNames
    Dim reader As SectionMap

    Set sh = ScratchSheet()

    Set writer = SectionMap.Create(sh)
    writer.Add "Demographics", 2, 7
    writer.Add "Laboratory", 8, 11

    'The linelist button module holds one store per sheet and hands it in, so
    'the map does not walk every name of the sheet to read a dozen back. The
    'blocks it answers have to be the ones a map building its own store reads.
    Set store = HiddenNames.Create(sh)
    Set reader = SectionMap.Create(sh, store)

    Assert.AreEqual CLng(2), reader.Count, _
                    "A map reading through a handed-in store finds both blocks"
    Assert.AreEqual "Demographics", reader.SectionNameAt(1), _
                    "The first block carries its title"
    Assert.AreEqual CLng(2), reader.StartAt(1), _
                    "And its first position"
    Assert.AreEqual CLng(8), reader.StartAt(2), _
                    "The second its first position"
    Assert.AreEqual CLng(11), reader.EndAt(2), _
                    "And its last"
    Assert.AreEqual CLng(2), reader.IndexAtPosition(9), _
                    "And a lookup lands in the block it should"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAHandedInStoreReadsTheSameBlocks", Err.Number, Err.Description
End Sub

'@TestMethod("SectionMap")
Public Sub TestAStoreOfAnotherSheetIsRefused()
    CustomTestSetTitles Assert, TESTMODULE, "TestAStoreOfAnotherSheetIsRefused"
    If Not FixtureReady("TestAStoreOfAnotherSheetIsRefused") Then Exit Sub
    On Error GoTo ExpectError

    Dim mapped As Worksheet
    Dim other As Worksheet
    Dim sut As SectionMap

    Set mapped = ScratchSheet()
    Set other = ScratchSheet()

    'A store reads the names of the sheet it was built over. Handed one built
    'over another sheet, the map would answer that sheet's blocks under this
    'sheet's name, so it refuses the store instead.
    Set sut = SectionMap.Create(mapped, HiddenNames.Create(other))

    Assert.LogFailure "Create should raise when the store belongs to another sheet."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), CLng(Err.Number), _
                    "A store of another sheet should raise InvalidArgument - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub

'@TestMethod("SectionMap")
Public Sub TestClearDropsEveryBlock()
    CustomTestSetTitles Assert, TESTMODULE, "TestClearDropsEveryBlock"
    If Not FixtureReady("TestClearDropsEveryBlock") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sut As SectionMap

    Set sh = ScratchSheet()

    Set sut = SectionMap.Create(sh)
    sut.Add "Demographics", 2, 7
    sut.Add "Laboratory", 8, 11
    sut.Clear

    Assert.AreEqual CLng(0), sut.Count, "Clear empties the instance"
    Assert.AreEqual CLng(0), SectionMap.Create(sh).Count, _
                    "And the sheet with it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestClearDropsEveryBlock", Err.Number, Err.Description
End Sub

'@TestMethod("SectionMap")
Public Sub TestARebuildLeavesNoStaleBlock()
    CustomTestSetTitles Assert, TESTMODULE, "TestARebuildLeavesNoStaleBlock"
    If Not FixtureReady("TestARebuildLeavesNoStaleBlock") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim firstBuild As SectionMap
    Dim secondBuild As SectionMap
    Dim reader As SectionMap

    Set sh = ScratchSheet()

    Set firstBuild = SectionMap.Create(sh)
    firstBuild.Add "Demographics", 2, 7
    firstBuild.Add "Laboratory", 8, 11
    firstBuild.Add "Outcome", 12, 15

    'A rebuild with fewer sections. The tail of the old map points at positions
    'the new sheet gives to something else, so it has to go.
    Set secondBuild = SectionMap.Create(sh)
    secondBuild.Clear
    secondBuild.Add "Everything", 2, 6

    Set reader = SectionMap.Create(sh)

    Assert.AreEqual CLng(1), reader.Count, _
                    "The rebuilt sheet carries one block, not four"
    Assert.AreEqual "Everything", reader.SectionNameAt(1), _
                    "And it is the block the rebuild wrote"
    Assert.AreEqual CLng(0), reader.IndexAtPosition(14), _
                    "A position the old third block covered falls in nothing now"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARebuildLeavesNoStaleBlock", Err.Number, Err.Description
End Sub


'@section Finding a block
'===============================================================================

'@TestMethod("SectionMap")
Public Sub TestIndexAtPositionFindsTheBlock()
    CustomTestSetTitles Assert, TESTMODULE, "TestIndexAtPositionFindsTheBlock"
    If Not FixtureReady("TestIndexAtPositionFindsTheBlock") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionMap

    Set sut = SectionMap.Create(ScratchSheet())
    sut.Add "Demographics", 2, 7
    sut.Add "Laboratory", 8, 11

    Assert.AreEqual CLng(1), sut.IndexAtPosition(2), _
                    "The first position of a block is in it"
    Assert.AreEqual CLng(1), sut.IndexAtPosition(5), _
                    "And so is a position in the middle"
    Assert.AreEqual CLng(1), sut.IndexAtPosition(7), _
                    "And the last one"
    Assert.AreEqual CLng(2), sut.IndexAtPosition(8), _
                    "The next position belongs to the next block"
    Assert.AreEqual CLng(0), sut.IndexAtPosition(1), _
                    "A position before every block falls in none"
    Assert.AreEqual CLng(0), sut.IndexAtPosition(40), _
                    "And so does one past them all"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIndexAtPositionFindsTheBlock", Err.Number, Err.Description
End Sub

'@TestMethod("SectionMap")
Public Sub TestTwoBlocksMayShareASingleName()
    CustomTestSetTitles Assert, TESTMODULE, "TestTwoBlocksMayShareASingleName"
    If Not FixtureReady("TestTwoBlocksMayShareASingleName") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionMap

    'Three blocks by hand, two of them titled "Controls" with "Status" between
    'them. A prepared dictionary would have brought the two Controls runs
    'together before the build saw them, so this shape is built here to hold the
    'class to keeping same-titled blocks apart on its own.
    Set sut = SectionMap.Create(ScratchSheet())
    sut.Add "Controls", 2, 7
    sut.Add "Status", 8, 11
    sut.Add "Controls", 12, 13

    Assert.AreEqual CLng(3), sut.Count, _
                    "A repeated title makes a block of its own"
    Assert.AreEqual CLng(1), sut.IndexAtPosition(5), _
                    "A position of the first run answers the first block"
    Assert.AreEqual CLng(3), sut.IndexAtPosition(12), _
                    "And one of the second run answers the third"
    Assert.AreEqual sut.SectionNameAt(1), sut.SectionNameAt(3), _
                    "Both still read as the same section title"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTwoBlocksMayShareASingleName", Err.Number, Err.Description
End Sub


'@section Finding every block a selection touches
'===============================================================================

'@TestMethod("SectionMap")
Public Sub TestIndicesInRangeCoversAGap()
    CustomTestSetTitles Assert, TESTMODULE, "TestIndicesInRangeCoversAGap"
    If Not FixtureReady("TestIndicesInRangeCoversAGap") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionMap
    Dim touched As BetterArray

    Set sut = SectionMap.Create(ScratchSheet())
    sut.Add "Demographics", 2, 7
    sut.Add "Laboratory", 8, 11
    sut.Add "Outcome", 12, 15

    'This is how a user reaches a section every position of which is hidden:
    'select from the visible position before the gap to the visible one after
    'it. The middle block has to come back even though nothing inside it can be
    'clicked.
    Set touched = sut.IndicesInRange(6, 13)

    Assert.AreEqual CLng(3), CLng(touched.Length), _
                    "A span across the gap touches all three blocks"
    Assert.AreEqual CLng(1), CLng(touched.Item(touched.LowerBound)), _
                    "The blocks come back in map order"
    Assert.AreEqual CLng(3), CLng(touched.Item(touched.UpperBound)), _
                    "Down to the last one the span reaches"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIndicesInRangeCoversAGap", Err.Number, Err.Description
End Sub

'@TestMethod("SectionMap")
Public Sub TestIndicesInRangeReadsItsBoundsEitherWay()
    CustomTestSetTitles Assert, TESTMODULE, "TestIndicesInRangeReadsItsBoundsEitherWay"
    If Not FixtureReady("TestIndicesInRangeReadsItsBoundsEitherWay") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionMap
    Dim forwards As BetterArray
    Dim backwards As BetterArray

    Set sut = SectionMap.Create(ScratchSheet())
    sut.Add "Demographics", 2, 7
    sut.Add "Laboratory", 8, 11

    Set forwards = sut.IndicesInRange(3, 9)
    Set backwards = sut.IndicesInRange(9, 3)

    Assert.AreEqual CLng(2), CLng(forwards.Length), _
                    "A span read left to right touches both blocks"
    Assert.AreEqual CLng(forwards.Length), CLng(backwards.Length), _
                    "And a span made the other way round touches the same"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIndicesInRangeReadsItsBoundsEitherWay", Err.Number, Err.Description
End Sub

'@TestMethod("SectionMap")
Public Sub TestIndicesInRangeIsEmptyOffTheMap()
    CustomTestSetTitles Assert, TESTMODULE, "TestIndicesInRangeIsEmptyOffTheMap"
    If Not FixtureReady("TestIndicesInRangeIsEmptyOffTheMap") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As SectionMap
    Dim touched As BetterArray

    Set sut = SectionMap.Create(ScratchSheet())
    sut.Add "Demographics", 2, 7

    Set touched = sut.IndicesInRange(20, 25)

    Assert.IsTrue Not touched Is Nothing, _
                  "A span touching nothing still answers a list"
    Assert.AreEqual CLng(0), CLng(touched.Length), _
                    "And that list is empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIndicesInRangeIsEmptyOffTheMap", Err.Number, Err.Description
End Sub


'@section What a section title may carry
'===============================================================================

'@TestMethod("SectionMap")
Public Sub TestATitleCarryingTheSeparatorArrivesWhole()
    CustomTestSetTitles Assert, TESTMODULE, "TestATitleCarryingTheSeparatorArrivesWhole"
    If Not FixtureReady("TestATitleCarryingTheSeparatorArrivesWhole") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim writer As SectionMap

    Set sh = ScratchSheet()

    'The stored form is "<start>|<end>|<name>", so a title carrying the
    'separator would be cut at its first one were the split not capped.
    Set writer = SectionMap.Create(sh)
    writer.Add "Signs | symptoms", 2, 7

    Assert.AreEqual "Signs | symptoms", SectionMap.Create(sh).SectionNameAt(1), _
                    "A title carrying the separator reads back whole"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestATitleCarryingTheSeparatorArrivesWhole", Err.Number, Err.Description
End Sub

'@TestMethod("SectionMap")
Public Sub TestALongTitleIsCutButTheBoundariesSurvive()
    CustomTestSetTitles Assert, TESTMODULE, "TestALongTitleIsCutButTheBoundariesSurvive"
    If Not FixtureReady("TestALongTitleIsCutButTheBoundariesSurvive") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim writer As SectionMap
    Dim reader As SectionMap
    Dim longTitle As String

    Set sh = ScratchSheet()
    longTitle = String$(400, "x")

    'A defined name carries its value in its RefersTo formula and that string is
    'length capped. The title is what gets cut, so the two numbers - the only
    'part anything acts on - always survive.
    Set writer = SectionMap.Create(sh)
    writer.Add longTitle, 2, 7

    Set reader = SectionMap.Create(sh)

    Assert.AreEqual CLng(1), reader.Count, "The block was still recorded"
    Assert.AreEqual CLng(2), reader.StartAt(1), "With its first position intact"
    Assert.AreEqual CLng(7), reader.EndAt(1), "And its last"
    Assert.IsTrue Len(reader.SectionNameAt(1)) < Len(longTitle), _
                  "And a title too long for a defined name was cut to fit"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALongTitleIsCutButTheBoundariesSurvive", Err.Number, Err.Description
End Sub


'@section Fixture helpers
'===============================================================================

'@fun-title Report a fixture that could not be built, once per test.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is there.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If SetupError = 0 And Not ScratchSheets Is Nothing Then
        FixtureReady = True
        Exit Function
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The fixture could not be built - " & SetupMessage
End Function

'@fun-title Hand the running test the next worksheet of the pool.
'@return Worksheet. A worksheet of the fixture workbook carrying no map.
Private Function ScratchSheet() As Worksheet
    Dim sh As Worksheet

    ScratchTaken = ScratchTaken + 1
    If ScratchTaken > ScratchSheets.Count Then
        Err.Raise vbObjectError + 3101, "TestSectionMap", _
                  "The scratch pool holds " & CStr(ScratchSheets.Count) & _
                  " sheets and this test asked for " & CStr(ScratchTaken)
    End If

    Set sh = ScratchSheets.Item(ScratchTaken)
    ResetScratchSheet sh
    Set ScratchSheet = sh
End Function

'@sub-title Put one pool sheet back to empty.
'@details
'The defined names go as well as the cells. The names ARE the map, so a sheet
'cleared of its contents alone would still hand the next test the sections the
'last one wrote.
'@param sh Worksheet. The sheet to clear.
Private Sub ResetScratchSheet(ByVal sh As Worksheet)
    Dim counter As Long

    On Error Resume Next
        sh.Cells.Clear

        For counter = sh.Names.Count To 1 Step -1
            sh.Names(counter).Delete
        Next counter
    On Error GoTo 0
End Sub
