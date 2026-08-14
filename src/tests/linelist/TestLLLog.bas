Attribute VB_Name = "TestLLLog"
Attribute VB_Description = "Tests for LLLog class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLLog class")

'@description
'Drives LLLog, the store of the linelist user log. The class binds to the
'very hidden __log worksheet, builds it when the workbook has none, and
'flushes one Checking bundle per event through CheckingOutput, which
'appends across renders. The suite covers the provisioning, the append
'across two flushes, the outcome colours, the rotation past the row cap
'and the separator guard on the detail.
'
'THE THREE SECTIONS
'-------------------------------------------------------------------------------
'The title of an entry is the section of its action, and the writer is in
'compact mode, so each of the three titles is written once however many
'actions follow it, and each heads a block of its own kind. Six tests hold
'that: the action code no longer heads a block, one title covers three
'actions of one section, a second section brings its own title, an action
'coming back to a section written earlier rejoins its block rather than
'landing at the foot of the sheet, a blank row stands between two blocks,
'and the mapping answers the two boundary cases.
'
'THE ROTATION IS REACHED BY SEEDING
'-------------------------------------------------------------------------------
'The cap is ten thousand rows. The test writes one cell past the cap in
'the output column and logs one event, so the rotation runs without the
'suite writing ten thousand lines.
'@depends LLLog, CustomTest, Checking

Private Assert As CustomTest
Private FixtureWkb As Workbook

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "LLLog"
Private Const LOG_SHEET As String = "__log"
Private Const TITLE_COLUMN As Long = 2
Private Const OUTPUT_COLUMN As Long = 3
Private Const DATE_COLUMN As Long = 4
Private Const DETAIL_COLUMN As Long = 5

'The three section titles of the log sheet, held here as the suite reads
'them off the sheet and the class holds them Private.
Private Const SECTION_OPENCLOSE As String = "open/close"
Private Const SECTION_DATAIO As String = "data input/output"
Private Const SECTION_LIFECYCLE As String = "linelist lifecycle"

'@section Lifecycle
'===============================================================================

'@sub-title Set up the assertion harness.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLLLog"
End Sub

'@sub-title Print results and tear down.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Build a fresh workbook before each test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureWkb = NewWorkbook()
End Sub

'@sub-title Flush assert state and drop the fixture workbook.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not FixtureWkb Is Nothing Then DeleteWorkbook FixtureWkb
    On Error GoTo 0

    Set FixtureWkb = Nothing
End Sub

'@section Test Fixture Helpers
'===============================================================================

'@sub-title Answer the last written row of one column of a worksheet.
'@param sh Worksheet. The log worksheet.
'@return Long. The last written row of the output column.
Private Function LastOutputRow(ByVal sh As Worksheet) As Long
    LastOutputRow = sh.Cells(sh.Rows.Count, OUTPUT_COLUMN).End(xlUp).Row
End Function

'@sub-title Answer the first row whose cell in a column contains a text.
'@details
'A loop rather than Range.Find, which inherits LookIn and SearchOrder
'from the last search of the Excel session. Answers 0 on a miss.
'@param sh Worksheet. The log worksheet.
'@param columnIndex Long. The column to read.
'@param searched String. The text to look for.
'@return Long. The first matching row, or 0.
Private Function RowOfText(ByVal sh As Worksheet, ByVal columnIndex As Long, _
                           ByVal searched As String) As Long
    Dim rowIndex As Long
    Dim lastRow As Long

    lastRow = LastOutputRow(sh)
    For rowIndex = 1 To lastRow
        If InStr(1, CStr(sh.Cells(rowIndex, columnIndex).Value), searched, vbTextCompare) > 0 Then
            RowOfText = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

'@sub-title Count the rows of a column that carry a text.
'@param sh Worksheet. The log worksheet.
'@param columnIndex Long. The column to read.
'@param searched String. The text to look for.
'@return Long. How many rows carry it.
Private Function CountOfText(ByVal sh As Worksheet, ByVal columnIndex As Long, _
                             ByVal searched As String) As Long
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim total As Long

    lastRow = LastOutputRow(sh)
    For rowIndex = 1 To lastRow
        If InStr(1, CStr(sh.Cells(rowIndex, columnIndex).Value), searched, vbTextCompare) > 0 Then
            total = total + 1
        End If
    Next rowIndex

    CountOfText = total
End Function

'@section Factory Tests
'===============================================================================

'@sub-title A Nothing workbook is refused, and the number says why.
'@TestMethod("LLLog")
Public Sub TestCreateRejectsNothingWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim errNumber As Long

    On Error Resume Next
        Set sut = LLLog.Create(Nothing)
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing workbook is refused and the number names the reason"
    Assert.IsTrue (sut Is Nothing), "Nothing comes back from the refused create"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingWorkbook", Err.Number, Err.Description
End Sub

'@sub-title A workbook without a log sheet grows one, very hidden.
'@TestMethod("LLLog")
Public Sub TestCreateBuildsTheLogSheetVeryHidden()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateBuildsTheLogSheetVeryHidden"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet

    Set sut = LLLog.Create(FixtureWkb)

    Assert.IsTrue (Not sut Is Nothing), "Create gives back an instance"

    Set sh = FixtureWkb.Worksheets(LOG_SHEET)
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(sh.Visible), _
                    "The built sheet is very hidden"
    Assert.AreEqual LOG_SHEET, sut.Wksh().Name, _
                    "The instance is bound to the log sheet"
    Assert.AreEqual LOG_SHEET, sut.SheetName(), _
                    "The class answers the internal sheet name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateBuildsTheLogSheetVeryHidden", Err.Number, Err.Description
End Sub

'@sub-title An existing log sheet is reused and keeps its visibility.
'@TestMethod("LLLog")
Public Sub TestCreateBindsToAnExistingLogSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateBindsToAnExistingLogSheet"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet
    Dim countBefore As Long

    Set sh = FixtureWkb.Worksheets.Add
    sh.Name = LOG_SHEET
    countBefore = FixtureWkb.Worksheets.Count

    Set sut = LLLog.Create(FixtureWkb)

    Assert.AreEqual countBefore, FixtureWkb.Worksheets.Count, _
                    "The existing sheet is reused"
    Assert.IsTrue (sut.Wksh() Is sh), "The instance is bound to the existing sheet"
    Assert.AreEqual CLng(xlSheetVisible), CLng(sh.Visible), _
                    "An existing sheet keeps its visibility"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateBindsToAnExistingLogSheet", Err.Number, Err.Description
End Sub

'@section Logging Tests
'===============================================================================

'@sub-title One logged event carries the section, the date, the action and the detail.
'@TestMethod("LLLog")
Public Sub TestLogSuccessWritesActionDateAndDetail()
    CustomTestSetTitles Assert, TESTMODULE, "TestLogSuccessWritesActionDateAndDetail"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet
    Dim titleRow As Long
    Dim entryRow As Long

    Set sut = LLLog.Create(FixtureWkb)
    sut.LogSuccess "import-data", "cases.xlsx"
    Set sh = sut.Wksh()

    titleRow = RowOfText(sh, OUTPUT_COLUMN, SECTION_DATAIO)
    Assert.IsTrue (titleRow > 0), "The section heads the block"

    entryRow = RowOfText(sh, OUTPUT_COLUMN, "Success")
    Assert.IsTrue (entryRow > titleRow), "The outcome line sits under the section"
    Assert.IsTrue IsDate(CStr(sh.Cells(entryRow, DATE_COLUMN).Value)), _
                  "The entry label opens with the date"
    Assert.AreEqual "import-data: cases.xlsx", _
                    CStr(sh.Cells(entryRow, DETAIL_COLUMN).Value), _
                    "The action and the detail sit in the last written column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLogSuccessWritesActionDateAndDetail", Err.Number, Err.Description
End Sub

'@sub-title The action code no longer heads a block of its own.
'@details
'The complaint the sections answer is a title per action. The action code
'belongs on the entry line now, and no row of the title column carries it.
'@TestMethod("LLLog")
Public Sub TestTheActionCodeIsNoLongerATitle()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheActionCodeIsNoLongerATitle"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet

    Set sut = LLLog.Create(FixtureWkb)
    sut.LogSuccess "add-rows", "10 rows"
    Set sh = sut.Wksh()

    Assert.AreEqual 0&, RowOfText(sh, OUTPUT_COLUMN, "add-rows"), _
                    "No title row carries the action code"
    Assert.IsTrue (RowOfText(sh, OUTPUT_COLUMN, SECTION_LIFECYCLE) > 0), _
                  "The section stands in its place"
    Assert.IsTrue (RowOfText(sh, DETAIL_COLUMN, "add-rows") > 0), _
                  "The action code moved to the entry line"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheActionCodeIsNoLongerATitle", Err.Number, Err.Description
End Sub

'@sub-title A section title is written once, however many actions follow it.
'@details
'The complaint the compact writer answers. Three actions of one section
'are three flushes, and each flush is its own PrintOutput call, so the
'title is only written once because the writer looks back at the sheet
'rather than at the call it is inside. The three entry lines follow one
'another with no blank row between them.
'@TestMethod("LLLog")
Public Sub TestTheSectionTitleIsWrittenOnce()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheSectionTitleIsWrittenOnce"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet
    Dim firstRow As Long
    Dim thirdRow As Long

    Set sut = LLLog.Create(FixtureWkb)
    sut.LogSuccess "add-rows", "10 rows"
    sut.LogWarning "sort", "no header"
    sut.LogSuccess "resize", "Main"
    Set sh = sut.Wksh()

    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_LIFECYCLE), _
                    "Three actions of one section stand under one title"

    firstRow = RowOfText(sh, DETAIL_COLUMN, "add-rows")
    thirdRow = RowOfText(sh, DETAIL_COLUMN, "resize")
    Assert.AreEqual 2&, thirdRow - firstRow, _
                    "The three entries follow one another with no blank row"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSectionTitleIsWrittenOnce", Err.Number, Err.Description
End Sub

'@sub-title A second section brings its own title, and the first keeps its own.
'@details
'A title heads a block of its own kind. The last action here comes back to
'a section whose title is already up, and its line lands at the foot of
'THAT block rather than at the foot of the sheet. The hidden title column
'carries the section of every line, which is what the title dropdown of the
'output sheet filters on.
'@TestMethod("LLLog")
Public Sub TestASecondSectionBringsItsOwnTitle()
    CustomTestSetTitles Assert, TESTMODULE, "TestASecondSectionBringsItsOwnTitle"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet

    Set sut = LLLog.Create(FixtureWkb)
    sut.LogInfo "open", "linelist.xlsb"
    sut.LogSuccess "add-rows", "10 rows"
    sut.LogSuccess "export", "cases.xlsx"
    sut.LogSuccess "sort", "Age"
    Set sh = sut.Wksh()

    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_OPENCLOSE), _
                    "The session section is written once"
    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_LIFECYCLE), _
                    "The lifecycle section is written once, both entries under it"
    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_DATAIO), _
                    "The data section is written once"
    Assert.AreEqual 1&, RowOfText(sh, DETAIL_COLUMN, "sort") - _
                        RowOfText(sh, DETAIL_COLUMN, "add-rows"), _
                    "The action that came back to a written section joined its block"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASecondSectionBringsItsOwnTitle", Err.Number, Err.Description
End Sub

'@sub-title An action rejoins its own block, above the sections opened after it.
'@details
'The owner report the grouping answers. The close of a session is the last
'thing that happens, and it used to land at the foot of the sheet, under
'whatever section had been written last. It belongs with the open it closes,
'so it is written at the foot of the open/close block and stands above every
'block opened after that one.
'@TestMethod("LLLog")
Public Sub TestAnActionRejoinsItsOwnBlock()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnActionRejoinsItsOwnBlock"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet
    Dim openRow As Long
    Dim closeRow As Long
    Dim sortRow As Long

    Set sut = LLLog.Create(FixtureWkb)
    sut.LogInfo "open", "linelist.xlsb"
    sut.LogSuccess "add-rows", "10 rows"
    sut.LogSuccess "sort", "Age"
    sut.LogInfo "close", "linelist.xlsb"
    Set sh = sut.Wksh()

    openRow = RowOfText(sh, DETAIL_COLUMN, "open: linelist.xlsb")
    closeRow = RowOfText(sh, DETAIL_COLUMN, "close: linelist.xlsb")
    sortRow = RowOfText(sh, DETAIL_COLUMN, "sort")

    Assert.IsTrue (openRow > 0), "The open line is on the sheet"
    Assert.AreEqual 1&, closeRow - openRow, _
                    "The close sits at the foot of the block the open opened"
    Assert.IsTrue (closeRow < sortRow), _
                  "The block opened after the session block stands below it"
    Assert.AreEqual SECTION_OPENCLOSE, _
                    Trim$(CStr(sh.Cells(closeRow, TITLE_COLUMN).Value)), _
                    "The close line carries its section in the hidden column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnActionRejoinsItsOwnBlock", Err.Number, Err.Description
End Sub

'@sub-title A blank row stands between two blocks.
'@details
'The blocks are read as blocks, so the title of a new one is not written
'straight under the last line of the one above it.
'@TestMethod("LLLog")
Public Sub TestABlankRowStandsBetweenTwoBlocks()
    CustomTestSetTitles Assert, TESTMODULE, "TestABlankRowStandsBetweenTwoBlocks"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet
    Dim firstTitleRow As Long
    Dim secondTitleRow As Long

    Set sut = LLLog.Create(FixtureWkb)
    sut.LogInfo "open", "linelist.xlsb"
    sut.LogSuccess "add-rows", "10 rows"
    Set sh = sut.Wksh()

    firstTitleRow = RowOfText(sh, OUTPUT_COLUMN, SECTION_OPENCLOSE)
    secondTitleRow = RowOfText(sh, OUTPUT_COLUMN, SECTION_LIFECYCLE)

    Assert.IsTrue (firstTitleRow > 0), "The first block opens the sheet"
    Assert.AreEqual 3&, secondTitleRow - firstTitleRow, _
                    "A title, its one line, then a blank row, then the next title"
    Assert.AreEqual 0&, CLng(LenB(CStr(sh.Cells(secondTitleRow - 1, OUTPUT_COLUMN).Value))), _
                    "The row between the two blocks carries nothing"
    Assert.AreEqual 0&, CLng(LenB(CStr(sh.Cells(secondTitleRow - 1, TITLE_COLUMN).Value))), _
                    "The row between the two blocks belongs to no section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestABlankRowStandsBetweenTwoBlocks", Err.Number, Err.Description
End Sub

'@sub-title Every action falls in one of the three sections.
'@details
'The two boundary cases are the ones worth stating: a code named by no
'case falls in the lifecycle section, and the restart line of the
'rotation reads as a session boundary.
'@TestMethod("LLLog")
Public Sub TestActionsFallInTheThreeSections()
    CustomTestSetTitles Assert, TESTMODULE, "TestActionsFallInTheThreeSections"
    On Error GoTo TestFail

    Dim sut As LLLog

    Set sut = LLLog.Create(FixtureWkb)

    Assert.AreEqual SECTION_OPENCLOSE, sut.SectionOf("open"), _
                    "The workbook open is a session boundary"
    Assert.AreEqual SECTION_OPENCLOSE, sut.SectionOf("log"), _
                    "The restart line is a session boundary"
    Assert.AreEqual SECTION_DATAIO, sut.SectionOf("export-migration"), _
                    "A migration export moves data out"
    Assert.AreEqual SECTION_DATAIO, sut.SectionOf("clear-data"), _
                    "Clearing the data belongs with the data"
    Assert.AreEqual SECTION_LIFECYCLE, sut.SectionOf("sort"), _
                    "A sort is work inside the workbook"
    Assert.AreEqual SECTION_LIFECYCLE, sut.SectionOf("MSG_ErrUpdate"), _
                    "A code named by no case falls in the lifecycle section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestActionsFallInTheThreeSections", Err.Number, Err.Description
End Sub

'@sub-title A second flush appends under the first one.
'@TestMethod("LLLog")
Public Sub TestTwoFlushesAppend()
    CustomTestSetTitles Assert, TESTMODULE, "TestTwoFlushesAppend"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet
    Dim firstRow As Long
    Dim secondRow As Long

    Set sut = LLLog.Create(FixtureWkb)
    sut.LogSuccess "add-rows", "10 rows"
    sut.LogFailure "calculate", "err 1004"
    Set sh = sut.Wksh()

    firstRow = RowOfText(sh, DETAIL_COLUMN, "add-rows")
    secondRow = RowOfText(sh, DETAIL_COLUMN, "calculate")

    Assert.IsTrue (firstRow > 0), "The first event survives the second flush"
    Assert.IsTrue (secondRow > firstRow), "The second event lands below the first"
    Assert.IsTrue (RowOfText(sh, OUTPUT_COLUMN, "Error") > 0), _
                  "The failure line carries the error outcome"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTwoFlushesAppend", Err.Number, Err.Description
End Sub

'@sub-title Each outcome paints its entry with its own colour.
'@details
'The colours come from CheckingOutput.ResolveFormatting: green for a
'success, red for an error, orange for a warning, grey for an info line.
'The action codes of this test stay clear of the outcome words, so each
'search lands on an entry line.
'@TestMethod("LLLog")
Public Sub TestColoursFollowTheOutcome()
    CustomTestSetTitles Assert, TESTMODULE, "TestColoursFollowTheOutcome"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet

    Set sut = LLLog.Create(FixtureWkb)
    sut.LogSuccess "add-rows"
    sut.LogFailure "calculate"
    sut.LogWarning "sort"
    sut.LogInfo "open"
    Set sh = sut.Wksh()

    Assert.AreEqual RGB(0, 120, 50), _
                    sh.Cells(RowOfText(sh, OUTPUT_COLUMN, "Success"), OUTPUT_COLUMN).Font.Color, _
                    "A success line is green"
    Assert.AreEqual RGB(192, 0, 0), _
                    sh.Cells(RowOfText(sh, OUTPUT_COLUMN, "Error"), OUTPUT_COLUMN).Font.Color, _
                    "A failure line is red"
    Assert.AreEqual RGB(167, 106, 0), _
                    sh.Cells(RowOfText(sh, OUTPUT_COLUMN, "Warning"), OUTPUT_COLUMN).Font.Color, _
                    "A warning line is orange"
    Assert.AreEqual RGB(96, 97, 100), _
                    sh.Cells(RowOfText(sh, OUTPUT_COLUMN, "Info"), OUTPUT_COLUMN).Font.Color, _
                    "An info line is grey"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestColoursFollowTheOutcome", Err.Number, Err.Description
End Sub

'@sub-title A separator inside the detail stays in the last column.
'@TestMethod("LLLog")
Public Sub TestSeparatorInDetailStaysInTheLastColumn()
    CustomTestSetTitles Assert, TESTMODULE, "TestSeparatorInDetailStaysInTheLastColumn"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet
    Dim entryRow As Long

    Set sut = LLLog.Create(FixtureWkb)
    sut.LogFailure "calculate", "range -- missing"
    Set sh = sut.Wksh()

    entryRow = RowOfText(sh, OUTPUT_COLUMN, "Error")
    Assert.AreEqual "calculate: range - missing", _
                    CStr(sh.Cells(entryRow, DETAIL_COLUMN).Value), _
                    "The softened detail sits whole in the last column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSeparatorInDetailStaysInTheLastColumn", Err.Number, Err.Description
End Sub

'@sub-title An empty action code is refused, and the number says why.
'@TestMethod("LLLog")
Public Sub TestEmptyActionIsRefused()
    CustomTestSetTitles Assert, TESTMODULE, "TestEmptyActionIsRefused"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim errNumber As Long

    Set sut = LLLog.Create(FixtureWkb)

    On Error Resume Next
        sut.LogInfo vbNullString
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "An empty action code is refused and the number names the reason"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEmptyActionIsRefused", Err.Number, Err.Description
End Sub

'@section Rotation Tests
'===============================================================================

'@sub-title A flush past the cap clears the sheet and restarts the log.
'@details
'The overflow is seeded as one cell past the cap in the output column,
'so the rotation runs without the suite writing ten thousand lines.
'@TestMethod("LLLog")
Public Sub TestRotationClearsThePastLog()
    CustomTestSetTitles Assert, TESTMODULE, "TestRotationClearsThePastLog"
    On Error GoTo TestFail

    Dim sut As LLLog
    Dim sh As Worksheet
    Dim restartRow As Long
    Dim eventRow As Long

    Set sut = LLLog.Create(FixtureWkb)
    sut.LogSuccess "add-rows", "before the cap"
    Set sh = sut.Wksh()
    sh.Cells(sut.MaxEntries + 1, OUTPUT_COLUMN).Value = "overflow marker"

    sut.LogInfo "open", "after the cap"

    Assert.AreEqual 0&, CLng(LenB(CStr(sh.Cells(sut.MaxEntries + 1, OUTPUT_COLUMN).Value))), _
                    "The overflow cell was cleared"
    Assert.AreEqual 0&, RowOfText(sh, DETAIL_COLUMN, "add-rows"), _
                    "The old log lines are gone"

    restartRow = RowOfText(sh, DETAIL_COLUMN, "restarted")
    eventRow = RowOfText(sh, DETAIL_COLUMN, "open: after the cap")
    Assert.IsTrue (restartRow > 0), "The fresh log opens with the restart line"
    Assert.IsTrue (eventRow > restartRow), _
                  "The event that met the cap lands after the restart line"
    Assert.IsTrue (LastOutputRow(sh) < sut.MaxEntries), _
                  "The fresh log sits far below the cap"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRotationClearsThePastLog", Err.Number, Err.Description
End Sub
