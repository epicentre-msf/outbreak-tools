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
Private Const OUTPUT_COLUMN As Long = 3
Private Const DATE_COLUMN As Long = 4
Private Const DETAIL_COLUMN As Long = 5

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

'@sub-title One logged event carries the action, the date and the detail.
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

    titleRow = RowOfText(sh, OUTPUT_COLUMN, "import-data")
    Assert.IsTrue (titleRow > 0), "The action code heads the block"

    entryRow = RowOfText(sh, OUTPUT_COLUMN, "Success")
    Assert.IsTrue (entryRow > titleRow), "The outcome line sits under the action"
    Assert.IsTrue IsDate(CStr(sh.Cells(entryRow, DATE_COLUMN).Value)), _
                  "The entry label opens with the date"
    Assert.AreEqual "cases.xlsx", CStr(sh.Cells(entryRow, DETAIL_COLUMN).Value), _
                    "The detail sits in the last written column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLogSuccessWritesActionDateAndDetail", Err.Number, Err.Description
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

    firstRow = RowOfText(sh, OUTPUT_COLUMN, "add-rows")
    secondRow = RowOfText(sh, OUTPUT_COLUMN, "calculate")

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
    Assert.AreEqual "range - missing", CStr(sh.Cells(entryRow, DETAIL_COLUMN).Value), _
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
    Assert.AreEqual 0&, RowOfText(sh, OUTPUT_COLUMN, "add-rows"), _
                    "The old log lines are gone"

    restartRow = RowOfText(sh, DETAIL_COLUMN, "restarted")
    eventRow = RowOfText(sh, OUTPUT_COLUMN, "open")
    Assert.IsTrue (restartRow > 0), "The fresh log opens with the restart line"
    Assert.IsTrue (eventRow > restartRow), _
                  "The event that met the cap lands after the restart line"
    Assert.IsTrue (LastOutputRow(sh) < sut.MaxEntries), _
                  "The fresh log sits far below the cap"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRotationClearsThePastLog", Err.Number, Err.Description
End Sub
