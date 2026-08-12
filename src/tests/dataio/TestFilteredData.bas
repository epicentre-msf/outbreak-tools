Attribute VB_Name = "TestFilteredData"
Attribute VB_Description = "Unit tests for FilteredData"

'@Folder("Tests.DataIO")
'@ModuleDescription("Unit tests for FilteredData")
'@TestModule
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument

'@description
'Validates the FilteredData class, which writes every HList table into its
'filtered companion sheet ahead of a filtered export and on every linelist
'filter refresh. Tests cover the factory, the copy itself, the hidden rows and
'empty rows the companion leaves out, a whole block of hidden rows, and the
'failure surface: a sheet with a missing filtered_sheet name or a missing
'companion is skipped and recorded in FailedSheets while the walk continues,
'and a run that skipped a sheet writes one failure line to the user log of the
'workbook. Each test builds its own fixture workbook with a small
'three-column table and tears it down, so the order the tests run in decides
'nothing.
'@depends FilteredData, HiddenNames, LLLog, TestHelpersLite, CustomTest

Option Explicit
Option Private Module

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SOURCE_SHEET As String = "hlist_source"
Private Const COMPANION_SHEET As String = "filtered_source"

Private Assert As CustomTest
Private FixtureWorkbook As Workbook


'@section Module Lifecycle
'===============================================================================

'@sub-title Initialise the test module before any tests run.
'@details
'Ensures the output worksheet exists and creates the CustomTest assertion
'object used by all test methods in this module. Called once before the
'first test executes.
'@ModuleInitialize
Public Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
End Sub

'@sub-title Tear down the module after all tests complete.
'@details
'Renders the collected results to the output worksheet, which is what the
'headless runner harvests into the results file, then restores the
'application state and releases the assertion object. Called once after
'the last test finishes.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Build a fresh fixture workbook before each test.
'@details
'A fresh workbook per test is what keeps the hidden rows of one test away
'from the next: a reused sheet keeps its row states and its hidden names.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
    Set FixtureWorkbook = NewWorkbook()
End Sub

'@sub-title Flush the assert state and delete the fixture workbook.
'@details
'Flush persists the assertions of the test that just ran into the results
'buffer. PrintResults renders that buffer alone, so a test that skips the
'flush leaves no trace in the results file.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then Assert.Flush
    If Not FixtureWorkbook Is Nothing Then DeleteWorkbook FixtureWorkbook
    Set FixtureWorkbook = Nothing
End Sub


'@section Fixture helpers
'===============================================================================

'@sub-title Build one HList source sheet with a three-column table.
'@details
'Writes an id, label and value column with one row per requested count, adds
'a ListObject over them, and stores the two worksheet-level names the sync
'reads. Row ids run 1 to rowCount, so a test can name the survivors of a
'delete by their ids. A rowCount of 0 answers a table whose data body was
'deleted, the shape an emptied data entry sheet carries.
'@param rowCount Long. The number of data rows to write.
'@param withFilteredName Boolean. Whether to store the filtered_sheet name.
'@param sheetName String. The name of the source sheet.
'@param companionName String. The value stored under filtered_sheet.
'@return Worksheet. The built source sheet.
Private Function BuildSourceSheet(ByVal rowCount As Long, _
                                  Optional ByVal withFilteredName As Boolean = True, _
                                  Optional ByVal sheetName As String = SOURCE_SHEET, _
                                  Optional ByVal companionName As String = COMPANION_SHEET) As Worksheet

    Dim sh As Worksheet
    Dim store As HiddenNames
    Dim lo As ListObject
    Dim rowCounter As Long
    Dim lastRow As Long

    Set sh = EnsureWorksheet(sheetName, FixtureWorkbook, clearSheet:=True)

    WriteRow sh.Cells(1, 1), "id", "label", "value"
    For rowCounter = 1 To rowCount
        WriteRow sh.Cells(rowCounter + 1, 1), rowCounter, "case " & rowCounter, rowCounter * 10
    Next rowCounter

    'A table needs at least one data row to be created, so the empty shape
    'is built over one blank row whose data body is then deleted
    lastRow = rowCount + 1
    If rowCount = 0 Then lastRow = 2

    Set lo = sh.ListObjects.Add(xlSrcRange, _
                                sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, 3)), , xlYes)
    If rowCount = 0 Then lo.DataBodyRange.Delete

    Set store = HiddenNames.Create(sh)
    store.EnsureName "sheet_type", "HList", HiddenNameTypeString
    store.SetValue "sheet_type", "HList"
    If withFilteredName Then
        store.EnsureName "filtered_sheet", companionName, HiddenNameTypeString
        store.SetValue "filtered_sheet", companionName
    End If

    Set BuildSourceSheet = sh
End Function

'@sub-title Build one filtered companion sheet with an empty table.
'@param companionName String. The name of the companion sheet.
'@return Worksheet. The built companion sheet.
Private Function BuildCompanionSheet( _
        Optional ByVal companionName As String = COMPANION_SHEET) As Worksheet

    Dim sh As Worksheet

    Set sh = EnsureWorksheet(companionName, FixtureWorkbook, clearSheet:=True)
    WriteRow sh.Cells(1, 1), "id", "label", "value"
    sh.ListObjects.Add xlSrcRange, sh.Range(sh.Cells(1, 1), sh.Cells(2, 3)), , xlYes

    Set BuildCompanionSheet = sh
End Function

'@sub-title The id column of a companion table, joined with commas.
'@details
'Answers an empty string for a table whose data body is gone, and the blank
'placeholder row of a freshly built companion answers a single empty id.
'One string carries the whole outcome of a sync, so each test states its
'expectation in one assertion.
'@param companionName String. The companion sheet to read.
'@return String. The ids top to bottom, e.g. "1,3,5".
Private Function CompanionIds( _
        Optional ByVal companionName As String = COMPANION_SHEET) As String

    Dim lo As ListObject
    Dim rowCounter As Long
    Dim result As String

    Set lo = FixtureWorkbook.Worksheets(companionName).ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function

    For rowCounter = 1 To lo.DataBodyRange.Rows.Count
        If LenB(result) > 0 Then result = result & ","
        result = result & CStr(lo.DataBodyRange.Cells(rowCounter, 1).Value)
    Next rowCounter

    CompanionIds = result
End Function

'@sub-title A FilteredData over the fixture workbook.
'@return FilteredData. A fresh instance.
Private Function FixtureSync() As FilteredData
    Set FixtureSync = FilteredData.Create(FixtureWorkbook)
End Function

'@sub-title Whether the user log of the fixture workbook mentions a text.
'@details
'Walks the used range of the log sheet and reads each cell with InStr. A
'workbook with no log sheet answers False.
'@param needle String. The text to look for.
'@return Boolean. True when one cell of the log carries the text.
Private Function LogMentions(ByVal needle As String) As Boolean

    Dim logSh As Worksheet
    Dim cell As Range

    If Not WorksheetExists(LLLog.SheetName, FixtureWorkbook) Then Exit Function
    Set logSh = FixtureWorkbook.Worksheets(LLLog.SheetName)

    For Each cell In logSh.UsedRange.Cells
        If InStr(1, CStr(cell.Value), needle, vbTextCompare) > 0 Then
            LogMentions = True
            Exit Function
        End If
    Next cell
End Function


'@section Factory Validation
'===============================================================================

'@sub-title Verify Create returns a valid FilteredData for a valid workbook.
'@TestMethod("FilteredData")
Public Sub FactoryCreatesWithValidArgs()
    CustomTestSetTitles Assert, "FilteredData", "FactoryCreatesWithValidArgs"
    On Error GoTo TestFail

    Dim sut As FilteredData
    Set sut = FixtureSync()
    Assert.IsNotNothing sut, "Factory should return a valid object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryCreatesWithValidArgs", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises an error when the workbook argument is Nothing.
'@TestMethod("FilteredData")
Public Sub FactoryRejectsNothingWorkbook()
    CustomTestSetTitles Assert, "FilteredData", "FactoryRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim sut As FilteredData
    On Error Resume Next
    Set sut = FilteredData.Create(Nothing)
    Assert.IsTrue Err.Number <> 0, "Should raise error for Nothing workbook"
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingWorkbook", Err.Number, Err.Description
End Sub

'@sub-title Verify FailedSheets answers an empty list ahead of the first sync.
'@TestMethod("FilteredData")
Public Sub FailedSheetsEmptyBeforeSync()
    CustomTestSetTitles Assert, "FilteredData", "FailedSheetsEmptyBeforeSync"
    On Error GoTo TestFail

    Dim sut As FilteredData
    Set sut = FixtureSync()
    Assert.AreEqual 0&, sut.FailedSheets.Length, _
                    "FailedSheets should be empty ahead of the first sync"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FailedSheetsEmptyBeforeSync", Err.Number, Err.Description
End Sub


'@section The copy
'===============================================================================

'@sub-title Verify Sync copies the whole table when every row is visible.
'@details
'Arranges three visible filled rows and a companion. Asserts the count of
'synced sheets, the companion ids, and one value cell, which pins that the
'copy carries values across.
'@TestMethod("FilteredData")
Public Sub SyncCopiesTheTableToTheCompanion()
    CustomTestSetTitles Assert, "FilteredData", "SyncCopiesTheTableToTheCompanion"
    On Error GoTo TestFail

    BuildCompanionSheet
    BuildSourceSheet 3

    Dim syncedCount As Long
    syncedCount = FixtureSync().Sync()

    Assert.AreEqual 1&, syncedCount, "One sheet should be synced"
    Assert.AreEqual "1,2,3", CompanionIds(), _
                    "The companion should hold every visible row"
    Assert.AreEqual 20, _
                    FixtureWorkbook.Worksheets(COMPANION_SHEET) _
                                   .ListObjects(1).DataBodyRange.Cells(2, 3).Value, _
                    "The companion should carry the copied values"
    Assert.IsFalse WorksheetExists(LLLog.SheetName, FixtureWorkbook), _
                   "A run that skipped nothing should write nothing to the log"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "SyncCopiesTheTableToTheCompanion", Err.Number, Err.Description
End Sub

'@sub-title Verify Sync leaves a sheet with an empty table as it is.
'@TestMethod("FilteredData")
Public Sub SyncSkipsASheetWithoutData()
    CustomTestSetTitles Assert, "FilteredData", "SyncSkipsASheetWithoutData"
    On Error GoTo TestFail

    BuildCompanionSheet
    BuildSourceSheet 0

    Dim sut As FilteredData
    Set sut = FixtureSync()

    Assert.AreEqual 0&, sut.Sync(), "A sheet with an empty table stays out of the count"
    Assert.AreEqual 0&, sut.FailedSheets.Length, _
                    "An empty table is a quiet skip"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "SyncSkipsASheetWithoutData", Err.Number, Err.Description
End Sub


'@section Hidden and empty rows
'===============================================================================

'@sub-title Verify hidden source rows leave the companion.
'@details
'Hides data rows 2 and 4 of five, which makes two separate runs. Asserts
'the companion keeps rows 1, 3 and 5 in order.
'@TestMethod("FilteredData")
Public Sub SyncDeletesHiddenRows()
    CustomTestSetTitles Assert, "FilteredData", "SyncDeletesHiddenRows"
    On Error GoTo TestFail

    BuildCompanionSheet

    Dim sh As Worksheet
    Set sh = BuildSourceSheet(5)
    sh.Rows(3).Hidden = True
    sh.Rows(5).Hidden = True

    Assert.AreEqual 1&, FixtureSync().Sync(), "One sheet should be synced"
    Assert.AreEqual "1,3,5", CompanionIds(), _
                    "The companion should keep the visible rows in order"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "SyncDeletesHiddenRows", Err.Number, Err.Description
End Sub

'@sub-title Verify empty source rows leave the companion.
'@details
'Clears data row 3 of four on the source, so its copy lands empty on the
'companion and the empty-row check takes it off.
'@TestMethod("FilteredData")
Public Sub SyncDeletesEmptyRows()
    CustomTestSetTitles Assert, "FilteredData", "SyncDeletesEmptyRows"
    On Error GoTo TestFail

    BuildCompanionSheet

    Dim sh As Worksheet
    Set sh = BuildSourceSheet(4)
    sh.Range(sh.Cells(4, 1), sh.Cells(4, 3)).ClearContents

    Assert.AreEqual 1&, FixtureSync().Sync(), "One sheet should be synced"
    Assert.AreEqual "1,2,4", CompanionIds(), _
                    "The companion should keep the filled rows in order"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "SyncDeletesEmptyRows", Err.Number, Err.Description
End Sub

'@sub-title Verify a whole block of hidden rows stays out of the companion.
'@details
'Hides data rows 2 to 4 of five, the shape a filter leaves behind. Asserts the
'companion keeps the first and last row, which pins the visible-area walk over
'its row arithmetic: SpecialCells answers two areas here and each has to land
'on the right row of the value block.
'@TestMethod("FilteredData")
Public Sub SyncDeletesAContiguousRun()
    CustomTestSetTitles Assert, "FilteredData", "SyncDeletesAContiguousRun"
    On Error GoTo TestFail

    BuildCompanionSheet

    Dim sh As Worksheet
    Set sh = BuildSourceSheet(5)
    sh.Rows("3:5").Hidden = True

    Assert.AreEqual 1&, FixtureSync().Sync(), "One sheet should be synced"
    Assert.AreEqual "1,5", CompanionIds(), _
                    "The companion should keep the rows around the run"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "SyncDeletesAContiguousRun", Err.Number, Err.Description
End Sub


'@section The failure surface
'===============================================================================

'@sub-title Verify a missing filtered_sheet name is recorded and skipped.
'@TestMethod("FilteredData")
Public Sub SyncRecordsAMissingFilteredName()
    CustomTestSetTitles Assert, "FilteredData", "SyncRecordsAMissingFilteredName"
    On Error GoTo TestFail

    BuildSourceSheet 3, withFilteredName:=False

    Dim sut As FilteredData
    Set sut = FixtureSync()

    Assert.AreEqual 0&, sut.Sync(), "A skipped sheet stays out of the count"
    Assert.AreEqual 1&, sut.FailedSheets.Length, "The skip should be recorded"
    Assert.IsTrue InStr(1, CStr(sut.FailedSheets.Item(sut.FailedSheets.LowerBound)), _
                        SOURCE_SHEET) > 0, _
                  "The record should name the skipped sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "SyncRecordsAMissingFilteredName", Err.Number, Err.Description
End Sub

'@sub-title Verify a run that skipped a sheet writes one failure line.
'@details
'Arranges an HList sheet with a missing filtered_sheet name and syncs.
'Asserts the log sheet grew in the fixture workbook and carries the action
'code and the skipped sheet's name. The clean-copy test asserts the other
'half: a run that skipped nothing writes nothing.
'@TestMethod("FilteredData")
Public Sub SyncLogsTheSkippedSheets()
    CustomTestSetTitles Assert, "FilteredData", "SyncLogsTheSkippedSheets"
    On Error GoTo TestFail

    BuildSourceSheet 3, withFilteredName:=False
    FixtureSync().Sync

    Assert.IsTrue WorksheetExists(LLLog.SheetName, FixtureWorkbook), _
                  "A skipped sheet should build the user log"
    Assert.IsTrue LogMentions("MSG_ErrUpdate"), _
                  "The log should carry the action code of the failure"
    Assert.IsTrue LogMentions(SOURCE_SHEET), _
                  "The log should name the skipped sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "SyncLogsTheSkippedSheets", Err.Number, Err.Description
End Sub

'@sub-title Verify the walk continues past a failed sheet.
'@details
'The first HList sheet points at a companion that is missing, and the second
'is healthy. Asserts the second still syncs while the first is recorded, so
'one broken sheet leaves the others synced.
'@TestMethod("FilteredData")
Public Sub SyncContinuesPastAFailedSheet()
    CustomTestSetTitles Assert, "FilteredData", "SyncContinuesPastAFailedSheet"
    On Error GoTo TestFail

    BuildSourceSheet 2, sheetName:="hlist_broken", companionName:="filtered_gone"
    BuildCompanionSheet
    BuildSourceSheet 3

    Dim sut As FilteredData
    Set sut = FixtureSync()

    Assert.AreEqual 1&, sut.Sync(), "The healthy sheet should still sync"
    Assert.AreEqual 1&, sut.FailedSheets.Length, "The broken sheet should be recorded"
    Assert.AreEqual "1,2,3", CompanionIds(), _
                    "The healthy companion should hold its rows"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "SyncContinuesPastAFailedSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify FailedSheets describes the last run alone.
'@details
'Runs a sync that records one failure, repairs the fixture, and runs again.
'Asserts the second run answers a synced sheet and an empty FailedSheets,
'which pins that the list is rebuilt per run.
'@TestMethod("FilteredData")
Public Sub SyncResetsFailedSheetsBetweenRuns()
    CustomTestSetTitles Assert, "FilteredData", "SyncResetsFailedSheetsBetweenRuns"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSourceSheet(3, withFilteredName:=False)

    Dim sut As FilteredData
    Set sut = FixtureSync()
    sut.Sync
    Assert.AreEqual 1&, sut.FailedSheets.Length, "The first run should record the skip"

    BuildCompanionSheet
    Dim store As HiddenNames
    Set store = HiddenNames.Create(sh)
    store.EnsureName "filtered_sheet", COMPANION_SHEET, HiddenNameTypeString
    store.SetValue "filtered_sheet", COMPANION_SHEET

    Assert.AreEqual 1&, sut.Sync(), "The repaired sheet should sync"
    Assert.AreEqual 0&, sut.FailedSheets.Length, _
                    "The second run should start from an empty list"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "SyncResetsFailedSheetsBetweenRuns", Err.Number, Err.Description
End Sub
