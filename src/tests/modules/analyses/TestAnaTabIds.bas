Attribute VB_Name = "TestAnaTabIds"
Attribute VB_Description = "Tests for AnaTabIds class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for AnaTabIds class")

'@description
'Validates the AnaTabIds class, which records analysis table and graph metadata
'into ListObjects on a hidden tracking worksheet for use during linelist export.
'Tests cover factory guard behaviour (Nothing worksheet, with and without
'CheckRequirements validation) and the AddTableInfos method that appends named
'range entries to the tracking ListObject. The BuildFixtureSheet helper creates
'a fully populated fixture with all 12 ListObjects and 4 named ranges so that
'AnaTabIds.Create with check:=True can pass validation without a real linelist.
'@depends AnaTabIds, IAnaTabIds, BetterArray, CustomTest, TestHelpers

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "AnaTabIdsFixture"

Private Assert As ICustomTest

'@section Helpers
'===============================================================================

'@sub-title Build a fixture sheet with all required ListObjects and named ranges
'@details
'Creates (or clears) a hidden worksheet and populates it with the 12 ListObjects
'that AnaTabIds expects (3 prefixes -- tab_ids, graph_ids, graph_formats -- times
'4 scope suffixes -- uba, sp, ts, sptemp). Each ListObject has columns "id",
'"name", and "export" with one empty data row. After the tables, four named
'ranges (RNG_SheetUAName, RNG_SheetTSName, RNG_SheetSPName, RNG_SheetSPTempName)
'are created pointing back to the fixture sheet itself, making the fixture
'self-contained for factory validation tests.
Private Function BuildFixtureSheet() As Worksheet
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim loNames As Variant
    Dim counter As Long
    Dim wb As Workbook
    Dim outSheetName As String

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)
    Set wb = sh.Parent

    ' Create all 12 required ListObjects (each with 1 header row + 1 empty data row)
    loNames = Array("tab_ids_uba", "tab_ids_sp", "tab_ids_ts", "tab_ids_sptemp", _
                    "graph_ids_uba", "graph_ids_sp", "graph_ids_ts", "graph_ids_sptemp", _
                    "graph_formats_uba", "graph_formats_sp", "graph_formats_ts", "graph_formats_sptemp")

    Dim startRow As Long
    startRow = 1

    For counter = LBound(loNames) To UBound(loNames)
        sh.Cells(startRow, 1).Value = "id"
        sh.Cells(startRow, 2).Value = "name"
        sh.Cells(startRow, 3).Value = "export"
        sh.Cells(startRow + 1, 1).Value = vbNullString
        Set lo = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                     Source:=sh.Range(sh.Cells(startRow, 1), sh.Cells(startRow + 1, 3)), _
                                     XlListObjectHasHeaders:=xlYes)
        lo.Name = CStr(loNames(counter))
        startRow = startRow + 3
    Next

    ' Create named ranges pointing to the fixture sheet itself (self-referential for testing)
    outSheetName = sh.Name
    sh.Cells(startRow, 1).Value = outSheetName
    sh.Cells(startRow, 1).Name = "RNG_SheetUAName"
    sh.Cells(startRow + 1, 1).Value = outSheetName
    sh.Cells(startRow + 1, 1).Name = "RNG_SheetTSName"
    sh.Cells(startRow + 2, 1).Value = outSheetName
    sh.Cells(startRow + 2, 1).Name = "RNG_SheetSPName"
    sh.Cells(startRow + 3, 1).Value = outSheetName
    sh.Cells(startRow + 3, 1).Name = "RNG_SheetSPTempName"

    Set BuildFixtureSheet = sh
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Set up the test output sheet and assertion harness
'@details
'Creates the shared test output worksheet (if absent), initialises the
'CustomTest assertion object, and registers the module name for result
'grouping. Called once before all tests in this module run.
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnaTabIds"
End Sub

'@sub-title Print results and tear down shared fixtures
'@details
'Prints accumulated test results to the output sheet, deletes the fixture
'worksheet created by BuildFixtureSheet, restores the Excel application state,
'and releases the assertion object. Called once after all tests have run.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet FIXTURE_SHEET
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updates before each test
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush pending assertions after each test
'@details
'Ensures that any assertions recorded during the test are written to the
'output sheet before the next test begins.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory validation tests
'===============================================================================

'@sub-title Verify Create rejects a Nothing worksheet
'@details
'Arranges by passing Nothing as the worksheet argument to AnaTabIds.Create
'under On Error Resume Next. Acts by attempting factory creation. Asserts
'that the returned reference is Nothing, confirming the guard clause
'prevents instantiation when no valid worksheet is supplied.
'@TestMethod("AnaTabIds")
Public Sub TestCreateRejectsNothingWorksheet()
    CustomTestSetTitles Assert, "AnaTabIds", "TestCreateRejectsNothingWorksheet"
    On Error GoTo TestFail

    On Error Resume Next
    Dim ids As IAnaTabIds
    Set ids = AnaTabIds.Create(Nothing)
    On Error GoTo 0

    Assert.IsTrue (ids Is Nothing), _
                  "Create with Nothing worksheet should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingWorksheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Create succeeds with check enabled on a valid fixture
'@details
'Arranges by calling BuildFixtureSheet to produce a worksheet with all 12
'ListObjects and 4 named ranges. Acts by calling AnaTabIds.Create with
'check:=True so CheckRequirements validation runs. Asserts the returned
'IAnaTabIds reference is not Nothing, confirming the fixture satisfies all
'structural requirements.
'@TestMethod("AnaTabIds")
Public Sub TestCreateWithCheckPasses()
    CustomTestSetTitles Assert, "AnaTabIds", "TestCreateWithCheckPasses"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildFixtureSheet()

    Dim ids As IAnaTabIds
    Set ids = AnaTabIds.Create(sh, check:=True)

    Assert.IsTrue (Not ids Is Nothing), _
                  "Create with valid fixture sheet should succeed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateWithCheckPasses", Err.Number, Err.Description
End Sub

'@sub-title Verify Create succeeds without check on any worksheet
'@details
'Arranges by creating a bare worksheet with no ListObjects or named ranges.
'Acts by calling AnaTabIds.Create with check:=False to skip structural
'validation. Asserts the returned IAnaTabIds reference is not Nothing,
'confirming that factory creation works on any worksheet when validation
'is bypassed.
'@TestMethod("AnaTabIds")
Public Sub TestCreateWithoutCheck()
    CustomTestSetTitles Assert, "AnaTabIds", "TestCreateWithoutCheck"
    On Error GoTo TestFail

    ' Any worksheet should work without check
    Dim sh As Worksheet
    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    Dim ids As IAnaTabIds
    Set ids = AnaTabIds.Create(sh, check:=False)

    Assert.IsTrue (Not ids Is Nothing), _
                  "Create without check should succeed on any worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateWithoutCheck", Err.Number, Err.Description
End Sub

'@section AddTableInfos tests
'===============================================================================

'@sub-title Verify AddTableInfos appends rows to the tracking ListObject
'@details
'Arranges by building the full fixture sheet and creating an AnaTabIds instance
'with validation enabled. Prepares a BetterArray containing three range names
'(TITLE, ROW_CATEGORIES, VALUES_COL_1) for a test table. Acts by calling
'AddTableInfos with AnalysisIdsScopeNormal scope and the prepared range names.
'Asserts that the "tab_ids_uba" ListObject has been resized to at least 3 rows,
'confirming that each named range was tracked as a separate entry.
'@TestMethod("AnaTabIds")
Public Sub TestAddTableInfosResizesListObject()
    CustomTestSetTitles Assert, "AnaTabIds", "TestAddTableInfosResizesListObject"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildFixtureSheet()

    Dim ids As IAnaTabIds
    Set ids = AnaTabIds.Create(sh, check:=True)

    Dim rangeNames As BetterArray
    Set rangeNames = New BetterArray
    rangeNames.Push "TITLE_test1", "ROW_CATEGORIES_test1", "VALUES_COL_1_test1"

    ids.AddTableInfos scope:=AnalysisIdsScopeNormal, tabId:="test1", _
                       tabRangesNames:=rangeNames

    ' Verify the ListObject was resized
    Dim lo As ListObject
    Set lo = sh.ListObjects("tab_ids_uba")

    Assert.IsTrue lo.ListRows.Count >= 3, _
                  "AddTableInfos should add rows to the tracking ListObject"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddTableInfosResizesListObject", Err.Number, Err.Description
End Sub
