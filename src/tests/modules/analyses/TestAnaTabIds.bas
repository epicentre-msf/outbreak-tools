Attribute VB_Name = "TestAnaTabIds"
Attribute VB_Description = "Tests for AnaTabIds class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for AnaTabIds class")

' AnaTabIds tests focus on factory validation and basic tracking operations.
' Full integration tests require a linelist workbook with all 12 ListObjects
' and 4 named sheet ranges, exercised through TestAnalysisOutput integration.

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "AnaTabIdsFixture"

Private Assert As ICustomTest

'@section Helpers
'===============================================================================

' @description Create a fixture sheet with all required ListObjects and named ranges
'              for AnaTabIds to pass CheckRequirements validation.
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

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnaTabIds"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet FIXTURE_SHEET
    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory validation tests
'===============================================================================

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
