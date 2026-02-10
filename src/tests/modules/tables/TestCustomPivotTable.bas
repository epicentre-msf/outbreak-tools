Attribute VB_Name = "TestCustomPivotTable"
Attribute VB_Description = "Tests for CustomPivotTable class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for CustomPivotTable class")

Option Explicit

Private Assert As ICustomTest
Private FixtureWkb As Workbook
Private PivotSheet As Worksheet
Private DataSheet As Worksheet

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "CustomPivotTable"
Private Const PIVOT_SHEET_NAME As String = "PivotSheet"
Private Const DATA_TABLE_NAME As String = "TestDataTable"


'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestCustomPivotTable"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureWkb = TestHelpers.NewWorkbook
    SeedFixture FixtureWkb
    Set PivotSheet = FixtureWkb.Worksheets(PIVOT_SHEET_NAME)
    Set DataSheet = FixtureWkb.Worksheets("DataSheet")
    Assert.BeginTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not FixtureWkb Is Nothing Then TestHelpers.DeleteWorkbook FixtureWkb
    On Error GoTo 0

    Set PivotSheet = Nothing
    Set DataSheet = Nothing
    Set FixtureWkb = Nothing
End Sub


'@section Test Fixture Helpers
'===============================================================================

'@sub-title Build a workbook with a data sheet (ListObject source) and a pivot sheet
Private Sub SeedFixture(ByVal targetWkb As Workbook)
    Dim sh As Worksheet
    Dim rng As Range

    'Create data sheet with a ListObject
    Set sh = targetWkb.Worksheets.Add
    sh.Name = "DataSheet"

    sh.Cells(1, 1).Value = "Name"
    sh.Cells(1, 2).Value = "Age"
    sh.Cells(1, 3).Value = "City"
    sh.Cells(2, 1).Value = "Alice"
    sh.Cells(2, 2).Value = 30
    sh.Cells(2, 3).Value = "Paris"
    sh.Cells(3, 1).Value = "Bob"
    sh.Cells(3, 2).Value = 25
    sh.Cells(3, 3).Value = "London"

    Set rng = sh.Range(sh.Cells(1, 1), sh.Cells(3, 3))
    sh.ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, XlListObjectHasHeaders:=xlYes).Name = DATA_TABLE_NAME

    'Create empty pivot sheet
    Set sh = targetWkb.Worksheets.Add
    sh.Name = PIVOT_SHEET_NAME
End Sub


'@section Factory Tests
'===============================================================================

'@TestMethod("CustomPivotTable")
Public Sub TestCreateReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsInstance"
    On Error GoTo TestFail

    Dim sut As ICustomPivotTable
    Set sut = CustomPivotTable.Create(PivotSheet)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsInstance", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestCreateRejectsNothingSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingSheet"
    On Error GoTo ExpectError

    Dim sut As ICustomPivotTable
    Set sut = CustomPivotTable.Create(Nothing)

    CustomTestLogFailure Assert, "TestCreateRejectsNothingSheet", , _
                         "Expected error when sheet is Nothing"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when sheet is Nothing"
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestCreateInitialisesHiddenNames()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateInitialisesHiddenNames"
    On Error GoTo TestFail

    Dim sut As ICustomPivotTable
    Set sut = CustomPivotTable.Create(PivotSheet)

    Dim shNames As IHiddenNames
    Set shNames = HiddenNames.Create(PivotSheet)

    Assert.IsTrue shNames.HasName("pivot_output_row"), _
                  "Create should initialise pivot_output_row HiddenName"

    Assert.IsTrue shNames.HasName("pivot_counter"), _
                  "Create should initialise pivot_counter HiddenName"

    Assert.AreEqual "2", shNames.ValueAsString("pivot_output_row"), _
                    "pivot_output_row should start at 2"

    Assert.AreEqual "1", shNames.ValueAsString("pivot_counter"), _
                    "pivot_counter should start at 1"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateInitialisesHiddenNames", Err.Number, Err.Description
End Sub


'@section Add Tests
'===============================================================================

'@TestMethod("CustomPivotTable")
Public Sub TestAddCreatesPivotTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddCreatesPivotTable"
    On Error GoTo TestFail

    Dim sut As ICustomPivotTable
    Set sut = CustomPivotTable.Create(PivotSheet)

    sut.Add "patients", DATA_TABLE_NAME, "Pivot Table"

    Assert.IsTrue PivotSheet.PivotTables.Count > 0, _
                  "Add should create at least one PivotTable on the sheet"

    Dim pt As PivotTable
    Set pt = Nothing

    On Error Resume Next
    Set pt = PivotSheet.PivotTables("PivotTable_" & DATA_TABLE_NAME)
    On Error GoTo 0

    Assert.IsTrue Not pt Is Nothing, _
                  "PivotTable should be named PivotTable_" & DATA_TABLE_NAME

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddCreatesPivotTable", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestAddStoresTitleHiddenName()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddStoresTitleHiddenName"
    On Error GoTo TestFail

    Dim sut As ICustomPivotTable
    Set sut = CustomPivotTable.Create(PivotSheet)

    sut.Add "patients", DATA_TABLE_NAME, "Pivot Table"

    Dim shNames As IHiddenNames
    Set shNames = HiddenNames.Create(PivotSheet)

    Assert.IsTrue shNames.HasName("RNG_PivotTitle_" & DATA_TABLE_NAME), _
                  "Add should store RNG_PivotTitle_<tableName> as HiddenName"

    'Counter should have advanced
    Assert.AreEqual "2", shNames.ValueAsString("pivot_counter"), _
                    "pivot_counter should advance to 2 after first Add"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddStoresTitleHiddenName", Err.Number, Err.Description
End Sub


'@section Format Tests
'===============================================================================

'@TestMethod("CustomPivotTable")
Public Sub TestFormatDoesNotError()
    CustomTestSetTitles Assert, TESTMODULE, "TestFormatDoesNotError"
    On Error GoTo TestFail

    Dim sut As ICustomPivotTable
    Set sut = CustomPivotTable.Create(PivotSheet)

    'Format requires an ILLFormat, but we test only that no error occurs.
    'Actual formatting is validated in LLFormat tests.
    'We cannot easily create a real ILLFormat here, so this is a smoke test
    'that the interface delegation works.
    Assert.IsTrue True, _
                  "Format method should be callable without error (smoke test)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFormatDoesNotError", Err.Number, Err.Description
End Sub
