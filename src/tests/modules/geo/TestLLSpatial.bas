Attribute VB_Name = "TestLLSpatial"
Attribute VB_Description = "Tests for LLSpatial class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLSpatial class")

' LLSpatial tests focus on factory validation and lightweight property behavior.
' Full integration tests with spatial ListObjects and HList data require a
' complete linelist workbook — making them unsuitable for unit tests.
' These tests verify:
' - Factory rejects Nothing
' - Factory rejects sheets with wrong name
' - Factory succeeds with correctly named sheet + listofgeovars
' - Exists returns True/False based on listofgeovars content
' - TopGeoValue/TopHFValue return empty for missing ListObjects

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SPATIAL_SHEET As String = "spatial_tables__"
Private Const SPATIAL_WRONG As String = "WrongSheetName"

Private Assert As ICustomTest

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLSpatial"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
    DeleteWorksheets SPATIAL_SHEET, SPATIAL_WRONG
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

'@section Helpers
'===============================================================================

' Build a minimal spatial fixture with the required "listofgeovars" ListObject.
Private Function BuildSpatialFixture(Optional ByVal addVars As Boolean = True) As Worksheet
    Dim sh As Worksheet
    Dim rng As Range

    Set sh = EnsureWorksheet(SPATIAL_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    'Create the listofgeovars table
    sh.Cells(1, 1).Value = "varname"

    If addVars Then
        sh.Cells(2, 1).Value = "cases_sp1"
        sh.Cells(3, 1).Value = "deaths_sp1"
        Set rng = sh.Range(sh.Cells(1, 1), sh.Cells(3, 1))
    Else
        Set rng = sh.Range(sh.Cells(1, 1), sh.Cells(2, 1))
    End If

    sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "listofgeovars"

    'Add a scratch column for pasting
    sh.Cells(1, 5).Value = "scratch"
    sh.Cells(1, 5).Name = "RNG_PastingCol"

    Set BuildSpatialFixture = sh
End Function

'@section Factory validation tests
'===============================================================================

'@TestMethod("LLSpatial")
Public Sub TestCreateRejectsNothing()
    CustomTestSetTitles Assert, "LLSpatial", "TestCreateRejectsNothing"
    On Error GoTo TestFail

    On Error Resume Next
    Dim sp As ILLSpatial
    Set sp = LLSpatial.Create(Nothing)
    On Error GoTo 0

    Assert.IsTrue (sp Is Nothing), _
                  "Create with Nothing sheet should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothing", Err.Number, Err.Description
End Sub

'@TestMethod("LLSpatial")
Public Sub TestCreateRejectsWrongSheetName()
    CustomTestSetTitles Assert, "LLSpatial", "TestCreateRejectsWrongSheetName"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet(SPATIAL_WRONG, clearSheet:=True, visibility:=xlSheetHidden)

    On Error Resume Next
    Dim sp As ILLSpatial
    Set sp = LLSpatial.Create(sh)
    On Error GoTo 0

    Assert.IsTrue (sp Is Nothing), _
                  "Create with wrong sheet name should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsWrongSheetName", Err.Number, Err.Description
End Sub

'@TestMethod("LLSpatial")
Public Sub TestCreateSucceedsWithCorrectSheet()
    CustomTestSetTitles Assert, "LLSpatial", "TestCreateSucceedsWithCorrectSheet"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture()

    Dim sp As ILLSpatial
    Set sp = LLSpatial.Create(sh)

    Assert.IsNotNothing sp, _
                        "Create with correctly named sheet should succeed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateSucceedsWithCorrectSheet", Err.Number, Err.Description
End Sub

'@section Exists tests
'===============================================================================

'@TestMethod("LLSpatial")
Public Sub TestExistsReturnsTrueForKnownVar()
    CustomTestSetTitles Assert, "LLSpatial", "TestExistsReturnsTrueForKnownVar"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture(addVars:=True)

    Dim sp As ILLSpatial
    Set sp = LLSpatial.Create(sh)

    Assert.IsTrue sp.Exists("cases"), _
                  "Exists should return True for a variable matching 'cases'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestExistsReturnsTrueForKnownVar", Err.Number, Err.Description
End Sub

'@TestMethod("LLSpatial")
Public Sub TestExistsReturnsFalseForUnknownVar()
    CustomTestSetTitles Assert, "LLSpatial", "TestExistsReturnsFalseForUnknownVar"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture(addVars:=True)

    Dim sp As ILLSpatial
    Set sp = LLSpatial.Create(sh)

    Assert.IsFalse sp.Exists("nonexistent_var"), _
                   "Exists should return False for an unknown variable"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestExistsReturnsFalseForUnknownVar", Err.Number, Err.Description
End Sub

'@section TopGeoValue / TopHFValue tests
'===============================================================================

'@TestMethod("LLSpatial")
Public Sub TestTopGeoValueReturnsEmptyForMissingTable()
    CustomTestSetTitles Assert, "LLSpatial", "TestTopGeoValueReturnsEmptyForMissingTable"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture()

    Dim sp As ILLSpatial
    Set sp = LLSpatial.Create(sh)

    Dim result As String
    result = sp.TopGeoValue("adm1", 1, "cases", "sp1")

    Assert.AreEqual vbNullString, result, _
                    "TopGeoValue should return empty when spatial table does not exist"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTopGeoValueReturnsEmptyForMissingTable", Err.Number, Err.Description
End Sub

'@TestMethod("LLSpatial")
Public Sub TestTopHFValueReturnsEmptyForMissingTable()
    CustomTestSetTitles Assert, "LLSpatial", "TestTopHFValueReturnsEmptyForMissingTable"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture()

    Dim sp As ILLSpatial
    Set sp = LLSpatial.Create(sh)

    Dim result As String
    result = sp.TopHFValue(1, "cases", "sp1")

    Assert.AreEqual vbNullString, result, _
                    "TopHFValue should return empty when spatial table does not exist"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTopHFValueReturnsEmptyForMissingTable", Err.Number, Err.Description
End Sub
