Attribute VB_Name = "TestLLSpatial"
Attribute VB_Description = "Tests for LLSpatial class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLSpatial class")

'@description
'Validates the LLSpatial class, which manages the "spatial_tables__"
'worksheet containing ListObjects for spatial analysis of geo variables.
'Tests focus on factory validation and lightweight property behaviour since
'full integration tests with spatial ListObjects and HList data require a
'complete linelist workbook. The fixture builds a minimal spatial worksheet
'with the required "listofgeovars" ListObject and a pasting named range,
'then tears it down in ModuleCleanup. Tests verify: factory rejects Nothing;
'factory rejects sheets with wrong name; factory succeeds with correctly
'named sheet containing listofgeovars; Exists returns True for known
'variables and False for unknown ones; TopGeoValue and TopHFValue return
'empty strings when the spatial ListObjects do not exist.
'@depends LLSpatial, ILLSpatial, CustomTest, TestHelpers

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SPATIAL_SHEET As String = "spatial_tables__"
Private Const SPATIAL_WRONG As String = "WrongSheetName"

Private Assert As ICustomTest

'@section Module lifecycle
'===============================================================================

'@sub-title Initialise the test module before any tests run.
'@details
'Suppresses screen updates via BusyApp, ensures the test output sheet
'exists, creates the CustomTest assertion object targeting that sheet,
'and sets the module name for result grouping.
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLSpatial"
End Sub

'@sub-title Tear down the module after all tests complete.
'@details
'Prints accumulated test results to the output sheet, restores the
'application state via RestoreApp, releases the assertion object, and
'deletes all temporary worksheets created during the test run.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
    DeleteWorksheets SPATIAL_SHEET, SPATIAL_WRONG
End Sub

'@sub-title Reset state before each individual test.
'@details
'Suppresses screen updates so worksheet operations during each test do
'not trigger flickering or event cascades.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Clean up after each individual test.
'@details
'Flushes any pending assertion results to the output sheet so each test's
'outcome is recorded before the next test begins.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Helpers
'===============================================================================

'@sub-title Build a minimal spatial fixture with the required "listofgeovars" ListObject.
'@details
'Creates a hidden worksheet named "spatial_tables__" containing a
'"listofgeovars" ListObject with a "varname" header column. When addVars
'is True (default), populates two data rows ("cases_sp1", "deaths_sp1")
'so Exists lookups can find matching entries. Also creates a
'RNG_PastingCol named range for the pasting scratch column.
'@param addVars Optional Boolean. True to populate sample variable rows. Defaults to True.
'@return Worksheet. The prepared spatial fixture sheet.
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

'@sub-title Verify Create returns Nothing when the worksheet argument is Nothing.
'@details
'Acts by calling LLSpatial.Create with Nothing under On Error Resume Next.
'Asserts that the result is Nothing, confirming the guard clause rejects
'invalid input without raising an unhandled error.
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

'@sub-title Verify Create returns Nothing when the worksheet has the wrong name.
'@details
'Arranges a hidden worksheet named "WrongSheetName" which does not match
'the required "spatial_tables__" name. Acts by calling LLSpatial.Create
'with that sheet under On Error Resume Next. Asserts that the result is
'Nothing, confirming the factory validates the sheet name before returning
'an instance.
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

'@sub-title Verify Create succeeds with a correctly named sheet containing listofgeovars.
'@details
'Arranges a spatial fixture via BuildSpatialFixture with the correct
'sheet name and the required "listofgeovars" ListObject. Acts by calling
'LLSpatial.Create with that sheet. Asserts that the result is not Nothing,
'confirming the factory accepts a well-formed spatial worksheet.
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

'@sub-title Verify Exists returns True for a variable present in listofgeovars.
'@details
'Arranges a spatial fixture with "cases_sp1" and "deaths_sp1" in the
'listofgeovars table. Acts by creating an LLSpatial instance and calling
'Exists("cases"). Asserts that the result is True, confirming partial
'name matching finds the variable in the spatial variables list.
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

'@sub-title Verify Exists returns False for a variable not in listofgeovars.
'@details
'Arranges a spatial fixture with known variables. Acts by creating an
'LLSpatial instance and calling Exists("nonexistent_var"). Asserts that
'the result is False, confirming the method correctly reports absence
'when no partial match is found.
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

'@sub-title Verify TopGeoValue returns empty when the spatial ListObject does not exist.
'@details
'Arranges a spatial fixture with listofgeovars but no admin-level spatial
'ListObjects. Acts by creating an LLSpatial instance and calling
'TopGeoValue("adm1", 1, "cases", "sp1"). Asserts that the result is
'vbNullString, confirming the property handles missing tables gracefully
'by returning an empty string rather than raising an error.
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

'@sub-title Verify TopHFValue returns empty when the spatial ListObject does not exist.
'@details
'Arranges a spatial fixture with listofgeovars but no health facility
'spatial ListObjects. Acts by creating an LLSpatial instance and calling
'TopHFValue(1, "cases", "sp1"). Asserts that the result is vbNullString,
'confirming the property handles missing tables gracefully by returning
'an empty string rather than raising an error.
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
