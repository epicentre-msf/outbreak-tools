Attribute VB_Name = "TestAnalysisOutput"
Attribute VB_Description = "Tests for AnalysisOutput class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for AnalysisOutput class")

'@description
'Validates the AnalysisOutput factory guard clauses. AnalysisOutput orchestrates
'the full analysis output pipeline (cross-tables, formulas, graphs, metadata
'tracking) for all four analysis scopes, so its Create method requires both a
'valid specs worksheet and a valid linelist facade. These tests verify that
'Create returns Nothing when either argument is Nothing. Full integration
'tests are not feasible at the unit level because the class requires a complete
'linelist workbook with analysis setup ListObjects, output worksheets,
'translation tables, formula data, and a dictionary.
'@depends AnalysisOutput, IAnalysisOutput, CustomTest, TestHelpers

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

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
    Assert.SetModuleName "TestAnalysisOutput"
End Sub

'@sub-title Print results and tear down shared state
'@details
'Prints accumulated test results to the output sheet, restores the Excel
'application state, and releases the assertion object. Called once after all
'tests have run.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
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

'@sub-title Verify Create rejects a Nothing specs worksheet
'@details
'Arranges by passing Nothing for both the specs sheet and the linelist
'arguments under On Error Resume Next. Acts by calling AnalysisOutput.Create.
'Asserts that the returned IAnalysisOutput reference is Nothing, confirming
'the factory guard prevents instantiation when the specs worksheet is missing.
'@TestMethod("AnalysisOutput")
Public Sub TestCreateRejectsNothingSpecSheet()
    CustomTestSetTitles Assert, "AnalysisOutput", "TestCreateRejectsNothingSpecSheet"
    On Error GoTo TestFail

    On Error Resume Next
    Dim ao As IAnalysisOutput
    Set ao = AnalysisOutput.Create(Nothing, Nothing)
    On Error GoTo 0

    Assert.IsTrue (ao Is Nothing), _
                  "Create with Nothing specs sheet should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingSpecSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Create rejects a Nothing linelist when specs sheet is valid
'@details
'Arranges by creating a temporary hidden worksheet to serve as a valid specs
'sheet, then passing Nothing for the linelist argument under On Error Resume
'Next. Acts by calling AnalysisOutput.Create with the valid sheet and Nothing.
'Asserts that the returned IAnalysisOutput reference is Nothing, confirming
'the factory guard catches a missing linelist even when the specs sheet is
'provided. Cleans up the temporary worksheet after the assertion.
'@TestMethod("AnalysisOutput")
Public Sub TestCreateRejectsNothingLinelist()
    CustomTestSetTitles Assert, "AnalysisOutput", "TestCreateRejectsNothingLinelist"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet("AOTestSheet", clearSheet:=True, visibility:=xlSheetHidden)

    On Error Resume Next
    Dim ao As IAnalysisOutput
    Set ao = AnalysisOutput.Create(sh, Nothing)
    On Error GoTo 0

    Assert.IsTrue (ao Is Nothing), _
                  "Create with Nothing linelist should fail"

    DeleteWorksheet "AOTestSheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingLinelist", Err.Number, Err.Description
End Sub
