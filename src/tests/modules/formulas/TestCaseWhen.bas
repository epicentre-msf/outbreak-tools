Attribute VB_Name = "TestCaseWhen"
Attribute VB_Description = "Verifies the CaseWhen parser"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@Folder("CustomTests")
'@ModuleDescription("Verifies the CaseWhen parser")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

'@description
'Tests the CaseWhen class, which parses CASE_WHEN custom formulas into nested
'Excel IF statements. The suite covers valid formulas with and without default
'branches, category label extraction, and rejection of malformed input. Each
'test creates a fresh ICaseWhen instance via the CreateCaseWhen helper using
'module-level formula constants as fixtures.
'@depends CaseWhen, ICaseWhen, BetterArray, CustomTest, ICustomTest

Private Const VALID_FORMULA_DEFAULT As String = _
    "CASE_WHEN(A1=""Yes"", ""Choice is A"", B1>0, ""Choice is B"", ""Default Choice"")"
Private Const VALID_FORMULA_NO_DEFAULT As String = _
    "CASE_WHEN(A1=""Yes"", ""Choice is A"", OR(B1>0, C1<5), ""Choice is B"")"
Private Const INVALID_FORMULA As String = "IF(CASE_WHEN(yes, true)"

Private Assert As ICustomTest
Private casewhenObject As ICaseWhen

'@section Helpers
'===============================================================================

'@sub-title Instantiate a CaseWhen parser for the provided formula
Private Function CreateCaseWhen(ByVal formula As String) As ICaseWhen
    Set CreateCaseWhen = CaseWhen.Create(formula)
End Function

'@section Module Lifecycle
'===============================================================================

'@ModuleInitialize
'@sub-title Prepare the test output sheet and assertion engine
'@details
'Creates the shared output worksheet (if absent) and initialises the CustomTest
'assertion object for the entire module run.
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCaseWhen"
End Sub

'@ModuleCleanup
'@sub-title Print results and release module-level references
'@details
'Writes accumulated test results to the output sheet, then tears down the
'assertion object and the shared CaseWhen reference.
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
    Set casewhenObject = Nothing
End Sub

'@TestInitialize
'@sub-title Reset the CaseWhen instance before each test
'@details
'Clears the module-level casewhenObject so each test begins with a clean state.
Private Sub TestInitialize()
    Set casewhenObject = Nothing
End Sub

'@TestCleanup
'@sub-title Flush assertion state and release the CaseWhen instance
'@details
'Flushes any buffered assertion output and resets the casewhenObject reference
'after each test method completes.
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set casewhenObject = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("CaseWhen")
'@sub-title Verify a valid formula with default branch parses to nested IF
'@details
'Arranges a CASE_WHEN formula containing two condition/result pairs plus a
'default branch. Acts by creating the parser and reading the parsedFormula
'property. Asserts that the formula is marked valid and that the output matches
'the expected nested IF(condition, result, IF(...)) structure with the default
'value as the innermost else.
Public Sub TestValidFormulaParsesToNestedIf()
    CustomTestSetTitles Assert, "CaseWhen", "TestValidFormulaParsesToNestedIf"
    On Error GoTo Fail

    Dim expected As String

    Set casewhenObject = CreateCaseWhen(VALID_FORMULA_DEFAULT)

    Assert.IsTrue casewhenObject.Valid, "CASE_WHEN formula should be recognised as valid"

    expected = "IF(A1=""Yes"", ""Choice is A"", IF(B1>0, ""Choice is B"", ""Default Choice""))"
    Assert.AreEqual expected, casewhenObject.parsedFormula, "Parsed formula does not match expected nested IF"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValidFormulaParsesToNestedIf", Err.Number, Err.Description
End Sub

'@TestMethod("CaseWhen")
'@sub-title Verify category extraction returns all branch labels
'@details
'Arranges a valid CASE_WHEN formula with two condition branches and a default.
'Acts by reading the Categories property which returns a BetterArray of labels.
'Asserts that exactly three categories are extracted in order: the two branch
'result strings and the default value.
Public Sub TestCategoriesExtractLabels()
    CustomTestSetTitles Assert, "CaseWhen", "TestCategoriesExtractLabels"
    On Error GoTo Fail

    Dim categories As BetterArray

    Set casewhenObject = CreateCaseWhen(VALID_FORMULA_DEFAULT)
    Set categories = casewhenObject.Categories

    Assert.IsTrue (categories.Length = 3), "Expected three categories including default. Lenght: " & categories.Length
    Assert.AreEqual "Choice is A", categories.Item(1), "First category should match first branch"
    Assert.AreEqual "Choice is B", categories.Item(2), "Second category should match second branch"
    Assert.AreEqual "Default Choice", categories.Item(3), "Default branch should supply final category"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCategoriesExtractLabels", Err.Number, Err.Description
End Sub

'@TestMethod("CaseWhen")
'@sub-title Verify missing default produces an empty-string else branch
'@details
'Arranges a CASE_WHEN formula with two condition/result pairs but no trailing
'default argument. Acts by parsing the formula and reading the output. Asserts
'that the innermost else of the nested IF is an empty string literal ("").
Public Sub TestMissingDefaultProducesEmptyString()
    CustomTestSetTitles Assert, "CaseWhen", "TestMissingDefaultProducesEmptyString"
    On Error GoTo Fail

    Dim expected As String

    Set casewhenObject = CreateCaseWhen(VALID_FORMULA_NO_DEFAULT)

    expected = "IF(A1=""Yes"", ""Choice is A"", IF(OR(B1>0, C1<5), ""Choice is B"", """"))"
    Assert.AreEqual expected, casewhenObject.parsedFormula, "Missing default should produce empty string literal"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMissingDefaultProducesEmptyString", Err.Number, Err.Description
End Sub

'@TestMethod("CaseWhen")
'@sub-title Verify an invalid formula is rejected with empty outputs
'@details
'Arranges a malformed formula that wraps CASE_WHEN inside IF with unbalanced
'parentheses. Acts by creating the parser and querying Valid, parsedFormula,
'and Categories. Asserts that the formula is marked invalid, the parsed output
'is an empty string, and the category collection has zero length.
Public Sub TestInvalidFormulaRejected()
    CustomTestSetTitles Assert, "CaseWhen", "TestInvalidFormulaRejected"
    On Error GoTo Fail

    Dim categories As BetterArray

    Set casewhenObject = CreateCaseWhen(INVALID_FORMULA)

    Assert.IsFalse casewhenObject.Valid, "Invalid CASE_WHEN wrapper should fail validation"
    Assert.AreEqual vbNullString, casewhenObject.parsedFormula, "Parsed formula should be empty when invalid"

    Set categories = casewhenObject.Categories
    Assert.IsTrue (categories.Length = 0), "Invalid formulas should not yield categories"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInvalidFormulaRejected", Err.Number, Err.Description
End Sub
