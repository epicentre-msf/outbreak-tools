Attribute VB_Name = "TestChoiceFormula"
Attribute VB_Description = "Verifies the ChoiceFormula parser"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@Folder("CustomTests")
'@ModuleDescription("Verifies the ChoiceFormula parser")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

'@description
'Tests the ChoiceFormula class, which parses CHOICE_FORMULA custom expressions
'into nested Excel IF statements via CaseWhen delegation. The suite covers valid
'formulas with and without default branches, category extraction, choice name
'parsing from the first argument, nested comma expressions (e.g. OR), and
'rejection of invalid input including missing condition/result pairs and wrong
'token types. Each test creates a fresh IChoiceFormula instance through the
'CreateChoiceFormula helper using module-level formula constants as fixtures.
'@depends ChoiceFormula, IChoiceFormula, BetterArray, CustomTest, ICustomTest

Private Const VALID_FORMULA_WITH_DEFAULT As String = _
    "CHOICE_FORMULA(list_multiple, A1=""Yes"", ""Choice is A"", B1>0, ""Choice is B"", ""Default Choice"")"
Private Const VALID_FORMULA_NO_DEFAULT As String = _
    "CHOICE_FORMULA(list_multiple, OR(B1>0, C1<5), ""Either positive"", A1=""Yes"", ""Choice is A"")"
Private Const INVALID_FORMULA_NO_PAIR As String = "CHOICE_FORMULA(list_multiple)"
Private Const INVALID_FORMULA_WRONG_TOKEN As String = _
    "CASE_WHEN(A1=""Yes"", ""Choice is A"", B1>0, ""Choice is B"")"

Private Assert As ICustomTest
Private choiceObj As IChoiceFormula

'@section Helpers
'===============================================================================

'@sub-title Instantiate a ChoiceFormula parser for the provided expression
Private Function CreateChoiceFormula(ByVal formula As String) As IChoiceFormula
    Set CreateChoiceFormula = ChoiceFormula.Create(formula)
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
    Assert.SetModuleName "TestChoiceFormula"
End Sub

'@ModuleCleanup
'@sub-title Print results and release module-level references
'@details
'Writes accumulated test results to the output sheet, then tears down the
'assertion object and the shared ChoiceFormula reference.
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
    Set choiceObj = Nothing
End Sub

'@TestInitialize
'@sub-title Reset the ChoiceFormula instance before each test
'@details
'Clears the module-level choiceObj so each test begins with a clean state.
Private Sub TestInitialize()
    Set choiceObj = Nothing
End Sub

'@TestCleanup
'@sub-title Flush assertion state and release the ChoiceFormula instance
'@details
'Flushes any buffered assertion output and resets the choiceObj reference
'after each test method completes.
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set choiceObj = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("ChoiceFormula")
'@sub-title Verify a valid formula with default branch parses to nested IF
'@details
'Arranges a CHOICE_FORMULA expression containing a choice name, two
'condition/result pairs, and a default branch. Acts by creating the parser
'and reading the parsedFormula property. Asserts that the formula is marked
'valid and that the output matches the expected nested IF structure with the
'default value as the innermost else.
Public Sub TestValidFormulaParsesToNestedIf()
    CustomTestSetTitles Assert, "ChoiceFormula", "TestValidFormulaParsesToNestedIf"
    On Error GoTo Fail

    Dim expected As String

    Set choiceObj = CreateChoiceFormula(VALID_FORMULA_WITH_DEFAULT)

    Assert.IsTrue choiceObj.Valid, "CHOICE_FORMULA should be recognised as valid"

    expected = "IF(A1=""Yes"", ""Choice is A"", IF(B1>0, ""Choice is B"", ""Default Choice""))"
    Assert.AreEqual expected, choiceObj.parsedFormula, "Parsed formula does not match expected nested IF"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValidFormulaParsesToNestedIf", Err.Number, Err.Description
End Sub

'@TestMethod("ChoiceFormula")
'@sub-title Verify category extraction returns all branch labels including default
'@details
'Arranges a valid CHOICE_FORMULA with two condition branches and a default.
'Acts by reading the Categories property which returns a BetterArray of labels.
'Asserts that exactly three categories are extracted in order: the two branch
'result strings and the default value.
Public Sub TestCategoriesReflectResults()
    CustomTestSetTitles Assert, "ChoiceFormula", "TestCategoriesReflectResults"
    On Error GoTo Fail

    Dim categories As BetterArray

    Set choiceObj = CreateChoiceFormula(VALID_FORMULA_WITH_DEFAULT)
    Set categories = choiceObj.Categories

    Assert.AreEqual 3, categories.Length, "Expected three categories including default branch"
    Assert.AreEqual "Choice is A", categories.Item(1), "First category should match first branch"
    Assert.AreEqual "Choice is B", categories.Item(2), "Second category should match second branch"
    Assert.AreEqual "Default Choice", categories.Item(3), "Default branch should supply final category"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCategoriesReflectResults", Err.Number, Err.Description
End Sub

'@TestMethod("ChoiceFormula")
'@sub-title Verify choice name is extracted and trimmed from the first argument
'@details
'Arranges a CHOICE_FORMULA expression with extra whitespace around the token
'and the first argument. Acts by creating the parser and reading the choiceName
'property. Asserts that the formula remains valid despite whitespace and that
'the choice name is correctly trimmed to "list_multiple".
Public Sub TestChoiceNameExtractedFromFirstArgument()
    CustomTestSetTitles Assert, "ChoiceFormula", "TestChoiceNameExtractedFromFirstArgument"
    On Error GoTo Fail

    Set choiceObj = CreateChoiceFormula("  CHOICE_FORMULA ( list_multiple , A1=""Yes"", ""Choice is A"" )")

    Assert.IsTrue choiceObj.Valid, "Choice formula with additional whitespace should remain valid"
    Assert.AreEqual "list_multiple", choiceObj.choiceName, "Choice name should be trimmed from the first argument"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestChoiceNameExtractedFromFirstArgument", Err.Number, Err.Description
End Sub

'@TestMethod("ChoiceFormula")
'@sub-title Verify nested comma expressions like OR are preserved correctly
'@details
'Arranges a CHOICE_FORMULA whose first condition is an OR expression containing
'commas inside parentheses. Acts by parsing the formula and reading the output.
'Asserts that the formula is valid and that the nested OR(...) sub-expression
'is preserved intact within the generated nested IF structure, with an empty
'string default because no default argument is provided.
Public Sub TestNestedCommaExpressionsHandled()
    CustomTestSetTitles Assert, "ChoiceFormula", "TestNestedCommaExpressionsHandled"
    On Error GoTo Fail

    Dim expected As String

    Set choiceObj = CreateChoiceFormula(VALID_FORMULA_NO_DEFAULT)

    Assert.IsTrue choiceObj.Valid, "Formula containing nested OR should be valid"

    expected = "IF(OR(B1>0, C1<5), ""Either positive"", IF(A1=""Yes"", ""Choice is A"", """"))"
    Assert.AreEqual expected, choiceObj.parsedFormula, "Nested OR expression should be preserved"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestNestedCommaExpressionsHandled", Err.Number, Err.Description
End Sub

'@TestMethod("ChoiceFormula")
'@sub-title Verify invalid formulas are rejected with empty outputs
'@details
'Tests two invalid scenarios. First, a CHOICE_FORMULA with only a choice name
'but no condition/result pairs: asserts that Valid is False, parsedFormula is
'empty, and Categories has zero length. Second, a CASE_WHEN formula (wrong
'token type) passed to the ChoiceFormula parser: asserts that it is also
'rejected as invalid.
Public Sub TestInvalidFormulaRejected()
    CustomTestSetTitles Assert, "ChoiceFormula", "TestInvalidFormulaRejected"
    On Error GoTo Fail

    Dim categories As BetterArray

    Set choiceObj = CreateChoiceFormula(INVALID_FORMULA_NO_PAIR)

    Assert.IsFalse choiceObj.Valid, "Formula without condition/result pair should be invalid"
    Assert.AreEqual vbNullString, choiceObj.parsedFormula, "Invalid specification should not produce parsed formula"

    Set categories = choiceObj.Categories
    Assert.AreEqual 0, categories.Length, "Invalid formulas should not yield categories"

    Set choiceObj = CreateChoiceFormula(INVALID_FORMULA_WRONG_TOKEN)
    Assert.IsFalse choiceObj.Valid, "Non-choice formula should be rejected"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInvalidFormulaRejected", Err.Number, Err.Description
End Sub
