Attribute VB_Name = "TestChoiceFormula"
Attribute VB_Description = "Verifies the ChoiceFormula parser"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

'@Folder("CustomTests")
'@Folder("Tests")
'@ModuleDescription("Verifies the ChoiceFormula parser")

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

'Instantiate a ChoiceFormula parser for the provided expression.
Private Function CreateChoiceFormula(ByVal formula As String) As IChoiceFormula
    Set CreateChoiceFormula = ChoiceFormula.Create(formula)
End Function

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestChoiceFormula"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
    Set choiceObj = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set choiceObj = Nothing
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set choiceObj = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("ChoiceFormula")
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
