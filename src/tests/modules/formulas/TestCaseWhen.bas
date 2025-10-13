Attribute VB_Name = "TestCaseWhen"
Attribute VB_Description = "Verifies the CaseWhen parser"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

'@Folder("CustomTests")

'@ModuleDescription("Verifies the CaseWhen parser")

Private Const VALID_FORMULA_DEFAULT As String = _
    "CASE_WHEN(A1=""Yes"", ""Choice is A"", B1>0, ""Choice is B"", ""Default Choice"")"
Private Const VALID_FORMULA_NO_DEFAULT As String = _
    "CASE_WHEN(A1=""Yes"", ""Choice is A"", OR(B1>0, C1<5), ""Choice is B"")"
Private Const INVALID_FORMULA As String = "IF(CASE_WHEN(yes, true)"

Private Assert As ICustomTest
Private casewhenObject As ICaseWhen

'@section Helpers
'===============================================================================

'Instantiate a CaseWhen parser for the provided formula.
Private Function CreateCaseWhen(ByVal formula As String) As ICaseWhen
    Set CreateCaseWhen = CaseWhen.Create(formula)
End Function

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCaseWhen"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
    Set casewhenObject = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set casewhenObject = Nothing
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set casewhenObject = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("CaseWhen")
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
