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
'branches, category label extraction, quoting rules, and rejection of malformed
'input. Each test creates a fresh CaseWhen instance via the CreateCaseWhen
'helper using module-level formula constants as fixtures.
'@depends CaseWhen, BetterArray, CustomTest, TestHelpersLite

Private Const VALID_FORMULA_DEFAULT As String = _
    "CASE_WHEN(A1=""Yes"", ""Choice is A"", B1>0, ""Choice is B"", ""Default Choice"")"
Private Const VALID_FORMULA_NO_DEFAULT As String = _
    "CASE_WHEN(A1=""Yes"", ""Choice is A"", OR(B1>0, C1<5), ""Choice is B"")"
Private Const INVALID_FORMULA As String = "IF(CASE_WHEN(yes, true)"
Private Const BARE_HEADER_FORMULA As String = "CASE_WHEN("

'A result written without quotes, next to results written with them.
Private Const MIXED_QUOTING_FORMULA As String = _
    "CASE_WHEN(A1=""Yes"", ""Choice is A"", B1>0, C1, ""Default Choice"")"

'A quoted result that carries a doubled quote inside it.
Private Const INNER_QUOTE_FORMULA As String = _
    "CASE_WHEN(A1=1, ""He said """"hi"""""", ""Other"")"

Private Assert As CustomTest
Private casewhenObject As CaseWhen

'@section Helpers
'===============================================================================

'@sub-title Instantiate a CaseWhen parser for the provided formula
Private Function CreateCaseWhen(ByVal formula As String) As CaseWhen
    Set CreateCaseWhen = CaseWhen.Create(formula)
End Function

'@section Module Lifecycle
'===============================================================================

'@sub-title Prepare the test output sheet and assertion engine
'@details
'Creates the shared output worksheet (if absent) and initialises the CustomTest
'assertion object for the entire module run.
'Public, so the headless runner reaches it: the runner calls every lifecycle
'hook through Application.Run, which cannot see a Private Sub. The handler
'matters as much. An error escaping a lifecycle hook aborts the WHOLE module,
'so a fixture problem here would drop every test and the run would report no
'failures, because nothing ran.
'@ModuleInitialize
Public Sub ModuleInitialize()
    On Error GoTo Fail
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCaseWhen"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "ModuleInitialize", Err.Number, Err.Description
End Sub

'@sub-title Print results and release module-level references
'@details
'Writes accumulated test results to the output sheet, then tears down the
'assertion object and the shared CaseWhen reference.
'Every step before RestoreApp is wrapped, because RestoreApp has to run
'whatever happened above: the hooks here call BusyApp, which puts Excel in
'manual calculation, and the next module in the run would inherit that.
'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0

    RestoreApp

    Set Assert = Nothing
    Set casewhenObject = Nothing
End Sub

'@sub-title Reset the CaseWhen instance before each test
'@details
'Clears the module-level casewhenObject so each test begins with a clean state.
'@TestInitialize
Public Sub TestInitialize()
    On Error GoTo Fail
    BusyApp
    Set casewhenObject = Nothing
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInitialize", Err.Number, Err.Description
End Sub

'@sub-title Flush assertion state and release the CaseWhen instance
'@details
'Flushes any buffered assertion output and resets the casewhenObject reference
'after each test method completes.
'@TestCleanup
Public Sub TestCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.Flush
        End If
    On Error GoTo 0

    Set casewhenObject = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("CaseWhen")
'@sub-title Verify a valid formula with default branch parses to nested IF
'@details
'Arranges a CASE_WHEN formula containing two condition/result pairs plus a
'default branch. Acts by creating the parser and reading the ParsedFormula
'property. Asserts that the formula is marked valid and that the output matches
'the expected nested IF(condition, result, IF(...)) structure with the default
'value as the innermost else.
Public Sub TestValidCaseWhenParsesToNestedIf()
    CustomTestSetTitles Assert, "CaseWhen", "TestValidCaseWhenParsesToNestedIf"
    On Error GoTo Fail

    Dim expected As String

    Set casewhenObject = CreateCaseWhen(VALID_FORMULA_DEFAULT)

    Assert.IsTrue casewhenObject.Valid, "CASE_WHEN formula should be recognised as valid"

    expected = "IF(A1=""Yes"", ""Choice is A"", IF(B1>0, ""Choice is B"", ""Default Choice""))"
    Assert.AreEqual expected, casewhenObject.ParsedFormula, "Parsed formula does not match expected nested IF"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValidCaseWhenParsesToNestedIf", Err.Number, Err.Description
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
'@sub-title Verify every category comes back as plain text
'@details
'Arranges a CASE_WHEN whose middle result carries no quotes while the others
'do. Acts by reading Categories. Asserts that all three labels come back with
'no quote character around them. SetupErrors compares these labels against the
'choice sheet, which holds plain text, and EventSetup draws them as labels, so
'a label wrapped in quotes matches nothing and is drawn with the quotes showing.
Public Sub TestCategoriesReturnPlainLabels()
    CustomTestSetTitles Assert, "CaseWhen", "TestCategoriesReturnPlainLabels"
    On Error GoTo Fail

    Dim categories As BetterArray

    Set casewhenObject = CreateCaseWhen(MIXED_QUOTING_FORMULA)
    Set categories = casewhenObject.Categories

    Assert.IsTrue (categories.Length = 3), "Expected three categories. Length: " & categories.Length
    Assert.AreEqual "Choice is A", categories.Item(1), "A quoted label loses its quotes"
    Assert.AreEqual "C1", categories.Item(2), "An unquoted label is returned as written"
    Assert.AreEqual "Default Choice", categories.Item(3), "The default label loses its quotes"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCategoriesReturnPlainLabels", Err.Number, Err.Description
End Sub

'@TestMethod("CaseWhen")
'@sub-title Verify a doubled quote inside a label survives
'@details
'Arranges a CASE_WHEN whose first result is a quoted string holding a doubled
'quote. Acts by reading Categories. Asserts that only the outer pair is removed
'and the doubled inner quote becomes one quote.
Public Sub TestCategoriesKeepInnerQuotes()
    CustomTestSetTitles Assert, "CaseWhen", "TestCategoriesKeepInnerQuotes"
    On Error GoTo Fail

    Dim categories As BetterArray

    Set casewhenObject = CreateCaseWhen(INNER_QUOTE_FORMULA)
    Set categories = casewhenObject.Categories

    Assert.IsTrue (categories.Length = 2), "Expected two categories. Length: " & categories.Length
    Assert.AreEqual "He said ""hi""", categories.Item(1), "Inner quotes should survive as one quote each"
    Assert.AreEqual "Other", categories.Item(2), "Default branch should supply final category"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCategoriesKeepInnerQuotes", Err.Number, Err.Description
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
    Assert.AreEqual expected, casewhenObject.ParsedFormula, "Missing default should produce empty string literal"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMissingDefaultProducesEmptyString", Err.Number, Err.Description
End Sub

'@TestMethod("CaseWhen")
'@sub-title Verify an invalid formula is rejected with empty outputs
'@details
'Arranges a malformed formula that wraps CASE_WHEN inside IF with unbalanced
'parentheses. Acts by creating the parser and querying Valid, ParsedFormula,
'and Categories. Asserts that the formula is marked invalid, the parsed output
'is an empty string, and the category collection has zero length.
Public Sub TestInvalidCaseWhenRejected()
    CustomTestSetTitles Assert, "CaseWhen", "TestInvalidCaseWhenRejected"
    On Error GoTo Fail

    Dim categories As BetterArray

    Set casewhenObject = CreateCaseWhen(INVALID_FORMULA)

    Assert.IsFalse casewhenObject.Valid, "Invalid CASE_WHEN wrapper should fail validation"
    Assert.AreEqual vbNullString, casewhenObject.ParsedFormula, "Parsed formula should be empty when invalid"

    Set categories = casewhenObject.Categories
    Assert.IsTrue (categories.Length = 0), "Invalid formulas should not yield categories"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInvalidCaseWhenRejected", Err.Number, Err.Description
End Sub

'@TestMethod("CaseWhen")
'@sub-title Verify a header with no body is rejected and says why
'@details
'Arranges the bare token "CASE_WHEN(" with nothing after it. Acts by creating
'the parser and querying Valid, ParsedFormula and FailureReason. Asserts that
'the formula is rejected and that a reason is available. The old class accepted
'this shape, handed back an empty string, and the caller reported it as "the
'formula is empty" for an expression the user did write.
Public Sub TestBareHeaderIsRejectedWithReason()
    CustomTestSetTitles Assert, "CaseWhen", "TestBareHeaderIsRejectedWithReason"
    On Error GoTo Fail

    Set casewhenObject = CreateCaseWhen(BARE_HEADER_FORMULA)

    Assert.IsFalse casewhenObject.Valid, "A CASE_WHEN header with no body should fail validation"
    Assert.AreEqual vbNullString, casewhenObject.ParsedFormula, "Parsed formula should be empty when invalid"
    Assert.IsTrue (LenB(casewhenObject.FailureReason) > 0), "A rejected formula should carry a reason"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBareHeaderIsRejectedWithReason", Err.Number, Err.Description
End Sub
