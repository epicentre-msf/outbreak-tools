Attribute VB_Name = "TestValueOfFormula"
Attribute VB_Description = "Verifies VALUE_OF formula conversion"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Verifies VALUE_OF formula conversion")

'@description
'Tests the ValueOfFormula class, which parses VALUE_OF custom expressions and
'resolves them into workbook-ready formulas containing sheet names and column
'indices. Covers the happy-path conversion of a valid three-argument VALUE_OF
'expression, rejection of cross-sheet arguments where the lookup and value
'variables reside on different worksheets, rejection of a token that only
'starts with VALUE_OF, an undefined variable, and a dictionary that carries no
'column index. Each test rebuilds a fresh dictionary fixture via
'PrepareDictionaryFixture so that dictionary state is isolated between runs.
'@depends ValueOfFormula, LLdictionary, LLVariables, BetterArray, CustomTest,
'  DictionaryTestFixture, TestHelpersLite


'@section Constants
'===============================================================================

Private Const DICTIONARY_SHEET As String = "ValueOfDictionary"

'@section Module State
'===============================================================================

Private Assert As CustomTest
Private dictionarySheet As Worksheet
Private dictionary As LLdictionary

'@section Helpers
'===============================================================================

'@sub-title Build a VALUE_OF parser bound to the shared dictionary
'@details
'Creates a ValueOfFormula instance by delegating to the predeclared factory
'method, passing the supplied expression string and the module-level
'dictionary fixture. Used by every test to obtain a parser under test.
'@param expression String. VALUE_OF expression to parse.
'@return ValueOfFormula. Parser instance ready for assertion.
Private Function BuildParser(ByVal expression As String) As ValueOfFormula
    Set BuildParser = ValueOfFormula.Create(expression, dictionary)
End Function

'@sub-title Wrap a sheet name in quotes matching workbook formula escaping
'@details
'Surrounds the sheet name with double-quote characters suitable for
'embedding inside a converted VALUE_OF formula string. Internal quotes
'are escaped by doubling them, matching Excel workbook formula conventions.
'@param sheetName String. Raw sheet name to quote.
'@return String. Quoted and escaped sheet name.
Private Function QuoteSheet(ByVal sheetName As String) As String
    QuoteSheet = """" & Replace(sheetName, """", """""") & """"
End Function

'@sub-title Locate a variable residing on a different sheet than the supplied lookup variable
'@details
'Iterates all variable names exposed by the dictionary and returns the first
'one whose sheet name differs from that of lookupVar. This helper is used by
'the cross-sheet rejection test to build a VALUE_OF expression whose lookup
'and value columns intentionally reside on different worksheets. Returns
'vbNullString when no such candidate exists.
'@param lookupVar String. Name of the variable whose sheet acts as the exclusion target.
'@return String. Name of a variable on a different sheet, or vbNullString if none found.
Private Function VariableOnDifferentSheet(ByVal lookupVar As String) As String
    Dim vars As LLVariables
    Dim names As BetterArray
    Dim idx As Long
    Dim candidate As String
    Dim targetSheet As String
    Dim candidateSheet As String

    Set vars = LLVariables.Create(dictionary)
    targetSheet = vars.Value("sheet name", lookupVar)

    Set names = vars.VariableNames
    For idx = names.LowerBound To names.UpperBound
        candidate = CStr(names.Item(idx))
        If StrComp(candidate, lookupVar, vbTextCompare) <> 0 Then
            candidateSheet = vars.Value("sheet name", candidate)
            If StrComp(candidateSheet, targetSheet, vbTextCompare) <> 0 Then
                VariableOnDifferentSheet = candidate
                Exit Function
            End If
        End If
    Next idx
End Function

'@section Module Lifecycle
'===============================================================================

'@sub-title Initialise the test harness output sheet and assertion object
'@details
'Runs once before any test in this module. Ensures the shared output
'worksheet exists and creates a CustomTest harness targeted at it.
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
    Assert.SetModuleName "TestValueOfFormula"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "ModuleInitialize", Err.Number, Err.Description
End Sub

'@sub-title Print results and tear down module-level state
'@details
'Runs once after all tests in this module have completed. Prints the
'accumulated test results, deletes the dictionary fixture worksheet, and
'releases all module-level object references.
'Every step before RestoreApp is wrapped, because RestoreApp has to run
'whatever happened above: the hooks here call BusyApp, which puts Excel in
'manual calculation, and the next module in the run would inherit that.
'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        DeleteWorksheet DICTIONARY_SHEET
    On Error GoTo 0

    RestoreApp

    Set Assert = Nothing
    Set dictionary = Nothing
    Set dictionarySheet = Nothing
End Sub

'@sub-title Prepare a fresh dictionary fixture before each test
'@details
'Populates the dictionary fixture worksheet via PrepareDictionaryFixture,
'then creates an LLdictionary instance backed by that worksheet. This
'ensures each test starts with a clean, consistent dictionary state. The
'fixture is left unprepared; a test that needs the prepared columns calls
'Prepare itself.
'@TestInitialize
Public Sub TestInitialize()
    On Error GoTo Fail
    BusyApp
    PrepareDictionaryFixture DICTIONARY_SHEET
    Set dictionarySheet = ThisWorkbook.Worksheets(DICTIONARY_SHEET)
    Set dictionary = LLdictionary.Create(dictionarySheet, 1, 1)
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInitialize", Err.Number, Err.Description
End Sub

'@sub-title Flush assertions and release per-test state
'@details
'Flushes any pending assertion output and releases the dictionary and
'worksheet references so the next test starts with a clean slate.
'@TestCleanup
Public Sub TestCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.Flush
        End If
    On Error GoTo 0

    Set dictionary = Nothing
    Set dictionarySheet = Nothing
End Sub

'@section Tests
'===============================================================================

'@sub-title Verify valid VALUE_OF formula converts to the new signature with indices and sheet name
'@details
'Arranges a valid three-argument VALUE_OF expression referencing dictionary
'variables lauto_drop_h2, choi_h2, and text_h2. Prepares the dictionary,
'resolves the expected sheet name and column indices via LLVariables, then
'asserts that the parser reports validity, produces the correct converted
'formula string with embedded sheet name and indices, and reports no
'failure reason. Also verifies that individual metadata properties
'(LookupSheetName, LookupColumnIndex, ValueColumnIndex) match the
'expected values from the dictionary.
'@TestMethod("ValueOfFormula")
Public Sub TestValueOfFormulaConvertsToNewSignature()
    CustomTestSetTitles Assert, "ValueOfFormula", "TestValueOfFormulaConvertsToNewSignature"
    On Error GoTo Fail

    Dim parser As ValueOfFormula
    Dim vars As LLVariables
    Dim lookupSheet As String
    Dim lookupIndex As Long
    Dim valueIndex As Long
    Dim expected As String

    Set parser = BuildParser("VALUE_OF(lauto_drop_h2, choi_h2, text_h2)")
    dictionary.Prepare
    Set vars = LLVariables.Create(dictionary)

    lookupSheet = vars.Value("sheet name", "choi_h2")
    lookupIndex = vars.Index("choi_h2")
    valueIndex = vars.Index("text_h2")

    expected = "VALUE_OF(lauto_drop_h2, " & QuoteSheet(lookupSheet) & ", " & CStr(lookupIndex) & ", " & CStr(valueIndex) & ")"

    Assert.IsTrue parser.Valid, "Expected VALUE_OF parser to accept valid arguments"
    Assert.AreEqual expected, parser.ConvertedFormula, "Converted VALUE_OF formula should include sheet name and column indices"
    Assert.AreEqual lookupSheet, parser.LookupSheetName, "Lookup sheet should match dictionary metadata"
    Assert.AreEqual lookupIndex, parser.LookupColumnIndex, "Lookup column index should be retrieved from dictionary"
    Assert.AreEqual valueIndex, parser.ValueColumnIndex, "Value column index should be retrieved from dictionary"
    Assert.AreEqual vbNullString, parser.FailureReason, "Valid formulas should not report failure reasons"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValueOfFormulaConvertsToNewSignature", Err.Number, Err.Description
End Sub

'@sub-title Verify VALUE_OF formula rejects lookup and value variables on different sheets
'@details
'Arranges a VALUE_OF expression whose lookup variable (choi_h2) and value
'variable reside on different worksheets by using the VariableOnDifferentSheet
'helper to find a mismatched candidate. Asserts that the parser reports
'invalid status, returns vbNullString for the converted formula, provides
'a descriptive failure reason about the sheet mismatch, and resets all
'metadata properties (LookupSheetName, LookupColumnIndex, ValueColumnIndex)
'to their default empty or zero values.
'@TestMethod("ValueOfFormula")
Public Sub TestValueOfFormulaRejectsCrossSheetArguments()
    CustomTestSetTitles Assert, "ValueOfFormula", "TestValueOfFormulaRejectsCrossSheetArguments"
    On Error GoTo Fail

    Dim mismatchedVar As String
    Dim parser As ValueOfFormula
    Dim failureMessage As String

    mismatchedVar = VariableOnDifferentSheet("choi_h2")
    Assert.IsTrue LenB(mismatchedVar) > 0, "Fixture should expose a variable on a different sheet for testing"

    Set parser = BuildParser("VALUE_OF(lauto_drop_h2, choi_h2, " & mismatchedVar & ")")

    failureMessage = "VALUE_OF expects lookup and value variables to share the same worksheet"

    Assert.IsFalse parser.Valid, "Cross-sheet VALUE_OF arguments must be rejected"
    Assert.AreEqual vbNullString, parser.ConvertedFormula, "Invalid formulas should not produce a converted expression"
    Assert.AreEqual failureMessage, parser.FailureReason, "Failure reason should highlight sheet mismatch"
    Assert.AreEqual vbNullString, parser.LookupSheetName, "Lookup sheet should be cleared after failure"
    Assert.AreEqual 0&, parser.LookupColumnIndex, "Lookup column index should reset on failure"
    Assert.AreEqual 0&, parser.ValueColumnIndex, "Value column index should reset on failure"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValueOfFormulaRejectsCrossSheetArguments", Err.Number, Err.Description
End Sub

'@sub-title Verify a token that only starts with VALUE_OF is rejected
'@details
'Arranges VALUE_OFX(...), which shares its first eight characters with the
'real token. Acts by reading Valid and ConvertedFormula. Asserts that the
'expression is rejected. The old class compared the leading characters only
'and then looked for the first bracket anywhere, so VALUE_OFX and
'VALUE_OF_EXTRA both parsed as a VALUE_OF.
'@TestMethod("ValueOfFormula")
Public Sub TestValueOfFormulaRejectsLongerToken()
    CustomTestSetTitles Assert, "ValueOfFormula", "TestValueOfFormulaRejectsLongerToken"
    On Error GoTo Fail

    Dim parser As ValueOfFormula

    Set parser = BuildParser("VALUE_OFX(lauto_drop_h2, choi_h2, text_h2)")

    Assert.IsFalse parser.Valid, "A longer token should not parse as VALUE_OF"
    Assert.AreEqual vbNullString, parser.ConvertedFormula, "A rejected token should produce no converted expression"
    Assert.IsTrue (LenB(parser.FailureReason) > 0), "A rejected expression should carry a reason"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValueOfFormulaRejectsLongerToken", Err.Number, Err.Description
End Sub

'@sub-title Verify an undefined variable is named in the failure reason
'@details
'Arranges a VALUE_OF whose value argument names a variable the dictionary
'does not hold. Acts by reading Valid and FailureReason. Asserts the
'expression is rejected and the message names the variable, so the user can
'find it in the dictionary.
'@TestMethod("ValueOfFormula")
Public Sub TestValueOfFormulaNamesTheMissingVariable()
    CustomTestSetTitles Assert, "ValueOfFormula", "TestValueOfFormulaNamesTheMissingVariable"
    On Error GoTo Fail

    Dim parser As ValueOfFormula
    Dim reason As String

    Set parser = BuildParser("VALUE_OF(lauto_drop_h2, choi_h2, not_a_variable)")
    reason = parser.FailureReason

    Assert.IsFalse parser.Valid, "An undefined variable should be rejected"
    Assert.IsTrue (InStr(1, reason, "not_a_variable", vbTextCompare) > 0), _
                  "Failure reason should name the missing variable. Found: " & reason
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValueOfFormulaNamesTheMissingVariable", Err.Number, Err.Description
End Sub

'@sub-title Verify a missing column index is reported as a missing column index
'@details
'Arranges a valid VALUE_OF over a dictionary that has NOT been prepared, so
'the column index column carries nothing. LLVariables.Index raises for that,
'and the old class swallowed the raise into one generic "Unable to parse
'VALUE_OF formula" message. Asserts the expression is rejected and the
'message says the column index is missing and names the variable.
'@TestMethod("ValueOfFormula")
Public Sub TestValueOfFormulaReportsMissingColumnIndex()
    CustomTestSetTitles Assert, "ValueOfFormula", "TestValueOfFormulaReportsMissingColumnIndex"
    On Error GoTo Fail

    Dim parser As ValueOfFormula
    Dim reason As String

    Set parser = BuildParser("VALUE_OF(lauto_drop_h2, choi_h2, text_h2)")
    reason = parser.FailureReason

    Assert.IsFalse parser.Valid, "An unprepared dictionary carries no column index, so the formula is rejected"
    Assert.IsTrue (InStr(1, reason, "column index", vbTextCompare) > 0), _
                  "Failure reason should say the column index is missing. Found: " & reason
    Assert.IsTrue (InStr(1, reason, "choi_h2", vbTextCompare) > 0), _
                  "Failure reason should name the variable. Found: " & reason
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValueOfFormulaReportsMissingColumnIndex", Err.Number, Err.Description
End Sub
