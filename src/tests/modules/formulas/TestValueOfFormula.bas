Attribute VB_Name = "TestValueOfFormula"
Attribute VB_Description = "Verifies VALUE_OF formula conversion"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Verifies VALUE_OF formula conversion")

'@section Constants
'===============================================================================

Private Const DICTIONARY_SHEET As String = "ValueOfDictionary"

'@section Module State
'===============================================================================

Private Assert As ICustomTest
Private dictionarySheet As Worksheet
Private dictionary As ILLdictionary

'@section Helpers
'===============================================================================

'@description Build a VALUE_OF parser bound to the shared dictionary.
'@param expression String VALUE_OF expression.
'@return IValueOfFormula parser instance.
Private Function BuildParser(ByVal expression As String) As IValueOfFormula
    Set BuildParser = ValueOfFormula.Create(expression, dictionary)
End Function

'@description Wrap a sheet name in quotes matching workbook formula escaping.
Private Function QuoteSheet(ByVal sheetName As String) As String
    QuoteSheet = """" & Replace(sheetName, """", """""") & """"
End Function

'@description Locate a variable residing on a different sheet than the supplied lookup variable.
Private Function VariableOnDifferentSheet(ByVal lookupVar As String) As String
    Dim vars As ILLVariables
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

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestValueOfFormula"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()

    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet DICTIONARY_SHEET
    Set Assert = Nothing
    Set dictionary = Nothing
    Set dictionarySheet = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    PrepareDictionaryFixture DICTIONARY_SHEET
    Set dictionarySheet = ThisWorkbook.Worksheets(DICTIONARY_SHEET)
    Set dictionary = LLdictionary.Create(dictionarySheet, 1, 1)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set dictionary = Nothing
    Set dictionarySheet = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("ValueOfFormula")
'@description Ensure VALUE_OF formulas convert to the new signature with indices and sheet name.
Public Sub TestValueOfFormulaConvertsToNewSignature()
    CustomTestSetTitles Assert, "ValueOfFormula", "TestValueOfFormulaConvertsToNewSignature"
    On Error GoTo Fail

    Dim parser As IValueOfFormula
    Dim vars As ILLVariables
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

'@TestMethod("ValueOfFormula")
'@description Reject VALUE_OF formulas when lookup and value variables reside on different sheets.
Public Sub TestValueOfFormulaRejectsCrossSheetArguments()
    CustomTestSetTitles Assert, "ValueOfFormula", "TestValueOfFormulaRejectsCrossSheetArguments"
    On Error GoTo Fail

    Dim mismatchedVar As String
    Dim parser As IValueOfFormula
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
