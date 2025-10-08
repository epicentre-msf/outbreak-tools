Attribute VB_Name = "TestFormulas"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")

'@section Constants
'===============================================================================

Private Const FORMULA_SHEET As String = "FormulasFixture"
Private Const FORMULAS_TABLE_NAME As String = "T_XlsFonctions"
Private Const CHARACTERS_TABLE_NAME As String = "T_ascii"
Private Const DICTIONARY_SHEET As String = "FormulasDictionary"
Private Const FORMULA_SUCCESS_MESSAGE As String = "The formula seems correct"
Private Const FORMULA_ANALYSIS_SINGLE_VAR_MESSAGE As String = "Analysis formula can not consist of only one variable, you should use aggregation function"
Private Const FORMULA_UNKNOWN_TOKEN_TEMPLATE As String = "Unknown token '%1' encountered while parsing"
Private Const FORMULA_PAREN_MISMATCH_MESSAGE As String = "The formula contains unmatched parentheses"
Private Const FORMULA_NEGATIVE_PAREN_MESSAGE As String = "Closing parenthesis detected before opening one"

'@section Module State
'===============================================================================

Private Assert As Object
Private Fakes As Object
Private FixtureSheet As Worksheet
Private DictionarySheet As Worksheet
Private FormulaDataSource As IFormulaData
Private LinelistDictionary As ILLdictionary

'@section Helpers
'===============================================================================

'@description Prepare the dictionary worksheet using the shared fixture.
Private Sub PrepareDictionary()
    PrepareDictionaryFixture DICTIONARY_SHEET
    Set DictionarySheet = ThisWorkbook.Worksheets(DICTIONARY_SHEET)
    Set LinelistDictionary = LLdictionary.Create(DictionarySheet, 1, 1)
End Sub

'@description Build a Formulas instance backed by the shared dependencies.
'@param expression String pseudo-code expression to parse.
'@return IFormulas configured with the shared dictionary and formula data.
Private Function BuildFormula(ByVal expression As String) As IFormulas
    Set BuildFormula = Formulas.Create(LinelistDictionary, FormulaDataSource, expression)
End Function

'@description Retrieve a variable name from the dictionary fixture by index.
'@param index Long zero-based row index.
'@return String variable name held at the requested row.
Private Function FixtureVariableName(ByVal index As Long) As String
    FixtureVariableName = DictionaryFixtureValue(index, "Variable Name")
End Function

'@description Retrieve the first available variable name.
'@return String variable name.
Private Function AnyVariableName() As String
    AnyVariableName = FixtureVariableName(0)
End Function

'@description Determine whether a dictionary row should be ignored during bulk validation.
'@param rowData Variant array representing the row values.
'@param noteIndex Long index of the Note column.
'@return Boolean True when the row should be skipped.
Private Function ShouldSkipFormulaRow( rowData As Variant, _
                                      ByVal noteIndex As Long) As Boolean
    Dim noteText As String
    If noteIndex >= LBound(rowData) And noteIndex <= UBound(rowData) Then
        noteText = CStr(rowData(noteIndex))
        If InStr(1, noteText, "should fail", vbTextCompare) > 0 Then
            ShouldSkipFormulaRow = True
        End If
    End If
End Function

'@description Retrieve a specific column value from a dictionary row.
'@param rowData Variant array representing the row values.
'@param columnIndex Long index to extract.
'@return String column value or vbNullString when index is out of range.
Private Function RowValue( rowData As Variant, ByVal columnIndex As Long) As String
    If columnIndex >= LBound(rowData) And columnIndex <= UBound(rowData) Then
        RowValue = CStr(rowData(columnIndex))
    End If
End Function

'@description Determine whether a control value is expected to hold a formula expression.
'@param controlValue String value from the Control column.
'@return Boolean True when the row should be parsed as a formula.
Private Function IsFormulaControl(ByVal controlValue As String) As Boolean
    Dim normalized As String
    normalized = LCase$(controlValue)
    Select Case normalized
        Case "formula", "formulas", "choice_formula", "choice_formulas", "choce_formulas", "case_when"
            IsFormulaControl = True
    End Select
End Function

'@description Format the expected unknown-token reason using the template defined by Formulas.
'@param token String token reported as invalid.
'@return String reason message matching the Formulas implementation.
Private Function UnknownTokenReason(ByVal token As String) As String
    UnknownTokenReason = Replace(FORMULA_UNKNOWN_TOKEN_TEMPLATE, "%1", token, 1, 1, vbTextCompare)
End Function

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet FORMULA_SHEET
    DeleteWorksheet DICTIONARY_SHEET
    RestoreApp
    Set Assert = Nothing
    Set Fakes = Nothing
    Set FixtureSheet = Nothing
    Set DictionarySheet = Nothing
    Set FormulaDataSource = Nothing
    Set LinelistDictionary = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureSheet = PrepareFormulaFixtureSheet(FORMULA_SHEET, FORMULAS_TABLE_NAME, CHARACTERS_TABLE_NAME)
    Set FormulaDataSource = FormulaData.Create(FixtureSheet)
    PrepareDictionary
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set FixtureSheet = Nothing
    Set DictionarySheet = Nothing
    Set FormulaDataSource = Nothing
    Set LinelistDictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Formulas")
'@description Ensure a single variable is valid in the simple context and produces no diagnostics.
Private Sub TestSimpleVariableValidForLinelist()
    Dim variableName As String
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    variableName = AnyVariableName()
    Set formulaInstance = BuildFormula(variableName)

    Assert.IsTrue formulaInstance.Valid("simple"), "Single variable should be valid for simple context"
    Assert.IsTrue formulaInstance.HasLiterals, "HasLiterals should be true when variable detected"
    Assert.AreEqual FORMULA_SUCCESS_MESSAGE, formulaInstance.Reason("simple"), "Reason should default to success message"
    Assert.IsFalse formulaInstance.HasChecking, "No diagnostics expected for simple variable"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestSimpleVariableValidForLinelist"
End Sub

'@TestMethod("Formulas")
'@description Verify analysis context rejects formulas consisting solely of a variable and records diagnostics.
Private Sub TestAnalysisSingleVariableRejected()
    Dim variableName As String
    Dim formulaInstance As IFormulas

    variableName = AnyVariableName()
    Set formulaInstance = BuildFormula(variableName)

    Assert.IsFalse formulaInstance.Valid("analysis"), "Single variable should be rejected in analysis context"
    Assert.AreEqual FORMULA_ANALYSIS_SINGLE_VAR_MESSAGE, formulaInstance.Reason("analysis"), "Reason should explain aggregation requirement"
    Assert.IsTrue formulaInstance.HasChecking, "Rejection should record a checking entry"
End Sub

'@TestMethod("Formulas")
'@description Ensure unknown tokens trigger failure, include the offending token, and log a diagnostic entry.
Private Sub TestUnknownTokenRecordsFailure()
    Dim formulaInstance As IFormulas
    Dim expectedReason As String

    Set formulaInstance = BuildFormula("UNKNOWN_TOKEN + 5")
    expectedReason = UnknownTokenReason("UNKNOWN_TOKEN")

    Assert.IsFalse formulaInstance.Valid("simple"), "Unknown token should mark formula invalid"
    Assert.AreEqual expectedReason, formulaInstance.Reason("simple"), "Reason should include token name"
    Assert.IsTrue formulaInstance.HasChecking, "Failure should create checking log"
End Sub

'@TestMethod("Formulas")
'@description Confirm custom aggregators convert to Excel equivalents during analysis parsing.
Private Sub TestCustomAggregatorTranslatesToAverage()
    Dim formulaInstance As IFormulas
    Dim condition As IFormulaCondition
    Dim conditionVars As BetterArray
    Dim conditionConds As BetterArray
    Dim tableName As String

    Set formulaInstance = BuildFormula("MEAN")
    Set conditionVars = BetterArrayFromList(AnyVariableName())
    Set conditionConds = BetterArrayFromList("=1")
    tableName = DictionaryFixtureValue(0, "Table Name")
    Set condition = FormulaCondition.Create(conditionVars, conditionConds)
    condition.Valid LinelistDictionary, tableName

    Assert.IsTrue formulaInstance.Valid("analysis"), "Custom MEAN should be accepted for analysis"
    Assert.AreEqual "AVERAGE", formulaInstance.ParsedAnalysisFormula(condition), "MEAN should translate to AVERAGE in analysis context"
End Sub

'@TestMethod("Formulas")
'@description Check that structured references are applied when requested for linelist formulas.
Private Sub TestParsedLinelistStructuredReference()
    Dim variableName As String
    Dim formulaInstance As IFormulas
    Dim parsed As String

    variableName = AnyVariableName()
    Set formulaInstance = BuildFormula(variableName & " + 5")

    parsed = formulaInstance.ParsedLinelistFormula(useTableName:=True, tablePrefix:="tbl_")
    Assert.IsTrue InStr(1, parsed, "tbl_", vbTextCompare) > 0, "Structured reference should include table prefix"
End Sub

'@TestMethod("Formulas")
'@description Validate that every formula-like dictionary entry parses without warnings.
Private Sub TestAllDictionaryFormulasParse()
    Dim rows As Variant
    Dim rowData As Variant
    Dim controlIndex As Long
    Dim detailsIndex As Long
    Dim noteIndex As Long
    Dim nameIndex As Long
    Dim controlValue As String
    Dim formulaText As String
    Dim variableName As String
    Dim formulaInstance As IFormulas
    Dim evaluatedCount As Long

    On Error GoTo Fail

    rows = DictionaryFixtureRows()
    controlIndex = DictionaryHeaderIndex("Control")
    detailsIndex = DictionaryHeaderIndex("Control Details")
    noteIndex = DictionaryHeaderIndex("Note")
    nameIndex = DictionaryHeaderIndex("Variable Name")

    For Each rowData In rows
        controlValue = RowValue(rowData, controlIndex)
        If IsFormulaControl(controlValue) Then
            If ShouldSkipFormulaRow(rowData, noteIndex) Then GoTo NextRow
            formulaText = RowValue(rowData, detailsIndex)
            If LenB(formulaText) = 0 Then GoTo NextRow
            variableName = RowValue(rowData, nameIndex)

            Set formulaInstance = BuildFormula(formulaText)
            Assert.IsTrue formulaInstance.Valid("analysis"), "Failed to parse formula for variable " & variableName
            Assert.IsFalse formulaInstance.HasChecking, "Parsing raised diagnostics for variable " & variableName
            Assert.AreEqual FORMULA_SUCCESS_MESSAGE, formulaInstance.Reason("analysis"), "Unexpected reason for variable " & variableName
            evaluatedCount = evaluatedCount + 1
        End If
NextRow:
    Next rowData

    Assert.IsTrue (evaluatedCount > 0), "No dictionary formulas were evaluated"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAllDictionaryFormulasParse"
End Sub

'@TestMethod("Formulas")
'@description Confirm nested parentheses and irregular whitespace are handled correctly.
Private Sub TestNestedParenthesesAndWhitespace()
    Dim variableName As String
    Dim expression As String
    Dim formulaInstance As IFormulas

    variableName = FixtureVariableName(1)
    expression = "   SUM  (  (" & variableName & "  +  2 )  * ( IF(1 = 1 , 3 , 4 ) ) )  "
    Set formulaInstance = BuildFormula(expression)

    Assert.IsTrue formulaInstance.Valid("analysis"), "Nested parentheses should parse"
    Assert.AreEqual FORMULA_SUCCESS_MESSAGE, formulaInstance.Reason("analysis"), "Nested formula should succeed"
End Sub

'@TestMethod("Formulas")
'@description Ensure escaped double-quotes inside string literals are recognised.
Private Sub TestHandlesEscapedQuotesWithinLiterals()
    Dim expression As String
    Dim formulaInstance As IFormulas

    expression = "IF(" & Chr$(34) & "Alpha" & Chr$(34) & Chr$(34) & "Beta" & Chr$(34) & ", ""OK"", ""KO"")"
    Set formulaInstance = BuildFormula(expression)

    Assert.IsTrue formulaInstance.Valid("analysis"), "Escaped quotes should parse"
    Assert.IsTrue formulaInstance.HasLiterals, "Literal strings should be detected"
End Sub

'@TestMethod("Formulas")
'@description Verify boolean literals participate in expressions without causing failures.
Private Sub TestBooleanLiteralsAccepted()
    Dim formulaInstance As IFormulas

    Set formulaInstance = BuildFormula("IF(TRUE, FALSE, TRUE)")

    Assert.IsTrue formulaInstance.Valid("analysis"), "Boolean literals should parse"
    Assert.IsTrue formulaInstance.HasLiterals, "Boolean literals count as literals"
End Sub

'@TestMethod("Formulas")
'@description Ensure unknown functions trigger the standard unknown-token failure message.
Private Sub TestInvalidFunctionRaisesChecking()
    Dim formulaInstance As IFormulas
    Dim expectedReason As String
    Dim variableName As String

    variableName = AnyVariableName()
    Set formulaInstance = BuildFormula("NOTAFUNCTION(" & variableName & ")")
    expectedReason = UnknownTokenReason("NOTAFUNCTION")

    Assert.IsFalse formulaInstance.Valid("analysis"), "Invalid function should fail"
    Assert.AreEqual expectedReason, formulaInstance.Reason("analysis"), "Reason should explain invalid function"
    Assert.IsTrue formulaInstance.HasChecking, "Failure must be logged"
End Sub

'@TestMethod("Formulas")
'@description Detect formulas missing closing parentheses and report descriptive feedback.
Private Sub TestUnmatchedParenthesesDetected()
    Dim variableName As String
    Dim formulaInstance As IFormulas

    variableName = FixtureVariableName(2)
    Set formulaInstance = BuildFormula("SUM((" & variableName & " + 1")

    Assert.IsFalse formulaInstance.Valid("analysis"), "Unmatched parentheses should fail"
    Assert.AreEqual FORMULA_PAREN_MISMATCH_MESSAGE, formulaInstance.Reason("analysis"), "Reason should mention unmatched parentheses"
    Assert.IsTrue formulaInstance.HasChecking, "Failure must be logged"
End Sub

'@TestMethod("Formulas")
'@description Detect closing parentheses that appear before any opening parenthesis.
Private Sub TestClosingParenthesisBeforeOpeningDetected()
    Dim formulaInstance As IFormulas

    Set formulaInstance = BuildFormula(")1")

    Assert.IsFalse formulaInstance.Valid("analysis"), "Leading closing parenthesis should fail"
    Assert.AreEqual FORMULA_NEGATIVE_PAREN_MESSAGE, formulaInstance.Reason("analysis"), "Reason should mention negative parentheses"
    Assert.IsTrue formulaInstance.HasChecking, "Failure must be logged"
End Sub

'@TestMethod("Formulas")
'@description Reject characters not present in the allowed special-character table.
Private Sub TestDisallowedCharacterRejected()
    Dim formulaInstance As IFormulas
    Dim expectedReason As String

    Set formulaInstance = BuildFormula("é")
    expectedReason = UnknownTokenReason("é")

    Assert.IsFalse formulaInstance.Valid("analysis"), "Disallowed characters should fail"
    Assert.AreEqual expectedReason, formulaInstance.Reason("analysis"), "Reason should reference the disallowed character"
    Assert.IsTrue formulaInstance.HasChecking, "Failure must be logged"
End Sub

'@TestMethod("Formulas")
'@description Confirm ParsedAnalysisFormula respects connectors provided by IFormulaCondition.
Private Sub TestParsedAnalysisFormulaAppliesConnector()
    Dim variableName As String
    Dim expression As String
    Dim formulaInstance As IFormulas
    Dim condition As IFormulaCondition
    Dim conditionVars As BetterArray
    Dim conditionConds As BetterArray
    Dim tableName As String
    Dim parsed As String

    variableName = FixtureVariableName(3)
    expression = "SUM(" & variableName & ")"
    Set formulaInstance = BuildFormula(expression)

    Set conditionVars = BetterArrayFromList(variableName, variableName)
    Set conditionConds = BetterArrayFromList("=1", "<>\"\"")
    tableName = DictionaryFixtureValue(3, "Table Name")
    Set condition = FormulaCondition.Create(conditionVars, conditionConds)

    parsed = formulaInstance.ParsedAnalysisFormula(condition, tablePrefix:="tbl_", Connector:=" + ")

    Assert.IsTrue InStr(1, parsed, " + ", vbTextCompare) > 0, "Connector should be injected between conditions"
    Assert.IsTrue InStr(1, parsed, "IF", vbTextCompare) > 0, "Expression should include IF wrapper from FormulaCondition"
End Sub

'@TestMethod("Formulas")
'@description Stress the parser with a lengthy expression to ensure repeated tokenisation remains successful.
Private Sub TestLargeFormulaParsesSuccessfully()
    Dim variableName As String
    Dim idx As Long
    Dim builder As String
    Dim formulaInstance As IFormulas

    variableName = FixtureVariableName(4)

    For idx = 1 To 25
        builder = builder & "SUM(" & variableName & " + " & CStr(idx) & ")"
        If idx < 25 Then builder = builder & " + "
    Next idx

    Set formulaInstance = BuildFormula(builder)

    Assert.IsTrue formulaInstance.Valid("analysis"), "Large formula should parse successfully"
    Assert.AreEqual FORMULA_SUCCESS_MESSAGE, formulaInstance.Reason("analysis"), "Reason should indicate success"
End Sub
