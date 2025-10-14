Attribute VB_Name = "TestFormulas"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")


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
Private Const FORMULA_GROUP_TABLE_MISMATCH_MESSAGE As String = "Grouped formulas require the first and third variables to belong to the same table."
Private Const FORMULA_EMPTY_MESSAGE As String = "The formula is empty. No formula were found"
Private Const FORMULA_GROUP_INVALID_GENERIC_MESSAGE As String = "Grouped formula '%1' must specify a valid aggregator after the GROUP prefix."
Private Const FORMULA_GROUP_UNKNOWN_AGGREGATOR_MESSAGE As String = "Grouped formula '%1' targets aggregator '%2', which is not allowed."
Private Const DEFAULTCRITERIAVAR As String = "lauto_man_h2"
Private Const DEFAULTCONDITIONVAR As String = "lauto_drop_h2"
Private Const DEFAULTRESULTVAR As String = "num_valid_h2"

'@section Module State
'===============================================================================

Private Assert As ICustomTest
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
    LinelistDictionary.Prepare
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
Private Function ShouldSkipFormulaRow(rowData As Variant, _
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
        Case "formula", "formulas", "choice_formula", "choice_formulas", "case_when"
            IsFormulaControl = True
    End Select
End Function

'@description Format the expected unknown-token reason using the template defined by Formulas.
'@param token String token reported as invalid.
'@return String reason message matching the Formulas implementation.
Private Function UnknownTokenReason(ByVal token As String) As String
    UnknownTokenReason = Replace(FORMULA_UNKNOWN_TOKEN_TEMPLATE, "%1", token, 1, 1, vbTextCompare)
End Function

'@description Identify two variables sharing the same table and a third variable used as condition.
Private Function SampleGroupedVariables(ByRef criteriaVar As String, _
                                        ByRef conditionVar As String, _
                                        ByRef resultVar As String, _
                                        ByRef tabName As String) As Boolean

    Dim vars As ILLVariables


    criteriaVar = DEFAULTCRITERIAVAR
    conditionVar = DEFAULTCONDITIONVAR
    resultVar = DEFAULTRESULTVAR

    Set vars = LLVariables.Create(LinelistDictionary)
    
    tabName = vars.Value(colName:="table name", varName:=criteriaVar)
    If tabName <> vars.Value(colName:="table name", varName:=resultVar) Then Exit Function
    If tabName = vars.Value(colName:="table name", varName:=conditionVar) Then Exit Function
    SampleGroupedVariables = True
End Function

'@description Retrieve a variable that belongs to a table different from the supplied one.
Private Function VariableFromDifferentTable(ByVal excludedTable As String, _
                                            ByRef variableName As String) As Boolean
    Dim vars As ILLVariables
    Set vars = LLVariables.Create(LinelistDictionary)

    variableName = DEFAULTCONDITIONVAR
    If vars.Value(colName:="table name", varName:=variableName) = excludedTable Then Exit Function
    VariableFromDifferentTable = True
End Function

'@description Build a grouped-reference string matching the production logic.
Private Function GroupedRangeReferenceForTest(ByVal variableName As String, _
                                              ByVal tableName As String, _
                                              ByVal useTableName As Boolean, _
                                              ByVal tablePrefix As String, _
                                              ByVal sheets As ILLSheets) As String
    If useTableName Then
        GroupedRangeReferenceForTest = tablePrefix & tableName & "[" & variableName & "]"
    Else
        GroupedRangeReferenceForTest = sheets.VariableAddress(variableName)
    End If
End Function

'@description Compose the expected SUMIFS formula for grouped parsing assertions.
Private Function ExpectedSumIfsFormula(ByVal criteriaVar As String, _
                                       ByVal conditionVar As String, _
                                       ByVal resultVar As String, _
                                       ByVal tableName As String, _
                                       ByVal tablePrefix As String, _
                                       ByVal useTableName As Boolean) As String
    Dim sheets As ILLSheets
    Dim criteriaRange As String
    Dim resultRange As String
    Dim conditionValue As String

    Set sheets = LLSheets.Create(LinelistDictionary)
    criteriaRange = GroupedRangeReferenceForTest(criteriaVar, tableName, useTableName, tablePrefix, sheets)
    resultRange = GroupedRangeReferenceForTest(resultVar, tableName, useTableName, tablePrefix, sheets)
    conditionValue = sheets.VariableAddress(conditionVar)

    ExpectedSumIfsFormula = "SUMIFS(" & resultRange & ", " & criteriaRange & ", " & conditionValue & ")"
End Function

'@description Compose the expected COUNTIFS formula with a non-empty criterion on the value range.
Private Function ExpectedCountIfsFormula(ByVal criteriaVar As String, _
                                         ByVal conditionVar As String, _
                                         ByVal resultVar As String, _
                                         ByVal tableName As String, _
                                         ByVal tablePrefix As String, _
                                         ByVal useTableName As Boolean) As String
    Dim sheets As ILLSheets
    Dim criteriaRange As String
    Dim resultRange As String
    Dim conditionValue As String

    Set sheets = LLSheets.Create(LinelistDictionary)
    criteriaRange = GroupedRangeReferenceForTest(criteriaVar, tableName, useTableName, tablePrefix, sheets)
    resultRange = GroupedRangeReferenceForTest(resultVar, tableName, useTableName, tablePrefix, sheets)
    conditionValue = sheets.VariableAddress(conditionVar)

    ExpectedCountIfsFormula = "COUNTIFS(" & criteriaRange & ", " & conditionValue & ", " & resultRange & ", " & Chr$(34) & "<>" & Chr$(34) & ")"
End Function

'@description Compose the expected array-style aggregator formula (e.g. AVERAGE(IF(...))).
Private Function ExpectedArrayGroupedFormula(ByVal aggregator As String, _
                                             ByVal criteriaVar As String, _
                                             ByVal conditionVar As String, _
                                             ByVal resultVar As String, _
                                             ByVal tableName As String, _
                                             ByVal tablePrefix As String, _
                                             ByVal useTableName As Boolean) As String
    Dim sheets As ILLSheets
    Dim criteriaRange As String
    Dim resultRange As String
    Dim conditionValue As String

    Set sheets = LLSheets.Create(LinelistDictionary)
    criteriaRange = GroupedRangeReferenceForTest(criteriaVar, tableName, useTableName, tablePrefix, sheets)
    resultRange = GroupedRangeReferenceForTest(resultVar, tableName, useTableName, tablePrefix, sheets)
    conditionValue = sheets.VariableAddress(conditionVar)

    ExpectedArrayGroupedFormula = aggregator & "(IF(" & criteriaRange & "=" & conditionValue & ", " & resultRange & "))"
End Function


'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestFormulas"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
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
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set FixtureSheet = Nothing
    Set DictionarySheet = Nothing
    Set FormulaDataSource = Nothing
    Set LinelistDictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Formulas")
'@description Ensure a single variable is valid in the simple context and produces no diagnostics.
Public Sub TestSimpleVariableValidForLinelist()
    CustomTestSetTitles Assert, "Formulas", "TestSimpleVariableValidForLinelist"
    Dim variableName As String
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    variableName = AnyVariableName()
    Set formulaInstance = BuildFormula(variableName)

    Assert.IsTrue formulaInstance.Valid("simple"), "Single variable should be valid for simple context"
    Assert.IsTrue formulaInstance.HasSetupVariables, "HasSetupVariables should be true when variable detected"
    Assert.AreEqual FORMULA_SUCCESS_MESSAGE, formulaInstance.Reason("simple"), "Reason should default to success message"
    Assert.IsFalse formulaInstance.HasChecking, "No diagnostics expected for simple variable"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSimpleVariableValidForLinelist", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Verify analysis context rejects formulas consisting solely of a variable and records diagnostics.
Public Sub TestAnalysisSingleVariableRejected()
    CustomTestSetTitles Assert, "Formulas", "TestAnalysisSingleVariableRejected"
    Dim variableName As String
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    variableName = AnyVariableName()
    Set formulaInstance = BuildFormula(variableName)

    Assert.IsFalse formulaInstance.Valid("analysis"), "Single variable should be rejected in analysis context"
    Assert.AreEqual FORMULA_ANALYSIS_SINGLE_VAR_MESSAGE, formulaInstance.Reason("analysis"), "Reason should explain aggregation requirement"
    Assert.IsTrue formulaInstance.HasChecking, "Rejection should record a checking entry"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAnalysisSingleVariableRejected", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Ensure empty expressions are rejected with the correct reason and diagnostics.
Public Sub TestEmptyFormulaRejected()
    CustomTestSetTitles Assert, "Formulas", "TestEmptyFormulaRejected"
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    Set formulaInstance = BuildFormula(vbNullString)

    Assert.IsFalse formulaInstance.Valid("analysis"), "Empty formula should be invalid"
    Assert.AreEqual FORMULA_EMPTY_MESSAGE, formulaInstance.Reason("analysis"), "Empty formula reason should explain the absence of tokens"
    Assert.IsTrue formulaInstance.HasChecking, "Empty formula should log a diagnostic entry"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEmptyFormulaRejected", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Ensure unknown tokens trigger failure, include the offending token, and log a diagnostic entry.
Public Sub TestUnknownTokenRecordsFailure()
    CustomTestSetTitles Assert, "Formulas", "TestUnknownTokenRecordsFailure"
    Dim formulaInstance As IFormulas
    Dim expectedReason As String

    On Error GoTo Fail

    Set formulaInstance = BuildFormula("UNKNOWN_TOKEN + 5")
    expectedReason = UnknownTokenReason("UNKNOWN_TOKEN")

    Assert.IsFalse formulaInstance.Valid("simple"), "Unknown token should mark formula invalid"
    Assert.AreEqual expectedReason, formulaInstance.Reason("simple"), "Reason should include token name"
    Assert.IsTrue formulaInstance.HasChecking, "Failure should create checking log"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUnknownTokenRecordsFailure", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Confirm custom aggregators convert to Excel equivalents during analysis parsing.
Public Sub TestCustomAggregatorTranslatesToAverage()
    CustomTestSetTitles Assert, "Formulas", "TestCustomAggregatorTranslatesToAverage"
    Dim formulaInstance As IFormulas
    Dim condition As IFormulaCondition
    Dim conditionVars As BetterArray
    Dim conditionConds As BetterArray
    Dim tableName As String

    On Error GoTo Fail

    Set formulaInstance = BuildFormula("MEAN")
    Set conditionVars = BetterArrayFromList(AnyVariableName())
    Set conditionConds = BetterArrayFromList("=1")
    tableName = LinelistDictionary.DataRange("table name").Cells(1, 1).Value
    Set condition = FormulaCondition.Create(conditionVars, conditionConds)
    
    Assert.IsTrue condition.Valid(LinelistDictionary, tableName), "Condition on custom aggregator should be valid"
    Assert.IsTrue formulaInstance.Valid("analysis"), "Custom MEAN should be accepted for analysis"
    Assert.AreEqual "AVERAGE", formulaInstance.ParsedAnalysisFormula(condition), "MEAN should translate to AVERAGE in analysis context"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCustomAggregatorTranslatesToAverage", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Check that structured references are applied when requested for linelist formulas.
Public Sub TestParsedLinelistStructuredReference()
    CustomTestSetTitles Assert, "Formulas", "TestParsedLinelistStructuredReference"
    Dim variableName As String
    Dim formulaInstance As IFormulas
    Dim parsed As String

    On Error GoTo Fail

    variableName = AnyVariableName()
    Set formulaInstance = BuildFormula(variableName & " + 5")

    parsed = formulaInstance.ParsedLinelistFormula(useTableName:=True, tablePrefix:="tbl_")
    Assert.IsTrue InStr(1, parsed, "tbl_", vbTextCompare) > 0, "Structured reference should include table prefix"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestParsedLinelistStructuredReference", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Ensure disabling structured references yields direct cell addresses in linelist formulas.
Public Sub TestParsedLinelistUsesCellReferencesWhenOptedOut()
    CustomTestSetTitles Assert, "Formulas", "TestParsedLinelistUsesCellReferencesWhenOptedOut"
    Dim variableName As String
    Dim formulaInstance As IFormulas
    Dim parsed As String
    Dim sheets As ILLSheets
    Dim expectedAddress As String

    On Error GoTo Fail

    variableName = AnyVariableName()
    Set formulaInstance = BuildFormula(variableName & " + 1")
    parsed = formulaInstance.ParsedLinelistFormula(useTableName:=False)

    LinelistDictionary.Prepare
    Set sheets = LLSheets.Create(LinelistDictionary)
    expectedAddress = sheets.VariableAddress(variableName)

    Assert.IsTrue formulaInstance.Valid("simple"), "Expression should be valid in simple context"
    Assert.IsTrue InStr(1, parsed, expectedAddress, vbTextCompare) > 0, "Parsed linelist formula should contain the direct cell address"
    Assert.IsFalse InStr(1, parsed, "[", vbBinaryCompare) > 0, "Parsed linelist formula should not include structured reference brackets when disabled"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestParsedLinelistUsesCellReferencesWhenOptedOut", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Validate that every formula-like dictionary entry parses without warnings.
Public Sub TestAllDictionaryFormulasParse()
    CustomTestSetTitles Assert, "Formulas", "TestAllDictionaryFormulasParse"
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
            Assert.IsTrue formulaInstance.Valid("simple"), "Failed to parse formula for variable " & variableName & "; formula is: " & formulaText
            Assert.IsFalse formulaInstance.HasChecking, "Parsing raised diagnostics for variable " & variableName
            Assert.AreEqual FORMULA_SUCCESS_MESSAGE, formulaInstance.Reason("simple"), "Unexpected reason for variable " & variableName
            evaluatedCount = evaluatedCount + 1
        End If
NextRow:
    Next rowData

    Assert.IsTrue (evaluatedCount > 0), "No dictionary formulas were evaluated"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAllDictionaryFormulasParse", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Confirm nested parentheses and irregular whitespace are handled correctly.
Public Sub TestNestedParenthesesAndWhitespace()
    CustomTestSetTitles Assert, "Formulas", "TestNestedParenthesesAndWhitespace"
    Dim variableName As String
    Dim expression As String
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    variableName = FixtureVariableName(1)
    expression = "   SUM  (  (" & variableName & "  +  2 )  * ( IF(1 = 1 , 3 , 4 ) ) )  "
    Set formulaInstance = BuildFormula(expression)

    Assert.IsTrue formulaInstance.Valid("analysis"), "Nested parentheses should parse"
    Assert.AreEqual FORMULA_SUCCESS_MESSAGE, formulaInstance.Reason("analysis"), "Nested formula should succeed"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestNestedParenthesesAndWhitespace", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Ensure escaped double-quotes inside string literals are recognised.
Public Sub TestHandlesEscapedQuotesWithinLiterals()
    CustomTestSetTitles Assert, "Formulas", "TestHandlesEscapedQuotesWithinLiterals"
    Dim expression As String
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    expression = "IF(" & Chr$(34) & "Alpha" & Chr$(34) & Chr$(34) & "Beta" & Chr$(34) & ", ""OK"", ""KO"")"
    Set formulaInstance = BuildFormula(expression)

    Assert.IsTrue formulaInstance.Valid("analysis"), "Escaped quotes should parse in analysis scope"
    Assert.IsTrue formulaInstance.Valid("simple"), "Escaped quotes should parse in simple scope"
    Assert.IsFalse formulaInstance.HasSetupVariables, "Formulas with no setup variables should be detected"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestHandlesEscapedQuotesWithinLiterals", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Verify boolean literals participate in expressions without causing failures.
Public Sub TestBooleanLiteralsAccepted()
    CustomTestSetTitles Assert, "Formulas", "TestBooleanLiteralsAccepted"
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    Set formulaInstance = BuildFormula("IF(TRUE, FALSE, TRUE)")

    Assert.IsTrue formulaInstance.Valid("analysis"), "Boolean literals should parse in analysis scope"
    Assert.IsTrue formulaInstance.Valid("simple"), "Boolean literals should parse in simple scope"
    Assert.IsFalse formulaInstance.HasSetupVariables, "Boolean literals count as literals"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBooleanLiteralsAccepted", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Ensure unknown functions trigger the standard unknown-token failure message.
Public Sub TestInvalidFunctionRaisesChecking()
    CustomTestSetTitles Assert, "Formulas", "TestInvalidFunctionRaisesChecking"
    Dim formulaInstance As IFormulas
    Dim expectedReason As String
    Dim variableName As String

    On Error GoTo Fail

    variableName = AnyVariableName()
    Set formulaInstance = BuildFormula("NOTAFUNCTION(" & variableName & ")")
    expectedReason = UnknownTokenReason("NOTAFUNCTION")

    Assert.IsFalse formulaInstance.Valid("analysis"), "Invalid function should fail"
    Assert.AreEqual expectedReason, formulaInstance.Reason("analysis"), "Reason should explain invalid function"
    Assert.IsTrue formulaInstance.HasChecking, "Failure must be logged"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInvalidFunctionRaisesChecking", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Ensure custom N aggregator translations do not leave empty parentheses in analysis output.
Public Sub TestCustomNRemovesEmptyInvocation()
    CustomTestSetTitles Assert, "Formulas", "TestCustomNRemovesEmptyInvocation"
    Const TABLE_PREFIX As String = "f_"
    Dim variableName As String
    Dim tableName As String
    Dim formulaInstance As IFormulas
    Dim conditionVars As BetterArray
    Dim conditionConds As BetterArray
    Dim formCondition As IFormulaCondition
    Dim parsed As String

    On Error GoTo Fail

    variableName = AnyVariableName()
    tableName = LinelistDictionary.DataRange("table name").Cells(1, 1).Value

    Set formulaInstance = BuildFormula("IF(N()>0, 1, 0)")
    Set conditionVars = BetterArrayFromList(variableName)
    Set conditionConds = BetterArrayFromList(">0")
    Set formCondition = FormulaCondition.Create(conditionVars, conditionConds)

    Assert.IsFalse formulaInstance.Valid("simple"), "Expression using custom N should not be valid in simple context"
    Assert.IsTrue formulaInstance.Valid("analysis"), "Expression using custom N should be valid in analysis context"

    parsed = formulaInstance.ParsedAnalysisFormula(formCondition, tablePrefix:=TABLE_PREFIX)

    Assert.IsTrue InStr(1, parsed, "COUNTIFS(", vbTextCompare) > 0, "Custom N should translate to COUNTIFS. Parsed formula: " & parsed
    Assert.IsTrue InStr(1, parsed, TABLE_PREFIX & tableName & "[" & variableName & "]", vbTextCompare) > 0, "COUNTIFS should reference the structured table column"
    Assert.AreEqual 0&, InStr(1, parsed, ")()", vbBinaryCompare), "COUNTIFS translation should not contain empty parentheses"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAnalysisCustomNRemovesEmptyInvocation", Err.Number, Err.Description
End Sub



'@TestMethod("Formulas")
'@description Detect formulas missing closing parentheses and report descriptive feedback.
Public Sub TestUnmatchedParenthesesDetected()
    CustomTestSetTitles Assert, "Formulas", "TestUnmatchedParenthesesDetected"
    Dim variableName As String
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    variableName = FixtureVariableName(2)
    Set formulaInstance = BuildFormula("SUM((" & variableName & " + 1")

    Assert.IsFalse formulaInstance.Valid("analysis"), "Unmatched parentheses should fail"
    Assert.AreEqual FORMULA_PAREN_MISMATCH_MESSAGE, formulaInstance.Reason("analysis"), "Reason should mention unmatched parentheses"
    Assert.IsTrue formulaInstance.HasChecking, "Failure must be logged"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestUnmatchedParenthesesDetected", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Detect closing parentheses that appear before any opening parenthesis.
Public Sub TestClosingParenthesisBeforeOpeningDetected()
    CustomTestSetTitles Assert, "Formulas", "TestClosingParenthesisBeforeOpeningDetected"
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    Set formulaInstance = BuildFormula(")1")

    Assert.IsFalse formulaInstance.Valid("analysis"), "Leading closing parenthesis should fail"
    Assert.AreEqual FORMULA_NEGATIVE_PAREN_MESSAGE, formulaInstance.Reason("analysis"), "Reason should mention negative parentheses"
    Assert.IsTrue formulaInstance.HasChecking, "Failure must be logged"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestClosingParenthesisBeforeOpeningDetected", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Reject characters not present in the allowed special-character table.
Public Sub TestDisallowedCharacterRejected()
    CustomTestSetTitles Assert, "Formulas", "TestDisallowedCharacterRejected"
    Dim formulaInstance As IFormulas
    Dim expectedReason As String

    On Error GoTo Fail

    Set formulaInstance = BuildFormula("é")
    expectedReason = UnknownTokenReason("é")

    Assert.IsFalse formulaInstance.Valid("analysis"), "Disallowed characters should fail"
    Assert.AreEqual expectedReason, formulaInstance.Reason("analysis"), "Reason should reference the disallowed character"
    Assert.IsTrue formulaInstance.HasChecking, "Failure must be logged"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDisallowedCharacterRejected", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Ensure grouped SUMIFS expressions emit the native SUMIFS function with structured references.
Public Sub TestGroupedSumIfsUsesNativeFunction()
    CustomTestSetTitles Assert, "Formulas", "TestGroupedSumIfsUsesNativeFunction"
    Const TABLE_PREFIX As String = "f_"
    Dim criteriaVar As String
    Dim conditionVar As String
    Dim resultVar As String
    Dim tableName As String
    Dim expression As String
    Dim formulaInstance As IFormulas
    Dim expected As String

    On Error GoTo Fail

    If Not SampleGroupedVariables(criteriaVar, conditionVar, resultVar, tableName) Then
        Assert.LogFailure "Unable to identify grouped variables for SUMIFS test"
        Exit Sub
    End If

    expression = "SUMIFS(" & criteriaVar & ", " & conditionVar & ", " & resultVar & ")"
    Set formulaInstance = BuildFormula(expression)

    Assert.IsTrue formulaInstance.Valid("analysis"), "Grouped SUMIFS should be valid"
    Assert.AreEqual "Yes", formulaInstance.IsGrouped, "Grouped SUMIFS should report grouped state"

    expected = ExpectedSumIfsFormula(criteriaVar, conditionVar, resultVar, tableName, TABLE_PREFIX, True)
    Assert.AreEqual expected, formulaInstance.ParsedLinelistFormula(useTableName:=True, tablePrefix:=TABLE_PREFIX), "Linelist SUMIFS output mismatch"
    Assert.AreEqual expected, formulaInstance.ParsedAnalysisFormula(formCond:=Nothing, tablePrefix:=TABLE_PREFIX), "Analysis SUMIFS output mismatch"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestGroupedSumIfsUsesNativeFunction", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Ensure grouped COUNTIFS appends a non-empty criterion for the aggregation range.
Public Sub TestGroupedCountIfsAddsNotBlankCriteria()
    CustomTestSetTitles Assert, "Formulas", "TestGroupedCountIfsAddsNotBlankCriteria"
    Const TABLE_PREFIX As String = "f_"
    Dim criteriaVar As String
    Dim conditionVar As String
    Dim resultVar As String
    Dim tableName As String
    Dim expression As String
    Dim formulaInstance As IFormulas
    Dim expected As String

    On Error GoTo Fail

    If Not SampleGroupedVariables(criteriaVar, conditionVar, resultVar, tableName) Then
        Assert.LogFailure "Unable to identify grouped variables for COUNTIFS test"
        Exit Sub
    End If

    expression = "COUNTIFS(" & criteriaVar & ", " & conditionVar & ", " & resultVar & ")"
    Set formulaInstance = BuildFormula(expression)

    Assert.IsTrue formulaInstance.Valid("analysis"), "Grouped COUNTIFS should be valid"
    Assert.AreEqual "Yes", formulaInstance.IsGrouped, "Grouped COUNTIFS should report grouped state"

    expected = ExpectedCountIfsFormula(criteriaVar, conditionVar, resultVar, tableName, TABLE_PREFIX, True)
    Assert.AreEqual expected, formulaInstance.ParsedLinelistFormula(useTableName:=True, tablePrefix:=TABLE_PREFIX), "Linelist COUNTIFS output mismatch"
    Assert.AreEqual expected, formulaInstance.ParsedAnalysisFormula(formCond:=Nothing, tablePrefix:=TABLE_PREFIX), "Analysis COUNTIFS output mismatch"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestGroupedCountIfsAddsNotBlankCriteria", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Ensure grouped MEANIFS expressions create array-style AVERAGE(IF()) formulas.
Public Sub TestGroupedMeanIfsBuildsArrayFormula()
    CustomTestSetTitles Assert, "Formulas", "TestGroupedMeanIfsBuildsArrayFormula"
    Const TABLE_PREFIX As String = "f_"
    Dim criteriaVar As String
    Dim conditionVar As String
    Dim resultVar As String
    Dim tableName As String
    Dim expression As String
    Dim formulaInstance As IFormulas
    Dim expectedLinelist As String
    Dim expectedAnalysis As String

    On Error GoTo Fail

    If Not SampleGroupedVariables(criteriaVar, conditionVar, resultVar, tableName) Then
        Assert.LogFailure "Unable to identify grouped variables for MEANIFS test"
        Exit Sub
    End If

    expression = "MEANIFS(" & criteriaVar & ", " & conditionVar & ", " & resultVar & ")"
    Set formulaInstance = BuildFormula(expression)

    Assert.IsTrue formulaInstance.Valid("analysis"), "Grouped MEANIFS should be valid in analysis scope"
    Assert.IsTrue formulaInstance.Valid("simple"), "Grouped MEANIFS should be valid in simple scope"
    Assert.AreEqual "Yes", formulaInstance.IsGrouped, "Grouped MEANIFS should report grouped state"

    expectedLinelist = ExpectedArrayGroupedFormula("AVERAGE", criteriaVar, conditionVar, resultVar, tableName, vbNullString, False)
    expectedAnalysis = ExpectedArrayGroupedFormula("AVERAGE", criteriaVar, conditionVar, resultVar, tableName, TABLE_PREFIX, True)

    Assert.AreEqual expectedLinelist, formulaInstance.ParsedLinelistFormula(), "Linelist MEANIFS output mismatch"
    Assert.AreEqual expectedAnalysis, formulaInstance.ParsedAnalysisFormula(formCond:=Nothing, tablePrefix:=TABLE_PREFIX), "Analysis MEANIFS output mismatch"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestGroupedMeanIfsBuildsArrayFormula", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Validate that generic GROUP_SUM expressions produce SUM(IF()) style formulas.
Public Sub TestGenericGroupSumBuildsArrayFormula()
    CustomTestSetTitles Assert, "Formulas", "TestGenericGroupSumBuildsArrayFormula"
    Const TABLE_PREFIX As String = "f_"
    Dim criteriaVar As String
    Dim conditionVar As String
    Dim resultVar As String
    Dim tableName As String
    Dim expression As String
    Dim formulaInstance As IFormulas
    Dim expectedLinelist As String
    Dim expectedAnalysis As String

    On Error GoTo Fail

    If Not SampleGroupedVariables(criteriaVar, conditionVar, resultVar, tableName) Then
        Assert.LogFailure "Unable to identify grouped variables for GROUP_SUM test"
        Exit Sub
    End If

    expression = "GROUP_SUM(" & criteriaVar & ", " & conditionVar & ", " & resultVar & ")"
    Set formulaInstance = BuildFormula(expression)

    Assert.IsTrue formulaInstance.Valid("simple"), "GROUP_SUM should be valid in simple scope"
    Assert.IsTrue formulaInstance.Valid("analysis"), "GROUP_SUM should be valid in analysis scope"
    Assert.AreEqual "Yes", formulaInstance.IsGrouped, "GROUP_SUM should report grouped state"

    expectedLinelist = ExpectedArrayGroupedFormula("SUM", criteriaVar, conditionVar, resultVar, tableName, vbNullString, False)
    expectedAnalysis = ExpectedArrayGroupedFormula("SUM", criteriaVar, conditionVar, resultVar, tableName, TABLE_PREFIX, True)

    Assert.AreEqual expectedLinelist, formulaInstance.ParsedLinelistFormula(), "Linelist GROUP_SUM output mismatch"
    Assert.AreEqual expectedAnalysis, formulaInstance.ParsedAnalysisFormula(formCond:=Nothing, tablePrefix:=TABLE_PREFIX), "Analysis GROUP_SUM output mismatch"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestGenericGroupSumBuildsArrayFormula", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Reject grouped formulas when the result variable does not share the criteria table.
Public Sub TestGroupedTableMismatchRejected()
    CustomTestSetTitles Assert, "Formulas", "TestGroupedTableMismatchRejected"
    Dim criteriaVar As String
    Dim conditionVar As String
    Dim resultVar As String
    Dim tableName As String
    Dim mismatchedVar As String
    Dim expression As String
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    If Not SampleGroupedVariables(criteriaVar, conditionVar, resultVar, tableName) Then
        Assert.LogFailure "Unable to identify grouped variables for mismatch test"
        Exit Sub
    End If

    If Not VariableFromDifferentTable(tableName, mismatchedVar) Then
        Assert.LogFailure "Unable to locate variable on a different table for mismatch validation"
        Exit Sub
    End If

    expression = "SUMIFS(" & criteriaVar & ", " & conditionVar & ", " & mismatchedVar & ")"
    Set formulaInstance = BuildFormula(expression)

    Assert.IsFalse formulaInstance.Valid("analysis"), "Grouped formula with mismatched tables must be rejected"
    Assert.AreEqual FORMULA_GROUP_TABLE_MISMATCH_MESSAGE, formulaInstance.Reason("analysis"), "Mismatch reason should explain table constraint"
    Assert.AreEqual "No", formulaInstance.IsGrouped, "Invalid grouped formula should not report grouped state"
    Assert.IsTrue formulaInstance.HasChecking, "Mismatch should register a checking entry"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestGroupedTableMismatchRejected", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Reject generic grouped formulas that omit an aggregator after the GROUP prefix.
Public Sub TestGenericGroupMissingAggregatorRejected()
    CustomTestSetTitles Assert, "Formulas", "TestGenericGroupMissingAggregatorRejected"
    Dim criteriaVar As String
    Dim conditionVar As String
    Dim resultVar As String
    Dim tableName As String
    Dim expression As String
    Dim formulaInstance As IFormulas
    Dim expectedReason As String

    On Error GoTo Fail

    If Not SampleGroupedVariables(criteriaVar, conditionVar, resultVar, tableName) Then
        Assert.LogFailure "Unable to identify grouped variables for missing aggregator test"
        Exit Sub
    End If

    expression = "GROUP_(" & criteriaVar & ", " & conditionVar & ", " & resultVar & ")"
    Set formulaInstance = BuildFormula(expression)

    Assert.IsFalse formulaInstance.Valid("analysis"), "GROUP_ should be rejected when no aggregator is specified"

    expectedReason = Replace(FORMULA_GROUP_INVALID_GENERIC_MESSAGE, "%1", "GROUP_", 1, 1, vbTextCompare)
    Assert.AreEqual expectedReason, formulaInstance.Reason("analysis"), "Missing aggregator reason mismatch"
    Assert.IsTrue formulaInstance.HasChecking, "Missing aggregator should log a diagnostic entry"
    Assert.AreEqual "No", formulaInstance.IsGrouped, "Invalid grouped formula should not report grouped state"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestGenericGroupMissingAggregatorRejected", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Reject generic grouped formulas targeting aggregators not present in the Excel catalog.
Public Sub TestGenericGroupUnknownAggregatorRejected()
    CustomTestSetTitles Assert, "Formulas", "TestGenericGroupUnknownAggregatorRejected"
    Dim criteriaVar As String
    Dim conditionVar As String
    Dim resultVar As String
    Dim tableName As String
    Dim expression As String
    Dim formulaInstance As IFormulas
    Dim expectedReason As String

    On Error GoTo Fail

    If Not SampleGroupedVariables(criteriaVar, conditionVar, resultVar, tableName) Then
        Assert.LogFailure "Unable to identify grouped variables for unknown aggregator test"
        Exit Sub
    End If

    expression = "GROUP_UNKNOWNFUNC(" & criteriaVar & ", " & conditionVar & ", " & resultVar & ")"
    Set formulaInstance = BuildFormula(expression)

    Assert.IsFalse formulaInstance.Valid("analysis"), "GROUP_UNKNOWNFUNC should be rejected when aggregator is not registered"

    expectedReason = Replace(FORMULA_GROUP_UNKNOWN_AGGREGATOR_MESSAGE, "%1", "GROUP_UNKNOWNFUNC", 1, 1, vbTextCompare)
    expectedReason = Replace(expectedReason, "%2", "UNKNOWNFUNC", 1, 1, vbTextCompare)
    Assert.AreEqual expectedReason, formulaInstance.Reason("analysis"), "Unknown aggregator reason mismatch"
    Assert.IsTrue formulaInstance.HasChecking, "Unknown aggregator should log a diagnostic entry"
    Assert.AreEqual "No", formulaInstance.IsGrouped, "Invalid grouped formula should not report grouped state"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestGenericGroupUnknownAggregatorRejected", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Confirm ParsedAnalysisFormula respects connectors provided by IFormulaCondition.
Public Sub TestParsedAnalysisFormulaAppliesConnector()
    CustomTestSetTitles Assert, "Formulas", "TestParsedAnalysisFormulaAppliesConnector"
    Dim variableName As String
    Dim expression As String
    Dim formulaInstance As IFormulas
    Dim condition As IFormulaCondition
    Dim conditionVars As BetterArray
    Dim conditionConds As BetterArray
    Dim tableName As String
    Dim parsed As String

    On Error GoTo Fail

    variableName = FixtureVariableName(3)
    expression = "SUM(" & variableName & ")"
    Set formulaInstance = BuildFormula(expression)

    Set conditionVars = BetterArrayFromList(variableName, variableName)
    Set conditionConds = BetterArrayFromList("=1", "<>""""")

    Dim vars As ILLVariables
    Set vars = LLVariables.Create(LinelistDictionary)
    tableName = vars.Value(colName:="table name", varName:=variableName)

    Set condition = FormulaCondition.Create(conditionVars, conditionConds)

    parsed = formulaInstance.ParsedAnalysisFormula(condition, tablePrefix:="tbl_", Connector:=" + ")

    Assert.IsTrue InStr(1, parsed, " + ", vbTextCompare) > 0, "Connector should be injected between conditions"
    Assert.IsTrue InStr(1, parsed, "IF", vbTextCompare) > 0, "Expression should include IF wrapper from FormulaCondition"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestParsedAnalysisFormulaAppliesConnector", Err.Number, Err.Description
End Sub

'@TestMethod("Formulas")
'@description Stress the parser with a lengthy expression to ensure repeated tokenisation remains successful.
Public Sub TestLargeFormulaParsesSuccessfully()
    CustomTestSetTitles Assert, "Formulas", "TestLargeFormulaParsesSuccessfully"
    Dim variableName As String
    Dim idx As Long
    Dim builder As String
    Dim formulaInstance As IFormulas

    On Error GoTo Fail

    variableName = FixtureVariableName(4)

    For idx = 1 To 25
        builder = builder & "SUM(" & variableName & " + " & CStr(idx) & ")"
        If idx < 25 Then builder = builder & " + "
    Next idx

    Set formulaInstance = BuildFormula(builder)

    Assert.IsTrue formulaInstance.Valid("analysis"), "Large formula should parse successfully"
    Assert.AreEqual FORMULA_SUCCESS_MESSAGE, formulaInstance.Reason("analysis"), "Reason should indicate success"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLargeFormulaParsesSuccessfully", Err.Number, Err.Description
End Sub
