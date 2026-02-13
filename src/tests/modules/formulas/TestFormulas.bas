Attribute VB_Name = "TestFormulas"
Attribute VB_Description = "Tests for the Formulas class parser and validator"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@ModuleDescription("Tests for the Formulas class parser and validator")
'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'Tests the Formulas class, which is responsible for parsing, validating, and
'translating pseudo-code formula expressions into Excel-native syntax. The
'module covers simple and analysis validation contexts, token recognition
'(variables, functions, operators, literals, parentheses), structured and
'cell-based reference generation, grouped formula translation (SUMIFS,
'COUNTIFS, MEANIFS, GROUP_*), custom aggregator mapping, error diagnostics,
'and edge cases such as empty input, escaped quotes, and large expressions.
'Each test builds a fresh dictionary and FormulaData fixture via worksheet
'helpers so tests run in isolation.
'@depends Formulas, IFormulas, FormulaData, IFormulaData, FormulaCondition,
'IFormulaCondition, LLdictionary, ILLdictionary, LLVariables, ILLVariables,
'LLSheets, ILLSheets, BetterArray, CustomTest, ICustomTest,
'DictionaryTestFixture, FormulaTestFixture


'@section Constants
'===============================================================================
'@description Fixture sheet names, expected message templates, and default
'variable names used across the test methods.

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
'@description Module-level variables holding the test assertion object, the
'formula fixture worksheet, the dictionary worksheet, and the shared
'FormulaData and LLdictionary instances rebuilt before every test.

Private Assert As ICustomTest
Private Fakes As Object
Private FixtureSheet As Worksheet
Private DictionarySheet As Worksheet
Private FormulaDataSource As IFormulaData
Private LinelistDictionary As ILLdictionary

'@section Helpers
'===============================================================================
'@description Private helper functions that build fixture objects, retrieve
'dictionary values, compose expected formula strings, and support the
'grouped-formula test methods.

'@sub-title Prepare the dictionary worksheet using the shared fixture
'@details
'Delegates to the DictionaryTestFixture helper to create or refresh the
'dictionary sheet, then wraps it in an LLdictionary instance and calls
'Prepare so that internal caches (sheets, variables) are ready for use.
Private Sub PrepareDictionary()
    PrepareDictionaryFixture DICTIONARY_SHEET
    Set DictionarySheet = ThisWorkbook.Worksheets(DICTIONARY_SHEET)
    Set LinelistDictionary = LLdictionary.Create(DictionarySheet, 1, 1)
    LinelistDictionary.Prepare
End Sub

'@sub-title Build a Formulas instance backed by the shared dependencies
'@details
'Creates a new Formulas object wired to the module-level dictionary and
'FormulaData source. The expression is parsed eagerly during Create, so
'the returned instance is immediately ready for Valid/Reason/Parsed calls.
'@param expression String. Pseudo-code expression to parse.
'@return IFormulas. A configured Formulas instance backed by the shared dictionary and formula data.
Private Function BuildFormula(ByVal expression As String) As IFormulas
    Set BuildFormula = Formulas.Create(LinelistDictionary, FormulaDataSource, expression)
End Function

'@sub-title Retrieve a variable name from the dictionary fixture by index
'@details
'Looks up a variable name from the dictionary fixture at the given
'zero-based row index. Used to obtain known-good variable names for
'constructing test expressions.
'@param index Long. Zero-based row index into the dictionary fixture.
'@return String. Variable name held at the requested row.
Private Function FixtureVariableName(ByVal index As Long) As String
    FixtureVariableName = DictionaryFixtureValue(index, "Variable Name")
End Function

'@sub-title Retrieve the first available variable name
'@details
'Convenience wrapper that returns the variable name at row index 0 of the
'dictionary fixture. Most tests that need any valid variable call this.
'@return String. The first variable name in the fixture.
Private Function AnyVariableName() As String
    AnyVariableName = FixtureVariableName(0)
End Function

'@sub-title Determine whether a dictionary row should be skipped during bulk validation
'@details
'Checks the Note column of a dictionary row for the marker text
'"should fail". Rows containing this marker are intentionally invalid
'formulas placed in the fixture and must be excluded from the bulk
'parse-success test.
'@param rowData Variant. Array representing the row values.
'@param noteIndex Long. Index of the Note column within rowData.
'@return Boolean. True when the row should be skipped.
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

'@sub-title Retrieve a specific column value from a dictionary row
'@details
'Safely extracts a string value at the given column index from a row
'array. Returns vbNullString when the index falls outside the array
'bounds, preventing runtime errors on sparse rows.
'@param rowData Variant. Array representing the row values.
'@param columnIndex Long. Index to extract.
'@return String. Column value or vbNullString when the index is out of range.
Private Function RowValue( rowData As Variant, ByVal columnIndex As Long) As String
    If columnIndex >= LBound(rowData) And columnIndex <= UBound(rowData) Then
        RowValue = CStr(rowData(columnIndex))
    End If
End Function

'@sub-title Determine whether a control value represents a formula expression
'@details
'Normalises the control string to lowercase and checks it against the
'known formula control types: "formula", "formulas", "choice_formula",
'"choice_formulas", and "case_when". Used by the bulk dictionary test
'to filter rows that should be parsed.
'@param controlValue String. Value from the Control column.
'@return Boolean. True when the row should be parsed as a formula.
Private Function IsFormulaControl(ByVal controlValue As String) As Boolean
    Dim normalized As String
    normalized = LCase$(controlValue)
    Select Case normalized
        Case "formula", "formulas", "choice_formula", "choice_formulas", "case_when"
            IsFormulaControl = True
    End Select
End Function

'@sub-title Format the expected unknown-token reason using the template
'@details
'Replaces the %1 placeholder in FORMULA_UNKNOWN_TOKEN_TEMPLATE with the
'supplied token string. The result matches the message the Formulas class
'produces for unrecognised tokens.
'@param token String. Token reported as invalid.
'@return String. Reason message matching the Formulas implementation.
Private Function UnknownTokenReason(ByVal token As String) As String
    UnknownTokenReason = Replace(FORMULA_UNKNOWN_TOKEN_TEMPLATE, "%1", token, 1, 1, vbTextCompare)
End Function

'@sub-title Identify two variables sharing the same table and a third from a different table
'@details
'Locates the default criteria, condition, and result variables from the
'dictionary fixture and verifies that criteria and result belong to the
'same table while condition belongs to a different table. This three-
'variable arrangement is required by grouped formula tests. Returns
'False when the fixture data does not satisfy the constraint.
'@param criteriaVar String. ByRef. Populated with the criteria variable name.
'@param conditionVar String. ByRef. Populated with the condition variable name.
'@param resultVar String. ByRef. Populated with the result variable name.
'@param tabName String. ByRef. Populated with the shared table name of criteria and result.
'@return Boolean. True when all three variables satisfy the grouping constraint.
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

'@sub-title Retrieve a variable that belongs to a different table
'@details
'Looks up the default condition variable and confirms its table differs
'from the excluded table name. Used by the table-mismatch rejection
'test to supply a variable that violates the same-table constraint.
'@param excludedTable String. Table name that must not match.
'@param variableName String. ByRef. Populated with the variable name on success.
'@return Boolean. True when a variable from a different table was found.
Private Function VariableFromDifferentTable(ByVal excludedTable As String, _
                                            ByRef variableName As String) As Boolean
    Dim vars As ILLVariables
    Set vars = LLVariables.Create(LinelistDictionary)

    variableName = DEFAULTCONDITIONVAR
    If vars.Value(colName:="table name", varName:=variableName) = excludedTable Then Exit Function
    VariableFromDifferentTable = True
End Function

'@sub-title Build a grouped-reference string matching the production logic
'@details
'Constructs either a structured table reference (prefix + tableName +
'[variableName]) or a direct cell address via LLSheets.VariableAddress,
'depending on the useTableName flag. Mirrors the reference resolution
'used by the Formulas class so tests can compute expected output.
'@param variableName String. The variable to reference.
'@param tableName String. The table that owns the variable.
'@param useTableName Boolean. When True, emit a structured reference.
'@param tablePrefix String. Prefix prepended to the table name for structured references.
'@param sheets ILLSheets. Sheet-address resolver used for cell references.
'@return String. The formatted range reference.
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

'@sub-title Compose the expected SUMIFS formula for grouped parsing assertions
'@details
'Builds the SUMIFS(resultRange, criteriaRange, conditionValue) string
'that the Formulas class is expected to produce for a grouped SUMIFS
'expression. Delegates range construction to GroupedRangeReferenceForTest.
'@param criteriaVar String. Criteria variable name.
'@param conditionVar String. Condition variable name.
'@param resultVar String. Result variable name.
'@param tableName String. Shared table name.
'@param tablePrefix String. Prefix for structured table references.
'@param useTableName Boolean. When True, uses structured references.
'@return String. The expected SUMIFS formula string.
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

'@sub-title Compose the expected COUNTIFS formula with a non-empty criterion
'@details
'Builds the COUNTIFS(criteriaRange, conditionValue, resultRange, "<>")
'string. COUNTIFS grouped formulas append a non-blank criterion on the
'result range, so the expected output includes an extra pair of arguments.
'@param criteriaVar String. Criteria variable name.
'@param conditionVar String. Condition variable name.
'@param resultVar String. Result variable name.
'@param tableName String. Shared table name.
'@param tablePrefix String. Prefix for structured table references.
'@param useTableName Boolean. When True, uses structured references.
'@return String. The expected COUNTIFS formula string.
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

'@sub-title Compose the expected array-style aggregator formula
'@details
'Builds an array-style formula such as AVERAGE(IF(criteriaRange =
'conditionValue, resultRange)) for grouped aggregators that do not have
'a native *IFS Excel function. The aggregator parameter names the outer
'function (e.g. "AVERAGE", "SUM").
'@param aggregator String. Outer aggregator function name.
'@param criteriaVar String. Criteria variable name.
'@param conditionVar String. Condition variable name.
'@param resultVar String. Result variable name.
'@param tableName String. Shared table name.
'@param tablePrefix String. Prefix for structured table references.
'@param useTableName Boolean. When True, uses structured references.
'@return String. The expected array-style grouped formula string.
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


'@section Module Lifecycle
'===============================================================================
'@description Module-level setup and teardown that initialise the assertion
'framework, create per-test fixture worksheets, and release all references
'on completion.

'@sub-title Initialise the assertion framework and set the module name
'@details
'Creates the CustomTest assertion object, suppresses screen updates via
'BusyApp, and registers "TestFormulas" as the current module so that
'test results are labelled correctly.
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestFormulas"
End Sub

'@sub-title Print results and release all module-level references
'@details
'Flushes accumulated assertion results to the output sheet, deletes the
'formula and dictionary fixture worksheets, restores the Excel application
'state via RestoreApp, and sets every module-level object to Nothing.
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

'@sub-title Create fresh fixture worksheets and shared dependencies before each test
'@details
'Suppresses UI updates, builds the formula fixture worksheet containing
'the functions and ASCII character tables, wraps it in a FormulaData
'instance, and prepares the dictionary via PrepareDictionary.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureSheet = PrepareFormulaFixtureSheet(FORMULA_SHEET, FORMULAS_TABLE_NAME, CHARACTERS_TABLE_NAME)
    Set FormulaDataSource = FormulaData.Create(FixtureSheet)
    PrepareDictionary
End Sub

'@sub-title Flush assertions and release per-test references
'@details
'Calls Assert.Flush to write any buffered results and then clears the
'per-test worksheet and object references so the next test starts clean.
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
'@description Public test methods covering validation, parsing, reference
'generation, grouped formulas, error diagnostics, and edge cases of the
'Formulas class.

'@sub-title Verify a single variable is valid in the simple context
'@details
'Arranges a formula containing only a known variable name. Asserts that
'the formula is valid in the "simple" context, that HasSetupVariables is
'True (the parser detected a dictionary variable), that the reason is the
'default success message, and that no diagnostic checking entries exist.
'@TestMethod("Formulas")
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

'@sub-title Verify analysis context rejects a single-variable formula
'@details
'Arranges a formula consisting of only one variable. Asserts that the
'"analysis" context rejects it (analysis formulas must include an
'aggregation function), that the reason matches the single-variable
'message constant, and that a checking entry is logged.
'@TestMethod("Formulas")
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

'@sub-title Verify empty expressions are rejected with the correct reason
'@details
'Arranges a formula with an empty string (vbNullString). Asserts that
'the formula is invalid in the "analysis" context, that the reason
'matches the empty-formula message, and that a diagnostic entry is logged.
'@TestMethod("Formulas")
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

'@sub-title Verify unknown tokens trigger the standard failure message
'@details
'Arranges a formula containing an unrecognised identifier "UNKNOWN_TOKEN".
'Asserts that the formula is invalid in "simple" context, that the reason
'includes the offending token via the template, and that a checking entry
'is recorded.
'@TestMethod("Formulas")
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

'@sub-title Confirm custom aggregators convert to Excel equivalents during analysis parsing
'@details
'Arranges a formula consisting of the custom aggregator "MEAN" and a
'FormulaCondition targeting a single variable. Asserts that the condition
'is valid, that the formula passes analysis validation, and that
'ParsedAnalysisFormula translates "MEAN" to the Excel-native "AVERAGE".
'@TestMethod("Formulas")
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

'@sub-title Verify structured references are applied in linelist formulas
'@details
'Arranges a formula "variable + 5" and calls ParsedLinelistFormula with
'useTableName:=True and tablePrefix:="tbl_". Asserts that the parsed
'output contains the "tbl_" prefix, confirming that the structured
'reference path was used instead of direct cell addresses.
'@TestMethod("Formulas")
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

'@sub-title Verify direct cell addresses when structured references are disabled
'@details
'Arranges a formula "variable + 1" and calls ParsedLinelistFormula with
'useTableName:=False. Resolves the expected cell address via LLSheets.
'Asserts that the parsed output contains the direct cell address and does
'not contain structured reference bracket syntax.
'@TestMethod("Formulas")
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

'@sub-title Validate that every formula-like dictionary entry parses without warnings
'@details
'Iterates all rows in the dictionary fixture, filtering for rows whose
'Control column matches a formula control type. Rows marked "should fail"
'in the Note column or with empty formula text are skipped. For each
'remaining row, builds a Formulas instance and asserts that it is valid
'in the "simple" context with no diagnostics. Finally asserts that at
'least one formula was evaluated to guard against a vacuously passing test.
'@TestMethod("Formulas")
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

'@sub-title Confirm nested parentheses and irregular whitespace are handled correctly
'@details
'Arranges a formula with deeply nested parentheses, mixed spacing, and
'an embedded IF function: "SUM(((var + 2) * (IF(1=1, 3, 4))))". Asserts
'that the formula is valid in analysis context with the standard success
'reason, confirming the tokeniser and parenthesis tracker are robust.
'@TestMethod("Formulas")
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

'@sub-title Verify escaped double-quotes inside string literals are recognised
'@details
'Arranges a formula containing doubled double-quotes within string
'literals (the VBA/Excel quoting convention). Asserts that the formula
'is valid in both analysis and simple contexts and that HasSetupVariables
'is False since the expression contains only literals, not dictionary
'variables.
'@TestMethod("Formulas")
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

'@sub-title Verify boolean literals are accepted in expressions
'@details
'Arranges a formula "IF(TRUE, FALSE, TRUE)" using boolean literal tokens.
'Asserts that the formula is valid in both analysis and simple contexts
'and that HasSetupVariables is False since boolean literals are not
'dictionary variables.
'@TestMethod("Formulas")
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

'@sub-title Verify unknown functions trigger the standard unknown-token failure
'@details
'Arranges a formula wrapping a variable inside "NOTAFUNCTION(...)". Asserts
'that the formula is invalid in analysis context, that the reason matches
'the unknown-token template with "NOTAFUNCTION", and that a diagnostic
'checking entry is logged.
'@TestMethod("Formulas")
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

'@sub-title Verify custom N aggregator does not leave empty parentheses in output
'@details
'Arranges a formula "IF(N()>0, 1, 0)" using the custom N aggregator.
'Asserts that N is invalid in simple context but valid in analysis, that
'the parsed output translates N() to COUNTIFS with structured references,
'and that the result does not contain the empty-parentheses artifact "()("
'which would indicate incomplete substitution.
'@TestMethod("Formulas")
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



'@sub-title Verify unmatched parentheses are detected with descriptive feedback
'@details
'Arranges a formula "SUM((variable + 1" that is missing a closing
'parenthesis. Asserts that the formula is invalid in analysis context,
'that the reason matches the parenthesis-mismatch message, and that a
'diagnostic checking entry is logged.
'@TestMethod("Formulas")
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

'@sub-title Verify a closing parenthesis before any opening parenthesis is detected
'@details
'Arranges a formula ")1" where the closing parenthesis precedes any
'opening one. Asserts that the formula is invalid in analysis context,
'that the reason matches the negative-parenthesis message, and that a
'diagnostic checking entry is logged.
'@TestMethod("Formulas")
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

'@sub-title Verify disallowed characters are rejected
'@details
'Arranges a formula containing the accented character "e-acute" which is
'not present in the allowed ASCII character table. Asserts that the formula
'is invalid in analysis context, that the reason includes the offending
'character via the unknown-token template, and that a checking entry is
'logged.
'@TestMethod("Formulas")
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

'@sub-title Verify grouped SUMIFS expressions emit the native SUMIFS function
'@details
'Arranges a SUMIFS(criteria, condition, result) expression using three
'variables that satisfy the grouping constraint (criteria and result share
'a table, condition is from a different table). Asserts that the formula
'is valid in analysis context, that IsGrouped reports "Yes", and that
'both ParsedLinelistFormula and ParsedAnalysisFormula produce the
'expected SUMIFS(resultRange, criteriaRange, conditionValue) string with
'structured table references.
'@TestMethod("Formulas")
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

'@sub-title Verify grouped COUNTIFS appends a non-empty criterion on the result range
'@details
'Arranges a COUNTIFS(criteria, condition, result) expression using grouped
'variables. Asserts that the formula is valid, that IsGrouped is "Yes",
'and that both parsed outputs match the expected COUNTIFS format which
'includes the extra "<>" criterion on the result range to exclude blanks.
'@TestMethod("Formulas")
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

'@sub-title Verify grouped MEANIFS expressions produce array-style AVERAGE(IF()) formulas
'@details
'Arranges a MEANIFS(criteria, condition, result) expression using grouped
'variables. Asserts validity in both analysis and simple contexts, that
'IsGrouped is "Yes", and that the parsed linelist output uses the
'cell-address form while the parsed analysis output uses structured
'references, both wrapping the result in AVERAGE(IF(...)).
'@TestMethod("Formulas")
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

'@sub-title Validate that generic GROUP_SUM expressions produce SUM(IF()) style formulas
'@details
'Arranges a GROUP_SUM(criteria, condition, result) expression using
'grouped variables. Asserts validity in both simple and analysis contexts,
'that IsGrouped is "Yes", and that both parsed outputs match the expected
'SUM(IF(criteriaRange = conditionValue, resultRange)) format.
'@TestMethod("Formulas")
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

'@sub-title Verify grouped formulas are rejected when the result variable is on a different table
'@details
'Arranges a SUMIFS expression where the third variable (result) belongs
'to a different table than the first variable (criteria). Asserts that the
'formula is invalid in analysis context, that the reason matches the
'table-mismatch message, that IsGrouped is "No", and that a checking
'entry is logged.
'@TestMethod("Formulas")
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

'@sub-title Verify generic grouped formulas without an aggregator are rejected
'@details
'Arranges a "GROUP_(criteria, condition, result)" expression where the
'underscore is followed by an opening parenthesis with no aggregator name.
'Asserts that the formula is invalid in analysis context, that the reason
'matches the invalid-generic-message template, that a checking entry is
'logged, and that IsGrouped is "No".
'@TestMethod("Formulas")
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

'@sub-title Verify generic grouped formulas with an unknown aggregator are rejected
'@details
'Arranges a "GROUP_UNKNOWNFUNC(criteria, condition, result)" expression
'where "UNKNOWNFUNC" is not registered in the Excel function catalog.
'Asserts that the formula is invalid in analysis context, that the reason
'matches the unknown-aggregator template with both the full token and the
'aggregator suffix, that a checking entry is logged, and that IsGrouped
'is "No".
'@TestMethod("Formulas")
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

'@sub-title Confirm ParsedAnalysisFormula respects connectors from IFormulaCondition
'@details
'Arranges a "SUM(variable)" formula with a FormulaCondition containing
'two conditions on the same variable. Calls ParsedAnalysisFormula with
'Connector:=" + ". Asserts that the parsed output contains the " + "
'connector and an IF wrapper, confirming that the condition's connector
'parameter propagates through to the final output.
'@TestMethod("Formulas")
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

'@sub-title Stress test the parser with a large expression
'@details
'Arranges a formula built by concatenating 25 "SUM(variable + N)"
'terms separated by " + ". Asserts that the formula is valid in
'analysis context with the standard success reason, confirming that
'repeated tokenisation across a lengthy expression does not cause
'failures or performance issues.
'@TestMethod("Formulas")
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
