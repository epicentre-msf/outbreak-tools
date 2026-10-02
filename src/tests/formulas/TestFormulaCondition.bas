Attribute VB_Name = "TestFormulaCondition"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@ModuleDescription("Verifies FormulaCondition creation, validation, and predicate rendering")
'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'Tests the FormulaCondition class, which pairs variable names with predicate
'fragments and renders structured Excel expressions for analysis formulas.
'Coverage includes factory guard clauses (mismatched array lengths), dictionary
'validation (same table, different tables, missing variables, table override),
'predicate rendering (ConditionPredicate and ConditionString), and the cached
'VariablesTable accessor. Each test builds lightweight BetterArray fixtures
'via BetterArrayFromList and a shared dictionary fixture seeded from
'DictionaryTestFixture.
'
'Every test arms its handler above the arrange. An error that escapes a test
'proc reaches the VBE as a modal dialog, the dialog blocks the Apple Event that
'drives the run, and the whole suite comes back with no results file.
'@depends FormulaCondition, LLdictionary, LLdictionary,
'  LLVariables, BetterArray, CustomTest,
'  DictionaryTestFixture, TestHelpersLite

Private Const DICT_SHEET As String = "FormulaConditionDict"
Private Const OTHER_DICT_SHEET As String = "FormulaConditionDict2"

Private Assert As CustomTest
Private Dictionary As LLdictionary

'@section Helpers
'===============================================================================

'@sub-title Resolve the table name for a given variable through the dictionary
Private Function TableNameFor(ByVal variableName As String) As String
    Dim vars As LLVariables
    Set vars = LLVariables.Create(Dictionary)
    TableNameFor = vars.TableName(variableName)
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Initialise the test harness and seed the shared dictionary fixture
'@details
'Creates the test output sheet, sets up the CustomTest assertion object,
'seeds a dictionary worksheet via PrepareDictionaryFixture, and wraps it
'in an LLdictionary instance used by all tests.
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
    Assert.SetModuleName "TestFormulaCondition"
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "ModuleInitialize", Err.Number, Err.Description
End Sub

'@sub-title Print results and tear down the dictionary fixture
'@details
'Flushes remaining assertion output to the test sheet, deletes the
'dictionary fixture worksheet, and releases object references.
'Every step before RestoreApp is wrapped, because RestoreApp has to run
'whatever happened above: the hooks here call BusyApp, which puts Excel in
'manual calculation, and the next module in the run would inherit that.
'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        DeleteWorksheet DICT_SHEET
        DeleteWorksheet OTHER_DICT_SHEET
    On Error GoTo 0

    RestoreApp

    Set Dictionary = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Re-seed the dictionary fixture before each test
'@details
'Recreates the dictionary worksheet and prepares it via LLdictionary.Prepare
'so that each test starts from a known clean state with prepared metadata.
'@TestInitialize
Public Sub TestInitialize()
    On Error GoTo Fail
    BusyApp
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    Dictionary.Prepare
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInitialize", Err.Number, Err.Description
End Sub

'@sub-title Flush assertions and release the dictionary after each test
'@TestCleanup
Public Sub TestCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.Flush
        End If
    On Error GoTo 0

    Set Dictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@sub-title Verify Create rejects variables and conditions arrays of different lengths
'@details
'Arranges a single-element variables array and a two-element conditions array,
'then calls FormulaCondition.Create. Asserts that an InvalidArgument error is
'raised, confirming the factory guard clause prevents mismatched inputs.
'The arrange runs under its own handler, and the act then switches to the
'handler that expects the raise.
'@TestMethod("FormulaCondition")
Public Sub TestCreateRejectsMismatchedLengths()
    CustomTestSetTitles Assert, "FormulaCondition", "TestCreateRejectsMismatchedLengths"
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As FormulaCondition

    On Error GoTo Fail
    Set vars = BetterArrayFromList("choi_v1")
    Set conds = BetterArrayFromList("=0", "=1")

    On Error GoTo ExpectError
        '@Ignored AssigmentNotUsed
        Set form = FormulaCondition.Create(vars, conds)
        Assert.LogFailure "Create should raise for mismatched inputs"
        Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Expected InvalidArgument when arrays lengths differ"
    Err.Clear
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateRejectsMismatchedLengths", Err.Number, Err.Description
End Sub

'@sub-title Verify validation succeeds when all variables belong to the same table
'@details
'Creates a FormulaCondition with two variables from the same dictionary table
'and two condition fragments. Asserts Valid returns True, HasCheckings returns
'False (no diagnostics), and VariablesTable returns the expected table name
'resolved from the first variable.
'@TestMethod("FormulaCondition")
Public Sub TestValidSucceedsForSameTable()
    CustomTestSetTitles Assert, "FormulaCondition", "TestValidSucceedsForSameTable"
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As FormulaCondition

    On Error GoTo Fail
    Set vars = BetterArrayFromList("choi_v1", "choi_mult_v1")
    Set conds = BetterArrayFromList(">0", "<5")

    Set form = FormulaCondition.Create(vars, conds)
    Assert.IsTrue form.Valid(Dictionary), "Valid should succeed when variables share a table"
    Assert.IsFalse form.HasCheckings, "Matching tables should not record diagnostics"
    Assert.AreEqual TableNameFor("choi_v1"), form.VariablesTable(Dictionary), _
                    "VariablesTable should cache the resolved table"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValidSucceedsForSameTable", Err.Number, Err.Description
End Sub

'@sub-title Verify validation fails and logs diagnostics when a variable is missing
'@details
'Uses one valid variable and one that does not exist in the dictionary.
'Asserts Valid returns False, HasCheckings returns True, and the
'CheckingValues object is available for diagnostic inspection.
'@TestMethod("FormulaCondition")
Public Sub TestValidLogsWhenVariableMissing()
    CustomTestSetTitles Assert, "FormulaCondition", "TestValidLogsWhenVariableMissing"
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As FormulaCondition

    On Error GoTo Fail
    Set vars = BetterArrayFromList("choi_v1", "missing_var")
    Set conds = BetterArrayFromList(">0", ">1")

    Set form = FormulaCondition.Create(vars, conds)

    Assert.IsFalse form.Valid(Dictionary), "Valid should return False when variables are missing"
    Assert.IsTrue form.HasCheckings, "Validation failures should produce checkings"
    Assert.IsTrue Not form.CheckingValues Is Nothing, "Checking log should be available after failure"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValidLogsWhenVariableMissing", Err.Number, Err.Description
End Sub

'@sub-title Verify ConditionPredicate and ConditionString render correct Excel expressions
'@details
'Creates a FormulaCondition from two same-table variables with conditions
'">0" and ">1", then calls ConditionPredicate with a "*" connector and
'asserts the joined predicate string. Also calls ConditionString with a
'"result" column and asserts the IF-wrapped expression is correctly formed.
'@TestMethod("FormulaCondition")
Public Sub TestConditionStringBuildsExpression()
    CustomTestSetTitles Assert, "FormulaCondition", "TestConditionStringBuildsExpression"
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As FormulaCondition
    Dim predicate As String

    On Error GoTo Fail
    Set vars = BetterArrayFromList("choi_v1", "choi_mult_v1")
    Set conds = BetterArrayFromList(">0", ">1")

    Set form = FormulaCondition.Create(vars, conds)

    predicate = form.ConditionPredicate("DataTable", Connector:="*")
    Assert.AreEqual "(DataTable[choi_v1]>0)*(DataTable[choi_mult_v1]>1)", predicate, _
                    "ConditionPredicate should join clauses with the provided connector"

    Assert.AreEqual "IF((DataTable[choi_v1]>0)*(DataTable[choi_mult_v1]>1) , DataTable[result])", _
                 form.ConditionString("DataTable", "result", Connector:="*"), _
                 "ConditionString should wrap the predicate in an IF expression"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestConditionStringBuildsExpression", Err.Number, Err.Description
End Sub

'@sub-title Verify VariablesTable returns the cached value after a prior Valid call
'@details
'Creates a FormulaCondition, explicitly calls Valid to populate the cache,
'then asserts that VariablesTable returns the same resolved table name
'from a single validation pass.
'@TestMethod("FormulaCondition")
Public Sub TestVariablesTableUsesCachedValue()
    CustomTestSetTitles Assert, "FormulaCondition", "TestVariablesTableUsesCachedValue"
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As FormulaCondition
    Dim expectedTable As String

    On Error GoTo Fail
    Set vars = BetterArrayFromList("choi_v1", "choi_mult_v1")
    Set conds = BetterArrayFromList(">0", ">1")
    expectedTable = TableNameFor("choi_v1")

    Set form = FormulaCondition.Create(vars, conds)
    form.Valid Dictionary
    Assert.AreEqual expectedTable, form.VariablesTable(Dictionary), _
                    "VariablesTable should reuse the cached table name after validation"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestVariablesTableUsesCachedValue", Err.Number, Err.Description
End Sub

'@sub-title Verify validation fails when variables belong to different tables
'@details
'Arranges two variables known to belong to different dictionary tables
'(confirmed by a fixture assumption assertion). Creates a FormulaCondition
'and asserts that Valid returns False and HasCheckings returns True,
'confirming cross-table usage is rejected.
'@TestMethod("FormulaCondition")
Public Sub TestValidFailsForDifferentTables()
    CustomTestSetTitles Assert, "FormulaCondition", "TestValidFailsForDifferentTables"
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As FormulaCondition
    Dim firstTable As String
    Dim secondTable As String

    On Error GoTo Fail
    Set vars = BetterArrayFromList("choi_v1", "cond_test_h1")
    Set conds = BetterArrayFromList(">0", ">1")

    firstTable = TableNameFor("choi_v1")
    secondTable = TableNameFor("cond_test_h1")
    Assert.IsFalse (StrComp(firstTable, secondTable, vbTextCompare) = 0), _
                   "Fixture assumption broken: expected variables from different tables"

    Set form = FormulaCondition.Create(vars, conds)

    Assert.IsFalse form.Valid(Dictionary), "Valid should fail when variables belong to different tables"
    Assert.IsTrue form.HasCheckings, "Cross-table validation failure should log diagnostics"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValidFailsForDifferentTables", Err.Number, Err.Description
End Sub

'@sub-title Verify a cached validity answer belongs to the dictionary it was computed against
'@details
'Validates against the shared dictionary, then validates the same instance
'against a SECOND dictionary whose fixture no longer holds one of the
'variables. Both calls pass the same (empty) table name, so a cache keyed on
'the table name alone hands the first dictionary's True answer back for the
'second one. Asserts the second call answers for its own dictionary.
'@TestMethod("FormulaCondition")
Public Sub TestValidCacheBelongsToOneDictionary()
    CustomTestSetTitles Assert, "FormulaCondition", "TestValidCacheBelongsToOneDictionary"
    On Error GoTo Fail

    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As FormulaCondition
    Dim otherDictionary As LLdictionary
    Dim otherVars As LLVariables
    Dim renamedCell As Range

    Set vars = BetterArrayFromList("choi_v1", "choi_mult_v1")
    Set conds = BetterArrayFromList(">0", "<5")

    PrepareDictionaryFixture OTHER_DICT_SHEET
    Set otherDictionary = LLdictionary.Create(ThisWorkbook.Worksheets(OTHER_DICT_SHEET), 1, 1)
    otherDictionary.Prepare

    Set otherVars = LLVariables.Create(otherDictionary)
    Set renamedCell = otherVars.CellRange("variable name", "choi_v1")
    Assert.IsTrue Not renamedCell Is Nothing, "Fixture assumption broken: choi_v1 should exist on the second sheet"
    renamedCell.Value = "renamed_choi_v1"
    otherVars.InvalidateCaches

    Set form = FormulaCondition.Create(vars, conds)

    Assert.IsTrue form.Valid(Dictionary), "The first dictionary holds both variables"
    Assert.IsFalse form.Valid(otherDictionary), _
                   "The second dictionary is missing a variable, so it answers False"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValidCacheBelongsToOneDictionary", Err.Number, Err.Description
End Sub

'@sub-title Verify the optional table override parameter of Valid
'@details
'Creates a FormulaCondition with same-table variables, then validates with
'an incorrect override table name and asserts failure with diagnostics.
'Next validates with the correct override table and asserts success,
'cleared diagnostics, and the expected VariablesTable cache value.
'@TestMethod("FormulaCondition")
Public Sub TestValidRespectsTableOverride()
    CustomTestSetTitles Assert, "FormulaCondition", "TestValidRespectsTableOverride"
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As FormulaCondition
    Dim expectedTable As String
    Dim wrongTable As String

    On Error GoTo Fail
    Set vars = BetterArrayFromList("choi_v1", "choi_mult_v1")
    Set conds = BetterArrayFromList(">=1", "<=5")
    expectedTable = TableNameFor("choi_v1")
    wrongTable = TableNameFor("cond_test_h1")
    Assert.IsFalse (StrComp(expectedTable, wrongTable, vbTextCompare) = 0), _
                   "Fixture assumption broken: override table should differ from expected"

    Set form = FormulaCondition.Create(vars, conds)

    Assert.IsFalse form.Valid(Dictionary, wrongTable), _
                   "Supplying an incorrect override table should fail validation"
    Assert.IsTrue form.HasCheckings, "Incorrect override should record diagnostics"

    Assert.IsTrue form.Valid(Dictionary, expectedTable), _
                  "Providing the matching override table should allow validation"
    Assert.IsFalse form.HasCheckings, "Successful validation should clear previous diagnostics"
    Assert.AreEqual expectedTable, form.VariablesTable(Dictionary), _
                    "VariablesTable should return the override value once validation succeeds"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValidRespectsTableOverride", Err.Number, Err.Description
End Sub
