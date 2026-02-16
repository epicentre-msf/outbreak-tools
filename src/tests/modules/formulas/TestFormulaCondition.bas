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
'@depends FormulaCondition, IFormulaCondition, LLdictionary, ILLdictionary,
'  LLVariables, ILLVariables, BetterArray, CustomTest, ICustomTest,
'  DictionaryTestFixture, TestHelpers

Private Const DICT_SHEET As String = "FormulaConditionDict"

Private Assert As ICustomTest
Private Dictionary As ILLdictionary

'@section Helpers
'===============================================================================

'@sub-title Resolve the table name for a given variable through the dictionary
Private Function TableNameFor(ByVal variableName As String) As String
    Dim vars As ILLVariables
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
'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestFormulaCondition"
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
End Sub

'@sub-title Print results and tear down the dictionary fixture
'@details
'Flushes remaining assertion output to the test sheet, deletes the
'dictionary fixture worksheet, and releases object references.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet DICT_SHEET
    Set Dictionary = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Re-seed the dictionary fixture before each test
'@details
'Recreates the dictionary worksheet and prepares it via LLdictionary.Prepare
'so that each test starts from a known clean state with prepared metadata.
'@TestInitialize
Private Sub TestInitialize()
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    Dictionary.Prepare
End Sub

'@sub-title Flush assertions and release the dictionary after each test
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set Dictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@sub-title Verify Create rejects variables and conditions arrays of different lengths
'@details
'Arranges a single-element variables array and a two-element conditions array,
'then calls FormulaCondition.Create. Asserts that an InvalidArgument error is
'raised, confirming the factory guard clause prevents mismatched inputs.
'@TestMethod("FormulaCondition")
Public Sub TestCreateRejectsMismatchedLengths()
    CustomTestSetTitles Assert, "FormulaCondition", "TestCreateRejectsMismatchedLengths"
    Dim vars As BetterArray
    Dim conds As BetterArray

    Set vars = BetterArrayFromList("choi_v1")
    Set conds = BetterArrayFromList("=0", "=1")

    On Error GoTo ExpectError
        Dim form As IFormulaCondition
        '@Ignored AssigmentNotUsed
        Set form = FormulaCondition.Create(vars, conds)
        Assert.LogFailure "Create should raise for mismatched inputs"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Expected InvalidArgument when arrays lengths differ"
    Err.Clear
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
    Dim form As IFormulaCondition

    Set vars = BetterArrayFromList("choi_v1", "choi_mult_v1")
    Set conds = BetterArrayFromList(">0", "<5")

    Set form = FormulaCondition.Create(vars, conds)
    Assert.IsTrue form.Valid(Dictionary), "Valid should succeed when variables share a table"
    Assert.IsFalse form.HasCheckings, "Matching tables should not record diagnostics"
    Assert.AreEqual TableNameFor("choi_v1"), form.VariablesTable(Dictionary), _
                    "VariablesTable should cache the resolved table"
End Sub

'@sub-title Verify validation fails and logs diagnostics when a variable is missing
'@details
'Uses one valid variable and one that does not exist in the dictionary.
'Asserts Valid returns False, HasCheckings returns True, and the
'CheckingValues object is available (not Nothing) for diagnostic inspection.
'@TestMethod("FormulaCondition")
Public Sub TestValidLogsWhenVariableMissing()
    CustomTestSetTitles Assert, "FormulaCondition", "TestValidLogsWhenVariableMissing"
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As IFormulaCondition

    Set vars = BetterArrayFromList("choi_v1", "missing_var")
    Set conds = BetterArrayFromList(">0", ">1")

    Set form = FormulaCondition.Create(vars, conds)

    Assert.IsFalse form.Valid(Dictionary), "Valid should return False when variables are missing"
    Assert.IsTrue form.HasCheckings, "Validation failures should produce checkings"
    Assert.IsTrue Not form.CheckingValues Is Nothing, "Checking log should be available after failure"
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
    Dim form As IFormulaCondition
    Dim predicate As String

    Set vars = BetterArrayFromList("choi_v1", "choi_mult_v1")
    Set conds = BetterArrayFromList(">0", ">1")

    Set form = FormulaCondition.Create(vars, conds)

    predicate = form.ConditionPredicate("DataTable", Connector:="*")
    Assert.AreEqual "(DataTable[choi_v1]>0)*(DataTable[choi_mult_v1]>1)", predicate, _
                    "ConditionPredicate should join clauses with the provided connector"

    Assert.AreEqual "IF((DataTable[choi_v1]>0)*(DataTable[choi_mult_v1]>1) , DataTable[result])", _
                 form.ConditionString("DataTable", "result", Connector:="*"), _
                 "ConditionString should wrap the predicate in an IF expression"
End Sub

'@sub-title Verify VariablesTable returns the cached value after a prior Valid call
'@details
'Creates a FormulaCondition, explicitly calls Valid to populate the cache,
'then asserts that VariablesTable returns the same resolved table name
'without requiring a second validation pass.
'@TestMethod("FormulaCondition")
Public Sub TestVariablesTableUsesCachedValue()
    CustomTestSetTitles Assert, "FormulaCondition", "TestVariablesTableUsesCachedValue"
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As IFormulaCondition
    Dim expectedTable As String

    Set vars = BetterArrayFromList("choi_v1", "choi_mult_v1")
    Set conds = BetterArrayFromList(">0", ">1")
    expectedTable = TableNameFor("choi_v1")

    Set form = FormulaCondition.Create(vars, conds)
    form.Valid Dictionary
    Assert.AreEqual expectedTable, form.VariablesTable(Dictionary), _
                    "VariablesTable should reuse the cached table name after validation"
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
    Dim form As IFormulaCondition
    Dim firstTable As String
    Dim secondTable As String

    Set vars = BetterArrayFromList("choi_v1", "cond_test_h1")
    Set conds = BetterArrayFromList(">0", ">1")

    firstTable = TableNameFor("choi_v1")
    secondTable = TableNameFor("cond_test_h1")
    Assert.IsFalse (StrComp(firstTable, secondTable, vbTextCompare) = 0), _
                   "Fixture assumption broken: expected variables from different tables"

    Set form = FormulaCondition.Create(vars, conds)

    Assert.IsFalse form.Valid(Dictionary), "Valid should fail when variables belong to different tables"
    Assert.IsTrue form.HasCheckings, "Cross-table validation failure should log diagnostics"
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
    Dim form As IFormulaCondition
    Dim expectedTable As String
    Dim wrongTable As String

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
End Sub
