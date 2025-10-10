Attribute VB_Name = "TestFormulaCondition"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@Folder("CustomTests")
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const DICT_SHEET As String = "FormulaConditionDict"

Private Assert As ICustomTest
Private Dictionary As ILLdictionary

Private Function TableNameFor(ByVal variableName As String) As String
    Dim vars As ILLVariables
    Set vars = LLVariables.Create(Dictionary)
    TableNameFor = vars.TableName(variableName)
End Function

'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestFormulaCondition"
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet DICT_SHEET
    Set Dictionary = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set Dictionary = Nothing
End Sub

'@TestMethod("FormulaCondition")
Public Sub TestCreateRejectsMismatchedLengths()
    CustomTestSetTitles Assert, "FormulaCondition", "TestCreateRejectsMismatchedLengths"
    Dim vars As BetterArray
    Dim conds As BetterArray

    Set vars = BetterArrayFromList("choi_v1")
    Set conds = BetterArrayFromList("=0", "=1")

    On Error GoTo ExpectError
        Dim form As IFormulaCondition
        Set form = FormulaCondition.Create(vars, conds)
        Assert.LogFailure "Create should raise for mismatched inputs"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Expected InvalidArgument when arrays lengths differ"
    Err.Clear
End Sub

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
