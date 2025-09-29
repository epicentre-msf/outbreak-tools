Attribute VB_Name = "TestFormulaCondition"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const DICT_SHEET As String = "FormulaConditionDict"

Private Assert As Object
Private Dictionary As ILLdictionary

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
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
    Set Dictionary = Nothing
End Sub

'@TestMethod("FormulaCondition")
Private Sub TestCreateRejectsMismatchedLengths()
    Dim vars As BetterArray
    Dim conds As BetterArray

    Set vars = BetterArrayFromList("choi_v1")
    Set conds = BetterArrayFromList("=0", "=1")

    On Error GoTo ExpectError
        Dim form As IFormulaCondition
        Set form = FormulaCondition.Create(vars, conds)
        Assert.Fail "Create should raise for mismatched inputs"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Expected InvalidArgument when arrays lengths differ"
    Err.Clear
End Sub

'@TestMethod("FormulaCondition")
Private Sub TestValidSucceedsForSameTable()
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As IFormulaCondition

    Set vars = BetterArrayFromList("choi_v1", "choi_mult_v1")
    Set conds = BetterArrayFromList(">0", "<5")

    Set form = FormulaCondition.Create(vars, conds)
    Assert.IsTrue form.Valid(Dictionary), "Valid should succeed when variables share a table"
    Assert.AreEqual DictionaryFixtureValue(0, "Table Name"), form.VariablesTable(Dictionary), _
                    "VariablesTable should cache the resolved table"
End Sub

'@TestMethod("FormulaCondition")
Private Sub TestValidLogsWhenVariableMissing()
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
Private Sub TestConditionStringBuildsExpression()
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
Private Sub TestVariablesTableUsesCachedValue()
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim form As IFormulaCondition
    Dim expectedTable As String

    Set vars = BetterArrayFromList("choi_v1", "choi_mult_v1")
    Set conds = BetterArrayFromList(">0", ">1")
    expectedTable = DictionaryFixtureValue(0, "Table Name")

    Set form = FormulaCondition.Create(vars, conds)
    form.Valid Dictionary
    Assert.AreEqual expectedTable, form.VariablesTable(Dictionary), _
                    "VariablesTable should reuse the cached table name after validation"
End Sub
