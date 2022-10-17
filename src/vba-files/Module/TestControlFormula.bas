Attribute VB_Name = "TestControlFormula"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private formcond As IFormulaCondition
Private var As BetterArray
Private cond As BetterArray

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    Set var = New BetterArray
    Set cond = New BetterArray
    

    var.Push "varb1", "varb2", "varb3", "varb4"
    cond.Push " > 0", " < 0", " > 1", " < 1"

    Set formcond = FormulaCondition.Create(var, cond)
End Sub


'@TestMethod
Private Sub TestFormInit()

    Assert.IsTrue (formcond.Variable.Length = 4), "Not all the variables are initilialized in conditions"
    Assert.IsTrue (formcond.Condition.Length = formcond.Variable.Length), "Conditions and variable length do not match"

End Sub


'@TestMethod
Private Sub TestFormulaValidity()
    Dim Wksh As Worksheet
    Dim dict As ILLdictionary

    Set Wksh = ThisWorkbook.Worksheets("Dictionary")
    Set dict = LLdictionary.Create(Wksh, 1, 1)
    
    Assert.IsTrue (Not dict.ColumnExists("table name")) Or formcond.Valid(dict, "tab2"), "Correct formula shows as incorrect (variable length = 4)"
    Assert.IsFalse formcond.Valid(dict, "tab3"), "Formula with false table name shows as correct"
    var.Pop
    Set formcond = FormulaCondition.Create(var, cond)
    Assert.IsFalse formcond.Valid(dict, "tab2"), "Formula with variable length < condition length shows as correct"
    cond.Pop
    cond.Pop
    Set formcond = FormulaCondition.Create(var, cond)
    Assert.IsFalse formcond.Valid(dict, "tab2"), "Formula with variable length > condition length shows as correct"
    var.Pop
    Set formcond = FormulaCondition.Create(var, cond)
    Assert.IsTrue (Not dict.ColumnExists("table name")) Or formcond.Valid(dict, "tab2"), "Correct formula shows as incorrect (variable length = 2)"
End Sub

'@TestMethod
Private Sub TestFormConversion()
    var.Pop
    var.Pop
    cond.Pop
    cond.Pop
    Set formcond = FormulaCondition.Create(var, cond)
    Assert.IsTrue (formcond.ConditionString("tab2", "varb2") = "IF((tab2[varb1] > 0)*(tab2[varb2] < 0) , tab2[varb2])"), "Formula not converted correctly (step 1)"
    Assert.IsTrue (formcond.ConditionString("filttable2", "varb5") = "IF((filttable2[varb1] > 0)*(filttable2[varb2] < 0) , filttable2[varb5])"), "Formula not converted correctly (step 2)"

End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set var = Nothing
    Set cond = Nothing
    Set formcond = Nothing
End Sub
