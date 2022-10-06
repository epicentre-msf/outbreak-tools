Attribute VB_Name = "TestControlFormula"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private formCond As IFormulaCondition
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

    var.Push "varb1", "varb2", "varb3", "varb4"
    cond.Push " > 0", " < 0", " > 1", " < 1"

    Set formCond = FormulaCondition.Create(var, cond)
End Sub


'@TestMethod
Private Sub TestFormInit()

    Assert.IsTrue (formCond.Variable.Length = 4), "Not all the variables are initilialized in conditions"
    Assert.IsTrue (formCond.Cond.Length = formCond.Variable.Length), "Conditions and variable length do not match"

End Sub


'@TestMethod
Private Sub TestFormulaValidity()
    Dim Wksh As Worksheet
    Dim Dict As ILLdictionary

    Set Wksh = ThisWorkbook.Worksheets("Dictionary")
    Dict = LLdictionary.Create(Wksh, 1, 1)

    Assert.IsTrue formCond.Valid(Dict, "table2"), "Correct formula shows as incorrect (variable length = 4)"
    Assert.IsFalse formCond.Valid(Dict, "table3"), "Formula with false table name shows as correct"
    var.Pop
    formCond = FormulaCondition.Create(var, cond)
    Assert.IsFalse formCond.Valid(Dict, "table2"), "Formula with variable length < condition length shows as correct"
    cond.Pop
    cond.Pop
    formCond = FormulaCondition.Create(var, cond)
    Assert.IsFalse formCond.Valid(Dict, "table2"), "Formula with variable length > condition length shows as correct"
    var.pop
    formCond = FormulaCondition.Create(var, cond)
    Assert.IsTrue formCond.Valid(Dict, "table2"), "Correct formula shows as incorrect (variable length = 2)"
End Sub

'@TestMethod
Private Sub TestFormConversion

    Assert.IsTrue (formCond.ConditionString("table2", "varb2") = "IF((table2[varb1] > 0)*(table2[varb2] < 0) , table2[varb2])"), "Formula not converted correctly (step 1)"
     Assert.IsTrue (formCond.ConditionString("filttable2", "varb5") = "IF((filttable2[varb1] > 0)*(filttable2[varb2] < 0) , filttable2[varb5])"), "Formula not converted correctly (step 2)"

End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
