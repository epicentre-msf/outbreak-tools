Attribute VB_Name = "DesignerTests"
Option Explicit
Option Private Module

Sub ShowWindows()
    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    EndWork xlsapp:=Application
End Sub


Sub testform()
    Dim form As IFormulaCondition
    Dim dictObject As ILLdictionary
    Dim formcond As BetterArray
    Dim formvar As BetterArray
    Dim Wksh As Worksheet
    
    Set formcond = New BetterArray
    Set formvar = New BetterArray
    Set Wksh = ThisWorkbook.Worksheets("Dictionary")
    
    Set dictObject = LLdictionary.Create(Wksh, 1, 1)
    
    formcond.Push "> 0", "< 1", ">0", "<1"
    formvar.Push "varb1", "varb2", "varb3", "varb4"
    
    Set form = FormulaCondition.Create(formcond, formvar)
     
    Debug.Print form.Variable.Length
    Debug.Print form.Valid(dictObject, "table2")
    Debug.Print form.ConditionString("table2", "varb2")
    Debug.Print form.Valid(dictObject, "table3")
    
End Sub
