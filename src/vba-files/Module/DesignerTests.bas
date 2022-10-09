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
    Dim var As BetterArray
    
    
    Dim varData As BetterArray
    Dim condData As BetterArray
    Dim retrData As BetterArray

    Set varData = New BetterArray
    Set condData = New BetterArray
    Set retrData = New BetterArray
    Set var = New BetterArray
    

    'Found values
    varData.Push "sheet name"
    condData.Push "A, B, C"
    retrData.Push "variable name", "sheet type"
    
    Set formcond = New BetterArray
    Set formvar = New BetterArray
    Set Wksh = ThisWorkbook.Worksheets("Dictionary")
    
    Set dictObject = LLdictionary.Create(Wksh, 1, 1)
    
    dictObject.Clean
    dictObject.Prepare
    
    formcond.Push "> 0", "< 1", ">0", "<1"
    formvar.Push "varb1", "varb2", "varb3", "varb4"
    
    Set var = dictObject.FilterData("export 1", "<>", "__all__", includeHeaders:=True)
    
    Set form = FormulaCondition.Create(formcond, formvar)
    
    
    Debug.Print form.Variable.Length
    Debug.Print form.Valid(dictObject, "table2")
    Debug.Print form.ConditionString("table2", "varb2")
    Debug.Print form.Valid(dictObject, "table3")
    
End Sub
