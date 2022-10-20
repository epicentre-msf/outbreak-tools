Attribute VB_Name = "DesignerTests"
Option Explicit
Option Private Module

Sub ShowWindows()
    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    EndWork xlsapp:=Application
End Sub


Sub test()
    Dim Dictionary As ILLdictionary
    Dim formData As IFormulaData
    Dim Wksh As Worksheet
    Dim dataWksh As Worksheet
    Dim setupForm As String
    Dim lform As IFormulas
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim parsedFormula As String
    Dim formCond As IFormulaCondition

    'This method runs before every test in the module..
    Set dataWksh = ThisWorkbook.Worksheets("TestDictionary")
    Set Dictionary = LLdictionary.Create(dataWksh, 1, 1)
    Dictionary.Prepare
    
    'Formulas Data
    Set Wksh = ThisWorkbook.Worksheets("ControleFormule")
    Set formData = FormulaData.Create(Wksh, "T_XlsFonctions", "T_ascii")

    setupForm = "N"

    Set lform = Formulas.Create(Dictionary, formData, setupForm)

    Debug.Print lform.Valid()
    Set vars = New BetterArray
    vars.Push "varb1"
    Set conds = New BetterArray
    conds.Push ">0"
    
    Set formCond = FormulaCondition.Create(vars, conds)
    
    'Conditions
    
    'testing a formula
    
    parsedFormula = lform.ParsedAnalysisFormula(formCond)
    Debug.Print parsedFormula

End Sub
