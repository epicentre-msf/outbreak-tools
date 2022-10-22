Attribute VB_Name = "DesignerTests"
Option Explicit
Option Private Module

Sub ShowWindows()
    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    EndWork xlsapp:=Application
End Sub


Sub test()
    Dim Dictionary As ILLchoice
    Dim formData As IFormulaData
    Dim Wksh As Worksheet
    Dim dataWksh As Worksheet
    Dim setupForm As String
    Dim lform As IFormulas
    Dim vars As BetterArray
    Dim conds As BetterArray
    Dim parsedFormula As String
    Dim formCond As IFormulaCondition
    Dim Wkb As Workbook

    'This method runs before every test in the module..
    Set dataWksh = ThisWorkbook.Worksheets("TestChoices")
    Set Dictionary = LLchoice.Create(dataWksh, 1, 1)
    
    Set Wkb = Workbooks.Add
    Dictionary.Export Wkb
    

End Sub
