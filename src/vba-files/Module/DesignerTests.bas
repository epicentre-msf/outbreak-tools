Attribute VB_Name = "DesignerTests"
Option Explicit
Option Private Module

Sub ShowWindows()
    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    EndWork xlsapp:=Application
End Sub


Sub testform()
    Dim form As ILLdictionary
    Dim cond As String
    Dim var As BetterArray

    Dim varData As BetterArray
    Dim condData As BetterArray
    Dim retrData As BetterArray
    
    Set var = New BetterArray
    Set varData = New BetterArray
    Set condData = New BetterArray
    Set retrData = New BetterArray
    
    varData.Push "Sheet Name", "Sub Section"
    condData.Push "A-V1D", "Sub section 1"
    retrData.Push "Variable Name"
    
    Dim Wksh As Worksheet
    Set Wksh = ThisWorkbook.Worksheets("Dictionary")
    
    Set form = LLdictionary.Create(Wksh, 1, 1)
    
    Set var = form.FiltersData(varData, condData, retrData)
    
    Debug.Print form.StartColumn
    Debug.Print form.StartLine
End Sub
