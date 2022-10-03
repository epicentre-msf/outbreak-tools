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
    Dim cond As String
    Dim var As String
    
    Dim wkb As Workbook
    Set wkb = ThisWorkbook
    
    Set form = FormulaCondition.Create(wkb, "Hello", "Hello")
    
    Debug.Print form.variable
    
    

    
    
End Sub
