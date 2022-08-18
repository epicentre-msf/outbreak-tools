Attribute VB_Name = "DesignerTests"
Option Explicit
Option Private Module

Sub ShowWindows()
    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    EndWork xlsapp:=Application
End Sub

