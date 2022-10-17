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
    Dim sheets As ILLSheets
    
    Dim dataWksh As Worksheet
    'This method runs before every test in the module..
    Set dataWksh = ThisWorkbook.Worksheets("Dictionary")
    Set Dictionary = LLdictionary.Create(dataWksh, 1, 1)
    Dictionary.Prepare
    Set sheets = LLSheets.Create(Dictionary)
    
    
End Sub
