Attribute VB_Name = "DesignerTests"
Option Explicit
Option Private Module

Sub ShowWindows()
    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    EndWork xlsapp:=Application
End Sub


Sub test()
    Dim ana As ILLdictionary
    Dim anash As Worksheet
    Dim trad As ITranslation
    
    Set anash = ThisWorkbook.Worksheets("TestDictionary")
    Set ana = LLdictionary.Create(anash, 1, 1)
    Set trad = Translation.Create(ThisWorkbook.Worksheets("Translations").ListObjects(1), "Français")
    
    ana.Translate trad
    
    
End Sub
