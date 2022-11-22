Attribute VB_Name = "DesignerTranslation"
Option Explicit
Option Private Module

'
Sub TranslateDesignerMain()
    Dim Trads As IDesTranslation
    Dim sh As Worksheet
    Dim shmain As Worksheet

    Set sh = ThisWorkbook.Worksheets("DesignerTranslation")
    Set shmain = ThisWorkbook.Worksheets("Main")
    Set Trads = DesTranslation.Create(sh)
    
    Trads.TranslateDesigner shmain
End Sub
