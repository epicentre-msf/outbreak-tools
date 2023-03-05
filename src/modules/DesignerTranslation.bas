Attribute VB_Name = "DesignerTranslation"
Option Explicit
Option Private Module

'
Sub TranslateDesignerMain()
    Dim trads As IDesTranslation
    Dim sh As Worksheet
    Dim shmain As Worksheet

    Set sh = ThisWorkbook.Worksheets("DesignerTranslation")
    Set shmain = ThisWorkbook.Worksheets("Main")
    Set trads = DesTranslation.Create(sh)
    
    trads.TranslateDesigner shmain
End Sub

