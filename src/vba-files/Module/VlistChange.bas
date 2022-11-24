Attribute VB_Name = "VlistChange"
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    EventValueChangeVList Target
    Application.EnableEvents = True
End Sub

