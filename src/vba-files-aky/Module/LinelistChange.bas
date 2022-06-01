Attribute VB_Name = "LinelistChange"

'This event is called when one sheet of type linelist changes.
'This is for controlling changes related to Geo dropdown control

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    Application.Cursor = xlNorthwestArrow
    Call EventValueChangeLinelist(Target)
    Application.EnableEvents = True
    Application.Cursor = xlDefault
End Sub

Private Sub Worksheet_Deactivate()
    Application.EnableEvents = False
    Dim sSheetName as String
    sSheetName = Me.Name
    Call EventDesactivateLinelist(sSheetName)
    Application.EnableEvents = True
End Sub


