Attribute VB_Name = "LinelistChange"

'This event is called when one sheet of type linelist changes.
'This is for controlling changes related to Geo dropdown control

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    Call EventValueChangeLinelist(Target)
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Activate()
    'Application.EnableEvents = False
    'Call EventOpenLinelist()
    'Application.EnableEvents = True
End Sub