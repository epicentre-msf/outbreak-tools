Attribute VB_Name = "LinelistChange"

'This event is called when one sheet of type linelist changes.
'This is for controlling changes related to Geo dropdown control

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Call EventSheetLineListPatient(Target)
End Sub

