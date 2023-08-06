Attribute VB_Name = "HlistChange"

'This event is called when one sheet of type linelist changes.
'This is for controlling changes related to Geo dropdown control

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.Cursor = xlNorthwestArrow
    Application.EnableEvents = False
    EventValueChangeLinelist Target
    Application.EnableEvents = True
    Application.Cursor = xlDefault
End Sub

Private Sub Worksheet_Deactivate()
    Application.EnableEvents = False
    EventDesactivateLinelist Me.Name
    Application.EnableEvents = True
End Sub


'@ntncmch here is the code for commenting. Just uncomment the next three lines.

' Private Sub Worksheet_Activate()
'     Me.Calculate
' End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.Cursor = xlNorthwestArrow
    Application.EnableEvents = False
    EventSelectionLinelist Target
    Application.EnableEvents = True
    Application.Cursor = xlDefault
End Sub
