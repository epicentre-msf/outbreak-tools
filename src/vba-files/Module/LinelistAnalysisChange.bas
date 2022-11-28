Attribute VB_Name = "LinelistAnalysisChange"

'This event is called when something changes in the sheet analysis.
'This is for controlling changes related  dropdowns on GoTo sections

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    EventValueChangeAnalysis Target
    Application.EnableEvents = True
End Sub

