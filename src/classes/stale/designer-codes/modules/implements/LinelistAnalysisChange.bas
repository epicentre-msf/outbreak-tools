Attribute VB_Name = "LinelistAnalysisChange"
Option Explicit

'This event is called when something changes in the sheet analysis.
'This is for controlling changes related  dropdowns on GoTo sections

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    EventValueChangeAnalysis Target
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Application.EnableEvents = False
    EventDoubleClickAnalysis Target
    Application.EnableEvents = True
End Sub