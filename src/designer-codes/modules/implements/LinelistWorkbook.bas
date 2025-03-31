Attribute VB_Name = "LinelistWorkbook"
Option Explicit

Private Sub Workbook_Open()
    Application.OnKey "^+g", "ClickGeoApp"
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = True
    
    'If you want to add a time stamp to the linelist

    On Error Resume Next
     Application.FormatStaleValues = False
    On Error GoTo 0
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Add functions to move directly in non debug mode
    On Error Resume Next
     Application.FormatStaleValues = True
    On Error GoTo 0
End Sub


'avoid calculation before save to reduce latency
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Application.CalculateBeforeSave = False
End Sub
