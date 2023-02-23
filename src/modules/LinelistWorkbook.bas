Attribute VB_Name = "LinelistWorkbook"
Option Explicit

Private Sub Workbook_Open()
    Application.OnKey "^+g", "ClicCmdGeoApp"
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Add functions to move directly in non debug mode
    
End Sub


'avoid calculation before save to reduce latency
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Application.CalculateBeforeSave = False
End Sub
