Attribute VB_Name = "LinelistWorkbook"
Option Explicit

Private Sub Workbook_Open()
    Application.OnKey "^+g", "ClicCmdGeoApp"
End Sub

Private Sub Workbook_BeforeClose(cancel As Boolean)
    'Add functions to move directly in debug mode
    
End Sub

