VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


'clear the path ranges once the workbook opens
Private Sub Workbook_Open()
  SheetMain.Range(C_RngPathDic).value = ""
  SheetMain.Range(C_RngPathGeo).value = ""
  SheetMain.Range(C_RngLLName).value = ""
  SheetMain.Range(C_RngLLDir).value = ""
End Sub




