VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignerWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit
'clear the path ranges once the workbook opens
Private Sub Workbook_Open()
  SheetMain.Range(C_sRngPathDic).value = ""
  SheetMain.Range(C_sRngPathGeo).value = ""
  SheetMain.Range(C_sRngLLName).value = ""
  SheetMain.Range(C_sRngLLDir).value = ""
  
  Call SetInputRangesToWhite
  SheetMain.Range(C_sRngEdition).value = TranslateMsg(C_sMsgSet)
  
  SheetGeo.ListObjects(C_sTabADM1).DataBodyRange.Clear
  SheetGeo.ListObjects(C_sTabADM2).DataBodyRange.Clear
  SheetGeo.ListObjects(C_sTabADM3).DataBodyRange.Clear
  SheetGeo.ListObjects(C_sTabADM4).DataBodyRange.Clear
  SheetGeo.ListObjects(C_sTabHF).DataBodyRange.Clear
  SheetGeo.ListObjects(C_sTabNames).DataBodyRange.Clear
  
End Sub




