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

  Call ResetField
  SheetMain.Range(C_sRngEdition).value = TranslateMsg(C_sMsgSet)

  SheetGeo.ListObjects(C_sTabadm1).DataBodyRange.Clear
  SheetGeo.ListObjects(C_sTabAdm2).DataBodyRange.Clear
  SheetGeo.ListObjects(C_sTabAdm3).DataBodyRange.Clear
  SheetGeo.ListObjects(C_sTabAdm4).DataBodyRange.Clear
  SheetGeo.ListObjects(C_sTabHF).DataBodyRange.Clear
  SheetGeo.ListObjects(C_sTabNames).DataBodyRange.Clear

  SheetLLTranslation.Range(C_sRngLLLanguage).value = SheetMain.Range(C_sRngLLFormLang).value 'check Language of linelist's forms

  With SheetDesTranslation
        .Range(.Cells(.Range("T_Lst_Lang").Row, .Range("T_Lst_Lang").Column), _
               .Cells(.Range("T_Lst_Lang").Row, .Range("T_Lst_Lang").End(xlToRight).Column)).ClearContents
  End With

End Sub



























































































