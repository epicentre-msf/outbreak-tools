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
    SheetDesTranslation.Range("LangDictList").ClearContents

End Sub




