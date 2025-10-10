VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistWorkbookReleaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Allows lifecycle helpers to clear linelist workbook references once they are closed"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Allows lifecycle helpers to clear linelist workbook references once they are closed")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, SuperfluousAnnotationArgument

Option Explicit

'@section Workbook release
'===============================================================================
'@sub-title Clears the cached workbook instance so a fresh output workbook can be created afterwards.
Public Sub ClearOutputWorkbook()
End Sub
