VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Export Linelist Data")
'@IgnoreModule

Option Explicit

Public Enum ExportScope
    ExportOne = 1
    ExportTwo = 2
    ExportThree = 3
    ExportFour = 4
    ExportFive = 5
    ExportAll = 10
End Enum

'Exposed methods

Public Sub Translate(ByVal trans As ITranslation): End Sub
Public Sub ExportSpecs(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden): End Sub

Public Sub ImportSpecs(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long): End Sub
