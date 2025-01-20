VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of export class"

'@Folder("Dictionary")
'@ModuleDescription("Interface of export class")
'@Interface

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

Public Sub ExportSpecs(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden): End Sub
Public Sub ImportSpecs(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long): End Sub
'Add and remove rows for the export
Public Property Get NumberOfExports() As Long: End Property
Public Sub AddRows(): End Sub
Public Sub RemoveRows(): End Sub
