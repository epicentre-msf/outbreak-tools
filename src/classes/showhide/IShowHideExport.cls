VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShowHideExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for importing and exporting show/hide plans to workbook tables"

'@Folder("ShowHideInterface")
'@ModuleDescription("Interface for importing and exporting show/hide plans to workbook tables")
'@Interface

Option Explicit

'@section Configuration
'===============================================================================

Public Sub BindTable(ByVal table As ICustomTable)
End Sub

Public Property Get BoundTable() As ICustomTable
End Property

'@section Persistence
'===============================================================================

Public Sub ExportPlan(ByVal plan As IShowHidePlan)
End Sub

Public Function ImportPlan(ByVal layer As ShowHideWorksheetLayer) As IShowHidePlan
End Function

Public Sub ClearLayer(ByVal layer As ShowHideWorksheetLayer)
End Sub

