VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShowHideOrchestrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface coordinating show/hide state building, strategies, and persistence"

'@Folder("ShowHideInterface")
'@ModuleDescription("Interface coordinating show/hide state building, strategies, and persistence")
'@Interface

Option Explicit

Public Sub RegisterStrategy(ByVal strategy As IShowHideStrategy)
End Sub

Public Function HasStrategy(ByVal layer As ShowHideWorksheetLayer) As Boolean
End Function

Public Function Apply(ByVal sheetName As String, _
                      ByVal layer As ShowHideWorksheetLayer, _
                      ByVal exportService As IShowHideExport, _
                      Optional ByVal table As ICustomTable) As IShowHidePlan
End Function
