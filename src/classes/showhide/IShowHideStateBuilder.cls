VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShowHideStateBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for constructing show/hide states from dictionary data"

'@Folder("ShowHideInterface")
'@ModuleDescription("Interface for constructing show/hide states from dictionary data")
'@Interface

Option Explicit

'@section Configuration
'===============================================================================

Public Property Set Dictionary(ByVal dict As ILLdictionary)
End Property

Public Property Set Variables(ByVal vars As ILLVariables)
End Property

Public Property Get TargetLayer() As ShowHideWorksheetLayer
End Property

Public Property Let TargetLayer(ByVal layer As ShowHideWorksheetLayer)
End Property

Public Property Get SourceSheetName() As String
End Property

Public Property Let SourceSheetName(ByVal sheetName As String)
End Property

'@section Plan Integration
'===============================================================================

Public Sub ApplyPlan(ByVal plan As IShowHidePlan)
End Sub

'@section Execution
'===============================================================================

Public Function Build() As IShowHideState
End Function

