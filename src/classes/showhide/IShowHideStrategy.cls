VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShowHideStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Strategy interface generating visibility plans for a specific worksheet layer"

'@Folder("ShowHideInterface")
'@ModuleDescription("Strategy interface generating visibility plans for a specific worksheet layer")
'@Interface

Option Explicit

Public Property Get TargetLayer() As ShowHideWorksheetLayer
End Property

Public Property Get StrategyName() As String
End Property

Public Function BuildPlan(ByVal state As IShowHideState, _
                          ByVal table As ICustomTable) As IShowHidePlan
End Function
