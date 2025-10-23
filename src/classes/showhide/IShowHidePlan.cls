VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShowHidePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing a set of column visibility actions for a worksheet layer"

'@Folder("ShowHideInterface")
'@ModuleDescription("Interface describing a set of column visibility actions for a worksheet layer")
'@Interface

Option Explicit

Public Property Get TargetLayer() As ShowHideWorksheetLayer
End Property

Public Sub AddVisibility(ByVal fieldKey As String, _
                         ByVal headerText As String, _
                         ByVal hidden As Boolean)
End Sub

Public Function Actions() As BetterArray
End Function

Public Function ContainsField(ByVal fieldKey As String) As Boolean
End Function

Public Function VisibilityFor(ByVal fieldKey As String) As Boolean
End Function
