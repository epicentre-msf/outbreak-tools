VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShowHideState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface exposing the aggregate show/hide state for a linelist"

'@Folder("ShowHideInterface")
'@ModuleDescription("Interface exposing the aggregate show/hide state for a linelist")
'@Interface

Option Explicit

'@section Rule management
'===============================================================================

Public Sub AddRule(ByVal rule As IShowHideRule)
End Sub

Public Function HasRule(ByVal fieldKey As String) As Boolean
End Function

Public Function Rule(ByVal fieldKey As String) As IShowHideRule
End Function

Public Function Rules() As BetterArray
End Function

Public Sub UpdateLayerVisibility(ByVal fieldKey As String, _
                                 ByVal layer As ShowHideWorksheetLayer, _
                                 ByVal hidden As Boolean)
End Sub

Public Sub MarkLayerAsApplicable(ByVal fieldKey As String, _
                                 ByVal layer As ShowHideWorksheetLayer, _
                                 ByVal applies As Boolean)
End Sub

Public Sub SetDictionaryHidden(ByVal fieldKey As String, ByVal hidden As Boolean)
End Sub

Public Sub SetForceHidden(ByVal fieldKey As String, ByVal hidden As Boolean)
End Sub

'@details Compute the final hidden flag for the provided rule/layer combination.
Public Function EvaluateHidden(ByVal fieldKey As String, _
                               ByVal layer As ShowHideWorksheetLayer) As Boolean
End Function

