VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShowHideRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing a show/hide rule for a linelist column"

'@Folder("ShowHideInterface")
'@ModuleDescription("Interface describing a show/hide rule for a linelist column")
'@Interface

Option Explicit

'@section Enumerations
'===============================================================================

'@details Worksheet layers that can host show/hide logic. Values align with legacy ShowHideScope for compatibility.
Public Enum ShowHideWorksheetLayer
    ShowHideLayerHList = 1
    ShowHideLayerCRF = 2
    ShowHideLayerPrinted = 3
    ShowHideLayerVList = 4
    ShowHideLayerDictionary = 5
End Enum


'@section Core properties
'===============================================================================
'@details A rule identifies a column and stores per-layer visibility flags together with dictionary-level constraints.

Public Property Get FieldKey() As String
End Property
Public Property Let FieldKey(ByVal value As String)
End Property

Public Property Get HeaderText() As String
End Property
Public Property Let HeaderText(ByVal value As String)
End Property

Public Property Get Mandatory() As Boolean
End Property
Public Property Let Mandatory(ByVal value As Boolean)
End Property

Public Property Get DictionaryHidden() As Boolean
End Property
Public Property Let DictionaryHidden(ByVal value As Boolean)
End Property

Public Property Get ForceHidden() As Boolean
End Property
Public Property Let ForceHidden(ByVal value As Boolean)
End Property


'@section Layer visibility
'===============================================================================

'@details Store or retrieve user visibility choice for the provided layer.
Public Sub SetLayerVisibility(ByVal layer As ShowHideWorksheetLayer, ByVal hidden As Boolean)
End Sub

'@details Retrieve the stored visibility, without accounting for dictionary or forced flags.
Public Property Get LayerVisibility(ByVal layer As ShowHideWorksheetLayer) As Boolean
End Property

'@details True when SetLayerVisibility was called for the provided layer.
Public Function HasLayerVisibility(ByVal layer As ShowHideWorksheetLayer) As Boolean
End Function

'@details Mark the rule as applicable for a given layer without setting a visibility choice.
Public Sub MarkLayerApplicability(ByVal layer As ShowHideWorksheetLayer, ByVal applies As Boolean)
End Sub

'@details Compute final hidden state after applying dictionary and forced flags.
Public Function EffectiveHidden(ByVal layer As ShowHideWorksheetLayer) As Boolean
End Function

'@details Returns True when the rule should be considered for the provided layer.
Public Function AppliesToLayer(ByVal layer As ShowHideWorksheetLayer) As Boolean
End Function
