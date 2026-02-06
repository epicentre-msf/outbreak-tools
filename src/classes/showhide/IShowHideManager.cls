VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShowHideManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for show/hide visibility management across worksheet layers"

'@Folder("ShowHide")
'@ModuleDescription("Interface for show/hide visibility management across worksheet layers")
'@Interface

Option Explicit

'@section Enumerations
'===============================================================================

' @description Worksheet layers that can receive show/hide visibility management.
' Values align with legacy ShowHideScope for backward compatibility.
Public Enum ShowHideWorksheetLayer
    ShowHideLayerHList = 1
    ShowHideLayerPrinted = 2
    ShowHideLayerVList = 3
    ShowHideLayerCRF = 4
End Enum


'@section Properties
'===============================================================================

' @description Number of entries managed by this instance.
Public Property Get EntryCount() As Long
End Property

' @description The worksheet layer this manager was built for.
Public Property Get TargetLayer() As ShowHideWorksheetLayer
End Property

' @description The sheet name used to filter dictionary variables during build.
Public Property Get SheetName() As String
End Property


'@section Entry access (1-based index)
'===============================================================================

' @description Variable name (dictionary field key) at the given index.
' @param index 1-based entry index.
' @return String variable name.
Public Function FieldKey(ByVal index As Long) As String
End Function

' @description Display header text at the given index.
' @param index 1-based entry index.
' @return String header suitable for form display.
Public Function HeaderText(ByVal index As Long) As String
End Function

' @description Effective hidden state at the given index.
' Accounts for mandatory, dictionary-hidden and force-hidden flags.
' @param index 1-based entry index.
' @return True when the entry is effectively hidden.
Public Function IsHidden(ByVal index As Long) As Boolean
End Function

' @description True when the entry is mandatory (cannot be hidden by user).
' @param index 1-based entry index.
Public Function IsMandatory(ByVal index As Long) As Boolean
End Function

' @description Column or row position index in the target worksheet.
' @param index 1-based entry index.
' @return Long position; 0 if not resolved.
Public Function PositionIndex(ByVal index As Long) As Long
End Function

' @description Control type string at the given index (e.g. "geo1", "formula").
' @param index 1-based entry index.
Public Function ControlType(ByVal index As Long) As String
End Function


'@section Lookup
'===============================================================================

' @description Find the 1-based index of a field key.
' @param fieldKey Variable name to look up.
' @return Long index; 0 when the field is not found.
Public Function IndexOf(ByVal fieldKey As String) As Long
End Function

' @description True when the field key exists in the manager.
' @param fieldKey Variable name to look up.
Public Function HasField(ByVal fieldKey As String) As Boolean
End Function


'@section Visibility management
'===============================================================================

' @description Set hidden state for the entry at the given index.
' Mandatory entries are silently skipped.
' @param index 1-based entry index.
' @param hidden True to hide, False to show.
Public Sub SetHidden(ByVal index As Long, ByVal hidden As Boolean)
End Sub

' @description Bulk show or hide all non-mandatory, non-locked entries.
' @param hidden True to hide all optional, False to show all optional.
Public Sub SetAllOptionalHidden(ByVal hidden As Boolean)
End Sub


'@section Plan generation
'===============================================================================

' @description Return visibility data for all entries as a BetterArray.
' Each element is a Variant array: Array(headerText, fieldKey, effectiveHidden).
' Suitable for populating form list controls.
' @return BetterArray of Variant arrays (1-based).
Public Function Plan() As BetterArray
End Function


'@section Persistence
'===============================================================================

' @description Export current visibility state to a ListObject.
' Clears existing rows for the target layer before writing.
' The ListObject must have columns: layer, field_key, header_text, hidden_flag.
' @param lo Target ListObject for persistence.
Public Sub ExportPlan(ByVal lo As ListObject)
End Sub

' @description Import persisted visibility from a ListObject, updating hidden states.
' Only updates entries whose field_key exists in the manager.
' @param lo Source ListObject containing persisted data.
Public Sub ImportPlan(ByVal lo As ListObject)
End Sub
