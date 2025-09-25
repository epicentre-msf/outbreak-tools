VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLVarLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Layout strategy describing where linelist variables are written")
'@IgnoreModule UseMeaningfulName, HungarianNotation
'
'@InterfaceDescription("Encapsulates anchor and offset logic for variable placement")

Option Explicit

'@section Lifecycle
'===============================================================================
Public Sub Configure(ByVal baseSheet As Worksheet, _
                     ByVal anchorIndex As Long, _
                     Optional ByVal printedSheet As Worksheet)
End Sub

Public Sub Reset()
End Sub

'@section Layout pointers
'===============================================================================
Public Property Get ValueCell() As Range
End Property

Public Property Get PrintedValueCell() As Range
End Property

Public Property Get LabelCell() As Range
End Property

Public Property Get SublabelTarget() As Range
End Property

Public Property Get NameCell() As Range
End Property

Public Property Get NoteTarget() As Range
End Property

Public Property Get ControlCell() As Range
End Property

Public Property Get AutoOriginCell() As Range
End Property

Public Property Get AnchorCell() As Range
End Property

Public Property Get HiddenRange() As Range
End Property

'@section Metadata
'===============================================================================
Public Function SupportsPrintedLayout() As Boolean
End Function

Public Property Get LayoutKey() As String
End Property

