VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSectionLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Layout strategy for linelist sections")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation
'@InterfaceDescription("Provides header and body ranges for linelist sections")

Option Explicit

'@section Lifecycle
'===============================================================================
Public Sub Configure(ByVal context As ILLSectionContext, _
                     ByVal startRow As Long, _
                     ByVal endRow As Long)
End Sub

Public Sub Reset()
End Sub

'@section Layout projections
'===============================================================================
Public Property Get HeaderRange() As Range
End Property

Public Property Get PrintHeaderRange() As Range
End Property

Public Property Get CrfHeaderRange() As Range
End Property

Public Property Get LayoutKey() As String
End Property

