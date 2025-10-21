VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSectionWriterHooksVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Vertical section writer hooks")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLSectionWriterHooks

'@section ILLSectionWriterHooks implementation
'===============================================================================
Private Sub ILLSectionWriterHooks_FormatSectionHeader(ByVal context As ILLSectionContext, _
                                                      ByVal layout As ILLSectionLayout, _
                                                      ByVal headerRange As Range, _
                                                      ByVal printHeaderRange As Range, _
                                                      ByVal crfHeaderRange As Range)

    Dim design As ILLFormat

    Set design = context.Design
    If design Is Nothing Then Exit Sub

    On Error Resume Next
        If Not headerRange Is Nothing Then design.ApplyFormat headerRange, scope:=VListSection
    On Error GoTo 0
End Sub

Private Function ILLSectionWriterHooks_CreateSubSectionLayout() As ILLSubSectionLayout
    Set ILLSectionWriterHooks_CreateSubSectionLayout = New LLSubSectionLayoutVertical
End Function

Private Function ILLSectionWriterHooks_SubSectionHooks() As ILLSubSectionWriterHooks
    Set ILLSectionWriterHooks_SubSectionHooks = New LLSubSectionWriterHooksVertical
End Function
