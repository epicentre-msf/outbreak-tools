VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSubSectionWriterHooksVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Hooks for vertical subsection writer")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLSubSectionWriterHooks

Private Sub ILLSubSectionWriterHooks_AfterVariableWrite(ByVal context As ILLSectionContext, _
                                                        ByVal layout As ILLSubSectionLayout, _
                                                        ByVal headerRange As Range)
    'No additional behaviour required
End Sub

Private Function ILLSubSectionWriterHooks_CreateVariableWriter(ByVal context As ILLVarContext) As ILLVarWriter
    Dim writer As ILLVarWriter
    Set writer = New LLVarWriterBase
    writer.Initialise context, New LLVarLayoutVertical, New LLVarWriterHooksVertical
    Set ILLSubSectionWriterHooks_CreateVariableWriter = writer
End Function

Private Sub ILLSubSectionWriterHooks_FormatHeader(ByVal context As ILLSectionContext, _
                                                  ByVal layout As ILLSubSectionLayout, _
                                                  ByVal headerRange As Range, _
                                                  ByVal printHeaderRange As Range, _
                                                  ByVal crfHeaderRange As Range)

    Dim design As ILLFormat

    Set design = context.Design
    If design Is Nothing Then Exit Sub

    On Error Resume Next
        If Not headerRange Is Nothing Then design.ApplyFormat headerRange, scope:=FormatScopeVListSubSection
    On Error GoTo 0
End Sub

