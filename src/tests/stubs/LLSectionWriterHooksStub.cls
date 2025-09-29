VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSectionWriterHooksStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests")
'@ModuleDescription("Test double for section writer hooks")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLSectionWriterHooks
Implements ILLSubSectionWriterHooks

Public HeaderCallCount As Long
Public SubSectionCallCount As Long
Public LastHeaderAddress As String
Public SubSectionLayoutRequests As Long
Public VariableWriteCount As Long

Private Sub ILLSectionWriterHooks_FormatSectionHeader(ByVal context As ILLSectionContext, _
                                                      ByVal layout As ILLSectionLayout, _
                                                      ByVal headerRange As Range, _
                                                      ByVal printHeaderRange As Range, _
                                                      ByVal crfHeaderRange As Range)

    HeaderCallCount = HeaderCallCount + 1
    If Not headerRange Is Nothing Then
        LastHeaderAddress = headerRange.Address(False, False)
    End If
End Sub

Private Function ILLSectionWriterHooks_CreateSubSectionLayout() As ILLSubSectionLayout
    SubSectionLayoutRequests = SubSectionLayoutRequests + 1
    Set ILLSectionWriterHooks_CreateSubSectionLayout = New LLSubSectionLayoutHorizontal
End Function

Private Function ILLSectionWriterHooks_SubSectionHooks() As ILLSubSectionWriterHooks
    Set ILLSectionWriterHooks_SubSectionHooks = Me
End Function

Private Sub ILLSubSectionWriterHooks_FormatHeader(ByVal context As ILLSectionContext, _
                                                  ByVal layout As ILLSubSectionLayout, _
                                                  ByVal headerRange As Range, _
                                                  ByVal printHeaderRange As Range, _
                                                  ByVal crfHeaderRange As Range)
    SubSectionCallCount = SubSectionCallCount + 1
End Sub

Private Function ILLSubSectionWriterHooks_CreateVariableWriter(ByVal context As ILLVarContext) As ILLVarWriter
    Dim writer As ILLVarWriter
    Set writer = New LLVarWriterStub
    writer.Initialise context, New LLVarLayoutHorizontal, New LLVarWriterHooksHorizontal
    Set ILLSubSectionWriterHooks_CreateVariableWriter = writer
    VariableWriteCount = VariableWriteCount + 1
End Function

Private Sub ILLSubSectionWriterHooks_AfterVariableWrite(ByVal context As ILLSectionContext, _
                                                        ByVal layout As ILLSubSectionLayout, _
                                                        ByVal headerRange As Range)
    'No-op
End Sub
