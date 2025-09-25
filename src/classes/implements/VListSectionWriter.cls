VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VListSectionWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Factory for vertical section writers")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

Public Function Create(ByVal context As ILLSectionContext) As ILLSectionWriter
    Dim writer As ILLSectionWriter
    Dim layout As ILLSectionLayout
    Dim hooks As ILLSectionWriterHooks

    Set writer = New LLSectionWriterBase
    Set layout = New LLSectionLayoutVertical
    Set hooks = New LLSectionWriterHooksVertical

    writer.Initialise context, layout, hooks
    Set Create = writer
End Function

