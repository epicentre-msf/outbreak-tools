VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HListSubSectionWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Factory for horizontal subsection writers")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

Public Function Create(ByVal context As ILLSectionContext) As ILLSubSectionWriter
    Dim writer As ILLSubSectionWriter
    Dim layout As ILLSubSectionLayout
    Dim hooks As ILLSubSectionWriterHooks

    Set writer = New LLSubSectionWriterBase
    Set layout = New LLSubSectionLayoutHorizontal
    Set hooks = New LLSubSectionWriterHooksHorizontal

    writer.Initialise context, layout, hooks
    Set Create = writer
End Function

