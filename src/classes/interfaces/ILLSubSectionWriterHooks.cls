VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSubSectionWriterHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Hook contract for subsection writers")
'@IgnoreModule UseMeaningfulName, HungarianNotation
'@InterfaceDescription("Provides orientation-specific behaviour to LLSubSectionWriterBase")

Option Explicit

'@section Writer hooks
'===============================================================================
Public Sub FormatHeader(ByVal context As ILLSectionContext, _
                        ByVal layout As ILLSubSectionLayout, _
                        ByVal headerRange As Range, _
                        ByVal printHeaderRange As Range, _
                        ByVal crfHeaderRange As Range)
End Sub

Public Function CreateVariableWriter(ByVal context As ILLVarContext) As ILLVarWriter
End Function

Public Sub AfterVariableWrite(ByVal context As ILLSectionContext, _
                              ByVal layout As ILLSubSectionLayout, _
                              ByVal headerRange As Range)
End Sub

