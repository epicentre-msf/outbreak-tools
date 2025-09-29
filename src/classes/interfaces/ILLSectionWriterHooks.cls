VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSectionWriterHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Hook interface for section writer customisation")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation
'@InterfaceDescription("Supplies orientation-specific behaviour to LLSectionWriterBase")

Option Explicit

'@section Hook methods
'===============================================================================
Public Sub FormatSectionHeader(ByVal context As ILLSectionContext, _
                               ByVal layout As ILLSectionLayout, _
                               ByVal headerRange As Range, _
                               ByVal printHeaderRange As Range, _
                               ByVal crfHeaderRange As Range)
End Sub

Public Function CreateSubSectionLayout() As ILLSubSectionLayout
End Function

Public Function SubSectionHooks() As ILLSubSectionWriterHooks
End Function
