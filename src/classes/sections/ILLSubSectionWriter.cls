VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSubSectionWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Reusable writer for linelist subsections")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Lifecycle
'===============================================================================
Public Sub Initialise(ByVal sectionContext As ILLSectionContext, _
                      ByVal layout As ILLSubSectionLayout, _
                      ByVal hooks As ILLSubSectionWriterHooks)
End Sub

Public Sub Reset()
End Sub

'@section Operations
'===============================================================================
Public Sub WriteSubSection(ByVal startRow As Long, ByVal endRow As Long)
End Sub

