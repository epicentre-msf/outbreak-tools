VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSectionWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Interface for reusable linelist section writers")
'@IgnoreModule UseMeaningfulName, HungarianNotation
'
Option Explicit

'@section Lifecycle
'===============================================================================
Public Sub Initialise(ByVal context As ILLSectionContext, _
                      ByVal layout As ILLSectionLayout, _
                      ByVal hooks As ILLSectionWriterHooks)
End Sub

Public Sub Reset()
End Sub

'@section Operations
'===============================================================================
Public Sub WriteSection(ByVal startRow As Long, ByVal endRow As Long)
End Sub

