VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTableEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("High-level coordinator driving analysis table generation pipeline")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Public Sub Run(ByVal specificationList As ListObject, _
               ByVal outputSheet As Worksheet, _
               ByVal linelistSpecs As ILinelistSpecs, _
               ByVal tableTypeCode As Byte, _
               ByVal writerContext As IAnalysisTableWriterContext, _
               Optional ByVal sectionLabelPrefix As String = vbNullString, _
               Optional ByVal headerLabelPrefix As String = vbNullString)
End Sub
