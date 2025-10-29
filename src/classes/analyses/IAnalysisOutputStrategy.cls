VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisOutputStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Defines hooks for rendering plan results to outputs via AnalysisOutputBase")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Public Sub Prepare(ByVal plan As IAnalysisTablePlanResult, _
                   ByVal outputSheet As Worksheet, _
                   ByVal context As IAnalysisTableWriterContext)
End Sub

Public Sub RenderItem(ByVal planItem As IAnalysisTablePlanItem, _
                      ByVal outputSheet As Worksheet, _
                      ByVal context As IAnalysisTableWriterContext)
End Sub

Public Sub FinalizePlan(ByVal plan As IAnalysisTablePlanResult, _
                        ByVal outputSheet As Worksheet, _
                        ByVal context As IAnalysisTableWriterContext)
End Sub
