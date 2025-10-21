VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableWriterStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisTableWriter stub capturing the plan passed for execution")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisTableWriter

Public WriteCount As Long
Public LastPlan As IAnalysisTablePlanResult
Public LastOutputSheet As Worksheet
Public LastContext As IAnalysisTableWriterContext

Private Sub IAnalysisTableWriter_Write(ByVal plan As IAnalysisTablePlanResult, _
                                       ByVal outputSheet As Worksheet, _
                                       ByVal context As IAnalysisTableWriterContext)
    WriteCount = WriteCount + 1
    Set LastPlan = plan
    Set LastOutputSheet = outputSheet
    Set LastContext = context
End Sub
