VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGraphEngineStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisGraphEngine stub capturing Run invocations")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisGraphEngine

Private Type TGraphEngineStubState
    runCount As Long
    lastPlan As IAnalysisTablePlanResult
End Type

Private this As TGraphEngineStubState

Public Property Get RunCount() As Long
    RunCount = this.runCount
End Property

Public Property Get LastPlan() As IAnalysisTablePlanResult
    Set LastPlan = this.lastPlan
End Property

Private Function IAnalysisGraphEngine_Run(ByVal plan As IAnalysisTablePlanResult, ByVal inputs As IAnalysisGraphEngineInputs) As IAnalysisGraphExecutionResult
    this.runCount = this.runCount + 1
    Set this.lastPlan = plan

    Dim result As AnalysisGraphExecutionResult
    Set result = New AnalysisGraphExecutionResult
    Set IAnalysisGraphEngine_Run = result
End Function
