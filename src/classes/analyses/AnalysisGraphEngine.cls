VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGraphEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Runs graph planning and orchestrator execution for analysis outputs")
'@depends ProjectError, AnalysisGraphPlanBuilder, AnalysisGraphExecutionResult, AnalysisGraphExecutionItem, AnalysisGraphOrchestratorFactory
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGraphEngine

Private Type TGraphEngineState
    planBuilder As IAnalysisGraphPlanBuilder
    orchestratorFactory As IAnalysisGraphOrchestratorFactory
End Type

Private this As TGraphEngineState

Private Const CLASSNAME As String = "AnalysisGraphEngine"

'@section Properties
'===============================================================================
Public Property Set PlanBuilder(ByVal builder As IAnalysisGraphPlanBuilder)
    ValidatePlanBuilder builder
    Set this.planBuilder = builder
End Property

Private Property Get PlanBuilderInternal() As IAnalysisGraphPlanBuilder
    If this.planBuilder Is Nothing Then
        Set this.planBuilder = New AnalysisGraphPlanBuilder
    End If
    Set PlanBuilderInternal = this.planBuilder
End Property

Public Property Set OrchestratorFactory(ByVal factory As IAnalysisGraphOrchestratorFactory)
    ValidateOrchestratorFactory factory
    Set this.orchestratorFactory = factory
End Property

Private Property Get OrchestratorFactoryInternal() As IAnalysisGraphOrchestratorFactory
    If this.orchestratorFactory Is Nothing Then
        Set this.orchestratorFactory = New AnalysisGraphOrchestratorFactory
    End If
    Set OrchestratorFactoryInternal = this.orchestratorFactory
End Property

'@section Validation helpers
'===============================================================================
Private Sub ValidatePlan(ByVal plan As IAnalysisTablePlanResult)
    If plan Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "plan cannot be Nothing"
    End If
End Sub

Private Sub ValidateInputs(ByVal inputs As IAnalysisGraphEngineInputs)
    If inputs Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "inputs cannot be Nothing"
    End If
End Sub

Private Sub ValidatePlanBuilder(ByVal builder As IAnalysisGraphPlanBuilder)
    If builder Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "builder cannot be Nothing"
    End If
End Sub

Private Sub ValidateOrchestratorFactory(ByVal factory As IAnalysisGraphOrchestratorFactory)
    If factory Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "factory cannot be Nothing"
    End If
End Sub

Private Sub ValidateCrossTable(ByVal crossTable As ICrossTable, ByVal planItem As IAnalysisGraphPlanItem)
    If crossTable Is Nothing Then
        RaiseError CLng(ProjectError.ErrorUnexpectedState), _
                   "Cross table resources are missing for plan item " & SafePlanIdentifier(planItem)
    End If
End Sub

Private Function SafePlanIdentifier(ByVal planItem As IAnalysisGraphPlanItem) As String
    On Error Resume Next
        SafePlanIdentifier = Trim$(CStr(planItem.Specification.TableId))
    On Error GoTo 0
End Function

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Build helpers
'===============================================================================
Private Function BuildGraphPlan(ByVal plan As IAnalysisTablePlanResult) As IAnalysisGraphPlanResult
    Set BuildGraphPlan = PlanBuilderInternal.Build(plan)
End Function

Private Sub CollectCrossTableGraphs(ByVal graphPlan As IAnalysisGraphPlanResult, _
                                    ByVal inputs As IAnalysisGraphEngineInputs, _
                                    ByVal result As AnalysisGraphExecutionResult)

    Dim items As BetterArray
    Dim idx As Long
    Dim planItem As IAnalysisGraphPlanItem
    Dim crossTable As ICrossTable
    Dim orchestrator As IGraphSpecsOrchestrator
    Dim charts As BetterArray
    Dim executionItem As IAnalysisGraphExecutionItem

    Set items = graphPlan.CrossTableItems
    If items Is Nothing Then Exit Sub
    If items.Length = 0 Then Exit Sub

    For idx = items.LowerBound To items.UpperBound
        Set planItem = items.Item(idx)
        If planItem Is Nothing Then GoTo ContinueLoop

        Set crossTable = inputs.CrossTableFor(planItem)
        ValidateCrossTable crossTable, planItem

        Set orchestrator = OrchestratorFactoryInternal.CreateForCrossTable(crossTable)
        If orchestrator Is Nothing Then GoTo ContinueLoop

        Set charts = orchestrator.Build
        If charts Is Nothing Then GoTo ContinueLoop
        If charts.Length = 0 Then GoTo ContinueLoop

        Set executionItem = AnalysisGraphExecutionItem.Create(planItem, charts)
        result.AddCrossTableItem executionItem

ContinueLoop:
        Set planItem = Nothing
        Set crossTable = Nothing
        Set orchestrator = Nothing
        Set charts = Nothing
        Set executionItem = Nothing
    Next idx
End Sub

Private Sub CollectTimeSeriesGraphs(ByVal graphPlan As IAnalysisGraphPlanResult, _
                                    ByVal inputs As IAnalysisGraphEngineInputs, _
                                    ByVal result As AnalysisGraphExecutionResult)

    Dim orchestrator As IGraphSpecsOrchestrator
    Dim charts As BetterArray

    If Not graphPlan.HasTimeSeriesGraphs Then Exit Sub

    Set orchestrator = OrchestratorFactoryInternal.CreateForTimeSeries( _
                            inputs.GraphList, _
                            inputs.TimeSeriesList, _
                            inputs.TitleList, _
                            inputs.Linelist)

    If orchestrator Is Nothing Then Exit Sub

    Set charts = orchestrator.Build
    If charts Is Nothing Then Exit Sub
    If charts.Length = 0 Then Exit Sub

    result.SetTimeSeriesCharts charts
End Sub

'@section Interface implementation
'===============================================================================
Private Function IAnalysisGraphEngine_Run(ByVal plan As IAnalysisTablePlanResult, _
                                          ByVal inputs As IAnalysisGraphEngineInputs) As IAnalysisGraphExecutionResult
    Dim graphPlan As IAnalysisGraphPlanResult
    Dim result As AnalysisGraphExecutionResult

    ValidatePlan plan
    ValidateInputs inputs

    Set graphPlan = BuildGraphPlan(plan)
    Set result = New AnalysisGraphExecutionResult

    CollectCrossTableGraphs graphPlan, inputs, result
    CollectTimeSeriesGraphs graphPlan, inputs, result

    Set IAnalysisGraphEngine_Run = result
End Function
