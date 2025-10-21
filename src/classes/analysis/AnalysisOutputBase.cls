VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisOutputBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Provides shared execution flow for rendering analysis plans using output strategies")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Private Const CLASSNAME As String = "AnalysisOutputBase"

'@section Execution entry point
'===============================================================================
Public Sub Execute(ByVal plan As IAnalysisTablePlanResult, _
                   ByVal outputSheet As Worksheet, _
                   ByVal context As IAnalysisTableWriterContext, _
                   ByVal strategy As IAnalysisOutputStrategy)

    ValidatePlan plan
    ValidateOutputSheet outputSheet
    ValidateContext context
    ValidateStrategy strategy

    context.BeginPlan outputSheet
    strategy.Prepare plan, outputSheet, context

    RenderPlanItems plan, outputSheet, context, strategy

    strategy.FinalizePlan plan, outputSheet, context
    context.CompletePlan outputSheet
End Sub

'@section Rendering helpers
'===============================================================================
Private Sub RenderPlanItems(ByVal plan As IAnalysisTablePlanResult, _
                            ByVal outputSheet As Worksheet, _
                            ByVal context As IAnalysisTableWriterContext, _
                            ByVal strategy As IAnalysisOutputStrategy)

    Dim items As BetterArray
    Dim idx As Long
    Dim planItem As IAnalysisTablePlanItem

    Set items = SafePlanItems(plan)
    If items Is Nothing Then Exit Sub
    If items.Length = 0 Then Exit Sub

    For idx = items.LowerBound To items.UpperBound
        Set planItem = items.Item(idx)
        If Not planItem Is Nothing Then
            strategy.RenderItem planItem, outputSheet, context
        End If
    Next idx
End Sub

Private Function SafePlanItems(ByVal plan As IAnalysisTablePlanResult) As BetterArray
    On Error Resume Next
        Set SafePlanItems = plan.Items
    On Error GoTo 0
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidatePlan(ByVal plan As IAnalysisTablePlanResult)
    If plan Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "plan cannot be Nothing"
    End If
End Sub

Private Sub ValidateOutputSheet(ByVal outputSheet As Worksheet)
    If outputSheet Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "outputSheet cannot be Nothing"
    End If
End Sub

Private Sub ValidateContext(ByVal context As IAnalysisTableWriterContext)
    If context Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "context cannot be Nothing"
    End If
End Sub

Private Sub ValidateStrategy(ByVal strategy As IAnalysisOutputStrategy)
    If strategy Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "strategy cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub
