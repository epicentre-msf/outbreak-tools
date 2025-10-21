VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableOutputStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Renders table plans via IAnalysisTableWriterContext using GoTo stack navigation")
'@depends ProjectError, AnalysisGoToStackBuilder
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisOutputStrategy

Private Type TStrategyState
    stackBuilder As IAnalysisGoToStackBuilder
End Type

Private this As TStrategyState
Private Const CLASSNAME As String = "AnalysisTableOutputStrategy"

'@section Configuration
'===============================================================================
Public Property Set StackBuilder(ByVal builder As IAnalysisGoToStackBuilder)
    Set this.stackBuilder = builder
End Property

Private Property Get StackBuilderInternal() As IAnalysisGoToStackBuilder
    Set StackBuilderInternal = this.stackBuilder
End Property

Private Function ResolveStackBuilder() As IAnalysisGoToStackBuilder
    Dim builder As IAnalysisGoToStackBuilder
    Set builder = StackBuilderInternal

    If builder Is Nothing Then
        Set builder = AnalysisGoToStackBuilder.Create
        Set this.stackBuilder = builder
    End If

    Set ResolveStackBuilder = builder
End Function

'@section Strategy helpers
'===============================================================================
Private Sub ApplyPrefixes(ByVal plan As IAnalysisTablePlanResult)
    Dim builder As IAnalysisGoToStackBuilder

    Set builder = ResolveStackBuilder()
    builder.SectionPrefix = plan.SectionPrefix
    builder.HeaderPrefix = plan.HeaderPrefix
End Sub

Private Sub ApplyNavigation(ByVal plan As IAnalysisTablePlanResult, _
                            ByVal outputSheet As Worksheet, _
                            ByVal context As IAnalysisTableWriterContext)

    Dim builder As IAnalysisGoToStackBuilder
    Dim stacks As AnalysisGoToStackContext

    Set builder = ResolveStackBuilder()
    ApplyPrefixes plan

    Set stacks = builder.Build(plan)
    context.ApplyNavigation stacks.SectionLabels, stacks.HeaderLabels, outputSheet
End Sub

Private Sub ValidatePlanItem(ByVal planItem As IAnalysisTablePlanItem)
    If planItem Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "planItem cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Sub IAnalysisOutputStrategy_Prepare(ByVal plan As IAnalysisTablePlanResult, _
                                            ByVal outputSheet As Worksheet, _
                                            ByVal context As IAnalysisTableWriterContext)
    ApplyPrefixes plan
End Sub

Private Sub IAnalysisOutputStrategy_RenderItem(ByVal planItem As IAnalysisTablePlanItem, _
                                               ByVal outputSheet As Worksheet, _
                                               ByVal context As IAnalysisTableWriterContext)
    ValidatePlanItem planItem
    context.WriteTable planItem, outputSheet
End Sub

Private Sub IAnalysisOutputStrategy_FinalizePlan(ByVal plan As IAnalysisTablePlanResult, _
                                                 ByVal outputSheet As Worksheet, _
                                                 ByVal context As IAnalysisTableWriterContext)
    ApplyNavigation plan, outputSheet, context
End Sub
