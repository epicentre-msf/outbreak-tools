VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGraphPlanBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Analyses table plans to identify and categorise graph generation work")
'@depends ProjectError, AnalysisGraphPlanItem, AnalysisGraphPlanResult
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGraphPlanBuilder

Private Const CLASSNAME As String = "AnalysisGraphPlanBuilder"

'@section Factory helpers
'===============================================================================
Private Function CreateBuffer() As BetterArray
    Dim buffer As BetterArray
    Set buffer = New BetterArray
    buffer.LowerBound = 1
    Set CreateBuffer = buffer
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidatePlan(ByVal plan As IAnalysisTablePlanResult)
    If plan Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "plan cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Build logic
'===============================================================================
Private Sub PopulateGraphCollections(ByVal planItems As BetterArray, _
                                     ByVal allGraphs As BetterArray, _
                                     ByVal crossTableGraphs As BetterArray, _
                                     ByVal timeSeriesGraphs As BetterArray)
    Dim idx As Long
    Dim planItem As IAnalysisTablePlanItem
    Dim policyResult As IAnalysisTablePolicyResult
    Dim graphItem As IAnalysisGraphPlanItem

    If planItems Is Nothing Then Exit Sub
    If planItems.Length = 0 Then Exit Sub

    For idx = planItems.LowerBound To planItems.UpperBound
        Set planItem = planItems.Item(idx)
        If planItem Is Nothing Then GoTo ContinueLoop

        Set policyResult = planItem.PolicyResult
        If policyResult Is Nothing Then GoTo ContinueLoop
        If Not policyResult.HasGraph Then GoTo ContinueLoop

        Set graphItem = AnalysisGraphPlanItem.Create(planItem)
        allGraphs.Push graphItem

        If graphItem.IsTimeSeries Then
            timeSeriesGraphs.Push graphItem
        Else
            crossTableGraphs.Push graphItem
        End If

ContinueLoop:
        Set planItem = Nothing
        Set policyResult = Nothing
        Set graphItem = Nothing
    Next idx
End Sub

'@section Interface implementation
'===============================================================================
Private Function IAnalysisGraphPlanBuilder_Build(ByVal plan As IAnalysisTablePlanResult) As IAnalysisGraphPlanResult
    Dim planItems As BetterArray
    Dim allGraphs As BetterArray
    Dim crossTableGraphs As BetterArray
    Dim timeSeriesGraphs As BetterArray

    ValidatePlan plan

    Set allGraphs = CreateBuffer()
    Set crossTableGraphs = CreateBuffer()
    Set timeSeriesGraphs = CreateBuffer()

    Set planItems = plan.Items
    PopulateGraphCollections planItems, allGraphs, crossTableGraphs, timeSeriesGraphs

    Set IAnalysisGraphPlanBuilder_Build = AnalysisGraphPlanResult.Create(allGraphs, crossTableGraphs, timeSeriesGraphs)
End Function
