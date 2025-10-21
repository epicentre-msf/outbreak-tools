VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGraphExecutionItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Concrete implementation binding a graph plan item to generated chart specs")
'@depends ProjectError, AnalysisArrayUtilities
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGraphExecutionItem

Private Type TExecutionItemState
    planItem As IAnalysisGraphPlanItem
    charts As BetterArray
End Type

Private this As TExecutionItemState

Private Const CLASSNAME As String = "AnalysisGraphExecutionItem"

'@section Factory
'===============================================================================
Public Function Create(ByVal planItem As IAnalysisGraphPlanItem, _
                       ByVal charts As BetterArray) As IAnalysisGraphExecutionItem

    ValidatePlanItem planItem
    ValidateCharts charts

    With New AnalysisGraphExecutionItem
        .Initialise planItem, charts
        Set Create = .Self
    End With
End Function

'Expose the instance as its interface type.
Public Property Get Self() As IAnalysisGraphExecutionItem
    Set Self = Me
End Property

'@section Initialisation
'===============================================================================
Public Sub Initialise(ByVal planItem As IAnalysisGraphPlanItem, _
                      ByVal charts As BetterArray)

    Set this.planItem = planItem
    Set this.charts = CloneCharts(charts)
End Sub

Private Function CloneCharts(ByVal charts As BetterArray) As BetterArray
    Set CloneCharts = AnalysisArrayUtilities.CloneCloneable(charts)
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidatePlanItem(ByVal planItem As IAnalysisGraphPlanItem)
    If planItem Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "planItem cannot be Nothing"
    End If
End Sub

Private Sub ValidateCharts(ByVal charts As BetterArray)
    If charts Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "charts cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisGraphExecutionItem_PlanItem() As IAnalysisGraphPlanItem
    Set IAnalysisGraphExecutionItem_PlanItem = this.planItem
End Property

Private Property Get IAnalysisGraphExecutionItem_Charts() As BetterArray
    Set IAnalysisGraphExecutionItem_Charts = AnalysisArrayUtilities.CloneCloneable(this.charts)
End Property

Private Property Get IAnalysisGraphExecutionItem_ChartCount() As Long
    If this.charts Is Nothing Then
        IAnalysisGraphExecutionItem_ChartCount = 0
    Else
        IAnalysisGraphExecutionItem_ChartCount = this.charts.Length
    End If
End Property
