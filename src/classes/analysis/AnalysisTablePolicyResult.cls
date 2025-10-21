VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTablePolicyResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Immutable value object storing table-type policy evaluation results")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTablePolicyResult

Private Type TPolicyResultState
    iterationItem As IAnalysisTableIterationItem
    policy As ITableTypePolicy
    context As ITableSpecsPolicyContext
    isValid As Boolean
    hasPercent As Boolean
    hasTotal As Boolean
    hasMissing As Boolean
    hasGraph As Boolean
End Type

Private this As TPolicyResultState
Private Const CLASSNAME As String = "AnalysisTablePolicyResult"

'@section Factory
'===============================================================================
'Create a populated policy result capturing evaluation flags.
Public Function Create(ByVal iterationItem As IAnalysisTableIterationItem, _
                       ByVal policy As ITableTypePolicy, _
                       ByVal context As ITableSpecsPolicyContext, _
                       ByVal isValid As Boolean, _
                       ByVal hasPercent As Boolean, _
                       ByVal hasTotal As Boolean, _
                       ByVal hasMissing As Boolean, _
                       ByVal hasGraph As Boolean) As IAnalysisTablePolicyResult

    ValidateIterationItem iterationItem
    ValidatePolicy policy
    ValidateContext context

    With New AnalysisTablePolicyResult
        Set .IterationItem = iterationItem
        Set .Policy = policy
        Set .Context = context
        .IsValid = isValid
        .HasPercent = hasPercent
        .HasTotal = hasTotal
        .HasMissing = hasMissing
        .HasGraph = hasGraph
        Set Create = .Self
    End With
End Function

'Expose the current instance as its interface.
Public Property Get Self() As IAnalysisTablePolicyResult
    Set Self = Me
End Property

'@section Properties
'===============================================================================
Public Property Set IterationItem(ByVal iterationItem As IAnalysisTableIterationItem)
    ValidateIterationItem iterationItem
    Set this.iterationItem = iterationItem
End Property

Private Property Get IterationItemInternal() As IAnalysisTableIterationItem
    Set IterationItemInternal = this.iterationItem
End Property

Public Property Set Policy(ByVal policy As ITableTypePolicy)
    ValidatePolicy policy
    Set this.policy = policy
End Property

Private Property Get PolicyInternal() As ITableTypePolicy
    Set PolicyInternal = this.policy
End Property

Public Property Set Context(ByVal context As ITableSpecsPolicyContext)
    ValidateContext context
    Set this.context = context
End Property

Private Property Get ContextInternal() As ITableSpecsPolicyContext
    Set ContextInternal = this.context
End Property

Public Property Let IsValid(ByVal isValid As Boolean)
    this.isValid = isValid
End Property

Private Property Get IsValidInternal() As Boolean
    IsValidInternal = this.isValid
End Property

Public Property Let HasPercent(ByVal hasPercent As Boolean)
    this.hasPercent = hasPercent
End Property

Private Property Get HasPercentInternal() As Boolean
    HasPercentInternal = this.hasPercent
End Property

Public Property Let HasTotal(ByVal hasTotal As Boolean)
    this.hasTotal = hasTotal
End Property

Private Property Get HasTotalInternal() As Boolean
    HasTotalInternal = this.hasTotal
End Property

Public Property Let HasMissing(ByVal hasMissing As Boolean)
    this.hasMissing = hasMissing
End Property

Private Property Get HasMissingInternal() As Boolean
    HasMissingInternal = this.hasMissing
End Property

Public Property Let HasGraph(ByVal hasGraph As Boolean)
    this.hasGraph = hasGraph
End Property

Private Property Get HasGraphInternal() As Boolean
    HasGraphInternal = this.hasGraph
End Property

'@section Validation helpers
'===============================================================================
Private Sub ValidateIterationItem(ByVal iterationItem As IAnalysisTableIterationItem)
    If iterationItem Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "iterationItem cannot be Nothing"
    End If
End Sub

Private Sub ValidatePolicy(ByVal policy As ITableTypePolicy)
    If policy Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "policy cannot be Nothing"
    End If
End Sub

Private Sub ValidateContext(ByVal context As ITableSpecsPolicyContext)
    If context Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "context cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisTablePolicyResult_IterationItem() As IAnalysisTableIterationItem
    Set IAnalysisTablePolicyResult_IterationItem = IterationItemInternal()
End Property

Private Property Get IAnalysisTablePolicyResult_Policy() As ITableTypePolicy
    Set IAnalysisTablePolicyResult_Policy = PolicyInternal()
End Property

Private Property Get IAnalysisTablePolicyResult_Context() As ITableSpecsPolicyContext
    Set IAnalysisTablePolicyResult_Context = ContextInternal()
End Property

Private Property Get IAnalysisTablePolicyResult_IsValid() As Boolean
    IAnalysisTablePolicyResult_IsValid = IsValidInternal()
End Property

Private Property Get IAnalysisTablePolicyResult_HasPercent() As Boolean
    IAnalysisTablePolicyResult_HasPercent = HasPercentInternal()
End Property

Private Property Get IAnalysisTablePolicyResult_HasTotal() As Boolean
    IAnalysisTablePolicyResult_HasTotal = HasTotalInternal()
End Property

Private Property Get IAnalysisTablePolicyResult_HasMissing() As Boolean
    IAnalysisTablePolicyResult_HasMissing = HasMissingInternal()
End Property

Private Property Get IAnalysisTablePolicyResult_HasGraph() As Boolean
    IAnalysisTablePolicyResult_HasGraph = HasGraphInternal()
End Property
