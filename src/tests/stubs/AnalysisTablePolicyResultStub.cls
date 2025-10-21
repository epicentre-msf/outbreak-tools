VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTablePolicyResultStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisTablePolicyResult stub exposing configurable flags")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisTablePolicyResult

Private Type TPolicyResultStubState
    iterationItem As IAnalysisTableIterationItem
    policy As ITableTypePolicy
    context As ITableSpecsPolicyContext
    isValid As Boolean
    hasPercent As Boolean
    hasTotal As Boolean
    hasMissing As Boolean
    hasGraph As Boolean
End Type

Private this As TPolicyResultStubState

Public Sub Configure(ByVal iterationItem As IAnalysisTableIterationItem, _
                     ByVal isValid As Boolean)
    Set this.iterationItem = iterationItem
    this.isValid = isValid
End Sub

Public Sub SetPolicy(ByVal policy As ITableTypePolicy)
    Set this.policy = policy
End Sub

Public Sub SetContext(ByVal context As ITableSpecsPolicyContext)
    Set this.context = context
End Sub

Public Sub SetFlags(ByVal hasPercent As Boolean, ByVal hasTotal As Boolean, _
                    ByVal hasMissing As Boolean, ByVal hasGraph As Boolean)
    this.hasPercent = hasPercent
    this.hasTotal = hasTotal
    this.hasMissing = hasMissing
    this.hasGraph = hasGraph
End Sub

Private Property Get IAnalysisTablePolicyResult_IterationItem() As IAnalysisTableIterationItem
    Set IAnalysisTablePolicyResult_IterationItem = this.iterationItem
End Property

Private Property Get IAnalysisTablePolicyResult_Policy() As ITableTypePolicy
    Set IAnalysisTablePolicyResult_Policy = this.policy
End Property

Private Property Get IAnalysisTablePolicyResult_Context() As ITableSpecsPolicyContext
    Set IAnalysisTablePolicyResult_Context = this.context
End Property

Private Property Get IAnalysisTablePolicyResult_IsValid() As Boolean
    IAnalysisTablePolicyResult_IsValid = this.isValid
End Property

Private Property Get IAnalysisTablePolicyResult_HasPercent() As Boolean
    IAnalysisTablePolicyResult_HasPercent = this.hasPercent
End Property

Private Property Get IAnalysisTablePolicyResult_HasTotal() As Boolean
    IAnalysisTablePolicyResult_HasTotal = this.hasTotal
End Property

Private Property Get IAnalysisTablePolicyResult_HasMissing() As Boolean
    IAnalysisTablePolicyResult_HasMissing = this.hasMissing
End Property

Private Property Get IAnalysisTablePolicyResult_HasGraph() As Boolean
    IAnalysisTablePolicyResult_HasGraph = this.hasGraph
End Property
