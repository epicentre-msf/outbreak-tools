VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTablePolicyResolverStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisTablePolicyResolver double returning preconfigured results")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisTablePolicyResolver

Private Type TResolverStubState
    results As Collection
    lastIterationItem As IAnalysisTableIterationItem
    invalidateCount As Long
End Type

Private this As TResolverStubState

Private Sub Class_Initialize()
    Set this.results = New Collection
End Sub

'Queue a policy result to be returned on the next Resolve call.
Public Sub EnqueueResult(ByVal policyResult As IAnalysisTablePolicyResult)
    this.results.Add policyResult
End Sub

'Retrieve the last iteration item provided to Resolve.
Public Property Get LastIterationItem() As IAnalysisTableIterationItem
    Set LastIterationItem = this.lastIterationItem
End Property

'Retrieve the number of times Invalidate was invoked.
Public Property Get InvalidateCount() As Long
    InvalidateCount = this.invalidateCount
End Property

'@section Interface implementation
'===============================================================================
Private Function IAnalysisTablePolicyResolver_Resolve(ByVal iterationItem As IAnalysisTableIterationItem) As IAnalysisTablePolicyResult
    Dim policyResult As IAnalysisTablePolicyResult

    Set this.lastIterationItem = iterationItem

    On Error Resume Next
        Set policyResult = this.results.Item(1)
        this.results.Remove 1
    On Error GoTo 0

    Set IAnalysisTablePolicyResolver_Resolve = policyResult
End Function

Private Sub IAnalysisTablePolicyResolver_Invalidate()
    this.invalidateCount = this.invalidateCount + 1
End Sub
