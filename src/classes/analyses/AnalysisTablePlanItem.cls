VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTablePlanItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Immutable implementation of IAnalysisTablePlanItem")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTablePlanItem

Private Type TPlanItemState
    policyResult As IAnalysisTablePolicyResult
    sequenceIndex As Long
End Type

Private this As TPlanItemState
Private Const CLASSNAME As String = "AnalysisTablePlanItem"

'@section Factory
'===============================================================================
'Create a plan item capturing the supplied policy result and sequence index.
Public Function Create(ByVal policyResult As IAnalysisTablePolicyResult, _
                       ByVal sequenceIndex As Long) As IAnalysisTablePlanItem
    ValidatePolicyResult policyResult
    ValidateSequenceIndex sequenceIndex

    With New AnalysisTablePlanItem
        Set .PolicyResult = policyResult
        .SequenceIndex = sequenceIndex
        Set Create = .Self
    End With
End Function

'Expose the instance through its interface.
Public Property Get Self() As IAnalysisTablePlanItem
    Set Self = Me
End Property

'@section Properties
'===============================================================================
Public Property Set PolicyResult(ByVal policyResult As IAnalysisTablePolicyResult)
    ValidatePolicyResult policyResult
    Set this.policyResult = policyResult
End Property

Private Property Get PolicyResultInternal() As IAnalysisTablePolicyResult
    Set PolicyResultInternal = this.policyResult
End Property

Public Property Let SequenceIndex(ByVal sequenceIndex As Long)
    ValidateSequenceIndex sequenceIndex
    this.sequenceIndex = sequenceIndex
End Property

Private Property Get SequenceIndexInternal() As Long
    SequenceIndexInternal = this.sequenceIndex
End Property

'@section Validation helpers
'===============================================================================
Private Sub ValidatePolicyResult(ByVal policyResult As IAnalysisTablePolicyResult)
    If policyResult Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "policyResult cannot be Nothing"
    End If
End Sub

Private Sub ValidateSequenceIndex(ByVal sequenceIndex As Long)
    If sequenceIndex < 0& Then
        RaiseError CLng(ProjectError.InvalidArgument), "sequenceIndex cannot be negative"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisTablePlanItem_PolicyResult() As IAnalysisTablePolicyResult
    Set IAnalysisTablePlanItem_PolicyResult = PolicyResultInternal()
End Property

Private Property Get IAnalysisTablePlanItem_SequenceIndex() As Long
    IAnalysisTablePlanItem_SequenceIndex = SequenceIndexInternal()
End Property
