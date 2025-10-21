VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGraphPlanItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Concrete implementation describing a plan item that requires graph generation")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGraphPlanItem

Private Type TGraphPlanItemState
    planItem As IAnalysisTablePlanItem
    policyResult As IAnalysisTablePolicyResult
    specification As ITablesSpecs
    tableType As Byte
    sequenceIndex As Long
    graphLabel As String
End Type

Private this As TGraphPlanItemState

Private Const CLASSNAME As String = "AnalysisGraphPlanItem"

'@section Factory
'===============================================================================
'Create a populated graph plan item based on the supplied table plan entry.
Public Function Create(ByVal planItem As IAnalysisTablePlanItem) As IAnalysisGraphPlanItem
    ValidatePlanItem planItem

    With New AnalysisGraphPlanItem
        .Initialise planItem
        Set Create = .Self
    End With
End Function

'Expose the current instance through its interface type.
Public Property Get Self() As IAnalysisGraphPlanItem
    Set Self = Me
End Property

'@section Initialisation
'===============================================================================
Private Sub Initialise(ByVal planItem As IAnalysisTablePlanItem)
    Dim policyResult As IAnalysisTablePolicyResult
    Dim iterationItem As IAnalysisTableIterationItem
    Dim specification As ITablesSpecs

    Set policyResult = planItem.PolicyResult
    ValidatePolicyResult policyResult

    Set iterationItem = policyResult.IterationItem
    ValidateIterationItem iterationItem

    Set specification = iterationItem.Specification
    ValidateSpecification specification

    Set this.planItem = planItem
    Set this.policyResult = policyResult
    Set this.specification = specification

    this.sequenceIndex = planItem.SequenceIndex
    this.tableType = ResolveTableType(specification)
    this.graphLabel = ResolveGraphLabel(specification)
End Sub

Private Function ResolveTableType(ByVal specification As ITablesSpecs) As Byte
    On Error Resume Next
        ResolveTableType = CByte(specification.TableType)
    On Error GoTo 0
End Function

Private Function ResolveGraphLabel(ByVal specification As ITablesSpecs) As String
    Dim labelText As String

    On Error Resume Next
        labelText = Trim$(CStr(specification.Value("label")))
    On Error GoTo 0

    If LenB(labelText) = 0 Then
        On Error Resume Next
            labelText = Trim$(CStr(specification.TableId))
        On Error GoTo 0
    End If

    ResolveGraphLabel = labelText
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidatePlanItem(ByVal planItem As IAnalysisTablePlanItem)
    If planItem Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "planItem cannot be Nothing"
    End If
End Sub

Private Sub ValidatePolicyResult(ByVal policyResult As IAnalysisTablePolicyResult)
    If policyResult Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "PolicyResult cannot be Nothing"
    End If
End Sub

Private Sub ValidateIterationItem(ByVal iterationItem As IAnalysisTableIterationItem)
    If iterationItem Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "IterationItem cannot be Nothing"
    End If
End Sub

Private Sub ValidateSpecification(ByVal specification As ITablesSpecs)
    If specification Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "Specification cannot be Nothing"
    End If

    If StrComp(TypeName(specification), "Nothing", vbBinaryCompare) = 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), "Specification must be a valid object"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

Private Function TableTypeInternal() As Byte
    TableTypeInternal = this.tableType
End Function

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisGraphPlanItem_PlanItem() As IAnalysisTablePlanItem
    Set IAnalysisGraphPlanItem_PlanItem = this.planItem
End Property

Private Property Get IAnalysisGraphPlanItem_PolicyResult() As IAnalysisTablePolicyResult
    Set IAnalysisGraphPlanItem_PolicyResult = this.policyResult
End Property

Private Property Get IAnalysisGraphPlanItem_Specification() As ITablesSpecs
    Set IAnalysisGraphPlanItem_Specification = this.specification
End Property

Private Property Get IAnalysisGraphPlanItem_TableType() As AnalysisTablesType
    IAnalysisGraphPlanItem_TableType = TableTypeInternal()
End Property

Private Property Get IAnalysisGraphPlanItem_SequenceIndex() As Long
    IAnalysisGraphPlanItem_SequenceIndex = this.sequenceIndex
End Property

Private Property Get IAnalysisGraphPlanItem_IsTimeSeries() As Boolean
    IAnalysisGraphPlanItem_IsTimeSeries = (TableTypeInternal() = TypeTimeSeries)
End Property

Private Property Get IAnalysisGraphPlanItem_IsSpatioTemporal() As Boolean
    IAnalysisGraphPlanItem_IsSpatioTemporal = (TableTypeInternal() = TypeSpatioTemporal)
End Property

Private Property Get IAnalysisGraphPlanItem_GraphLabel() As String
    IAnalysisGraphPlanItem_GraphLabel = this.graphLabel
End Property
