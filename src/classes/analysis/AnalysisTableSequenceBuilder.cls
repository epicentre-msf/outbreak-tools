VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableSequenceBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Coordinates enumerator and policy resolver to build write-ready table sequences")
'@depends ProjectError, AnalysisTableEnumerator, AnalysisTablePolicyResolver, BetterArray
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTableSequenceBuilder

Private Type TSequenceBuilderState
    enumerator As IAnalysisTableEnumerator
    policyResolver As IAnalysisTablePolicyResolver
End Type

Private this As TSequenceBuilderState
Private Const CLASSNAME As String = "AnalysisTableSequenceBuilder"

'@section Factory
'===============================================================================
'Create a sequence builder with optional dependency overrides.
Public Function Create(ByVal linelistSpecs As ILinelistSpecs, _
                       Optional ByVal enumerator As IAnalysisTableEnumerator = Nothing, _
                       Optional ByVal policyResolver As IAnalysisTablePolicyResolver = Nothing) As IAnalysisTableSequenceBuilder

    ValidateLinelistSpecs linelistSpecs

    With New AnalysisTableSequenceBuilder
        If enumerator Is Nothing Then
            Set .Enumerator = AnalysisTableEnumerator.Create
        Else
            Set .Enumerator = enumerator
        End If

        If policyResolver Is Nothing Then
            Set .PolicyResolver = AnalysisTablePolicyResolver.Create(linelistSpecs)
        Else
            Set .PolicyResolver = policyResolver
        End If

        Set Create = .Self
    End With
End Function

'Expose the current instance through its interface type.
Public Property Get Self() As IAnalysisTableSequenceBuilder
    Set Self = Me
End Property

'@section Properties
'===============================================================================
Public Property Set Enumerator(ByVal enumerator As IAnalysisTableEnumerator)
    ValidateEnumerator enumerator
    Set this.enumerator = enumerator
End Property

Private Property Get EnumeratorInternal() As IAnalysisTableEnumerator
    Set EnumeratorInternal = this.enumerator
End Property

Public Property Set PolicyResolver(ByVal policyResolver As IAnalysisTablePolicyResolver)
    ValidatePolicyResolver policyResolver
    Set this.policyResolver = policyResolver
End Property

Private Property Get PolicyResolverInternal() As IAnalysisTablePolicyResolver
    Set PolicyResolverInternal = this.policyResolver
End Property

'@section Internal helpers
'===============================================================================
Private Function CreateResultBuffer() As BetterArray
    Dim buffer As BetterArray

    Set buffer = New BetterArray
    buffer.LowerBound = 1
    Set CreateResultBuffer = buffer
End Function

Private Sub AppendResult(ByVal results As BetterArray, ByVal policyResult As IAnalysisTablePolicyResult)
    If policyResult Is Nothing Then Exit Sub
    If Not policyResult.IsValid Then Exit Sub
    results.Push policyResult
End Sub

'@section Validation helpers
'===============================================================================
Private Sub ValidateEnumerator(ByVal enumerator As IAnalysisTableEnumerator)
    If enumerator Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "enumerator cannot be Nothing"
    End If
End Sub

Private Sub ValidatePolicyResolver(ByVal policyResolver As IAnalysisTablePolicyResolver)
    If policyResolver Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "policyResolver cannot be Nothing"
    End If
End Sub

Private Sub ValidateLinelistSpecs(ByVal linelistSpecs As ILinelistSpecs)
    If linelistSpecs Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "linelistSpecs cannot be Nothing"
    End If
End Sub

Private Sub ValidateSpecificationList(ByVal specificationList As ListObject)
    If specificationList Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "specificationList cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IAnalysisTableSequenceBuilder_Build(ByVal specificationList As ListObject, _
                                                     ByVal linelistSpecs As ILinelistSpecs, _
                                                     ByVal tableTypeCode As Byte) As BetterArray
    Dim iterationItems As BetterArray
    Dim idx As Long
    Dim policyResult As IAnalysisTablePolicyResult
    Dim results As BetterArray

    ValidateSpecificationList specificationList
    ValidateLinelistSpecs linelistSpecs

    Set results = CreateResultBuffer()

    Set iterationItems = EnumeratorInternal.Enumerate(specificationList, linelistSpecs, tableTypeCode)
    If iterationItems Is Nothing Then
        Set IAnalysisTableSequenceBuilder_Build = results
        Exit Function
    End If

    For idx = iterationItems.LowerBound To iterationItems.UpperBound
        Set policyResult = PolicyResolverInternal.Resolve(iterationItems.Item(idx))
        AppendResult results, policyResult
    Next idx

    Set IAnalysisTableSequenceBuilder_Build = results
End Function
