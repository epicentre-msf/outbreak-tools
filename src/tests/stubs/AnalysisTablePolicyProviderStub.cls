VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTablePolicyProviderStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisTablePolicyProvider double mapping table types to stub policies")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisTablePolicyProvider

Private Type TProviderStubState
    policies As Collection
End Type

Private this As TProviderStubState

Private Sub Class_Initialize()
    Set this.policies = New Collection
End Sub

'Store a policy instance resolved for the supplied table type code.
Public Sub AddPolicy(ByVal tableTypeCode As Byte, ByVal policy As ITableTypePolicy)
    Dim key As String

    key = CStr(tableTypeCode)

    On Error Resume Next
        this.policies.Remove key
    On Error GoTo 0

    On Error Resume Next
        this.policies.Add policy, key
    On Error GoTo 0
End Sub

'@section Interface implementation
'===============================================================================
Private Function IAnalysisTablePolicyProvider_ResolvePolicy(ByVal tableTypeCode As Byte) As ITableTypePolicy
    Dim key As String
    Dim policy As ITableTypePolicy

    key = CStr(tableTypeCode)

    On Error Resume Next
        Set policy = this.policies.Item(key)
    On Error GoTo 0

    Set IAnalysisTablePolicyProvider_ResolvePolicy = policy
End Function
