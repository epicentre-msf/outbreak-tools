VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableTypePolicyFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Factory returning strategy objects for TableSpecs table types")
'@depends ProjectError, TableSpecsPolicyHelpers
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Private Type TFactoryState
    cache As Collection
End Type

Private this As TFactoryState
Private Const CLASSNAME As String = "TableTypePolicyFactory"

'@section Lifecycle
'===============================================================================
'Ensure the internal cache exists before the factory is used.
Private Sub Class_Initialize()
    Set this.cache = New Collection
End Sub

'@section Factory API
'===============================================================================
'Return a policy instance matching the requested table type code.
Public Function Create(ByVal tableTypeCode As Byte) As ITableTypePolicy
    Set Create = ResolvePolicy(tableTypeCode)
End Function

'@section Resolution helpers
'===============================================================================
'Locate a cached policy or create it on demand.
Private Function ResolvePolicy(ByVal tableTypeCode As Byte) As ITableTypePolicy
    Dim cacheKey As String
    Dim policy As ITableTypePolicy

    cacheKey = CStr(tableTypeCode)

    On Error Resume Next
        Set policy = this.cache.Item(cacheKey)
    On Error GoTo 0

    If Not policy Is Nothing Then
        Set ResolvePolicy = policy
        Exit Function
    End If

    Set policy = CreatePolicyInstance(tableTypeCode)
    this.cache.Add policy, cacheKey
    Set ResolvePolicy = policy
End Function

'Create a fresh policy for the supplied table type code.
Private Function CreatePolicyInstance(ByVal tableTypeCode As Byte) As ITableTypePolicy
    Select Case tableTypeCode
        Case TABLE_TYPE_GLOBAL_SUMMARY
            Set CreatePolicyInstance = New TableTypePolicyGlobalSummary
        Case TABLE_TYPE_UNIVARIATE
            Set CreatePolicyInstance = New TableTypePolicyUnivariate
        Case TABLE_TYPE_BIVARIATE
            Set CreatePolicyInstance = New TableTypePolicyBivariate
        Case TABLE_TYPE_TIME_SERIES
            Set CreatePolicyInstance = New TableTypePolicyTimeSeries
        Case TABLE_TYPE_SPATIAL
            Set CreatePolicyInstance = New TableTypePolicySpatial
        Case TABLE_TYPE_SPATIO_TEMPORAL
            Set CreatePolicyInstance = New TableTypePolicySpatioTemporal
        Case Else
            ThrowError CLng(ProjectError.InvalidArgument), _
                       "Unsupported table type code: " & CStr(tableTypeCode)
    End Select
End Function

'@section Error handling
'===============================================================================
'Centralised error helper to guarantee consistent error metadata.
Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

