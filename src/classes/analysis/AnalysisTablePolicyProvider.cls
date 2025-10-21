VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTablePolicyProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Wraps TableTypePolicyFactory to expose IAnalysisTablePolicyProvider")
'@depends ProjectError, TableTypePolicyFactory
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTablePolicyProvider

Private Type TProviderState
    factory As TableTypePolicyFactory
End Type

Private this As TProviderState
Private Const CLASSNAME As String = "AnalysisTablePolicyProvider"

'@section Factory
'===============================================================================
'Create a provider backed by the default TableTypePolicyFactory implementation.
Public Function Create(Optional ByVal factory As TableTypePolicyFactory = Nothing) As IAnalysisTablePolicyProvider
    With New AnalysisTablePolicyProvider
        If factory Is Nothing Then
            Set .Factory = New TableTypePolicyFactory
        Else
            Set .Factory = factory
        End If
        Set Create = .Self
    End With
End Function

'Return the current instance as the provider interface.
Public Property Get Self() As IAnalysisTablePolicyProvider
    Set Self = Me
End Property

'@section Properties
'===============================================================================
'Assign the underlying policy factory reference.
Public Property Set Factory(ByVal factory As TableTypePolicyFactory)
    ValidateFactory factory
    Set this.factory = factory
End Property

'Expose the current factory reference.
Private Property Get FactoryInternal() As TableTypePolicyFactory
    Set FactoryInternal = this.factory
End Property

'@section Helpers
'===============================================================================
Private Sub ValidateFactory(ByVal factory As TableTypePolicyFactory)
    If factory Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "factory cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IAnalysisTablePolicyProvider_ResolvePolicy(ByVal tableTypeCode As Byte) As ITableTypePolicy
    On Error GoTo CleanFail
        Set IAnalysisTablePolicyProvider_ResolvePolicy = FactoryInternal.Create(tableTypeCode)
        Exit Function
CleanFail:
    RaiseError CLng(ProjectError.ErrorUnexpectedState), _
               "Unable to resolve policy for table type " & CStr(tableTypeCode)
End Function
