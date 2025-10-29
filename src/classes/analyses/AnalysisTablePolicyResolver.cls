VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTablePolicyResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Resolves table-type strategies and evaluation flags for analysis specifications")
'@depends ProjectError, AnalysisTablePolicyResult, AnalysisTablePolicyProvider, TableSpecsColumnMap, TableSpecsDictionaryAdapter, TableSpecsVariablesAdapter, TableSpecsPolicyContext, TableSpecsVariablesCache
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTablePolicyResolver

Private Type TResolverState
    variablesCache As ITableSpecsVariablesCache
    policyProvider As IAnalysisTablePolicyProvider
End Type

Private this As TResolverState
Private Const CLASSNAME As String = "AnalysisTablePolicyResolver"

'@section Factory
'===============================================================================
'Create a resolver bound to the provided linelist specifications.
Public Function Create(ByVal linelistSpecs As ILinelistSpecs, _
                       Optional ByVal variablesCache As ITableSpecsVariablesCache = Nothing, _
                       Optional ByVal policyProvider As IAnalysisTablePolicyProvider = Nothing) As IAnalysisTablePolicyResolver

    ValidateLinelistSpecs linelistSpecs

    With New AnalysisTablePolicyResolver
        If variablesCache Is Nothing Then
            Set .VariablesCache = TableSpecsVariablesCache.Create(linelistSpecs)
        Else
            Set .VariablesCache = variablesCache
        End If

        If policyProvider Is Nothing Then
            Set .PolicyProvider = AnalysisTablePolicyProvider.Create()
        Else
            Set .PolicyProvider = policyProvider
        End If

        Set Create = .Self
    End With
End Function

'Expose the current instance as its interface type.
Public Property Get Self() As IAnalysisTablePolicyResolver
    Set Self = Me
End Property

'@section Properties
'===============================================================================
Public Property Set VariablesCache(ByVal variablesCache As ITableSpecsVariablesCache)
    ValidateVariablesCache variablesCache
    Set this.variablesCache = variablesCache
End Property

Private Property Get VariablesCacheInternal() As ITableSpecsVariablesCache
    Set VariablesCacheInternal = this.variablesCache
End Property

Public Property Set PolicyProvider(ByVal policyProvider As IAnalysisTablePolicyProvider)
    ValidatePolicyProvider policyProvider
    Set this.policyProvider = policyProvider
End Property

Private Property Get PolicyProviderInternal() As IAnalysisTablePolicyProvider
    Set PolicyProviderInternal = this.policyProvider
End Property

'@section Public helpers
'===============================================================================
Public Sub InvalidateCaches()
    If Not this.variablesCache Is Nothing Then this.variablesCache.Invalidate
End Sub

'@section Internal helpers
'===============================================================================
Private Function ResolveColumnMap(ByVal iterationItem As IAnalysisTableIterationItem) As ITableSpecsColumnMap
    Set ResolveColumnMap = TableSpecsColumnMap.Create(iterationItem.HeaderRange, iterationItem.RowRange)
End Function

Private Function CreateVariablesAccessor() As ITableSpecsVariableAccessor
    Dim variables As ILLVariables

    Set variables = VariablesCacheInternal.Variables
    Set CreateVariablesAccessor = TableSpecsVariablesAdapter.Create(variables)
End Function

Private Function CreateDictionaryAccessor() As ITableSpecsDictionaryAccessor
    Dim dictionary As ILLdictionary

    Set dictionary = VariablesCacheInternal.Dictionary
    Set CreateDictionaryAccessor = TableSpecsDictionaryAdapter.Create(dictionary)
End Function

Private Function CreatePolicyContext(ByVal tableTypeCode As Byte, _
                                     ByVal columnMap As ITableSpecsColumnMap) As ITableSpecsPolicyContext
    Set CreatePolicyContext = TableSpecsPolicyContext.Create(tableTypeCode, columnMap, CreateVariablesAccessor(), CreateDictionaryAccessor())
End Function

Private Function CreatePolicyResult(ByVal iterationItem As IAnalysisTableIterationItem, _
                                    ByVal policy As ITableTypePolicy, _
                                    ByVal context As ITableSpecsPolicyContext) As IAnalysisTablePolicyResult
    Dim isValid As Boolean
    Dim hasPercent As Boolean
    Dim hasTotal As Boolean
    Dim hasMissing As Boolean
    Dim hasGraph As Boolean

    isValid = policy.IsValid(context)
    hasPercent = policy.HasPercent(context)
    hasTotal = policy.HasTotal(context)
    hasMissing = policy.HasMissing(context)
    hasGraph = policy.HasGraph(context)

    Set CreatePolicyResult = AnalysisTablePolicyResult.Create(iterationItem, policy, context, isValid, hasPercent, hasTotal, hasMissing, hasGraph)
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidateLinelistSpecs(ByVal linelistSpecs As ILinelistSpecs)
    If linelistSpecs Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "linelistSpecs cannot be Nothing"
    End If
End Sub

Private Sub ValidateVariablesCache(ByVal variablesCache As ITableSpecsVariablesCache)
    If variablesCache Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "variablesCache cannot be Nothing"
    End If
End Sub

Private Sub ValidatePolicyProvider(ByVal policyProvider As IAnalysisTablePolicyProvider)
    If policyProvider Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "policyProvider cannot be Nothing"
    End If
End Sub

Private Sub ValidateIterationItem(ByVal iterationItem As IAnalysisTableIterationItem)
    If iterationItem Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "iterationItem cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IAnalysisTablePolicyResolver_Resolve(ByVal iterationItem As IAnalysisTableIterationItem) As IAnalysisTablePolicyResult
    Dim specification As ITablesSpecs
    Dim policy As ITableTypePolicy
    Dim columnMap As ITableSpecsColumnMap
    Dim context As ITableSpecsPolicyContext

    ValidateIterationItem iterationItem
    Set specification = iterationItem.Specification

    Set policy = PolicyProviderInternal.ResolvePolicy(specification.TableType)
    If policy Is Nothing Then
        RaiseError CLng(ProjectError.ElementNotFound), _
                   "Policy provider did not return a strategy for table type " & CStr(specification.TableType)
    End If
    Set columnMap = ResolveColumnMap(iterationItem)
    Set context = CreatePolicyContext(specification.TableType, columnMap)

    Set IAnalysisTablePolicyResolver_Resolve = CreatePolicyResult(iterationItem, policy, context)
End Function

Private Sub IAnalysisTablePolicyResolver_Invalidate()
    InvalidateCaches
End Sub
