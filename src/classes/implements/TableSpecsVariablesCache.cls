VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecsVariablesCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Provides reusable dictionary/variables references for TableSpecs workflows")
'@depends ProjectError, TableSpecsPolicyHelpers
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableSpecsVariablesCache

Private Type TVariablesCacheState
    linelistSpecs As ILinelistSpecs
    dictionary As ILLdictionary
    variables As ILLVariables
End Type

Private this As TVariablesCacheState
Private Const CLASSNAME As String = "TableSpecsVariablesCache"

'@section Factory
'===============================================================================
'@label:Create
'Create a cache backed by the provided linelist specifications reference.
Public Function Create(ByVal linelistSpecs As ILinelistSpecs) As ITableSpecsVariablesCache
    ValidateLinelistSpecs linelistSpecs

    With New TableSpecsVariablesCache
        Set .LinelistSpecifications = linelistSpecs
        .Refresh
        Set Create = .Self
    End With
End Function

'@label:Self
'Expose the current instance as its interface type for fluent chaining.
Public Property Get Self() As ITableSpecsVariablesCache
    Set Self = Me
End Property

'@section Source configuration
'===============================================================================
'@label:GetLinelistSpecifications
'Return the linelist specifications currently bound to the cache.
Public Property Get LinelistSpecifications() As ILinelistSpecs
    Set LinelistSpecifications = this.linelistSpecs
End Property

'@label:SetLinelistSpecifications
'Assign a new linelist specifications reference and reset caches whenever it changes.
Public Property Set LinelistSpecifications(ByVal linelistSpecs As ILinelistSpecs)
    ValidateLinelistSpecs linelistSpecs
    Set this.linelistSpecs = linelistSpecs
    ResetCaches
End Property

'@section Cached accessors
'===============================================================================
'@label:GetDictionary
'Retrieve the dictionary reference, fetching it from the linelist specifications when required.
Public Property Get Dictionary() As ILLdictionary
    EnsureDictionaryReady
    Set Dictionary = this.dictionary
End Property

'@label:GetVariables
'Retrieve a cached ILLVariables instance, creating it on demand from the dictionary.
Public Property Get Variables() As ILLVariables
    EnsureVariablesReady
    Set Variables = this.variables
End Property

'@section Cache management
'===============================================================================
'@label:Invalidate
'Explicitly clear cached references.
Public Sub Invalidate()
    ResetCaches
End Sub

'@label:Refresh
'Rebuild cached collaborators immediately.
Public Sub Refresh()
    ResetCaches
    EnsureDictionaryReady
    EnsureVariablesReady
End Sub

'@section Internal helpers
'===============================================================================
'@label:ResetCaches
'Clear cached dictionary and variables without altering the bound linelist specifications.
Private Sub ResetCaches()
    Set this.dictionary = Nothing
    Set this.variables = Nothing
End Sub

'@label:EnsureDictionaryReady
'Retrieve the dictionary from linelist specifications when we do not have a cached instance.
Private Sub EnsureDictionaryReady()
    If this.dictionary Is Nothing Then
        ValidateLinelistSpecs this.linelistSpecs
        Set this.dictionary = FetchDictionary()
        ValidateDictionary this.dictionary
    End If
End Sub

'@label:EnsureVariablesReady
'Create the variables helper when missing, reusing the cached dictionary.
Private Sub EnsureVariablesReady()
    If this.variables Is Nothing Then
        EnsureDictionaryReady
        Set this.variables = LLVariables.Create(this.dictionary)
    End If
End Sub

'@label:FetchDictionary
'Wrap the linelist specifications dictionary retrieval so we can centralise error handling.
Private Function FetchDictionary() As ILLdictionary
    On Error GoTo CleanFail
        Set FetchDictionary = this.linelistSpecs.Dictionary
        Exit Function
CleanFail:
    ThrowError CLng(ProjectError.ErrorUnexpectedState), _
               "Unable to retrieve dictionary from linelist specifications"
End Function

'@label:ValidateLinelistSpecs
'Ensure callers provide a valid linelist specifications reference.
Private Sub ValidateLinelistSpecs(ByVal linelistSpecs As ILinelistSpecs)
    If linelistSpecs Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "linelistSpecs cannot be Nothing"
    End If

    If NormalizeValue(TypeName(linelistSpecs)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "linelistSpecs must expose a concrete implementation"
    End If
End Sub

'@label:ValidateDictionary
'Ensure a usable dictionary reference is available.
Private Sub ValidateDictionary(ByVal dictionary As ILLdictionary)
    If dictionary Is Nothing Then
        ThrowError CLng(ProjectError.ElementNotFound), _
                   "Linelist specifications did not provide a dictionary reference"
    End If
End Sub

'@label:ThrowError
'Centralised error helper to guarantee consistent metadata on failures.
Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function ITableSpecsVariablesCache_Create(ByVal linelistSpecs As ILinelistSpecs) As ITableSpecsVariablesCache
    Set ITableSpecsVariablesCache_Create = Create(linelistSpecs)
End Function

Private Property Get ITableSpecsVariablesCache_Self() As ITableSpecsVariablesCache
    Set ITableSpecsVariablesCache_Self = Self
End Property

Private Property Get ITableSpecsVariablesCache_LinelistSpecifications() As ILinelistSpecs
    Set ITableSpecsVariablesCache_LinelistSpecifications = LinelistSpecifications
End Property

Private Property Set ITableSpecsVariablesCache_LinelistSpecifications(ByVal linelistSpecs As ILinelistSpecs)
    Set Me.LinelistSpecifications = linelistSpecs
End Property

Private Property Get ITableSpecsVariablesCache_Dictionary() As ILLdictionary
    Set ITableSpecsVariablesCache_Dictionary = Dictionary
End Property

Private Property Get ITableSpecsVariablesCache_Variables() As ILLVariables
    Set ITableSpecsVariablesCache_Variables = Variables
End Property

Private Sub ITableSpecsVariablesCache_Invalidate()
    Invalidate
End Sub

Private Sub ITableSpecsVariablesCache_Refresh()
    Refresh
End Sub
