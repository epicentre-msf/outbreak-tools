VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecsDictionaryAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Wraps ILLdictionary for use inside TableSpecs policies")
'@depends ProjectError
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableSpecsDictionaryAccessor

Private Type TDictionaryAdapterState
    dictionary As ILLdictionary
End Type

Private this As TDictionaryAdapterState
Private Const CLASSNAME As String = "TableSpecsDictionaryAdapter"

'@section Factory
'===============================================================================
'Create a new adapter using the supplied dictionary reference.
Public Function Create(ByVal dictionary As ILLdictionary) As ITableSpecsDictionaryAccessor
    ValidateDictionary dictionary

    With New TableSpecsDictionaryAdapter
        Set .Dictionary = dictionary
        Set Create = .Self
    End With
End Function

'Expose the instance as its interface type.
Public Property Get Self() As ITableSpecsDictionaryAccessor
    Set Self = Me
End Property

'@section Properties
'===============================================================================
'Assign the dictionary reference once it has been validated.
Public Property Set Dictionary(ByVal dictionary As ILLdictionary)
    ValidateDictionary dictionary
    Set this.dictionary = dictionary
End Property

'Retrieve the wrapped dictionary reference.
Private Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dictionary
End Property

'@section Validation helpers
'===============================================================================
'Ensure the adapter receives a non-empty dictionary reference.
Private Sub ValidateDictionary(ByVal dictionary As ILLdictionary)
    If dictionary Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "dictionary cannot be Nothing"
    End If

    If NormalizeValue(TypeName(dictionary)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "dictionary must expose a concrete implementation"
    End If
End Sub

'Centralised error helper so every failure surfaces with a consistent source.
Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function ITableSpecsDictionaryAccessor_VariableExists(ByVal variableName As String) As Boolean
    ITableSpecsDictionaryAccessor_VariableExists = Dictionary.VariableExists(variableName)
End Function

