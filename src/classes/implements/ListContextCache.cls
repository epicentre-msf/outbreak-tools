VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListContextCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Linelist.Shared")
'@ModuleDescription("Stores linelist-related collaborators so builders can reuse them without repeated lookups")
'@depends ProjectError
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IListContextCache

Private Type TCacheState
    linelist As Object
    dictionary As Object
    variables As Object
End Type

Private this As TCacheState
Private Const CLASSNAME As String = "ListContextCache"

'@section Factory
'===============================================================================
'Create a cache optionally primed with a linelist reference.
Public Function Create(Optional ByVal linelist As Object = Nothing) As IListContextCache
    With New ListContextCache
        If Not linelist Is Nothing Then .AttachLinelist linelist
        Set Create = .Self
    End With
End Function

'Expose the instance as its interface type.
Public Property Get Self() As IListContextCache
    Set Self = Me
End Property

'@section Validation helpers
'===============================================================================
Private Sub ValidateObject(ByVal candidate As Object, ByVal argumentName As String)
    If candidate Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Sub IListContextCache_AttachLinelist(ByVal linelist As Object)
    ValidateObject linelist, "linelist"
    Set this.linelist = linelist
End Sub

Private Sub IListContextCache_AttachDictionary(ByVal dictionary As Object)
    ValidateObject dictionary, "dictionary"
    Set this.dictionary = dictionary
End Sub

Private Sub IListContextCache_AttachVariables(ByVal variables As Object)
    ValidateObject variables, "variables"
    Set this.variables = variables
End Sub

Private Function IListContextCache_Linelist() As Object
    If this.linelist Is Nothing Then
        ThrowError CLng(ProjectError.ObjectNotInitialized), _
                  "Linelist reference has not been attached to the cache"
    End If
    Set IListContextCache_Linelist = this.linelist
End Function

Private Function IListContextCache_Dictionary() As Object
    If this.dictionary Is Nothing Then
        ThrowError CLng(ProjectError.ObjectNotInitialized), _
                  "Dictionary reference has not been attached to the cache"
    End If
    Set IListContextCache_Dictionary = this.dictionary
End Function

Private Function IListContextCache_Variables() As Object
    If this.variables Is Nothing Then
        ThrowError CLng(ProjectError.ObjectNotInitialized), _
                  "Variables reference has not been attached to the cache"
    End If
    Set IListContextCache_Variables = this.variables
End Function

Private Sub IListContextCache_Invalidate()
    Set this.dictionary = Nothing
    Set this.variables = Nothing
End Sub

