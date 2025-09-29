VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLFormatLogStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests")
'@ModuleDescription("ILLFormat test double capturing ApplyFormat calls")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLFormat

Private Type TFormatLog
    scopes As Collection
End Type

Private this As TFormatLog

Private Sub Class_Initialize()
    Set this.scopes = New Collection
End Sub

Public Function ScopeCount(ByVal scopeValue As Byte) As Long
    Dim item As Variant
    Dim count As Long
    For Each item In this.scopes
        If CLng(item) = CLng(scopeValue) Then count = count + 1
    Next item
    ScopeCount = count
End Function

Public Property Get AppliedScopes() As Collection
    Set AppliedScopes = this.scopes
End Property

'@section ILLFormat implementation
'===============================================================================
Private Sub ILLFormat_ApplyFormat(ByVal formattingObject As Variant, ByVal scope As Byte, _
                                  Optional ByVal nbDecimals As Integer = 0)
    this.scopes.Add scope
End Sub

Private Function ILLFormat_DesignValue(ByVal labelName As String, _
                                       Optional ByVal returnColor As Boolean = True) As Variant
    ILLFormat_DesignValue = Empty
End Function

Private Sub ILLFormat_Import(ByVal inpsh As Worksheet)
    'Not required for stub
End Sub

Private Property Get ILLFormat_HasCheckings() As Boolean
    ILLFormat_HasCheckings = False
End Property

Private Property Get ILLFormat_CheckingValues() As IChecking
    Set ILLFormat_CheckingValues = Nothing
End Property

