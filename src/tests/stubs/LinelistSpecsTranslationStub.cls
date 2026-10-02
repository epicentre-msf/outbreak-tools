VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistSpecsTranslationStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Lightweight ITranslationObject stub for unit tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ITranslationObject

Private Type TTranslationStub
    nameText As String
    values As Collection
    hasChecks As Boolean
End Type

Private this As TTranslationStub

'@section Configuration
'===============================================================================
Public Sub Initialise(Optional ByVal translationName As String = "TranslationStub")
    this.nameText = translationName
    Set this.values = New Collection
End Sub

Public Sub SetTranslation(ByVal keyText As String, ByVal valueText As String)
    Dim key As String

    key = NormaliseKey(keyText)
    EnsureStorage

    On Error Resume Next
        this.values.Remove key
    On Error GoTo 0

    this.values.Add valueText, key
End Sub

Private Sub EnsureStorage()
    If this.values Is Nothing Then Set this.values = New Collection
End Sub

Private Function NormaliseKey(ByVal keyText As String) As String
    NormaliseKey = LCase$(Trim$(keyText))
End Function

Private Function ResolveValue(ByVal keyText As String) As String
    Dim key As String
    Dim result As String

    key = NormaliseKey(keyText)
    EnsureStorage

    On Error Resume Next
        result = CStr(this.values.Item(key))
    On Error GoTo 0

    ResolveValue = result
End Function

'@section ITranslationObject implementation
'===============================================================================
Private Property Get ITranslationObject_Name() As String
    If LenB(this.nameText) = 0 Then this.nameText = "TranslationStub"
    ITranslationObject_Name = this.nameText
End Property

Private Property Get ITranslationObject_ValueExists(ByVal text As String) As Boolean
    ITranslationObject_ValueExists = LenB(ResolveValue(text)) > 0
End Property

Private Function ITranslationObject_TranslatedValue(ByVal Text As String, Optional ByVal containsFormula As Boolean = False) As String
    Dim stored As String
    stored = ResolveValue(Text)

    If LenB(stored) = 0 Then
        ITranslationObject_TranslatedValue = Text
    Else
        ITranslationObject_TranslatedValue = stored
    End If
End Function

Private Sub ITranslationObject_TranslateForm(ByVal UserFrm As Object)
    'No-op for stub
End Sub

Private Sub ITranslationObject_TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
    'No-op for stub
End Sub

Private Property Get ITranslationObject_HasCheckings() As Boolean
    ITranslationObject_HasCheckings = this.hasChecks
End Property

Private Property Get ITranslationObject_CheckingValues() As Object
    'Stub does not produce checking output
End Property

