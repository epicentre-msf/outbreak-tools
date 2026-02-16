VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistTranslationCounterStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Counts translation lookups while providing deterministic results")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements ITranslationObject

Private Type TTranslation
    values As Collection
    lookupCount As Long
End Type

Private this As TTranslation

Public Sub Initialise()
    Set this.values = New Collection
    this.lookupCount = 0
End Sub

Public Sub SetValue(ByVal key As String, ByVal value As String)
    Dim normalized As String

    normalized = NormalizeKey(key)
    EnsureStorage

    On Error Resume Next
        this.values.Remove normalized
    On Error GoTo 0

    this.values.Add value, normalized
End Sub

Public Property Get LookupCount() As Long
    LookupCount = this.lookupCount
End Property

Private Sub EnsureStorage()
    If this.values Is Nothing Then Set this.values = New Collection
End Sub

Private Function NormalizeKey(ByVal key As String) As String
    NormalizeKey = LCase$(Trim$(key))
End Function

Private Function ResolveValue(ByVal key As String) As String
    Dim normalized As String
    Dim stored As String

    normalized = NormalizeKey(key)
    EnsureStorage

    On Error Resume Next
        stored = CStr(this.values.Item(normalized))
    On Error GoTo 0

    ResolveValue = stored
End Function

Private Property Get ITranslationObject_Name() As String
    ITranslationObject_Name = "TranslationCounterStub"
End Property

Private Property Get ITranslationObject_ValueExists(ByVal text As String) As Boolean
    ITranslationObject_ValueExists = (LenB(ResolveValue(text)) > 0)
End Property

Private Function ITranslationObject_TranslatedValue(ByVal Text As String, Optional ByVal containsFormula As Boolean = False) As String
    Dim stored As String

    stored = ResolveValue(Text)
    this.lookupCount = this.lookupCount + 1

    If LenB(stored) = 0 Then
        ITranslationObject_TranslatedValue = Text
    Else
        ITranslationObject_TranslatedValue = stored
    End If
End Function

Private Sub ITranslationObject_TranslateForm(ByVal UserFrm As Object)
    'Not used in tests
End Sub

Private Sub ITranslationObject_TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
    'Not used in tests
End Sub

Private Property Get ITranslationObject_HasCheckings() As Boolean
    ITranslationObject_HasCheckings = False
End Property

Private Property Get ITranslationObject_CheckingValues() As Object
    Set ITranslationObject_CheckingValues = Nothing
End Property
