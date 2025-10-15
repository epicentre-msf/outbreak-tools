VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseTranslationCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Caches translation lookups for repeated disease export keys"

'@Folder("Msetup")
'@ModuleDescription("Caches translation lookups for repeated disease export keys")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDiseaseTranslationCache

'@section State
'===============================================================================

Private Type TCache
    translations As ITranslationObject
    keys As BetterArray
    values As BetterArray
End Type

Private this As TCache

'@section Factory
'===============================================================================

'@sub-title Create a new cache instance and initialise it with translations.
'@param translations Optional ITranslationObject providing translation values.
'@return IDiseaseTranslationCache ready for use.
Public Function Create(Optional ByVal translations As ITranslationObject) As IDiseaseTranslationCache
    With New DiseaseTranslationCache
        .Initialise translations
        Set Create = .Self
    End With
End Function

'@sub-title Return the current instance as its interface.
Public Property Get Self() As IDiseaseTranslationCache
    Set Self = Me
End Property

'@section Configuration
'===============================================================================

'@sub-title Prime the cache with the translation source.
Public Sub Initialise(Optional ByVal translations As ITranslationObject)
    Set this.translations = translations
    Set this.keys = New BetterArray
    Set this.values = New BetterArray
End Sub

'@section Queries
'===============================================================================

'@sub-title Retrieve a cached translation or resolve it from the source.
Public Function Resolve(ByVal key As String, ByVal fallback As String) As String
    Dim cached As String

    ValidateKey key

    If TryGetCached(key, cached) Then
        Resolve = cached
        Exit Function
    End If

    Resolve = Store(key, ResolveTranslation(key, fallback))
End Function

'@section Maintenance
'===============================================================================

'@sub-title Clear cached values and detach translation object.
Public Sub Reset()
    Initialise
End Sub

'@section Internal Helpers
'===============================================================================

'@sub-title Validate incoming key values.
Private Sub ValidateKey(ByVal key As String)
    If LenB(key) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Translation key cannot be empty"
    End If
End Sub

'@sub-title Attempt to read a cached translation value.
Private Function TryGetCached(ByVal key As String, ByRef cachedValue As String) As Boolean
    Dim index As Long

    index = FindKeyIndex(key)
    If index = -1 Then Exit Function

    cachedValue = CStr(this.values.Item(index))
    TryGetCached = True
End Function

'@sub-title Resolve translation value from the backing source.
Private Function ResolveTranslation(ByVal key As String, ByVal fallback As String) As String
    Dim translations As ITranslationObject

    Set translations = this.translations
    If translations Is Nothing Then
        ResolveTranslation = fallback
        Exit Function
    End If

    If Not translations.ValueExists(key) Then
        ResolveTranslation = fallback
        Exit Function
    End If

    ResolveTranslation = translations.TranslatedValue(key)
End Function

'@sub-title Persist the resolved value in the cache and return it.
Private Function Store(ByVal key As String, ByVal value As String) As String
    this.keys.Push LCase$(key)
    this.values.Push value
    Store = value
End Function

Private Function FindKeyIndex(ByVal key As String) As Long
    Dim idx As Long

    FindKeyIndex = -1

    If this.keys Is Nothing Then Exit Function
    If this.keys.Length = 0 Then Exit Function

    For idx = this.keys.LowerBound To this.keys.UpperBound
        If StrComp(CStr(this.keys.Item(idx)), LCase$(key), vbBinaryCompare) = 0 Then
            FindKeyIndex = idx
            Exit Function
        End If
    Next idx
End Function

'@section ErrorHandling
'===============================================================================

'@sub-title Raise ProjectError-based exceptions.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseTranslationCache", errMessage
End Sub

'@section Interface
'===============================================================================

Private Sub IDiseaseTranslationCache_Initialise(Optional ByVal translations As ITranslationObject)
    Initialise translations
End Sub

Private Function IDiseaseTranslationCache_Resolve(ByVal key As String, ByVal fallback As String) As String
    IDiseaseTranslationCache_Resolve = Resolve(key, fallback)
End Function

Private Sub IDiseaseTranslationCache_Reset()
    Reset
End Sub
