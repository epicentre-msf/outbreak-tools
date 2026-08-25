VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseTranslationCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Caches translated header labels for disease exports"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Caches translated header labels for disease exports")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'DiseaseTranslationCache.cls without interface -- single implementation, no
'test fake, no compile cycle (see project-rules.md 4.4).
'
'A thin wrapper over TranslationObject. The first Resolve of a key reads the
'translation table; the answer itself is cached in a keyed Collection, so
'later edits of the table never change what an export already resolved.
'Lookups are kept in a keyed Collection. Dictionary is missing on Mac Excel.

'@section Types
'===============================================================================

Private Type TTranslationCache
    Translations As TranslationObject
    Answers As Collection
End Type

Private this As TTranslationCache

'@section Factory
'===============================================================================

'@sub-title Build a cache over the given translation object.
'@param translations TranslationObject supplying the translated labels.
'@return DiseaseTranslationCache ready to resolve headers.
Public Function Create(ByVal translations As TranslationObject) As DiseaseTranslationCache
    Dim cache As DiseaseTranslationCache

    Set cache = New DiseaseTranslationCache
    cache.Initialise translations
    Set Create = cache
End Function

'@section Configuration
'===============================================================================

'@sub-title Attach the translation object backing this cache.
Public Sub Initialise(ByVal translations As TranslationObject)
    If translations Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Translation object cannot be Nothing"
    End If

    Set this.Translations = translations
    Set this.Answers = New Collection
End Sub

'@section Queries
'===============================================================================

'@sub-title Resolve a header key to its translation, remembering the answer.
'@details A key found in the table is answered from the cache from then on.
'A key the table does not carry answers the supplied fallback and stays
'uncached, so it can pick the translation up if the table gains it later.
'@param key String tag to translate.
'@param fallback String answered when the table does not carry the key.
'@return String translated label, or the fallback.
Public Function Resolve(ByVal key As String, ByVal fallback As String) As String
    Dim answer As String

    If LenB(key) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Translation key cannot be empty"
    End If

    If TryCached(key, answer) Then
        Resolve = answer
        Exit Function
    End If

    If this.Translations.ValueExists(key) Then
        answer = this.Translations.TranslatedValue(key)
        this.Answers.Add answer, key
        Resolve = answer
    Else
        Resolve = fallback
    End If
End Function

'@section Internal Helpers
'===============================================================================

'@sub-title Read a cached answer, reporting whether it was there.
Private Function TryCached(ByVal key As String, ByRef answer As String) As Boolean
    ' A Collection has no existence test; the scoped error probe is the
    ' cheapest one. Restored right after.
    On Error Resume Next
    answer = this.Answers.Item(key)
    TryCached = (Err.Number = 0)
    On Error GoTo 0
End Function

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised error raising aligned with ProjectError codes.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseTranslationCache", errMessage
End Sub
