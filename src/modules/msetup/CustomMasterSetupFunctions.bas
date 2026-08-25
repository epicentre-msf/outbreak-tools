Attribute VB_Name = "CustomMasterSetupFunctions"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("Worksheet functions filling disease sheet lines from the Variables and Choices sheets.")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed

'The disease table carries two formulas per line, written by DiseaseSheet:
'MainLabelValue answers the label of the picked variable, ChoiceValues answers
'the joined values of the picked choice. Both take the language of the sheet,
'so switching that one cell retranslates every line.
'
'A worksheet function can never raise and can never write. Every resolver
'failure is answered with an empty string, and the resolved managers are
'cached at module level. MasterSetupEventsManager clears the caches when the
'workbook opens or the translations change.

Private Const CHOICE_SEPARATOR As String = " | "

Private cachedVariables As MasterSetupVariables
Private cachedChoices As LLChoices
Private cachedTranslations As Collection

'@section Worksheet functions
'===============================================================================

'@sub-title Answer the label of a variable, translated to the given language.
'@param variableName Variant cell value naming the variable.
'@param languageTag Variant cell value naming the language column.
'@return String label, or an empty string when the variable is unknown.
Public Function MainLabelValue(ByVal variableName As Variant, ByVal languageTag As Variant) As String
    Dim resolvedName As String
    Dim label As String
    Dim trads As TranslationObject

    resolvedName = CleanText(variableName)
    If LenB(resolvedName) = 0 Then Exit Function

    ' A worksheet function answers empty on any resolver failure.
    On Error Resume Next
    label = VariablesManager().LabelFor(resolvedName)
    On Error GoTo 0

    If LenB(label) = 0 Then Exit Function

    Set trads = TranslationsFor(CleanText(languageTag))
    If trads Is Nothing Then
        MainLabelValue = label
    Else
        MainLabelValue = trads.TranslatedValue(label)
    End If
End Function

'@sub-title Answer the joined values of a choice, translated to the given language.
'@param choiceName Variant cell value naming the choice list.
'@param languageTag Variant cell value naming the language column.
'@return String values joined with " | ", or an empty string.
Public Function ChoiceValues(ByVal choiceName As Variant, ByVal languageTag As Variant) As String
    Dim resolvedName As String
    Dim choices As LLChoices

    resolvedName = CleanText(choiceName)
    If LenB(resolvedName) = 0 Then Exit Function

    Set choices = ChoicesManager()
    If choices Is Nothing Then Exit Function

    ' A worksheet function answers empty on any resolver failure.
    On Error Resume Next
    ChoiceValues = choices.ConcatenateCategories(resolvedName, CHOICE_SEPARATOR, _
                                                 TranslationsFor(CleanText(languageTag)))
    On Error GoTo 0
End Function

'@section Cache management
'===============================================================================

'@sub-title Drop every cached manager so the next call re-reads the sheets.
Public Sub ResetMasterSetupFunctionCaches()
    Set cachedVariables = Nothing
    Set cachedChoices = Nothing
    Set cachedTranslations = Nothing
End Sub

'@section Resolvers
'===============================================================================

Private Function VariablesManager() As MasterSetupVariables
    If cachedVariables Is Nothing Then
        Set cachedVariables = MasterSetupHelpers.ResolveMasterSetupVariables()
    End If
    Set VariablesManager = cachedVariables
End Function

Private Function ChoicesManager() As LLChoices
    If cachedChoices Is Nothing Then
        Set cachedChoices = MasterSetupHelpers.ResolveMasterChoices()
    End If
    Set ChoicesManager = cachedChoices
End Function

'@sub-title Answer a translation object for the language, cached per language.
'@details An unknown or empty language answers Nothing, and callers then keep
'the untranslated text. Lookups are kept in a keyed Collection. Dictionary is
'missing on Mac Excel.
Private Function TranslationsFor(ByVal languageTag As String) As TranslationObject
    Dim table As ListObject
    Dim trads As TranslationObject

    If LenB(languageTag) = 0 Then Exit Function

    If cachedTranslations Is Nothing Then Set cachedTranslations = New Collection

    ' A Collection has no existence test; the scoped error probe is the
    ' cheapest one. Restored right after.
    On Error Resume Next
    Set TranslationsFor = cachedTranslations.Item(languageTag)
    On Error GoTo 0
    If Not TranslationsFor Is Nothing Then Exit Function

    Set table = ResolveTranslationsTable()
    If table Is Nothing Then Exit Function

    ' An unknown language makes the factory raise; the caller then keeps the
    ' untranslated text.
    On Error Resume Next
    Set trads = TranslationObject.Create(table, languageTag)
    On Error GoTo 0
    If trads Is Nothing Then Exit Function

    cachedTranslations.Add trads, languageTag
    Set TranslationsFor = trads
End Function

Private Function ResolveTranslationsTable() As ListObject
    Dim translationsSheet As Worksheet

    Set translationsSheet = MasterSetupHelpers.ResolveMasterTranslationsSheet()
    If translationsSheet Is Nothing Then Exit Function
    If translationsSheet.ListObjects.Count = 0 Then Exit Function

    Set ResolveTranslationsTable = translationsSheet.ListObjects(1)
End Function

Private Function CleanText(ByVal candidate As Variant) As String
    If IsError(candidate) Then Exit Function
    CleanText = Trim$(CStr(candidate))
End Function
