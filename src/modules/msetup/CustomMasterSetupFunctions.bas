Attribute VB_Name = "CustomMasterSetupFunctions"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("Worksheet functions filling disease sheet lines from the Variables and Choices sheets.")
'@depends MasterSetupHelpers, MasterSetupVariables, TranslationObject
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
'
'THE CHOICES ARE JOINED IN MEMORY
'-------------------------------------------------------------------------------
'Excel refuses a change of the sheet from inside a calculation, and
'LLChoices reads the categories of a list through an AutoFilter. Called from
'a formula, that filter fails without a word and every label of the sheet
'comes back joined. So the two columns of the Choices block, list name and
'label, are read once into memory and the join runs there.

Private Const CHOICE_SEPARATOR As String = " | "
Private Const CHOICES_HEADER_ROW As Long = 4
Private Const CHOICES_START_COLUMN As Long = 1
Private Const HEADER_LIST_NAME As String = "list name"
Private Const HEADER_LABEL As String = "label"

Private cachedVariables As MasterSetupVariables
Private cachedTranslations As Collection
Private cachedChoiceNames As Variant       'One column: the list name of every Choices line
Private cachedChoiceLabels As Variant      'One column: the label of every Choices line
Private cachedChoiceRows As Long           'Lines the two blocks carry
Private cachedChoiceLoaded As Boolean      'The block was read once, whatever it answered

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

    ' The untranslated label is the answer from here on. A fault in the
    ' translations table then costs the translation, never the label: the
    ' resolver ran outside a guard once and the cell read #VALUE!.
    MainLabelValue = label

    On Error Resume Next
    Set trads = TranslationsFor(CleanText(languageTag))
    On Error GoTo 0
    If trads Is Nothing Then Exit Function

    MainLabelValue = TranslatedOrSame(trads, label)
End Function

'@sub-title Answer the joined values of a choice, translated to the given language.
'@details The disease sheets pass their language cell. The Variables sheet
'passes no language: the master setup is in English, and the function then
'answers the labels of the choice as the Choices sheet carries them.
'@param choiceName Variant cell value naming the choice list.
'@param languageTag Optional Variant cell value naming the language column.
'@return String values joined with " | ", or an empty string.
Public Function ChoiceValues(ByVal choiceName As Variant, Optional ByVal languageTag As Variant) As String
    Dim resolvedName As String
    Dim resolvedLanguage As String
    Dim trads As TranslationObject

    resolvedName = CleanText(choiceName)
    If LenB(resolvedName) = 0 Then Exit Function

    If Not IsMissing(languageTag) Then resolvedLanguage = CleanText(languageTag)

    ' A worksheet function answers empty on any resolver failure. The
    ' translator is resolved on a guard of its own: resolved inside the join,
    ' a fault there abandoned the join and the cell came back empty.
    On Error Resume Next
    Set trads = TranslationsFor(resolvedLanguage)
    On Error GoTo 0

    On Error Resume Next
    ChoiceValues = JoinedLabels(resolvedName, trads)
    On Error GoTo 0
End Function

'@section Cache management
'===============================================================================

'@sub-title Drop every cached manager so the next call re-reads the sheets.
Public Sub ResetMasterSetupFunctionCaches()
    Set cachedVariables = Nothing
    Set cachedTranslations = Nothing
    cachedChoiceNames = Empty
    cachedChoiceLabels = Empty
    cachedChoiceRows = 0
    cachedChoiceLoaded = False
End Sub

'@section Resolvers
'===============================================================================

Private Function VariablesManager() As MasterSetupVariables
    If cachedVariables Is Nothing Then
        Set cachedVariables = MasterSetupHelpers.ResolveMasterSetupVariables()
    End If
    Set VariablesManager = cachedVariables
End Function

'@sub-title Join the labels of one list, in sheet order, translated when asked.
Private Function JoinedLabels(ByVal listName As String, ByVal trads As TranslationObject) As String
    Dim idx As Long
    Dim label As String
    Dim joined As String

    EnsureChoiceBlock

    For idx = 1 To cachedChoiceRows
        If StrComp(CleanText(cachedChoiceNames(idx, 1)), listName, vbTextCompare) = 0 Then
            label = CleanText(cachedChoiceLabels(idx, 1))
            If Not trads Is Nothing Then label = TranslatedOrSame(trads, label)
            If LenB(joined) > 0 Then joined = joined & CHOICE_SEPARATOR
            joined = joined & label
        End If
    Next idx

    JoinedLabels = joined
End Function

'@sub-title The translated label, or the label itself when the translation faults.
'@details A translations table the translator cannot read is a fault of the
'workbook, not of the line being written. The line then carries the label as
'it is typed, which is what a master setup in English reads anyway.
Private Function TranslatedOrSame(ByVal trads As TranslationObject, ByVal label As String) As String
    TranslatedOrSame = label

    On Error Resume Next
    TranslatedOrSame = trads.TranslatedValue(label)
    On Error GoTo 0
End Function

'@sub-title Read the list name and label columns of the Choices sheet once.
'@details The master file carries its choices in a table; a sheet without
'one is read as the block under the header row. A sheet missing either
'header leaves the block empty, and every join answers empty.
Private Sub EnsureChoiceBlock()
    Dim choicesSheet As Worksheet
    Dim block As Range
    Dim headers As Variant
    Dim nameIndex As Long
    Dim labelIndex As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim bodyRows As Long

    If cachedChoiceLoaded Then Exit Sub
    cachedChoiceLoaded = True
    cachedChoiceRows = 0

    Set choicesSheet = MasterSetupHelpers.ResolveMasterChoicesSheet()
    If choicesSheet Is Nothing Then Exit Sub

    If choicesSheet.ListObjects.Count > 0 Then
        Set block = choicesSheet.ListObjects(1).Range
    Else
        lastRow = choicesSheet.Cells(choicesSheet.Rows.Count, CHOICES_START_COLUMN).End(xlUp).Row
        lastColumn = choicesSheet.Cells(CHOICES_HEADER_ROW, choicesSheet.Columns.Count).End(xlToLeft).Column
        If lastRow <= CHOICES_HEADER_ROW Then Exit Sub
        If lastColumn < CHOICES_START_COLUMN Then Exit Sub
        Set block = choicesSheet.Range(choicesSheet.Cells(CHOICES_HEADER_ROW, CHOICES_START_COLUMN), _
                                       choicesSheet.Cells(lastRow, lastColumn))
    End If

    bodyRows = block.Rows.Count - 1
    If bodyRows < 1 Then Exit Sub
    If block.Columns.Count < 2 Then Exit Sub

    headers = block.Rows(1).Value2
    nameIndex = HeaderIndex(headers, HEADER_LIST_NAME)
    labelIndex = HeaderIndex(headers, HEADER_LABEL)
    If nameIndex = 0 Or labelIndex = 0 Then Exit Sub

    cachedChoiceNames = ColumnBlock(block.Columns(nameIndex).Offset(1).Resize(bodyRows))
    cachedChoiceLabels = ColumnBlock(block.Columns(labelIndex).Offset(1).Resize(bodyRows))
    cachedChoiceRows = bodyRows
End Sub

'@sub-title The index of a header in a one-row block, 0 when it is missing.
Private Function HeaderIndex(ByRef headers As Variant, ByVal headerName As String) As Long
    Dim idx As Long

    For idx = 1 To UBound(headers, 2)
        If StrComp(CleanText(headers(1, idx)), headerName, vbTextCompare) = 0 Then
            HeaderIndex = idx
            Exit Function
        End If
    Next idx
End Function

'@sub-title Read a one-column range as a 2D block, whatever its height.
Private Function ColumnBlock(ByVal columnRange As Range) As Variant
    Dim values As Variant
    Dim scalar(1 To 1, 1 To 1) As Variant

    values = columnRange.Value2
    If IsArray(values) Then
        ColumnBlock = values
    Else
        scalar(1, 1) = values
        ColumnBlock = scalar
    End If
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
