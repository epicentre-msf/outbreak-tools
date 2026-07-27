VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TranslationObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Translate various objects from one language to another"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Translate various objects from one language to another")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation
'@depends BetterArray, Checking


Option Explicit

'@class TranslationObject
'@description
'Tag-based translator backed by a ListObject table. Looks up tags in
'the first column and returns the value from the active language column.
'Supports plain text, formula-aware translation (quoted chunks only),
'range bulk-translation, and user-form caption translation.
'Each public call reads the table once into memory and matches every
'value against that copy, so a call costs the same two worksheet reads
'whatever the number of values it translates.
'@depends BetterArray, Checking
'@version 1.0 (2026-02-09)


Private Const CLASSNAME As String = "TranslationObject"

Private Type TTranslationObject
    transLo As ListObject
    transLng As String
    hasChecks As Boolean
    sealed As Boolean
    lookupTags() As String
    lookupValues() As String
    lookupCount As Long
End Type

Private this As TTranslationObject

'checking objects
Private checkCounter As Long
Private internalChecks As Checking

'@section Factory
'===============================================================================
'Initialisation and self-reference.

'@label:create
'@sub-title Factory returning an initialised interface
'@details
'A translation object wraps a ListObject whose first column contains tags
'and remaining columns hold language translations. Tags not found during
'translation are returned as-is.
'@param translationLo ListObject. The translation table.
'@param translationLang String. The target language column header.
'@return TranslationObject. Ready to use.
'@export
Public Function Create(ByVal translationLo As ListObject, ByVal translationLang As String) As TranslationObject

    ValidateListObject translationLo, "Create"

    Dim newTranslation As TranslationObject
    Set newTranslation = New TranslationObject

    Set newTranslation.Table = translationLo
    newTranslation.Language = Trim$(translationLang)
    newTranslation.Seal
    Set Create = newTranslation
End Function

'@section Elements
'===============================================================================
'Core translation accessors and lookups.

'@label:table
'@prop-title The backing ListObject translation table
'@return ListObject. The translation table.
Public Property Get Table() As ListObject
    Set Table = this.transLo
End Property

'@label:table-set
'@prop-title Assign the backing ListObject
'@param lo ListObject. The translation table.
Public Property Set Table(ByVal lo As ListObject)
    GuardNotSealed "Table"
    ValidateListObject lo, "Table"
    Set this.transLo = lo
End Property

'@label:language
'@prop-title Current target language
'@return String. One of the language column headers.
'@export
Public Property Get Language() As String
    Language = this.transLng
End Property

'@label:languages-list
'@prop-title List of available language headers
'@return BetterArray. Each language header, excluding the helper/tag column.
'@export
Public Property Get LanguagesList() As BetterArray
    Dim headerRow As Range
    Dim headerBlock As Variant
    Dim languages As BetterArray
    Dim idx As Long
    Dim valueText As String

    Set languages = New BetterArray
    languages.LowerBound = 1

    'Read the header row afresh on every call: columns can be added or
    'removed from the table while this object is alive.
    Set headerRow = HeaderRowRange()

    If Not headerRow Is Nothing Then
        headerBlock = RangeBlock(headerRow.Offset(, 1))
        For idx = LBound(headerBlock, 2) To UBound(headerBlock, 2)
            valueText = Trim$(TextOf(headerBlock(1, idx)))
            If LenB(valueText) > 0 Then languages.Push valueText
        Next idx
    End If

    Set LanguagesList = languages
End Property


'@label:language-set
'@prop-title Set the target language
'@param transLng String. The new language column header.
Public Property Let Language(ByVal transLng As String)
    GuardNotSealed "Language"
    this.transLng = Trim$(transLng)
End Property

'@label:name
'@prop-title Descriptive name including table and language
'@return String. Identifier for diagnostic output.
'@export
Public Property Get Name() As String
    If this.transLo Is Nothing Then
        Name = CLASSNAME & " (uninitialised)"
    Else
        Name = "Table: " & this.transLo.Name & " - Language: " & Language
    End If
End Property

'@label:value-exists
'@prop-title Whether a tag exists in the translation table
'@param Text String. Tag to search.
'@return Boolean. True when the tag is found.
'@export
Public Function ValueExists(ByVal Text As String) As Boolean
    Dim sanitized As String

    RefreshLookup

    If LookupRow(Text) > 0 Then
        ValueExists = True
        Exit Function
    End If

    'Second chance on the cleaned text, in case the caller passed a raw value.
    sanitized = SanitizeText(Text)
    If StrComp(sanitized, Text, vbBinaryCompare) <> 0 Then
        ValueExists = (LookupRow(sanitized) > 0)
    End If
End Function

'@label:translated-value
'@sub-title Translate a text tag or formula string
'@details
'When containsFormula is True, only double-quoted chunks inside the text
'are translated; unquoted characters pass through unchanged.
'@param textToTranslate String. Text to translate.
'@param containsFormula Boolean. Translate as formula. Defaults to False.
'@return String. The translated value, or the original text when not found.
'@export
Public Function TranslatedValue(ByVal textToTranslate As String, _
                                Optional ByVal containsFormula As Boolean = False) As String

    RefreshLookup
    TranslatedValue = TranslateText(textToTranslate, containsFormula)
End Function

'@section Outside
'===============================================================================
'Bulk translation of forms and ranges.

'@label:translate-range
'@sub-title Translate all values in a column range
'@param colRng Range. Column range whose values should be translated.
'@param containsFormula Boolean. When True, translates as formulas. Defaults to False.
'@export
Public Sub TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
    Dim targetColumn As Range
    Dim cellValues As Variant
    Dim rowCount As Long
    Dim idx As Long

    If colRng Is Nothing Then Exit Sub

    Set targetColumn = colRng.Columns(1)
    rowCount = targetColumn.Cells.Count
    If rowCount = 0 Then Exit Sub

    RefreshLookup

    'One read, translate in memory, one write back.
    cellValues = RangeBlock(targetColumn)
    For idx = 1 To rowCount
        cellValues(idx, 1) = TranslateText(TextOf(cellValues(idx, 1)), containsFormula)
    Next idx

    targetColumn.Value = cellValues
End Sub

'@label:translate-form
'@sub-title Translate all supported controls on a user form
'@param UserFrm Object. Any user form exposing a Controls collection.
'@export
Public Sub TranslateForm(ByVal UserFrm As Object)
    Dim formControl As Control
    Dim controlType As String

    If Not SupportsUserFormContract(UserFrm) Then
        LogInfo "Attempt to translate an object of type '" & TypeName(UserFrm) & "' that is not a form", checkingNote
        Exit Sub
    End If

    RefreshLookup

    For Each formControl In UserFrm.Controls
        controlType = TypeName(formControl)
        If StrComp(controlType, "MultiPage", vbTextCompare) = 0 Then
            TranslateMultiPageControl formControl
        ElseIf ShouldTranslateControl(controlType) Then
            If Trim$(CStr(formControl.Caption)) <> vbNullString Then
                formControl.Caption = TranslateTag(formControl.Name)
            End If
        Else
            LogInfo "Attempt to translate a control that is not supported :'" & controlType & "'", checkingNote
        End If
    Next formControl
End Sub

'@section Checkings
'===============================================================================
'Diagnostic logging.

'@label:log-info
'@sub-title Record a diagnostic entry
'@param label String. The message to add.
'@param scope Byte. The checking scope. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "Translation " & Name() & " object trace"

    'First, instanciante the internalCheks Objects
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)

    'Add new informations to the checking object
    internalChecks.Add CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:has-checkings
'@prop-title Whether diagnostic entries have been recorded
'@return Boolean. True when at least one entry exists.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Object. The checking instance, or Nothing.
'@export
Public Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
End Property

'@section Helpers
'===============================================================================
'Internal utilities.

'@label:sanitize-text
'@sub-title Clean a raw value the way the worksheet functions used to
'@details
'Does the same three jobs as SUBSTITUTE + CLEAN + TRIM did, in plain VBA
'so a call no longer goes out to Excel: the '<' character becomes a space,
'characters below code 32 are dropped, then leading, trailing and repeated
'spaces are squeezed. Worksheet TRIM squeezes inner runs of spaces down to
'one, which VBA Trim$ does not do, so that part is written out here.
'@param rawValue String. The raw text.
'@return String. Cleaned text.
Private Function SanitizeText(ByVal rawValue As String) As String
    Dim substituted As String
    Dim cleaned As String
    Dim currentChar As String
    Dim charCode As Long
    Dim idx As Long
    Dim cleanedLength As Long
    Dim pendingSpace As Boolean

    substituted = Replace(rawValue, Chr$(60), " ", Compare:=vbBinaryCompare)
    If LenB(substituted) = 0 Then Exit Function

    cleaned = Space$(Len(substituted))

    For idx = 1 To Len(substituted)
        currentChar = Mid$(substituted, idx, 1)
        charCode = AscW(currentChar)

        If charCode >= 0 And charCode < 32 Then
            'dropped, as CLEAN does
        ElseIf charCode = 32 Then
            'a space is only kept when a printable character follows it
            If cleanedLength > 0 Then pendingSpace = True
        Else
            If pendingSpace Then
                cleanedLength = cleanedLength + 1
                Mid$(cleaned, cleanedLength, 1) = " "
                pendingSpace = False
            End If
            cleanedLength = cleanedLength + 1
            Mid$(cleaned, cleanedLength, 1) = currentChar
        End If
    Next idx

    SanitizeText = Left$(cleaned, cleanedLength)
End Function

'@label:header-row-range
'@sub-title Retrieve the header row range of the translation table
'@return Range. The header row, or Nothing.
Private Function HeaderRowRange() As Range
    On Error Resume Next
    If Not this.transLo Is Nothing Then
        Set HeaderRowRange = this.transLo.HeaderRowRange
    End If
    On Error GoTo 0
End Function

'@label:range-block
'@sub-title Read a range into a two dimensional array in one go
'@details
'Range.Value hands back a plain value for a single cell, so that case is
'wrapped to keep one shape for every caller.
'@param target Range. The range to read.
'@return Variant. A two dimensional array, or Empty when target is Nothing.
Private Function RangeBlock(ByVal target As Range) As Variant
    Dim oneCell(1 To 1, 1 To 1) As Variant

    If target Is Nothing Then Exit Function

    If target.Cells.Count = 1 Then
        oneCell(1, 1) = target.Value
        RangeBlock = oneCell
    Else
        RangeBlock = target.Value
    End If
End Function

'@label:text-of
'@sub-title Turn a cell value into text without ever raising
'@param cellValue Variant. The value read from a cell.
'@return String. The value as text, empty for errors and blanks.
Private Function TextOf(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    TextOf = CStr(cellValue)
End Function

'@label:language-index
'@sub-title Resolve the one-based column index of the active language
'@return Long. Column index, or -1 when not found.
Private Function LanguageIndex() As Long
    Dim headerRow As Range
    Dim headerBlock As Variant
    Dim idx As Long

    LanguageIndex = -1

    If this.transLo Is Nothing Then Exit Function
    If LenB(this.transLng) = 0 Then Exit Function

    Set headerRow = HeaderRowRange()
    If headerRow Is Nothing Then Exit Function

    headerBlock = RangeBlock(headerRow)

    'the match is exact and case sensitive, as the worksheet search was
    For idx = LBound(headerBlock, 2) To UBound(headerBlock, 2)
        If StrComp(TextOf(headerBlock(1, idx)), this.transLng, vbBinaryCompare) = 0 Then
            LanguageIndex = idx
            Exit Function
        End If
    Next idx
End Function

'@label:refresh-lookup
'@sub-title Read the tags and the active language column into memory
'@details
'Called at the start of every public translation call. The table can be
'edited while this object is alive, so the copy is never kept between
'calls; inside a call it serves every value.
Private Sub RefreshLookup()
    Dim bodyRange As Range
    Dim bodyBlock As Variant
    Dim languageColumn As Long
    Dim rowCount As Long
    Dim idx As Long

    this.lookupCount = 0

    If this.transLo Is Nothing Then Exit Sub

    languageColumn = LanguageIndex()
    If languageColumn <= 0 Then Exit Sub

    Set bodyRange = this.transLo.DataBodyRange
    If bodyRange Is Nothing Then Exit Sub

    bodyBlock = RangeBlock(bodyRange)
    If languageColumn > UBound(bodyBlock, 2) Then Exit Sub

    rowCount = UBound(bodyBlock, 1)
    ReDim this.lookupTags(1 To rowCount)
    ReDim this.lookupValues(1 To rowCount)

    For idx = 1 To rowCount
        this.lookupTags(idx) = TextOf(bodyBlock(idx, 1))
        this.lookupValues(idx) = TextOf(bodyBlock(idx, languageColumn))
    Next idx

    this.lookupCount = rowCount
End Sub

'@label:lookup-row
'@sub-title Find the row of a tag in the copy held in memory
'@param token String. The tag to look up.
'@return Long. The one-based row, or 0 when the tag is not there.
Private Function LookupRow(ByVal token As String) As Long
    Dim idx As Long

    If LenB(token) = 0 Then Exit Function

    'exact and case sensitive, as the worksheet search was
    For idx = 1 To this.lookupCount
        If StrComp(this.lookupTags(idx), token, vbBinaryCompare) = 0 Then
            LookupRow = idx
            Exit Function
        End If
    Next idx
End Function

'@label:translate-text
'@sub-title Translate one value against the copy already in memory
'@param textToTranslate String. Text to translate.
'@param containsFormula Boolean. Translate as formula.
'@return String. The translated value.
Private Function TranslateText(ByVal textToTranslate As String, _
                               ByVal containsFormula As Boolean) As String

    If containsFormula Then
        TranslateText = TranslateFormulaText(SanitizeText(textToTranslate))
        Exit Function
    End If

    TranslateText = TranslateTag(textToTranslate)
End Function

'@label:translate-tag
'@sub-title Return the translation of one tag, or the cleaned tag itself
'@param token String. The tag to look up.
'@return String. The translation, or the cleaned token when there is none.
Private Function TranslateTag(ByVal token As String) As String
    Dim matchRow As Long
    Dim translated As String

    matchRow = LookupRow(token)

    If matchRow > 0 Then
        translated = SanitizeText(this.lookupValues(matchRow))
        'an empty translation cell falls back to the text to translate
        If LenB(translated) > 0 Then
            TranslateTag = translated
            Exit Function
        End If
    End If

    TranslateTag = SanitizeText(token)
End Function

'@label:translate-formula-text
'@sub-title Translate only the double-quoted chunks within a formula string
'@param formulaText String. The formula text.
'@return String. The formula with quoted chunks translated.
Private Function TranslateFormulaText(ByVal formulaText As String) As String
    Dim strippedText As String
    Dim workingText As String
    Dim replacements As BetterArray
    Dim quoteStart As Long
    Dim idx As Long
    Dim chunk As String
    Dim translated As String
    Dim DOUBLE_QUOTE As String 'A constant

    DOUBLE_QUOTE = Chr$(34)

    workingText = formulaText
    strippedText = Replace(formulaText, DOUBLE_QUOTE & DOUBLE_QUOTE, vbNullString)

    If InStr(1, strippedText, DOUBLE_QUOTE, vbBinaryCompare) = 0 Then
        TranslateFormulaText = workingText
        Exit Function
    End If

    Set replacements = New BetterArray

    For idx = 1 To Len(strippedText)
        If Mid$(strippedText, idx, 1) = DOUBLE_QUOTE Then
            If quoteStart = 0 Then
                quoteStart = idx + 1
            Else
                chunk = Mid$(strippedText, quoteStart, idx - quoteStart)
                If LenB(chunk) > 0 Then
                    If Not replacements.Includes(chunk) Then
                        translated = TranslateTag(chunk)
                        workingText = Replace(workingText, DOUBLE_QUOTE & chunk & DOUBLE_QUOTE, DOUBLE_QUOTE & translated & DOUBLE_QUOTE)
                        replacements.Push chunk
                    End If
                End If
                quoteStart = 0
            End If
        End If
    Next idx

    TranslateFormulaText = workingText
End Function

'@label:supports-user-form-contract
'@sub-title Check whether an object exposes a Controls collection
'@param candidate Object. The object to test.
'@return Boolean. True when a Controls property is accessible.
Private Function SupportsUserFormContract(ByVal candidate As Object) As Boolean
    Dim controls As Object

    If candidate Is Nothing Then Exit Function

    On Error Resume Next
    Set controls = candidate.Controls
    SupportsUserFormContract = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'@label:should-translate-control
'@sub-title Determine whether a control type supports caption translation
'@param controlType String. TypeName of the control.
'@return Boolean. True for CommandButton, Label, OptionButton, Page, Frame, CheckBox.
Private Function ShouldTranslateControl(ByVal controlType As String) As Boolean
    Select Case LCase$(controlType)
        Case "commandbutton", "label", "optionbutton", "page", "frame", "checkbox"
            ShouldTranslateControl = True
    End Select
End Function

'@label:translate-multi-page-control
'@sub-title Translate page captions within a MultiPage control
'@param multiPageControl Object. The MultiPage control.
Private Sub TranslateMultiPageControl(ByVal multiPageControl As Object)
    Dim pageIndex As Long
    Dim controlName As String

    If multiPageControl Is Nothing Then Exit Sub

    controlName = CStr(multiPageControl.Name)

    If controlName <> "MultiPage1" And controlName <> "MultiPage2" Then Exit Sub

    For pageIndex = 0 To multiPageControl.Pages.Count - 1
        multiPageControl.Pages(pageIndex).Caption = TranslateTag(multiPageControl.Pages(pageIndex).Name)
    Next pageIndex
End Sub

'@label:validate-list-object
'@sub-title Guard against Nothing ListObject arguments
'@param source ListObject. The table to validate.
'@param sourceMember String. Calling member name for error context.
Private Sub ValidateListObject(ByVal source As ListObject, ByVal sourceMember As String)
    If source Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "ListObject is not defined.", sourceMember
    End If
End Sub

'@label:raise-error
'@sub-title Raise a typed project error with optional member context
'@param errNumber ProjectError. The error code.
'@param message String. Descriptive error message.
'@param sourceMember String. Optional calling member name.
Private Sub RaiseError(ByVal errNumber As ProjectError, ByVal message As String, Optional ByVal sourceMember As String = vbNullString)
    Dim sourceName As String

    sourceName = CLASSNAME
    If LenB(sourceMember) > 0 Then
        sourceName = sourceName & "." & sourceMember
    End If

    Err.Raise CLng(errNumber), sourceName, message
End Sub

'@label:seal
'@sub-title Seal the instance so setup setters can no longer be written
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against writes after sealing
'@param propName String. Name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        RaiseError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub
