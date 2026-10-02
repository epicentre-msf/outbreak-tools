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
'@depends BetterArray, Checking, IChecking, ProjectError


Option Explicit
Implements ITranslationObject

'@class TranslationObject
'@description
'Tag-based translator backed by a ListObject table. Looks up tags in
'the first column and returns the value from the active language column.
'Supports plain text, formula-aware translation (quoted chunks only),
'range bulk-translation, and user-form caption translation.
'@depends BetterArray, Checking, IChecking, ProjectError
'@version 1.0 (2026-02-09)


Private Const CLASSNAME As String = "TranslationObject"

Private Type TTranslationObject
    transLo As ListObject
    transLng As String
    hasChecks As Boolean
    languageColumnIndex As Long
End Type

Private this As TTranslationObject

'checking objects
Private checkCounter As Long
Private internalChecks As IChecking

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
'@return ITranslationObject. Ready to use.
'@export
Public Function Create(ByVal translationLo As ListObject, ByVal translationLang As String) As ITranslationObject

    ValidateListObject translationLo, "Create"

    With New TranslationObject
        Set .Table = translationLo
        .Language = Trim$(translationLang)
        Set Create = .Self
    End With
End Function

'@section Elements
'===============================================================================
'Core translation accessors and lookups.

'@label:self
'@prop-title Self-reference as ITranslationObject
'@return ITranslationObject. Self-reference.
Public Property Get Self() As ITranslationObject
    Set Self = Me
End Property

'@label:table
'@prop-title The backing ListObject translation table
'@return ListObject. The translation table.
Public Property Get Table() As ListObject
    Set Table = this.transLo
End Property

'@label:table-set
'@prop-title Assign the backing ListObject and reset caches
'@param lo ListObject. The translation table.
Public Property Set Table(ByVal lo As ListObject)
    ValidateListObject lo, "Table"
    Set this.transLo = lo
    ResetCaches
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
    Dim languages As BetterArray
    Dim idx As Long
    Dim valueText As String

    Set languages = New BetterArray
    languages.LowerBound = 1

    Set headerRow = HeaderRowRange()
    If headerRow Is Nothing Then
        Set LanguagesList = languages
        Exit Property
    End If

    languages.FromExcelRange headerRow.Offset(, 1), DetectLastRow:=False, DetectLastColumn:=False

    For idx = languages.LowerBound To languages.UpperBound
        valueText = Trim$(CStr(languages.Item(idx)))
        languages.Item(idx) = valueText
    Next idx

    languages.Filter vbNullString, Include:=False

    Set LanguagesList = languages
End Property


'@label:language-set
'@prop-title Set the target language and reset caches
'@param transLng String. The new language column header.
Public Property Let Language(ByVal transLng As String)
    this.transLng = Trim$(transLng)
    ResetCaches
End Property

'@label:name
'@prop-title Descriptive name including table and language
'@return String. Identifier for diagnostic output.
'@export
Private Property Get Name() As String
    If this.transLo Is Nothing Then
        Name = CLASSNAME & " (uninitialised)"
    Else
        Name = "Table: " & this.transLo.Name & " - Language: " & Language
    End If
End Property

'@label:language-index
'@sub-title Resolve the one-based column index of the active language
'@return Long. Column index, or -1 when not found.
Private Function LanguageIndex() As Long

    Dim headerRow As Range
    Dim found As Range
    Dim currentName As String

    If this.transLo Is Nothing Then Exit Function
    If LenB(this.transLng) = 0 Then Exit Function

    If this.languageColumnIndex > 0 Then
        On Error Resume Next
        currentName = this.transLo.ListColumns(this.languageColumnIndex).Name
        On Error GoTo 0
        If StrComp(currentName, this.transLng, vbBinaryCompare) <> 0 Then
            this.languageColumnIndex = 0
        End If
    End If

    If this.languageColumnIndex <= 0 Then
        this.languageColumnIndex = -1
        Set headerRow = HeaderRowRange()
        If headerRow Is Nothing Then
            LanguageIndex = this.languageColumnIndex
            Exit Function
        End If
        Set found = headerRow.Find(What:=this.transLng, LookAt:=xlWhole, MatchCase:=True)
        If Not found Is Nothing Then
            this.languageColumnIndex = found.Column - headerRow.Column + 1
        End If
    End If

    LanguageIndex = this.languageColumnIndex
End Function

'@label:value-exists
'@prop-title Whether a tag exists in the translation table
'@param Text String. Tag to search.
'@return Boolean. True when the tag is found.
'@export
Private Function ValueExists(ByVal Text As String) As Boolean
    Dim translationCell As Range
    Dim sanitized As String

    Set translationCell = ResolveTranslationCell(Text)

    If translationCell Is Nothing Then
        sanitized = SanitizeText(Text)
        If StrComp(sanitized, Text, vbBinaryCompare) <> 0 Then
            Set translationCell = ResolveTranslationCell(sanitized)
        End If
    End If

    ValueExists = Not translationCell Is Nothing
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
Private Function TranslatedValue(ByVal textToTranslate As String, _ 
                                 Optional ByVal containsFormula As Boolean = False) As String

    Dim cleanedText As String
    Dim translationCell As Range
    Dim returnedValue As String

    cleanedText = SanitizeText(textToTranslate)

    If containsFormula Then
        TranslatedValue = TranslateFormulaText(cleanedText)
        Exit Function
    End If

    Set translationCell = ResolveTranslationCell(textToTranslate)

    If translationCell Is Nothing Then
        TranslatedValue = cleanedText
    Else
        returnedValue = SanitizeText(CStr(translationCell.Value))
        'If returned value is an empty string, return the text to translate
        TranslatedValue = IIF(LenB(returnedValue) = 0, cleanedText, returnedValue)
    End If
End Function

'@section Outside
'===============================================================================
'Bulk translation of forms and ranges.

'@label:translate-range
'@sub-title Translate all values in a column range
'@param colRng Range. Column range whose values should be translated.
'@param containsFormula Boolean. When True, translates as formulas. Defaults to False.
'@export
Private Sub TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
    Dim targetCell As Range

    If colRng Is Nothing Then Exit Sub

    For Each targetCell In colRng.Columns(1).Cells
        targetCell.Value = TranslatedValue(targetCell.Value, containsFormula:=containsFormula)
    Next targetCell
End Sub

'@label:translate-form
'@sub-title Translate all supported controls on a user form
'@param UserFrm Object. Any user form exposing a Controls collection.
'@export
Private Sub TranslateForm(ByVal UserFrm As Object)
    Dim formControl As Control
    Dim controlType As String

    If Not SupportsUserFormContract(UserFrm) Then
        LogInfo "Attempt to translate an object of type '" & TypeName(UserFrm) & "' that is not a form", checkingNote
        Exit Sub
    End If

    For Each formControl In UserFrm.Controls
        controlType = TypeName(formControl)
        If StrComp(controlType, "MultiPage", vbTextCompare) = 0 Then
            TranslateMultiPageControl formControl
        ElseIf ShouldTranslateControl(controlType) Then
            If Trim$(CStr(formControl.Caption)) <> vbNullString Then
                formControl.Caption = TranslatedValue(formControl.Name)
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
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Object. The checking instance, or Nothing.
'@export
Private Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
End Property

'@label:show-debug
'@sub-title Print diagnostic information to the Immediate window
Private Sub ShowDebug()
    Debug.Print "Check the informations of Translation object " & Name()
End Sub

'@section Helpers
'===============================================================================
'Internal utilities.

'@label:sanitize-text
'@sub-title Clean a raw value using worksheet functions
'@param rawValue String. The raw text.
'@return String. Cleaned text.
Private Function SanitizeText(ByVal rawValue As String) As String
    Dim cleaned As String

    cleaned = Application.WorksheetFunction.SUBSTITUTE(rawValue, Chr$(60), " ")
    cleaned = Application.WorksheetFunction.CLEAN(cleaned)
    cleaned = Application.WorksheetFunction.TRIM(cleaned)

    SanitizeText = cleaned
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

'@label:keys-data-range
'@sub-title Retrieve the data body range of the first (tags) column
'@return Range. The keys column data, or Nothing.
Private Function KeysDataRange() As Range
    On Error Resume Next
    If Not this.transLo Is Nothing Then
        Set KeysDataRange = this.transLo.ListColumns(1).DataBodyRange
    End If
    On Error GoTo 0
End Function

'@label:resolve-translation-cell
'@sub-title Find the translated cell for a given tag token
'@param token String. The tag to look up.
'@return Range. The cell in the language column, or Nothing.
Private Function ResolveTranslationCell(ByVal token As String) As Range
    Dim keyRange As Range
    Dim found As Range
    Dim bodyRange As Range
    Dim relativeRow As Long
    Dim languageColumn As Long

    If LenB(token) = 0 Then Exit Function
    If this.transLo Is Nothing Then Exit Function

    languageColumn = LanguageIndex()
    If languageColumn <= 0 Then Exit Function

    Set keyRange = KeysDataRange()
    If keyRange Is Nothing Then Exit Function

    Set found = keyRange.Find(What:=token, LookAt:=xlWhole, MatchCase:=True)
    If found Is Nothing Then Exit Function

    Set bodyRange = this.transLo.DataBodyRange
    If bodyRange Is Nothing Then Exit Function

    relativeRow = found.Row - keyRange.Row + 1
    Set ResolveTranslationCell = bodyRange.Cells(relativeRow, languageColumn)
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
                        translated = TranslatedValue(chunk)
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
        multiPageControl.Pages(pageIndex).Caption = TranslatedValue(multiPageControl.Pages(pageIndex).Name)
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

'@label:reset-caches
'@sub-title Invalidate cached column index
'@param resetLanguageIndex Boolean. When True, clears the language column cache. Defaults to True.
Private Sub ResetCaches(Optional ByVal resetLanguageIndex As Boolean = True)
    If resetLanguageIndex Then
        this.languageColumnIndex = 0
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

'@section Interface
'===============================================================================
'ITranslationObject delegation wiring.

Private Property Get ITranslationObject_Name() As String
    ITranslationObject_Name = Name()
End Property

Private Property Get ITranslationObject_Language() As String
    ITranslationObject_Language = Language
End Property

Private Property Get ITranslationObject_LanguagesList() As BetterArray
    Set ITranslationObject_LanguagesList = LanguagesList
End Property

Private Property Get ITranslationObject_ValueExists(ByVal text As String) As Boolean
    ITranslationObject_ValueExists = ValueExists(text)
End Property

Private Function ITranslationObject_TranslatedValue( _
    ByVal Text As String, _
    Optional ByVal containsFormula As Boolean = False) As String
    ITranslationObject_TranslatedValue = TranslatedValue(Text, containsFormula)
End Function

Private Sub ITranslationObject_TranslateRange( _
    ByVal colRng As Range, _
    Optional ByVal containsFormula As Boolean = False)
    TranslateRange colRng:=colRng, containsFormula:=containsFormula
End Sub

Private Sub ITranslationObject_TranslateForm(ByVal UserFrm As Object)
    TranslateForm UserFrm
End Sub

Private Property Get ITranslationObject_HasCheckings() As Boolean
    ITranslationObject_HasCheckings = HasCheckings()
End Property

Private Property Get ITranslationObject_CheckingValues() As Object
    Set ITranslationObject_CheckingValues = CheckingValues()
End Property
