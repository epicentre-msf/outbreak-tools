VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignerTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Translation of the designer"


'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation
'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("Translation of the designer")

Option Explicit

Implements IDesignerTranslation

'@class DesignerTranslation
'@description
'Manages translation of the designer workbook by reading translation
'tables (shapes, ranges, messages, dropdowns) from a dedicated
'translation worksheet. Provides factory creation with requirement
'validation, returns cached ITranslationObject instances scoped to
'each content type, and applies translations to an entire designer
'worksheet by updating shape text, named range values, and dropdown
'validations. The current language code is extracted from a language
'string (e.g. "FRA - Francais") and persisted as a worksheet-level
'HiddenName for cross-session persistence. All translation objects
'and name lists are cached for fast repeated access.
'Consumers interact through the IDesignerTranslation interface.
'@depends TranslationObject, ITranslationObject, BetterArray, HiddenNames, IHiddenNames



'@section State
'===============================================================================

Private Type TDesignerTranslation
    sh As Worksheet
    langCode As String
    msgTrans As ITranslationObject
    shapeTrans As ITranslationObject
    rangeTrans As ITranslationObject
    dropTrans As ITranslationObject
    shapeNames As BetterArray
    rangeNames As BetterArray
    dropNames As BetterArray
    wkshNames As IHiddenNames
    isCached As Boolean
End Type

Private this As TDesignerTranslation

'@section Constants
'===============================================================================

Private Const TABSHAPE As String = "T_tradShape"
Private Const TABRANGE As String = "T_tradRange"
Private Const TABSMSG As String = "T_tradMsg"
Private Const TABDROP As String = "T_tradDrop"
Private Const CLASS_NAME As String = "DesignerTranslation"
Private Const LANG_HIDDEN_NAME As String = "TAG_DES_LANG"

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new DesignerTranslation instance from a translation worksheet
'@details
'Factory method that validates all required ListObjects on the supplied
'worksheet, then constructs and returns a new instance through the
'IDesignerTranslation interface. Called on the predeclared instance.
'@param sh Worksheet. The translation worksheet containing T_tradShape, T_tradRange, T_tradMsg, T_tradDrop.
'@return IDesignerTranslation. A fully initialised instance ready for use.
'@export
Public Function Create(ByVal sh As Worksheet) As IDesignerTranslation
    CheckRequirements sh
    Dim instance As DesignerTranslation
    Set instance = New DesignerTranslation
    Set instance.Wksh = sh
    Set Create = instance
End Function

'@label:wksh
'@prop-title Translation worksheet
'@return Worksheet. The backing translation worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:wksh-set
'@prop-title Assign the translation worksheet
'@param sh Worksheet. The worksheet to store.
Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

'@section Translation Objects
'===============================================================================

'@label:designer-language
'@prop-title Current designer language code
'@details
'Returns the cached language code if available, otherwise reads from the
'worksheet-level HiddenName. Returns an empty string when no language
'code is stored.
'@return String. The current language code, or vbNullString when not set.
Private Property Get DesignerLanguage() As String
    If LenB(this.langCode) > 0 Then
        DesignerLanguage = this.langCode
        Exit Property
    End If

    Dim names As IHiddenNames
    Set names = WkshNames()

    If names.HasName(LANG_HIDDEN_NAME) Then
        this.langCode = names.ValueAsString(LANG_HIDDEN_NAME)
    End If

    DesignerLanguage = this.langCode
End Property

'@label:trans-object
'@prop-title Translation object for a given content type
'@details
'Returns a cached ITranslationObject for the requested translation
'scope. Ensures caches are built using the current language code.
'Returns Nothing when no language code is configured.
'@param trans Optional Byte. TranslationType value selecting the content scope. Defaults to TranslationOfMessages.
'@return ITranslationObject. The translation object for the requested scope.
'@export
Public Property Get TransObject(Optional ByVal trans As Byte = DesignerTranslationOfMessages) As ITranslationObject
    Dim langCode As String
    langCode = DesignerLanguage()

    If LenB(langCode) = 0 Then Exit Property

    EnsureCaches langCode

    Select Case trans
    Case DesignerTranslationOfMessages
        Set TransObject = this.msgTrans
    Case DesignerTranslationOfShapes
        Set TransObject = this.shapeTrans
    Case DesignerTranslationOfRanges
        Set TransObject = this.rangeTrans
    Case DesignerTranslationOfDropdowns
        Set TransObject = this.dropTrans
    End Select
End Property

'@section Translation Application
'===============================================================================

'@label:translate-designer
'@sub-title Translate an entire designer worksheet
'@details
'Applies translations to all shapes, named ranges, and dropdown
'validations on the supplied worksheet. When a non-empty language string
'is provided (e.g. "FRA - Francais"), extracts the language code and
'persists it to the worksheet-level HiddenName. When the language string
'is empty, reads the stored language code. Exits silently when no
'language code is available.
'@param mainsh Worksheet. The designer worksheet to translate.
'@param language String. Language string in "CODE - Label" format. Pass vbNullString to use the stored language.
'@export
Private Sub TranslateDesigner(ByVal mainsh As Worksheet, ByVal language As String)
    Dim langCode As String

    If LenB(language) > 0 Then
        langCode = ExtractLangCode(language)
        UpdateLanguageCode langCode
    Else
        langCode = DesignerLanguage()
    End If

    If LenB(langCode) = 0 Then Exit Sub

    EnsureCaches langCode

    TranslateShapes mainsh
    TranslateRanges mainsh
    TranslateDropdowns mainsh
End Sub

'@label:translate-shapes
'@sub-title Apply shape translations on the target worksheet
'@details
'Iterates over all shapes in the worksheet and replaces text content
'for shapes whose names appear in the cached shape names list. Preserves
'the original font name on each shape.
'@param mainsh Worksheet. The worksheet containing shapes to translate.
Private Sub TranslateShapes(ByVal mainsh As Worksheet)
    Dim shp As Shape
    Dim prevFont As String

    If this.shapeNames Is Nothing Then Exit Sub

    For Each shp In mainsh.Shapes
        If this.shapeNames.Includes(shp.Name) Then
            With shp
                prevFont = .TextFrame.Characters.Font.Name
                .TextFrame.Characters.Text = this.shapeTrans.TranslatedValue(shp.Name)
                .TextFrame.Characters.Font.Name = prevFont
            End With
        End If
    Next
End Sub

'@label:translate-ranges
'@sub-title Apply named range translations on the target worksheet
'@details
'Iterates over the cached range names and updates each matching named
'range on the worksheet with its translated value.
'@param mainsh Worksheet. The worksheet containing named ranges to translate.
Private Sub TranslateRanges(ByVal mainsh As Worksheet)
    Dim counter As Long
    Dim rngCode As String
    Dim mainRng As Range

    If this.rangeNames Is Nothing Then Exit Sub

    For counter = this.rangeNames.LowerBound To this.rangeNames.UpperBound
        rngCode = this.rangeNames.Item(counter)
        On Error Resume Next
        Set mainRng = mainsh.Range(rngCode)
        On Error GoTo 0

        If Not (mainRng Is Nothing) Then
            mainRng.Value = this.rangeTrans.TranslatedValue(rngCode)
            Set mainRng = Nothing
        End If
    Next
End Sub

'@label:translate-dropdowns
'@sub-title Apply dropdown validation translations on the target worksheet
'@details
'Iterates over the cached dropdown names, deletes existing validation on
'each matching range, and creates a new list validation with the
'translated formula. Clears any existing cell content after applying
'the validation.
'@param mainsh Worksheet. The worksheet containing dropdowns to translate.
Private Sub TranslateDropdowns(ByVal mainsh As Worksheet)
    Dim counter As Long
    Dim rngCode As String
    Dim mainRng As Range

    If this.dropNames Is Nothing Then Exit Sub

    For counter = this.dropNames.LowerBound To this.dropNames.UpperBound
        rngCode = this.dropNames.Item(counter)
        On Error Resume Next
        Set mainRng = mainsh.Range(rngCode)

        If Not (mainRng Is Nothing) Then
            With mainRng.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="= " & this.dropTrans.TranslatedValue(rngCode)

                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = vbNullString
                .ErrorTitle = vbNullString
                .InputMessage = vbNullString
                .ErrorMessage = vbNullString
                .ShowInput = True
                .ShowError = True
            End With

            mainRng.ClearContents
            Set mainRng = Nothing
        End If
        On Error GoTo 0
    Next
End Sub

'@section Message Lookup
'===============================================================================

'@label:translation-msg
'@sub-title Get the translated value of a message code
'@details
'Returns the translated text for the supplied message code using the
'cached messages translation object. Ensures caches are built before
'lookup. Returns vbNullString when no language code is configured.
'@param msgCode String. The message code to translate.
'@return String. The translated message text.
'@export
Private Function TranslatedValue(ByVal msgCode As String) As String
    Dim langCode As String
    langCode = DesignerLanguage()

    If LenB(langCode) = 0 Then Exit Function

    EnsureCaches langCode
    TranslatedValue = this.msgTrans.TranslatedValue(msgCode)
End Function

'@section Language Code Management
'===============================================================================

'@label:extract-lang-code
'@sub-title Extract the language code from a language string
'@details
'Parses a language string in "CODE - Label" format (e.g. "FRA - Francais")
'and returns the code portion. If no separator is found, returns the
'trimmed input as-is.
'@param language String. The language string to parse.
'@return String. The extracted language code.
Private Function ExtractLangCode(ByVal language As String) As String
    Dim pos As Long
    pos = InStr(1, language, " - ", vbBinaryCompare)
    If pos > 0 Then
        ExtractLangCode = Trim$(Left$(language, pos - 1))
    Else
        ExtractLangCode = Trim$(language)
    End If
End Function

'@label:update-language-code
'@sub-title Persist the language code to the worksheet-level HiddenName
'@details
'Writes the language code to the TAG_DES_LANG HiddenName on the
'translation worksheet. Creates the name if it does not exist yet.
'Does not update the in-memory langCode field; EnsureCaches owns that
'so the cache-invalidation check remains correct.
'@param langCode String. The language code to persist.
Private Sub UpdateLanguageCode(ByVal langCode As String)
    Dim names As IHiddenNames
    Set names = WkshNames()
    names.EnsureName LANG_HIDDEN_NAME, langCode, HiddenNameTypeString
    names.SetValue LANG_HIDDEN_NAME, langCode
End Sub

'@label:wksh-names
'@prop-title Cached HiddenNames accessor for the translation worksheet
'@details
'Lazily creates and returns an IHiddenNames instance scoped to the
'translation worksheet.
'@return IHiddenNames. The HiddenNames manager for the translation worksheet.
Private Property Get WkshNames() As IHiddenNames
    If this.wkshNames Is Nothing Then
        Set this.wkshNames = HiddenNames.Create(this.sh)
    End If
    Set WkshNames = this.wkshNames
End Property

'@section Cache Management
'===============================================================================

'@label:ensure-caches
'@sub-title Build or refresh translation caches for the given language code
'@details
'When the caches are not yet built or the language code has changed,
'creates four TranslationObject instances and reads three name lists
'from the translation ListObjects. Skips rebuilding when the caches
'are already populated for the same language code.
'@param langCode String. The language code to cache translations for.
Private Sub EnsureCaches(ByVal langCode As String)
    Dim sh As Worksheet

    If this.isCached And StrComp(this.langCode, langCode, vbTextCompare) = 0 Then Exit Sub

    Set sh = Wksh()
    this.langCode = langCode

    Set this.msgTrans = TranslationObject.Create(sh.ListObjects(TABSMSG), langCode)
    Set this.shapeTrans = TranslationObject.Create(sh.ListObjects(TABSHAPE), langCode)
    Set this.rangeTrans = TranslationObject.Create(sh.ListObjects(TABRANGE), langCode)
    Set this.dropTrans = TranslationObject.Create(sh.ListObjects(TABDROP), langCode)

    Set this.shapeNames = New BetterArray
    this.shapeNames.FromExcelRange sh.ListObjects(TABSHAPE).ListColumns(1).DataBodyRange, _
                                   DetectLastRow:=False, DetectLastColumn:=False

    Set this.rangeNames = New BetterArray
    this.rangeNames.FromExcelRange sh.ListObjects(TABRANGE).ListColumns(1).DataBodyRange, _
                                   DetectLastRow:=False, DetectLastColumn:=False

    Set this.dropNames = New BetterArray
    this.dropNames.FromExcelRange sh.ListObjects(TABDROP).ListColumns(1).DataBodyRange, _
                                  DetectLastRow:=False, DetectLastColumn:=False

    this.isCached = True
End Sub

'@section Validation
'===============================================================================

'@label:check-requirements
'@sub-title Validate that the worksheet has all required ListObjects
'@details
'Checks that the supplied worksheet is not Nothing, then verifies the
'presence of four required ListObjects (T_tradShape, T_tradRange,
'T_tradMsg, T_tradDrop). Raises an error when any requirement is
'missing.
'@param sh Worksheet. The worksheet to validate.
Private Sub CheckRequirements(ByVal sh As Worksheet)
    Dim tabl As BetterArray
    If sh Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, _
        "Unable to instanciate, no linelist translation sheet"
    Set tabl = New BetterArray
    tabl.Push TABSHAPE, TABRANGE, TABSMSG, TABDROP
    LoExists sh, tabl
End Sub

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Raises a VBA error using the supplied ProjectError code and message.
'@param errNumber Long. The ProjectError error code.
'@param message String. Human-readable description of the failure.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise errNumber, CLASS_NAME, "Error : " & message
End Sub

'@label:lo-exists
'@sub-title Verify that all listed ListObjects exist on a worksheet
'@details
'Iterates over the supplied BetterArray of table names and checks that
'each ListObject can be found on the given worksheet. Raises an error
'for the first missing table encountered.
'@param sh Worksheet. The worksheet to check.
'@param loNames BetterArray. List of ListObject names to verify.
Private Sub LoExists(sh As Worksheet, ByVal loNames As BetterArray)
    Dim Lo As ListObject
    Dim counter As Long
    Dim lName As String
    For counter = loNames.LowerBound To loNames.UpperBound
        lName = loNames.Item(counter)
        On Error Resume Next
        Set Lo = sh.ListObjects(lName)
        On Error GoTo 0
        If Lo Is Nothing Then ThrowError ProjectError.ElementNotFound, _
            "Table " & lName & " does not exists in Worksheet " & sh.Name
    Next
End Sub

'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the IDesignerTranslation contract.
'See the corresponding members above for full documentation.

Private Property Get IDesignerTranslation_TransObject(Optional ByVal trans As Byte = DesignerTranslationOfMessages) As ITranslationObject
    Set IDesignerTranslation_TransObject = TransObject(trans)
End Property

Private Sub IDesignerTranslation_TranslateDesigner(ByVal mainsh As Worksheet, ByVal language As String)
    TranslateDesigner mainsh, language
End Sub

Private Function IDesignerTranslation_TranslatedValue(ByVal msgCode As String) As String
    IDesignerTranslation_TranslatedValue = TranslatedValue(msgCode)
End Function
