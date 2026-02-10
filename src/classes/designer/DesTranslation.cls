VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesTranslation"
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

Implements IDesTranslation

'@class DesTranslation
'@description
'Manages translation of the designer workbook by reading translation
'tables (shapes, ranges, messages, dropdowns, languages) from a
'dedicated translation worksheet. Provides factory creation with
'requirement validation, returns ITranslationObject instances scoped to
'each content type, and applies translations to an entire designer
'worksheet by updating shape text, named range values, and dropdown
'validations. Consumers interact through the IDesTranslation interface.
'@depends TranslationObject, ITranslationObject, BetterArray


'@section State
'===============================================================================

Private Type TDesTranslation
    sh As Worksheet
End Type

Private this As TDesTranslation

'@section Constants
'===============================================================================

Private Const TABSHAPE As String = "T_tradShape"
Private Const TABRANGE As String = "T_tradRange"
Private Const TABSMSG As String = "T_tradMsg"
Private Const TABLANG As String = "T_Lang"
Private Const TABDROP As String = "T_tradDrop"
Private Const TABLANGDICT As String = "T_LanguageDictionary"
Private Const CLASSNAME As String = "DesTranslation"
Private Const RNGMAINLANG As String = "RNG_MainLang"
Private Const RNGMAINLANGCODE As String = "RNG_MainLangCode"
Private Const RNGLANGDICTLIST As String = "LangDictList"

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new DesTranslation instance from a translation worksheet
'@details
'Factory method that validates all required ListObjects and named ranges
'on the supplied worksheet, then constructs and returns a new instance
'through the IDesTranslation interface. Called on the predeclared
'instance.
'@param sh Worksheet. The translation worksheet containing T_tradShape, T_tradRange, T_tradMsg, T_Lang, T_LanguageDictionary, and required named ranges.
'@return IDesTranslation. A fully initialised instance ready for use.
'@export
Public Function Create(ByVal sh As Worksheet) As IDesTranslation

    CheckRequirements sh

    With New DesTranslation
        Set .Wksh = sh
        Set Create = .Self()
    End With

End Function

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor so consuming code can fluently retrieve the
'interface reference from the predeclared Create method.
'@return IDesTranslation. The current instance cast to the interface.
Public Property Get Self() As IDesTranslation
    Set Self = Me
End Property

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
'Reads the language code from the RNG_MainLangCode named range on the
'translation worksheet. Returns an empty string when the value is not a
'string type, indicating no language is configured.
'@return String. The current language code, or vbNullString when not set.
Private Property Get DesignerLanguage() As String
    Dim sh As Worksheet
    Dim lang As Variant
    Set sh = Wksh()

    lang = sh.Range(RNGMAINLANGCODE).Value

    If (TypeName(lang) = "String") Then
        DesignerLanguage = lang
    Else
        DesignerLanguage = vbNullString
    End If
End Property

'@label:trans-object
'@prop-title Translation object for a given content type
'@details
'Returns an ITranslationObject configured for the requested translation
'scope. Uses the TranslationType enum values: TranslationOfMessages
'selects T_tradMsg, TranslationOfShapes selects T_tradShape,
'TranslationOfRanges selects T_tradRange, and TranslationOfDropdowns
'selects T_tradDrop. The returned object is bound to the current
'designer language code.
'@param trans Optional Byte. TranslationType value selecting the content scope. Defaults to TranslationOfMessages.
'@return ITranslationObject. The translation object for the requested scope.
'@export
Public Property Get TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject

    Dim Lo As ListObject
    Dim lang As String
    Dim sh As Worksheet

    Set sh = Wksh()
    lang = DesignerLanguage()

    Select Case trans
    Case TranslationOfMessages
        Set Lo = sh.ListObjects(TABSMSG)
    Case TranslationOfShapes
        Set Lo = sh.ListObjects(TABSHAPE)
    Case TranslationOfRanges
        Set Lo = sh.ListObjects(TABRANGE)
    Case TranslationOfDropdowns
        Set Lo = sh.ListObjects(TABDROP)
    End Select

    Set TransObject = TranslationObject.Create(Lo, lang)
End Property

'@section Translation Application
'===============================================================================

'@label:translate-designer
'@sub-title Translate an entire designer worksheet
'@details
'Applies translations to all shapes, named ranges, and dropdown
'validations on the supplied worksheet. Reads shape names from
'T_tradShape and replaces matching shape text while preserving font.
'Reads range codes from T_tradRange and updates matching named range
'values. Reads dropdown codes from T_tradDrop and replaces data
'validation lists with translated formulas, clearing existing content.
'Exits silently when no language code is configured.
'@param mainsh Worksheet. The designer worksheet to translate.
'@export
Private Sub TranslateDesigner(ByVal mainsh As Worksheet)

    Dim tradobj As ITranslationObject
    Dim shp As Shape
    Dim counter As Long
    Dim shapesNames As BetterArray               'Name of all shapes to translate in worksheet main
    Dim rngNames As BetterArray                  'List of all named ranges in the worksheet main
    Dim Lo As ListObject
    Dim sh As Worksheet
    Dim langCode As String                       'Language for the designer
    Dim prevFont As String
    Dim rngCode As String                        'Code Name of the range
    Dim mainRng As Range

    Set sh = Wksh()

    langCode = DesignerLanguage()

    If (langCode = vbNullString) Then Exit Sub

     'Translation of shapes
     Set shapesNames = New BetterArray
     Set Lo = sh.ListObjects(TABSHAPE)
     Set tradobj = TransObject(TranslationOfShapes)
     shapesNames.FromExcelRange Lo.ListColumns(1).DataBodyRange, DetectLastRow:=False, DetectLastColumn:=False

    For Each shp In mainsh.Shapes
         If shapesNames.Includes(shp.Name) Then
             With shp
                 'keep the previous font
                 prevFont = .TextFrame.Characters.Font.Name
                 .TextFrame.Characters.Text = tradobj.TranslatedValue(shp.Name)
                 .TextFrame.Characters.Font.Name = prevFont
             End With
         End If
    Next

   'Translation of label Ranges
     Set rngNames = New BetterArray
     Set Lo = sh.ListObjects(TABRANGE)
     Set tradobj = TransObject(TranslationOfRanges)
     rngNames.FromExcelRange Lo.ListColumns(1).DataBodyRange, DetectLastRow:=False, DetectLastColumn:=False

    For counter = rngNames.LowerBound To rngNames.UpperBound
         rngCode = rngNames.Item(counter)
         'Test if the range exists before changing its value
         On Error Resume Next
         Set mainRng = mainsh.Range(rngCode)
         On Error GoTo 0

       If Not (mainRng Is Nothing) Then
             mainRng.Value = tradobj.TranslatedValue(rngCode)
             Set mainRng = Nothing
         End If
    Next

    'Translation of dropdowns
    Set rngNames = New BetterArray
     Set Lo = sh.ListObjects(TABDROP)
     Set tradobj = TransObject(TranslationOfDropdowns)
     rngNames.FromExcelRange Lo.ListColumns(1).DataBodyRange, DetectLastRow:=False, DetectLastColumn:=False

    For counter = rngNames.LowerBound To rngNames.UpperBound
        rngCode = rngNames.Item(counter)
        'Test if the range exists before changing its value
        On Error Resume Next
        Set mainRng = mainsh.Range(rngCode)


        If Not (mainRng Is Nothing) Then

            With mainRng.validation
                .Delete
                'for all the others, add an information alert
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="= " & tradobj.TranslatedValue(rngCode)

                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = vbNullString
                .errorTitle = vbNullString
                .InputMessage = vbNullString
                .errorMessage = vbNullString
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
'Retrieves the TransObject for messages and looks up the supplied
'message code to return its translated text in the current designer
'language.
'@param msgCode String. The message code to translate.
'@return String. The translated message text.
'@export
Private Function TranslationMsg(ByVal msgCode As String) As String
    Dim msgObject As ITranslationObject
    Set msgObject = TransObject(TranslationOfMessages)
    TranslationMsg = msgObject.TranslatedValue(msgCode)
End Function

'@label:lang-list-rng
'@prop-title Language list range on the translation sheet
'@details
'Returns the first cell of the LangDictList named range on the
'translation worksheet. Used by consumers to read or write the current
'language selection.
'@return Range. The language list range cell.
'@export
Private Property Get LangListRng() As Range
    Dim sh As Worksheet
    Set sh = Wksh()
    Set LangListRng = sh.Range(RNGLANGDICTLIST).Cells(1, 1)
End Property

'@section Validation
'===============================================================================

'@label:check-requirements
'@sub-title Validate that the worksheet has all required tables and named ranges
'@details
'Checks that the supplied worksheet is not Nothing, then verifies the
'presence of five required ListObjects (T_tradShape, T_tradRange,
'T_tradMsg, T_Lang, T_LanguageDictionary) and three required named
'ranges (RNG_MainLang, RNG_MainLangCode, LangDictList). Raises an
'error when any requirement is missing.
'@param sh Worksheet. The worksheet to validate.
Private Sub CheckRequirements(ByVal sh As Worksheet)
    Dim tabl As BetterArray
    If sh Is Nothing Then ThrowError "Unable to instanciate, no linelist translation sheet"
    'Required listObjects fo translations
    Set tabl = New BetterArray
    tabl.Push TABSHAPE, TABRANGE, TABSMSG, TABLANG, TABLANGDICT
    LoExists sh, tabl

    RangeExists sh, RNGMAINLANG
    RangeExists sh, RNGMAINLANGCODE
    RangeExists sh, RNGLANGDICTLIST
End Sub

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a VBA error with a descriptive message
'@details
'Raises a custom VBA error using vbObjectError + 516 as the error
'number. Used by validation helpers to signal missing requirements.
'@param errorMessage String. Human-readable description of the failure.
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
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
        If Lo Is Nothing Then ThrowError "Table " & lName & " does not exists in Worksheet " & sh.Name
    Next
End Sub

'@label:range-exists
'@sub-title Verify that a named range exists on a worksheet
'@details
'Attempts to resolve a named range on the given worksheet. Raises an
'error when the range cannot be found.
'@param sh Worksheet. The worksheet to check.
'@param rngName String. The named range to look for.
Private Sub RangeExists(ByVal sh As Worksheet, ByVal rngName As String)

    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then ThrowError "Range " & rngName & " does not exists in worksheet " & sh.Name

End Sub

'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the IDesTranslation contract.
'See the corresponding members above for full documentation.

Private Property Get IDesTranslation_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject
    Set IDesTranslation_TransObject = TransObject(trans)
End Property

Private Sub IDesTranslation_TranslateDesigner(ByVal mainsh As Worksheet)
    TranslateDesigner mainsh
End Sub

Private Property Get IDesTranslation_LangListRng() As Range
    Set IDesTranslation_LangListRng = LangListRng()
End Property

Private Function IDesTranslation_TranslationMsg(ByVal msgCode As String) As String
    IDesTranslation_TranslationMsg = TranslationMsg(msgCode)
End Function
