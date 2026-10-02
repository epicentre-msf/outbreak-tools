VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITranslationObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of translation object"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of translation object")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ITranslationObject
'@description
'Defines the public contract for the TranslationObject class, exposing
'tag-based translation of strings, ranges, and user-form controls from
'a ListObject translation table keyed by language column.


'@section Elements
'===============================================================================
'@description Core translation accessors and lookups.

'@jump:name
'@prop-title Name of the translation object
'@return String. Descriptive identifier including table and language.
Public Property Get Name() As String
End Property

'@jump:language
'@prop-title Current language header associated with the translation table
'@return String. The active language.
Public Property Get Language() As String
End Property

'@jump:languages-list
'@prop-title List of available language headers
'@return BetterArray. Each language header, excluding the helper/tag column.
Public Property Get LanguagesList() As BetterArray
End Property

'@jump:value-exists
'@prop-title Whether a tag exists in the translation table
'@param text String. Tag to search in the translation table.
'@return Boolean. True when the tag exists for the configured language.
Public Property Get ValueExists(ByVal text As String) As Boolean
End Property

'@jump:translated-value
'@sub-title Translate a text tag or formula string
'@param Text String. Text to translate.
'@param containsFormula Boolean. When True, only quoted chunks are translated. Defaults to False.
'@return String. The translated value, or the original text when no translation was found.
Public Function TranslatedValue(ByVal Text As String, Optional ByVal containsFormula As Boolean = False) As String
End Function


'@section Outside
'===============================================================================
'@description Bulk translation of forms and ranges.

'@jump:translate-form
'@sub-title Translate all supported controls on a user form
'@param UserFrm Object. Any user form exposing a Controls collection.
Public Sub TranslateForm(ByVal UserFrm As Object)
End Sub

'@jump:translate-range
'@sub-title Translate all values in a column range
'@param colRng Range. Column range whose values should be translated.
'@param containsFormula Boolean. When True, translates as formulas. Defaults to False.
Public Sub TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
End Sub


'@section Checkings
'===============================================================================
'@description Diagnostic logging access.

'@jump:has-checkings
'@prop-title Whether the translator logged diagnostic notes
'@return Boolean. True when at least one entry exists.
Public Property Get HasCheckings() As Boolean: End Property

'@jump:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Object. The checking instance, or Nothing.
Public Property Get CheckingValues() As Object: End Property
