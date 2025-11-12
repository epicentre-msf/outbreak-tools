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
'@Folder("GeneralInterface")
'@ModuleDescription("Interface of translation object")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'TranslationObject is a class for translation of elements such 
'as ranges, values including formulas. Do not confuse with the 
'translation class of the setup or other translation elements. This
'one is only for translating chunks.


'@jump:elements
'@section Elements

'@jump:Name
Public Property Get Name() As String
End Property

'@jump:Language
'@return Current language header associated with the translation table.
Public Property Get Language() As String
End Property

'@jump:LanguagesList
'@return BetterArray containing each language header, excluding the helper/tag column.
Public Property Get LanguagesList() As BetterArray
End Property

'@jump:ValueExists
'@param text String tag to search in the translation table.
'@return True when the tag exists for the configured language.
Public Property Get ValueExists(ByVal text As String) As Boolean
End Property

'@jump:TranslatedValue
'@param Text String to translate.
'@param containsFormula Optional flag; when True only quoted chunks inside the text are translated.
'@return Either the translated value or the original text when no translation was found.
Public Function TranslatedValue(ByVal Text As String, Optional ByVal containsFormula As Boolean = False) As String
End Function

'@jump:outside
'@section Outside

'@jump:TranslateForm
'@param UserFrm Any userform exposing a Controls collection whose captions need translation.
Public Sub TranslateForm(ByVal UserFrm As Object)
End Sub

'@jump:TranslateRange
'@param colRng Column range whose values should be translated.
'@param containsFormula Optional flag to translate the range as formulas instead of plain text.
Public Sub TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
End Sub


'@section checkings

'@jump:HasCheckings
'Indicates whether the translator logged checking notes.
Public Property Get HasCheckings() As Boolean: End Property

'@jump:CheckingValues
'Returns the logged checking values for inspection.
Public Property Get CheckingValues() As Object: End Property
