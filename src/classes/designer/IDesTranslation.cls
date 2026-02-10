VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDesTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for designer translation management"

'@Interface
'@Folder("Designer")
'@ModuleDescription("Interface for designer translation management")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IDesTranslation
'@description
'Defines the public contract for the DesTranslation class, exposing
'methods to retrieve translation objects for different content types
'(messages, shapes, ranges, dropdowns), translate an entire designer
'worksheet, resolve the language list range, and look up individual
'translated message values.


'@section Properties
'===============================================================================

'@jump:trans-object
'@prop-title Translation object for a given content type
'@details
'Returns an ITranslationObject configured for the requested translation
'scope (messages, shapes, ranges, or dropdowns). The returned object
'reads from the corresponding ListObject on the translation sheet using
'the current designer language.
'@param trans Optional Byte. TranslationType value selecting the content scope. Defaults to TranslationOfMessages.
'@return ITranslationObject. The translation object for the requested scope.
'@export
Public Property Get TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject: End Property

'@jump:lang-list-rng
'@prop-title Language list range on the translation sheet
'@details
'Returns the first cell of the LangDictList named range on the
'translation sheet. Used by consumers to read or write the current
'language code.
'@return Range. The language list range cell.
'@export
Public Property Get LangListRng() As Range: End Property


'@section Methods
'===============================================================================

'@jump:translate-designer
'@sub-title Translate an entire designer worksheet
'@details
'Applies translations to all shapes, named ranges, and dropdown
'validations on the supplied worksheet. Reads shape names, range codes,
'and dropdown codes from the translation ListObjects and replaces their
'content with translated values in the current designer language. Exits
'silently when no language code is configured.
'@param mainsh Worksheet. The designer worksheet to translate.
'@export
Public Sub TranslateDesigner(ByVal mainsh As Worksheet): End Sub

'@jump:translation-msg
'@sub-title Get the translated value of a message code
'@details
'Looks up a message code in the messages translation table and returns
'the translated string for the current designer language.
'@param msgCode String. The message code to translate.
'@return String. The translated message text.
'@export
Public Function TranslationMsg(ByVal msgCode As String) As String: End Function
