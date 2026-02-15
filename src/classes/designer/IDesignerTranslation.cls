VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDesignerTranslation"
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

'@class IDesignerTranslation
'@description
'Defines the public contract for the DesignerTranslation class, exposing
'methods to retrieve translation objects for different content types
'(messages, shapes, ranges, dropdowns), translate an entire designer
'worksheet with a language string, and look up individual translated
'message values. The language code is persisted as a worksheet-level
'HiddenName for cross-session persistence.



Public Enum DesignerTranslationScope
    DesignerTranslationOfMessages = 0
    DesignerTranslationOfShapes = 2
    DesignerTranslationOfRanges = 5
    DesignerTranslationOfDropdowns = 6
End Enum

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
Public Property Get TransObject(Optional ByVal trans As Byte = DesignerTranslationOfMessages) As ITranslationObject: End Property


'@section Methods
'===============================================================================

'@jump:translate-designer
'@sub-title Translate an entire designer worksheet
'@details
'Applies translations to all shapes, named ranges, and dropdown
'validations on the supplied worksheet. When a non-empty language string
'is provided (e.g. "FRA - Francais"), extracts the language code and
'persists it to the worksheet-level HiddenName. When language is empty,
'reads the stored language code from the HiddenName. Exits silently
'when no language code is available.
'@param mainsh Worksheet. The designer worksheet to translate.
'@param language String. Language string in "CODE - Label" format (e.g. "FRA - Francais"). Pass vbNullString to use the stored language.
'@export
Public Sub TranslateDesigner(ByVal mainsh As Worksheet, ByVal language As String): End Sub

'@jump:translation-msg
'@sub-title Get the translated value of a message code
'@details
'Looks up a message code in the messages translation table and returns
'the translated string for the current designer language.
'@param msgCode String. The message code to translate.
'@return String. The translated message text.
'@export
Public Function TranslatedValue(ByVal msgCode As String) As String: End Function
