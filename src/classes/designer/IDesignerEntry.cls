VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDesignerEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for managing designer entry sheet interactions"

'@Interface
'@Folder("Designer")
'@ModuleDescription("Coordinates clearing and translation actions for designer entry sheets.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class IDesignerEntry
'@description
'Defines the public contract for the DesignerEntry class, exposing
'methods to bind to a designer entry worksheet, inject test doubles for
'the entry manager and translation provider, clear user input ranges,
'translate the worksheet to a given language, read and write info
'fields on the Main sheet, and translate message codes.


'@section Lifecycle
'===============================================================================

'@jump:configure
'@sub-title Bind the entry helper to the worksheet hosting designer inputs
'@details
'Sets the internal host worksheet reference. Must be called before any
'other operation. Raises an error when the worksheet is Nothing.
'@param hostSheet Worksheet. The worksheet hosting designer inputs.
'@export
Public Sub Configure(ByVal hostSheet As Worksheet)
End Sub

'@jump:use-entry-manager
'@sub-title Swap in a custom entry manager implementation
'@details
'Replaces the internally resolved IMain entry manager with the supplied
'instance. Primarily used in tests to inject a stub or mock.
'@param entryManager IMain. The entry manager to use.
'@export
Public Sub UseEntryManager(ByVal entryManager As IMain)
End Sub

'@jump:use-translator
'@sub-title Swap in a custom translations provider
'@details
'Replaces the internally resolved IDesignerTranslation translator with the
'supplied instance. Primarily used in tests to inject a stub or mock.
'@param translator IDesignerTranslation. The translations provider to use.
'@export
Public Sub UseTranslator(ByVal translator As IDesignerTranslation)
End Sub


'@section Operations
'===============================================================================

'@jump:clear
'@sub-title Clear user entry ranges on the configured worksheet
'@details
'Delegates to the entry manager to clear all input ranges on the
'configured worksheet, removing user-entered values.
'@export
Public Sub Clear()
End Sub

'@jump:translate
'@sub-title Translate the configured worksheet to the provided language
'@details
'Delegates to the translations provider to apply translated labels,
'shapes, and dropdowns to the configured worksheet. When a non-empty
'language string is supplied (e.g. "FRA - Francais"), the translator
'extracts the language code and persists it.
'@param language Optional String. Language string in "CODE - Label" format. Defaults to vbNullString.
'@export
Public Sub Translate(Optional ByVal language As String = vbNullString)
End Sub

'@jump:add-info
'@sub-title Write a value to a named info range on the host worksheet
'@details
'Resolves the named range matching infoName on the host worksheet,
'writes the supplied value, and sets the cell background to white.
'Silently exits when the named range does not exist.
'Valid infoName values: "edition", "setuppath", "geopath", "lldir",
'"llname", "setuplang", "temppath".
'@param value String. The value to write (path, message text, etc.).
'@param infoName String. The logical name identifying the target range.
'@export
Public Sub AddInfo(ByVal value As String, ByVal infoName As String)
End Sub

'@jump:value-of
'@prop-title Read a value from a named info range on the host worksheet
'@details
'Resolves the named range matching infoName on the host worksheet and
'returns its current value. Returns vbNullString when the range does
'not exist.
'@param infoName String. The logical name identifying the source range.
'@return String. The current value of the named range.
'@export
Public Property Get ValueOf(ByVal infoName As String) As String
End Property

'@jump:translate-message
'@sub-title Translate a message code using the designer translations
'@details
'Delegates to the internal translations provider to translate a message
'code (e.g. "MSG_ChemFich") into the current designer language. Returns
'the raw msgCode when no translation is available.
'@param msgCode String. The message code to translate.
'@return String. The translated message text.
'@export
Public Function TranslateMessage(ByVal msgCode As String) As String
End Function


'@section State
'===============================================================================

'@jump:host-sheet
'@prop-title The worksheet managed by this helper
'@details
'Returns the designer entry worksheet that was bound via Configure.
'Raises an error when accessed before configuration.
'@return Worksheet. The host worksheet.
'@export
Public Property Get HostSheet() As Worksheet
End Property
