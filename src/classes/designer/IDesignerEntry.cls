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
'and translate the worksheet to a given language.


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
