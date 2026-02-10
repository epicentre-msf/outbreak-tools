VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDesignerPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract preparing designer workbooks and ribbon settings"

'@Interface
'@Folder("Designer")
'@ModuleDescription("Prepares designer workbooks for ribbon interactions and exposes persisted settings.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class IDesignerPreparation
'@description
'Defines the public contract for the DesignerPreparation class, exposing
'methods to bind to a designer workbook, run preparation tasks, and
'persist or read ribbon checkbox flags (alert/instruct) through
'workbook-level hidden names.


'@section Lifecycle
'===============================================================================

'@jump:configure
'@sub-title Bind the preparation helper to a designer workbook
'@details
'Sets the internal host workbook reference and resets any previously
'cached hidden names store. Must be called before any other operation.
'@param hostBook Workbook. The designer workbook to prepare.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
End Sub

'@jump:prepare
'@sub-title Execute preparation tasks using the supplied development manager
'@details
'Ensures the hidden names store and default ribbon flags exist on the
'host workbook. The devManager parameter is accepted for interface parity
'with setup preparation but is not currently used for designer-specific
'work.
'@param devManager IDevelopment. ByRef development manager instance, or Nothing.
'@export
Public Sub Prepare(ByRef devManager As IDevelopment)
End Sub


'@section Settings
'===============================================================================

'@jump:set-flag
'@sub-title Persist a ribbon checkbox flag as a hidden name
'@details
'Stores a boolean value as "Yes" or "No" in the workbook-level hidden
'names store, identified by the supplied name. Ensures the store and
'default flags exist before writing.
'@param nameId String. The hidden name identifier for the flag.
'@param enabled Boolean. True to store "Yes", False to store "No".
'@export
Public Sub SetFlag(ByVal nameId As String, ByVal enabled As Boolean)
End Sub

'@jump:get-flag
'@sub-title Read a ribbon checkbox flag from hidden names
'@details
'Reads the named value from the hidden names store and returns True
'when the stored value equals "Yes" (case-insensitive). Falls back to
'the defaultValue when the name does not exist.
'@param nameId String. The hidden name identifier for the flag.
'@param defaultValue Optional Boolean. Value returned when the flag is missing. Defaults to False.
'@return Boolean. True when the stored value is "Yes".
'@export
Public Function GetFlag(ByVal nameId As String, Optional ByVal defaultValue As Boolean = False) As Boolean
End Function


'@section State
'===============================================================================

'@jump:host-workbook
'@prop-title The configured host workbook
'@details
'Returns the designer workbook that was bound via Configure. Raises an
'error when accessed before configuration.
'@return Workbook. The host workbook.
'@export
Public Property Get HostWorkbook() As Workbook
End Property

'@jump:hidden-store
'@prop-title The hidden names store used for designer settings
'@details
'Returns the workbook-level IHiddenNames instance used to persist ribbon
'flags. The store is lazily initialised on first access from the host
'workbook.
'@return IHiddenNames. The hidden names manager.
'@export
Public Property Get HiddenStore() As IHiddenNames
End Property
