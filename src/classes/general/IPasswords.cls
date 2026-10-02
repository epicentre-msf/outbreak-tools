VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Passwords"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of Passwords")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IPasswords
'@description
'Defines the public contract for the Passwords class, exposing
'workbook/worksheet protection management, key generation,
'debug-mode toggling, and import/export of password metadata.
'Uses TypeName dispatch to handle Workbook, Worksheet, and String
'target arguments through a single Protect/UnProtect entry point.


'@section Enumerations
'===============================================================================
'@description Tag identifiers for password named ranges.

Public Enum PaswordTag
  PasswordTagDebug = 1
  PasswordTagPublic = 2
  PasswordTagPrivate = 3
  PasswordTagVersion = 4
  PasswordTagMode = 5
  PasswordTagLabPublic = 6
  PasswordTagLabPrivate = 7
End Enum


'@section Core Accessors
'===============================================================================
'@description Properties exposing the password worksheet and stored values.

'@jump:password-sheet
'@prop-title Worksheet hosting the password tables
'@return Worksheet. The password worksheet.
Public Property Get PasswordSheet() As Worksheet: End Property

'@jump:value
'@prop-title Retrieve the value for a named tag
'@details
'Looks up the named range corresponding to the supplied tag label
'and returns its current value.
'@param tagName String. Supported tags: debuggingpassword, publickey, privatekey, version, labpublickey, labprivatekey.
'@return String. The stored value.
Public Property Get Value(ByVal tagName As String) As String: End Property

'@jump:display-prompts
'@prop-title Whether UI prompts are displayed during operations
'@return Boolean. True when prompts are enabled.
Public Property Get DisplayPrompts() As Boolean: End Property

'@jump:display-prompts-set
'@prop-title Toggle UI prompt display
'@param value Boolean. True to enable prompts.
Public Property Let DisplayPrompts(ByVal value As Boolean): End Property

'@jump:last-private-prompt
'@prop-title Most recent private key prompt text
'@return String. The prompt text.
Public Property Get LastPrivatePrompt() As String: End Property

'@jump:last-private-prompt-title
'@prop-title Most recent private key prompt title
'@return String. The prompt title.
Public Property Get LastPrivatePromptTitle() As String: End Property

'@jump:table-range
'@prop-title Retrieve the range of a password-managed table
'@param tableName String. Table identifier (T_keys or T_ProtectedSheets).
'@param includeHeaders Boolean. When True, includes the header row. Defaults to False.
'@return Range. The requested table range.
Public Property Get TableRange(ByVal tableName As String, Optional ByVal includeHeaders As Boolean = False) As Range: End Property

'@jump:has-checkings
'@prop-title Whether any diagnostic entries have been captured
'@return Boolean. True when at least one entry exists.
Public Property Get HasCheckings() As Boolean: End Property

'@jump:checking-values
'@prop-title Retrieve the aggregated checking entries
'@return IChecking. The checking object, or Nothing.
Public Property Get CheckingValues() As IChecking: End Property

'@jump:is-in-debug-mode
'@prop-title Whether the workbook is currently in debug mode
'@return Boolean. True when debug mode is active.
Public Property Get IsInDebugMode() As Boolean: End Property


'@section Debug Mode
'===============================================================================
'@description Toggle debug mode on the target workbook.

'@jump:enter-debug-mode
'@sub-title Put the workbook into a fully unprotected debug state
'@details
'Removes all worksheet and workbook protections and sets the debug
'mode flag. Defaults to the password workbook when no target is supplied.
'@param targetWorkbook Workbook. Optional target workbook.
Public Sub EnterDebugMode(Optional ByVal targetWorkbook As Workbook): End Sub

'@jump:leave-debug-mode
'@sub-title Restore protections after debug mode
'@details
'Reapplies worksheet-level protections from the protection matrix
'and re-locks the workbook structure.
'@param targetWorkbook Workbook. Optional target workbook.
Public Sub LeaveDebugMode(Optional ByVal targetWorkbook As Workbook): End Sub

'@jump:ensure-debug-exit-handler
'@sub-title Inject before-close procedures that leave debug mode automatically
'@param targetWorkbook Workbook. Optional target workbook.
Public Sub EnsureDebugExitHandler(Optional ByVal targetWorkbook As Workbook):
End Sub


'@section Protection
'===============================================================================
'@description Protect and unprotect worksheets and workbooks.

'@jump:protect
'@sub-title Protect a worksheet or workbook
'@details
'Accepts a Workbook, Worksheet, or String. The string "_active"
'targets the active worksheet; "_wbactive" targets the active workbook.
'@param targetObject Variant. The object to protect. Defaults to "_active".
'@param allowShapes Boolean. Allow manipulating drawing objects. Defaults to True.
'@param allowDeletingRows Boolean. Allow deleting rows. Defaults to True.
'@param registerState Boolean. When True, reuses stored table settings. Defaults to True.
Public Sub Protect(Optional targetObject As Variant = "_active", _
                   Optional ByVal allowShapes As Boolean = True, _
                   Optional ByVal allowDeletingRows As Boolean = True, _
                   Optional ByVal registerState As Boolean = True)
End Sub

'@jump:unprotect
'@sub-title Remove protection from a worksheet or workbook
'@param targetObject Variant. The object to unprotect. Defaults to "_active".
Public Sub UnProtect(Optional targetObject As Variant = "_active")
End Sub

'@jump:ensure-protected-sheets-name
'@sub-title Ensure the workbook exposes a named range referencing the protection table
'@param targetWorkbook Workbook. Optional target workbook.
Public Sub EnsureProtectedSheetsName(Optional ByVal targetWorkbook As Workbook)
End Sub

'@jump:apply-protection-matrix
'@sub-title Apply stored protection settings to all listed worksheets
'@param wb Workbook. The workbook whose sheets to protect.
Public Sub ApplyProtectionMatrix(ByVal wb As Workbook)
End Sub


'@section Keys
'===============================================================================
'@description Key generation and display.

'@jump:generate-key
'@sub-title Randomly select a public/private key pair
'@param trads ITranslationObject. Translation helper for prompts.
Public Sub GenerateKey(ByVal trads As ITranslationObject): End Sub

'@jump:display-private-key
'@sub-title Display the current private key via a translated prompt
'@param trads ITranslationObject. Translation helper for the message.
Public Sub DisplayPrivateKey(ByVal trads As ITranslationObject): End Sub


'@section DataExchange
'===============================================================================
'@description Import, export, and cloning operations.

'@jump:export-to-workbook
'@sub-title Export the password worksheet into another workbook
'@param wb Workbook. Destination workbook.
Public Sub ExportToWorkbook(ByVal wb As Workbook): End Sub

'@jump:import-from
'@sub-title Import keys and metadata from another password handler
'@param source IPasswords. Source password object.
Public Sub ImportFrom(ByVal source As IPasswords)
End Sub

'@jump:clone-to-worksheet
'@sub-title Clone the password worksheet into another sheet
'@return IPasswords. Handler bound to the cloned sheet.
Public Function CloneToWorksheet(ByVal targetSheet As Worksheet) As IPasswords
End Function

'@jump:clone-to-workbook
'@sub-title Clone the password worksheet into another workbook
'@details
'Exports the password data to the target workbook and returns a
'new handler bound to the cloned sheet. The new workbook receives
'keys, protection table, and a debug mode exit handler.
'@param wb Workbook. Destination workbook.
'@return IPasswords. Handler bound to the cloned sheet.
Public Function CloneToWorkbook(ByVal wb As Workbook) As IPasswords
End Function


'@section Utilities
'===============================================================================
'@description Miscellaneous password operations.

'@jump:new-debuging-password
'@sub-title Update the debugging password stored on the password worksheet
'@param newValue String. The new debugging password.
Public Sub NewDebugingPassword(ByVal newValue As String)
End Sub
