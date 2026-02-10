VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignerPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Prepares designer workbooks and ribbon settings"

'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("Prepares designer workbooks and persists ribbon settings through hidden names.")
'@depends ProjectError, HiddenNames, IHiddenNames, IDevelopment
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ParameterNotUsed

Option Explicit
Implements IDesignerPreparation

'@class DesignerPreparation
'@description
'Prepares a designer workbook for ribbon interactions by initialising a
'workbook-level hidden names store and ensuring default checkbox flags
'(alert, instruct) exist. Provides methods to read and write ribbon
'settings as boolean flags persisted through hidden names, so the
'designer ribbon can restore its state across sessions. Consumers
'interact through the IDesignerPreparation interface.
'@depends HiddenNames, IHiddenNames, IDevelopment


'@section State
'===============================================================================

Private Type TDesignerPreparation
    hostBook As Workbook       'Designer workbook being prepared
    store As IHiddenNames      'Workbook-level hidden names manager
End Type

Private this As TDesignerPreparation

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "DesignerPreparation"
Private Const FLAG_ALERT As String = "chkAlert"
Private Const FLAG_INSTRUCT As String = "chkInstruct"
Private Const FLAG_ON As String = "Yes"
Private Const FLAG_OFF As String = "No"


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create and configure a designer preparation helper
'@details
'Factory method that creates a new DesignerPreparation instance, binds
'it to the supplied designer workbook via Configure, and returns the
'result through the IDesignerPreparation interface. Called on the
'predeclared instance.
'@param hostBook Workbook. The designer workbook to prepare.
'@return IDesignerPreparation. A fully initialised instance ready for use.
'@export
Public Function Create(ByVal hostBook As Workbook) As IDesignerPreparation
    Dim instance As DesignerPreparation

    Set instance = New DesignerPreparation
    instance.Configure hostBook
    Set Create = instance.Self
End Function

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor so consuming code can fluently retrieve the
'interface reference from the predeclared Create method.
'@return IDesignerPreparation. The current instance cast to the interface.
Public Property Get Self() As IDesignerPreparation
    Set Self = Me
End Property


'@section Public API
'===============================================================================

'@label:configure
'@sub-title Bind the helper to the designer workbook
'@details
'Stores a reference to the host workbook and clears any previously
'cached hidden names store so it will be rebuilt on next access. Raises
'an error when the workbook is Nothing.
'@param hostBook Workbook. The designer workbook to bind to.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
    If hostBook Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host workbook is required."
    End If

    Set this.hostBook = hostBook
    Set this.store = Nothing
End Sub

'@label:prepare
'@sub-title Prepare workbook-level settings and ensure ribbon flags exist
'@details
'Ensures the hidden names store is initialised from the host workbook and
'that both default ribbon flags (chkAlert, chkInstruct) exist with "No"
'as default. The devManager parameter is accepted for interface parity
'with setup preparation but is not currently used for designer work.
'@param devManager IDevelopment. ByRef development manager instance, or Nothing.
'@export
Public Sub Prepare(ByRef devManager As IDevelopment)
    EnsureHiddenStore
    EnsureDefaultFlags

    'Keep the parameter consumed to satisfy interface parity with setup preparation.
    If Not devManager Is Nothing Then
        'No designer-specific dev work required yet.
    End If
End Sub

'@label:set-flag
'@sub-title Persist a ribbon checkbox flag as Yes/No in hidden names
'@details
'Writes a boolean value to the workbook-level hidden names store as
'"Yes" or "No". Ensures the store and default flags exist before writing.
'@param nameId String. The hidden name identifier for the flag (e.g. "chkAlert").
'@param enabled Boolean. True to store "Yes", False to store "No".
'@export
Public Sub SetFlag(ByVal nameId As String, ByVal enabled As Boolean)
    EnsureHiddenStore
    EnsureDefaultFlags
    this.store.SetValue nameId, IIf(enabled, FLAG_ON, FLAG_OFF)
End Sub

'@label:get-flag
'@sub-title Read a ribbon checkbox flag as a boolean from hidden names
'@details
'Reads the named value from the hidden names store and returns True when
'the stored value equals "Yes" (case-insensitive comparison). Falls back
'to the defaultValue when the name is missing.
'@param nameId String. The hidden name identifier for the flag.
'@param defaultValue Optional Boolean. Value returned when the flag is missing. Defaults to False.
'@return Boolean. True when the stored value is "Yes".
'@export
Public Function GetFlag(ByVal nameId As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim rawValue As String
    EnsureHiddenStore
    EnsureDefaultFlags

    rawValue = this.store.ValueAsString(nameId, IIf(defaultValue, FLAG_ON, FLAG_OFF))
    GetFlag = (StrComp(rawValue, FLAG_ON, vbTextCompare) = 0)
End Function

'@label:host-workbook
'@prop-title The configured host workbook
'@details
'Returns the designer workbook that was bound via Configure. Raises an
'error when accessed before configuration.
'@return Workbook. The host workbook.
'@throws ProjectError.ObjectNotInitialized When the workbook was not configured.
'@export
Public Property Get HostWorkbook() As Workbook
    If this.hostBook Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host workbook was not configured."
    End If
    Set HostWorkbook = this.hostBook
End Property

'@label:hidden-store
'@prop-title The workbook-level hidden names manager
'@details
'Returns the IHiddenNames instance lazily created from the host workbook.
'Used internally to persist and retrieve ribbon flag values.
'@return IHiddenNames. The hidden names manager.
'@export
Public Property Get HiddenStore() As IHiddenNames
    EnsureHiddenStore
    Set HiddenStore = this.store
End Property


'@section Internal Helpers
'===============================================================================

'@label:ensure-hidden-store
'@sub-title Lazily initialise the hidden names store from the host workbook
'@details
'Creates the IHiddenNames instance on first access by calling
'HiddenNames.Create with the host workbook. Subsequent calls are no-ops.
Private Sub EnsureHiddenStore()
    If this.store Is Nothing Then
        Set this.store = HiddenNames.Create(HostWorkbook)
    End If
End Sub

'@label:ensure-default-flags
'@sub-title Ensure both default ribbon flags exist in the hidden store
'@details
'Calls EnsureName for chkAlert and chkInstruct with "No" as default,
'creating the hidden names if they do not already exist.
Private Sub EnsureDefaultFlags()
    this.store.EnsureName FLAG_ALERT, FLAG_OFF, HiddenNameTypeString
    this.store.EnsureName FLAG_INSTRUCT, FLAG_OFF, HiddenNameTypeString
End Sub

'@label:raise-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME.
'@param errorCode ProjectError. The error code to raise.
'@param messageText String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub RaiseError(ByVal errorCode As ProjectError, ByVal messageText As String)
    Err.Raise CLng(errorCode), CLASS_NAME, messageText
End Sub


'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the IDesignerPreparation contract.
'See the corresponding members above for full documentation.

Private Sub IDesignerPreparation_Configure(ByVal hostBook As Workbook)
    Configure hostBook
End Sub

Private Function IDesignerPreparation_GetFlag(ByVal nameId As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    IDesignerPreparation_GetFlag = GetFlag(nameId, defaultValue)
End Function

Private Property Get IDesignerPreparation_HiddenStore() As IHiddenNames
    Set IDesignerPreparation_HiddenStore = HiddenStore
End Property

Private Property Get IDesignerPreparation_HostWorkbook() As Workbook
    Set IDesignerPreparation_HostWorkbook = HostWorkbook
End Property

Private Sub IDesignerPreparation_Prepare(ByRef devManager As IDevelopment)
    Prepare devManager
End Sub

Private Sub IDesignerPreparation_SetFlag(ByVal nameId As String, ByVal enabled As Boolean)
    SetFlag nameId, enabled
End Sub
