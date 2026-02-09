VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Guards Excel application settings while the linelist runs heavy operations"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Guards Excel application settings while the linelist runs heavy operations")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends ProjectError

Option Explicit

Implements IApplicationState

'@class ApplicationState
'@description
'Cell-based guard that snapshots Excel Application settings before
'entering a locked-down busy mode and restores them afterwards.
'Handles ScreenUpdating, DisplayAlerts, Calculation, EnableAnimations,
'EnableEvents, AutomationSecurity, cursor, and CalculateBeforeSave.
'@depends ProjectError
'@version 1.0 (2026-02-09)
'@author OutbreakTools

'@section State and constants
'===============================================================================
Private Type TScopeState
    screenUpdating As Boolean                   'Cached ScreenUpdating value
    displayAlerts As Boolean                    'Cached DisplayAlerts value
    calculationMode As XlCalculation            'Cached Calculation mode
    enableAnimations As Boolean                 'Cached EnableAnimations value
    supportsAnimations As Boolean               'Flag indicating whether EnableAnimations is available
    supportsEvents As Boolean
    cursor As XlMousePointer                    'Cached cursor value
    snapshotCaptured As Boolean                 'Flag indicating a snapshot is ready for restoration
    isBusy As Boolean
    calcBeforeSave As Boolean                    'Flag indicating busy mode is currently applied
    automationSecurity As MsoAutomationSecurity  'Cached automation security level
    securityBlocked As Boolean                   'Indicates whether automation security is currently blocked
End Type

Private this As TScopeState
Private excelApp As Application

Private Const CLASSNAME As String = "ApplicationState"


'@section Factory
'===============================================================================
'Initialisation and self-reference.

'@label:create
'@sub-title Factory returning an initialised interface
'@details
'Captures the current application state immediately so callers can
'switch to busy mode without an additional call.
'@param excelApplication Application. Target application. Defaults to Excel.Application.
'@return IApplicationState. Ready to use.
'@export
Public Function Create(Optional ByVal excelApplication As Application) As IApplicationState
    Dim scope As ApplicationState

    Set scope = New ApplicationState
    scope.BindApplication excelApplication
    scope.CaptureSnapshot

    Set Create = scope.Self
End Function

'@label:self
'@prop-title Self-reference as IApplicationState
'@return IApplicationState. Self-reference.
Public Property Get Self() As IApplicationState
    Set Self = Me
End Property


'@section Introspection
'===============================================================================
'Read-only state accessors.

'@label:application-object
'@prop-title Application reference currently guarded
'@return Application. The Excel Application instance.
Public Property Get ApplicationObject() As Application
    Set ApplicationObject = TargetApplication
End Property

'@label:is-busy
'@prop-title Whether the busy state is currently active
'@return Boolean. True when ApplyBusyState has been invoked without a matching Restore.
Public Property Get IsBusy() As Boolean
    IsBusy = this.isBusy
End Property

'@label:has-snapshot
'@prop-title Whether a snapshot exists and restoration is possible
'@return Boolean. True when a snapshot has been captured.
Public Property Get HasSnapshot() As Boolean
    HasSnapshot = this.snapshotCaptured
End Property


'@section Internal Workflow
'===============================================================================
'Friend helpers called by the factory.

'@label:bind-application
'@sub-title Bind a target application
'@param excelApplication Application. The Excel Application to guard.
Friend Sub BindApplication(ByVal excelApplication As Application)
    If excelApplication Is Nothing Then
        Set excelApp = Nothing
    ElseIf StrComp(TypeName(excelApplication), "Application", vbTextCompare) = 0 Then
        Set excelApp = excelApplication
    Else
        ThrowError ProjectError.InvalidArgument, _
                  "excelApplication must be an Excel Application reference"
    End If
End Sub

'@label:capture-snapshot
'@sub-title Capture current application settings
Friend Sub CaptureSnapshot()
    CaptureSnapshotCore
End Sub


'@section Core Logic
'===============================================================================
'Private implementation of busy-state, snapshot, and restore workflows.

'@label:apply-busy-state
'@sub-title Apply the busy state to the target application
'@details
'Switches the target Excel Application into a locked-down mode for heavy
'operations: disables screen updating, alerts, automatic calculation, and
'optionally suppresses events and blocks automation security.  Does nothing
'when the scope is already busy.  On failure the busy flag is cleared and a
'ProjectError is raised so the caller can attempt recovery.
'@param suppressEvents Optional Boolean. When True, disables Application.EnableEvents. Defaults to False.
'@param calculateOnSave Optional Boolean. Value assigned to CalculateBeforeSave. Defaults to True.
'@param busyCursor Optional XlMousePointer. Cursor shown while busy. Defaults to xlNorthWestArrow.
'@param blockSecurity Optional Boolean. When True, forces msoAutomationSecurityForceDisable. Defaults to False.
'@throws ProjectError.ErrorUnexpectedState When the application settings cannot be applied.
'@export
Private Sub ApplyBusyStateCore(Optional ByVal suppressEvents As Boolean = False, _
                               Optional ByVal calculateOnSave As Boolean = True, _
                               Optional ByVal busyCursor As XlMousePointer = xlNorthWestArrow, _
                               Optional ByVal blockSecurity As Boolean = False)
    Dim target As Application

    If this.isBusy Then Exit Sub

    EnsureSnapshot "ApplyBusyState"
    Set target = TargetApplication

    On Error GoTo ApplyFail
        target.ScreenUpdating = False
        target.DisplayAlerts = False
        target.Calculation = xlCalculationManual
        If suppressEvents Then target.EnableEvents = (Not suppressEvents)
        target.CalculateBeforeSave = calculateOnSave
        target.Cursor = busyCursor
        If blockSecurity And Not this.securityBlocked Then
            this.automationSecurity = target.AutomationSecurity
            target.AutomationSecurity = msoAutomationSecurityForceDisable
            this.securityBlocked = True
        End If
        If this.supportsAnimations Then
            SafeSetEnableAnimations target, False
        End If
        this.isBusy = True
    On Error GoTo 0
    Exit Sub

ApplyFail:
    this.isBusy = False
    ThrowError ProjectError.ErrorUnexpectedState, _
              "Unable to apply busy state: " & Err.Description
End Sub

'@label:capture-snapshot-core
'@sub-title Capture all application settings into the snapshot
Private Sub CaptureSnapshotCore()
    Dim target As Application

    EnsureNotBusy "CaptureSnapshot"
    Set target = TargetApplication

    this.screenUpdating = target.ScreenUpdating
    this.displayAlerts = target.DisplayAlerts
    this.calculationMode = target.Calculation
    this.supportsAnimations = TryReadEnableAnimations(target, this.enableAnimations)
    this.supportsEvents = target.EnableEvents
    this.calcBeforeSave = target.CalculateBeforeSave
    this.automationSecurity = target.AutomationSecurity
    this.cursor = target.Cursor
    this.snapshotCaptured = True
End Sub

'@label:refresh-snapshot
'@sub-title Re-capture the snapshot from the current application state
'@export
Private Sub RefreshSnapshotCore()
    EnsureNotBusy "RefreshSnapshot"
    CaptureSnapshotCore
End Sub

'@label:target-application
'@sub-title Resolve the Excel Application reference
'@return Application. The guarded Application instance.
Private Function TargetApplication() As Application
    If excelApp Is Nothing Then
        Set excelApp = Application
    End If

    If excelApp Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Excel application reference is required"
    End If

    Set TargetApplication = excelApp
End Function

'@label:ensure-snapshot
'@sub-title Guard that a snapshot exists before proceeding
'@param operationName String. Name of the calling operation for error messages.
Private Sub EnsureSnapshot(ByVal operationName As String)
    If Not this.snapshotCaptured Then
        CaptureSnapshotCore
    End If

    If Not this.snapshotCaptured Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                  operationName & " requires a captured snapshot. Unexpected error"
    End If
End Sub

'@label:ensure-not-busy
'@sub-title Guard that the scope is not busy before proceeding
'@param operationName String. Name of the calling operation for error messages.
Private Sub EnsureNotBusy(ByVal operationName As String)
    If this.isBusy Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                  operationName & " cannot execute while the scope is busy. Call Restore first."
    End If
End Sub

'@label:restore
'@sub-title Restore all captured settings to the application
'@param silent Boolean. When True, swallows errors instead of raising them.
'@export
Private Sub RestoreInternal(ByVal silent As Boolean)
    Dim target As Application

    If Not this.snapshotCaptured Then Exit Sub

    Set target = TargetApplication

    On Error GoTo RestoreFail
        target.ScreenUpdating = this.screenUpdating
        target.DisplayAlerts = this.displayAlerts
        target.Calculation = this.calculationMode
        target.EnableEvents = this.supportsEvents
        target.CalculateBeforeSave = this.calcBeforeSave
        If this.securityBlocked Then
            target.AutomationSecurity = this.automationSecurity
            this.securityBlocked = False
        End If
        target.Cursor = this.cursor
        If this.supportsAnimations Then
            SafeSetEnableAnimations target, this.enableAnimations
        End If
        this.isBusy = False
    On Error GoTo 0
    Exit Sub

RestoreFail:
    this.isBusy = False
    If silent Then
        Err.Clear
    Else
        ThrowError ProjectError.ErrorUnexpectedState, _
                  "Unable to restore application state: " & Err.Description
    End If
End Sub

'@label:safe-set-enable-animations
'@sub-title Set EnableAnimations with graceful fallback on unsupported platforms
'@param target Application. The Excel Application instance.
'@param value Boolean. The desired EnableAnimations state.
Private Sub SafeSetEnableAnimations(ByVal target As Application, ByVal value As Boolean)
    On Error GoTo Handler
        target.EnableAnimations = value
    On Error GoTo 0
    Exit Sub
Handler:
    'Mac Excel may not expose EnableAnimations; mark capability as unavailable for future calls.
    this.supportsAnimations = False
    Err.Clear
End Sub

'@label:try-read-enable-animations
'@sub-title Attempt to read EnableAnimations from the target application
'@param target Application. The Excel Application instance.
'@param value Boolean. ByRef. Receives the current EnableAnimations value on success.
'@return Boolean. True when the property was read successfully.
Private Function TryReadEnableAnimations(ByVal target As Application, _
                                         ByRef value As Boolean) As Boolean
    On Error GoTo Unavailable
        value = target.EnableAnimations
        TryReadEnableAnimations = True
    On Error GoTo 0
    Exit Function

Unavailable:
    value = False
    TryReadEnableAnimations = False
    Err.Clear
End Function

'@label:throw-error
'@sub-title Raise a typed project error
'@param errNumber ProjectError. The error code.
'@param message String. Descriptive error message.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, "ApplicationState", message
End Sub

'@section Interface
'===============================================================================
'IApplicationState delegation wiring.

Private Sub IApplicationState_ApplyBusyState( _
    Optional ByVal suppressEvents As Boolean = False, _
    Optional ByVal calculateOnSave As Boolean = True, _
    Optional ByVal busyCursor As XlMousePointer = xlNorthWestArrow, _
    Optional ByVal blockSecurity As Boolean = False)
    ApplyBusyStateCore suppressEvents, calculateOnSave, busyCursor, blockSecurity
End Sub

Private Sub IApplicationState_Restore()
    RestoreInternal False
End Sub

Private Sub IApplicationState_RefreshSnapshot()
    RefreshSnapshotCore
End Sub

Private Property Get IApplicationState_ApplicationObject() As Application
    Set IApplicationState_ApplicationObject = TargetApplication
End Property

Private Property Get IApplicationState_IsBusy() As Boolean
    IApplicationState_IsBusy = this.isBusy
End Property

Private Property Get IApplicationState_HasSnapshot() As Boolean
    IApplicationState_HasSnapshot = this.snapshotCaptured
End Property
