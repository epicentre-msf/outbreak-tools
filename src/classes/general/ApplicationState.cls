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

'@class ApplicationState
'@description
'Holds a copy of the Excel Application settings a long operation changes, and
'puts them back when it ends. Create binds an Application and takes the copy
'straight away. ApplyBusyState turns screen updating, alerts and automatic
'calculation off, and on request also silences events, forces the automation
'security level and sets the cursor. Restore writes every saved value back.
'IsBusy and HasSnapshot answer where the object stands.
'
'THE COPY IS TAKEN ONCE
'-------------------------------------------------------------------------------
'ApplyBusyState restores from the copy the object already holds, and takes no
'new one. RefreshSnapshot takes a fresh copy, which is what a caller that
'changed a setting between two busy blocks calls first.
'
'THE SETTINGS THIS CLASS COVERS
'-------------------------------------------------------------------------------
'ScreenUpdating, DisplayAlerts, Calculation, EnableAnimations, EnableEvents,
'AutomationSecurity, the mouse cursor and CalculateBeforeSave. EnableAnimations
'is missing on some hosts, and the object records whether it is there.
'@depends ProjectError
'@version 1.0 (2026-02-09)


'@section State and constants
'===============================================================================
Private Type TApplicationState
    excelApp As Application                      'Guarded Excel Application reference
    screenUpdating As Boolean                    'Cached ScreenUpdating value
    displayAlerts As Boolean                     'Cached DisplayAlerts value
    calculationMode As XlCalculation             'Cached Calculation mode
    enableAnimations As Boolean                  'Cached EnableAnimations value
    supportsAnimations As Boolean                'Flag indicating whether EnableAnimations is available
    enableEvents As Boolean                      'Cached EnableEvents value
    cursor As XlMousePointer                     'Cached cursor value
    snapshotCaptured As Boolean                  'Flag indicating a snapshot is ready for restoration
    isBusy As Boolean                            'Flag indicating busy mode is currently applied
    calcBeforeSave As Boolean                    'Cached CalculateBeforeSave value
    automationSecurity As MsoAutomationSecurity  'Cached automation security level
    securityBlocked As Boolean                   'Indicates whether automation security is currently blocked
End Type

Private this As TApplicationState

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
'@return ApplicationState. Ready to use.
'@export
Public Function Create(Optional ByVal excelApplication As Application) As ApplicationState
    Dim scope As ApplicationState

    Set scope = New ApplicationState
    scope.BindApplication excelApplication
    scope.RefreshSnapshot

    Set Create = scope
End Function

'@section Introspection
'===============================================================================
'Read-only state accessors.

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
        Set this.excelApp = Nothing
    ElseIf StrComp(TypeName(excelApplication), "Application", vbTextCompare) = 0 Then
        Set this.excelApp = excelApplication
    Else
        ThrowError ProjectError.InvalidArgument, _
                  "excelApplication must be an Excel Application reference"
    End If
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
'The snapshot restored later is the one this object already holds; it belongs
'to the whole life of the object and is NOT re-taken here.  A caller that
'changed a setting on purpose between two busy blocks must ask for a new
'snapshot with RefreshSnapshot, otherwise the next Restore puts the old value
'back.
'@param suppressEvents Optional Boolean. When True, disables Application.EnableEvents. Defaults to False.
'@param calculateOnSave Optional Boolean. Value assigned to CalculateBeforeSave. Defaults to True.
'@param busyCursor Optional Long. Cursor shown while busy. When 0 (default), leaves cursor unchanged.
'@param blockSecurity Optional Boolean. When True, forces msoAutomationSecurityForceDisable. Defaults to False.
'@throws ProjectError.ErrorUnexpectedState When the application settings cannot be applied.
'@export
Public Sub ApplyBusyState(Optional ByVal suppressEvents As Boolean = False, _
                               Optional ByVal calculateOnSave As Boolean = True, _
                               Optional ByVal busyCursor As Long = 0, _
                               Optional ByVal blockSecurity As Boolean = False)
    Dim target As Application

    If this.isBusy Then Exit Sub

    EnsureSnapshot
    Set target = TargetApplication

    On Error GoTo ApplyFail
        target.ScreenUpdating = False
        target.DisplayAlerts = False
        'Setting Calculation to Manual can trigger a flush of pending
        'worksheet calculations.  If any worksheet has a formula that
        'cannot evaluate, the flush raises error 1004.  Isolate so the
        'remaining settings are still applied.
        On Error Resume Next
        target.Calculation = xlCalculationManual
        On Error GoTo ApplyFail
        If suppressEvents Then target.EnableEvents = False
        target.CalculateBeforeSave = calculateOnSave
        If busyCursor <> 0 Then target.Cursor = busyCursor
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

    Set target = TargetApplication

    this.screenUpdating = target.ScreenUpdating
    this.displayAlerts = target.DisplayAlerts
    this.calculationMode = target.Calculation
    this.supportsAnimations = TryReadEnableAnimations(target, this.enableAnimations)
    this.enableEvents = target.EnableEvents
    this.calcBeforeSave = target.CalculateBeforeSave
    this.automationSecurity = target.AutomationSecurity
    this.cursor = target.Cursor
    this.snapshotCaptured = True
End Sub

'@label:refresh-snapshot
'@sub-title Re-capture the snapshot from the current application state
'@throws ProjectError.ErrorUnexpectedState When the scope is busy.
'@export
Public Sub RefreshSnapshot()
    EnsureNotBusy "RefreshSnapshot"
    CaptureSnapshotCore
End Sub

'@label:target-application
'@sub-title Resolve the Excel Application reference
'@return Application. The guarded Application instance.
Private Function TargetApplication() As Application
    If this.excelApp Is Nothing Then
        Set this.excelApp = Application
    End If

    If this.excelApp Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Excel application reference is required"
    End If

    Set TargetApplication = this.excelApp
End Function

'@label:ensure-snapshot
'@sub-title Guard that a snapshot exists before proceeding
'@details
'Takes the snapshot only when the object does not have one yet.  The snapshot
'belongs to the whole life of the object: once taken it stays until a caller
'asks for a new one with RefreshSnapshot.
Private Sub EnsureSnapshot()
    If Not this.snapshotCaptured Then
        CaptureSnapshotCore
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
'@param silent Optional Boolean. When True, swallows errors instead of raising them. Defaults to False.
'@export
Public Sub Restore(Optional ByVal silent As Boolean = False)
    Dim target As Application
    Dim failMessage As String

    If Not this.snapshotCaptured Then Exit Sub

    Set target = TargetApplication

    On Error GoTo RestoreFail
        'Restore Calculation BEFORE ScreenUpdating so any recalculation
        'triggered by switching from Manual to Automatic happens with
        'the screen still suppressed, preventing a visible flash.
        'Isolated: restoring Auto can flush pending calculations and
        'raise error 1004 if a worksheet formula cannot evaluate.
        On Error Resume Next
        target.Calculation = this.calculationMode
        On Error GoTo RestoreFail
        target.DisplayAlerts = this.displayAlerts
        target.EnableEvents = this.enableEvents
        target.CalculateBeforeSave = this.calcBeforeSave
        If this.securityBlocked Then
            target.AutomationSecurity = this.automationSecurity
            this.securityBlocked = False
        End If
        target.Cursor = this.cursor
        If this.supportsAnimations Then
            SafeSetEnableAnimations target, this.enableAnimations
        End If
        'ScreenUpdating must be restored LAST so that all other settings
        '(especially Calculation mode) are already in place before Excel
        'repaints the screen.
        target.ScreenUpdating = this.screenUpdating
        this.isBusy = False
    On Error GoTo 0
    Exit Sub

RestoreFail:
    failMessage = Err.Description
    this.isBusy = False
    'Put the screen back even when the restore failed.  Without this line a
    'failure above leaves the whole Excel session with a frozen screen, which
    'is the state this class exists to prevent.
    On Error Resume Next
    target.ScreenUpdating = this.screenUpdating
    On Error GoTo 0
    If silent Then
        Err.Clear
    Else
        ThrowError ProjectError.ErrorUnexpectedState, _
                  "Unable to restore application state: " & failMessage
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
    Err.Raise errNumber, CLASSNAME, message
End Sub
