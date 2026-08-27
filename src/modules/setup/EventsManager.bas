Attribute VB_Name = "EventsManager"
Option Explicit

'@Folder("Setup")
'@ModuleDescription("Centralised workbook-level event and BusyState manager delegating to EventSetup")
'@depends EventSetup, ApplicationState, HiddenNames
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private setupService As EventSetup
Private appScope As ApplicationState
Private busyDepth As Long
Private persisted As Boolean
Private eventsOnly As Boolean
Private priorEvents As Boolean

Private Const SNAPSHOT_KEY As String = "APPSTATE_SNAPSHOT"

'The pointer the setup session rests on, and the pointer every busy state of it
'shows. The two being equal is what makes an operation leave no visible pointer
'change: ApplicationState snapshots the standing cursor and puts it back, so
'arrow follows arrow. On xlDefault Excel shows its own busy pointer for the
'length of every macro, which is the hourglass the setup used to flash.
Private Const RESTING_CURSOR As Long = xlNorthwestArrow

'Sheets whose change handler writes to a worksheet other than the hidden
'registry. Everything else the workbook lets through only records a flag.
Private Const SHEET_ANALYSIS As String = "Analysis"


'@section Centralised BusyState
'===============================================================================

'@sub-title Enter busy state with crash-recovery and reference-counted nesting
'@details
'On the first (outermost) call: optionally persists current Application
'properties to a HiddenName for crash recovery, creates an ApplicationState
'snapshot, and applies the locked-down busy mode.  Nested calls only increment
'busyDepth.  When persist is False, HiddenNames I/O is skipped entirely for
'fast event handlers and lightweight ribbon operations.
'@param calculateOnSave Optional Boolean. Value for CalculateBeforeSave. Defaults to True.
'@param busyCursor Optional Long. Cursor shown while busy. Defaults to the resting arrow.
'@param persist Optional Boolean. When True (default), persists snapshot to HiddenNames for crash recovery.
Public Sub EnterBusyState(Optional ByVal calculateOnSave As Boolean = True, _
                          Optional ByVal busyCursor As Long = RESTING_CURSOR, _
                          Optional ByVal persist As Boolean = True)

    If persist Then RecoverIfNeeded

    busyDepth = busyDepth + 1
    If busyDepth > 1 Then Exit Sub

    persisted = persist
    eventsOnly = False
    If persist Then PersistCurrentState

    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, _
                            calculateOnSave:=calculateOnSave, _
                            busyCursor:=busyCursor
End Sub

'@sub-title Suppress events alone, for a handler that only records a flag
'@details
'The full busy state costs thirteen Application property writes, an
'ApplicationState object, a cursor change and a Calculation flip. A handler that
'writes one flag into the hidden registry needs one of those things: events off,
'so the write does not raise the change event again. This is that state.
'
'It shares busyDepth with EnterBusyState, so ExitBusyState closes either one and
'the nesting count stays right. Enter the full state instead when the handler
'writes to a visible sheet.
Public Sub EnterQuietState()
    busyDepth = busyDepth + 1
    If busyDepth > 1 Then Exit Sub

    persisted = False
    eventsOnly = True
    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
End Sub

'@sub-title Exit busy state, restoring Application properties on the outermost call
'@details
'Decrements the nesting counter.  On the outermost exit: restores the
'ApplicationState snapshot, clears the persisted HiddenName (only when
'persistence was used), parks the pointer back on the resting arrow, and
'releases the scope reference.
'When the outermost call was EnterQuietState, only Application.EnableEvents is
'put back.
Public Sub ExitBusyState()
    If busyDepth <= 0 Then
        busyDepth = 0
        Exit Sub
    End If

    busyDepth = busyDepth - 1
    If busyDepth > 0 Then Exit Sub

    If eventsOnly Then
        On Error Resume Next
        Application.EnableEvents = priorEvents
        On Error GoTo 0
        eventsOnly = False
        Exit Sub
    End If

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    If persisted Then ClearPersistedSnapshot
    On Error GoTo 0

    'The restore has already put the snapshot cursor back. This is for the
    'snapshot that was taken on xlDefault -- a session that had another
    'workbook open before the setup arrived, or a modal form that handed the
    'pointer back on the default. Writing the arrow here is what keeps the
    'invariant the busy cursor relies on.
    RestPointer

    Set appScope = Nothing
    persisted = False
End Sub

'@sub-title Whether the manager is currently in busy state
Public Property Get IsBusyState() As Boolean
    IsBusyState = (busyDepth > 0)
End Property

'@sub-title Whether the state currently held suppresses events alone
Public Property Get IsQuietState() As Boolean
    IsQuietState = ((busyDepth > 0) And eventsOnly)
End Property

'@sub-title Put the mouse pointer back on the arrow the session rests on
'@details
'The setup opens on this arrow and every busy state shows it, so an operation
'leaves no visible pointer change at all. A modal form breaks that: Excel hands
'the pointer back on the default cursor once the form closes, and from there
'every macro of the session flashes Excel's own busy pointer. One call on the
'way out of a form session puts the invariant back.
'
'It is here rather than inline because the arrow is the manager's own answer,
'and a caller naming it would be a second owner of it.
Public Sub RestPointer()
    On Error Resume Next
    Application.Cursor = RESTING_CURSOR
    On Error GoTo 0
End Sub


'@section Crash Recovery
'===============================================================================

'@sub-title Detect and recover from a VBA state reset that occurred mid-operation
'@details
'If no scope exists and we are not nested, checks whether a persisted snapshot
'HiddenName survives from a prior interrupted operation.  If found, restores
'Application properties from the pipe-delimited string and removes the key.
'All operations are wrapped in On Error Resume Next for resilience.
Private Sub RecoverIfNeeded()
    Dim hn As HiddenNames
    Dim raw As String
    Dim parts() As String

    If Not appScope Is Nothing Then Exit Sub
    If busyDepth > 0 Then Exit Sub

    On Error Resume Next
    Set hn = HiddenNames.Create(ThisWorkbook)
    On Error GoTo 0
    If hn Is Nothing Then Exit Sub

    If Not hn.HasName(SNAPSHOT_KEY) Then Exit Sub

    raw = hn.ValueAsString(SNAPSHOT_KEY)
    hn.RemoveName SNAPSHOT_KEY

    If LenB(raw) = 0 Then Exit Sub

    parts = Split(raw, "|")
    If UBound(parts) < 5 Then Exit Sub

    'Restore in same order as ApplicationState.RestoreInternal:
    'Calculation first, ScreenUpdating last.
    On Error Resume Next
    Application.Calculation = CLng(parts(2))
    Application.DisplayAlerts = CBool(parts(1))
    Application.EnableEvents = CBool(parts(3))
    Application.CalculateBeforeSave = CBool(parts(5))
    Application.Cursor = CLng(parts(4))
    Application.ScreenUpdating = CBool(parts(0))
    On Error GoTo 0
End Sub

'@sub-title Persist current Application properties to a HiddenName before entering busy mode
'@details
'Writes a pipe-delimited string containing six Application properties to a
'single workbook-level HiddenName.  On failure the operation is silently skipped.
'Format: ScreenUpdating|DisplayAlerts|Calculation|EnableEvents|Cursor|CalcBeforeSave
'
'THE NAME IS ENSURED BEFORE IT IS SET
'-------------------------------------------------------------------------------
'HiddenNames.SetValue raises ElementNotFound on a name that is not there yet, and
'the whole routine sits under On Error Resume Next. Without the EnsureName line
'below the snapshot was never written once, so RecoverIfNeeded had nothing to
'read and the crash recovery did nothing at all.
Private Sub PersistCurrentState()
    Dim hn As HiddenNames
    Dim raw As String

    On Error Resume Next
    Set hn = HiddenNames.Create(ThisWorkbook)
    If hn Is Nothing Then GoTo CleanExit

    raw = CStr(Application.ScreenUpdating) & "|" & _
          CStr(Application.DisplayAlerts) & "|" & _
          CStr(CLng(Application.Calculation)) & "|" & _
          CStr(Application.EnableEvents) & "|" & _
          CStr(CLng(Application.Cursor)) & "|" & _
          CStr(Application.CalculateBeforeSave)

    hn.EnsureName SNAPSHOT_KEY, raw, HiddenNameTypeString
    hn.SetValue SNAPSHOT_KEY, raw
CleanExit:
    On Error GoTo 0
End Sub

'@sub-title Remove the persisted snapshot HiddenName after a successful restore
Private Sub ClearPersistedSnapshot()
    Dim hn As HiddenNames

    On Error Resume Next
    Set hn = HiddenNames.Create(ThisWorkbook)
    If Not hn Is Nothing Then
        If hn.HasName(SNAPSHOT_KEY) Then hn.RemoveName SNAPSHOT_KEY
    End If
    On Error GoTo 0
End Sub


'@section Service Lifecycle
'===============================================================================

Private Function Service() As EventSetup
    If setupService Is Nothing Then
        Set setupService = EventSetup.Create(ThisWorkbook)
    End If
    Set Service = setupService
End Function

'@sub-title Drop the cached managers the service holds
'@details
'Called after an import or a clean, both of which rewrite whole sheets. The
'managers the service cached before that point were built against columns the
'sheets may no longer carry.
Public Sub ResetEventSetupCaches()
    If Not setupService Is Nothing Then
        setupService.ResetCaches
    End If
End Sub

'@sub-title Release the service, called from Workbook_BeforeClose
Public Sub DisposeEventSetup()
    Set setupService = Nothing
End Sub


'@section Workbook Entry Points
'===============================================================================

'@sub-title Answer the open, and leave the session on the resting pointer
'@details
'The open runs inside a busy state whose snapshot was taken before the setup
'had said anything about how it wants to run, so the exit puts back whatever
'pointer the host Excel was on. The arrow is applied again once that state has
'closed, and every busy state of the session then opens and closes on it.
Public Sub WorkbookOpened()
    On Error GoTo Cleanup
    EnterBusyState
    Service.OnWorkbookOpen
Cleanup:
    If Err.Number <> 0 Then LogHandlerFailure "WorkbookOpened", ThisWorkbook.Name
    ExitBusyState
    RestPointer
End Sub

'@sub-title Route an activation under the lightest state it needs
'@details
'Activating the Analysis sheet primes three table wrappers and rebuilds the
'dropdowns over them, which writes to the sheet and wants the screen still.
'Every other activation invalidates at most three ribbon controls, and only
'when the Translations boundary is crossed.
'
'The workbook handler used to switch screen updating off for all of them and
'never switch it back, leaving Excel to do that when the handler returned. So
'each sheet the user moved to repainted the window a second time, for work that
'in nearly every case had not touched a cell.
'@param sh Worksheet. The worksheet that became active.
Public Sub SheetActivated(ByVal sh As Worksheet)
    If sh Is Nothing Then Exit Sub

    On Error GoTo Cleanup
    If StrComp(sh.Name, SHEET_ANALYSIS, vbTextCompare) = 0 Then EnterBusyState persist:=False

    Service.OnSheetActivate sh
    ExitBusyState
    Exit Sub

Cleanup:
    'An exit with nothing open does nothing, so this is owed whichever branch
    'above the failure came from.
    ExitBusyState
    LogHandlerFailure "SheetActivated", sh.Name
End Sub

'@sub-title Route a committed edit to the service under the lightest state it needs
'@details
'The workbook handler has already dropped every sheet that watches nothing, so
'what arrives here is one of the four watched sheets. Only the Analysis sheet
'clears cells, rewrites validation and unlocks a sheet, so only Analysis is
'given the full lockdown. The other three write one flag into the hidden
'registry and take the quiet state.
Public Sub SheetChanged(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error GoTo Cleanup

    If StrComp(sh.Name, SHEET_ANALYSIS, vbTextCompare) = 0 Then
        EnterBusyState persist:=False
    Else
        EnterQuietState
    End If

    Service.OnSheetChange sh, target

Cleanup:
    If Err.Number <> 0 Then LogHandlerFailure "SheetChanged", sh.Name
    ExitBusyState
End Sub

Public Sub RefreshAnalysisDropdowns(Optional ByVal forceUpdate As Boolean = False)
    On Error GoTo Cleanup
    EnterBusyState persist:=False
    Service.UpdateAnalysisDropdowns forceUpdate
Cleanup:
    If Err.Number <> 0 Then LogHandlerFailure "RefreshAnalysisDropdowns", ThisWorkbook.Name
    ExitBusyState
End Sub

Public Sub RecalculateAnalysis()
    On Error GoTo Cleanup
    EnterBusyState persist:=False
    Service.RecalculateAnalysis
Cleanup:
    If Err.Number <> 0 Then LogHandlerFailure "RecalculateAnalysis", ThisWorkbook.Name
    ExitBusyState
End Sub

Public Sub ResetTranslationCounter()
    Service.ResetTranslationCounter
End Sub

Public Function EventSetupService() As EventSetup
    Set EventSetupService = Service()
End Function


'@section Reporting
'===============================================================================

'@sub-title Write a handler failure to the immediate window
'@details
'The event handlers stay silent for the user: a message box on every committed
'edit is worse than a missing feature. They used to stay silent for the
'developer too, which made every report of a broken column unreproducible. The
'setup workbook carries no log file, so the immediate window is where this goes,
'the same place every callback in SetupRibbon already writes.
'@param handlerName String. Name of the routine that failed.
'@param context String. Sheet or workbook name the failure happened on.
Private Sub LogHandlerFailure(ByVal handlerName As String, ByVal context As String)
    On Error Resume Next
    Debug.Print "EventsManager." & handlerName & " on '" & context & "': " & _
                Err.Number & " " & Err.Description
    On Error GoTo 0
End Sub
