Attribute VB_Name = "LinelistEventsManager"
Option Explicit

'@Folder("Linelist Events")
'@ModuleDescription("Centralised workbook-level event and BusyState manager delegating to EventLinelist")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

'@description
'Holds the single EventLinelist instance and the busy-state counter, and every
'workbook-level event of a generated linelist comes through here.
'
'WHY SIX NAMES CARRY AN LL PREFIX
'-------------------------------------------------------------------------------
'LLEnterBusyState, LLExitBusyState, LLIsBusyState, LLWorkbookOpened,
'LLSheetActivated and LLSheetChanged answer to the same six jobs that
'EventsManager answers to on the setup side. Both modules live in one VBA
'project while the test harness runs. Two Public procedures sharing a name make
'every unqualified call to that name "Ambiguous name detected", which stops the
'whole project compiling, so the linelist side of each pair carries the prefix.
'The three handlers below it -- SheetDeactivated, SelectionChanged and
'DoubleClicked -- have no twin on the setup side and keep their plain names.

Private linelistService As EventLinelist
Private appScope As ApplicationState
Private busyDepth As Long
Private persisted As Boolean

Private Const SNAPSHOT_KEY As String = "APPSTATE_SNAPSHOT"


'@section Centralised BusyState
'===============================================================================

'@sub-title Enter busy state with crash-recovery and reference-counted nesting
'@details
'On the first (outermost) call: optionally persists current Application
'properties to a HiddenName for crash recovery, creates an ApplicationState
'snapshot, and applies the locked-down busy mode. Nested calls only increment
'busyDepth. When persist is False, HiddenNames I/O is skipped entirely for
'fast event handlers and lightweight operations.
'@param calculateOnSave Optional Boolean. Value for CalculateBeforeSave. Defaults to True.
'@param busyCursor Optional Long. Cursor shown while busy. When 0 (default), leaves cursor unchanged.
'@param persist Optional Boolean. When True (default), persists snapshot to HiddenNames for crash recovery.
Public Sub LLEnterBusyState(Optional ByVal calculateOnSave As Boolean = True, _
                          Optional ByVal busyCursor As Long = 0, _
                          Optional ByVal persist As Boolean = True)

    If persist Then RecoverIfNeeded

    busyDepth = busyDepth + 1
    If busyDepth > 1 Then Exit Sub

    persisted = persist
    If persist Then PersistCurrentState

    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, _
                            calculateOnSave:=calculateOnSave, _
                            busyCursor:=busyCursor
End Sub

'@sub-title Exit busy state, restoring Application properties on the outermost call
'@details
'Decrements the nesting counter. On the outermost exit: restores the
'ApplicationState snapshot, clears the persisted HiddenName (only when
'persistence was used), resets the cursor, and releases the scope reference.
Public Sub LLExitBusyState()
    If busyDepth <= 0 Then
        busyDepth = 0
        Exit Sub
    End If

    busyDepth = busyDepth - 1
    If busyDepth > 0 Then Exit Sub

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    If persisted Then ClearPersistedSnapshot
    Application.Cursor = xlDefault
    On Error GoTo 0

    Set appScope = Nothing
    persisted = False
End Sub

'@sub-title Whether the manager is currently in busy state
Public Property Get LLIsBusyState() As Boolean
    LLIsBusyState = (busyDepth > 0)
End Property


'@section Crash Recovery
'===============================================================================

'@sub-title Detect and recover from a VBA state reset that occurred mid-operation
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

Private Function Service() As EventLinelist
    If linelistService Is Nothing Then
        Set linelistService = EventLinelist.Create(ThisWorkbook)
    End If
    Set Service = linelistService
End Function

Public Sub ResetEventLinelistCaches()
    If Not linelistService Is Nothing Then
        linelistService.ResetCaches
    End If
End Sub

Public Sub DisposeEventLinelist()
    Set linelistService = Nothing
End Sub

Public Function EventLinelistService() As EventLinelist
    Set EventLinelistService = Service()
End Function


'@section Workbook Entry Points
'===============================================================================

Public Sub LLWorkbookOpened()
    On Error GoTo Cleanup
    LLEnterBusyState
    Service.OnWorkbookOpen
Cleanup:
    LLExitBusyState
End Sub

Public Sub LLSheetActivated(ByVal sh As Worksheet)
    'No specific activate handling needed for linelist (yet)
    'Placeholder for future use (ribbon invalidation, etc.)
    If sh Is Nothing Then Exit Sub
End Sub

Public Sub SheetDeactivated(ByVal sh As Worksheet)
    If sh Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    LLEnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.OnSheetDeactivate sh.Name
Cleanup:
    LLExitBusyState
End Sub

Public Sub LLSheetChanged(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error GoTo Cleanup
    ' The edited sheet may be a VALUE_OF lookup table, and that cache outlives
    ' a recalculation. Drop its slot before anything recalculates.
    CustomLinelistFunctions.ResetValueOfCache sh.Name
    LLEnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.OnSheetChange sh, target
Cleanup:
    LLExitBusyState
End Sub

Public Sub SelectionChanged(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error GoTo Cleanup
    LLEnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.OnSelectionChange sh, target
Cleanup:
    LLExitBusyState
End Sub

'@sub-title Route a double-click and answer the geo picker it asked for.
'@details
'The picker is a UserForm and lives in the workbook, so the answer travels up
'to EventLinelistWorkbook and that module opens it.
'@param sh Worksheet. The worksheet where the double-click happened.
'@param target Range. The double-clicked cell.
'@return Long. A GeoScope value, or a negative number when no picker is wanted.
Public Function DoubleClicked(ByVal sh As Worksheet, ByVal target As Range) As Long
    DoubleClicked = -1
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Function

    On Error GoTo Cleanup
    LLEnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    DoubleClicked = Service.OnDoubleClick(sh, target)
Cleanup:
    LLExitBusyState
End Function


'@section Public Entry Points for External Callers
'===============================================================================
'These subs are called by name from other modules (GeoModule, EventsLinelistButtons,
'AnalysisOutput via Application.Run) and must remain publicly accessible.

'@EntryPoint
Public Sub UpdateFilterTables(Optional ByVal calculate As Boolean = True)
    On Error GoTo Cleanup
    LLEnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.UpdateFilterTables calculate
Cleanup:
    LLExitBusyState
End Sub

'@EntryPoint
Public Sub UpdateAllListAuto()
    On Error GoTo Cleanup
    LLEnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.UpdateAllListAuto
Cleanup:
    LLExitBusyState
End Sub
