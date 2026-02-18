Attribute VB_Name = "LinelistEventsManager"
Option Explicit

'@Folder("Linelist Events")
'@ModuleDescription("Centralised workbook-level event and BusyState manager delegating to EventLinelist")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private linelistService As IEventLinelist
Private appScope As IApplicationState
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
Public Sub EnterBusyState(Optional ByVal calculateOnSave As Boolean = True, _
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
Public Sub ExitBusyState()
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
Public Property Get IsBusyState() As Boolean
    IsBusyState = (busyDepth > 0)
End Property


'@section Crash Recovery
'===============================================================================

'@sub-title Detect and recover from a VBA state reset that occurred mid-operation
Private Sub RecoverIfNeeded()
    Dim hn As IHiddenNames
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
    Dim hn As IHiddenNames
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
    Dim hn As IHiddenNames

    On Error Resume Next
    Set hn = HiddenNames.Create(ThisWorkbook)
    If Not hn Is Nothing Then
        If hn.HasName(SNAPSHOT_KEY) Then hn.RemoveName SNAPSHOT_KEY
    End If
    On Error GoTo 0
End Sub


'@section Service Lifecycle
'===============================================================================

Private Function Service() As IEventLinelist
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


'@section Workbook Entry Points
'===============================================================================

Public Sub WorkbookOpened()
    On Error GoTo Cleanup
    EnterBusyState
    Service.OnWorkbookOpen
Cleanup:
    ExitBusyState
End Sub

Public Sub SheetActivated(ByVal sh As Worksheet)
    'No specific activate handling needed for linelist (yet)
    'Placeholder for future use (ribbon invalidation, etc.)
    If sh Is Nothing Then Exit Sub
End Sub

Public Sub SheetDeactivated(ByVal sh As Worksheet)
    If sh Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    EnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.OnSheetDeactivate sh.Name
Cleanup:
    ExitBusyState
End Sub

Public Sub SheetChanged(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error GoTo Cleanup
    EnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.OnSheetChange sh, target
Cleanup:
    ExitBusyState
End Sub

Public Sub SelectionChanged(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error GoTo Cleanup
    EnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.OnSelectionChange sh, target
Cleanup:
    ExitBusyState
End Sub

Public Sub DoubleClicked(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error GoTo Cleanup
    EnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.OnDoubleClick sh, target
Cleanup:
    ExitBusyState
End Sub


'@section Public Entry Points for External Callers
'===============================================================================
'These subs are called by name from other modules (GeoModule, EventsLinelistButtons,
'AnalysisOutput via Application.Run) and must remain publicly accessible.

'@EntryPoint
Public Sub UpdateFilterTables(Optional ByVal calculate As Boolean = True)
    On Error GoTo Cleanup
    EnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.UpdateFilterTables calculate
Cleanup:
    ExitBusyState
End Sub

'@EntryPoint
Public Sub UpdateAllListAuto()
    On Error GoTo Cleanup
    EnterBusyState busyCursor:=xlNorthwestArrow, persist:=False
    Application.ScreenUpdating = False
    Service.UpdateAllListAuto
Cleanup:
    ExitBusyState
End Sub
