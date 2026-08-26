Attribute VB_Name = "LinelistEventsManager"
Option Explicit

'@Folder("Linelist Events")
'@ModuleDescription("Centralised workbook-level event and BusyState manager delegating to EventLinelist")
'@depends EventLinelist, CustomLinelistFunctions, ApplicationState
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

'@description
'Holds the single EventLinelist instance and the busy-state counter, and every
'workbook-level event of a generated linelist comes through here.
'
'WHY FOUR NAMES CARRY AN LL PREFIX
'-------------------------------------------------------------------------------
'LLEnterBusyState, LLExitBusyState, LLWorkbookOpened and LLSheetChanged answer to
'jobs that EventsManager answers to on the setup side under the same plain names.
'Both modules live in one VBA project while the test harness runs. Two Public
'procedures sharing a name make every unqualified call to that name "Ambiguous
'name detected", which stops the whole project compiling, so the linelist side of
'each pair carries the prefix. The three handlers below them -- SheetDeactivated,
'SelectionChanged and DoubleClicked -- have no twin on the setup side and keep
'their plain names.

Private linelistService As EventLinelist
Private appScope As ApplicationState
Private busyDepth As Long
Private quietDepth As Long
Private quietEvents As Boolean


'@section Centralised BusyState
'===============================================================================

'@sub-title Enter busy state, with reference-counted nesting
'@details
'On the first (outermost) call, takes a fresh snapshot of the application
'settings and applies the locked-down busy mode. Nested calls only increment
'busyDepth.
'
'One ApplicationState serves the whole session. Building a new one per event
'reads eight Application properties, and one of them, EnableAnimations, is read
'through a raise-and-catch on Mac Excel, so a keystroke paid for that twice.
'RefreshSnapshot re-reads the same eight into the instance already held, which
'is what makes the Restore below put back what the user had.
'@param calculateOnSave Optional Boolean. Value for CalculateBeforeSave. Defaults to True.
'@param busyCursor Optional Long. Cursor shown while busy. When 0 (default), leaves cursor unchanged.
Public Sub LLEnterBusyState(Optional ByVal calculateOnSave As Boolean = True, _
                          Optional ByVal busyCursor As Long = 0)

    busyDepth = busyDepth + 1
    If busyDepth > 1 Then Exit Sub

    If appScope Is Nothing Then
        'Create takes the first snapshot itself.
        Set appScope = ApplicationState.Create(Application)
    ElseIf Not appScope.IsBusy Then
        'RefreshSnapshot raises while the scope is busy, which is the state an
        'exit that never ran leaves behind. The snapshot it still holds is the
        'one to restore from, so skipping the refresh is the safe answer.
        appScope.RefreshSnapshot
    End If

    appScope.ApplyBusyState suppressEvents:=True, _
                            calculateOnSave:=calculateOnSave, _
                            busyCursor:=busyCursor
End Sub

'@sub-title Exit busy state, restoring Application properties on the outermost call
'@details
'Decrements the nesting counter. On the outermost exit: restores the
'ApplicationState snapshot, which puts back the cursor it captured. The scope
'itself is kept for the next event, and the next LLEnterBusyState refreshes its
'snapshot.
Public Sub LLExitBusyState()
    If busyDepth <= 0 Then
        busyDepth = 0
        Exit Sub
    End If

    busyDepth = busyDepth - 1
    If busyDepth > 0 Then Exit Sub

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0
End Sub


'@section Centralised quiet state
'===============================================================================
'The events-only half of the busy state, for work that needs nothing else than
'silence: the import resizing every data entry table before it reads a file, and
'the geo form writing a place into the cells beside it. Both used to write
'Application.EnableEvents themselves, so one flag had three owners in the
'linelist and a run that ended between two of those lines left worksheet events
'off for the rest of the session. This is the one owner.
'
'It is deliberately NOT the busy state. Screen updating, manual calculation and
'the cursor are visible, and a form on screen writing four cells has no business
'taking any of them.

'@sub-title Turn worksheet events off for a stretch of work, counting the nesting.
'@details
'What was found is saved rather than assumed, so a quiet stretch that opens
'inside a busy one puts the busy value back and lets the busy exit restore what
'the user had. That is what makes the two safe to nest either way round.
Public Sub LLEnterQuietState()
    quietDepth = quietDepth + 1
    If quietDepth > 1 Then Exit Sub

    quietEvents = Application.EnableEvents
    Application.EnableEvents = False
End Sub

'@sub-title Give worksheet events back on the outermost exit.
'@details
'An exit with nothing open does nothing, so a handler may call this whether or
'not the raise happened inside the quiet stretch. Every caller of
'LLEnterQuietState owes this one on every path it can take, error label
'included -- that is the whole point of it being here rather than inline.
Public Sub LLExitQuietState()
    If quietDepth <= 0 Then
        quietDepth = 0
        Exit Sub
    End If

    quietDepth = quietDepth - 1
    If quietDepth > 0 Then Exit Sub

    On Error Resume Next
    Application.EnableEvents = quietEvents
    On Error GoTo 0
End Sub


'@section The resting pointer
'===============================================================================

'@sub-title Put the mouse pointer back on the arrow the session rests on.
'@details
'OnWorkbookOpen parks the pointer on the north-west arrow, and every busy state
'of the session shows that same arrow. The two being equal is what makes an
'event leave no visible pointer change: ApplicationState snapshots the standing
'cursor and puts it back, so arrow follows arrow and the user sees nothing.
'
'A modal form breaks that. Excel hands the pointer back on the default cursor
'once the form closes, so the standing cursor is no longer the arrow, and from
'there every selection on a data entry sheet flicks the pointer twice -- to the
'arrow on the way in and to the default on the way out. The form is gone by
'then, so the flicking outlives the thing that caused it.
'
'One call on the way out of a form session puts the invariant back. It is here
'rather than inline because the arrow is the manager's own answer, and a caller
'naming xlNorthwestArrow itself would be a second owner of it.
Public Sub LLRestPointer()
    On Error Resume Next
    Application.Cursor = xlNorthwestArrow
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

Public Sub DisposeEventLinelist()
    Set linelistService = Nothing
    Set appScope = Nothing
End Sub

Public Function EventLinelistService() As EventLinelist
    Set EventLinelistService = Service()
End Function


'@section Workbook Entry Points
'===============================================================================

'@sub-title Answer the open, and leave the session on the linelist resting state.
'@details
'The resting state is applied a second time, after the busy state has closed.
'The open runs inside a busy state whose snapshot was taken before the linelist
'had said anything about how it wants to run, so the exit put manual
'calculation and the resting pointer straight back to whatever the host Excel
'was on. A session that already had a workbook open when the linelist arrived
'therefore ran the linelist on automatic calculation, and from there every
'event of the workbook -- a cell selected, a sheet left, a click on Add rows --
'ended in a recalculation of the whole linelist on the way out of its busy
'state.
Public Sub LLWorkbookOpened()
    On Error GoTo Cleanup
    LLEnterBusyState
    Service.OnWorkbookOpen
Cleanup:
    LLExitBusyState

    'Nothing on the open is worth an error box, and the line above has already
    'given the user their screen back.
    On Error Resume Next
    Service.ApplyRestingState
    On Error GoTo 0
End Sub

'@sub-title Route the deactivation of one HList sheet.
'@details
'Worksheet_Deactivate is written into the code module of the HList sheets alone,
'so this runs when the user leaves a data entry sheet and nowhere else. It used
'to answer Workbook_SheetDeactivate, which fires for every analysis, geo,
'dropdown and temporary sheet of the workbook, and each of those paid a busy
'state and a read of the workbook hidden names to check one flag that only an
'HList edit ever sets.
'@param sh Worksheet. The sheet that was left.
Public Sub SheetDeactivated(ByVal sh As Worksheet)
    If sh Is Nothing Then Exit Sub

    'The busy state is asked for only when the sheet has a rebuild waiting on
    'it. Taking it is not free and it is not invisible: screen updating off and
    'back on repaints the window, so leaving a sheet nobody had edited used to
    'flicker for a flag that said there was nothing to do.
    On Error GoTo Cleanup
    If Not Service.SheetHasListAutoWork(sh) Then Exit Sub

    LLEnterBusyState busyCursor:=xlNorthwestArrow
    Service.OnSheetDeactivate sh
Cleanup:
    LLExitBusyState
End Sub

Public Sub LLSheetChanged(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error GoTo Cleanup
    ' The edited sheet may be a VALUE_OF lookup table, and that cache outlives
    ' a recalculation. Drop its slot before anything recalculates.
    CustomLinelistFunctions.ResetValueOfCache sh.Name
    LLEnterBusyState busyCursor:=xlNorthwestArrow
    Service.OnSheetChange sh, target
Cleanup:
    LLExitBusyState
End Sub

'@sub-title Route a selection under the lightest state that fits it.
'@details
'Landing on an admin cell refills the dropdown under it from the geobase, which
'is the slowest thing a selection can start and the only one worth hiding the
'screen for. It takes the busy state, and the north-west arrow with it: the
'edit that may follow runs under the same arrow, and the cursor changing
'between the two is a flick the owner reported.
'
'Every other selection recalculates the row under the cursor and stops. That
'used to take the busy state as well, and it is the visible one of the two:
'screen updating off and back on repaints the window, so moving round a data
'entry sheet with the arrow keys flickered once per key. The quiet state writes
'nothing the user can see.
Public Sub SelectionChanged(ByVal sh As Worksheet, ByVal target As Range)
    Dim needsBusyState As Boolean

    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error GoTo Cleanup
    needsBusyState = Service.SelectionIsHeavy(sh, target)

    If needsBusyState Then
        LLEnterBusyState busyCursor:=xlNorthwestArrow
    Else
        LLEnterQuietState
    End If

    Service.OnSelectionChange sh, target

Cleanup:
    If needsBusyState Then
        LLExitBusyState
    Else
        LLExitQuietState
    End If
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
    LLEnterBusyState busyCursor:=xlNorthwestArrow
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
    LLEnterBusyState busyCursor:=xlNorthwestArrow
    Service.UpdateFilterTables calculate
Cleanup:
    LLExitBusyState
End Sub

'@EntryPoint
Public Sub UpdateAllListAuto()
    On Error GoTo Cleanup
    LLEnterBusyState busyCursor:=xlNorthwestArrow
    Service.UpdateAllListAuto
Cleanup:
    LLExitBusyState
End Sub
