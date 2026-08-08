Attribute VB_Name = "LinelistEventsManager"
Option Explicit

'@Folder("Linelist Events")
'@ModuleDescription("Centralised workbook-level event and BusyState manager delegating to EventLinelist")
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

Public Sub LLWorkbookOpened()
    On Error GoTo Cleanup
    LLEnterBusyState
    Service.OnWorkbookOpen
Cleanup:
    LLExitBusyState
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
    On Error GoTo Cleanup
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

Public Sub SelectionChanged(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error GoTo Cleanup
    'No busyCursor here: this runs on every arrow key press, and flipping the
    'mouse cursor per keystroke is visible.
    LLEnterBusyState
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
