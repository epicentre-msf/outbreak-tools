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
'On the first (outermost) call, creates an ApplicationState snapshot and applies
'the locked-down busy mode. Nested calls only increment busyDepth.
'@param calculateOnSave Optional Boolean. Value for CalculateBeforeSave. Defaults to True.
'@param busyCursor Optional Long. Cursor shown while busy. When 0 (default), leaves cursor unchanged.
Public Sub LLEnterBusyState(Optional ByVal calculateOnSave As Boolean = True, _
                          Optional ByVal busyCursor As Long = 0)

    busyDepth = busyDepth + 1
    If busyDepth > 1 Then Exit Sub

    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, _
                            calculateOnSave:=calculateOnSave, _
                            busyCursor:=busyCursor
End Sub

'@sub-title Exit busy state, restoring Application properties on the outermost call
'@details
'Decrements the nesting counter. On the outermost exit: restores the
'ApplicationState snapshot, resets the cursor, and releases the scope reference.
Public Sub LLExitBusyState()
    If busyDepth <= 0 Then
        busyDepth = 0
        Exit Sub
    End If

    busyDepth = busyDepth - 1
    If busyDepth > 0 Then Exit Sub

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    Set appScope = Nothing
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

Public Sub SheetDeactivated(ByVal sh As Worksheet)
    If sh Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    LLEnterBusyState busyCursor:=xlNorthwestArrow
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
    LLEnterBusyState busyCursor:=xlNorthwestArrow
    Application.ScreenUpdating = False
    Service.OnSheetChange sh, target
Cleanup:
    LLExitBusyState
End Sub

Public Sub SelectionChanged(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error GoTo Cleanup
    LLEnterBusyState busyCursor:=xlNorthwestArrow
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
    LLEnterBusyState busyCursor:=xlNorthwestArrow
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
    LLEnterBusyState busyCursor:=xlNorthwestArrow
    Application.ScreenUpdating = False
    Service.UpdateFilterTables calculate
Cleanup:
    LLExitBusyState
End Sub

'@EntryPoint
Public Sub UpdateAllListAuto()
    On Error GoTo Cleanup
    LLEnterBusyState busyCursor:=xlNorthwestArrow
    Application.ScreenUpdating = False
    Service.UpdateAllListAuto
Cleanup:
    LLExitBusyState
End Sub
