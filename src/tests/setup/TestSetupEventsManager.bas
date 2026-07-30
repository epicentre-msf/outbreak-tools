Attribute VB_Name = "TestSetupEventsManager"
Attribute VB_Description = "Unit tests for the setup busy-state manager"

Option Explicit

'@Folder("CustomTests.Setup")
'@ModuleDescription("Exercises the busy-state counter, the quiet state and the crash-recovery snapshot")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

'WHAT THIS MODULE COVERS
'-------------------------------------------------------------------------------
'SetupEventsManager is a standard module because it holds two pieces of state
'that must survive between calls: the single EventSetup instance and the
'busy-state nesting counter. Ribbon callbacks and worksheet functions both reach
'them. None of that needs an Excel event to be raised, so all of it is testable
'from here.
'
'The counter, the nesting and the snapshot round trip are what the change-handler
'rework moved, so they are what this module pins.
'
'THE ONE THING THIS MODULE TOUCHES OUTSIDE ITSELF
'-------------------------------------------------------------------------------
'The manager reads and writes real Application settings and one workbook hidden
'name on ThisWorkbook. Every test here puts back what it changed, and
'TestInitialize drains the counter first, so a failed test cannot leave the next
'one entering a state that is already held.

Private Assert As CustomTest

Private Const OUTPUT_SHEET As String = "testsOutputs"
Private Const SNAPSHOT_KEY As String = "APPSTATE_SNAPSHOT"

'@ModuleInitialize
Public Sub ModuleInitialize()
    On Error GoTo Fail
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, OUTPUT_SHEET)
    Assert.SetModuleName "TestSetupEventsManager"
    Exit Sub

Fail:
    Debug.Print "TestSetupEventsManager.ModuleInitialize: "; Err.Number; Err.Description
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        DrainBusyState
        ClearSnapshotName
        If Not Assert Is Nothing Then
            Assert.PrintResults OUTPUT_SHEET
        End If
        Set Assert = Nothing
        RestoreApp
    On Error GoTo 0
End Sub

'@TestInitialize
Public Sub TestInitialize()
    On Error Resume Next
        BusyApp
        DrainBusyState
        ClearSnapshotName
        SetupEventsManager.DisposeEventSetup
    On Error GoTo 0
End Sub

'@TestCleanup
Public Sub TestCleanup()
    On Error Resume Next
        DrainBusyState
        ClearSnapshotName
        Application.EnableEvents = True
        Application.Calculation = xlCalculationManual
        If Not Assert Is Nothing Then Assert.Flush
    On Error GoTo 0
End Sub


'@section The busy-state counter
'===============================================================================

'@TestMethod("SetupEventsManager")
Public Sub TestEnterAndExitReportTheState()
    CustomTestSetTitles Assert, "SetupEventsManager", "Enter reports busy and exit reports idle"
    On Error GoTo Fail

    Assert.IsFalse SetupEventsManager.IsBusyState, "The manager should start idle"

    SetupEventsManager.EnterBusyState persist:=False
    Assert.IsTrue SetupEventsManager.IsBusyState, "The manager should report busy after entering"

    SetupEventsManager.ExitBusyState
    Assert.IsFalse SetupEventsManager.IsBusyState, "The manager should report idle after exiting"
    Exit Sub

Fail:
    DrainBusyState
    CustomTestLogFailure Assert, "TestEnterAndExitReportTheState", Err.Number, Err.Description
End Sub

'@TestMethod("SetupEventsManager")
Public Sub TestNestingRestoresOnTheOutermostExitAlone()
    CustomTestSetTitles Assert, "SetupEventsManager", "Only the outermost exit restores the state"
    On Error GoTo Fail

    'The import flow nests three deep: ImportOrCleanSetup enters, and
    'PostImportMaintenance enters twice more through the two manager routines it
    'calls. An inner exit must not put the state back under the outer job.
    SetupEventsManager.EnterBusyState persist:=False
    SetupEventsManager.EnterBusyState persist:=False
    SetupEventsManager.EnterBusyState persist:=False

    SetupEventsManager.ExitBusyState
    Assert.IsTrue SetupEventsManager.IsBusyState, "Two levels are still held after the first exit"

    SetupEventsManager.ExitBusyState
    Assert.IsTrue SetupEventsManager.IsBusyState, "One level is still held after the second exit"

    SetupEventsManager.ExitBusyState
    Assert.IsFalse SetupEventsManager.IsBusyState, "The third exit is the outermost one and releases the state"
    Exit Sub

Fail:
    DrainBusyState
    CustomTestLogFailure Assert, "TestNestingRestoresOnTheOutermostExitAlone", Err.Number, Err.Description
End Sub

'@TestMethod("SetupEventsManager")
Public Sub TestExitOnAnIdleManagerCannotDriveTheCounterNegative()
    CustomTestSetTitles Assert, "SetupEventsManager", "An unmatched exit leaves the counter at zero"
    On Error GoTo Fail

    'clickExport and clickResetTag both exit on the happy path and again in their
    'handler, so an unmatched exit is a case the workbook reaches. A counter
    'driven below zero would need as many enters to answer busy again.
    SetupEventsManager.ExitBusyState
    SetupEventsManager.ExitBusyState
    SetupEventsManager.ExitBusyState

    SetupEventsManager.EnterBusyState persist:=False
    Assert.IsTrue SetupEventsManager.IsBusyState, "One enter after three unmatched exits should report busy"

    SetupEventsManager.ExitBusyState
    Assert.IsFalse SetupEventsManager.IsBusyState, "One exit should then release it"
    Exit Sub

Fail:
    DrainBusyState
    CustomTestLogFailure Assert, "TestExitOnAnIdleManagerCannotDriveTheCounterNegative", Err.Number, Err.Description
End Sub


'@section The quiet state
'===============================================================================

'@TestMethod("SetupEventsManager")
Public Sub TestQuietStateSuppressesEventsAndPutsThemBack()
    CustomTestSetTitles Assert, "SetupEventsManager", "The quiet state suppresses events and puts them back"
    On Error GoTo Fail

    Application.EnableEvents = True

    SetupEventsManager.EnterQuietState

    Assert.IsFalse Application.EnableEvents, _
        "The watcher writes a flag into the hidden registry, so events must be off while it runs"
    Assert.IsTrue SetupEventsManager.IsQuietState, "The manager should report the quiet state"

    SetupEventsManager.ExitBusyState

    Assert.IsTrue Application.EnableEvents, "Events should be back on after the quiet state ends"
    Assert.IsFalse SetupEventsManager.IsQuietState, "The quiet state should be released"
    Exit Sub

Fail:
    DrainBusyState
    Application.EnableEvents = True
    CustomTestLogFailure Assert, "TestQuietStateSuppressesEventsAndPutsThemBack", Err.Number, Err.Description
End Sub

'@TestMethod("SetupEventsManager")
Public Sub TestQuietStateLeavesCalculationAlone()
    CustomTestSetTitles Assert, "SetupEventsManager", "The quiet state leaves calculation alone"
    On Error GoTo Fail

    'This is the flicker fix. Recording a watcher flag needs no manual
    'calculation, and it is the flip back to automatic at the end of every
    'committed edit that triggers the recalculation pass the user sees.
    Application.Calculation = xlCalculationAutomatic

    SetupEventsManager.EnterQuietState

    Assert.AreEqual CLng(xlCalculationAutomatic), CLng(Application.Calculation), _
        "The quiet state should not switch calculation to manual"

    SetupEventsManager.ExitBusyState
    Application.Calculation = xlCalculationManual
    Exit Sub

Fail:
    DrainBusyState
    Application.Calculation = xlCalculationManual
    CustomTestLogFailure Assert, "TestQuietStateLeavesCalculationAlone", Err.Number, Err.Description
End Sub

'@TestMethod("SetupEventsManager")
Public Sub TestBusyStateSwitchesCalculationToManual()
    CustomTestSetTitles Assert, "SetupEventsManager", "The full state switches calculation to manual"
    On Error GoTo Fail

    Application.Calculation = xlCalculationAutomatic

    SetupEventsManager.EnterBusyState persist:=False

    Assert.AreEqual CLng(xlCalculationManual), CLng(Application.Calculation), _
        "The Analysis branch clears cells and rewrites validation, so it keeps the full lockdown"

    SetupEventsManager.ExitBusyState

    Assert.AreEqual CLng(xlCalculationAutomatic), CLng(Application.Calculation), _
        "The outermost exit should put the calculation mode back"

    Application.Calculation = xlCalculationManual
    Exit Sub

Fail:
    DrainBusyState
    Application.Calculation = xlCalculationManual
    CustomTestLogFailure Assert, "TestBusyStateSwitchesCalculationToManual", Err.Number, Err.Description
End Sub

'@TestMethod("SetupEventsManager")
Public Sub TestQuietStateSharesTheCounterWithTheFullState()
    CustomTestSetTitles Assert, "SetupEventsManager", "Both states share one nesting counter"
    On Error GoTo Fail

    'One counter is what makes the two states safe to mix. A quiet outer state
    'with a full inner one must still take two exits.
    SetupEventsManager.EnterQuietState
    SetupEventsManager.EnterBusyState persist:=False

    SetupEventsManager.ExitBusyState
    Assert.IsTrue SetupEventsManager.IsBusyState, "The outer quiet state is still held"

    SetupEventsManager.ExitBusyState
    Assert.IsFalse SetupEventsManager.IsBusyState, "The second exit releases it"
    Exit Sub

Fail:
    DrainBusyState
    Application.EnableEvents = True
    CustomTestLogFailure Assert, "TestQuietStateSharesTheCounterWithTheFullState", Err.Number, Err.Description
End Sub


'@section The crash-recovery snapshot
'===============================================================================

'@TestMethod("SetupEventsManager")
Public Sub TestPersistedSnapshotIsWrittenThenCleared()
    CustomTestSetTitles Assert, "SetupEventsManager", "The snapshot is written on entry and cleared on exit"
    On Error GoTo Fail

    'A VBA state reset in the middle of a long job leaves the Application
    'locked down. The snapshot is what the next entry reads to undo that.
    SetupEventsManager.EnterBusyState persist:=True

    Assert.IsTrue SnapshotExists(), "Entering with persistence on should write the snapshot"

    SetupEventsManager.ExitBusyState

    Assert.IsFalse SnapshotExists(), "A clean exit should clear the snapshot"
    Exit Sub

Fail:
    DrainBusyState
    ClearSnapshotName
    CustomTestLogFailure Assert, "TestPersistedSnapshotIsWrittenThenCleared", Err.Number, Err.Description
End Sub

'@TestMethod("SetupEventsManager")
Public Sub TestSnapshotCarriesSixValuesInTheRestoreOrder()
    CustomTestSetTitles Assert, "SetupEventsManager", "The snapshot carries six values"
    On Error GoTo Fail

    'RecoverIfNeeded reads the six back by position and stops when there are
    'fewer than six, so the count is the contract between the two routines.
    SetupEventsManager.EnterBusyState persist:=True

    Dim parts() As String
    parts = Split(SnapshotValue(), "|")

    Assert.AreEqual CLng(5), CLng(UBound(parts)), _
        "RecoverIfNeeded reads six pipe-delimited values, so the writer must produce six"

    SetupEventsManager.ExitBusyState
    Exit Sub

Fail:
    DrainBusyState
    ClearSnapshotName
    CustomTestLogFailure Assert, "TestSnapshotCarriesSixValuesInTheRestoreOrder", Err.Number, Err.Description
End Sub

'@TestMethod("SetupEventsManager")
Public Sub TestQuietStateWritesNoSnapshot()
    CustomTestSetTitles Assert, "SetupEventsManager", "The quiet state writes no snapshot"
    On Error GoTo Fail

    SetupEventsManager.EnterQuietState

    Assert.IsFalse SnapshotExists(), _
        "A handler that only records a flag has no locked-down state to recover"

    SetupEventsManager.ExitBusyState
    Exit Sub

Fail:
    DrainBusyState
    ClearSnapshotName
    Application.EnableEvents = True
    CustomTestLogFailure Assert, "TestQuietStateWritesNoSnapshot", Err.Number, Err.Description
End Sub


'@section The service instance
'===============================================================================

'@TestMethod("SetupEventsManager")
Public Sub TestServiceIsBuiltOnceAndHandedBack()
    CustomTestSetTitles Assert, "SetupEventsManager", "One service instance is handed to every caller"
    On Error GoTo Fail

    'The four worksheet functions, the ribbon and the event handlers all reach
    'the same instance, which is what makes the lazy caches worth having.
    Dim firstCall As EventSetup
    Dim secondCall As EventSetup

    Set firstCall = SetupEventsManager.EventSetupService
    Set secondCall = SetupEventsManager.EventSetupService

    Assert.IsTrue (firstCall Is secondCall), "Two calls should answer the same instance"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestServiceIsBuiltOnceAndHandedBack", Err.Number, Err.Description
End Sub

'@TestMethod("SetupEventsManager")
Public Sub TestDisposeReleasesTheServiceSoTheNextCallRebuilds()
    CustomTestSetTitles Assert, "SetupEventsManager", "Dispose releases the service"
    On Error GoTo Fail

    'Workbook_BeforeClose calls this. It had no caller at all, so the setup
    'workbook held its service for the whole Excel session.
    Dim firstCall As EventSetup
    Dim afterDispose As EventSetup

    Set firstCall = SetupEventsManager.EventSetupService
    SetupEventsManager.DisposeEventSetup
    Set afterDispose = SetupEventsManager.EventSetupService

    Assert.IsFalse (firstCall Is afterDispose), "The call after Dispose should build a new instance"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDisposeReleasesTheServiceSoTheNextCallRebuilds", Err.Number, Err.Description
End Sub

'@TestMethod("SetupEventsManager")
Public Sub TestResetCachesIsSafeBeforeAnyServiceExists()
    CustomTestSetTitles Assert, "SetupEventsManager", "Resetting caches with no service is safe"
    On Error GoTo Fail

    'PostImportMaintenance and the clean flow both call this, and either can run
    'before anything has asked for the service.
    SetupEventsManager.DisposeEventSetup
    SetupEventsManager.ResetEventSetupCaches

    Assert.IsTrue True, "Resetting the caches with no service held should not raise"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestResetCachesIsSafeBeforeAnyServiceExists", Err.Number, Err.Description
End Sub


'@section Helpers
'===============================================================================

'@sub-title Bring the manager back to idle whatever a failed test left behind
Private Sub DrainBusyState()
    Dim guard As Long

    On Error Resume Next
    Do While SetupEventsManager.IsBusyState
        SetupEventsManager.ExitBusyState
        guard = guard + 1
        If guard > 20 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

'@sub-title Whether the crash-recovery snapshot is on the driver workbook
Private Function SnapshotExists() As Boolean
    Dim names As HiddenNames

    On Error Resume Next
        Set names = HiddenNames.Create(ThisWorkbook)
        If Not names Is Nothing Then SnapshotExists = names.HasName(SNAPSHOT_KEY)
    On Error GoTo 0
End Function

'@sub-title The raw snapshot string, or empty when there is none
Private Function SnapshotValue() As String
    Dim names As HiddenNames

    On Error Resume Next
        Set names = HiddenNames.Create(ThisWorkbook)
        If Not names Is Nothing Then SnapshotValue = names.ValueAsString(SNAPSHOT_KEY)
    On Error GoTo 0
End Function

'@sub-title Remove the snapshot so no test reads one another test wrote
Private Sub ClearSnapshotName()
    Dim names As HiddenNames

    On Error Resume Next
        Set names = HiddenNames.Create(ThisWorkbook)
        If Not names Is Nothing Then
            If names.HasName(SNAPSHOT_KEY) Then names.RemoveName SNAPSHOT_KEY
        End If
    On Error GoTo 0
End Sub
