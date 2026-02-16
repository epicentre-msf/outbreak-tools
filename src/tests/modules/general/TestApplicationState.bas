Attribute VB_Name = "TestApplicationState"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the ApplicationState class")

'@description
'Validates the ApplicationState class, which wraps Excel Application-level
'settings (ScreenUpdating, DisplayAlerts, Calculation, EnableEvents,
'CalculateBeforeSave, EnableAnimations) in an RAII-style scope object.
'Tests confirm that ApplyBusyState switches each property to its expected
'performance mode, Restore returns all properties to their captured
'snapshot, RefreshSnapshot guards against misuse while busy, and the
'optional suppressEvents / calculateOnSave overrides behave correctly.
'Each test creates a fresh ApplicationState scope and restores the
'original environment in TestInitialize / TestCleanup to prevent
'cross-test interference.
'@depends ApplicationState, IApplicationState, CustomTest, TestHelpers

Private Assert As ICustomTest
Private initialScreenUpdating As Boolean
Private initialDisplayAlerts As Boolean
Private initialEnableEvents As Boolean
Private initialCalculation As XlCalculation
Private initialCalculateBeforeSave As Boolean
Private initialEnableAnimations As Boolean
Private animationsAvailable As Boolean


'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestApplicationState"
    CaptureInitialState
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    ResetApplicationState
    Set Assert = Nothing
    RestoreApp
End Sub

'@TestInitialize
Private Sub TestInitialize()
    ResetApplicationState
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    ResetApplicationState
End Sub


'@section Helper routines
'===============================================================================

'@sub-title Snapshot the current Application settings before any test runs.
Private Sub CaptureInitialState()
    initialScreenUpdating = Application.ScreenUpdating
    initialDisplayAlerts = Application.DisplayAlerts
    initialEnableEvents = Application.EnableEvents
    initialCalculation = Application.Calculation
    initialCalculateBeforeSave = Application.CalculateBeforeSave
    animationsAvailable = TryReadAnimations(initialEnableAnimations)
End Sub

'@sub-title Restore every Application property to its pre-test value.
Private Sub ResetApplicationState()
    Application.ScreenUpdating = initialScreenUpdating
    Application.DisplayAlerts = initialDisplayAlerts
    Application.EnableEvents = initialEnableEvents
    Application.Calculation = initialCalculation
    Application.CalculateBeforeSave = initialCalculateBeforeSave
    If animationsAvailable Then
        On Error Resume Next
            Application.EnableAnimations = initialEnableAnimations
        On Error GoTo 0
    End If
End Sub

'@sub-title Probe whether EnableAnimations is available on this host.
'@details
'Some Excel versions or hosts do not expose EnableAnimations. This helper
'attempts to read the property; on success the captured value and True are
'returned via ByRef. On failure the value defaults to False so that
'animation-related assertions are skipped gracefully.
Private Function TryReadAnimations(ByRef value As Boolean) As Boolean
    On Error GoTo MissingProperty
        value = Application.EnableAnimations
        TryReadAnimations = True
    On Error GoTo 0
    Exit Function
MissingProperty:
    value = False
    TryReadAnimations = False
    Err.Clear
End Function


'@section Test cases
'===============================================================================

'@sub-title Verify ApplyBusyState switches all settings to performance mode.
'@details
'Creates a fresh scope, calls ApplyBusyState with default parameters, then
'asserts ScreenUpdating=False, DisplayAlerts=False, Calculation=xlManual,
'EnableEvents unchanged, and CalculateBeforeSave=True. When animations are
'available, also checks EnableAnimations=False. Restores after assertions.
'@TestMethod("ApplicationState")
Public Sub TestApplyBusyStateSwitchesSettings()
    CustomTestSetTitles Assert, "ApplicationState", "ApplyBusyStateSwitchesSettings"

    Dim scope As IApplicationState
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState

    Assert.IsFalse Application.ScreenUpdating, "ApplyBusyState must disable screen updating"
    Assert.IsFalse Application.DisplayAlerts, "ApplyBusyState must disable alerts"
    Assert.AreEqual xlCalculationManual, Application.Calculation, _
                     "ApplyBusyState must set calculation to manual"
    Assert.AreEqual initialEnableEvents, Application.EnableEvents, _
                     "Default ApplyBusyState should leave events unchanged"
    Assert.IsTrue Application.CalculateBeforeSave, _
                  "Default ApplyBusyState should leave CalculateBeforeSave enabled"

    If animationsAvailable Then
        Assert.IsFalse Application.EnableAnimations, "ApplyBusyState must disable animations when supported"
    End If

    scope.Restore
End Sub

'@sub-title Verify Restore returns every setting to its captured snapshot.
'@details
'Creates a scope, applies busy state to mutate all settings, then calls
'Restore. Each Application property is compared against the snapshot values
'captured in CaptureInitialState to confirm full restoration.
'@TestMethod("ApplicationState")
Public Sub TestRestoreReturnsOriginalSettings()
    CustomTestSetTitles Assert, "ApplicationState", "RestoreReturnsOriginalSettings"

    Dim scope As IApplicationState
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState

    scope.Restore

    Assert.AreEqual initialScreenUpdating, Application.ScreenUpdating, _
                     "Restore must reapply the original ScreenUpdating value"
    Assert.AreEqual initialDisplayAlerts, Application.DisplayAlerts, _
                     "Restore must reapply the original DisplayAlerts value"
    Assert.AreEqual initialEnableEvents, Application.EnableEvents, _
                     "Restore must reapply the original EnableEvents value"
    Assert.AreEqual initialCalculation, Application.Calculation, _
                     "Restore must reapply the original calculation mode"
    Assert.AreEqual initialCalculateBeforeSave, Application.CalculateBeforeSave, _
                     "Restore must reapply the original CalculateBeforeSave flag"

    If animationsAvailable Then
        Assert.AreEqual initialEnableAnimations, Application.EnableAnimations, _
                         "Restore must reapply the original animation preference"
    End If
End Sub

'@sub-title Verify RefreshSnapshot raises when called while busy.
'@details
'ApplyBusyState puts the scope into the "busy" state. Calling
'RefreshSnapshot in that state is a programming error, so the class must
'raise ErrorUnexpectedState. This test confirms the error number matches
'ProjectError.ErrorUnexpectedState.
'@TestMethod("ApplicationState")
Public Sub TestRefreshSnapshotRequiresIdle()
    CustomTestSetTitles Assert, "ApplicationState", "RefreshSnapshotRequiresIdle"

    Dim scope As IApplicationState
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState

    On Error GoTo ExpectError
        scope.RefreshSnapshot
        Assert.LogFailure "RefreshSnapshot should raise when called while busy"
        scope.Restore
        Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ErrorUnexpectedState, Err.Number, _
                     "RefreshSnapshot should raise ErrorUnexpectedState while busy"
    Err.Clear
    scope.Restore
End Sub

'@sub-title Verify suppressEvents parameter disables EnableEvents.
'@details
'By default, ApplyBusyState does not touch EnableEvents. Passing
'suppressEvents:=True must set EnableEvents to False. After Restore the
'original value must be reinstated.
'@TestMethod("ApplicationState")
Public Sub TestApplyBusyStateSuppressEventsWhenRequested()
    CustomTestSetTitles Assert, "ApplicationState", "TestApplyBusyStateSuppressEventsWhenRequested"

    Dim scope As IApplicationState
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState suppressEvents:=True

    Assert.IsFalse Application.EnableEvents, "ApplyBusyState suppressEvents:=True must disable events"

    scope.Restore
    Assert.AreEqual initialEnableEvents, Application.EnableEvents, _
                     "Restore must bring back original EnableEvents value"
End Sub

'@sub-title Verify calculateOnSave parameter disables CalculateBeforeSave.
'@details
'The default busy state leaves CalculateBeforeSave enabled. Passing
'calculateOnSave:=False should flip it to False. After Restore the
'original value is confirmed.
'@TestMethod("ApplicationState")
Public Sub TestApplyBusyStateRespectsCalculateOnSaveParameter()
    CustomTestSetTitles Assert, "ApplicationState", "TestApplyBusyStateRespectsCalculateOnSaveParameter"

    Dim scope As IApplicationState
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState calculateOnSave:=False

    Assert.IsFalse Application.CalculateBeforeSave, _
                  "ApplyBusyState calculateOnSave:=False must disable CalculateBeforeSave"

    scope.Restore
    Assert.AreEqual initialCalculateBeforeSave, Application.CalculateBeforeSave, _
                     "Restore must reapply initial CalculateBeforeSave value"
End Sub
