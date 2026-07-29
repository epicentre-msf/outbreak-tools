Attribute VB_Name = "TestApplicationState"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the ApplicationState class")

'@description
'Validates the ApplicationState class, which wraps Excel Application-level
'settings (ScreenUpdating, DisplayAlerts, Calculation, EnableEvents,
'CalculateBeforeSave, EnableAnimations, AutomationSecurity, cursor) in an
'RAII-style scope object.
'Tests confirm that ApplyBusyState switches each property to its expected
'performance mode, Restore returns all properties to their captured
'snapshot, RefreshSnapshot guards against misuse while busy, and the
'optional suppressEvents / calculateOnSave / busyCursor / blockSecurity
'overrides behave correctly.
'Each test sets a known starting state before it builds the scope, so an
'assertion can never compare a value against itself. TestInitialize and
'TestCleanup put the pre-test environment back to prevent cross-test
'interference.
'@depends ApplicationState, CustomTest, TestHelpersLite

Private Assert As CustomTest
Private initialScreenUpdating As Boolean
Private initialDisplayAlerts As Boolean
Private initialEnableEvents As Boolean
Private initialCalculation As XlCalculation
Private initialCalculateBeforeSave As Boolean
Private initialEnableAnimations As Boolean
Private initialAutomationSecurity As MsoAutomationSecurity
Private initialCursor As XlMousePointer
Private animationsAvailable As Boolean


'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    'The snapshot MUST be taken before BusyApp, otherwise every initial*
    'field holds the busy value and every restore assertion compares the
    'busy state against the busy state.
    CaptureInitialState
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestApplicationState"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
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
    initialAutomationSecurity = Application.AutomationSecurity
    initialCursor = Application.Cursor
    animationsAvailable = TryReadAnimations(initialEnableAnimations)
End Sub

'@sub-title Restore every Application property to its pre-test value.
Private Sub ResetApplicationState()
    Application.ScreenUpdating = initialScreenUpdating
    Application.DisplayAlerts = initialDisplayAlerts
    Application.EnableEvents = initialEnableEvents
    Application.Calculation = initialCalculation
    Application.CalculateBeforeSave = initialCalculateBeforeSave
    Application.AutomationSecurity = initialAutomationSecurity
    Application.Cursor = initialCursor
    If animationsAvailable Then
        On Error Resume Next
            Application.EnableAnimations = initialEnableAnimations
        On Error GoTo 0
    End If
End Sub

'@sub-title Put the Application into a known idle state before a test builds a scope.
'@details
'Every flip test needs a starting value that differs from the busy value,
'otherwise the assertion after ApplyBusyState passes even when the class
'does nothing at all.
Private Sub SetIdleApplicationState()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateBeforeSave = False
    If animationsAvailable Then
        On Error Resume Next
            Application.EnableAnimations = True
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
'Puts the Application in a known idle state first (screen on, alerts on,
'automatic calculation, CalculateBeforeSave off), then creates a scope and
'calls ApplyBusyState with default parameters. Each assertion therefore
'checks a real flip.
'@TestMethod("ApplicationState")
Public Sub TestApplyBusyStateSwitchesSettings()
    CustomTestSetTitles Assert, "ApplicationState", "ApplyBusyStateSwitchesSettings"

    Dim scope As ApplicationState

    SetIdleApplicationState
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState

    Assert.IsFalse Application.ScreenUpdating, "ApplyBusyState must disable screen updating"
    Assert.IsFalse Application.DisplayAlerts, "ApplyBusyState must disable alerts"
    Assert.AreEqual xlCalculationManual, Application.Calculation, _
                     "ApplyBusyState must set calculation to manual"
    Assert.IsTrue Application.EnableEvents, _
                  "Default ApplyBusyState should leave events unchanged"
    Assert.IsTrue Application.CalculateBeforeSave, _
                  "Default ApplyBusyState should enable CalculateBeforeSave"

    If animationsAvailable Then
        Assert.IsFalse Application.EnableAnimations, "ApplyBusyState must disable animations when supported"
    End If

    scope.Restore
End Sub

'@sub-title Verify Restore returns every setting to its captured snapshot.
'@details
'Sets a known idle state, creates a scope over it, applies busy state to
'mutate all settings, then calls Restore. Each Application property is
'compared against the idle value the scope captured, so an empty Restore
'would fail every assertion.
'@TestMethod("ApplicationState")
Public Sub TestRestoreReturnsOriginalSettings()
    CustomTestSetTitles Assert, "ApplicationState", "RestoreReturnsOriginalSettings"

    Dim scope As ApplicationState

    SetIdleApplicationState
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState suppressEvents:=True

    scope.Restore

    Assert.IsTrue Application.ScreenUpdating, _
                  "Restore must reapply the original ScreenUpdating value"
    Assert.IsTrue Application.DisplayAlerts, _
                  "Restore must reapply the original DisplayAlerts value"
    Assert.IsTrue Application.EnableEvents, _
                  "Restore must reapply the original EnableEvents value"
    Assert.AreEqual xlCalculationAutomatic, Application.Calculation, _
                     "Restore must reapply the original calculation mode"
    Assert.IsFalse Application.CalculateBeforeSave, _
                   "Restore must reapply the original CalculateBeforeSave flag"

    If animationsAvailable Then
        Assert.IsTrue Application.EnableAnimations, _
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

    Dim scope As ApplicationState

    SetIdleApplicationState
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState

    On Error GoTo ExpectError
        scope.RefreshSnapshot
    On Error GoTo 0

    'The handler must be off before the failure is logged, otherwise a raise
    'inside the harness lands on the label below and is read as a pass.
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
    CustomTestSetTitles Assert, "ApplicationState", "ApplyBusyStateSuppressEventsWhenRequested"

    Dim scope As ApplicationState

    SetIdleApplicationState
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState suppressEvents:=True

    Assert.IsFalse Application.EnableEvents, "ApplyBusyState suppressEvents:=True must disable events"

    scope.Restore
    Assert.IsTrue Application.EnableEvents, _
                  "Restore must bring back original EnableEvents value"
End Sub

'@sub-title Verify calculateOnSave parameter disables CalculateBeforeSave.
'@details
'The default busy state enables CalculateBeforeSave. Passing
'calculateOnSave:=False should leave it False. The starting value is set to
'True first so the assertion checks a real flip.
'@TestMethod("ApplicationState")
Public Sub TestApplyBusyStateRespectsCalculateOnSaveParameter()
    CustomTestSetTitles Assert, "ApplicationState", "ApplyBusyStateRespectsCalculateOnSaveParameter"

    Dim scope As ApplicationState

    SetIdleApplicationState
    Application.CalculateBeforeSave = True
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState calculateOnSave:=False

    Assert.IsFalse Application.CalculateBeforeSave, _
                  "ApplyBusyState calculateOnSave:=False must disable CalculateBeforeSave"

    scope.Restore
    Assert.IsTrue Application.CalculateBeforeSave, _
                  "Restore must reapply the CalculateBeforeSave value held at creation"
End Sub

'@sub-title Verify blockSecurity forces automation security off and Restore puts it back.
'@details
'blockSecurity has a single caller in the whole tree
'(SetupImport), so nothing else would surface a break here.
'@TestMethod("ApplicationState")
Public Sub TestApplyBusyStateBlocksAutomationSecurity()
    CustomTestSetTitles Assert, "ApplicationState", "ApplyBusyStateBlocksAutomationSecurity"

    Dim scope As ApplicationState
    Dim securityBefore As MsoAutomationSecurity

    SetIdleApplicationState
    Application.AutomationSecurity = msoAutomationSecurityByUI
    securityBefore = Application.AutomationSecurity
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState blockSecurity:=True

    Assert.AreEqual msoAutomationSecurityForceDisable, Application.AutomationSecurity, _
                     "blockSecurity:=True must force automation security to disabled"

    scope.Restore

    Assert.AreEqual securityBefore, Application.AutomationSecurity, _
                     "Restore must bring back the original automation security level"
End Sub

'@sub-title Verify the default busy state leaves automation security alone.
'@TestMethod("ApplicationState")
Public Sub TestApplyBusyStateKeepsSecurityByDefault()
    CustomTestSetTitles Assert, "ApplicationState", "ApplyBusyStateKeepsSecurityByDefault"

    Dim scope As ApplicationState
    Dim securityBefore As MsoAutomationSecurity

    SetIdleApplicationState
    Application.AutomationSecurity = msoAutomationSecurityByUI
    securityBefore = Application.AutomationSecurity
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState

    Assert.AreEqual securityBefore, Application.AutomationSecurity, _
                     "Default ApplyBusyState must not touch automation security"

    scope.Restore
End Sub

'@sub-title Verify busyCursor sets the cursor while busy and Restore puts it back.
'@TestMethod("ApplicationState")
Public Sub TestApplyBusyStateSetsBusyCursor()
    CustomTestSetTitles Assert, "ApplicationState", "ApplyBusyStateSetsBusyCursor"

    Dim scope As ApplicationState

    SetIdleApplicationState
    Application.Cursor = xlDefault
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState busyCursor:=xlWait

    Assert.AreEqual xlWait, Application.Cursor, _
                     "busyCursor must be applied while the scope is busy"

    scope.Restore

    Assert.AreEqual xlDefault, Application.Cursor, _
                     "Restore must bring back the cursor held at creation"
End Sub

'@sub-title Verify the default busy state leaves the cursor alone.
'@TestMethod("ApplicationState")
Public Sub TestApplyBusyStateKeepsCursorByDefault()
    CustomTestSetTitles Assert, "ApplicationState", "ApplyBusyStateKeepsCursorByDefault"

    Dim scope As ApplicationState

    SetIdleApplicationState
    Application.Cursor = xlDefault
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState

    Assert.AreEqual xlDefault, Application.Cursor, _
                     "Default ApplyBusyState must leave the cursor unchanged"

    scope.Restore
End Sub

'@sub-title Verify IsBusy follows the busy state through a full cycle.
'@TestMethod("ApplicationState")
Public Sub TestIsBusyFollowsTheCycle()
    CustomTestSetTitles Assert, "ApplicationState", "IsBusyFollowsTheCycle"

    Dim scope As ApplicationState

    SetIdleApplicationState
    Set scope = ApplicationState.Create(Application)

    Assert.IsFalse scope.IsBusy, "A fresh scope must not report itself busy"

    scope.ApplyBusyState
    Assert.IsTrue scope.IsBusy, "ApplyBusyState must mark the scope busy"

    scope.Restore
    Assert.IsFalse scope.IsBusy, "Restore must clear the busy flag"
End Sub

'@sub-title Verify Create captures a snapshot straight away.
'@details
'LinelistSpecs asks HasSnapshot before it decides whether to build a new
'scope, so a False answer after Create would rebuild the object on every
'busy block.
'@TestMethod("ApplicationState")
Public Sub TestCreateCapturesSnapshot()
    CustomTestSetTitles Assert, "ApplicationState", "CreateCapturesSnapshot"

    Dim scope As ApplicationState

    SetIdleApplicationState
    Set scope = ApplicationState.Create(Application)

    Assert.IsTrue scope.HasSnapshot, "Create must capture the snapshot itself"

    scope.ApplyBusyState
    Assert.IsTrue scope.HasSnapshot, "The snapshot must survive ApplyBusyState"

    scope.Restore
    Assert.IsTrue scope.HasSnapshot, "The snapshot must survive Restore"
End Sub

'@sub-title Verify a second ApplyBusyState on a busy scope does nothing.
'@details
'A setting changed by hand between two ApplyBusyState calls must be left
'alone by the second call, because the scope is already busy.
'@TestMethod("ApplicationState")
Public Sub TestSecondApplyBusyStateIsIgnored()
    CustomTestSetTitles Assert, "ApplicationState", "SecondApplyBusyStateIsIgnored"

    Dim scope As ApplicationState

    SetIdleApplicationState
    Set scope = ApplicationState.Create(Application)

    scope.ApplyBusyState calculateOnSave:=False
    Application.CalculateBeforeSave = True

    scope.ApplyBusyState calculateOnSave:=False

    Assert.IsTrue Application.CalculateBeforeSave, _
                  "A second ApplyBusyState on a busy scope must change nothing"

    scope.Restore
End Sub

'@sub-title Verify Create without an argument guards the host application.
'@details
'Most call sites in the tree use ApplicationState.Create() with no
'argument, so the lazy fallback to the host Application has to work.
'@TestMethod("ApplicationState")
Public Sub TestCreateWithoutArgumentUsesHostApplication()
    CustomTestSetTitles Assert, "ApplicationState", "CreateWithoutArgumentUsesHostApplication"

    Dim scope As ApplicationState

    SetIdleApplicationState
    Set scope = ApplicationState.Create()

    Assert.IsTrue scope.HasSnapshot, "Create with no argument must capture a snapshot"

    scope.ApplyBusyState

    Assert.IsFalse Application.ScreenUpdating, _
                   "Create with no argument must guard the host Application"

    scope.Restore

    Assert.IsTrue Application.ScreenUpdating, _
                  "Restore must bring back the host ScreenUpdating value"
End Sub

'@sub-title Verify BindApplication refuses an object that is not an Application.
'@details
'The reference is passed through an Object variable so the check happens at
'run time. Whatever raises first, the object must not silently accept a
'worksheet as its guarded application.
'@TestMethod("ApplicationState")
Public Sub TestBindApplicationRejectsWrongType()
    CustomTestSetTitles Assert, "ApplicationState", "BindApplicationRejectsWrongType"

    Dim scope As ApplicationState
    Dim wrongType As Object

    SetIdleApplicationState
    Set scope = ApplicationState.Create(Application)
    Set wrongType = ThisWorkbook.Worksheets(1)

    On Error GoTo ExpectError
        scope.BindApplication wrongType
    On Error GoTo 0

    Assert.LogFailure "BindApplication should refuse an object that is not an Application"
    Exit Sub

ExpectError:
    Assert.AreNotEqual 0, Err.Number, _
                       "BindApplication must raise when handed a worksheet"
    Err.Clear
End Sub
