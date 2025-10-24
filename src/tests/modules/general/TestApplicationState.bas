Attribute VB_Name = "TestApplicationState"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

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
Private Sub CaptureInitialState()
    initialScreenUpdating = Application.ScreenUpdating
    initialDisplayAlerts = Application.DisplayAlerts
    initialEnableEvents = Application.EnableEvents
    initialCalculation = Application.Calculation
    initialCalculateBeforeSave = Application.CalculateBeforeSave
    animationsAvailable = TryReadAnimations(initialEnableAnimations)
End Sub

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
