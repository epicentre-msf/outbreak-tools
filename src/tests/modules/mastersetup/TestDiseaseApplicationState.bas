Attribute VB_Name = "TestDiseaseApplicationState"
Attribute VB_Description = "Tests ensuring DiseaseApplicationState restores Excel state after guarded operations"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests ensuring DiseaseApplicationState restores Excel state after guarded operations")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest
Private Guard As IDiseaseApplicationState

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseApplicationState"
    Set Guard = New DiseaseApplicationState
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        If Not Guard Is Nothing Then
            Guard.Restore
        End If
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
    Set Guard = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Guard Is Nothing Then Guard.Restore
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseApplicationState")
Public Sub TestGuardDisablesAndRestoresState()
    CustomTestSetTitles Assert, "DiseaseApplicationState", "TestGuardDisablesAndRestoresState"

    Dim app As Application
    Dim originalScreenUpdating As Boolean
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalCalculation As XlCalculation

    On Error GoTo Fail

    Set app = Application
    originalScreenUpdating = app.ScreenUpdating
    originalDisplayAlerts = app.DisplayAlerts
    originalEnableEvents = app.EnableEvents
    originalCalculation = app.Calculation

    Guard.BeginGuard app

    Assert.IsFalse app.ScreenUpdating, "ScreenUpdating should be disabled while guarded"
    Assert.IsFalse app.DisplayAlerts, "DisplayAlerts should be disabled while guarded"
    Assert.IsFalse app.EnableEvents, "EnableEvents should be disabled while guarded"
    Assert.AreEqual xlCalculationManual, app.Calculation, "Calculation should be manual while guarded"

    Guard.Restore

    Assert.AreEqual originalScreenUpdating, app.ScreenUpdating, "ScreenUpdating should be restored"
    Assert.AreEqual originalDisplayAlerts, app.DisplayAlerts, "DisplayAlerts should be restored"
    Assert.AreEqual originalEnableEvents, app.EnableEvents, "EnableEvents should be restored"
    Assert.AreEqual originalCalculation, app.Calculation, "Calculation should be restored"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestGuardDisablesAndRestoresState", Err.Number, Err.Description
    Guard.Restore
End Sub

'@TestMethod("DiseaseApplicationState")
Public Sub TestRestoreIsIdempotent()
    CustomTestSetTitles Assert, "DiseaseApplicationState", "TestRestoreIsIdempotent"

    Dim app As Application
    Dim originalDisplayAlerts As Boolean
    Dim raisedError As Boolean

    On Error GoTo Fail

    Set app = Application
    originalDisplayAlerts = app.DisplayAlerts

    Guard.BeginGuard app, disableStatusBar:=False
    Guard.Restore

    On Error Resume Next
        Guard.Restore
        raisedError = (Err.Number <> 0)
        Err.Clear
    On Error GoTo Fail

    Assert.IsFalse raisedError, "Restore should be safe to call multiple times"
    Assert.AreEqual originalDisplayAlerts, app.DisplayAlerts, "DisplayAlerts should remain restored"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRestoreIsIdempotent", Err.Number, Err.Description
    Guard.Restore
End Sub
