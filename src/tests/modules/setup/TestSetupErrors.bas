Attribute VB_Name = "TestSetupErrors"
Attribute VB_Description = "Verifies SetupErrors orchestrator initialisation"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Setup")
'@ModuleDescription("Verifies SetupErrors orchestrator initialisation")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private setupChecker As ISetupErrors

'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set setupChecker = SetupErrors.Create(ThisWorkbook)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set setupChecker = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("SetupErrors")
Public Sub TestCreateReturnsInterface()
    On Error GoTo Fail

    '@ Given
    '@ When
    '@ Then
    Assert.IsNotNothing setupChecker, "Factory should return an ISetupErrors instance"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestCreateReturnsInterface"
End Sub

'@TestMethod("SetupErrors")
Public Sub TestCheckingsInitialisedEmpty()
    On Error GoTo Fail

    Dim results As BetterArray

    Set results = setupChecker.Checkings

    Assert.IsNotNothing results, "Checkings container should be initialised during setup"
    Assert.AreEqual 0&, results.Length, "Expected empty checkings before running the workflow"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestCheckingsInitialisedEmpty"
End Sub

