Attribute VB_Name = "TestListWorksheetPreparer"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Unit tests for the default ListWorksheetPreparer implementation")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private Preparer As ListWorksheetPreparer
Private Context As ListBuildContextStub

'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@section Test lifecycle
'===============================================================================
'@TestInitialize
Private Sub TestInitialize()
    Set Preparer = New ListWorksheetPreparer
    Set Context = New ListBuildContextStub
    Context.Configure "Worksheet_Main", CByte(2)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Preparer = Nothing
    Set Context = Nothing
End Sub

'@section Tests
'===============================================================================
'@TestMethod("ListWorksheetPreparer")
Private Sub TestBeginAndCompleteToggleBusyState()
    Preparer.Begin Context
    Assert.IsTrue Preparer.IsBusy, "Preparer should be busy after Begin"
    Assert.AreEqual 1&, Preparer.BeginCount, "Begin count should increment"

    Preparer.Complete Context
    Assert.IsFalse Preparer.IsBusy, "Preparer should not be busy after Complete"
    Assert.AreEqual 1&, Preparer.CompleteCount, "Complete count should increment"
End Sub

'@TestMethod("ListWorksheetPreparer")
Private Sub TestBeginFailsWhenAlreadyBusy()
    On Error GoTo ExpectError

    Preparer.Begin Context
    Preparer.Begin Context
    Assert.Fail "Second Begin should raise"
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), Err.Number
    Err.Clear
End Sub

'@TestMethod("ListWorksheetPreparer")
Private Sub TestAbortClearsBusyState()
    Preparer.Begin Context
    Preparer.Abort Context, CLng(1001), "failure"

    Assert.IsFalse Preparer.IsBusy, "Abort should reset busy flag"
    Assert.AreEqual 1&, Preparer.AbortCount, "Abort count should increment"
End Sub

