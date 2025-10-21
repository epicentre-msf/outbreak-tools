Attribute VB_Name = "TestAnalysisSafeRunner"
Attribute VB_Description = "Unit tests for AnalysisSafeRunner"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests verifying AnalysisSafeRunner wraps actions with ProjectError flow")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisSafeRunner"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@TestMethod("AnalysisSafeRunner")
Public Sub TestRunExecutesAction()
    CustomTestSetTitles Assert, "AnalysisSafeRunner", "TestRunExecutesAction"

    Dim runner As AnalysisSafeRunner
    Dim actionStub As AnalysisActionStub

    Set runner = New AnalysisSafeRunner
    Set actionStub = New AnalysisActionStub

    runner.Run actionStub

    Assert.AreEqual 1&, actionStub.ExecuteCount, "Run should invoke action"
End Sub

'@TestMethod("AnalysisSafeRunner")
Public Sub TestRunRaisesProjectErrorOnFailure()
    CustomTestSetTitles Assert, "AnalysisSafeRunner", "TestRunRaisesProjectErrorOnFailure"

    Dim runner As AnalysisSafeRunner
    Dim actionStub As AnalysisActionStub
    Dim raisedError As Boolean

    Set runner = New AnalysisSafeRunner
    Set actionStub = New AnalysisActionStub
    actionStub.ConfigureError vbObjectError + 100, "boom"

    On Error Resume Next
        runner.Run actionStub, "Context"
        raisedError = (Err.Number = ProjectError.ErrorUnexpectedState)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Runner should raise ProjectError.ErrorUnexpectedState"
End Sub
