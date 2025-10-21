Attribute VB_Name = "TestAnalysisErrorHelper"
Attribute VB_Description = "Unit tests for AnalysisErrorHelper"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests verifying AnalysisErrorHelper raises expected ProjectError values")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisErrorHelper"
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

'@TestMethod("AnalysisErrorHelper")
Public Sub TestRaiseInvalidArgument()
    CustomTestSetTitles Assert, "AnalysisErrorHelper", "TestRaiseInvalidArgument"
    Dim helper As AnalysisErrorHelper
    Dim raisedError As Boolean

    Set helper = New AnalysisErrorHelper

    On Error Resume Next
        helper.RaiseInvalidArgument "plan"
        raisedError = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "RaiseInvalidArgument should raise ProjectError.InvalidArgument"
End Sub

'@TestMethod("AnalysisErrorHelper")
Public Sub TestRaiseMissingDependency()
    CustomTestSetTitles Assert, "AnalysisErrorHelper", "TestRaiseMissingDependency"
    Dim helper As AnalysisErrorHelper
    Dim raisedError As Boolean

    Set helper = New AnalysisErrorHelper

    On Error Resume Next
        helper.RaiseMissingDependency "GraphSpecsOrchestrator"
        raisedError = (Err.Number = ProjectError.ObjectNotInitialized)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "RaiseMissingDependency should raise ProjectError.ObjectNotInitialized"
End Sub

'@TestMethod("AnalysisErrorHelper")
Public Sub TestRaiseUnexpectedState()
    CustomTestSetTitles Assert, "AnalysisErrorHelper", "TestRaiseUnexpectedState"
    Dim helper As AnalysisErrorHelper
    Dim raisedError As Boolean

    Set helper = New AnalysisErrorHelper

    On Error Resume Next
        helper.RaiseUnexpectedState "Pipeline state invalid"
        raisedError = (Err.Number = ProjectError.ErrorUnexpectedState)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "RaiseUnexpectedState should raise ProjectError.ErrorUnexpectedState"
End Sub
