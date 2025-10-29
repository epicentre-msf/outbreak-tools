Attribute VB_Name = "TestAnalysisTableWriterContext"
Attribute VB_Description = "Unit tests for AnalysisTableWriterContext baseline behaviour"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests for AnalysisTableWriterContext baseline behaviour")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest
Private DummySheet As Worksheet

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisTableWriterContext"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
    Set DummySheet = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set DummySheet = EnsureWorksheet("WriterContextDummy")
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    DeleteWorksheet "WriterContextDummy"
    Set DummySheet = Nothing
End Sub

Private Function BuildPlanItemStub() As IAnalysisTablePlanItem
    Dim policyStub As AnalysisTablePolicyResultStub
    policyStub.Configure Nothing, True
    Set BuildPlanItemStub = AnalysisTablePlanItem.Create(policyStub, 0)
End Function

'@TestMethod("AnalysisTableWriterContext")
Public Sub TestBeginPlanResetsCollections()
    CustomTestSetTitles Assert, "AnalysisTableWriterContext", "TestBeginPlanResetsCollections"
    Dim context As AnalysisTableWriterContext

    Set context = New AnalysisTableWriterContext
    context.BeginPlan DummySheet
    context.WriteTable BuildPlanItemStub(), DummySheet

    context.BeginPlan DummySheet

    Assert.AreEqual 0&, context.WrittenItems.Length, "Written items should be cleared on BeginPlan"
End Sub

'@TestMethod("AnalysisTableWriterContext")
Public Sub TestWriteTableValidatesPlanState()
    CustomTestSetTitles Assert, "AnalysisTableWriterContext", "TestWriteTableValidatesPlanState"
    Dim context As AnalysisTableWriterContext
    Dim raisedError As Boolean

    Set context = New AnalysisTableWriterContext

    On Error Resume Next
        context.WriteTable BuildPlanItemStub(), DummySheet
        raisedError = (Err.Number = ProjectError.ErrorUnexpectedState)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "WriteTable should enforce BeginPlan call"
End Sub

'@TestMethod("AnalysisTableWriterContext")
Public Sub TestApplyNavigationStoresLabels()
    CustomTestSetTitles Assert, "AnalysisTableWriterContext", "TestApplyNavigationStoresLabels"
    Dim context As AnalysisTableWriterContext
    Dim sections As BetterArray
    Dim headers As BetterArray

    Set context = New AnalysisTableWriterContext
    context.BeginPlan DummySheet

    Set sections = New BetterArray
    sections.LowerBound = 1
    sections.Push AnalysisGoToEntry.Create("section", "One", "sec: ")

    Set headers = New BetterArray
    headers.LowerBound = 1
    headers.Push AnalysisGoToEntry.Create("header", "One", "hdr: ")

    context.ApplyNavigation sections, headers, DummySheet

    Assert.AreEqual 1&, context.SectionLabels.Length, "Section labels should be stored"
    Assert.AreEqual 1&, context.HeaderLabels.Length, "Header labels should be stored"
End Sub
