Attribute VB_Name = "TestAnalysisWorksheetPipeline"
Attribute VB_Description = "Unit tests for AnalysisWorksheetPipeline behaviour"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests verifying worksheet pipeline iteration")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisWorksheetPipeline"
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

Private Function BuildSequence() As BetterArray
    Dim builder As AnalysisWorksheetSpecSequenceBuilder
    Dim sequence As BetterArray
    Dim graphSpecs As BetterArray
    Dim graphSpec As IAnalysisWorksheetGraphSpec

    Set builder = New AnalysisWorksheetSpecSequenceBuilder
    builder.AddSpec AnalysisScopeNormal, Array("Tab_global_summary", "Tab_Univariate_Analysis"), Array(CByte(1), CByte(2)), Nothing, "ua_", vbNullString, vbNullString

    Set graphSpecs = New BetterArray
    graphSpecs.LowerBound = 1
    Set graphSpec = AnalysisWorksheetGraphSpec.Create(AnalysisWorksheetGraphTimeSeries, "Tab_TimeSeries_Analysis", "Tab_Graph_TimeSeries", "Tab_Label_TSGraph", "ts_", "_graph")
    graphSpecs.Push graphSpec

    builder.AddSpec AnalysisScopeTimeSeries, Array("Tab_TimeSeries_Analysis"), Array(CByte(4)), graphSpecs, "ts_", "ts_", "_graph"

    Set sequence = builder.Build
    Set BuildSequence = sequence
End Function

'@TestMethod("AnalysisWorksheetPipeline")
Public Sub TestRunInvokesHandlerForEachSpec()
    CustomTestSetTitles Assert, "AnalysisWorksheetPipeline", "TestRunInvokesHandlerForEachSpec"

    Dim pipeline As IAnalysisWorksheetPipeline
    Dim handler As AnalysisWorksheetHandlerStub
    Dim sequence As BetterArray

    Set pipeline = New AnalysisWorksheetPipeline
    Set handler = New AnalysisWorksheetHandlerStub
    Set sequence = BuildSequence()

    pipeline.Run sequence, handler

    Assert.AreEqual 1&, handler.BeginCount, "BeginSequence should be called once"
    Assert.AreEqual 1&, handler.CompleteCount, "CompleteSequence should be called once"
    Assert.AreEqual sequence.Length, handler.Processed.Length, "Handler should process each specification"
End Sub

'@TestMethod("AnalysisWorksheetPipeline")
Public Sub TestRunValidatesArguments()
    CustomTestSetTitles Assert, "AnalysisWorksheetPipeline", "TestRunValidatesArguments"

    Dim pipeline As IAnalysisWorksheetPipeline
    Dim handler As AnalysisWorksheetHandlerStub
    Dim raisedError As Boolean

    Set pipeline = New AnalysisWorksheetPipeline
    Set handler = New AnalysisWorksheetHandlerStub

    On Error Resume Next
        pipeline.Run Nothing, handler
        raisedError = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Pipeline should validate the sequence argument"
End Sub

'@TestMethod("AnalysisWorksheetPipeline")
Public Sub TestRunValidatesSpecifications()
    CustomTestSetTitles Assert, "AnalysisWorksheetPipeline", "TestRunValidatesSpecifications"

    Dim pipeline As IAnalysisWorksheetPipeline
    Dim handler As AnalysisWorksheetHandlerStub
    Dim sequence As BetterArray
    Dim raisedError As Boolean

    Set pipeline = New AnalysisWorksheetPipeline
    Set handler = New AnalysisWorksheetHandlerStub

    Set sequence = New BetterArray
    sequence.LowerBound = 1
    sequence.Push Nothing

    On Error Resume Next
        pipeline.Run sequence, handler
        raisedError = (Err.Number = ProjectError.InvalidState)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Pipeline should validate individual worksheet specifications"
End Sub
