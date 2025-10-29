Attribute VB_Name = "TestAnalysisWorksheetCoordinator"
Attribute VB_Description = "Unit tests for AnalysisWorksheetCoordinator orchestration"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests verifying AnalysisWorksheetCoordinator integrates engines and pipeline")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const OUTPUT_SHEET_NORMAL As String = "CoordinatorNormal"
Private Const OUTPUT_SHEET_TS As String = "CoordinatorTimeSeries"

Private Assert As ICustomTest
Private NormalSheet As Worksheet
Private TimeSeriesSheet As Worksheet

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisWorksheetCoordinator"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet OUTPUT_SHEET_NORMAL
    DeleteWorksheet OUTPUT_SHEET_TS
    RestoreApp
    Set Assert = Nothing
    Set NormalSheet = Nothing
    Set TimeSeriesSheet = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set NormalSheet = EnsureWorksheet(OUTPUT_SHEET_NORMAL)
    Set TimeSeriesSheet = EnsureWorksheet(OUTPUT_SHEET_TS)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    ClearWorksheet NormalSheet
    ClearWorksheet TimeSeriesSheet
End Sub

Private Function BuildSequence() As BetterArray
    Dim builder As AnalysisWorksheetSpecSequenceBuilder
    Dim graphs As BetterArray
    Dim graphSpec As IAnalysisWorksheetGraphSpec

    Set builder = New AnalysisWorksheetSpecSequenceBuilder
    builder.AddSpec AnalysisScopeNormal, Array("Tab_global_summary"), Array(CByte(1)), Nothing, "ua_", vbNullString, vbNullString

    Set graphs = New BetterArray
    graphs.LowerBound = 1
    Set graphSpec = AnalysisWorksheetGraphSpec.Create(AnalysisWorksheetGraphTimeSeries, "Tab_TimeSeries_Analysis", "Tab_Graph_TimeSeries", "Tab_Label_TSGraph", "ts_", "_graph")
    graphs.Push graphSpec

    builder.AddSpec AnalysisScopeTimeSeries, Array("Tab_TimeSeries_Analysis"), Array(CByte(4)), graphs, "ts_", "ts_", "_graph"

    Set BuildSequence = builder.Build
End Function

Private Function CreateInputsStub() As AnalysisWorksheetInputsStub
    Dim inputsStub As AnalysisWorksheetInputsStub

    Set inputsStub = New AnalysisWorksheetInputsStub
    inputsStub.RegisterOutputSheet AnalysisScopeNormal, NormalSheet
    inputsStub.RegisterOutputSheet AnalysisScopeTimeSeries, TimeSeriesSheet

    Set CreateInputsStub = inputsStub
End Function

'@TestMethod("AnalysisWorksheetCoordinator")
Public Sub TestCoordinatorRunsEngines()
    CustomTestSetTitles Assert, "AnalysisWorksheetCoordinator", "TestCoordinatorRunsEngines"

    Dim coordinator As AnalysisWorksheetCoordinator
    Dim tableEngineStub As AnalysisTableEngineStub
    Dim graphEngineStub As AnalysisGraphEngineStub
    Dim safeRunnerStub As AnalysisSafeRunnerStub
    Dim sequence As BetterArray
    Dim inputsStub As AnalysisWorksheetInputsStub

    Set coordinator = New AnalysisWorksheetCoordinator
    Set tableEngineStub = New AnalysisTableEngineStub
    Set graphEngineStub = New AnalysisGraphEngineStub
    Set safeRunnerStub = New AnalysisSafeRunnerStub
    Set sequence = BuildSequence()
    Set inputsStub = CreateInputsStub()

    coordinator.TableEngine = tableEngineStub
    coordinator.GraphEngine = graphEngineStub
    coordinator.SafeRunner = safeRunnerStub

    coordinator.Run sequence, inputsStub

    Assert.AreEqual 2&, tableEngineStub.RunCount, "Table engine should be invoked for each table specification"
    Assert.AreEqual 1&, graphEngineStub.RunCount, "Graph engine should be invoked for time-series graphs"
    Assert.AreEqual 2&, coordinator.GoToEntries.SectionEntries.Length, "GoTo entries should accumulate section labels"
    Assert.AreEqual 2&, coordinator.DropdownRequests.Length, "Coordinator should create dropdown requests"
    Assert.AreEqual 3&, safeRunnerStub.RunCount, "Safe runner should wrap each engine call"
End Sub
