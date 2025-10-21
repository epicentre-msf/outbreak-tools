Attribute VB_Name = "TestAnalysisGraphEngine"
Attribute VB_Description = "Unit tests for AnalysisGraphEngine orchestration"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests verifying AnalysisGraphEngine behaviour")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisGraphEngine"
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

'@section Helpers
'===============================================================================
Private Function BuildPolicyResult(ByVal tableType As AnalysisTablesType, _
                                   ByVal hasGraph As Boolean, _
                                   ByVal isNewSection As Boolean, _
                                   ByVal tableId As String, _
                                   ByVal sectionText As String, _
                                   ByVal labelText As String) As IAnalysisTablePolicyResult

    Dim specStub As GraphTablesSpecsStub
    Dim iterationStub As AnalysisTableIterationItemStub
    Dim policyStub As AnalysisTypePolicyStub
    Dim contextStub As AnalysisPolicyContextStub

    Set specStub = New GraphTablesSpecsStub
    specStub.Configure tableType, tableId, "section_" & tableId
    specStub.SetHasGraph hasGraph
    specStub.SetValue "section", sectionText
    specStub.SetValue "label", labelText

    Set iterationStub = New AnalysisTableIterationItemStub
    iterationStub.Configure specStub.Self, isNewSection

    Set policyStub = New AnalysisTypePolicyStub
    policyStub.Configure tableType, True, False, False, False, hasGraph

    Set contextStub = New AnalysisPolicyContextStub
    contextStub.Configure tableType

    Set BuildPolicyResult = AnalysisTablePolicyResult.Create(iterationStub, policyStub, contextStub, True, False, False, False, hasGraph)
End Function

Private Function BuildPlanResult(Optional ByVal includeCrossTable As Boolean = True, _
                                 Optional ByVal includeTimeSeries As Boolean = False) As IAnalysisTablePlanResult

    Dim policyResults As BetterArray
    Dim planBuilder As AnalysisTablePlanBuilder

    Set policyResults = New BetterArray
    policyResults.LowerBound = 1

    If includeCrossTable Then
        policyResults.Push BuildPolicyResult(TypeUnivariate, True, True, "table_cross_1", "Cross Section", "Cross Label 1")
    End If

    If includeTimeSeries Then
        policyResults.Push BuildPolicyResult(TypeTimeSeries, True, True, "table_ts_1", "TS Section", "TS Label 1")
    End If

    'Add an entry with HasGraph = False to ensure filtering works
    policyResults.Push BuildPolicyResult(TypeBivariate, False, False, "table_ignore", "Ignore Section", "Ignore Label")

    Set planBuilder = New AnalysisTablePlanBuilder
    Set BuildPlanResult = planBuilder.Build(policyResults, "sec: ", "hdr: ")
End Function

Private Function BuildChartList(ParamArray ids() As Variant) As BetterArray
    Dim charts As BetterArray
    Dim idx As Long
    Dim chart As IGraphChartSpec
    Dim graphId As String

    Set charts = New BetterArray
    charts.LowerBound = 1

    For idx = LBound(ids) To UBound(ids)
        graphId = CStr(ids(idx))
        Set chart = GraphChartSpec.Create(graphId, "Title " & graphId)
        charts.Push chart
    Next idx

    Set BuildChartList = charts
End Function

Private Function CreateEngine(Optional ByVal factory As IAnalysisGraphOrchestratorFactory = Nothing, _
                              Optional ByVal builder As IAnalysisGraphPlanBuilder = Nothing) As AnalysisGraphEngine
    Dim engine As AnalysisGraphEngine
    Set engine = New AnalysisGraphEngine

    If Not factory Is Nothing Then
        engine.OrchestratorFactory = factory
    End If

    If Not builder Is Nothing Then
        engine.PlanBuilder = builder
    End If

    Set CreateEngine = engine
End Function

'@section Tests
'===============================================================================
'@TestMethod("AnalysisGraphEngine")
Public Sub TestRunProducesCrossTableCharts()
    CustomTestSetTitles Assert, "AnalysisGraphEngine", "TestRunProducesCrossTableCharts"

    Dim plan As IAnalysisTablePlanResult
    Dim factoryStub As AnalysisGraphOrchestratorFactoryStub
    Dim inputsStub As AnalysisGraphEngineInputsStub
    Dim crossTableStub As CrossTableStub
    Dim engine As AnalysisGraphEngine
    Dim result As IAnalysisGraphExecutionResult
    Dim item As IAnalysisGraphExecutionItem

    Set plan = BuildPlanResult(includeCrossTable:=True, includeTimeSeries:=False)

    Set factoryStub = New AnalysisGraphOrchestratorFactoryStub
    factoryStub.ConfigureCrossTableCharts BuildChartList("cross-1")

    Set inputsStub = New AnalysisGraphEngineInputsStub
    Set crossTableStub = New CrossTableStub
    inputsStub.ConfigureCrossTable "table_cross_1", crossTableStub

    Set engine = CreateEngine(factoryStub)
    Set result = engine.Run(plan, inputsStub)

    Assert.IsTrue result.HasGraphs, "Engine should report graphs produced"
    Assert.IsFalse result.HasTimeSeriesCharts, "No time-series charts expected"
    Assert.AreEqual 1&, result.GraphCount, "Exactly one chart expected"
    Assert.AreEqual 1&, result.CrossTableItems.Length, "One cross-table execution item expected"

    Set item = result.CrossTableItems.Item(1)
    Assert.AreEqual 1&, item.ChartCount, "Execution item should track generated chart count"
    Assert.AreEqual "table_cross_1", Trim$(item.PlanItem.Specification.TableId), "Plan item should reference cross-table specification"
End Sub

'@TestMethod("AnalysisGraphEngine")
Public Sub TestRunProducesTimeSeriesCharts()
    CustomTestSetTitles Assert, "AnalysisGraphEngine", "TestRunProducesTimeSeriesCharts"

    Dim plan As IAnalysisTablePlanResult
    Dim factoryStub As AnalysisGraphOrchestratorFactoryStub
    Dim inputsStub As AnalysisGraphEngineInputsStub
    Dim engine As AnalysisGraphEngine
    Dim result As IAnalysisGraphExecutionResult

    Set plan = BuildPlanResult(includeCrossTable:=False, includeTimeSeries:=True)

    Set factoryStub = New AnalysisGraphOrchestratorFactoryStub
    factoryStub.ConfigureTimeSeriesCharts BuildChartList("ts-1", "ts-2")

    Set inputsStub = New AnalysisGraphEngineInputsStub

    Set engine = CreateEngine(factoryStub)
    Set result = engine.Run(plan, inputsStub)

    Assert.IsTrue result.HasGraphs, "Engine should report graphs produced"
    Assert.IsTrue result.HasTimeSeriesCharts, "Time-series charts should be reported"
    Assert.AreEqual 2&, result.TimeSeriesCharts.Length, "Two time-series charts expected"
    Assert.AreEqual 2&, result.GraphCount, "Total chart count should match time-series charts"
    Assert.AreEqual 0&, result.CrossTableItems.Length, "No cross-table execution items expected"
End Sub

'@TestMethod("AnalysisGraphEngine")
Public Sub TestRunValidatesPlan()
    CustomTestSetTitles Assert, "AnalysisGraphEngine", "TestRunValidatesPlan"

    Dim engine As AnalysisGraphEngine
    Dim inputsStub As AnalysisGraphEngineInputsStub
    Dim raisedError As Boolean

    Set engine = CreateEngine()
    Set inputsStub = New AnalysisGraphEngineInputsStub

    On Error Resume Next
        engine.Run Nothing, inputsStub
        raisedError = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Engine should validate the plan argument"
End Sub

'@TestMethod("AnalysisGraphEngine")
Public Sub TestRunRequiresCrossTableWhenNeeded()
    CustomTestSetTitles Assert, "AnalysisGraphEngine", "TestRunRequiresCrossTableWhenNeeded"

    Dim plan As IAnalysisTablePlanResult
    Dim engine As AnalysisGraphEngine
    Dim inputsStub As AnalysisGraphEngineInputsStub
    Dim raisedError As Boolean

    Set plan = BuildPlanResult(includeCrossTable:=True, includeTimeSeries:=False)
    Set engine = CreateEngine(New AnalysisGraphOrchestratorFactoryStub)
    Set inputsStub = New AnalysisGraphEngineInputsStub

    On Error Resume Next
        engine.Run plan, inputsStub
        raisedError = (Err.Number = ProjectError.ErrorUnexpectedState)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Engine should require a cross table when plan requests a graph"
End Sub
