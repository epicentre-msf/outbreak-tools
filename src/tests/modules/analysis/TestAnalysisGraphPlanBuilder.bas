Attribute VB_Name = "TestAnalysisGraphPlanBuilder"
Attribute VB_Description = "Unit tests for AnalysisGraphPlanBuilder classification logic"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests exercising AnalysisGraphPlanBuilder")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisGraphPlanBuilder"
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
Private Function CreatePolicyResult(ByVal tableType As AnalysisTablesType, _
                                    ByVal hasGraph As Boolean, _
                                    ByVal tableId As String) As IAnalysisTablePolicyResult
    Dim specStub As GraphTablesSpecsStub
    Dim iterationStub As AnalysisTableIterationItemStub
    Dim policyStub As AnalysisTypePolicyStub
    Dim contextStub As AnalysisPolicyContextStub

    Set specStub = New GraphTablesSpecsStub
    specStub.Configure tableType, tableId
    specStub.SetHasGraph hasGraph
    specStub.SetValue "label", "Label " & tableId

    Set iterationStub = New AnalysisTableIterationItemStub
    iterationStub.Configure specStub.Self, False

    Set policyStub = New AnalysisTypePolicyStub
    policyStub.Configure tableType, True, False, False, False, hasGraph

    Set contextStub = New AnalysisPolicyContextStub
    contextStub.Configure tableType

    Set CreatePolicyResult = AnalysisTablePolicyResult.Create(iterationStub, policyStub, contextStub, True, False, False, False, hasGraph)
End Function

Private Function BuildPlanResult() As IAnalysisTablePlanResult
    Dim items As BetterArray
    Dim sectionLabels As BetterArray
    Dim headerLabels As BetterArray
    Dim crossPolicy As IAnalysisTablePolicyResult
    Dim timePolicy As IAnalysisTablePolicyResult
    Dim ignoredPolicy As IAnalysisTablePolicyResult

    Set items = New BetterArray
    items.LowerBound = 1

    Set sectionLabels = New BetterArray
    sectionLabels.LowerBound = 1

    Set headerLabels = New BetterArray
    headerLabels.LowerBound = 1

    Set crossPolicy = CreatePolicyResult(TypeUnivariate, True, "cross_1")
    Set timePolicy = CreatePolicyResult(TypeTimeSeries, True, "ts_1")
    Set ignoredPolicy = CreatePolicyResult(TypeBivariate, False, "no_graph")

    items.Push AnalysisTablePlanItem.Create(crossPolicy, 0&)
    items.Push AnalysisTablePlanItem.Create(timePolicy, 1&)
    items.Push AnalysisTablePlanItem.Create(ignoredPolicy, 2&)

    Set BuildPlanResult = AnalysisTablePlanResult.Create(items, sectionLabels, headerLabels)
End Function

'@section Tests
'===============================================================================
'@TestMethod("AnalysisGraphPlanBuilder")
Public Sub TestBuildFiltersGraphItems()
    CustomTestSetTitles Assert, "AnalysisGraphPlanBuilder", "TestBuildFiltersGraphItems"
    Dim builder As IAnalysisGraphPlanBuilder
    Dim plan As IAnalysisTablePlanResult
    Dim graphPlan As IAnalysisGraphPlanResult

    Set builder = New AnalysisGraphPlanBuilder
    Set plan = BuildPlanResult()

    Set graphPlan = builder.Build(plan)

    Assert.AreEqual 2&, graphPlan.GraphCount, "Only policy results with HasGraph should be emitted"
    Assert.IsTrue graphPlan.HasGraphs, "Plan should report graphs present"
End Sub

'@TestMethod("AnalysisGraphPlanBuilder")
Public Sub TestBuildCategorisesTimeSeriesGraphs()
    CustomTestSetTitles Assert, "AnalysisGraphPlanBuilder", "TestBuildCategorisesTimeSeriesGraphs"
    Dim builder As IAnalysisGraphPlanBuilder
    Dim plan As IAnalysisTablePlanResult
    Dim graphPlan As IAnalysisGraphPlanResult
    Dim timeSeriesItem As IAnalysisGraphPlanItem
    Dim crossItem As IAnalysisGraphPlanItem

    Set builder = New AnalysisGraphPlanBuilder
    Set plan = BuildPlanResult()

    Set graphPlan = builder.Build(plan)

    Assert.AreEqual 1&, graphPlan.TimeSeriesItems.Length, "Exactly one time-series graph expected"
    Assert.AreEqual 1&, graphPlan.CrossTableItems.Length, "Exactly one cross-table graph expected"
    Assert.IsTrue graphPlan.HasTimeSeriesGraphs, "Plan should flag presence of time-series graphs"

    Set timeSeriesItem = graphPlan.TimeSeriesItems.Item(1)
    Set crossItem = graphPlan.CrossTableItems.Item(1)

    Assert.IsTrue timeSeriesItem.IsTimeSeries, "Time-series item should report time-series classification"
    Assert.IsFalse crossItem.IsTimeSeries, "Cross-table item should not report time-series classification"
End Sub

'@TestMethod("AnalysisGraphPlanBuilder")
Public Sub TestBuildValidatesPlan()
    CustomTestSetTitles Assert, "AnalysisGraphPlanBuilder", "TestBuildValidatesPlan"
    Dim builder As IAnalysisGraphPlanBuilder
    Dim raisedError As Boolean

    Set builder = New AnalysisGraphPlanBuilder

    On Error Resume Next
        builder.Build Nothing
        raisedError = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Builder should validate the plan argument"
End Sub
