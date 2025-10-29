Attribute VB_Name = "TestAnalysisWorksheetSpec"
Attribute VB_Description = "Unit tests for AnalysisWorksheetSpec and related builders"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests verifying AnalysisWorksheetSpec and graph spec behaviour")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TABLE_GLOBALSUMMARY As String = "Tab_global_summary"
Private Const TABLE_UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const TABLE_BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TABLE_TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const TABLE_TIMESERIES_GRAPHS As String = "Tab_Graph_TimeSeries"
Private Const TABLE_TIMESERIES_TITLES As String = "Tab_Label_TSGraph"
Private Const TABLE_SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const TYPE_GLOBALSUMMARY As Byte = 1
Private Const TYPE_UNIVARIATE As Byte = 2
Private Const TYPE_BIVARIATE As Byte = 3
Private Const TYPE_TIMESERIES As Byte = 4

Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisWorksheetSpec"
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

Private Function CreateGraphSpecs() As BetterArray
    Dim graphs As BetterArray
    Dim spec As IAnalysisWorksheetGraphSpec

    Set graphs = New BetterArray
    graphs.LowerBound = 1

    Set spec = AnalysisWorksheetGraphSpec.Create(AnalysisWorksheetGraphTimeSeries, TABLE_TIMESERIES, TABLE_TIMESERIES_GRAPHS, TABLE_TIMESERIES_TITLES, "ts_", "_graph")
    graphs.Push spec

    Set CreateGraphSpecs = graphs
End Function

'@TestMethod("AnalysisWorksheetSpec")
Public Sub TestCreateSpecStoresMetadata()
    CustomTestSetTitles Assert, "AnalysisWorksheetSpec", "TestCreateSpecStoresMetadata"

    Dim spec As IAnalysisWorksheetSpec
    Dim tables As Variant
    Dim graphs As BetterArray

    tables = Array(TABLE_GLOBALSUMMARY, TABLE_UNIVARIATE, TABLE_BIVARIATE)
    Set graphs = CreateGraphSpecs()
    Dim tableTypes As Variant
    tableTypes = Array(TYPE_GLOBALSUMMARY, TYPE_UNIVARIATE, TYPE_BIVARIATE)

    Set spec = AnalysisWorksheetSpec.Create(AnalysisScopeTimeSeries, tables, tableTypes, graphs, "ts_", "ts_", "_graph")

    Assert.AreEqual AnalysisScopeTimeSeries, spec.Scope, "Scope should be stored"
    Assert.AreEqual 3&, spec.TableIds.Length, "Table ids should be preserved"
    Assert.AreEqual 3&, spec.TableTypeCodes.Length, "Table type codes should align with table ids"
    Assert.IsTrue spec.HasGraphs, "Spec should flag graphs present"
    Assert.AreEqual "ts_", spec.SectionGoToPrefix, "Section prefix should be stored"
    Assert.AreEqual "ts_", spec.GraphGoToPrefix, "Graph prefix should be stored"
    Assert.AreEqual "_graph", spec.GraphGoToSuffix, "Graph suffix should be stored"
End Sub

'@TestMethod("AnalysisWorksheetSpec")
Public Sub TestGraphSpecEncapsulatesMetadata()
    CustomTestSetTitles Assert, "AnalysisWorksheetSpec", "TestGraphSpecEncapsulatesMetadata"

    Dim graphSpec As IAnalysisWorksheetGraphSpec
    Set graphSpec = AnalysisWorksheetGraphSpec.Create(AnalysisWorksheetGraphSpatioTemporal, TABLE_SPATIOTEMPORAL, TABLE_SPATIOTEMPORAL, vbNullString, "spt_", "_graph")

    Assert.AreEqual AnalysisWorksheetGraphSpatioTemporal, graphSpec.GraphType, "Graph type should match constructor"
    Assert.AreEqual TABLE_SPATIOTEMPORAL, graphSpec.TableId, "Table id should match constructor"
    Assert.AreEqual TABLE_SPATIOTEMPORAL, graphSpec.GraphListId, "Graph list id should match constructor"
    Assert.AreEqual "spt_", graphSpec.GoToPrefix, "GoTo prefix should match constructor"
    Assert.AreEqual "_graph", graphSpec.GoToSuffix, "GoTo suffix should match constructor"
End Sub

'@TestMethod("AnalysisWorksheetSpec")
Public Sub TestSequenceBuilderAggregatesSpecs()
    CustomTestSetTitles Assert, "AnalysisWorksheetSpec", "TestSequenceBuilderAggregatesSpecs"

    Dim builder As IAnalysisWorksheetSpecSequenceBuilder
    Dim sequence As BetterArray
    Dim spec As IAnalysisWorksheetSpec

    Set builder = New AnalysisWorksheetSpecSequenceBuilder
    builder.AddSpec AnalysisScopeNormal, Array(TABLE_GLOBALSUMMARY, TABLE_UNIVARIATE, TABLE_BIVARIATE), Array(TYPE_GLOBALSUMMARY, TYPE_UNIVARIATE, TYPE_BIVARIATE), Nothing, "ua_", vbNullString, vbNullString
    builder.AddSpec AnalysisScopeTimeSeries, Array(TABLE_TIMESERIES), Array(TYPE_TIMESERIES), CreateGraphSpecs(), "ts_", "ts_", "_graph"

    Assert.AreEqual 2&, builder.Count, "Builder should track number of specs"

    Set sequence = builder.Build
    Assert.AreEqual 2&, sequence.Length, "Sequence should contain added specs"

    Set spec = sequence.Item(1)
    Assert.AreEqual AnalysisScopeNormal, spec.Scope, "First spec should preserve scope"
    Assert.IsFalse spec.HasGraphs, "First spec should not have graphs"

    Set spec = sequence.Item(2)
    Assert.AreEqual AnalysisScopeTimeSeries, spec.Scope, "Second spec should preserve scope"
    Assert.IsTrue spec.HasGraphs, "Second spec should have graphs"
End Sub
