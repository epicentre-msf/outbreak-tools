Attribute VB_Name = "TestCrossTableGraphBuilder"
Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests covering CrossTableGraphBuilder behaviour")

Private Assert As Object
Private Builder As ICrossTableGraphBuilder

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Builder = Nothing
    Set Assert = Nothing
End Sub

'@section Test lifecycle
'===============================================================================

'@TestInitialize
Private Sub TestInitialize()
    Set Builder = CrossTableGraphBuilder.Create
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Builder = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("CrossTableGraphBuilder")
Private Sub TestUnivariateSeriesConfiguration()
    Dim specsStub As GraphTablesSpecsStub
    Dim tableStub As CrossTableGraphTableStub
    Dim chart As IGraphChartSpec
    Dim seriesList As BetterArray
    Dim firstSeries As IGraphSeriesSpec
    Dim secondSeries As IGraphSeriesSpec

    On Error GoTo Fail

    Set specsStub = New GraphTablesSpecsStub
    specsStub.Configure TABLE_TYPE_UNIVARIATE, "TAB001"
    specsStub.SetHasPercentage True

    Set tableStub = New CrossTableGraphTableStub
    tableStub.Configure specsStub.Self, 1

    Set chart = Builder.BuildGraph(tableStub.Self, "Graph_TAB001", "Tab 001")
    Set seriesList = chart.Series

    Assert.AreEqual 2&, chart.SeriesCount
    Assert.AreEqual 2&, seriesList.Length

    Set firstSeries = seriesList.Item(seriesList.LowerBound)
    Set secondSeries = seriesList.Item(seriesList.LowerBound + 1)

    Assert.AreEqual "VALUES_COL_1_TAB001", firstSeries.SeriesName
    Assert.AreEqual "ROW_CATEGORIES_TAB001", firstSeries.RowLabel
    Assert.AreEqual "PERC_COL_1_TAB001", secondSeries.SeriesName
    Assert.AreEqual "PERC_LABEL_COL_TAB001", secondSeries.ColumnLabel
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestUnivariateSeriesConfiguration"
End Sub

'@TestMethod("CrossTableGraphBuilder")
Private Sub TestBivariateValuesMode()
    Dim specsStub As GraphTablesSpecsStub
    Dim tableStub As CrossTableGraphTableStub
    Dim chart As IGraphChartSpec
    Dim seriesList As BetterArray

    On Error GoTo Fail

    Set specsStub = New GraphTablesSpecsStub
    specsStub.Configure TABLE_TYPE_BIVARIATE, "TAB010"
    specsStub.SetValue "graph", "values"

    Set tableStub = New CrossTableGraphTableStub
    tableStub.Configure specsStub.Self, 3

    Set chart = Builder.BuildGraph(tableStub.Self, "Graph_TAB010")
    Set seriesList = chart.Series

    Assert.AreEqual 3&, chart.SeriesCount
    Assert.AreEqual "VALUES_COL_2_TAB010", _
        CStr(seriesList.Item(seriesList.LowerBound + 1).SeriesName)
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestBivariateValuesMode"
End Sub

'@TestMethod("CrossTableGraphBuilder")
Private Sub TestSpatioTemporalResolvesSectionLabels()
    Dim specsStub As GraphTablesSpecsStub
    Dim tableStub As CrossTableGraphTableStub
    Dim chart As IGraphChartSpec
    Dim seriesList As BetterArray
    Dim firstSeries As IGraphSeriesSpec

    On Error GoTo Fail

    Set specsStub = New GraphTablesSpecsStub
    specsStub.Configure TABLE_TYPE_SPATIO_TEMPORAL, "TAB100", "SEC0001"
    specsStub.SetValue "graph", "values"
    specsStub.SetValue "n geo", "2"

    Set tableStub = New CrossTableGraphTableStub
    tableStub.Configure specsStub.Self, 5

    Set chart = Builder.BuildGraph(tableStub.Self, "Graph_TAB100")
    Set seriesList = chart.Series
    Set firstSeries = seriesList.Item(seriesList.LowerBound)

    Assert.AreEqual "ROW_CATEGORIES_SEC0001", firstSeries.RowLabel
    Assert.AreEqual 2&, chart.SeriesCount
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestSpatioTemporalResolvesSectionLabels"
End Sub

