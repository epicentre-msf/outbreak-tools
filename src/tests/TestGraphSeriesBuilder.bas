Attribute VB_Name = "TestGraphSeriesBuilder"
Option Explicit
Option Private Module

'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests validating the GraphSeriesBuilder and related classes")

Private Assert As Object
Private Builder As IGraphSeriesBuilder

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
    Set Builder = GraphSeriesBuilder.Create
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Builder = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("GraphSeriesBuilder")
Private Sub TestAddSeriesAndAssignLabels()
    Dim spec As IGraphSeriesSpec

    On Error GoTo Fail

    Set spec = Builder.AddSeries("cases", "bar", "primary")
    Builder.AssignLabels "row label", "column label", "prefix"

    Assert.AreEqual "cases", spec.SeriesName
    Assert.AreEqual "bar", spec.SeriesType
    Assert.AreEqual "primary", spec.SeriesPosition
    Assert.AreEqual "row label", spec.RowLabel
    Assert.AreEqual "column label", spec.ColumnLabel
    Assert.AreEqual "prefix", spec.LabelPrefix
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddSeriesAndAssignLabels"
End Sub

'@TestMethod("GraphSeriesBuilder")
Private Sub TestAssignLabelsWithoutSeriesRaises()
    On Error Resume Next
    Builder.AssignLabels "row", "col"

    Dim errNumber As Long
    errNumber = Err.Number
    Err.Clear
    On Error GoTo Fail

    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), errNumber, _
        "Assigning labels without registering a series should raise ErrorUnexpectedState"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAssignLabelsWithoutSeriesRaises"
End Sub

'@TestMethod("GraphSeriesBuilder")
Private Sub TestBuildGraphReturnsChartSpec()
    Dim chartSpec As IGraphChartSpec
    Dim seriesList As BetterArray

    On Error GoTo Fail

    Builder.AddSeries "cases", "bar", "primary"
    Builder.AssignLabels "row", "col"
    Builder.AddSeries "deaths", "line", "secondary"
    Builder.AssignLabels "row2", "col2", "pref"

    Set chartSpec = Builder.BuildGraph("graph_1", "Cases over time")
    Set seriesList = chartSpec.Series

    Assert.AreEqual "graph_1", chartSpec.GraphId
    Assert.AreEqual "Cases over time", chartSpec.Title
    Assert.AreEqual 2&, chartSpec.SeriesCount
    Assert.AreEqual 2&, seriesList.Length

    Dim firstSeries As IGraphSeriesSpec
    Dim secondSeries As IGraphSeriesSpec

    Set firstSeries = seriesList.Item(seriesList.LowerBound)
    Set secondSeries = seriesList.Item(seriesList.LowerBound + 1)

    Assert.AreEqual "cases", firstSeries.SeriesName
    Assert.AreEqual "deaths", secondSeries.SeriesName
    Assert.AreEqual "secondary", secondSeries.SeriesPosition

    'Builder should clear its cache after build
    Assert.AreEqual 0&, Builder.SeriesCount
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestBuildGraphReturnsChartSpec"
End Sub

'@TestMethod("GraphSeriesBuilder")
Private Sub TestClearResetsBuilder()
    On Error GoTo Fail

    Builder.AddSeries "cases", "bar", "primary"
    Builder.Clear

    Assert.AreEqual 0&, Builder.SeriesCount, "Clear should drop buffered series"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestClearResetsBuilder"
End Sub

