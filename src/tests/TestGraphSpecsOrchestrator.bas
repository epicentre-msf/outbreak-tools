Attribute VB_Name = "TestGraphSpecsOrchestrator"
Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests ensuring GraphSpecsOrchestrator coordinates graph builders correctly")

Private Const SIMPLE_SHEET As String = "OrchestratorSimple"
Private Const GRAPH_SHEET As String = "OrchestratorGraph"
Private Const SERIES_SHEET As String = "OrchestratorSeries"
Private Const TITLE_SHEET As String = "OrchestratorTitles"

Private Assert As Object

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet SIMPLE_SHEET
    DeleteWorksheet GRAPH_SHEET
    DeleteWorksheet SERIES_SHEET
    DeleteWorksheet TITLE_SHEET
    ClearWorkbookName "COLUMN_CATEGORIES_TAB500"
    ClearWorkbookName "LABEL_COL_1_TAB500"
    ClearWorkbookName "LABEL_COL_2_TAB500"
    Set Assert = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("GraphSpecsOrchestrator")
Private Sub TestSimpleBuildUsesCrossBuilder()
    Dim specsStub As GraphTablesSpecsStub
    Dim tableStub As CrossTableGraphTableStub
    Dim builderStub As CrossTableGraphBuilderStub
    Dim orchestrator As IGraphSpecsOrchestrator
    Dim charts As BetterArray
    Dim resultChart As IGraphChartSpec

    On Error GoTo Fail

    Set specsStub = New GraphTablesSpecsStub
    specsStub.Configure CInt(TABLE_TYPE_BIVARIATE), "TAB001"
    specsStub.SetValue "graph title", "Ward Graph"
    specsStub.SetHasPercentage True

    Set tableStub = New CrossTableGraphTableStub
    tableStub.Configure specsStub.Self, 3

    Set builderStub = New CrossTableGraphBuilderStub
    builderStub.ConfigureReturnChart CreateSampleChart("TAB001", "Ward Graph", Array("seriesA", "seriesB"))

    Set orchestrator = GraphSpecsOrchestrator.CreateForCrossTable(tableStub.Self, builderStub)

    Set charts = orchestrator.Build

    Assert.AreEqual 1&, orchestrator.GraphCount
    Assert.AreEqual 2&, orchestrator.SeriesCount
    Assert.IsFalse orchestrator.IsComplex
    Assert.AreEqual "TAB001", builderStub.LastGraphId
    Assert.AreEqual "Ward Graph", builderStub.LastGraphTitle

    Set resultChart = charts.Item(charts.LowerBound)
    Assert.AreEqual "TAB001", resultChart.GraphId
    Assert.AreEqual 2&, resultChart.SeriesCount
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestSimpleBuildUsesCrossBuilder"
End Sub

'@TestMethod("GraphSpecsOrchestrator")
Private Sub TestComplexBuildUsesTimeSeriesBuilder()
    Dim graphList As ListObject
    Dim seriesList As ListObject
    Dim titleList As ListObject
    Dim linelist As ILinelistSpecs
    Dim timeStub As TimeSeriesGraphBuilderStub
    Dim orchestrator As IGraphSpecsOrchestrator
    Dim charts As BetterArray

    On Error GoTo Fail

    Set graphList = SeedGraphList()
    Set seriesList = SeedSeriesList()
    Set titleList = SeedTitleList()
    Set linelist = New TableSpecsLinelistStub

    Set timeStub = New TimeSeriesGraphBuilderStub
    timeStub.ConfigureCharts CreateChartList(Array("GraphTS1", "GraphTS2"))

    Set orchestrator = GraphSpecsOrchestrator.CreateForTimeSeries(graphList, seriesList, titleList, linelist, Nothing, Nothing, timeStub)

    Set charts = orchestrator.Build

    Assert.IsTrue orchestrator.IsComplex
    Assert.AreEqual 2&, orchestrator.GraphCount
    Assert.AreEqual 4&, orchestrator.SeriesCount
    Assert.AreEqual 2&, charts.Length
    Assert.AreEqual "GraphTS1", charts.Item(charts.LowerBound).GraphId

    orchestrator.Invalidate
    Assert.AreEqual 1&, timeStub.InvalidateCount
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestComplexBuildUsesTimeSeriesBuilder"
End Sub

'@section Helper functions
'===============================================================================

Private Function CreateSampleChart(ByVal graphId As String, ByVal graphTitle As String,  seriesNames As Variant) As IGraphChartSpec
    Dim builder As IGraphSeriesBuilder
    Dim idx As Long
    Dim nameText As String

    Set builder = GraphSeriesBuilder.Create

    For idx = LBound(seriesNames) To UBound(seriesNames)
        nameText = CStr(seriesNames(idx))
        builder.AddSeries nameText, IIf(idx Mod 2 = 0, "bar", "line"), IIf(idx Mod 2 = 0, "left", "right")
        builder.AssignLabels "row" & CStr(idx + 1), "col" & CStr(idx + 1)
    Next idx

    Set CreateSampleChart = builder.BuildGraph(graphId, graphTitle)
End Function

Private Function CreateChartList( graphIds As Variant) As BetterArray
    Dim charts As BetterArray
    Dim idx As Long
    Dim chart As IGraphChartSpec

    Set charts = New BetterArray
    charts.LowerBound = 1

    For idx = LBound(graphIds) To UBound(graphIds)
        Set chart = CreateSampleChart(CStr(graphIds(idx)), "Title " & CStr(idx + 1), Array("S" & idx, "T" & idx))
        charts.Push chart
    Next idx

    Set CreateChartList = charts
End Function

Private Function SeedGraphList() As ListObject
    Dim sh As Worksheet
    Dim dataRange As Range

    Set sh = EnsureWorksheet(GRAPH_SHEET)
    sh.Cells.Clear

    sh.Range("A1:G1").Value = Array("graph id", "series id", "axis", "type", "percentages", "choices", "label")
    sh.Range("A2:G3").Value = Array( _
        Array("GraphTS1", "SeriesA", "primary", "line", "", "ChoiceA", ""), _
        Array("GraphTS2", "SeriesB", "secondary", "line", "", "ChoiceB", ""))

    If sh.ListObjects.Count > 0 Then sh.ListObjects(1).Delete
    Set dataRange = sh.Range("A1").CurrentRegion
    Set SeedGraphList = sh.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
End Function

Private Function SeedSeriesList() As ListObject
    Dim sh As Worksheet
    Dim dataRange As Range

    Set sh = EnsureWorksheet(SERIES_SHEET)
    sh.Cells.Clear

    sh.Range("A1:D1").Value = Array("series id", "table id", "placeholder", "value")
    sh.Range("A2:D3").Value = Array( _
        Array("SeriesA", "TAB500", "", ""), _
        Array("SeriesB", "TAB500", "", ""))

    If sh.ListObjects.Count > 0 Then sh.ListObjects(1).Delete
    Set dataRange = sh.Range("A1").CurrentRegion
    Set SeedSeriesList = sh.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
End Function

Private Function SeedTitleList() As ListObject
    Dim sh As Worksheet
    Dim dataRange As Range

    Set sh = EnsureWorksheet(TITLE_SHEET)
    sh.Cells.Clear

    sh.Range("A1:C1").Value = Array("title", "unused", "graph id")
    sh.Range("A2:C3").Value = Array( _
        Array("Admissions", "", "GraphTS1"), _
        Array("Recoveries", "", "GraphTS2"))

    If sh.ListObjects.Count > 0 Then sh.ListObjects(1).Delete
    Set dataRange = sh.Range("A1").CurrentRegion
    Set SeedTitleList = sh.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
End Function

Private Sub ClearWorkbookName(ByVal nameText As String)
    Dim wb As Workbook

    Set wb = Application.ThisWorkbook

    On Error Resume Next
        wb.Names(nameText).Delete
    Err.Clear
    On Error GoTo 0
End Sub

