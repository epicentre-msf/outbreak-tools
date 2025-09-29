Attribute VB_Name = "TestTimeSeriesGraphBuilder"
Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests for the TimeSeriesGraphBuilder helper")

Private Const GRAPH_SHEET As String = "TSBuilderGraph"
Private Const SERIES_SHEET As String = "TSBuilderSeries"
Private Const TITLE_SHEET As String = "TSBuilderTitles"

Private Assert As Object
Private GraphList As ListObject
Private SeriesList As ListObject
Private TitleList As ListObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet GRAPH_SHEET
    DeleteWorksheet SERIES_SHEET
    DeleteWorksheet TITLE_SHEET
    Set GraphList = Nothing
    Set SeriesList = Nothing
    Set TitleList = Nothing
    Set Assert = Nothing
End Sub

'@section Test lifecycle
'===============================================================================

'@TestInitialize
Private Sub TestInitialize()
    Set GraphList = SeedGraphList
    Set SeriesList = SeedSeriesList
    Set TitleList = SeedTitleList
    SeedCategoryNames "TAB500"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set GraphList = Nothing
    Set SeriesList = Nothing
    Set TitleList = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("TimeSeriesGraphBuilder")
Private Sub TestBuildGraphCreatesExpectedSeries()
    Dim cache As IGraphSpecsCache
    Dim builder As ITimeSeriesGraphBuilder
    Dim linelist As ILinelistSpecs
    Dim specsStubA As GraphTablesSpecsStub
    Dim specsStubB As GraphTablesSpecsStub
    Dim factoryStub As TimeSeriesSpecsFactoryStub
    Dim chart As IGraphChartSpec
    Dim seriesList As BetterArray
    Dim firstSeries As IGraphSeriesSpec
    Dim secondSeries As IGraphSeriesSpec

    On Error GoTo Fail

    Set cache = GraphSpecsCache.Create(GraphList)
    Set linelist = New TableSpecsLinelistStub

    Set specsStubA = New GraphTablesSpecsStub
    specsStubA.Configure CInt(TABLE_TYPE_TIME_SERIES), "TAB500"

    Set specsStubB = New GraphTablesSpecsStub
    specsStubB.Configure CInt(TABLE_TYPE_TIME_SERIES), "TAB500"

    Set factoryStub = New TimeSeriesSpecsFactoryStub
    factoryStub.AddSpec "SeriesA", specsStubA.Self
    factoryStub.AddSpec "SeriesB", specsStubB.Self

    Set builder = TimeSeriesGraphBuilder.Create(cache, SeriesList, TitleList, linelist, , factoryStub.Self)

    Set chart = builder.BuildGraph("GraphTS1")
    Set seriesList = chart.Series

    Assert.AreEqual "Total admissions", chart.Title
    Assert.AreEqual 2&, chart.SeriesCount

    Set firstSeries = seriesList.Item(seriesList.LowerBound)
    Set secondSeries = seriesList.Item(seriesList.LowerBound + 1)

    Assert.AreEqual "VALUES_COL_1_TAB500", firstSeries.SeriesName
    Assert.AreEqual "ROW_CATEGORIES_TAB500", firstSeries.RowLabel
    Assert.AreEqual "LABEL_COL_1_TAB500", firstSeries.ColumnLabel

    Assert.AreEqual "PERC_COL_1_TAB500", secondSeries.SeriesName
    Assert.AreEqual "ROW_CATEGORIES_TAB500", secondSeries.RowLabel
    Assert.AreEqual "LABEL_COL_1_TAB500", secondSeries.ColumnLabel
    Assert.AreEqual "secondary", secondSeries.SeriesPosition
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestBuildGraphCreatesExpectedSeries"
End Sub

'@TestMethod("TimeSeriesGraphBuilder")
Private Sub TestBuildAllReturnsCharts()
    Dim cache As IGraphSpecsCache
    Dim builder As ITimeSeriesGraphBuilder
    Dim linelist As ILinelistSpecs
    Dim specsStub As GraphTablesSpecsStub
    Dim factoryStub As TimeSeriesSpecsFactoryStub
    Dim charts As BetterArray

    On Error GoTo Fail

    Set cache = GraphSpecsCache.Create(GraphList)
    Set linelist = New TableSpecsLinelistStub

    Set specsStub = New GraphTablesSpecsStub
    specsStub.Configure CInt(TABLE_TYPE_TIME_SERIES), "TAB500"

    Set factoryStub = New TimeSeriesSpecsFactoryStub
    factoryStub.AddSpec "SeriesA", specsStub.Self
    factoryStub.AddSpec "SeriesB", specsStub.Self

    Set builder = TimeSeriesGraphBuilder.Create(cache, SeriesList, TitleList, linelist, , factoryStub.Self)
    Set charts = builder.BuildAll

    Assert.AreEqual 1&, charts.Length
    Assert.AreEqual "GraphTS1", charts.Item(charts.LowerBound).GraphId
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestBuildAllReturnsCharts"
End Sub

'@section Helpers
'===============================================================================

Private Function SeedGraphList() As ListObject
    Dim sh As Worksheet
    Dim dataRange As Range

    Set sh = EnsureWorksheet(GRAPH_SHEET)
    sh.Cells.Clear

    sh.Range("A1:G1").Value = Array("graph id", "series id", "axis", "type", "percentages", "choices", "label")
    sh.Range("A2:G3").Value = Array( _
        Array("GraphTS1", "SeriesA", "primary", "line", "", "ChoiceA", ""), _
        Array("GraphTS1", "SeriesB", "secondary", "column", "percentages", "ChoiceA", "PrefixB"))

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
    sh.Range("A2:C2").Value = Array("Total admissions", "", "GraphTS1")

    If sh.ListObjects.Count > 0 Then sh.ListObjects(1).Delete
    Set dataRange = sh.Range("A1").CurrentRegion
    Set SeedTitleList = sh.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
End Function

Private Sub SeedCategoryNames(ByVal tableId As String)
    Dim sh As Worksheet

    Set sh = SeriesList.Parent

    sh.Range("F2").Value = "ChoiceA"
    sh.Range("F3").Value = "ChoiceB"

    RemoveName "COLUMN_CATEGORIES_" & tableId
    RemoveName "LABEL_COL_1_" & tableId
    RemoveName "LABEL_COL_2_" & tableId

    With sh.Parent
        .Names.Add Name:="COLUMN_CATEGORIES_" & tableId, RefersTo:=sh.Range("F2:F3")
        .Names.Add Name:="LABEL_COL_1_" & tableId, RefersTo:=sh.Range("F2")
        .Names.Add Name:="LABEL_COL_2_" & tableId, RefersTo:=sh.Range("F3")
    End With
End Sub

Private Sub RemoveName(ByVal nameText As String)
    Dim wb As Workbook

    Set wb = SeriesList.Parent.Parent

    On Error Resume Next
        wb.Names(nameText).Delete
    Err.Clear
    On Error GoTo 0
End Sub
