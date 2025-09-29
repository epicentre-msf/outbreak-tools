Attribute VB_Name = "TestGraphs"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulNames
'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private GeneralGraph As IGraphs

Private Const GRAPHOUT As String = "GraphOut"
Private Const SERIESNAME As String = "GraphSeriesData"
Private Const CATEGORYNAME As String = "GraphCategoryData"
Private Const LABELNAME As String = "GraphLabelValue"
Private Const SERIESSECONDARY As String = "GraphSeriesSecondary"

'@section Helpers
'===============================================================================

Private Function GraphSheet() As Worksheet
    Set GraphSheet = ThisWorkbook.Worksheets(GRAPHOUT)
End Function

Private Sub AssignNamedRange(ByVal hostSheet As Worksheet, ByVal nameText As String, ByVal target As Range)
    On Error Resume Next
        hostSheet.Names(nameText).Delete
    On Error GoTo 0

    hostSheet.Names.Add Name:=nameText, RefersTo:="='" & hostSheet.Name & "'!" & target.Address(True, True)
End Sub

Private Sub ResetGraphSheet()
    Dim hostSheet As Worksheet

    Set hostSheet = GraphSheet()
    ClearWorksheet hostSheet

    On Error Resume Next
        Do While hostSheet.ChartObjects.Count > 0
            hostSheet.ChartObjects(1).Delete
        Loop
    On Error GoTo 0
End Sub

Private Function BuildGraph(Optional ByVal graphName As String = "General Graph") As IGraphs
    Dim hostSheet As Worksheet

    Set hostSheet = GraphSheet()
    Set BuildGraph = Graphs.Create(hostSheet, hostSheet.Cells(5, 5), graphName)
End Function


'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Dim wb As Workbook

    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set wb = ThisWorkbook
    BusyApp

    EnsureWorksheet GRAPHOUT, wb
    ResetGraphSheet
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing

    ResetGraphSheet
    DeleteWorksheet GRAPHOUT
End Sub


'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    ResetGraphSheet
    Set GeneralGraph = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Graphs")
Private Sub TestCreate()

    Dim graphInstance As IGraphs
    Dim targetCell As Range
    Dim hostSheet As Worksheet

    On Error GoTo Fail

    Set hostSheet = GraphSheet()
    Set targetCell = hostSheet.Cells(5, 5)
    Set graphInstance = Graphs.Create(hostSheet, targetCell, "TestGraph")
    Assert.IsTrue (TypeName(graphInstance) = "Graphs"), "Create should return an IGraphs instance"
    Set GeneralGraph = BuildGraph()

    On Error Resume Next
        Set targetCell = Nothing
        Err.Clear
        '@Ignore AssignmentNotUsed
        Set graphInstance = Graphs.Create(hostSheet, targetCell)
        Assert.IsTrue (Err.Number = ProjectError.ObjectNotInitialized), "Create should raise for missing range"

        Err.Clear
        Set targetCell = GraphSheet().Cells(5, 1)
        Set hostSheet = Nothing
        '@Ignore AssignmentNotUsed
        Set graphInstance = Graphs.Create(hostSheet, targetCell)
        Assert.IsTrue (Err.Number = ProjectError.ObjectNotInitialized), "Create should raise for missing worksheet"
    On Error GoTo Fail

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestCreate"
End Sub



'@TestMethod("Graphs")
Private Sub TestAddCreatesChartObject()

    Dim hostSheet As Worksheet

    On Error GoTo Fail

    Set hostSheet = GraphSheet()
    Set GeneralGraph = BuildGraph("Add Graph")

    GeneralGraph.Add

    Assert.IsTrue (hostSheet.ChartObjects.Count = 1), "Add should create a chart object on the worksheet"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddCreatesChartObject"
End Sub


'@TestMethod("Graphs")
Private Sub TestAddSeriesAddsChartSeries()

    Dim hostSheet As Worksheet
    Dim chartObject As ChartObject

    On Error GoTo Fail

    Set hostSheet = GraphSheet()
    WriteColumn hostSheet.Range("A1"), 10, 20, 30
    AssignNamedRange hostSheet, SERIESNAME, hostSheet.Range("A1:A3")

    Set GeneralGraph = BuildGraph("Series Graph")
    GeneralGraph.AddSeries SERIESNAME, "bar"

    Assert.IsTrue (hostSheet.ChartObjects.Count = 1), "AddSeries should ensure a chart exists"

    Set chartObject = hostSheet.ChartObjects(1)
    Assert.IsTrue (chartObject.Chart.SeriesCollection.Count = 1), "AddSeries should append a single series"
    Assert.IsTrue (chartObject.Chart.SeriesCollection(1).ChartType = xlColumnClustered), "AddSeries should respect the requested chart type"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddSeriesAddsChartSeries"
End Sub


'@TestMethod("Graphs")
Private Sub TestAddSeriesHandlesMultipleSeries()

    Dim hostSheet As Worksheet
    Dim chartObject As ChartObject

    On Error GoTo Fail

    Set hostSheet = GraphSheet()
    WriteColumn hostSheet.Range("A1"), "Cat A", "Cat B", "Cat C"
    WriteColumn hostSheet.Range("B1"), 5, 10, 15
    WriteColumn hostSheet.Range("C1"), 1, 4, 9
    hostSheet.Range("D1").Value = "Primary Series"

    AssignNamedRange hostSheet, CATEGORYNAME, hostSheet.Range("A1:A3")
    AssignNamedRange hostSheet, SERIESNAME, hostSheet.Range("B1:B3")
    AssignNamedRange hostSheet, SERIESSECONDARY, hostSheet.Range("C1:C3")
    AssignNamedRange hostSheet, LABELNAME, hostSheet.Range("D1")

    Set GeneralGraph = BuildGraph("Multiple Series Graph")
    GeneralGraph.AddSeries SERIESNAME, "bar"
    GeneralGraph.AddSeries SERIESSECONDARY, "line", "right"
    GeneralGraph.AddLabels CATEGORYNAME, LABELNAME

    Set chartObject = hostSheet.ChartObjects(1)

    Assert.IsTrue (chartObject.Chart.SeriesCollection.Count = 2), "AddSeries should append each requested series"
    Assert.IsTrue (chartObject.Chart.SeriesCollection(1).AxisGroup = xlPrimary), "First series should stay on the primary axis"
    Assert.IsTrue (chartObject.Chart.SeriesCollection(2).AxisGroup = xlSecondary), "Secondary series should move to the secondary axis"
    Assert.IsTrue (chartObject.Chart.HasAxis(xlValue, xlSecondary)), "Adding a secondary axis should enable it on the chart"
    Assert.IsTrue (chartObject.Chart.SeriesCollection(2).ChartType = xlLineMarkers), "Secondary series should respect the requested line marker type"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddSeriesHandlesMultipleSeries"
End Sub


'@TestMethod("Graphs")
Private Sub TestAddLabelsAnnotatesSeries()

    Dim hostSheet As Worksheet
    Dim chartObject As ChartObject

    On Error GoTo Fail

    Set hostSheet = GraphSheet()
    WriteColumn hostSheet.Range("A1"), "Cat 1", "Cat 2", "Cat 3"
    WriteColumn hostSheet.Range("B1"), 5, 15, 25
    hostSheet.Range("C1").Value = "Confirmed Cases"

    AssignNamedRange hostSheet, CATEGORYNAME, hostSheet.Range("A1:A3")
    AssignNamedRange hostSheet, SERIESNAME, hostSheet.Range("B1:B3")
    AssignNamedRange hostSheet, LABELNAME, hostSheet.Range("C1")

    Set GeneralGraph = BuildGraph("Labels Graph")
    GeneralGraph.AddSeries SERIESNAME, "bar"
    GeneralGraph.AddLabels CATEGORYNAME, LABELNAME, "FY24"

    Set chartObject = hostSheet.ChartObjects(1)

    Assert.IsTrue chartObject.Chart.SeriesCollection(1).HasDataLabels, "AddLabels should enable data labels"
    Assert.IsTrue (chartObject.Chart.SeriesCollection(1).Name = "FY24 - " & hostSheet.Range(LABELNAME).Value), "AddLabels should apply the requested prefix and label"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddLabelsAnnotatesSeries"
End Sub


'@TestMethod("Graphs")
Private Sub TestFormatAdjustsChartLayout()

    Dim hostSheet As Worksheet
    Dim chartObject As ChartObject

    On Error GoTo Fail

    Set hostSheet = GraphSheet()
    WriteColumn hostSheet.Range("A1"), 100, 200, 300
    AssignNamedRange hostSheet, SERIESNAME, hostSheet.Range("A1:A3")

    Set GeneralGraph = BuildGraph("Format Graph")
    GeneralGraph.AddSeries SERIESNAME, "line"
    GeneralGraph.Format valuesTitle:="Values", catTitle:="Dates", plotTitle:="Case Trend", scope:=GraphScopeTimeSeries, heightFactor:=2

    Set chartObject = hostSheet.ChartObjects(1)

    Assert.AreEqual "Values", chartObject.Chart.Axes(xlValue, xlPrimary).AxisTitle.Caption, "Format should set the values axis title"
    Assert.AreEqual "Dates", chartObject.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Caption, "Format should set the category axis title"
    Assert.IsTrue chartObject.Chart.HasTitle, "Format should add a chart title when provided"
    Assert.AreEqual "Case Trend", chartObject.Chart.ChartTitle.Caption, "Format should apply the provided chart title"
    Assert.IsTrue chartObject.Width > 488, "Format should expand time series charts"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestFormatAdjustsChartLayout"
End Sub
