Attribute VB_Name = "TestGraphs"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulNames
'@TestModule
'@Folder("Tests")

'@description
'Validates the Graphs class, which wraps Excel ChartObject creation and
'series management for analysis worksheets. Tests cover factory validation
'(Create with valid arguments, rejection of Nothing worksheet and Nothing
'range), chart creation via Add, single and multiple series attachment via
'AddSeries (including primary and secondary axis placement and chart type
'mapping), category and legend label assignment via AddLabels, and layout
'formatting via Format (axis titles, chart title, scope-based sizing). The
'fixture uses a dedicated "GraphOut" worksheet that is cleared before each
'test and torn down in ModuleCleanup. Named ranges are created as
'worksheet-level names to provide series, category, and label data sources.
'@depends Graphs, IGraphs, CustomTest, TestHelpers

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

'@sub-title Return the dedicated graph output worksheet.
'@return Worksheet. The "GraphOut" worksheet in ThisWorkbook.
Private Function GraphSheet() As Worksheet
    Set GraphSheet = ThisWorkbook.Worksheets(GRAPHOUT)
End Function

'@sub-title Create or replace a worksheet-level named range.
'@details
'Deletes any existing name with the same text on the host worksheet,
'then creates a new named range pointing to the supplied target range.
'Used to wire up series, category, and label data sources for chart
'construction in each test.
'@param hostSheet Worksheet. The worksheet owning the named range.
'@param nameText String. The name to assign.
'@param target Range. The cell range the name should reference.
Private Sub AssignNamedRange(ByVal hostSheet As Worksheet, ByVal nameText As String, ByVal target As Range)
    On Error Resume Next
        hostSheet.Names(nameText).Delete
    On Error GoTo 0

    hostSheet.Names.Add Name:=nameText, RefersTo:="='" & hostSheet.Name & "'!" & target.Address(True, True)
End Sub

'@sub-title Clear the graph output worksheet and remove all chart objects.
'@details
'Calls ClearWorksheet to wipe cell data and named ranges, then iterates
'and deletes every ChartObject on the sheet so the next test starts with
'a pristine canvas.
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

'@sub-title Build a default IGraphs instance anchored at cell E5 on the graph sheet.
'@details
'Creates a Graphs object using the "GraphOut" worksheet and cell E5 as
'the chart anchor position. The optional graphName parameter controls the
'chart title label used during creation.
'@param graphName Optional String. Display name for the chart. Defaults to "General Graph".
'@return IGraphs. A fully initialised Graphs instance ready for Add/AddSeries.
Private Function BuildGraph(Optional ByVal graphName As String = "General Graph") As IGraphs
    Dim hostSheet As Worksheet

    Set hostSheet = GraphSheet()
    Set BuildGraph = Graphs.Create(hostSheet, hostSheet.Cells(5, 5), graphName)
End Function


'@section Module lifecycle
'===============================================================================

'@sub-title Initialise the test module before any tests run.
'@details
'Creates the Rubberduck AssertClass and FakesProvider objects, suppresses
'screen updates via BusyApp, ensures the "GraphOut" worksheet exists,
'and resets it to a clean state.
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

'@sub-title Tear down the module after all tests complete.
'@details
'Releases the Rubberduck assertion and fakes objects, resets the graph
'output worksheet, and deletes it to leave the workbook clean.
'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing

    ResetGraphSheet
    DeleteWorksheet GRAPHOUT
End Sub

'@sub-title Reset state before each individual test.
'@details
'Suppresses screen updates, clears the graph output worksheet and all
'its chart objects, and releases any previously held GeneralGraph
'reference so each test starts with a pristine environment.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    ResetGraphSheet
    Set GeneralGraph = Nothing
End Sub

'@section Tests
'===============================================================================

'@sub-title Verify Create returns a valid IGraphs instance and rejects invalid arguments.
'@details
'Acts by calling Graphs.Create with a valid worksheet and target cell.
'Asserts that the returned object has TypeName "Graphs". Then verifies
'that Create raises ProjectError.ObjectNotInitialized when the target
'range is Nothing and when the worksheet is Nothing, confirming both
'guard clauses in the factory method.
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



'@sub-title Verify Add creates a single ChartObject on the worksheet.
'@details
'Arranges a fresh IGraphs instance via BuildGraph. Acts by calling Add.
'Asserts that the host worksheet's ChartObjects count is exactly 1,
'confirming that Add creates an empty chart at the anchor position.
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


'@sub-title Verify AddSeries appends a single chart series with the correct type.
'@details
'Arranges a named range containing three numeric values and builds a
'graph instance. Acts by calling AddSeries with "bar" as the chart type.
'Asserts that a chart was created (lazy Add), that exactly one series
'exists, and that its ChartType matches xlColumnClustered, confirming
'the "bar" type mapping logic.
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


'@sub-title Verify AddSeries handles multiple series with primary and secondary axes.
'@details
'Arranges category labels, two data series, and a legend label on the
'graph sheet with corresponding named ranges. Acts by adding two series:
'a "bar" on the primary axis and a "line" on the secondary ("right")
'axis, then calling AddLabels. Asserts that two series exist, that
'their axis groups are correct (primary vs secondary), that the
'secondary axis is enabled on the chart, and that the secondary series
'renders as xlLineMarkers, confirming multi-series and dual-axis support.
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


'@sub-title Verify AddLabels enables data labels and applies the legend entry prefix.
'@details
'Arranges category labels, series data, and a legend label cell on the
'graph sheet. Acts by adding a "bar" series and calling AddLabels with
'a "FY24" prefix. Asserts that the series has data labels enabled and
'that the series Name combines the prefix and the label cell value,
'confirming that AddLabels wires category data, enables labels, and
'constructs the legend entry correctly.
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


'@sub-title Verify Format sets axis titles, chart title, and applies scope-based sizing.
'@details
'Arranges a "line" series and calls Format with axis titles, a chart
'title, time series scope, and a height factor of 2. Asserts that the
'value and category axis titles match the supplied strings, that the
'chart has a title matching "Case Trend", and that the chart width
'exceeds the default 488px threshold for time series scope, confirming
'that Format applies all layout adjustments correctly.
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
