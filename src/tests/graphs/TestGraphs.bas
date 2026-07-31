Attribute VB_Name = "TestGraphs"
Attribute VB_Description = "Tests for Graphs class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for Graphs class")

'@description
'Drives Graphs, which wraps one Excel ChartObject: the frame, the series, the
'labels and the scope-dependent layout. This module had never been imported,
'compiled or run before 2026-07-31. The version it replaces built its assertion
'object with CreateObject("Rubberduck.AssertClass"), carried Option Private
'Module and kept every test procedure Private, so registering it as it stood
'would have stopped the whole project compiling.
'
'THE FIXTURE IS ONE HIDDEN WORKSHEET AND A HANDFUL OF NAMED RANGES
'-------------------------------------------------------------------------------
'Graphs reads every series, category and label off a named range of its own
'worksheet, so the fixture writes three short columns and names them. A second
'worksheet exists for one test alone: the factory refuses a position range that
'sits somewhere other than the worksheet the values come from.
'
'THE SHEET RESET IS BOUNDED AND IT SWEEPS THE CHARTS
'-------------------------------------------------------------------------------
'EnsureWorksheet(clearSheet:=True) runs Cells.Clear over the whole
'1,048,576 by 16,384 sheet, which took a green analyses run past the runner's
'cap. Every fixture here lives inside the first thirty rows. A chart left behind
'would be counted by the next test, so the reset drops the chart objects too,
'and the names are swept through RefersToRange, which answers the owning
'worksheet whatever the RefersTo text spells.
'@depends Graphs, Checking, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "GraphsFixture"
Private Const OTHER_SHEET As String = "GraphsOther"

' The named ranges the fixture writes.
Private Const SERIES_NAME As String = "GraphSeriesData"
Private Const SECOND_SERIES_NAME As String = "GraphSeriesSecondary"
Private Const CATEGORY_NAME As String = "GraphCategoryData"
Private Const LABEL_NAME As String = "GraphLabelValue"
Private Const TITLE_NAME As String = "GraphTitleValue"

' A name nothing ever creates.
Private Const ABSENT_NAME As String = "GraphNameNobodyCreates"

' What the fixture writes into the label cell and the title cell.
Private Const LABEL_TEXT As String = "Confirmed Cases"
Private Const TITLE_TEXT As String = "Case count"

' The frame a chart is born with, and the coefficient a time series applies.
Private Const BASE_WIDTH As Double = 488
Private Const TIME_SERIES_WIDTH As Double = 854

Private Assert As CustomTest

'@section Fixture helpers
'===============================================================================

'@sub-title Delete every workbook name that points at one worksheet.
'@details
'A worksheet-scoped name is held in the workbook collection too, so one walk
'catches both. The owning worksheet is read off RefersToRange, which holds
'whatever the sheet is called. A name carrying a value has no range and raises,
'so the read is guarded.
'@param sh Worksheet. The worksheet whose names go.
Private Sub ClearSheetNames(ByVal sh As Worksheet)
    Dim idx As Long
    Dim nameItem As Name
    Dim owner As Worksheet

    For idx = ThisWorkbook.Names.Count To 1 Step -1
        Set nameItem = ThisWorkbook.Names(idx)
        Set owner = Nothing

        On Error Resume Next
        Set owner = nameItem.RefersToRange.Worksheet
        Err.Clear
        On Error GoTo 0

        If Not owner Is Nothing Then
            If StrComp(owner.Name, sh.Name, vbTextCompare) = 0 Then nameItem.Delete
        End If
    Next idx
End Sub

'@sub-title Drop every chart on a fixture sheet.
'@param sh Worksheet. The fixture worksheet.
Private Sub RemoveCharts(ByVal sh As Worksheet)
    Dim idx As Long

    For idx = sh.ChartObjects.Count To 1 Step -1
        sh.ChartObjects(idx).Delete
    Next idx
End Sub

'@sub-title Empty a fixture worksheet without clearing the whole sheet.
'@param sheetName String. The worksheet to reset.
'@return Worksheet. The empty worksheet.
Private Function ResetFixtureSheet(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    ' The sheet stays visible. A chart is drawn on it, and some Excel hosts
    ' place a ChartObject through the selection.
    Set sh = EnsureWorksheet(sheetName, clearSheet:=False)
    RemoveCharts sh
    ClearSheetNames sh
    sh.Range("A1:T30").Clear

    Set ResetFixtureSheet = sh
End Function

'@sub-title Name a range on its own worksheet.
'@param sh Worksheet. The worksheet owning the name.
'@param nameText String. The name to give.
'@param target Range. The range the name points at.
Private Sub AssignName(ByVal sh As Worksheet, ByVal nameText As String, _
                       ByVal target As Range)
    target.Name = nameText
End Sub

'@sub-title Build the fixture worksheet and its named ranges.
'@details
'Three categories, two value columns, one legend label cell and one axis title
'cell. Every test that draws a chart starts from this.
'@return Worksheet. The fixture worksheet.
Private Function BuildFixture() As Worksheet
    Dim sh As Worksheet

    Set sh = ResetFixtureSheet(FIXTURE_SHEET)

    WriteColumn sh.Range("A1"), "Cat A", "Cat B", "Cat C"
    WriteColumn sh.Range("B1"), 10, 20, 30
    WriteColumn sh.Range("C1"), 1, 4, 9
    sh.Range("D1").Value = LABEL_TEXT
    sh.Range("E1").Value = TITLE_TEXT

    AssignName sh, CATEGORY_NAME, sh.Range("A1:A3")
    AssignName sh, SERIES_NAME, sh.Range("B1:B3")
    AssignName sh, SECOND_SERIES_NAME, sh.Range("C1:C3")
    AssignName sh, LABEL_NAME, sh.Range("D1")
    AssignName sh, TITLE_NAME, sh.Range("E1")

    Set BuildFixture = sh
End Function

'@sub-title Build a graph anchored at cell G5 of the fixture worksheet.
'@param sh Worksheet. The fixture worksheet.
'@param graphName Optional String. Display name of the chart.
'@return Graphs. A graph before Add.
Private Function BuildGraph(ByVal sh As Worksheet, _
                            Optional ByVal graphName As String = vbNullString) As Graphs
    Set BuildGraph = Graphs.Create(sh, sh.Cells(5, 7), graphName)
End Function

'@sub-title The one chart of the fixture worksheet.
'@param sh Worksheet. The fixture worksheet.
'@return ChartObject. The first chart, or Nothing when there is none.
Private Function FirstChart(ByVal sh As Worksheet) As ChartObject
    If sh.ChartObjects.Count = 0 Then Exit Function
    Set FirstChart = sh.ChartObjects(1)
End Function

'@sub-title A graph carrying one bar series, ready for Format.
'@param sh Worksheet. The fixture worksheet.
'@return Graphs. A graph with one series on it.
Private Function GraphWithOneSeries(ByVal sh As Worksheet) As Graphs
    Dim gr As Graphs

    Set gr = BuildGraph(sh)
    gr.Add
    gr.AddSeries SERIES_NAME, "bar"

    Set GraphWithOneSeries = gr
End Function

'@sub-title Whether two lengths are the same to within a point.
'@details
'A chart dimension is a Double that Excel stores as it was given. The tolerance
'covers the rounding a host may apply when it lays the frame out.
'@param expected Double. The length asked for.
'@param actual Double. The length the chart carries.
'@return Boolean. True when the two agree.
Private Function SameLength(ByVal expected As Double, ByVal actual As Double) As Boolean
    SameLength = (Abs(expected - actual) < 1)
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Build the assertion harness.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run. A Private lifecycle hook is the trap that has cost five
'modules a run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestGraphs"
End Sub

'@sub-title Print the results and drop every fixture this module made.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    DeleteWorksheets FIXTURE_SHEET, OTHER_SHEET

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updates before each test.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run. The first assertion of each test opens the checking, which
'picks up the titles set a line above it. Calling BeginTest here would open it
'with whatever titles are pending and file every result under the default label.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush the results of each test.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory validation
'===============================================================================

'@sub-title Verify Create answers a sealed instance carrying its name.
'@TestMethod("Graphs")
Public Sub TestCreateAnswersAnInstance()
    CustomTestSetTitles Assert, "Graphs", "TestCreateAnswersAnInstance"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs

    Set sh = BuildFixture()
    Set gr = BuildGraph(sh, "A chart")

    Assert.IsTrue (TypeName(gr) = "Graphs"), "Create answers a Graphs instance"
    Assert.AreEqual "A chart", gr.Name, "It carries the name it was given"
    Assert.AreEqual sh.Name, gr.Wksh.Name, "And the worksheet it was given"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateAnswersAnInstance", Err.Number, Err.Description
End Sub

'@sub-title Verify Create refuses a Nothing worksheet.
'@TestMethod("Graphs")
Public Sub TestCreateRejectsNothingWorksheet()
    CustomTestSetTitles Assert, "Graphs", "TestCreateRejectsNothingWorksheet"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim errNumber As Long

    Set sh = BuildFixture()

    On Error Resume Next
    Set gr = Graphs.Create(Nothing, sh.Cells(5, 7))
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (gr Is Nothing), "A Nothing worksheet gives nothing back"
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing worksheet raises ObjectNotInitialized"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingWorksheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Create refuses a Nothing position range.
'@TestMethod("Graphs")
Public Sub TestCreateRejectsNothingRange()
    CustomTestSetTitles Assert, "Graphs", "TestCreateRejectsNothingRange"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim errNumber As Long

    Set sh = BuildFixture()

    On Error Resume Next
    Set gr = Graphs.Create(sh, Nothing)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (gr Is Nothing), "A Nothing range gives nothing back"
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing range raises ObjectNotInitialized"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingRange", Err.Number, Err.Description
End Sub

'@sub-title Verify Create refuses a position range on another worksheet.
'@details
'The chart is drawn on the parent of the position range and every named range is
'read off the worksheet the class holds. Two different sheets would draw a chart
'in one place, feed it from another, and qualify its legend references with the
'wrong sheet name.
'@TestMethod("Graphs")
Public Sub TestCreateRejectsARangeOnAnotherWorksheet()
    CustomTestSetTitles Assert, "Graphs", "TestCreateRejectsARangeOnAnotherWorksheet"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim otherSh As Worksheet
    Dim gr As Graphs
    Dim errNumber As Long

    Set sh = BuildFixture()
    Set otherSh = ResetFixtureSheet(OTHER_SHEET)

    On Error Resume Next
    Set gr = Graphs.Create(sh, otherSh.Cells(5, 7))
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (gr Is Nothing), "A range on another worksheet gives nothing back"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A range on another worksheet raises InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsARangeOnAnotherWorksheet", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the worksheet cannot be swapped after creation.
'@TestMethod("Graphs")
Public Sub TestTheInstanceIsSealedAfterCreation()
    CustomTestSetTitles Assert, "Graphs", "TestTheInstanceIsSealedAfterCreation"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim otherSh As Worksheet
    Dim gr As Graphs
    Dim errNumber As Long

    Set sh = BuildFixture()
    Set otherSh = ResetFixtureSheet(OTHER_SHEET)
    Set gr = BuildGraph(sh)

    On Error Resume Next
    Set gr.Wksh = otherSh
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.SomethingWentWrong), errNumber, _
                    "Writing the worksheet after creation raises"
    Assert.AreEqual sh.Name, gr.Wksh.Name, "And the worksheet is the one Create was given"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheInstanceIsSealedAfterCreation", _
                         Err.Number, Err.Description
End Sub

'@section The chart frame
'===============================================================================

'@sub-title Verify Add draws one chart at the anchor cell.
'@details
'The chart used to be placed three times: once through the arguments of
'ChartObjects.Add and twice more by writing Left and Top again, each write
'moving the frame.
'@TestMethod("Graphs")
Public Sub TestAddDrawsOneChartAtTheAnchor()
    CustomTestSetTitles Assert, "Graphs", "TestAddDrawsOneChartAtTheAnchor"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject
    Dim anchor As Range

    Set sh = BuildFixture()
    Set anchor = sh.Cells(5, 7)
    Set gr = BuildGraph(sh)

    gr.Add

    Assert.AreEqual 1&, CLng(sh.ChartObjects.Count), "Add draws one chart"

    Set co = FirstChart(sh)
    Assert.IsTrue SameLength(anchor.Left, co.Left), "It sits at the left edge of its anchor"
    Assert.IsTrue SameLength(anchor.Top, co.Top), "And at the top edge of its anchor"
    Assert.IsTrue SameLength(BASE_WIDTH, co.width), "And it opens at the standard width"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddDrawsOneChartAtTheAnchor", Err.Number, Err.Description
End Sub

'@section Series
'===============================================================================

'@sub-title Verify AddSeries attaches one series carrying the values of its range.
'@details
'The values used to be written twice: SeriesCollection.Add takes a Source, which
'fills them, and the line below resolved the same name again and pushed the same
'data across. This test is what says the single write is enough.
'@TestMethod("Graphs")
Public Sub TestAddSeriesAttachesTheValuesOfItsRange()
    CustomTestSetTitles Assert, "Graphs", "TestAddSeriesAttachesTheValuesOfItsRange"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject
    Dim seriesValues As Variant

    Set sh = BuildFixture()
    Set gr = BuildGraph(sh)
    gr.AddSeries SERIES_NAME, "bar"

    Assert.AreEqual 1&, CLng(sh.ChartObjects.Count), "AddSeries draws the chart when there is none"

    Set co = FirstChart(sh)
    Assert.AreEqual 1&, CLng(co.Chart.SeriesCollection.Count), "One series is attached"
    Assert.AreEqual CLng(xlColumnClustered), CLng(co.Chart.SeriesCollection(1).chartType), _
                    "A bar series is drawn as clustered columns"

    seriesValues = co.Chart.SeriesCollection(1).Values
    Assert.AreEqual 3&, CLng(UBound(seriesValues) - LBound(seriesValues) + 1), _
                    "The series carries the three cells of its range"
    Assert.AreEqual 10, seriesValues(LBound(seriesValues)), "And their values"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddSeriesAttachesTheValuesOfItsRange", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a series asking for the right axis moves to the secondary one.
'@TestMethod("Graphs")
Public Sub TestASeriesOnTheRightMovesToTheSecondaryAxis()
    CustomTestSetTitles Assert, "Graphs", "TestASeriesOnTheRightMovesToTheSecondaryAxis"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = BuildGraph(sh)
    gr.Add
    gr.AddSeries SERIES_NAME, "bar"
    gr.AddSeries SECOND_SERIES_NAME, "line", "right"

    Set co = FirstChart(sh)

    Assert.AreEqual 2&, CLng(co.Chart.SeriesCollection.Count), "Both series are attached"
    Assert.AreEqual CLng(xlPrimary), CLng(co.Chart.SeriesCollection(1).AxisGroup), _
                    "The first series stays on the primary axis"
    Assert.AreEqual CLng(xlSecondary), CLng(co.Chart.SeriesCollection(2).AxisGroup), _
                    "The second moves to the secondary axis"
    Assert.IsTrue co.Chart.HasAxis(xlValue, xlSecondary), "Which the chart now carries"
    Assert.AreEqual CLng(xlLineMarkers), CLng(co.Chart.SeriesCollection(2).chartType), _
                    "A line series is drawn with markers"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASeriesOnTheRightMovesToTheSecondaryAxis", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a chart type nobody recognises is drawn as bars.
'@TestMethod("Graphs")
Public Sub TestAnUnknownChartTypeIsDrawnAsBars()
    CustomTestSetTitles Assert, "Graphs", "TestAnUnknownChartTypeIsDrawnAsBars"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = BuildGraph(sh)
    gr.AddSeries SERIES_NAME, "sunburst"

    Set co = FirstChart(sh)
    Assert.AreEqual CLng(xlColumnClustered), CLng(co.Chart.SeriesCollection(1).chartType), _
                    "An unknown chart type falls back to clustered columns"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnUnknownChartTypeIsDrawnAsBars", Err.Number, Err.Description
End Sub

'@sub-title Verify a series whose range is absent is left out and reported.
'@TestMethod("Graphs")
Public Sub TestASeriesWithNoRangeIsLeftOutAndReported()
    CustomTestSetTitles Assert, "Graphs", "TestASeriesWithNoRangeIsLeftOutAndReported"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs

    Set sh = BuildFixture()
    Set gr = BuildGraph(sh)
    gr.Add
    gr.AddSeries ABSENT_NAME, "bar"

    Assert.AreEqual 0&, CLng(FirstChart(sh).Chart.SeriesCollection.Count), _
                    "A name the worksheet does not carry attaches no series"
    Assert.IsTrue gr.HasCheckings, "And the miss is reported"
    Assert.IsTrue (Not gr.CheckingValues Is Nothing), "The report is handed over"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASeriesWithNoRangeIsLeftOutAndReported", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a failed series leaves the series before it labelled as it was.
'@details
'This is the fault that reached a user's chart. AddSeries logged the miss and
'returned without moving its index, and every caller runs AddSeries then
'AddLabels with no test between them, so the labels of the series that failed
'were written onto the series before it: its legend entry, its category axis and
'its data labels. The chart showed one series fewer than asked for and the last
'one carried the wrong name.
'@TestMethod("Graphs")
Public Sub TestAFailedSeriesLeavesThePreviousOneLabelledAsItWas()
    CustomTestSetTitles Assert, "Graphs", "TestAFailedSeriesLeavesThePreviousOneLabelledAsItWas"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = BuildGraph(sh)
    gr.Add

    gr.AddSeries SERIES_NAME, "bar"
    gr.AddLabels CATEGORY_NAME, LABEL_NAME, "First", prefixOnly:=True

    gr.AddSeries ABSENT_NAME, "bar"
    gr.AddLabels CATEGORY_NAME, LABEL_NAME, "Second", prefixOnly:=True

    Set co = FirstChart(sh)

    Assert.AreEqual 1&, CLng(co.Chart.SeriesCollection.Count), _
                    "Only the series that resolved is on the chart"
    Assert.AreEqual "First", co.Chart.SeriesCollection(1).Name, _
                    "And it keeps the legend entry of its own labels"
    Assert.IsTrue gr.HasCheckings, "Both the missing series and its skipped labels are reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFailedSeriesLeavesThePreviousOneLabelledAsItWas", _
                         Err.Number, Err.Description
End Sub

'@section Labels
'===============================================================================

'@sub-title Verify the prefix and the label cell make the legend entry.
'@TestMethod("Graphs")
Public Sub TestTheLegendEntryJoinsThePrefixAndTheLabelCell()
    CustomTestSetTitles Assert, "Graphs", "TestTheLegendEntryJoinsThePrefixAndTheLabelCell"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME, "FY24"

    Set co = FirstChart(sh)

    Assert.AreEqual "FY24 - " & LABEL_TEXT, co.Chart.SeriesCollection(1).Name, _
                    "The prefix and the label cell are joined by a dash"
    Assert.IsTrue co.Chart.SeriesCollection(1).HasDataLabels, "The data labels are on"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheLegendEntryJoinsThePrefixAndTheLabelCell", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a prefix-only label leaves the label cell out.
'@TestMethod("Graphs")
Public Sub TestAPrefixOnlyLabelIsTheWholeLegendEntry()
    CustomTestSetTitles Assert, "Graphs", "TestAPrefixOnlyLabelIsTheWholeLegendEntry"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME, "Week 12", prefixOnly:=True

    Assert.AreEqual "Week 12", FirstChart(sh).Chart.SeriesCollection(1).Name, _
                    "The prefix alone names the series"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAPrefixOnlyLabelIsTheWholeLegendEntry", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a dynamic label reads the cell it points at.
'@details
'A caller asking for dynamic labels gets a reference written into the series
'name, so a user editing that cell edits the legend.
'@TestMethod("Graphs")
Public Sub TestADynamicLabelReadsTheCellItPointsAt()
    CustomTestSetTitles Assert, "Graphs", "TestADynamicLabelReadsTheCellItPointsAt"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME, hardCodeLabels:=False

    Assert.AreEqual LABEL_TEXT, FirstChart(sh).Chart.SeriesCollection(1).Name, _
                    "The legend entry shows the text of the cell it references"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestADynamicLabelReadsTheCellItPointsAt", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify labels asked for before any series are reported.
'@TestMethod("Graphs")
Public Sub TestLabelsBeforeAnySeriesAreReported()
    CustomTestSetTitles Assert, "Graphs", "TestLabelsBeforeAnySeriesAreReported"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs

    Set sh = BuildFixture()
    Set gr = BuildGraph(sh)
    gr.Add
    gr.AddLabels CATEGORY_NAME, LABEL_NAME

    Assert.AreEqual 0&, CLng(FirstChart(sh).Chart.SeriesCollection.Count), _
                    "Labels on their own attach no series"
    Assert.IsTrue gr.HasCheckings, "And they are reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLabelsBeforeAnySeriesAreReported", _
                         Err.Number, Err.Description
End Sub

'@section Layout
'===============================================================================

'@sub-title Verify a standard chart keeps the base width.
'@TestMethod("Graphs")
Public Sub TestAStandardChartKeepsTheBaseWidth()
    CustomTestSetTitles Assert, "Graphs", "TestAStandardChartKeepsTheBaseWidth"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.Format valuesTitle:="Values", catTitle:="Dates", scope:=GraphScopeNormal

    Set co = FirstChart(sh)
    Assert.IsTrue SameLength(BASE_WIDTH, co.width), "A standard chart is 488 points wide"
    Assert.IsTrue (co.height > 0), "And it has a height"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAStandardChartKeepsTheBaseWidth", Err.Number, Err.Description
End Sub

'@sub-title Verify a time series chart is stretched sideways.
'@details
'The width coefficient is 1.75 and it was held in a Long, so it was stored as 2
'and every time series chart came out 14 per cent wider than the code asks for.
'@TestMethod("Graphs")
Public Sub TestATimeSeriesChartIsStretchedSideways()
    CustomTestSetTitles Assert, "Graphs", "TestATimeSeriesChartIsStretchedSideways"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.Format valuesTitle:="Values", catTitle:="Dates", scope:=GraphScopeTimeSeries

    Set co = FirstChart(sh)
    Assert.IsTrue SameLength(TIME_SERIES_WIDTH, co.width), _
                  "A time series chart is 854 points wide, which is 488 times 1.75"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestATimeSeriesChartIsStretchedSideways", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the height of a chart never collapses, whatever the geo count.
'@details
'The height coefficient is (geo count + 1) times 0.08 and it was held in a Long
'too. Five geographic units gave 0.48, which rounded down to zero, and the chart
'came out with no height at all. The floor keeps a small chart at its normal
'size and lets a large one grow. Twenty Format calls on one chart cost far less
'than twenty charts.
'@TestMethod("Graphs")
Public Sub TestTheChartHeightNeverCollapses()
    CustomTestSetTitles Assert, "Graphs", "TestTheChartHeightNeverCollapses"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject
    Dim geoCount As Long
    Dim smallestHeight As Double
    Dim heightAtOne As Double
    Dim heightAtTwenty As Double

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME

    Set co = FirstChart(sh)
    smallestHeight = 100000

    For geoCount = 1 To 20
        gr.Format valuesTitle:="Values", catTitle:="Units", _
                  scope:=GraphScopeSpatial, heightFactor:=geoCount

        If co.height < smallestHeight Then smallestHeight = co.height
        If geoCount = 1 Then heightAtOne = co.height
        If geoCount = 20 Then heightAtTwenty = co.height
    Next geoCount

    Assert.IsTrue (smallestHeight > 0), _
                  "Every geographic unit count from 1 to 20 gives a chart with a height"
    Assert.IsTrue (heightAtTwenty > heightAtOne), _
                  "And twenty units give a taller chart than one"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheChartHeightNeverCollapses", Err.Number, Err.Description
End Sub

'@sub-title Verify a spatial chart reverses its categories and moves its legend.
'@TestMethod("Graphs")
Public Sub TestASpatialChartReversesItsCategories()
    CustomTestSetTitles Assert, "Graphs", "TestASpatialChartReversesItsCategories"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.Format valuesTitle:="Values", catTitle:="Units", scope:=GraphScopeSpatial

    Set co = FirstChart(sh)

    Assert.IsTrue co.Chart.Axes(xlCategory, xlPrimary).ReversePlotOrder, _
                  "A spatial chart reads from the largest unit downwards"
    Assert.AreEqual CLng(xlLegendPositionBottom), CLng(co.Chart.Legend.Position), _
                    "And its legend sits under it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASpatialChartReversesItsCategories", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a standard chart keeps its categories in order.
'@TestMethod("Graphs")
Public Sub TestAStandardChartKeepsItsCategoriesInOrder()
    CustomTestSetTitles Assert, "Graphs", "TestAStandardChartKeepsItsCategoriesInOrder"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.Format valuesTitle:="Values", catTitle:="Dates", scope:=GraphScopeNormal

    Set co = FirstChart(sh)

    Assert.IsTrue (Not co.Chart.Axes(xlCategory, xlPrimary).ReversePlotOrder), _
                  "A standard chart reads left to right"
    Assert.AreEqual CLng(xlLegendPositionTop), CLng(co.Chart.Legend.Position), _
                    "And its legend sits above it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAStandardChartKeepsItsCategoriesInOrder", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a scope this class does not know takes the standard layout.
'@details
'AnalysisOutput maps its four analysis scopes onto these three before it calls,
'and AnaTabIds does the same on the export path. A number outside the three used
'to fall through every branch and take the standard layout silently.
'@TestMethod("Graphs")
Public Sub TestAnUnknownScopeTakesTheStandardLayoutAndIsReported()
    CustomTestSetTitles Assert, "Graphs", "TestAnUnknownScopeTakesTheStandardLayoutAndIsReported"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.Format valuesTitle:="Values", catTitle:="Dates", scope:=99

    Set co = FirstChart(sh)

    Assert.IsTrue SameLength(BASE_WIDTH, co.width), "An unknown scope is laid out as a standard chart"
    Assert.IsTrue gr.HasCheckings, "And it is reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnUnknownScopeTakesTheStandardLayoutAndIsReported", _
                         Err.Number, Err.Description
End Sub

'@section Titles
'===============================================================================

'@sub-title Verify a hardcoded title is written as it stands.
'@TestMethod("Graphs")
Public Sub TestHardcodedTitlesAreWrittenAsTheyStand()
    CustomTestSetTitles Assert, "Graphs", "TestHardcodedTitlesAreWrittenAsTheyStand"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.Format valuesTitle:="Values", catTitle:="Dates", plotTitle:="Case Trend"

    Set co = FirstChart(sh)

    Assert.AreEqual "Values", co.Chart.Axes(xlValue, xlPrimary).AxisTitle.Caption, _
                    "The value axis carries the title it was given"
    Assert.AreEqual "Dates", co.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Caption, _
                    "So does the category axis"
    Assert.IsTrue co.Chart.HasTitle, "The chart carries a title"
    Assert.AreEqual "Case Trend", co.Chart.ChartTitle.Caption, "Which is the one it was given"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHardcodedTitlesAreWrittenAsTheyStand", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a dynamic title shows the text of the cell it names.
'@details
'Caption takes text and Formula takes a reference. The class used to assign
'"= 'Analysis'!$C$7" to Caption, which prints that string on the chart. This test
'is what settles it: the caption reads back as the cell's text.
'@TestMethod("Graphs")
Public Sub TestADynamicTitleShowsTheTextOfItsCell()
    CustomTestSetTitles Assert, "Graphs", "TestADynamicTitleShowsTheTextOfItsCell"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.Format valuesTitle:=TITLE_NAME, catTitle:="Dates", plotTitle:=TITLE_NAME, _
              hardCodeLabels:=False

    Set co = FirstChart(sh)

    Assert.AreEqual TITLE_TEXT, co.Chart.Axes(xlValue, xlPrimary).AxisTitle.Caption, _
                    "The value axis title shows the text of the cell it names"
    Assert.AreEqual TITLE_TEXT, co.Chart.ChartTitle.Caption, _
                    "And so does the chart title"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestADynamicTitleShowsTheTextOfItsCell", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify both title properties bind a title to the cell it names.
'@details
'The audit asked for a look at a real chart before anything here moved, and
'nobody had taken it. This is that look, taken by the harness and measured on
'2026-07-31. It expected Caption to print "= 'Analysis'!$C$7" on the chart as
'text, because Caption is documented as a text property and Formula as the one
'that takes a reference. On this host Caption parses the string and binds the
'title to the cell, so both properties answer the same text and no axis title in
'the product was ever showing a formula. The class writes Formula, which is what
'the object model gives for a reference, and this test is what says the two
'agree.
'@TestMethod("Graphs")
Public Sub TestBothTitlePropertiesBindTheTitleToItsCell()
    CustomTestSetTitles Assert, "Graphs", "TestBothTitlePropertiesBindTheTitleToItsCell"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim ax As Axis
    Dim cellAddress As String

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.Format valuesTitle:="Values", catTitle:="Dates"

    Set ax = FirstChart(sh).Chart.Axes(xlValue, xlPrimary)
    cellAddress = "'" & sh.Name & "'!" & sh.Range(TITLE_NAME).Address

    ax.AxisTitle.Caption = "= " & cellAddress
    Assert.AreEqual TITLE_TEXT, ax.AxisTitle.Caption, _
                    "Caption parses a formula string and shows the text of the cell"

    ax.AxisTitle.Caption = "A literal title"
    ax.AxisTitle.Formula = "=" & cellAddress
    Assert.AreEqual TITLE_TEXT, ax.AxisTitle.Caption, _
                    "Formula binds the title to the same cell and reads the same text"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBothTitlePropertiesBindTheTitleToItsCell", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a dynamic title naming an absent range is reported.
'@details
'AnalysisOutput passes "LABEL_ROW_CATEGORIES_" and the table identifier as the
'category title with hardCodeLabels False, so a table whose label range was
'never created used to get that raw string printed as its axis title with
'nothing said about it.
'@TestMethod("Graphs")
Public Sub TestADynamicTitleWithNoRangeIsReported()
    CustomTestSetTitles Assert, "Graphs", "TestADynamicTitleWithNoRangeIsReported"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = GraphWithOneSeries(sh)
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.Format valuesTitle:=ABSENT_NAME, catTitle:="Dates", hardCodeLabels:=False

    Set co = FirstChart(sh)

    Assert.AreEqual ABSENT_NAME, co.Chart.Axes(xlValue, xlPrimary).AxisTitle.Caption, _
                    "The name is written as it stands when the worksheet does not carry it"
    Assert.IsTrue gr.HasCheckings, "And the caller is told the reference was missed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestADynamicTitleWithNoRangeIsReported", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the secondary axis is scaled for fractions.
'@TestMethod("Graphs")
Public Sub TestTheSecondaryAxisIsScaledForFractions()
    CustomTestSetTitles Assert, "Graphs", "TestTheSecondaryAxisIsScaledForFractions"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim co As ChartObject

    Set sh = BuildFixture()
    Set gr = BuildGraph(sh)
    gr.Add
    gr.AddSeries SERIES_NAME, "bar"
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.AddSeries SECOND_SERIES_NAME, "point", "right"
    gr.AddLabels CATEGORY_NAME, LABEL_NAME
    gr.Format valuesTitle:="Values", catTitle:="Dates"

    Set co = FirstChart(sh)

    Assert.AreEqual 1, co.Chart.Axes(xlValue, xlSecondary).MaximumScale, _
                    "The secondary axis runs up to one, which is what a fraction needs"
    Assert.AreEqual 0.1, co.Chart.Axes(xlValue, xlSecondary).MajorUnit, _
                    "With a tick every tenth"
    Assert.AreEqual "%", co.Chart.Axes(xlValue, xlSecondary).AxisTitle.Caption, _
                    "And a per cent sign as its title"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSecondaryAxisIsScaledForFractions", _
                         Err.Number, Err.Description
End Sub

'@section The empty chart
'===============================================================================

'@sub-title Verify a chart that ends with no series has its frame removed.
'@details
'A chart with zero series has no axes, and Format opens by reading them, so it
'raised 1004. Every empty-chart path in GraphSpecs ended here, and the two chart
'loops of AnalysisOutput had nothing above them to catch it. An empty bordered
'frame on the analysis sheet reads as a broken chart, and the export would draw
'it again, so the frame goes.
'@TestMethod("Graphs")
Public Sub TestAChartWithNoSeriesLosesItsFrame()
    CustomTestSetTitles Assert, "Graphs", "TestAChartWithNoSeriesLosesItsFrame"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim errNumber As Long

    Set sh = BuildFixture()
    Set gr = BuildGraph(sh)
    gr.Add

    Assert.AreEqual 1&, CLng(sh.ChartObjects.Count), "The empty frame is drawn"

    On Error Resume Next
    gr.Format valuesTitle:="Values", catTitle:="Dates"
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual 0&, errNumber, "Formatting a chart with no series raises nothing"
    Assert.AreEqual 0&, CLng(sh.ChartObjects.Count), "And the empty frame is removed"
    Assert.IsTrue gr.HasCheckings, "And the reason is reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAChartWithNoSeriesLosesItsFrame", Err.Number, Err.Description
End Sub

'@sub-title Verify Format before Add is reported.
'@TestMethod("Graphs")
Public Sub TestFormatBeforeAddIsReported()
    CustomTestSetTitles Assert, "Graphs", "TestFormatBeforeAddIsReported"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim gr As Graphs
    Dim errNumber As Long

    Set sh = BuildFixture()
    Set gr = BuildGraph(sh)

    On Error Resume Next
    gr.Format valuesTitle:="Values", catTitle:="Dates"
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual 0&, errNumber, "Formatting before Add raises nothing"
    Assert.AreEqual 0&, CLng(sh.ChartObjects.Count), "And draws nothing"
    Assert.IsTrue gr.HasCheckings, "And it is reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFormatBeforeAddIsReported", Err.Number, Err.Description
End Sub

'@section The checking keys
'===============================================================================

'@sub-title Verify two charts can file into one report.
'@details
'Checking.Add raises ElementShouldNotExists on a duplicate key and
'Checking.Append replays every key into the target, so a key made of a bare
'counter took the whole generation down on the second chart that filed anything.
'AnalysisOutput merges every chart of a sheet into one report. The key names the
'anchor cell of the chart, and every chart of a sheet sits on a cell of its own.
'@TestMethod("Graphs")
Public Sub TestTwoChartsProduceDistinctCheckingKeys()
    CustomTestSetTitles Assert, "Graphs", "TestTwoChartsProduceDistinctCheckingKeys"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim firstGraph As Graphs
    Dim secondGraph As Graphs
    Dim report As Checking
    Dim errNumber As Long

    Set sh = BuildFixture()

    Set firstGraph = Graphs.Create(sh, sh.Cells(5, 7))
    firstGraph.Add
    firstGraph.AddSeries ABSENT_NAME, "bar"

    Set secondGraph = Graphs.Create(sh, sh.Cells(25, 7))
    secondGraph.Add
    secondGraph.AddSeries ABSENT_NAME, "bar"

    Assert.IsTrue firstGraph.HasCheckings, "The first chart filed an entry"
    Assert.IsTrue secondGraph.HasCheckings, "So did the second"

    Set report = Checking.Create("Analysis output")

    On Error Resume Next
    report.Append firstGraph.CheckingValues
    report.Append secondGraph.CheckingValues
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual 0&, errNumber, "Two charts file into one report without a key collision"
    Assert.IsTrue (report.Length > 0), "And the report carries their entries"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTwoChartsProduceDistinctCheckingKeys", _
                         Err.Number, Err.Description
End Sub
