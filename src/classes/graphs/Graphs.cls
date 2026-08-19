VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Graphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manage Excel graphs"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Manage Excel graphs")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class Graphs
'@description
'Wraps one Excel ChartObject on an analysis worksheet. Add draws the chart at a
'caller-supplied Range, AddSeries queues one data series from a named range,
'AddLabels gives that series its category and legend labels, and Format draws
'the queue and applies the look the GraphScope calls for. Entries filed along
'the way leave through HasCheckings and CheckingValues.
'
'THE PHASE ORDER
'-------------------------------------------------------------------------------
'Add, then AddSeries and AddLabels once per series, then Format. All four
'callers drive it that way: AnalysisOutput builds a chart per table and a chart
'per time series graph, and AnaTabIds replays both from its tracking tables
'when a linelist is exported.
'
'THE SERIES ARE QUEUED AND DRAWN IN ONE CALL
'-------------------------------------------------------------------------------
'AddSeries and AddLabels write to a queue and leave the chart alone. CommitSeries
'unions the queued ranges and hands the chart one SetSourceData, then walks the
'series it made to give each its type, its axis and its labels. Format calls
'CommitSeries; a caller that reads the chart before it formats calls it itself.
'
'A SERIES THAT FAILS TAKES ITS LABELS WITH IT
'-------------------------------------------------------------------------------
'AddSeries answers a missing named range by filing a warning and queueing no
'series. The labels that follow belong to that series, so AddLabels reads the
'flag AddSeries leaves and skips them.
'
'EVERY SIZE THIS CLASS APPLIES COMES FROM ONE FUNCTION
'-------------------------------------------------------------------------------
'SizeForScope holds the whole policy: the width coefficient that stretches a
'time series chart, the height coefficient that grows a spatial one, the plot
'area box and the legend position. It is arithmetic over two numbers, so a test
'reaches it through Format with no chart to automate.
'@depends Checking
'@version 1.1 (2026-07-31)



'@section State
'===============================================================================
'@description Backing UDTs, layout constants, and chart state.

Private Type TGraph
    sh As Worksheet
    strtRng As Range
    co As ChartObject
    grName As String
    lastSeriesOk As Boolean
    sealed As Boolean
    checkCounter As Long
    internalChecks As Checking
    pending As Collection
End Type

'@description Every number Format applies that depends on the graph scope.
Private Type TGraphLayout
    chartWidth As Double
    chartHeight As Double
    plotLeft As Double
    plotTop As Double
    plotWidth As Double
    plotHeight As Double
    legendPos As XlLegendPosition
    reverseCategories As Boolean
End Type

Private this As TGraph

'The frame a chart is born with, and the frame Format reapplies.
Private Const CHART_WIDTH As Double = 488
Private Const CHART_HEIGHT As Double = 214

'A time series chart is stretched sideways. A spatial one grows downwards with
'the number of geographic units it shows.
Private Const TIME_SERIES_COEF As Double = 1.75
Private Const HEIGHT_STEP As Double = 0.08

'The plot area box, in the same units as the chart frame.
Private Const PLOT_LEFT_TIME_SERIES As Double = 25
Private Const PLOT_LEFT_STANDARD As Double = 16.618
Private Const PLOT_TOP As Double = 16.618
Private Const PLOT_WIDTH As Double = 425
Private Const PLOT_HEIGHT As Double = 178

'The percentage columns of an analysis table hold fractions. CrossTableFormula
'decides that, and the secondary axis follows it.
Private Const PERCENT_AXIS_MAX As Double = 1
Private Const PERCENT_AXIS_STEP As Double = 0.1

'What one queued series holds until CommitSeries draws it. AddSeries writes the
'first three, AddLabels the rest.
Private Const PENDING_RANGE As Long = 0
Private Const PENDING_TYPE As Long = 1
Private Const PENDING_AXIS As Long = 2
Private Const PENDING_HAS_LABELS As Long = 3
Private Const PENDING_CAT As Long = 4
Private Const PENDING_LABEL As Long = 5
Private Const PENDING_PREFIX As Long = 6
Private Const PENDING_PREFIX_ONLY As Long = 7
Private Const PENDING_HARDCODE As Long = 8

'Type sizes and the theme font used across every chart.
Private Const TICK_FONT_SIZE As Long = 10
Private Const PLOT_TITLE_FONT_SIZE As Long = 14
Private Const DATA_LABEL_FONT_SIZE As Long = 6
Private Const THEME_FONT As String = "+mj-lt"


'@section Enumerations
'===============================================================================
'@description Layout scope controlling graph sizing.

Public Enum GraphScope
  GraphScopeTimeSeries = 1 'A time series graph (long graph)
  GraphScopeSpatial = 2 'A spatial analysis graph
  GraphScopeNormal = 3 'A standard univariate / bivariate graph
End Enum


'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a Graph object bound to a worksheet and position
'@details
'Graph objects simplify the manipulation of Excel graphs. They are based on a basic
'graphic grammar with the ability to stack graphs on top of each other.
'At its core, a graph object is created with a worksheet and a range where
'to write the graph. Values to be plotted should be in a named range on the
'worksheet used to instantiate the class.
'
'The position range has to sit on the worksheet the values come from. The class
'draws the chart on the parent of the range and reads every named range off the
'worksheet, so two different sheets would draw a chart in one place and feed it
'from another, with the legend references qualified by the wrong sheet name.
'@param sh Worksheet. The worksheet containing the values to be plotted.
'@param posRng Range. The cell or range determining where to position the graph.
'@param grName Optional String. Display name for the graph. Defaults to vbNullString.
'@return Graphs. A ready-to-use graph instance.
'@throws ProjectError.ObjectNotInitialized When sh is Nothing.
'@throws ProjectError.ObjectNotInitialized When posRng is Nothing.
'@throws ProjectError.InvalidArgument When posRng sits on another worksheet.
'@depends Checking
'@export
Public Function Create(ByVal sh As Worksheet, ByVal posRng As Range, _
                       Optional ByVal grName As String = vbNullString) As Graphs

    If (sh Is Nothing) Then ThrowError ProjectError.ObjectNotInitialized, "Worksheet not initialized, unable to create graph"
    If (posRng Is Nothing) Then ThrowError ProjectError.ObjectNotInitialized, "Range not initialized, unable to create graph"

    If Not (posRng.Parent Is sh) Then _
        ThrowError ProjectError.InvalidArgument, _
                   "The position range sits on " & posRng.Parent.Name & _
                   " and the graph worksheet is " & sh.Name

    Dim newGraph As Graphs
    Set newGraph = New Graphs
    Set newGraph.Wksh = sh
    Set newGraph.StartRange = posRng
    newGraph.Name = grName
    newGraph.Seal
    Set Create = newGraph
End Function

'@label:seal
'@sub-title Seal the instance against further setup writes
'@details
'Called by the factory after the last setup write. Once sealed, the
'creation-only setters raise through ThrowError.
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against post-creation writes
'@param propName String. The property name reported in the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:wksh
'@prop-title Host worksheet for the graph
'@details
'Returns the worksheet reference stored during creation. All named
'ranges for series data reside on this worksheet.
'@return Worksheet. The worksheet hosting the chart.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:wksh-set
'@prop-title Assign the host worksheet
'@param sh Worksheet. The worksheet where graph data resides.
Public Property Set Wksh(ByVal sh As Worksheet)
    GuardNotSealed "Wksh"
    Set this.sh = sh
End Property

'@label:start-range
'@prop-title Starting cell or range for the graph position
'@details
'Returns the Range that determines the top-left anchor point of the
'chart on the worksheet.
'@return Range. The positioning range.
Public Property Get StartRange() As Range
    Set StartRange = this.strtRng
End Property

'@label:start-range-set
'@prop-title Assign the starting position range
'@param strtRng Range. The cell or range used to position the graph.
Public Property Set StartRange(ByVal strtRng As Range)
    GuardNotSealed "StartRange"
    Set this.strtRng = strtRng
End Property

'@label:name-set
'@prop-title Assign the graph display name
'@param grName String. The graph name.
Public Property Let Name(ByVal grName As String)
    GuardNotSealed "Name"
    this.grName = grName
End Property

'@label:name
'@prop-title Retrieve the graph display name
'@return String. The graph name.
Public Property Get Name() As String
    Name = this.grName
End Property


'@section GraphElements
'===============================================================================
'@description Internal chart object state.

'@label:graph-chart
'@prop-title Retrieve the internal ChartObject
'@details
'Returns the ChartObject created by Add. Nothing until Add is called, and
'Nothing again after Format removes an empty frame.
'@return ChartObject. The underlying Excel chart object.
Private Property Get graphChart() As ChartObject
    Set graphChart = this.co
End Property


'@section Operations
'===============================================================================
'@description Chart creation, series attachment, labelling, and formatting.

'@label:add
'@sub-title Create an empty chart on the worksheet
'@details
'Creates an empty graph on a worksheet, ready for data to be
'attached to the graph. The graph is empty at creation, so the legend and the
'plot area are left to Format, which owns both and runs once the series exist.
'
'The frame comes from Shapes.AddChart2 rather than ChartObjects.Add. AddChart2
'builds the chart in one host call where the older one builds a frame and then
'takes its style, its type and its plot area in several. It plots whatever range
'is selected when it runs, so the series it invents are deleted here and the
'chart starts empty either way.
'@export
Public Sub Add()

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim rng As Range
    Dim shp As Shape
    Dim anchorLeft As Double
    Dim anchorTop As Double

    Set rng = StartRange()
    Set sh = rng.Parent
    anchorLeft = rng.Left
    anchorTop = rng.Top

    Set shp = sh.Shapes.AddChart2(-1, xlColumnClustered, anchorLeft, anchorTop, _
                                  CHART_WIDTH, CHART_HEIGHT, False)
    Set co = sh.ChartObjects(shp.Name)

    Do While co.Chart.SeriesCollection.Count > 0
        co.Chart.SeriesCollection(1).Delete
    Loop

    Set this.co = co
End Sub

'@label:addseries
'@sub-title Attach a data series from a named range
'@details
'One graph can have multiple series added. Series are referenced using
'named ranges on the graph object worksheet. If the named range does not
'exist, the method files a warning, queues no series, and marks the attempt as
'failed so the labels that follow are skipped too. If no chart exists
'yet, Add is called lazily before the first series is queued.
'
'The series is queued, not drawn. CommitSeries hands the whole queue to the
'chart in one SetSourceData call, and Format calls it. A caller that wants to
'read the chart before it formats calls CommitSeries itself.
'@param rngName String. Named range containing the series values.
'@param chrtType String. Chart type: "line", "point", "bar", or "hbar". Defaults to "bar" for unknown values.
'@param axisPos Optional String. Axis position: "right" adds a secondary axis; any other value uses the primary axis. Defaults to vbNullString.
'@export
Public Sub AddSeries(ByVal rngName As String, ByVal chrtType As String, _
                    Optional ByVal axisPos As String = vbNullString)

    Dim seriesRng As Range
    Dim entry As Variant

    this.lastSeriesOk = False

    'If the provided range does not exist, exit the sub
    If Not TryRange(rngName, seriesRng) Then
        LogInfo "Worksheet " & Wksh().Name & " carries no range named " & rngName & _
                ", so this series is left out of the graph", checkingWarning
        Exit Sub
    End If

    'If no graph exists, add one first
    If graphChart() Is Nothing Then Add

    'The fields are in the order the PENDING_ constants name.
    entry = Array(rngName, chrtType, axisPos, False, _
                  vbNullString, vbNullString, vbNullString, False, True)

    PendingSeries.Add entry

    this.lastSeriesOk = True
End Sub

'@label:pending-series
'@sub-title The queue of series waiting to be drawn
'@return Collection. The queue, created empty on first use.
Private Function PendingSeries() As Collection
    If this.pending Is Nothing Then Set this.pending = New Collection
    Set PendingSeries = this.pending
End Function

'@label:chart-type-of
'@sub-title The Excel chart type one series name stands for
'@param chrtType String. Chart type: "line", "point", "bar", or "hbar".
'@return XlChartType. The matching type, clustered columns for anything else.
Private Function ChartTypeOf(ByVal chrtType As String) As XlChartType
    Select Case chrtType
        Case "line": ChartTypeOf = xlLineMarkers
        Case "point": ChartTypeOf = xlXYScatter
        Case "bar": ChartTypeOf = xlColumnClustered
        Case "hbar": ChartTypeOf = xlBarClustered
        Case Else: ChartTypeOf = xlColumnClustered
    End Select
End Function


'@label:addlabels
'@sub-title Assign category and legend labels to the current series
'@details
'Sets the category axis values and legend entry for the most recently
'added series. The optional prefix is prepended to the legend name,
'separated by " - " unless prefixOnly is True. When hardCodeLabels is
'False, the legend entry references the named range dynamically rather
'than using a literal string.
'
'The labels belong to the series added just before them, so a failed AddSeries
'skips them. Writing them anyway put series n's category axis and legend entry
'onto series n-1.
'@param catName String. Named range for category axis labels.
'@param lblName String. Named range for the legend entry text.
'@param prefix Optional String. Display prefix prepended to the legend name. Defaults to vbNullString.
'@param prefixOnly Optional Boolean. When True, uses only the prefix as the legend entry. Defaults to False.
'@param hardCodeLabels Optional Boolean. When True, legend text is literal; when False, references the named range. Defaults to True.
Public Sub AddLabels(ByVal catName As String, _
                     ByVal lblName As String, _
                     Optional ByVal prefix As String = vbNullString, _
                     Optional ByVal prefixOnly As Boolean = False, _
                     Optional ByVal hardCodeLabels As Boolean = True)

    Dim entry As Variant
    Dim ind As Long

    If graphChart() Is Nothing Then
        LogInfo "Labels " & catName & " were asked for before any chart was made", _
                checkingWarning
        Exit Sub
    End If

    ind = PendingSeries.Count
    If ind = 0 Then
        LogInfo "Labels " & catName & " were asked for and this graph carries no series", _
                checkingWarning
        Exit Sub
    End If

    If Not this.lastSeriesOk Then
        LogInfo "Labels " & catName & " belong to a series that was left out, " & _
                "so they are skipped", checkingWarning
        Exit Sub
    End If

    'The labels join the series they belong to, in the queue. A Collection has
    'no way to write an item back, so the entry leaves the end of the queue and
    'returns to it filled in.
    entry = PendingSeries.Item(ind)
    entry(PENDING_HAS_LABELS) = True
    entry(PENDING_CAT) = catName
    entry(PENDING_LABEL) = lblName
    entry(PENDING_PREFIX) = prefix
    entry(PENDING_PREFIX_ONLY) = prefixOnly
    entry(PENDING_HARDCODE) = hardCodeLabels

    PendingSeries.Remove ind
    PendingSeries.Add entry
End Sub

'@label:commit-series
'@sub-title Draw every queued series in one call to the chart
'@details
'The queued ranges are the columns of one analysis table, so their union is a
'single Source Excel splits back into the same series, in the same order. One
'SetSourceData replaces the two data writes each series used to cost: the older
'path added a series from its own Source and then wrote the values over it,
'because SeriesCollection.Add guesses how to split a Source whose cells are
'mostly empty and a column with data in its last row alone came out as a
'one-cell series.
'
'Excel still does the splitting here, so the count it made is checked against
'the count that was asked for. When the two differ, or SetSourceData refuses the
'Source outright, the series are attached one at a time and the values pinned
'range by range. The look of a series -- its type, its axis, its labels -- is
'applied afterwards either way.
'
'Format calls this. A caller that reads the chart before it formats calls it
'itself. Calling it twice costs nothing: the queue is empty the second time.
'@export
Public Sub CommitSeries()

    Dim co As ChartObject
    Dim live As Collection
    Dim rngs As Collection
    Dim seriesRng As Range
    Dim allRng As Range
    Dim entry As Variant
    Dim ind As Long
    Dim sourceTaken As Boolean

    If PendingSeries.Count = 0 Then Exit Sub

    Set co = graphChart()
    If co Is Nothing Then
        Set this.pending = Nothing
        Exit Sub
    End If

    'A name can go between the queueing and the drawing. Only the series whose
    'range still resolves are drawn, and they are the ones counted.
    Set live = New Collection
    Set rngs = New Collection

    For ind = 1 To PendingSeries.Count
        entry = PendingSeries.Item(ind)
        If TryRange(CStr(entry(PENDING_RANGE)), seriesRng) Then
            live.Add entry
            rngs.Add seriesRng
            If allRng Is Nothing Then
                Set allRng = seriesRng
            Else
                Set allRng = Application.Union(allRng, seriesRng)
            End If
        Else
            LogInfo "The range named " & CStr(entry(PENDING_RANGE)) & " went away " & _
                    "before the chart was drawn, so this series is left out", checkingWarning
        End If
    Next

    Set this.pending = Nothing
    If live.Count = 0 Then Exit Sub

    On Error Resume Next
        co.Chart.SetSourceData Source:=allRng, PlotBy:=xlColumns
        sourceTaken = (Err.Number = 0)
        Err.Clear
    On Error GoTo 0

    If sourceTaken Then sourceTaken = InSheetOrder(rngs)
    If sourceTaken Then sourceTaken = (co.Chart.SeriesCollection.Count = live.Count)
    If sourceTaken Then sourceTaken = PointsMatch(co, rngs.Item(1))

    If Not sourceTaken Then
        'Filed as information, not as a warning. Nothing is wrong with the chart
        'and nothing is asked of the person reading the report: it says which of
        'the two paths drew the series.
        LogInfo "Graph " & GraphKey() & " did not take its series in one call, " & _
                "so they are attached one at a time", checkingInfo
        AttachOneByOne co, rngs
    End If

    StyleSeries co, live
End Sub

'@label:in-sheet-order
'@sub-title Whether the queued ranges run left to right down the sheet
'@details
'A Source is plotted in the order the sheet holds it, and the styling that
'follows walks the series by index against the queue. The two agree only while
'the queue is in sheet order, and the analysis tables name their columns that
'way: CrossTable writes VALUES_COL_1, PERC_COL_1, VALUES_COL_2 across the table
'as GraphSeries queues them.
'
'A queue that arrives in some other order is not wrong, it just cannot be read
'back by index, so those series are attached one at a time instead. Nothing here
'guesses: it compares the column and the row each range starts at.
'@param rngs Collection. The series ranges, in the order they were queued.
'@return Boolean. True when each range starts at or after the one before it.
Private Function InSheetOrder(ByVal rngs As Collection) As Boolean
    Dim ind As Long
    Dim thisCol As Long
    Dim thisRow As Long
    Dim lastCol As Long
    Dim lastRow As Long

    InSheetOrder = True

    For ind = 1 To rngs.Count
        thisCol = rngs.Item(ind).Column
        thisRow = rngs.Item(ind).Row

        If ind > 1 Then
            If thisCol < lastCol Then
                InSheetOrder = False
                Exit Function
            End If
            If (thisCol = lastCol) And (thisRow < lastRow) Then
                InSheetOrder = False
                Exit Function
            End If
        End If

        lastCol = thisCol
        lastRow = thisRow
    Next
End Function

'@label:points-match
'@sub-title Whether the chart plotted a whole column and not part of one
'@details
'The series count alone does not say the split was right. A column holding one
'number in its last row and nothing above it comes back as a one-point series,
'and one series is what was asked for, so the count agrees while the chart shows
'a single bar over the first period. The points of the first series are counted
'against the cells of its range, which is the same layout every other series of
'the chart has.
'
'A range wider than one column is left alone: its points were never meant to
'match its cells, and a count that far off is caught before this runs.
'@param co ChartObject. The chart that took the Source.
'@param firstRng Range. The range the first series was asked for.
'@return Boolean. True when the first series carries one point per cell.
Private Function PointsMatch(ByVal co As ChartObject, ByVal firstRng As Range) As Boolean
    Dim plotted As Variant

    PointsMatch = True
    If firstRng.Columns.Count > 1 Then Exit Function

    On Error Resume Next
        plotted = co.Chart.SeriesCollection(1).Values
        If Err.Number <> 0 Then
            Err.Clear
            PointsMatch = False
            Exit Function
        End If
    On Error GoTo 0

    If Not IsArray(plotted) Then
        PointsMatch = False
        Exit Function
    End If

    PointsMatch = (UBound(plotted) - LBound(plotted) + 1 = firstRng.Cells.Count)
End Function

'@label:attach-one-by-one
'@sub-title Give the chart one series per range, values pinned
'@details
'The fallback of CommitSeries. Whatever Excel made of the single Source goes,
'and each range becomes an empty series with its values written once.
'@param co ChartObject. The chart being filled.
'@param rngs Collection. The series ranges, in order.
Private Sub AttachOneByOne(ByVal co As ChartObject, ByVal rngs As Collection)
    Dim ind As Long
    Dim srs As Series

    Do While co.Chart.SeriesCollection.Count > 0
        co.Chart.SeriesCollection(1).Delete
    Loop

    For ind = 1 To rngs.Count
        Set srs = co.Chart.SeriesCollection.NewSeries
        srs.Values = rngs.Item(ind)
    Next
End Sub

'@label:style-series
'@sub-title Give every drawn series its type, its axis and its labels
'@details
'The Series objects are taken before the first type is written. Turning one
'series into a scatter moves it between chart groups, and an index read after
'that can answer a different series than the one the queue describes.
'@param co ChartObject. The chart carrying the series.
'@param live Collection. The queue entries that were drawn, in order.
Private Sub StyleSeries(ByVal co As ChartObject, ByVal live As Collection)
    Dim ind As Long
    Dim drawn() As Series
    Dim entry As Variant

    If co.Chart.SeriesCollection.Count < live.Count Then
        LogInfo "Graph " & GraphKey() & " carries " & co.Chart.SeriesCollection.Count & _
                " series where " & live.Count & " were asked for", checkingWarning
        Exit Sub
    End If

    ReDim drawn(1 To live.Count)
    For ind = 1 To live.Count
        Set drawn(ind) = co.Chart.SeriesCollection(ind)
    Next

    For ind = 1 To live.Count
        entry = live.Item(ind)

        'Add a right axis if required
        If CStr(entry(PENDING_AXIS)) = "right" Then
            co.Chart.HasAxis(xlValue, xlSecondary) = True
            drawn(ind).AxisGroup = xlSecondary
        End If

        drawn(ind).chartType = ChartTypeOf(CStr(entry(PENDING_TYPE)))

        If CBool(entry(PENDING_HAS_LABELS)) Then ApplySeriesLabels drawn(ind), entry
    Next
End Sub

'@label:apply-series-labels
'@sub-title Write the category values, the legend entry and the data labels
'@param srs Series. The series being labelled.
'@param entry Variant. The queue entry describing it.
Private Sub ApplySeriesLabels(ByVal srs As Series, ByRef entry As Variant)

    Dim labelsObj As DataLabels
    Dim catRng As Range
    Dim lblRng As Range
    Dim hasLabelRange As Boolean

    hasLabelRange = TryRange(CStr(entry(PENDING_LABEL)), lblRng)

    If TryRange(CStr(entry(PENDING_CAT)), catRng) Then srs.XValues = catRng

    'Labels can be hardcoded or dynamic; by default, they are hardcoded
    If (Not CBool(entry(PENDING_HARDCODE))) And hasLabelRange Then
        'The user can modify by changing values in the range, so the legend
        'entry is written as a reference
        srs.Name = "=" & QualifiedAddress(lblRng)
    Else
        srs.Name = LegendText(CStr(entry(PENDING_PREFIX)), _
                              CBool(entry(PENDING_PREFIX_ONLY)), _
                              hasLabelRange, lblRng)
    End If

    srs.ApplyDataLabels xlDataLabelsShowValue, False
    Set labelsObj = srs.DataLabels
    With labelsObj.Font
        .Size = DATA_LABEL_FONT_SIZE
        .color = RGB(2, 6, 91)
        .Name = THEME_FONT
    End With
End Sub

'@label:legend-text
'@sub-title Build the literal legend entry of a series
'@details
'Reads the label cell only on the branch that uses it. Every simple-mode series
'asks for a reference instead, so the read used to be paid for and thrown away.
'@param prefix String. Display prefix.
'@param prefixOnly Boolean. True answers the prefix alone.
'@param hasLabelRange Boolean. True when lblRng holds a resolved range.
'@param lblRng Range. The cell carrying the label text.
'@return String. The text to write as the series name.
Private Function LegendText(ByVal prefix As String, ByVal prefixOnly As Boolean, _
                            ByVal hasLabelRange As Boolean, _
                            ByVal lblRng As Range) As String
    Dim labelValue As String

    labelValue = prefix

    If prefixOnly Then
        LegendText = labelValue
        Exit Function
    End If

    'Add a prefix to the label name (useful for time series graphs)
    If LenB(labelValue) > 0 Then labelValue = labelValue & " - "

    LegendText = labelValue
    If hasLabelRange Then LegendText = labelValue & CStr(lblRng.Value)
End Function

'@label:format
'@sub-title Apply layout and formatting to the chart
'@details
'Formats are defined based on the scope of the graph. For time series, graphs are larger;
'for spatio-temporal analyses, they are wider. The scope controls dimension coefficients,
'axis gridline colours, legend placement, and plot area reverse ordering. When
'hardCodeLabels is False, axis titles and the chart title reference named ranges
'dynamically rather than using literal strings.
'
'A chart with no series has no axes, so reading them raised 1004 and took the
'whole chart pass of a sheet with it. Format removes such a frame and files the
'reason.
'@param valuesTitle Optional String. Title for the value axis (Y-axis). Defaults to vbNullString.
'@param catTitle Optional String. Title for the category axis (X-axis). Defaults to vbNullString.
'@param plotTitle Optional String. Title displayed above the chart. Defaults to vbNullString.
'@param secondAxisTitle Optional String. Title for the secondary value axis. Defaults to vbNullString.
'@param scope Optional Byte. GraphScope value controlling layout sizing. Defaults to GraphScopeNormal.
'@param heightFactor Optional Long. Multiplicative factor applied to chart height. Defaults to 1.
'@param hardCodeLabels Optional Boolean. When True, axis titles are literal text. Defaults to True.
Public Sub Format(Optional ByVal valuesTitle As String = vbNullString, _
                  Optional ByVal catTitle As String = vbNullString, _
                  Optional ByVal plotTitle As String = vbNullString, _
                  Optional ByVal secondAxisTitle As String = vbNullString, _
                  Optional ByVal scope As Byte = GraphScopeNormal, _
                  Optional ByVal heightFactor As Long = 1, _
                  Optional ByVal hardCodeLabels As Boolean = True)

    Dim co As ChartObject
    Dim layout As TGraphLayout

    CommitSeries

    Set co = graphChart()
    If Not ReadyToFormat(co) Then Exit Sub

    layout = SizeForScope(KnownScope(scope), heightFactor)

    ApplyChartFrame co, layout

    With co.Chart
        FormatCategoryAxis .Axes(xlCategory, xlPrimary), catTitle, hardCodeLabels, layout
        FormatValueAxis .Axes(xlValue, xlPrimary), valuesTitle, hardCodeLabels
        FormatChartTitle co.Chart, plotTitle, hardCodeLabels

        If .HasAxis(xlValue, xlSecondary) Then _
            FormatSecondaryAxis .Axes(xlValue, xlSecondary), secondAxisTitle

        ApplyPlotArea .PlotArea, layout
        ApplyLegend co.Chart, layout
    End With
End Sub


'@section Layout
'===============================================================================
'@description The scope-to-dimensions policy, and the six steps Format applies.

'@label:known-scope
'@sub-title Answer a scope this class knows how to lay out
'@details
'AnalysisOutput maps its own four analysis scopes onto these three before it
'calls, and AnaTabIds does the same when it replays a chart. A number outside
'the three used to fall through every branch and take the standard layout with
'nothing said about it.
'@param scope Byte. A GraphScope value.
'@return Byte. The scope given, or GraphScopeNormal when it is unknown.
Private Function KnownScope(ByVal scope As Byte) As Byte
    Select Case scope
    Case GraphScopeTimeSeries, GraphScopeSpatial, GraphScopeNormal
        KnownScope = scope
    Case Else
        LogInfo "Graph scope " & scope & " is outside the three this class knows, " & _
                "so the standard layout is used", checkingWarning
        KnownScope = GraphScopeNormal
    End Select
End Function

'@label:height-in-rows
'@sub-title How many worksheet rows a chart of this scope covers
'@details
'A caller that draws its charts after its tables still has to leave room for
'them while the tables are placed, and the chart is not there to be measured
'yet. This answers the height the chart WILL have, in rows of the sheet it will
'sit on, from the same SizeForScope policy Format applies.
'
'The scope is sanitised here rather than through KnownScope, because that one
'files a checking and this is called on the predeclared instance, which owns no
'report. An unknown scope takes the standard layout, the same answer.
'
'The row height comes from the sheet. Analysis output sheets are left at the
'standard height -- the row-height rules in LLFormat belong to the HList print
'and pivot sheets -- so a division answers the same row a drawn chart would
'reach. One row is added rather than rounding, so the reservation is never
'short.
'@param sh Worksheet. The sheet the chart will be drawn on.
'@param scope Byte. A GraphScope value.
'@param heightFactor Long. The geographic unit count, 1 or more.
'@return Long. The number of rows to reserve, or 0 when sh is Nothing.
'@export
Public Function HeightInRows(ByVal sh As Worksheet, ByVal scope As Byte, _
                             ByVal heightFactor As Long) As Long
    Dim layout As TGraphLayout
    Dim rowHeight As Double
    Dim safeScope As Byte

    If sh Is Nothing Then Exit Function

    safeScope = scope
    Select Case safeScope
    Case GraphScopeTimeSeries, GraphScopeSpatial, GraphScopeNormal
    Case Else
        safeScope = GraphScopeNormal
    End Select

    layout = SizeForScope(safeScope, heightFactor)

    rowHeight = sh.StandardHeight
    If rowHeight <= 0 Then rowHeight = 15

    HeightInRows = Int(layout.chartHeight / rowHeight) + 1
End Function

'@label:size-for-scope
'@sub-title Answer the chart and plot-area dimensions for one scope
'@details
'Holds every number that depends on the graph scope. The width coefficient
'stretches a time series graph. The height coefficient grows the chart with
'the number of geographic units it shows.
'
'Both coefficients are fractional and both were declared As Long once. 1.75 was
'stored as 2, so every time series chart came out 14 per cent larger than the
'code asks for, and a spatial chart of five geographic units got 0.48 rounded
'down to zero, which is a chart of no height at all.
'
'The height curve anchors at one geographic unit and each further unit adds
'8 per cent, so five units give 1.32 and twenty give 2.52. Every count a user
'can set moves the chart, and the bars keep their thickness as the count rises.
'The floor answers a count below 1, which the callers already clamp away.
'@param scope Byte. A GraphScope value.
'@param heightFactor Long. The geographic unit count, 1 or more.
'@return TGraphLayout. The dimensions to apply.
Private Function SizeForScope(ByVal scope As Byte, _
                              ByVal heightFactor As Long) As TGraphLayout
    Dim layout As TGraphLayout
    Dim widthCoef As Double
    Dim heightCoef As Double

    widthCoef = 1
    If scope = GraphScopeTimeSeries Then widthCoef = TIME_SERIES_COEF

    heightCoef = 1 + (heightFactor - 1) * HEIGHT_STEP
    If heightCoef < 1 Then heightCoef = 1

    layout.chartWidth = CHART_WIDTH * widthCoef
    layout.chartHeight = CHART_HEIGHT * widthCoef * heightCoef

    layout.plotLeft = PLOT_LEFT_STANDARD * widthCoef
    If scope = GraphScopeTimeSeries Then layout.plotLeft = PLOT_LEFT_TIME_SERIES * widthCoef

    layout.plotTop = PLOT_TOP * widthCoef
    layout.plotWidth = PLOT_WIDTH * widthCoef
    layout.plotHeight = PLOT_HEIGHT * widthCoef * heightCoef

    layout.legendPos = xlLegendPositionTop
    If scope = GraphScopeSpatial Then layout.legendPos = xlLegendPositionBottom

    layout.reverseCategories = (scope = GraphScopeSpatial)

    SizeForScope = layout
End Function

'@label:ready-to-format
'@sub-title Whether the chart can take a layout
'@details
'A chart with zero series has no axes, and reading Axes on one raises 1004.
'An empty bordered frame on the analysis sheet reads as a broken chart, and the
'export would faithfully draw it again, so the frame goes.
'@param co ChartObject. The chart Format is about to lay out.
'@return Boolean. True when there is a chart carrying at least one series.
Private Function ReadyToFormat(ByVal co As ChartObject) As Boolean
    If co Is Nothing Then
        LogInfo "Format ran before Add, so this graph has nothing to lay out", _
                checkingWarning
        Exit Function
    End If

    If co.Chart.SeriesCollection.Count = 0 Then
        LogInfo "Graph " & GraphKey() & " carries no series, so its empty frame is removed", _
                checkingWarning
        co.Delete
        Set this.co = Nothing
        Exit Function
    End If

    ReadyToFormat = True
End Function

'@label:apply-chart-frame
'@sub-title Size the chart and draw its border
'@param co ChartObject. The chart being laid out.
'@param layout TGraphLayout. The dimensions of its scope.
Private Sub ApplyChartFrame(ByVal co As ChartObject, ByRef layout As TGraphLayout)
    co.width = layout.chartWidth
    co.height = layout.chartHeight

    With co.ShapeRange
        .Shadow.Type = msoShadow21
        .line.Visible = msoTrue
        .line.ForeColor.RGB = RGB(173, 216, 230)
    End With

    'Avoid changing graph width when adjusting row height
    co.Placement = xlMove
End Sub

'@label:format-category-axis
'@sub-title Format the category axis and write its title
'@param ax Axis. The primary category axis.
'@param catTitle String. Axis title, or the name of a range holding one.
'@param hardCodeLabels Boolean. True writes the title as literal text.
'@param layout TGraphLayout. The dimensions of its scope.
Private Sub FormatCategoryAxis(ByVal ax As Axis, ByVal catTitle As String, _
                               ByVal hardCodeLabels As Boolean, _
                               ByRef layout As TGraphLayout)
    With ax
        .HasTitle = True
        .HasMajorGridlines = True
        .MajorGridlines.Border.color = RGB(255, 255, 255)
        .TickLabels.Font.Size = TICK_FONT_SIZE
        'Spatial analysis reads from the largest unit downwards
        .ReversePlotOrder = layout.reverseCategories
    End With

    SetTitleFromName ax.AxisTitle, catTitle, hardCodeLabels
End Sub

'@label:format-value-axis
'@sub-title Format the primary value axis and write its title
'@param ax Axis. The primary value axis.
'@param valuesTitle String. Axis title, or the name of a range holding one.
'@param hardCodeLabels Boolean. True writes the title as literal text.
Private Sub FormatValueAxis(ByVal ax As Axis, ByVal valuesTitle As String, _
                            ByVal hardCodeLabels As Boolean)
    With ax
        .HasTitle = True
        .HasMajorGridlines = True
        .MajorGridlines.Border.color = RGB(255, 255, 255)
        .TickLabels.Font.Size = TICK_FONT_SIZE
        .TickLabels.Font.Name = THEME_FONT
    End With

    SetTitleFromName ax.AxisTitle, valuesTitle, hardCodeLabels
End Sub

'@label:format-chart-title
'@sub-title Write the chart title, when one was asked for
'@param cht Chart. The chart being laid out.
'@param plotTitle String. Chart title, or the name of a range holding one.
'@param hardCodeLabels Boolean. True writes the title as literal text.
Private Sub FormatChartTitle(ByVal cht As Chart, ByVal plotTitle As String, _
                             ByVal hardCodeLabels As Boolean)
    cht.HasTitle = False
    If LenB(plotTitle) = 0 Then Exit Sub

    cht.HasTitle = True
    With cht.ChartTitle.Font
        .Size = PLOT_TITLE_FONT_SIZE
        .Bold = True
        .Name = THEME_FONT
    End With

    SetTitleFromName cht.ChartTitle, plotTitle, hardCodeLabels
End Sub

'@label:format-secondary-axis
'@sub-title Format the secondary value axis
'@details
'The scale runs from 0 to 1 with a tick every tenth, which is what a fraction
'needs. CrossTableFormula is what decides that the percentage columns hold
'fractions.
'@param ax Axis. The secondary value axis.
'@param secondAxisTitle String. Axis title. An empty one gives a per cent sign.
Private Sub FormatSecondaryAxis(ByVal ax As Axis, ByVal secondAxisTitle As String)
    Dim titleText As String

    titleText = "%"
    If LenB(secondAxisTitle) > 0 Then titleText = secondAxisTitle

    With ax
        .HasTitle = True
        .HasMajorGridlines = False
        .TickLabels.Font.Size = TICK_FONT_SIZE
        .TickLabels.Font.Name = THEME_FONT
        .AxisTitle.Caption = titleText
        .MaximumScale = PERCENT_AXIS_MAX
        .MajorUnit = PERCENT_AXIS_STEP
    End With
End Sub

'@label:apply-plot-area
'@sub-title Colour and place the plot area
'@param plotBox PlotArea. The plot area of the chart.
'@param layout TGraphLayout. The dimensions of its scope.
Private Sub ApplyPlotArea(ByVal plotBox As PlotArea, ByRef layout As TGraphLayout)
    With plotBox
        .Interior.color = RGB(235, 235, 235)
        .Left = layout.plotLeft
        .Top = layout.plotTop
        .width = layout.plotWidth
        .height = layout.plotHeight
    End With
End Sub

'@label:apply-legend
'@sub-title Place the legend
'@details
'Format owns the legend position. Add used to set it as well, on a chart that
'had no series and so may have had no legend at all.
'@param cht Chart. The chart being laid out.
'@param layout TGraphLayout. The dimensions of its scope.
Private Sub ApplyLegend(ByVal cht As Chart, ByRef layout As TGraphLayout)
    cht.HasLegend = True
    cht.Legend.Position = layout.legendPos
End Sub

'@label:set-title-from-name
'@sub-title Write a title as literal text or as a reference to a cell
'@details
'Formula is the property the object model gives for a reference, so that is what
'a dynamic title is written through. Caption was measured on 2026-07-31 doing
'the same thing on this host: it parses "= 'Analysis'!$C$7" and binds the title
'to the cell, which is what the class relied on before. TestGraphs holds that
'measurement.
'
'A caller asking for a reference against a name the worksheet does not carry is
'told about it, because that combination means the caller expected the name to
'be there.
'@param titleObj Object. An AxisTitle or a ChartTitle.
'@param nameOrText String. The title text, or the name of a range holding it.
'@param hardCoded Boolean. True writes the text as it stands.
Private Sub SetTitleFromName(ByVal titleObj As Object, ByVal nameOrText As String, _
                             ByVal hardCoded As Boolean)
    Dim rng As Range

    If hardCoded Then
        titleObj.Caption = nameOrText
        Exit Sub
    End If

    If TryRange(nameOrText, rng) Then
        titleObj.Formula = "=" & QualifiedAddress(rng)
        Exit Sub
    End If

    If LenB(nameOrText) > 0 Then _
        LogInfo "Graph title " & nameOrText & " was asked for as a reference and worksheet " & _
                Wksh().Name & " carries no such referenced name, so the text is written as it stands", _
                checkingInfo

    titleObj.Caption = nameOrText
End Sub


'@section Helpers
'===============================================================================
'@description Private utility functions for range resolution.

'@label:try-range
'@sub-title Resolve a named range on the host worksheet, answering whether it exists
'@details
'The class used to ask twice for every name it wanted: once to learn whether the
'name resolved and once to read it, and the address helper asked twice more.
'Labelling one series cost seven resolutions of two names. A miss costs more
'than a hit, because a miss is a trapped COM error, and the export path asks for
'an empty name on every time series series it replays.
'@param rngName String. The name to resolve.
'@param outRng Range. Receives the resolved range, or Nothing.
'@return Boolean. True when the name resolved.
Private Function TryRange(ByVal rngName As String, ByRef outRng As Range) As Boolean
    Set outRng = Nothing
    If LenB(rngName) = 0 Then Exit Function

    On Error Resume Next
        Set outRng = Wksh().Range(rngName)
        Err.Clear
    On Error GoTo 0

    TryRange = Not (outRng Is Nothing)
End Function

'@label:qualified-address
'@sub-title Return the full worksheet-qualified address of a range
'@details
'The address is in the format 'SheetName'!$A$1, which a chart formula takes.
'@param rng Range. The range to describe.
'@return String. Worksheet-qualified address.
Private Function QualifiedAddress(ByVal rng As Range) As String
    QualifiedAddress = "'" & rng.Worksheet.Name & "'!" & rng.Address
End Function


'@section Checking
'===============================================================================
'@description Logging and validation support via the Checking interface.

'@label:graph-key
'@sub-title A name for this chart that no other chart shares
'@details
'AnalysisOutput builds one chart per table and one per time series graph, and
'the caller of a checking merges them all into one report. Checking.Add raises
'on a duplicate key, so a key made of a bare counter took the whole generation
'down on the second chart that filed anything. The display name is used when
'the caller gave one; otherwise the cell the chart sits on names it, and every
'chart of a sheet sits on a cell of its own.
'@return String. The identifier used in every key this instance files.
Private Function GraphKey() As String
    If LenB(Name()) > 0 Then
        GraphKey = Name()
        Exit Function
    End If

    If Not this.strtRng Is Nothing Then
        GraphKey = this.strtRng.Address(False, False)
        Exit Function
    End If

    GraphKey = "unplaced"
End Function

'@label:log-info
'@sub-title Record a diagnostic entry in the internal checking store
'@details
'Lazily creates the Checking instance on first use and appends the message
'under a key naming this class, this chart and a counter. The checking store is
'exposed through HasCheckings and CheckingValues.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level from CheckingLogType. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "Graph " & GraphKey() & " object trace"

    'First, instantiate the internalChecks object
    If this.internalChecks Is Nothing Then Set this.internalChecks = Checking.Create(checkName)

    'Add new information to the checking object
    this.internalChecks.Add "Graphs-" & GraphKey() & "-" & CStr(this.checkCounter), _
                            label, scope
    this.checkCounter = this.checkCounter + 1
End Sub

'@label:has-checkings
'@prop-title Determine whether diagnostics have been captured
'@details
'Returns True when at least one LogInfo call has been made,
'indicating that diagnostic information is available.
'@return Boolean. True when internal checks exist.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = Not (this.internalChecks Is Nothing)
End Property

'@label:checking-values
'@prop-title Retrieve the collected checking entries
'@details
'Returns the internal Checking store when at least one entry has
'been recorded. Returns Nothing otherwise.
'@return Object. The Checking instance with logged diagnostics.
'@export
Public Property Get CheckingValues() As Checking

    'If there are no checks, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = this.internalChecks
End Property


'@section ErrorHandling
'===============================================================================
'@description Centralised error-raising helper.

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source for consistent
'stack traces across all methods in this class.
'@param errNumb Long. ProjectError code describing the failure.
'@param errorMessage String. Descriptive error message.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "Graphs", "Error : " & Chr(34) & errorMessage
End Sub
