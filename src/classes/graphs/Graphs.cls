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
'caller-supplied Range, AddSeries attaches one data series from a named range,
'AddLabels gives that series its category and legend labels, and Format applies
'the look the GraphScope calls for. Entries filed along the way leave through
'HasCheckings and CheckingValues.
'
'THE PHASE ORDER
'-------------------------------------------------------------------------------
'Add, then AddSeries and AddLabels once per series, then Format. All four
'callers drive it that way: AnalysisOutput builds a chart per table and a chart
'per time series graph, and AnaTabIds replays both from its tracking tables
'when a linelist is exported.
'
'A SERIES THAT FAILS TAKES ITS LABELS WITH IT
'-------------------------------------------------------------------------------
'AddSeries answers a missing named range by filing a warning and adding no
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

'@label:series-count
'@sub-title How many series the chart carries
'@details
'The chart's own SeriesCollection is the truth. The class used to keep a counter
'beside it and add one per AddSeries call, which drifted from the chart the
'moment a series was left out.
'@return Long. The number of series on the chart, 0 when there is no chart.
Private Function SeriesCount() As Long
    If this.co Is Nothing Then Exit Function
    SeriesCount = this.co.Chart.SeriesCollection.Count
End Function

'@label:last-series
'@sub-title The series added most recently
'@return Series. The last series of the chart, or Nothing when there is none.
Private Function LastSeries() As Series
    Dim ind As Long

    ind = SeriesCount()
    If ind = 0 Then Exit Function

    Set LastSeries = this.co.Chart.SeriesCollection(ind)
End Function


'@section Operations
'===============================================================================
'@description Chart creation, series attachment, labelling, and formatting.

'@label:add
'@sub-title Create an empty chart on the worksheet
'@details
'Creates an empty graph on a worksheet, ready for data to be
'attached to the graph. The graph is empty at creation, so the legend and the
'plot area are left to Format, which owns both and runs once the series exist.
'@export
Public Sub Add()

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim rng As Range
    Dim anchorLeft As Double
    Dim anchorTop As Double

    Set rng = StartRange()
    Set sh = rng.Parent
    anchorLeft = rng.Left
    anchorTop = rng.Top

    Set co = sh.ChartObjects.Add(anchorLeft, anchorTop, CHART_WIDTH, CHART_HEIGHT)

    Set this.co = co
End Sub

'@label:addseries
'@sub-title Attach a data series from a named range
'@details
'One graph can have multiple series added. Series are referenced using
'named ranges on the graph object worksheet. If the named range does not
'exist, the method files a warning, adds no series, and marks the attempt as
'failed so the labels that follow are skipped too. If no chart exists
'yet, Add is called lazily before attaching the series.
'@param rngName String. Named range containing the series values.
'@param chrtType String. Chart type: "line", "point", "bar", or "hbar". Defaults to "bar" for unknown values.
'@param axisPos Optional String. Axis position: "right" adds a secondary axis; any other value uses the primary axis. Defaults to vbNullString.
'@export
Public Sub AddSeries(ByVal rngName As String, ByVal chrtType As String, _
                    Optional ByVal axisPos As String = vbNullString)

    Dim co As ChartObject
    Dim srs As Series
    Dim seriesRng As Range
    Dim customType As XlChartType

    this.lastSeriesOk = False

    'If the provided range does not exist, exit the sub
    If Not TryRange(rngName, seriesRng) Then
        LogInfo "Worksheet " & Wksh().Name & " carries no range named " & rngName & _
                ", so this series is left out of the graph", checkingWarning
        Exit Sub
    End If

    'If no graph exists, add one first
    Set co = graphChart()
    If co Is Nothing Then
        Add
        Set co = graphChart()
    End If

    Select Case chrtType
        Case "line": customType = xlLineMarkers
        Case "point": customType = xlXYScatter
        Case "bar": customType = xlColumnClustered
        Case "hbar": customType = xlBarClustered
        Case Else: customType = xlColumnClustered
    End Select

    co.Chart.SeriesCollection.Add _
        Source:=seriesRng, RowCol:=xlColumns, _
        SeriesLabels:=False, CategoryLabels:=False, Replace:=False

    Set srs = LastSeries()
    If srs Is Nothing Then
        LogInfo "Excel accepted no series for " & rngName & _
                ", so this series is left out of the graph", checkingWarning
        Exit Sub
    End If

    'The values are written again, on purpose. Add guesses how to split a
    'Source whose cells are mostly empty, whatever the two label arguments
    'say, and a time series column with data in its last row alone came out
    'as a one-cell series: the chart showed one bar over the first period.
    'This line pins the whole range as the data. The range is already in
    'hand, so the write costs one chart recalculation and no name resolution.
    srs.Values = seriesRng

    'Add a right axis if required
    If axisPos = "right" Then
        co.Chart.HasAxis(xlValue, xlSecondary) = True
        srs.AxisGroup = xlSecondary
    End If

    srs.chartType = customType

    this.lastSeriesOk = True
End Sub


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

    Dim srs As Series
    Dim labelsObj As DataLabels
    Dim catRng As Range
    Dim lblRng As Range
    Dim hasLabelRange As Boolean

    If graphChart() Is Nothing Then
        LogInfo "Labels " & catName & " were asked for before any chart was made", _
                checkingWarning
        Exit Sub
    End If

    Set srs = LastSeries()
    If srs Is Nothing Then
        LogInfo "Labels " & catName & " were asked for and this graph carries no series", _
                checkingWarning
        Exit Sub
    End If

    If Not this.lastSeriesOk Then
        LogInfo "Labels " & catName & " belong to a series that was left out, " & _
                "so they are skipped", checkingWarning
        Exit Sub
    End If

    hasLabelRange = TryRange(lblName, lblRng)

    If TryRange(catName, catRng) Then srs.XValues = catRng

    'Labels can be hardcoded or dynamic; by default, they are hardcoded
    If (Not hardCodeLabels) And hasLabelRange Then
        'The user can modify by changing values in the range, so the legend
        'entry is written as a reference
        srs.Name = "=" & QualifiedAddress(lblRng)
    Else
        srs.Name = LegendText(prefix, prefixOnly, hasLabelRange, lblRng)
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
