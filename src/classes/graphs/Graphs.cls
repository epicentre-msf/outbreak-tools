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
Implements IGraphs

'@class Graphs
'@description
'Simplifies the creation and formatting of Excel ChartObject instances on
'analysis worksheets. The class wraps a single chart, allowing callers to
'attach data series from named ranges, assign category and legend labels,
'and apply scope-dependent formatting (time series, spatial, or standard).
'Charts are positioned at a caller-supplied Range and sized according to
'the GraphScope enum. Consumers interact through the IGraphs interface.
'@depends Checking
'@version 1.0 (2026-02-09)



'@section State
'===============================================================================
'@description Backing UDT, module-level checking variables, and chart state.

Private Type TGraph
    sh As Worksheet
    strtRng As Range
    co As ChartObject
    Index As Long
    grName As String
    hasChecks As Boolean
End Type

Private this As TGraph

'Checking objects
Private checkCounter As Long
Private internalChecks As IChecking


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
'@param sh Worksheet. The worksheet containing the values to be plotted.
'@param posRng Range. The cell or range determining where to position the graph.
'@param grName Optional String. Display name for the graph. Defaults to vbNullString.
'@return IGraphs. A ready-to-use graph instance.
'@throws ProjectError.ObjectNotInitialized When sh is Nothing.
'@throws ProjectError.ObjectNotInitialized When posRng is Nothing.
'@depends Checking
'@export
Public Function Create(sh As Worksheet, posRng As Range, Optional ByVal grName As String = vbNullString) As IGraphs

    If (sh Is Nothing) Then ThrowError ProjectError.ObjectNotInitialized, "Worksheet not initialized, unable to create graph"
    If (posRng Is Nothing) Then ThrowError ProjectError.ObjectNotInitialized, "Range not initialized, unable to create graph"

    With New Graphs
        Set .Wksh = sh
        Set .StartRange = posRng
        .Name = grName
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Current object instance
'@details
'Returns the current instance cast to its interface. Used internally
'by the factory method to return an IGraphs reference.
'@return IGraphs. Reference to the current object.
Public Property Get Self() As IGraphs
    Set Self = Me
End Property

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
    Set this.strtRng = strtRng
End Property

'@label:name-set
'@prop-title Assign the graph display name
'@param grName String. The graph name.
Public Property Let Name(ByVal grName As String)
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
'@description Internal chart object and series index state.

'@label:graph-chart
'@prop-title Retrieve the internal ChartObject
'@details
'Returns the ChartObject created by Add. Nothing until Add is called.
'@return ChartObject. The underlying Excel chart object.
Private Property Get graphChart() As ChartObject
    Set graphChart = this.co
End Property

'@label:series-index
'@prop-title Latest series index
'@details
'Tracks the 1-based index of the most recently added series.
'Incremented by AddSeries and used by AddLabels to target the
'correct series entry.
'@return Long. Current series index.
Private Property Get seriesIndex() As Long
    seriesIndex = this.Index
End Property

'@label:series-index-set
'@prop-title Update the latest series index
'@param ind Long. New index value.
Private Property Let seriesIndex(ByVal ind As Long)
    this.Index = ind
End Property


'@section Operations
'===============================================================================
'@description Chart creation, series attachment, labelling, and formatting.

'@label:add
'@sub-title Create an empty chart on the worksheet
'@details
'Creates an empty graph on a worksheet, ready for data to be
'attached to the graph. The graph is empty at creation. The width and height
'of the graph object may change depending on the scope in the future.
'@export
Private Sub Add()
    'Upon creation, the width and height of graphs are constant

    Const CW As Long = 488 'graph width
    Const RH As Long = 214  'graph height

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim rng As Range

    Set rng = StartRange()
    Set sh = rng.Parent

    Set co = sh.ChartObjects.Add(rng.Left, rng.Top, CW, RH)

    co.Left = rng.Left
    co.Top = rng.Top

    With co.Chart
        .PlotArea.Interior.color = RGB(235, 235, 235)
        .Legend.Position = xlLegendPositionTop
    End With

    Set this.co = co
End Sub

'@label:addseries
'@sub-title Attach a data series from a named range
'@details
'One graph can have multiple series added. Series are referenced using
'named ranges on the graph object worksheet. If the named range does not
'exist, the method logs a warning and exits silently. If no chart exists
'yet, Add is called lazily before attaching the series.
'@param rngName String. Named range containing the series values.
'@param chrtType String. Chart type: "line", "point", "bar", or "hbar". Defaults to "bar" for unknown values.
'@param axisPos Optional String. Axis position: "right" adds a secondary axis; any other value uses the primary axis. Defaults to vbNullString.
'@export
Private Sub AddSeries(ByVal rngName As String, ByVal chrtType As String, _
                    Optional axisPos As String = vbNullString)

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim ind As Long
    Dim customType As XlChartType
    Dim axGrp As XlAxisGroup

    'The index is initialized at 0 by default
    Set co = graphChart()
    Set sh = Wksh()

    'If the provided range does not exist, exit the sub
    If Not RangeExist(rngName) Then
        LogInfo "Attempt to add series " & rngName & " which does not exists in worksheet " & sh.Name, checkingWarning
        Exit Sub
    End If

    'If no graph exists, add one first
    If co Is Nothing Then
        Add
        Set co = graphChart()
    End If

    ind = seriesIndex() + 1

    Select Case chrtType
        Case "line": customType = xlLineMarkers
        Case "point": customType = xlXYScatter
        Case "bar": customType = xlColumnClustered
        Case "hbar": customType = xlBarClustered
        Case Else: customType = xlColumnClustered
    End Select

    axGrp = IIf(axisPos = "right", xlSecondary, xlPrimary)

    With co.Chart
        .SeriesCollection.Add _
        Source:=sh.Range(rngName), RowCol:=xlColumns, _
        CategoryLabels:=False, Replace:=False
        .SeriesCollection(ind).Values = sh.Range(rngName)

        'Add a right axis if required
        If axGrp = xlSecondary Then
            .HasAxis(xlValue, xlSecondary) = True
            .SeriesCollection(ind).AxisGroup = axGrp
        End If

        .SeriesCollection(ind).chartType = customType
    End With

    seriesIndex = ind
End Sub


'@label:addlabels
'@sub-title Assign category and legend labels to the current series
'@details
'Sets the category axis values and legend entry for the most recently
'added series. The optional prefix is prepended to the legend name,
'separated by " - " unless prefixOnly is True. When hardCodeLabels is
'False, the legend entry references the named range dynamically rather
'than using a literal string.
'@param catName String. Named range for category axis labels.
'@param lblName String. Named range for the legend entry text.
'@param prefix Optional String. Display prefix prepended to the legend name. Defaults to vbNullString.
'@param prefixOnly Optional Boolean. When True, uses only the prefix as the legend entry. Defaults to False.
'@param hardCodeLabels Optional Boolean. When True, legend text is literal; when False, references the named range. Defaults to True.
Private Sub AddLabels(ByVal catName As String, _
                     ByVal lblName As String, _
                     Optional ByVal prefix As String = vbNullString, _
                     Optional ByVal prefixOnly As Boolean = False, _
                     Optional ByVal hardCodeLabels As Boolean = True)

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim labelValue As String
    Dim ind As Long
    Dim currPrefix As String

    Set co = graphChart
    Set sh = Wksh()

    'Labels are added to the current index
    ind = seriesIndex()

    If ind <= 0 Then
        LogInfo "Attempt to add labels, but there a no series on this graph", checkingWarning
        Exit Sub
    End If

    'Add a prefix to the label name (useful for time series graphs)
    currPrefix = prefix
    If currPrefix <> vbNullString And (Not prefixOnly) Then currPrefix = currPrefix & " - "

    If prefixOnly Then
        labelValue = currPrefix
    ElseIf RangeExist(lblName) Then
        labelValue = currPrefix & CStr(sh.Range(lblName).Value)
    Else
        labelValue = currPrefix
    End If

    'Throw an error when co is Nothing
    With co.Chart
        If RangeExist(catName) Then .SeriesCollection(ind).XValues = sh.Range(catName)
        'Labels can be hardcoded or dynamic; by default, they are hardcoded
        If (Not hardCodeLabels) And RangeExist(lblName) Then
            'The user can modify by changing values in the range, but
            'You have to assume the range exists and is not empty
            .SeriesCollection(ind).Name = "= " & RngAddress(lblName)
        Else
            .SeriesCollection(ind).Name = labelValue
        End If

        .SeriesCollection(ind).ApplyDataLabels xlDataLabelsShowValue, False
        .SeriesCollection(ind).DataLabels.Font.Size = 6
        .SeriesCollection(ind).DataLabels.Font.color = RGB(2, 6, 91)
        .SeriesCollection(ind).DataLabels.Font.Name = "+mj-lt"
    End With
End Sub

'@label:format
'@sub-title Apply layout and formatting to the chart
'@details
'Formats are defined based on the scope of the graph. For time series, graphs are larger;
'for spatio-temporal analyses, they are wider. The scope controls dimension coefficients,
'axis gridline colours, legend placement, and plot area reverse ordering. When
'hardCodeLabels is False, axis titles and the chart title reference named ranges
'dynamically rather than using literal strings.
'@param valuesTitle Optional String. Title for the value axis (Y-axis). Defaults to vbNullString.
'@param catTitle Optional String. Title for the category axis (X-axis). Defaults to vbNullString.
'@param plotTitle Optional String. Title displayed above the chart. Defaults to vbNullString.
'@param secondAxisTitle Optional String. Title for the secondary value axis. Defaults to vbNullString.
'@param scope Optional Byte. GraphScope value controlling layout sizing. Defaults to GraphScopeNormal.
'@param heightFactor Optional Long. Multiplicative factor applied to chart height. Defaults to 1.
'@param hardCodeLabels Optional Boolean. When True, axis titles are literal text. Defaults to True.
Private Sub Format(Optional ByVal valuesTitle As String = vbNullString, _
                  Optional ByVal catTitle As String = vbNullString, _
                  Optional ByVal plotTitle As String = vbNullString, _
                  Optional ByVal secondAxisTitle As String = vbNullString, _
                  Optional ByVal scope As Byte = GraphScopeNormal, _
                  Optional ByVal heightFactor As Long = 1, _
                  Optional ByVal hardCodeLabels As Boolean = True)

    Const CW As Long = 488
    Const RH As Long = 214

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim coef As Long
    Dim heightCoef As Long

    Set co = graphChart
    Set sh = Wksh()

    coef = 1
    If (scope = GraphScopeTimeSeries) Then coef = 1.75
    heightCoef = IIf(heightFactor <= 1, 1, (heightFactor + 1) * 0.08)

    'Reapply the width and height
    co.width = CW * coef
    co.height = RH * coef * heightCoef

    With sh.Shapes(co.Name)
        .Shadow.Type = msoShadow21
        .line.Visible = msoTrue
        .line.ForeColor.RGB = RGB(173, 216, 230)
        'Avoid changing graph width when adjusting row height
        .Placement = xlMove
    End With

    With co.Chart

        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .HasMajorGridlines = True
            .MajorGridlines.Border.color = RGB(255, 255, 255)
            'If not hardcoded, titles may be range names. Test their existence
            'and use references instead of literal values
            If (Not hardCodeLabels) And RangeExist(catTitle) Then
                .AxisTitle.Caption = "= " & RngAddress(catTitle)
            Else
                .AxisTitle.Caption = catTitle
            End If

            .TickLabels.Font.Size = 10

            'Reverse plot order for spatial analysis
            If (scope = GraphScopeSpatial) Then .ReversePlotOrder = True
        End With

        'Formatting on primary axis
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .HasMajorGridlines = True
            .MajorGridlines.Border.color = RGB(255, 255, 255)
            .TickLabels.Font.Size = 10
            .TickLabels.Font.Name = "+mj-lt"

            If (Not hardCodeLabels) And RangeExist(valuesTitle) Then
                .AxisTitle.Caption = "= " & RngAddress(valuesTitle)
            Else
                .AxisTitle.Caption = valuesTitle
            End If
        End With

        'Add a title if needed
        .HasTitle = False
        If plotTitle <> vbNullString Then
            .HasTitle = True
            With .ChartTitle
                'You can add either address or value of plot title
                If (Not hardCodeLabels) And RangeExist(plotTitle) Then
                    .Caption = "= " & RngAddress(plotTitle)
                Else
                    .Caption = plotTitle
                End If
                .Font.Size = 14
                .Font.Bold = True
                .Font.Name = "+mj-lt"
            End With
        End If

        'Formatting on the secondary axis
        If .HasAxis(xlValue, xlSecondary) Then
            With .Axes(xlValue, xlSecondary)
                .HasTitle = True
                .HasMajorGridlines = False
                .MajorGridlines.Border.color = RGB(255, 255, 255)
                .TickLabels.Font.Size = 10
                .TickLabels.Font.Name = "+mj-lt"
                .AxisTitle.Text = "%"
                If secondAxisTitle <> vbNullString Then .AxisTitle.Text = secondAxisTitle
                .MaximumScale = 1
                .MajorUnit = 0.1
            End With
        End If

        With .PlotArea

            If scope = GraphScopeTimeSeries Then
                .Left = 25 * coef
            Else
                .Left = 16.618 * coef
            End If

            .Top = 16.618 * coef
            .width = 425 * coef
            .height = 178 * coef * heightCoef
        End With

        'Legend position
        If (scope = GraphScopeSpatial) Then
            .Legend.Position = xlLegendPositionBottom
        End If
    End With

End Sub


'@section Helpers
'===============================================================================
'@description Private utility functions for range validation and address resolution.

'@label:range-exist
'@sub-title Test whether a named range exists on the host worksheet
'@details
'Uses On Error Resume Next to attempt resolving the named range.
'Returns True when the range is valid and non-Nothing.
'@param rngName String. Named range to check.
'@return Boolean. True when the named range exists.
Private Function RangeExist(ByVal rngName As String) As Boolean
    Dim sh As Worksheet
    Dim rng As Range
    Set sh = Wksh()
    On Error Resume Next
        Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExist = Not (rng Is Nothing)
End Function

'@label:rng-address
'@sub-title Return the full worksheet-qualified address of a named range
'@details
'Resolves the named range on the host worksheet and returns its address
'in the format 'SheetName'!$A$1. Returns vbNullString when the range
'does not exist.
'@param rngName String. Named range to resolve.
'@return String. Worksheet-qualified address or vbNullString.
Private Function RngAddress(ByVal rngName As String) As String
    Dim sh As Worksheet
    Dim rng As Range

    If Not RangeExist(rngName) Then Exit Function 'address can be empty
    Set sh = Wksh()
    Set rng = sh.Range(rngName)

    RngAddress = "'" & sh.Name & "'" & "!" & rng.Address
End Function


'@section Checking
'===============================================================================
'@description Logging and validation support via the IChecking interface.

'@label:log-info
'@sub-title Record a diagnostic entry in the internal checking store
'@details
'Lazily creates the Checking instance on first use and appends the
'message using a monotonically increasing counter key. The checking
'store is exposed through HasCheckings and CheckingValues.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level from CheckingLogType. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "CustomTable " & Name() & " object trace"

    'First, instantiate the internalChecks object
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)

    'Add new information to the checking object
    internalChecks.Add Cstr(checkCounter), label, scope
    checkCounter = checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:has-checkings
'@prop-title Determine whether diagnostics have been captured
'@details
'Returns True when at least one LogInfo call has been made,
'indicating that diagnostic information is available.
'@return Boolean. True when internal checks exist.
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the collected checking entries
'@details
'Returns the internal IChecking store when at least one entry has
'been recorded. Returns Nothing otherwise.
'@return Object. The IChecking instance with logged diagnostics.
'@export
Private Property Get CheckingValues() As Object

    'If there are no checks, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
End Property


'@section ErrorHandling
'===============================================================================
'@description Centralised error-raising and debug helpers.

'@label:show-debug
'@sub-title Print debug information to the Immediate window
'@details
'Outputs basic debug information including the graph name.
Private Sub ShowDebug()
    Debug.Print "Check the informations of choice object " & Name()
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source for consistent
'stack traces across all methods in this class.
'@param errNumb Long. ProjectError code describing the failure.
'@param errorMessage String. Descriptive error message.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "DataSheet", "Error : " & Chr(34) & errorMessage
End Sub


'@section Interface
'===============================================================================
'@description Delegated members satisfying the IGraphs contract.

Private Sub IGraphs_Add()
    Add
End Sub

Private Sub IGraphs_AddSeries(ByVal rngName As String, ByVal chrtType As String, Optional axisPos As String = vbNullString)
    AddSeries rngName, chrtType, axisPos
End Sub

Private Sub IGraphs_AddLabels(ByVal catName As String, ByVal lblName As String, Optional ByVal prefix As String = vbNullString, _
                            Optional ByVal prefixOnly As Boolean = False, Optional ByVal hardCodeLabels As Boolean = True)
    AddLabels catName, lblName, prefix, prefixOnly, hardCodeLabels
End Sub

Private Sub IGraphs_Format(Optional ByVal valuesTitle As String = vbNullString, _
                          Optional ByVal catTitle As String = vbNullString, _
                          Optional ByVal plotTitle As String = vbNullString, _
                          Optional ByVal scope As Byte = GraphScopeNormal, _
                          Optional ByVal heightFactor As Long = 1, _
                          Optional ByVal hardCodeLabels As Boolean = True)
    Format valuesTitle:=valuesTitle, catTitle:=catTitle, plotTitle:=plotTitle, scope:=scope, _
           heightFactor:=heightFactor, hardCodeLabels:=hardCodeLabels
End Sub
