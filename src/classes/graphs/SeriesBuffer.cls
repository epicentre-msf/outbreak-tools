VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeriesBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "One chart's worth of series definitions"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("One chart's worth of series definitions")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class SeriesBuffer
'@description
'One chart's worth of series definitions. Entry N describes one complete series:
'the data range name, the chart type, the axis side, the category range name,
'the legend range name, and the display prefix. Both graph builders fill one of
'these and every consumer walks it by index.
'
'ONE ADD, SO THE SIX VALUES CANNOT COME APART
'-------------------------------------------------------------------------------
'The builder that came before this held six parallel BetterArrays and pushed
'into them through two routines, PushSeries and PushLabels. Nothing made the two
'calls happen together; a doc comment asked the caller to remember. Six values
'that describe one series arrive in one Add here, and an entry is one array of
'six, so a half-written series cannot exist.
'@depends BetterArray, Checking
'@version 1.0 (2026-07-31)



'@section State
'===============================================================================
'@description The entry list and the position of each value inside an entry.

Private Type TSeriesBuffer
    entries As BetterArray
End Type

Private this As TSeriesBuffer

Private Const CLASS_NAME As String = "SeriesBuffer"

' Where each value sits inside one entry. An entry is a six-element zero-based
' Variant array, which is what Array() hands back.
Private Const FIELD_RANGE_NAME As Long = 0
Private Const FIELD_CHART_TYPE As Long = 1
Private Const FIELD_AXIS_SIDE As Long = 2
Private Const FIELD_CATEGORY_RANGE As Long = 3
Private Const FIELD_LEGEND_RANGE As Long = 4
Private Const FIELD_LABEL_PREFIX As Long = 5


'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create an empty series buffer
'@details
'The buffer starts with no entry. Every builder makes one per chart, so a chart
'that ends with no series answers a Count of 0 and the caller draws nothing.
'@return SeriesBuffer. An empty buffer ready for Add.
'@export
Public Function Create() As SeriesBuffer
    Dim newBuffer As SeriesBuffer

    Set newBuffer = New SeriesBuffer
    newBuffer.Initialise
    Set Create = newBuffer
End Function

'@label:initialise
'@sub-title Build the entry list
'@details
'Public because the factory calls it on the instance it just made. The list is
'1-based, which is what every consumer walks.
Public Sub Initialise()
    Set this.entries = New BetterArray
    this.entries.LowerBound = 1
End Sub


'@section Operations
'===============================================================================

'@label:add
'@sub-title Append one complete series to the buffer
'@details
'The six values that describe a series arrive together, so index N of the buffer
'always holds a whole series.
'@param rangeName String. Named range holding the series data.
'@param chartType String. How the series is drawn: "bar", "hbar", "point" or "line".
'@param axisSide String. "left" for the primary axis, "right" for the secondary one.
'@param categoryRange String. Named range holding the category axis labels.
'@param legendRange String. Named range holding the legend entry text.
'@param labelPrefix Optional String. Text put in front of the legend entry.
'@export
Public Sub Add(ByVal rangeName As String, _
               ByVal chartType As String, _
               ByVal axisSide As String, _
               ByVal categoryRange As String, _
               ByVal legendRange As String, _
               Optional ByVal labelPrefix As String = vbNullString)

    If this.entries Is Nothing Then Initialise

    this.entries.Push Array(rangeName, chartType, axisSide, _
                            categoryRange, legendRange, labelPrefix)
End Sub


'@section Reading
'===============================================================================

'@label:count
'@prop-title How many series the buffer holds
'@return Long. The number of entries.
'@export
Public Property Get Count() As Long
    If this.entries Is Nothing Then Exit Property
    Count = this.entries.Length
End Property

'@label:range-name
'@prop-title Named range holding the data of one series
'@param index Long. 1-based position in the buffer.
'@return String. The data range name.
'@throws ProjectError.InvalidArgument When index is outside the buffer.
'@export
Public Property Get RangeName(ByVal index As Long) As String
    RangeName = FieldOf(index, FIELD_RANGE_NAME)
End Property

'@label:chart-type
'@prop-title How one series is drawn
'@param index Long. 1-based position in the buffer.
'@return String. "bar", "hbar", "point" or "line".
'@throws ProjectError.InvalidArgument When index is outside the buffer.
'@export
Public Property Get ChartType(ByVal index As Long) As String
    ChartType = FieldOf(index, FIELD_CHART_TYPE)
End Property

'@label:axis-side
'@prop-title Which axis one series is plotted against
'@param index Long. 1-based position in the buffer.
'@return String. "left" or "right".
'@throws ProjectError.InvalidArgument When index is outside the buffer.
'@export
Public Property Get AxisSide(ByVal index As Long) As String
    AxisSide = FieldOf(index, FIELD_AXIS_SIDE)
End Property

'@label:category-range
'@prop-title Named range holding the category axis labels of one series
'@param index Long. 1-based position in the buffer.
'@return String. The category range name.
'@throws ProjectError.InvalidArgument When index is outside the buffer.
'@export
Public Property Get CategoryRange(ByVal index As Long) As String
    CategoryRange = FieldOf(index, FIELD_CATEGORY_RANGE)
End Property

'@label:legend-range
'@prop-title Named range holding the legend entry of one series
'@param index Long. 1-based position in the buffer.
'@return String. The legend range name.
'@throws ProjectError.InvalidArgument When index is outside the buffer.
'@export
Public Property Get LegendRange(ByVal index As Long) As String
    LegendRange = FieldOf(index, FIELD_LEGEND_RANGE)
End Property

'@label:label-prefix
'@prop-title Text put in front of the legend entry of one series
'@details
'A time series series is labelled by its prefix alone. The other builders leave
'it empty.
'@param index Long. 1-based position in the buffer.
'@return String. The prefix, or an empty string.
'@throws ProjectError.InvalidArgument When index is outside the buffer.
'@export
Public Property Get LabelPrefix(ByVal index As Long) As String
    LabelPrefix = FieldOf(index, FIELD_LABEL_PREFIX)
End Property


'@section Helpers
'===============================================================================

'@label:field-of
'@sub-title Read one value of one entry
'@param index Long. 1-based position in the buffer.
'@param field Long. Which of the six values is wanted.
'@return String. The value, as text.
'@throws ProjectError.InvalidArgument When index is outside the buffer.
Private Function FieldOf(ByVal index As Long, ByVal field As Long) As String
    Dim entry As Variant

    ValidateIndex index

    ' The position is read from the lower bound of the entry itself. Array()
    ' hands back a zero-based array, and BetterArray re-bases what it stores to
    ' its own lower bound of 1, so a fixed 0 read one field short of the end and
    ' every other field answered its neighbour.
    entry = this.entries.Item(index)
    FieldOf = CStr(entry(LBound(entry) + field))
End Function

'@label:validate-index
'@sub-title Refuse a position the buffer does not hold
'@param index Long. The position to test.
'@throws ProjectError.InvalidArgument When index is outside the buffer.
Private Sub ValidateIndex(ByVal index As Long)
    If (index < 1) Or (index > Count()) Then _
        ThrowError ProjectError.InvalidArgument, _
                   "Series index " & index & " out of bounds (1 to " & Count() & ")"
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber Long. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
