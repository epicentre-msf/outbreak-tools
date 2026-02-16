VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGraphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Graph object"


'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface of Graph object")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IGraphs
'@description
'Defines the public contract for the Graphs class, exposing methods to
'create an empty chart on a worksheet, attach data series from named
'ranges, assign axis labels, and apply formatting based on graph scope.
'The GraphScope enum controls layout sizing for time series, spatial,
'and standard analysis charts.

'Differents scopes for graphs. Depending on the scope, the size and other
'adjustments are made. For example, on Time Series, the graph will appear
'bigger, compare to basic normal graph.

Public Enum GraphScope
  GraphScopeTimeSeries = 1 'A time series graph (long graph)
  GraphScopeSpatial = 2 'A spatial analysis graph
  GraphScopeNormal = 3 'A standard univariate / bivariate graph
End Enum


'@section Operations
'===============================================================================
'@description Chart creation, series attachment, labelling, and formatting.

'@jump:add
'@sub-title Create an empty chart object on the worksheet
'@details
'Initialises an empty chart at the position defined during factory
'construction. The chart is ready to receive series via AddSeries.
'Must be called before AddSeries if no chart exists yet; AddSeries
'also calls Add lazily when needed.
Public Sub Add(): End Sub

'@jump:addseries
'@sub-title Attach a data series from a named range
'@details
'Adds one chart series sourced from the named range on the host
'worksheet. The chart type controls the rendering style and the axis
'position determines whether the series uses the primary or secondary
'value axis. Call AddLabels immediately after to keep the parallel
'label arrays synchronised.
'@param rngName String. Named range containing the series values.
'@param chrtType String. Rendering style: "bar", "hbar", "line", or "point".
'@param axisPos Optional String. Axis placement: "left" (primary) or "right" (secondary). Defaults to vbNullString.
Public Sub AddSeries(ByVal rngName As String, ByVal chrtType As String, _
                     Optional axisPos As String = vbNullString): End Sub

'@jump:addlabels
'@sub-title Assign category and legend labels to the current series
'@details
'Sets the category axis values and legend entry for the most recently
'added series. The optional prefix is prepended to the legend name,
'useful for time series graphs where a user-defined label augments
'the column header.
'@param catName String. Named range for category axis labels.
'@param lblName String. Named range for the legend entry text.
'@param prefix Optional String. Display prefix prepended to the legend name. Defaults to vbNullString.
'@param prefixOnly Optional Boolean. When True, uses only the prefix as the legend entry. Defaults to False.
'@param hardCodeLabels Optional Boolean. When True, legend text is literal; when False, references the named range. Defaults to True.
Public Sub AddLabels(ByVal catName As String, ByVal lblName As String, _
                     Optional ByVal prefix As String = vbNullString, _
                     Optional ByVal prefixOnly As Boolean = False, _
                     Optional ByVal hardCodeLabels As Boolean = True): End Sub

'@jump:format
'@sub-title Apply layout and formatting to the chart
'@details
'Adjusts chart dimensions, axis titles, gridlines, legend position,
'and plot area sizing based on the supplied scope. Time series scope
'produces larger charts; spatial scope reverses the category axis plot
'order and moves the legend to the bottom.
'@param valuesTitle Optional String. Title for the value axis (Y-axis). Defaults to vbNullString.
'@param catTitle Optional String. Title for the category axis (X-axis). Defaults to vbNullString.
'@param plotTitle Optional String. Chart title displayed above the plot area. Defaults to vbNullString.
'@param scope Optional Byte. GraphScope value controlling layout sizing. Defaults to GraphScopeNormal.
'@param heightFactor Optional Long. Multiplicative factor for chart height. Defaults to 1.
'@param hardCodeLabels Optional Boolean. When True, axis titles are literal text. Defaults to True.
Public Sub Format(Optional ByVal valuesTitle As String = vbNullString, _
                  Optional ByVal catTitle As String = vbNullString, _
                  Optional ByVal plotTitle As String = vbNullString, _
                  Optional ByVal scope As Byte = GraphScopeNormal, _
                  Optional ByVal heightFactor As Long = 1, _
                  Optional ByVal hardCodeLabels As Boolean = True)
End Sub
