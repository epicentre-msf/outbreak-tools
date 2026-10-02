VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGraphSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for graph series specifications"

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface for graph series specifications")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IGraphSpecs
'@description
'Defines the public contract for the GraphSpecs class, exposing series
'accessors for simple (non-time-series) charts and a nested collection
'of graph specification entries for complex (time series) charts.
'Consumers call CreateSeries to populate the series data, then read
'individual series properties or the SpecsLists collection depending
'on the graph mode.

' Graph specs creation mode
Public Enum GraphSpecsType
    GraphTypeSimple = 1                          'Non-time-series (univariate, bivariate, spatial, spatio-temporal)
    GraphTypeComplex = 2                         'Time series (multiple graphs from setup listobjects)
End Enum


'@section SeriesAccess
'===============================================================================
'@description Series accessors for simple mode (non-time-series charts).

'@jump:number-of-series
'@prop-title Number of chart series in simple mode
'@details
'Returns the count of series pushed during simple mode building.
'Always returns 0 in complex mode or before CreateSeries is called.
'@return Long. Count of series entries (1-based).
Public Property Get NumberOfSeries() As Long: End Property

'@jump:series-name
'@prop-title Named range reference for series data at a given index
'@details
'Returns the Excel named range string (e.g. "VALUES_COL_1_table1")
'that a chart builder uses as the data source for this series.
'@param index Long. 1-based index of the series (1 to NumberOfSeries).
'@return String. Named range reference for the series data.
Public Property Get SeriesName(ByVal index As Long) As String: End Property

'@jump:series-type
'@prop-title Chart rendering type for a series at a given index
'@details
'Returns the chart type string controlling how the series is rendered.
'Common values are "bar", "hbar", "point", and "line".
'@param index Long. 1-based series index.
'@return String. Chart rendering type.
Public Property Get SeriesType(ByVal index As Long) As String: End Property

'@jump:series-pos
'@prop-title Axis position for a series at a given index
'@details
'Returns "left" for the primary value axis or "right" for the
'secondary axis, typically used for percentage overlays.
'@param index Long. 1-based series index.
'@return String. Axis placement ("left" or "right").
Public Property Get SeriesPos(ByVal index As Long) As String: End Property

'@jump:series-label
'@prop-title Row category named range for a series at a given index
'@details
'Returns the named range providing category axis labels (e.g.
'"ROW_CATEGORIES_table1") displayed along the chart axis.
'@param index Long. 1-based series index.
'@return String. Named range for row category labels.
Public Property Get SeriesLabel(ByVal index As Long) As String: End Property

'@jump:series-column-label
'@prop-title Column label named range for a series at a given index
'@details
'Returns the named range providing the legend entry text (e.g.
'"LABEL_COL_2_table1") that identifies this series in the chart legend.
'@param index Long. 1-based series index.
'@return String. Named range for the legend entry.
Public Property Get SeriesColumnLabel(ByVal index As Long) As String: End Property

'@jump:series-label-prefix
'@prop-title Label prefix for a series at a given index
'@details
'Returns the optional display prefix prepended to the series legend
'entry. Always vbNullString in simple mode; populated in complex mode
'from the user-defined "label" column of the graph setup table.
'@param index Long. 1-based series index.
'@return String. Prefix string or vbNullString when none is set.
Public Property Get SeriesLabelPrefix(ByVal index As Long) As String: End Property


'@section ComplexAccess
'===============================================================================
'@description Multi-graph accessors for complex (time series) mode.

'@jump:number-of-graphs
'@prop-title Number of distinct graphs in complex mode
'@details
'Returns the count of unique graph IDs whose series were successfully
'built. Returns 0 in simple mode or before CreateSeries is called.
'@return Long. Count of graph entries in the SpecsLists collection.
Public Property Get NumberOfGraphs() As Long: End Property

'@jump:specs-lists
'@prop-title Nested collection of all graph specifications
'@details
'Each element is a 7-element BetterArray representing one graph:
'[seriesNames, seriesTypes, seriesPositions, seriesLabels,
'seriesColumnLabels, seriesLabelPrefixes, titleString]. Returns
'an empty BetterArray in simple mode.
'@return BetterArray. Cloned collection of graph spec entries.
Public Property Get SpecsLists() As BetterArray: End Property

'@jump:graph-ids
'@prop-title Unique graph identifiers for complex mode
'@details
'Returns a deduplicated list of graph ID strings in the order of
'first appearance in the graph setup ListObject. Useful for iterating
'over graphs independently of SpecsLists.
'@return BetterArray. Cloned list of unique graph ID strings.
Public Property Get GraphIDs() As BetterArray: End Property


'@section State
'===============================================================================
'@description Validity and worksheet accessors.

'@jump:valid
'@prop-title Whether the graph specifications were successfully built
'@details
'Triggers a lazy CreateSeries call on first access. In simple mode
'always returns True; in complex mode returns True only when at
'least one graph produced valid series definitions.
'@return Boolean. True when series data is available for consumption.
Public Property Get Valid() As Boolean: End Property

'@jump:wksh
'@prop-title Output worksheet where the chart will be drawn
'@details
'In simple mode returns the cross-table worksheet; in complex mode
'returns the dedicated output worksheet supplied during construction.
'@return Worksheet. Target worksheet for chart creation.
Public Property Get Wksh() As Worksheet: End Property


'@section Operations
'===============================================================================
'@description Series building entry point.

'@jump:create-series
'@sub-title Build all series from the stored specifications
'@details
'Dispatches to the appropriate builder based on the graph mode.
'Must be called before accessing series properties or SpecsLists.
'Also called lazily by the Valid property on first access.
Public Sub CreateSeries(): End Sub
