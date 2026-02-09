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

' Graph specs creation mode
Public Enum GraphSpecsType
    GraphTypeSimple = 1                          'Non-time-series (univariate, bivariate, spatial, spatio-temporal)
    GraphTypeComplex = 2                         'Time series (multiple graphs from setup listobjects)
End Enum

'@section Series access (simple mode)
'===============================================================================

' @description Number of series in the current graph (simple mode only)
Public Property Get NumberOfSeries() As Long: End Property

' @description Named range name for series data at given index
Public Property Get SeriesName(ByVal index As Long) As String: End Property

' @description Chart type for series at given index ("bar", "hbar", "point")
Public Property Get SeriesType(ByVal index As Long) As String: End Property

' @description Axis position for series at given index ("left", "right")
Public Property Get SeriesPos(ByVal index As Long) As String: End Property

' @description Row categories range name for series at given index
Public Property Get SeriesLabel(ByVal index As Long) As String: End Property

' @description Column label range name for series at given index
Public Property Get SeriesColumnLabel(ByVal index As Long) As String: End Property

' @description Label prefix for series at given index (complex mode)
Public Property Get SeriesLabelPrefix(ByVal index As Long) As String: End Property

'@section Multi-graph access (complex mode)
'===============================================================================

' @description Number of graphs (complex mode only)
Public Property Get NumberOfGraphs() As Long: End Property

' @description Nested BetterArray of all graph specs (complex mode).
'              Each graph is a BetterArray of 7 BetterArrays:
'              [series, types, positions, labels, catLabels, prefixes, title]
Public Property Get SpecsLists() As BetterArray: End Property

' @description List of unique graph IDs (complex mode)
Public Property Get GraphIDs() As BetterArray: End Property

'@section State
'===============================================================================

' @description Whether this graph spec is valid and has series to plot
Public Property Get Valid() As Boolean: End Property

' @description Output worksheet where the graph will be drawn
Public Property Get Wksh() As Worksheet: End Property

'@section Operations
'===============================================================================

' @description Build all series from the specifications.
'              Must be called before accessing series properties.
Public Sub CreateSeries(): End Sub
