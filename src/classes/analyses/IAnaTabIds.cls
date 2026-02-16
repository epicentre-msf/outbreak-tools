VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnaTabIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for analysis table and graph ID tracking"

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface for analysis table and graph ID tracking")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IAnaTabIds
'@description
'Defines the public contract for the AnaTabIds class. Exposes methods to
'register graph series information, graph formatting metadata, table named
'ranges, and to write all tracked graphs to the output worksheet.


' Scope for analysis ID operations
Public Enum AnalysisIdsScope
    AnalysisIdsScopeTimeSeries = 1
    AnalysisIdsScopeSpatial = 2
    AnalysisIdsScopeNormal = 3
    AnalysisIdsScopeSpatioTemporal = 4
End Enum


'@section GraphTracking
'===============================================================================
'@description Methods for registering graph series and formatting metadata.

'@jump:add-graph-info
'@sub-title Register a graph series in the tracking table
'@details
'Adds one series entry to the internal graph-tracking table. Each call
'records the series name, chart type, position, label text, column label,
'and output range address. Optional parameters control label hard-coding
'and prefix filtering for spatial and spatio-temporal scopes.
'@param scope Byte. AnalysisIdsScope value selecting the target tracking table.
'@param tabId String. Identifier of the parent analysis table.
'@param graphId String. Identifier of the graph receiving this series.
'@param seriesName String. Name of the data series.
'@param seriesType String. Chart series type (e.g. "bar", "line").
'@param seriesPos String. Position index of the series within the graph.
'@param seriesLabel String. Display label for the series.
'@param seriesColumnLabel String. Column header label for the series data.
'@param outRangeAddress String. Worksheet address of the series data range.
'@param hardCodeLabels Optional Boolean. When True, labels are written as
'literal values rather than cell references. Defaults to False.
'@param prefix Optional String. Admin-level prefix for spatial graphs.
'Defaults to vbNullString.
'@param prefixOnly Optional Boolean. When True, only the prefix series is
'written. Defaults to False.
'@export
Public Sub AddGraphInfo( _
    ByVal scope As Byte, _
    ByVal tabId As String, _
    ByVal graphId As String, _
    ByVal seriesName As String, _
    ByVal seriesType As String, _
    ByVal seriesPos As String, _
    ByVal seriesLabel As String, _
    ByVal seriesColumnLabel As String, _
    ByVal outRangeAddress As String, _
    Optional ByVal hardCodeLabels As Boolean = False, _
    Optional ByVal prefix As String = vbNullString, _
    Optional ByVal prefixOnly As Boolean = False)
End Sub

'@jump:add-graph-format
'@sub-title Register graph formatting metadata in the tracking table
'@details
'Adds formatting information for a graph: category axis title, value axis
'title, optional plot title, and a height-factor multiplier. This data is
'consumed by WriteGraphs when rendering charts on the output worksheet.
'@param scope Byte. AnalysisIdsScope value selecting the target tracking table.
'@param tabId String. Identifier of the parent analysis table.
'@param graphId String. Identifier of the graph to format.
'@param catTitle String. Category (X) axis title.
'@param valuesTitle String. Value (Y) axis title.
'@param hardCodeLabels Optional Boolean. When True, axis labels are written
'as literal values. Defaults to False.
'@param heightFactor Optional Long. Multiplier for graph height. Defaults to 1.
'@param plotTitle Optional String. Title displayed above the chart area.
'Defaults to vbNullString.
'@export
Public Sub AddGraphFormat( _
    ByVal scope As Byte, _
    ByVal tabId As String, _
    ByVal graphId As String, _
    ByVal catTitle As String, _
    ByVal valuesTitle As String, _
    Optional ByVal hardCodeLabels As Boolean = False, _
    Optional ByVal heightFactor As Long = 1, _
    Optional ByVal plotTitle As String = vbNullString)
End Sub


'@section TableTracking
'===============================================================================
'@description Methods for registering table named ranges for export.

'@jump:add-table-infos
'@sub-title Register table named ranges in the tracking table
'@details
'Adds the list of named range names created by a cross-table to the
'internal tracking table. These entries are used later during export to
'identify which ranges belong to each analysis table.
'@param scope Byte. AnalysisIdsScope value selecting the target tracking table.
'@param tabId String. Identifier of the analysis table.
'@param tabRangesNames BetterArray. Collection of named range names to register.
'@param exportTable Optional String. Export flag ("yes" or "no"). Defaults to "yes".
'@export
Public Sub AddTableInfos( _
    ByVal scope As Byte, _
    ByVal tabId As String, _
    ByVal tabRangesNames As BetterArray, _
    Optional ByVal exportTable As String = "yes")
End Sub


'@section GraphWriting
'===============================================================================
'@description Methods for rendering tracked graphs on output worksheets.

'@jump:write-graphs
'@sub-title Write all tracked graphs to the output worksheet
'@details
'Iterates through the graph-tracking table for the specified scope and
'creates Excel chart objects on the output worksheet. Each graph is built
'from its registered series data and formatted according to the stored
'formatting metadata.
'@param outsh Worksheet. The output worksheet receiving the chart objects.
'@param scope Byte. AnalysisIdsScope value selecting which tracking table to read.
'@export
Public Sub WriteGraphs(ByVal outsh As Worksheet, ByVal scope As Byte): End Sub
