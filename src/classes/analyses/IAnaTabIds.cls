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

' Scope for analysis ID operations
Public Enum AnalysisIdsScope
    AnalysisIdsScopeTimeSeries = 1
    AnalysisIdsScopeSpatial = 2
    AnalysisIdsScopeNormal = 3
    AnalysisIdsScopeSpatioTemporal = 4
End Enum

' @description Add graph series information to the tracking table
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

' @description Add graph formatting information to the tracking table
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

' @description Add table named ranges to the tracking table for export
Public Sub AddTableInfos( _
    ByVal scope As Byte, _
    ByVal tabId As String, _
    ByVal tabRangesNames As BetterArray, _
    Optional ByVal exportTable As String = "yes")
End Sub

' @description Write all tracked graphs to the output worksheet
Public Sub WriteGraphs(ByVal outsh As Worksheet, ByVal scope As Byte): End Sub
