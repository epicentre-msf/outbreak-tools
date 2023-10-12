VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "IAnaTabIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Dictionary")
'@ModuleDescription("Interface of Analysis Table Ids")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods
Public Enum AnalysisIdsScope
  AnalysisIdsScopeTimeSeries = 1 'A time series analysis
  AnalysisIdsScopeSpatial = 2 'A spatial analysis
  AnalysisIdsScopeNormal = 3 'A standard univariate / bivariate analysis
  AnalysisIdsScopeSpatioTemporal = 4 'spatio-Temporal analysis
End Enum

Public Sub AddGraphInfo( _
  ByVal scope As Byte, _ 
  ByVal tabId As String, _
  ByVal graphId As String, _
  ByVal SeriesName As String, _ 
  ByVal SeriesType As String, _ 
  ByVal SeriesPos As String, _ 
  ByVal SeriesLabel As String, _ 
  ByVal SeriesColumnLabel As String, _ 
  ByVal outRangeAddress As String, _ 
  Optional ByVal hardCodeLabels As Boolean = False, _
  Optional ByVal prefix As String = vbNullString, _ 
  Optional ByVal prefixOnly As Boolean = False)
End Sub

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

Public Sub AddTableInfos( _ 
  ByVal scope As Byte, _
  ByVal tabId As String, _  
  ByVal tabRangesNames As BetterArray, _ 
  Optional ByVal exportTable As String = "yes")
End Sub

'Write Graphs on one scope
Public Sub WriteGraphs(ByVal outsh As Worksheet, ByVal scope As Byte): End Sub