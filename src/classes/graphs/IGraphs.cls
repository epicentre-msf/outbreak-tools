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


'@Folder("Analysis")
'@ModuleDescription("Interface of Graph object")
'@IgnoreModule SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  UseMeaningfulName, HungarianNotation

'Differents scopes for graphs. Depending on the scope, the size and other
'adjustments are made. For example, on Time Series, the graph will appear
'bigger, compare to basic normal graph.

Public Enum GraphScope
  GraphScopeTimeSeries = 1 'A time series graph (long graph)
  GraphScopeSpatial = 2 'A spatial analysis graph
  GraphScopeNormal = 3 'A standard univariate / bivariate graph
End Enum

Option Explicit

'@jump:operate
'@section Operate with the outside world
'===============================================================================
'Exposed methods
'@jump:add
Public Sub Add(): End Sub

'@jump:addseries
Public Sub AddSeries(ByVal rngName As String, ByVal chrtType As String, _
                     Optional axisPos As String = vbNullString): End Sub

'@jump:addlabels
Public Sub AddLabels(ByVal catName As String, ByVal lblName As String, _
                     Optional ByVal prefix As String = vbNullString, _
                     Optional ByVal prefixOnly As Boolean = False, _
                     Optional ByVal hardCodeLabels As Boolean = True): End Sub
'Maybe format
'@jump:format
Public Sub Format(Optional ByVal valuesTitle As String = vbNullString, _
                  Optional ByVal catTitle As String = vbNullString, _
                  Optional ByVal plotTitle As String = vbNullString, _
                  Optional ByVal scope As Byte = GraphScopeNormal, _
                  Optional ByVal heightFactor As Long = 1, _
                  Optional ByVal hardCodeLabels As Boolean = True)
End Sub
