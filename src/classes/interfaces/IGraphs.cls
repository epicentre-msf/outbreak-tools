VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGraphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

'Differents scopes for graphs
Public Enum GraphScope
  GraphScopeTimeSeries = 1 'A time series graph (long graph)
  GraphScopeSpatial = 2 'A spatial analysis graph
  GraphScopeNormal = 3 'A standard univariate / bivariate graph
End Enum

Option Explicit

'Exposed methods
Public Sub Add(): End Sub

Public Sub AddSeries(ByVal rngName As String, ByVal chrtType As String, _
                     Optional axisPos As String = vbNullString): End Sub

Public Sub AddLabels(ByVal catName As String, ByVal lblName As String, _
                     Optional ByVal prefix As String = vbNullString, _
                     Optional ByVal prefixOnly As Boolean = False, _
                     Optional ByVal hardCodeLabels As Boolean = True): End Sub
'Maybe format
Public Sub Format(Optional ByVal valuesTitle As String = vbNullString, _
                  Optional ByVal catTitle As String = vbNullString, _
                  Optional ByVal plotTitle As String = vbNullString, _
                  Optional ByVal scope As GraphScope = GraphScopeNormal, _
                  Optional ByVal heightFactor As Long = 1, _
                  Optional ByVal hardCodeLabels As Boolean = True)
End Sub


