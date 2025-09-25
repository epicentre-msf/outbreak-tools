VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeSeriesGraphBuilderStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of ITimeSeriesGraphBuilder returning pre-configured chart lists")
'@IgnoreModule UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ITimeSeriesGraphBuilder

Private Type TStubState
    charts As BetterArray
    buildGraphCalls As Collection
    invalidateCount As Long
End Type

Private this As TStubState

Public Sub ConfigureCharts(ByVal chartList As BetterArray)
    Set this.charts = chartList
End Sub

Public Property Get BuildGraphCallCount() As Long
    If this.buildGraphCalls Is Nothing Then
        BuildGraphCallCount = 0
    Else
        BuildGraphCallCount = this.buildGraphCalls.Count
    End If
End Property

Public Property Get InvalidateCount() As Long
    InvalidateCount = this.invalidateCount
End Property

Private Function ITimeSeriesGraphBuilder_Create(ByVal graphCache As IGraphSpecsCache, _
                                                ByVal timeSeriesList As ListObject, _
                                                ByVal titleList As ListObject, _
                                                ByVal linelist As ILinelistSpecs, _
                                                Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing, _
                                                Optional ByVal specsFactory As ITimeSeriesTableSpecsFactory = Nothing) As ITimeSeriesGraphBuilder
    Set ITimeSeriesGraphBuilder_Create = Me
End Function

Private Property Get ITimeSeriesGraphBuilder_Self() As ITimeSeriesGraphBuilder
    Set ITimeSeriesGraphBuilder_Self = Me
End Property

Private Function ITimeSeriesGraphBuilder_BuildGraph(ByVal graphId As String) As IGraphChartSpec
    If this.buildGraphCalls Is Nothing Then Set this.buildGraphCalls = New Collection
    this.buildGraphCalls.Add graphId

    If this.charts Is Nothing Or this.charts.Length = 0 Then
        Set ITimeSeriesGraphBuilder_BuildGraph = Nothing
    Else
        Set ITimeSeriesGraphBuilder_BuildGraph = this.charts.Item(this.charts.LowerBound).Clone
    End If
End Function

Private Function ITimeSeriesGraphBuilder_BuildAll() As BetterArray
    Dim cloneList As BetterArray
    Dim idx As Long
    Dim chart As IGraphChartSpec

    Set cloneList = New BetterArray
    cloneList.LowerBound = 1

    If Not this.charts Is Nothing Then
        For idx = this.charts.LowerBound To this.charts.UpperBound
            Set chart = this.charts.Item(idx)
            If Not chart Is Nothing Then cloneList.Push chart.Clone
        Next idx
    End If

    Set ITimeSeriesGraphBuilder_BuildAll = cloneList
End Function

Private Sub ITimeSeriesGraphBuilder_Invalidate()
    this.invalidateCount = this.invalidateCount + 1
End Sub

