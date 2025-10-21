VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphSpecsOrchestratorStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of IGraphSpecsOrchestrator returning configured chart specs")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IGraphSpecsOrchestrator

Private Type TStubState
    charts As BetterArray
    isComplex As Boolean
End Type

Private this As TStubState

Public Sub Configure(ByVal charts As BetterArray, Optional ByVal isComplex As Boolean = False)
    Set this.charts = CloneCharts(charts)
    this.isComplex = isComplex
End Sub

Private Function CloneCharts(ByVal charts As BetterArray) As BetterArray
    Dim cloneList As BetterArray
    Dim idx As Long
    Dim chart As IGraphChartSpec

    Set cloneList = New BetterArray
    cloneList.LowerBound = 1

    If charts Is Nothing Then
        Set CloneCharts = cloneList
        Exit Function
    End If

    If charts.Length = 0 Then
        Set CloneCharts = cloneList
        Exit Function
    End If

    For idx = charts.LowerBound To charts.UpperBound
        Set chart = charts.Item(idx)
        If Not chart Is Nothing Then cloneList.Push chart.Clone
    Next idx

    Set CloneCharts = cloneList
End Function

'@section Interface implementation
'===============================================================================
Private Function IGraphSpecsOrchestrator_CreateForCrossTable(ByVal table As ICrossTable, _
                                                             Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing) As IGraphSpecsOrchestrator
    Set IGraphSpecsOrchestrator_CreateForCrossTable = Me
End Function

Private Function IGraphSpecsOrchestrator_CreateForTimeSeries(ByVal graphList As ListObject, _
                                                             ByVal timeSeriesList As ListObject, _
                                                             ByVal titleList As ListObject, _
                                                             ByVal linelist As ILinelistSpecs, _
                                                             Optional ByVal cache As IGraphSpecsCache = Nothing, _
                                                             Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing, _
                                                             Optional ByVal timeBuilder As ITimeSeriesGraphBuilder = Nothing, _
                                                             Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing, _
                                                             Optional ByVal specsFactory As ITimeSeriesTableSpecsFactory = Nothing) As IGraphSpecsOrchestrator
    Set IGraphSpecsOrchestrator_CreateForTimeSeries = Me
End Function

Private Property Get IGraphSpecsOrchestrator_Self() As IGraphSpecsOrchestrator
    Set IGraphSpecsOrchestrator_Self = Me
End Property

Private Function IGraphSpecsOrchestrator_Build() As BetterArray
    Set IGraphSpecsOrchestrator_Build = CloneCharts(this.charts)
End Function

Private Function IGraphSpecsOrchestrator_Charts() As BetterArray
    Set IGraphSpecsOrchestrator_Charts = CloneCharts(this.charts)
End Function

Private Property Get IGraphSpecsOrchestrator_GraphCount() As Long
    If this.charts Is Nothing Then
        IGraphSpecsOrchestrator_GraphCount = 0
    Else
        IGraphSpecsOrchestrator_GraphCount = this.charts.Length
    End If
End Property

Private Property Get IGraphSpecsOrchestrator_SeriesCount() As Long
    Dim idx As Long
    Dim chart As IGraphChartSpec
    Dim count As Long

    If Not this.charts Is Nothing Then
        For idx = this.charts.LowerBound To this.charts.UpperBound
            Set chart = this.charts.Item(idx)
            If Not chart Is Nothing Then count = count + chart.SeriesCount
        Next idx
    End If

    IGraphSpecsOrchestrator_SeriesCount = count
End Property

Private Property Get IGraphSpecsOrchestrator_IsComplex() As Boolean
    IGraphSpecsOrchestrator_IsComplex = this.isComplex
End Property

Private Sub IGraphSpecsOrchestrator_Invalidate()
    Set this.charts = Nothing
End Sub
