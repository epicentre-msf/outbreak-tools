VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableGraphBuilderStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of ICrossTableGraphBuilder capturing invocation details")
'@IgnoreModule UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ICrossTableGraphBuilder

Private Type TStubState
    chart As IGraphChartSpec
    lastTable As ICrossTable
    lastGraphId As String
    lastGraphTitle As String
    seriesBuilder As IGraphSeriesBuilder
End Type

Private this As TStubState

Public Sub ConfigureReturnChart(ByVal chartSpec As IGraphChartSpec)
    Set this.chart = chartSpec
End Sub

Public Property Get LastTable() As ICrossTable
    Set LastTable = this.lastTable
End Property

Public Property Get LastGraphId() As String
    LastGraphId = this.lastGraphId
End Property

Public Property Get LastGraphTitle() As String
    LastGraphTitle = this.lastGraphTitle
End Property

Private Function ICrossTableGraphBuilder_Create(Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing) As ICrossTableGraphBuilder
    If Not seriesBuilder Is Nothing Then Set this.seriesBuilder = seriesBuilder
    Set ICrossTableGraphBuilder_Create = Me
End Function

Private Property Get ICrossTableGraphBuilder_Self() As ICrossTableGraphBuilder
    Set ICrossTableGraphBuilder_Self = Me
End Property

Private Function ICrossTableGraphBuilder_BuildGraph(ByVal table As ICrossTable, _
                                                    ByVal graphId As String, _
                                                    Optional ByVal graphTitle As String = vbNullString) As IGraphChartSpec
    Set this.lastTable = table
    this.lastGraphId = graphId
    this.lastGraphTitle = graphTitle

    If this.chart Is Nothing Then
        Set ICrossTableGraphBuilder_BuildGraph = Nothing
    Else
        Set ICrossTableGraphBuilder_BuildGraph = this.chart.Clone
    End If
End Function

Private Property Get ICrossTableGraphBuilder_SeriesBuilder() As IGraphSeriesBuilder
    Set ICrossTableGraphBuilder_SeriesBuilder = this.seriesBuilder
End Property

Private Sub ICrossTableGraphBuilder_Clear()
    'No-op for stub
End Sub

