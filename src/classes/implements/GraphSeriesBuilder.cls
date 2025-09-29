VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphSeriesBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis.Graphs")
'@ModuleDescription("Builder that accumulates graph series data before producing chart specs")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IGraphSeriesBuilder

Private Type TBuilderState
    seriesList As BetterArray
End Type

Private this As TBuilderState
Private Const CLASSNAME As String = "GraphSeriesBuilder"

'@section Factory
'===============================================================================
Public Function Create() As IGraphSeriesBuilder
    With New GraphSeriesBuilder
        .ResetState
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IGraphSeriesBuilder
    Set Self = Me
End Property

'@section Series recording
'===============================================================================
Public Function AddSeries(ByVal seriesName As String, _
                          ByVal seriesType As String, _
                          ByVal seriesPosition As String) As IGraphSeriesSpec
    Dim spec As IGraphSeriesSpec

    EnsureStateReady

    Set spec = GraphSeriesSpec.Create(seriesName, seriesType, seriesPosition)
    this.seriesList.Push spec
    Set AddSeries = spec
End Function

Public Sub AssignLabels(ByVal rowLabel As String, _
                        ByVal columnLabel As String, _
                        Optional ByVal labelPrefix As String = vbNullString)
    Dim lastIndex As Long
    Dim spec As IGraphSeriesSpec

    EnsureStateReady

    If this.seriesList.Length = 0 Then
        ThrowError CLng(ProjectError.ErrorUnexpectedState), _
                   "Cannot assign labels without adding a series first"
    End If

    lastIndex = this.seriesList.UpperBound
    Set spec = this.seriesList.Item(lastIndex)
    spec.AssignLabels rowLabel, columnLabel, labelPrefix
End Sub

'@section Graph creation
'===============================================================================
Public Function BuildGraph(ByVal graphId As String, _
                           ByVal graphTitle As String) As IGraphChartSpec
    Dim chartSpec As IGraphChartSpec
    Dim idx As Long
    Dim spec As IGraphSeriesSpec

    Set chartSpec = GraphChartSpec.Create(graphId, graphTitle)

    If (Not this.seriesList Is Nothing) And (this.seriesList.Length > 0) Then
        For idx = this.seriesList.LowerBound To this.seriesList.UpperBound
            Set spec = this.seriesList.Item(idx)
            If Not spec Is Nothing Then chartSpec.AddSeries spec
        Next idx
    End If

    Set BuildGraph = chartSpec
    Clear
End Function

'@section Cache management
'===============================================================================
Public Sub Clear()
    ResetState
End Sub

'@section Diagnostics
'===============================================================================
Public Property Get SeriesCount() As Long
    If this.seriesList Is Nothing Then
        SeriesCount = 0
    Else
        SeriesCount = this.seriesList.Length
    End If
End Property

'@section Internal helpers
'===============================================================================
Private Sub ResetState()
    Set this.seriesList = New BetterArray
    this.seriesList.LowerBound = 1
End Sub

Private Sub EnsureStateReady()
    If this.seriesList Is Nothing Then ResetState
End Sub

Private Function NormalizeGraphValue(ByVal valueText As String) As String
    NormalizeGraphValue = LCase$(Trim$(valueText))
End Function

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IGraphSeriesBuilder_Create() As IGraphSeriesBuilder
    Set IGraphSeriesBuilder_Create = Create()
End Function

Private Property Get IGraphSeriesBuilder_Self() As IGraphSeriesBuilder
    Set IGraphSeriesBuilder_Self = Self
End Property

Private Function IGraphSeriesBuilder_AddSeries(ByVal seriesName As String, _
                                               ByVal seriesType As String, _
                                               ByVal seriesPosition As String) As IGraphSeriesSpec
    Set IGraphSeriesBuilder_AddSeries = AddSeries(seriesName, seriesType, seriesPosition)
End Function

Private Sub IGraphSeriesBuilder_AssignLabels(ByVal rowLabel As String, _
                                             ByVal columnLabel As String, _
                                             Optional ByVal labelPrefix As String = vbNullString)
    AssignLabels rowLabel, columnLabel, labelPrefix
End Sub

Private Function IGraphSeriesBuilder_BuildGraph(ByVal graphId As String, _
                                                ByVal graphTitle As String) As IGraphChartSpec
    Set IGraphSeriesBuilder_BuildGraph = BuildGraph(graphId, graphTitle)
End Function

Private Sub IGraphSeriesBuilder_Clear()
    Clear
End Sub

Private Property Get IGraphSeriesBuilder_SeriesCount() As Long
    IGraphSeriesBuilder_SeriesCount = SeriesCount
End Property

