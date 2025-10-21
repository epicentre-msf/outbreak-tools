VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphChartSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis.Graphs")
'@ModuleDescription("Concrete chart specification storing graph series metadata")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IGraphChartSpec

Private Type TChartState
    graphId As String
    title As String
    seriesList As BetterArray
End Type

Private this As TChartState
Private Const CLASSNAME As String = "GraphChartSpec"

'@section Factory
'===============================================================================
Public Function Create(ByVal graphId As String, _
                       ByVal graphTitle As String) As IGraphChartSpec
    ValidateRequiredText graphId, "graphId"

    With New GraphChartSpec
        .Initialise graphId, graphTitle
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IGraphChartSpec
    Set Self = Me
End Property

Public Sub Initialise(ByVal graphId As String, ByVal graphTitle As String)
    this.graphId = Trim$(graphId)
    this.title = Trim$(graphTitle)
    Set this.seriesList = New BetterArray
    this.seriesList.LowerBound = 1
End Sub

'@section Properties
'===============================================================================
Public Property Get GraphId() As String
    GraphId = this.graphId
End Property

Public Property Get Title() As String
    Title = this.title
End Property

Public Function Series() As BetterArray
    Dim cloneList As BetterArray
    Dim idx As Long
    Dim spec As IGraphSeriesSpec

    Set cloneList = New BetterArray
    cloneList.LowerBound = 1

    If (Not this.seriesList Is Nothing) And (this.seriesList.Length > 0) Then
        For idx = this.seriesList.LowerBound To this.seriesList.UpperBound
            Set spec = this.seriesList.Item(idx)
            If Not spec Is Nothing Then cloneList.Push spec.Clone
        Next idx
    End If

    Set Series = cloneList
End Function

Public Property Get SeriesCount() As Long
    If this.seriesList Is Nothing Then
        SeriesCount = 0
    Else
        SeriesCount = this.seriesList.Length
    End If
End Property

'@section Mutators
'===============================================================================
Public Sub AddSeries(ByVal seriesSpec As IGraphSeriesSpec)
    ValidateSeries seriesSpec

    If this.seriesList Is Nothing Then
        Set this.seriesList = New BetterArray
        this.seriesList.LowerBound = 1
    End If

    this.seriesList.Push seriesSpec.Clone
End Sub

'@section Utilities
'===============================================================================
Public Function Clone() As IGraphChartSpec
    Dim cloneChart As GraphChartSpec
    Dim idx As Long
    Dim spec As IGraphSeriesSpec

    Set cloneChart = New GraphChartSpec
    cloneChart.Initialise this.graphId, this.title

    If Not this.seriesList Is Nothing Then
        For idx = this.seriesList.LowerBound To this.seriesList.UpperBound
            Set spec = this.seriesList.Item(idx)
            If Not spec Is Nothing Then cloneChart.AddSeries spec
        Next idx
    End If

    Set Clone = cloneChart.Self
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidateRequiredText(ByVal candidate As String, ByVal argumentName As String)
    If NormalizeGraphValue(candidate) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   argumentName & " cannot be empty when defining a chart"
    End If
End Sub

Private Sub ValidateSeries(ByVal seriesSpec As IGraphSeriesSpec)
    If seriesSpec Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "seriesSpec cannot be Nothing"
    End If

    If NormalizeGraphValue(TypeName(seriesSpec)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "seriesSpec must expose a concrete implementation"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IGraphChartSpec_Create(ByVal graphId As String, _
                                        ByVal graphTitle As String) As IGraphChartSpec
    Set IGraphChartSpec_Create = Create(graphId, graphTitle)
End Function

Private Property Get IGraphChartSpec_Self() As IGraphChartSpec
    Set IGraphChartSpec_Self = Self
End Property

Private Property Get IGraphChartSpec_GraphId() As String
    IGraphChartSpec_GraphId = GraphId
End Property

Private Property Get IGraphChartSpec_Title() As String
    IGraphChartSpec_Title = Title
End Property

Private Function IGraphChartSpec_Series() As BetterArray
    Set IGraphChartSpec_Series = Series()
End Function

Private Property Get IGraphChartSpec_SeriesCount() As Long
    IGraphChartSpec_SeriesCount = SeriesCount
End Property

Private Sub IGraphChartSpec_AddSeries(ByVal seriesSpec As IGraphSeriesSpec)
    AddSeries seriesSpec
End Sub

Private Function IGraphChartSpec_Clone() As IGraphChartSpec
    Set IGraphChartSpec_Clone = Clone()
End Function

