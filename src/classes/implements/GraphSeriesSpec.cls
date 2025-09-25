VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphSeriesSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis.Graphs")
'@ModuleDescription("Concrete implementation of IGraphSeriesSpec capturing chart series metadata")
'@depends ProjectError
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IGraphSeriesSpec

Private Type TSeriesState
    seriesName As String
    seriesType As String
    seriesPosition As String
    rowLabel As String
    columnLabel As String
    labelPrefix As String
End Type

Private this As TSeriesState
Private Const CLASSNAME As String = "GraphSeriesSpec"

'@section Factory
'===============================================================================
Public Function Create(ByVal seriesName As String, _
                       ByVal seriesType As String, _
                       ByVal seriesPosition As String) As IGraphSeriesSpec

    ValidateRequiredText seriesName, "seriesName"
    ValidateRequiredText seriesType, "seriesType"
    ValidateRequiredText seriesPosition, "seriesPosition"

    With New GraphSeriesSpec
        .InitialiseSeries Trim$(seriesName), Trim$(seriesType), Trim$(seriesPosition)
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IGraphSeriesSpec
    Set Self = Me
End Property

Friend Sub InitialiseSeries(ByVal seriesName As String, _
                            ByVal seriesType As String, _
                            ByVal seriesPosition As String)
    this.seriesName = seriesName
    this.seriesType = seriesType
    this.seriesPosition = seriesPosition
    this.rowLabel = vbNullString
    this.columnLabel = vbNullString
    this.labelPrefix = vbNullString
End Sub

'@section Properties
'===============================================================================
Public Property Get SeriesName() As String
    SeriesName = this.seriesName
End Property

Public Property Get SeriesType() As String
    SeriesType = this.seriesType
End Property

Public Property Get SeriesPosition() As String
    SeriesPosition = this.seriesPosition
End Property

Public Property Get RowLabel() As String
    RowLabel = this.rowLabel
End Property

Public Property Get ColumnLabel() As String
    ColumnLabel = this.columnLabel
End Property

Public Property Get LabelPrefix() As String
    LabelPrefix = this.labelPrefix
End Property

'@section Mutators
'===============================================================================
Public Sub AssignLabels(ByVal rowLabel As String, _
                        ByVal columnLabel As String, _
                        Optional ByVal labelPrefix As String = vbNullString)

    this.rowLabel = Trim$(rowLabel)
    this.columnLabel = Trim$(columnLabel)
    this.labelPrefix = Trim$(labelPrefix)
End Sub

'@section Utilities
'===============================================================================
Public Function Clone() As IGraphSeriesSpec
    Dim cloneSeries As GraphSeriesSpec

    Set cloneSeries = New GraphSeriesSpec
    cloneSeries.InitialiseSeries this.seriesName, this.seriesType, this.seriesPosition
    cloneSeries.AssignLabels this.rowLabel, this.columnLabel, this.labelPrefix
    Set Clone = cloneSeries.Self
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidateRequiredText(ByVal candidate As String, ByVal argumentName As String)
    If NormalizeGraphValue(candidate) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   argumentName & " cannot be empty when defining a graph series"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IGraphSeriesSpec_Create(ByVal seriesName As String, _
                                         ByVal seriesType As String, _
                                         ByVal seriesPosition As String) As IGraphSeriesSpec
    Set IGraphSeriesSpec_Create = Create(seriesName, seriesType, seriesPosition)
End Function

Private Property Get IGraphSeriesSpec_Self() As IGraphSeriesSpec
    Set IGraphSeriesSpec_Self = Self
End Property

Private Property Get IGraphSeriesSpec_SeriesName() As String
    IGraphSeriesSpec_SeriesName = SeriesName
End Property

Private Property Get IGraphSeriesSpec_SeriesType() As String
    IGraphSeriesSpec_SeriesType = SeriesType
End Property

Private Property Get IGraphSeriesSpec_SeriesPosition() As String
    IGraphSeriesSpec_SeriesPosition = SeriesPosition
End Property

Private Property Get IGraphSeriesSpec_RowLabel() As String
    IGraphSeriesSpec_RowLabel = RowLabel
End Property

Private Property Get IGraphSeriesSpec_ColumnLabel() As String
    IGraphSeriesSpec_ColumnLabel = ColumnLabel
End Property

Private Property Get IGraphSeriesSpec_LabelPrefix() As String
    IGraphSeriesSpec_LabelPrefix = LabelPrefix
End Property

Private Sub IGraphSeriesSpec_AssignLabels(ByVal rowLabel As String, _
                                          ByVal columnLabel As String, _
                                          Optional ByVal labelPrefix As String = vbNullString)
    AssignLabels rowLabel, columnLabel, labelPrefix
End Sub

Private Function IGraphSeriesSpec_Clone() As IGraphSeriesSpec
    Set IGraphSeriesSpec_Clone = Clone()
End Function

