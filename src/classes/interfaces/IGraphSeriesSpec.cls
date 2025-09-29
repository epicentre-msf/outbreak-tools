VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGraphSeriesSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis.Graphs")
'@ModuleDescription("Describes a single chart series extracted from graph specifications")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Factory
'===============================================================================
'Create a new graph series descriptor.
Public Function Create(ByVal seriesName As String, _
                       ByVal seriesType As String, _
                       ByVal seriesPosition As String) As IGraphSeriesSpec
End Function

'Return the instance as the interface type for fluent usage.
Public Property Get Self() As IGraphSeriesSpec
End Property

'@section Properties
'===============================================================================
'Retrieves the friendly series identifier used when defining the chart.
Public Property Get SeriesName() As String
End Property

'Retrieves the requested chart type for the series (e.g. bar, line).
Public Property Get SeriesType() As String
End Property

'Retrieves the target axis/position information for the series.
Public Property Get SeriesPosition() As String
End Property

'Retrieves the label describing the row dimension for the series.
Public Property Get RowLabel() As String
End Property

'Retrieves the label applied to column categories.
Public Property Get ColumnLabel() As String
End Property

'Optional prefix stored for label translation.
Public Property Get LabelPrefix() As String
End Property

'@section Mutators
'===============================================================================
'Assigns or updates the descriptive labels captured for the series.
Public Sub AssignLabels(ByVal rowLabel As String, _
                        ByVal columnLabel As String, _
                        Optional ByVal labelPrefix As String = vbNullString)
End Sub

'@section Utilities
'===============================================================================
'Returns a deep copy of the current series specification.
Public Function Clone() As IGraphSeriesSpec
End Function

