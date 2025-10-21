VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGraphChartSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis.Graphs")
'@ModuleDescription("Represents a chart definition composed of multiple series")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Factory
'===============================================================================
'Create a chart definition with the supplied identifier and title.
Public Function Create(ByVal graphId As String, _
                       ByVal graphTitle As String) As IGraphChartSpec
End Function

'Return the instance as the interface type.
Public Property Get Self() As IGraphChartSpec
End Property

'@section Properties
'===============================================================================
'Retrieve the unique identifier associated with the chart.
Public Property Get GraphId() As String
End Property

'Retrieve the chart title displayed to end users.
Public Property Get Title() As String
End Property

'Retrieve a snapshot of the configured series collection.
Public Function Series() As BetterArray
End Function

'Retrieve the number of series configured for this chart.
Public Property Get SeriesCount() As Long
End Property

'@section Mutators
'===============================================================================
'Add a new series definition to the chart.
Public Sub AddSeries(ByVal seriesSpec As IGraphSeriesSpec)
End Sub

'@section Utilities
'===============================================================================
'Return a deep copy of the chart specification.
Public Function Clone() As IGraphChartSpec
End Function

