VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICrossTableGraphBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis.Graphs")
'@ModuleDescription("Builds chart specifications for cross-table based graphs")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Factory
'===============================================================================
'Create a builder that will use the provided series builder instance when assembling graphs.
Public Function Create(Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing) As ICrossTableGraphBuilder
End Function

'Return the instance as an interface for fluent consumption.
Public Property Get Self() As ICrossTableGraphBuilder
End Property

'@section Build operations
'===============================================================================
'Build a chart specification for a non time-series cross table graph.
Public Function BuildGraph(ByVal table As ICrossTable, _
                           ByVal graphId As String, _
                           Optional ByVal graphTitle As String = vbNullString) As IGraphChartSpec
End Function

'Expose the active series builder for scenarios where callers need further control.
Public Property Get SeriesBuilder() As IGraphSeriesBuilder
End Property

