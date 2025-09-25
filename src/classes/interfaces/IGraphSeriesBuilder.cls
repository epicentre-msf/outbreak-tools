VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGraphSeriesBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis.Graphs")
'@ModuleDescription("Accumulator used to build graph chart specifications")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Factory
'===============================================================================
'Create a new builder ready to record series details.
Public Function Create() As IGraphSeriesBuilder
End Function

'Return the current instance as the interface type.
Public Property Get Self() As IGraphSeriesBuilder
End Property

'@section Series recording
'===============================================================================
'Append a new series definition with core metadata.
Public Function AddSeries(ByVal seriesName As String, _
                          ByVal seriesType As String, _
                          ByVal seriesPosition As String) As IGraphSeriesSpec
End Function

'Attach labels to the most recently added series.
Public Sub AssignLabels(ByVal rowLabel As String, _
                        ByVal columnLabel As String, _
                        Optional ByVal labelPrefix As String = vbNullString)
End Sub

'@section Graph creation
'===============================================================================
'Build a chart specification with the accumulated series.
Public Function BuildGraph(ByVal graphId As String, _
                           ByVal graphTitle As String) As IGraphChartSpec
End Function

'@section Cache management
'===============================================================================
'Reset the builder to its initial empty state.
Public Sub Clear()
End Sub

'@section Diagnostics
'===============================================================================
'Return the number of series currently buffered.
Public Property Get SeriesCount() As Long
End Property

