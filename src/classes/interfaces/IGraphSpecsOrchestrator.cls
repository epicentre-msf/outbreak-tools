VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGraphSpecsOrchestrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis.Graphs")
'@ModuleDescription("Coordinates graph specification builders for simple and complex scenarios")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Create an orchestrator configured for a simple cross-table graph.
Public Function CreateForCrossTable(ByVal table As ICrossTable, _
                                    Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing) As IGraphSpecsOrchestrator
End Function

'Create an orchestrator configured for time-series (complex) graph specifications.
Public Function CreateForTimeSeries(ByVal graphList As ListObject, _
                                    ByVal timeSeriesList As ListObject, _
                                    ByVal titleList As ListObject, _
                                    ByVal linelist As ILinelistSpecs, _
                                    Optional ByVal cache As IGraphSpecsCache = Nothing, _
                                    Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing, _
                                    Optional ByVal timeBuilder As ITimeSeriesGraphBuilder = Nothing, _
                                    Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing, _
                                    Optional ByVal specsFactory As ITimeSeriesTableSpecsFactory = Nothing) As IGraphSpecsOrchestrator
End Function

'Expose the current instance for fluent usage.
Public Property Get Self() As IGraphSpecsOrchestrator
End Property

'Build and return the chart specifications captured by the orchestrator.
Public Function Build() As BetterArray
End Function

'Retrieve the cached chart specifications without rebuilding.
Public Function Charts() As BetterArray
End Function

'Return the number of charts produced by the orchestrator.
Public Property Get GraphCount() As Long
End Property

'Return the total series count across all charts handled by the orchestrator.
Public Property Get SeriesCount() As Long
End Property

'Identify whether the orchestrator manages complex (time-series) graphs.
Public Property Get IsComplex() As Boolean
End Property

'Invalidate cached charts so future builds re-evaluate the source data.
Public Sub Invalidate()
End Sub

