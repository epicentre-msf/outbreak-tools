VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITimeSeriesGraphBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis.Graphs")
'@ModuleDescription("Builds graph specifications for time-series analysis configurations")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Create a builder bound to the supplied cache and list objects.
Public Function Create(ByVal graphCache As IGraphSpecsCache, _
                       ByVal timeSeriesList As ListObject, _
                       ByVal titleList As ListObject, _
                       ByVal linelist As ILinelistSpecs, _
                       Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing, _
                       Optional ByVal specsFactory As ITimeSeriesTableSpecsFactory = Nothing) As ITimeSeriesGraphBuilder
End Function

'Return the instance as an interface for fluent usage.
Public Property Get Self() As ITimeSeriesGraphBuilder
End Property

'Build a complete chart specification for the requested graph identifier.
Public Function BuildGraph(ByVal graphId As String) As IGraphChartSpec
End Function

'Build chart specifications for every graph indexed in the cache.
Public Function BuildAll() As BetterArray
End Function

'Reset any cached state (e.g. series index) so builders can react to list changes.
Public Sub Invalidate()
End Sub

