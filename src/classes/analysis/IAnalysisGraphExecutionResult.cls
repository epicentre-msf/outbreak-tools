VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisGraphExecutionResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Aggregates charts produced by the AnalysisGraphEngine")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return cross-table graph execution items keyed by their plan entry.
Public Property Get CrossTableItems() As BetterArray
End Property

'Return the chart specifications produced for time-series graphs.
Public Property Get TimeSeriesCharts() As BetterArray
End Property

'Return a flattened list of all charts produced.
Public Property Get AllCharts() As BetterArray
End Property

'Return the total number of charts produced.
Public Property Get GraphCount() As Long
End Property

'Flag indicating whether any charts were produced.
Public Property Get HasGraphs() As Boolean
End Property

'Flag indicating whether time-series charts were produced.
Public Property Get HasTimeSeriesCharts() As Boolean
End Property
