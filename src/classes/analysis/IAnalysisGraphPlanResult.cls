VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisGraphPlanResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Aggregates graph-capable plan items with convenient categorisation")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return every graph request discovered within the table plan.
Public Property Get Items() As BetterArray
End Property

'Return graph requests tied to cross-table analysis outputs.
Public Property Get CrossTableItems() As BetterArray
End Property

'Return graph requests routed through the time-series graph orchestrator.
Public Property Get TimeSeriesItems() As BetterArray
End Property

'Return the total number of graph requests identified.
Public Property Get GraphCount() As Long
End Property

'Indicates whether the plan contains any graph work.
Public Property Get HasGraphs() As Boolean
End Property

'Indicates whether at least one time-series graph is required.
Public Property Get HasTimeSeriesGraphs() As Boolean
End Property
