VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisGraphPlanItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Describes a graph-enabled table plan entry produced during analysis orchestration")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the originating plan item associated with the graph request.
Public Property Get PlanItem() As IAnalysisTablePlanItem
End Property

'Expose the policy evaluation that marked the table as graph-capable.
Public Property Get PolicyResult() As IAnalysisTablePolicyResult
End Property

'Return the concrete table specification tied to the plan item.
Public Property Get Specification() As ITablesSpecs
End Property

'Return the analysis table type code for downstream routing decisions.
Public Property Get TableType() As AnalysisTablesType
End Property

'Return the zero-based position of the table within the overall plan.
Public Property Get SequenceIndex() As Long
End Property

'Flag indicating whether the graph should be processed via the time-series orchestrator.
Public Property Get IsTimeSeries() As Boolean
End Property

'Flag indicating whether the graph belongs to the spatio-temporal family.
Public Property Get IsSpatioTemporal() As Boolean
End Property

'Return the preferred label to surface when building navigation entries.
Public Property Get GraphLabel() As String
End Property
