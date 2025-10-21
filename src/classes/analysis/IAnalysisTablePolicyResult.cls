VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTablePolicyResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Encapsulates the outcome of applying a table-type policy to an analysis specification")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the iteration item for which the policy was resolved.
Public Property Get IterationItem() As IAnalysisTableIterationItem
End Property

'Return the resolved table-type strategy.
Public Property Get Policy() As ITableTypePolicy
End Property

'Expose the reusable policy context built for the specification row.
Public Property Get Context() As ITableSpecsPolicyContext
End Property

'Indicate whether the specification satisfies the policy requirements.
Public Property Get IsValid() As Boolean
End Property

'Indicate whether the policy exposes percentage output.
Public Property Get HasPercent() As Boolean
End Property

'Indicate whether the policy exposes totals.
Public Property Get HasTotal() As Boolean
End Property

'Indicate whether the policy exposes missing values.
Public Property Get HasMissing() As Boolean
End Property

'Indicate whether downstream graph generation is required.
Public Property Get HasGraph() As Boolean
End Property
