VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTablePlanItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Represents a single table entry scheduled for writing during analysis output")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the policy evaluation driving the table output.
Public Property Get PolicyResult() As IAnalysisTablePolicyResult
End Property

'Return the zero-based position of the table within the planned sequence.
Public Property Get SequenceIndex() As Long
End Property
