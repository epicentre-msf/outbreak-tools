VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTablePolicyResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Coordinates policy evaluation for analysis table iteration items")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Evaluate the supplied iteration item and return the resolved policy metadata.
Public Function Resolve(ByVal iterationItem As IAnalysisTableIterationItem) As IAnalysisTablePolicyResult
End Function

'Reset cached dictionary/variables references.
Public Sub Invalidate()
End Sub
