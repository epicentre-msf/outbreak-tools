VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTablePlanBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Builds table plan results from policy evaluation output")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Public Function Build(ByVal policyResults As BetterArray, _
                      Optional ByVal sectionLabelPrefix As String = vbNullString, _
                      Optional ByVal headerLabelPrefix As String = vbNullString) As IAnalysisTablePlanResult
End Function
