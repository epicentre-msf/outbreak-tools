VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisGraphEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Coordinates graph planning and orchestrator execution for analysis outputs")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Evaluate the supplied plan and inputs to produce graph chart specifications.
Public Function Run(ByVal plan As IAnalysisTablePlanResult, _
                    ByVal inputs As IAnalysisGraphEngineInputs) As IAnalysisGraphExecutionResult
End Function
