VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTableWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Writes analysis tables using prepared plan results and orchestrates downstream formatting")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Write the tables defined by the supplied plan to the target worksheet context.
Public Sub Write(ByVal plan As IAnalysisTablePlanResult, _
                 ByVal outputSheet As Worksheet, _
                 ByVal context As IAnalysisTableWriterContext)
End Sub
