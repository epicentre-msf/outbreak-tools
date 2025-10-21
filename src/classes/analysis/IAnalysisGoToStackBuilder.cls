VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisGoToStackBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Defines operations for computing section and header navigation stacks")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Public Property Let SectionPrefix(ByVal prefix As String)
End Property

Public Property Let HeaderPrefix(ByVal prefix As String)
End Property

Public Function Build(ByVal plan As IAnalysisTablePlanResult) As AnalysisGoToStackContext
End Function
