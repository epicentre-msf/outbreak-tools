VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisSafeRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Defines a contract for executing analysis actions with error management")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Public Sub Run(ByVal action As IAnalysisAction, _
               Optional ByVal contextDescription As String = "Analysis operation")
End Sub
