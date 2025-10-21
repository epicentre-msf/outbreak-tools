VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisWorksheetCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Coordinates worksheet pipeline, table writer, and graph engine during analysis output")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Execute the worksheet sequence with the given inputs.
Public Sub Run(ByVal sequence As BetterArray, ByVal inputs As IAnalysisWorksheetInputs)
End Sub
