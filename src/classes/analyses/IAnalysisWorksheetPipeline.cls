VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisWorksheetPipeline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Coordinates execution of worksheet specifications via handlers")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Execute the supplied sequence of worksheet specifications with the provided handler.
Public Sub Run(ByVal sequence As BetterArray, ByVal handler As IAnalysisWorksheetHandler)
End Sub
