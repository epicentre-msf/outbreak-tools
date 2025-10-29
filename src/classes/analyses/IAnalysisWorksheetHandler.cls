VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisWorksheetHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Callback interface invoked by AnalysisWorksheetPipeline for each worksheet specification")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Invoked before processing any worksheet specifications.
Public Sub BeginSequence()
End Sub

'Process the supplied worksheet specification.
Public Sub ProcessWorksheet(ByVal spec As IAnalysisWorksheetSpec)
End Sub

'Invoked after all worksheet specifications have been processed.
Public Sub CompleteSequence()
End Sub
