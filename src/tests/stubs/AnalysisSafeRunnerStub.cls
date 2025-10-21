VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisSafeRunnerStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of IAnalysisSafeRunner that executes actions inline")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

Implements IAnalysisSafeRunner

Public RunCount As Long
Public LastContext As String

Private Sub IAnalysisSafeRunner_Run(ByVal action As IAnalysisAction, _
                                    Optional ByVal contextDescription As String = "Analysis operation")
    RunCount = RunCount + 1
    LastContext = contextDescription
    action.Execute
End Sub
