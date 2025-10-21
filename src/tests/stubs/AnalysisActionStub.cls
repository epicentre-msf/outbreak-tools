VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisActionStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Test double implementing IAnalysisAction")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

Implements IAnalysisAction

Private shouldRaise As Boolean
Private raiseNumber As Long
Private raiseDescription As String
Public ExecuteCount As Long

Public Sub ConfigureError(ByVal errNumber As Long, ByVal errDescription As String)
    shouldRaise = True
    raiseNumber = errNumber
    raiseDescription = errDescription
End Sub

Private Sub IAnalysisAction_Execute()
    ExecuteCount = ExecuteCount + 1

    If shouldRaise Then
        Err.Raise raiseNumber, "AnalysisActionStub", raiseDescription
    End If
End Sub
