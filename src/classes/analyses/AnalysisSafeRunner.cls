VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisSafeRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Executes analysis actions with scoped error handling")
'@depends AnalysisErrorHelper
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisSafeRunner

Private helper As AnalysisErrorHelper

Private Sub Class_Initialize()
    Set helper = New AnalysisErrorHelper
End Sub


Public Sub Run(ByVal action As IAnalysisAction, _
               Optional ByVal contextDescription As String = "Analysis operation")
    IAnalysisSafeRunner_Run action, contextDescription
End Sub

Private Sub IAnalysisSafeRunner_Run(ByVal action As IAnalysisAction, _
                                    Optional ByVal contextDescription As String = "Analysis operation")

    ValidateAction action

    On Error GoTo Handler
        action.Execute
    On Error GoTo 0
    Exit Sub

Handler:
    helper.RaiseUnexpectedState contextDescription & ": " & Err.Description
End Sub

Private Sub ValidateAction(ByVal action As IAnalysisAction)
    If action Is Nothing Then
        helper.RaiseInvalidArgument "action"
    End If
End Sub
