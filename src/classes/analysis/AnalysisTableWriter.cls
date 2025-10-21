VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Delegates table plan rendering to AnalysisOutputBase with a table strategy")
'@depends AnalysisOutputBase, AnalysisTableOutputStrategy
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTableWriter

Private runner As AnalysisOutputBase
Private tableStrategy As AnalysisTableOutputStrategy

'@section Factory
'===============================================================================
'Create a new writer instance.
Public Function Create() As IAnalysisTableWriter
    With New AnalysisTableWriter
        Set Create = .Self
    End With
End Function

'Expose the instance as its interface type.
Public Property Get Self() As IAnalysisTableWriter
    Set Self = Me
End Property

Public Property Set StackBuilder(ByVal builder As IAnalysisGoToStackBuilder)
    StrategyInternal.StackBuilder = builder
End Property

Private Property Get RunnerInternal() As AnalysisOutputBase
    If runner Is Nothing Then
        Set runner = New AnalysisOutputBase
    End If

    Set RunnerInternal = runner
End Property

Private Property Get StrategyInternal() As AnalysisTableOutputStrategy
    If tableStrategy Is Nothing Then
        Set tableStrategy = New AnalysisTableOutputStrategy
    End If

    Set StrategyInternal = tableStrategy
End Property

'@section Interface implementation
'===============================================================================
Private Sub IAnalysisTableWriter_Write(ByVal plan As IAnalysisTablePlanResult, _
                                       ByVal outputSheet As Worksheet, _
                                       ByVal context As IAnalysisTableWriterContext)
    RunnerInternal.Execute plan, outputSheet, context, StrategyInternal
End Sub
