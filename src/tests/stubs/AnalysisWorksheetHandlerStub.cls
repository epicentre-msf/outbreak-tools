VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisWorksheetHandlerStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementing IAnalysisWorksheetHandler for pipeline tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisWorksheetHandler

Private Type THandlerStubState
    beginCount As Long
    processed As BetterArray
    completeCount As Long
End Type

Private this As THandlerStubState

Private Sub Class_Initialize()
    Set this.processed = New BetterArray
    this.processed.LowerBound = 1
End Sub

Public Property Get BeginCount() As Long
    BeginCount = this.beginCount
End Property

Public Property Get CompleteCount() As Long
    CompleteCount = this.completeCount
End Property

Public Property Get Processed() As BetterArray
    Set Processed = this.processed
End Property

Private Sub IAnalysisWorksheetHandler_BeginSequence()
    this.beginCount = this.beginCount + 1
End Sub

Private Sub IAnalysisWorksheetHandler_ProcessWorksheet(ByVal spec As IAnalysisWorksheetSpec)
    this.processed.Push spec
End Sub

Private Sub IAnalysisWorksheetHandler_CompleteSequence()
    this.completeCount = this.completeCount + 1
End Sub
