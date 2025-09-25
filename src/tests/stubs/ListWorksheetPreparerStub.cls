VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListWorksheetPreparerStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub IListWorksheetPreparer tracking lifecycle calls")
'@IgnoreModule UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IListWorksheetPreparer

Private Type TPreparerState
    beginCount As Long
    completeCount As Long
    abortCount As Long
    lastAbortMessage As String
End Type

Private this As TPreparerState

Public Property Get BeginCount() As Long
    BeginCount = this.beginCount
End Property

Public Property Get CompleteCount() As Long
    CompleteCount = this.completeCount
End Property

Public Property Get AbortCount() As Long
    AbortCount = this.abortCount
End Property

Public Property Get LastAbortMessage() As String
    LastAbortMessage = this.lastAbortMessage
End Property

'@section Interface implementation
'===============================================================================
Private Sub IListWorksheetPreparer_Begin(ByVal context As IListBuildContext)
    this.beginCount = this.beginCount + 1
End Sub

Private Sub IListWorksheetPreparer_Complete(ByVal context As IListBuildContext)
    this.completeCount = this.completeCount + 1
End Sub

Private Sub IListWorksheetPreparer_Abort(ByVal context As IListBuildContext, _
                                         ByVal errNumber As Long, _
                                         ByVal messageText As String)
    this.abortCount = this.abortCount + 1
    this.lastAbortMessage = messageText
End Sub

