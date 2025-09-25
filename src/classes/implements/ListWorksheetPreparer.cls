VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListWorksheetPreparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Linelist.Shared")
'@ModuleDescription("Default worksheet preparer tracking busy state for list builders")
'@depends ProjectError
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IListWorksheetPreparer

Private Type TPreparerState
    isBusy As Boolean
    beginCount As Long
    completeCount As Long
    abortCount As Long
    lastWorksheet As String
End Type

Private this As TPreparerState
Private Const CLASSNAME As String = "ListWorksheetPreparer"

'@section Diagnostics
'===============================================================================
'Expose whether a Begin/Complete pair is currently active.
Public Property Get IsBusy() As Boolean
    IsBusy = this.isBusy
End Property

'Number of successful Begin calls (useful for tests/monitoring).
Public Property Get BeginCount() As Long
    BeginCount = this.beginCount
End Property

'Number of successful Complete calls.
Public Property Get CompleteCount() As Long
    CompleteCount = this.completeCount
End Property

'Number of Abort calls triggered by upstream failures.
Public Property Get AbortCount() As Long
    AbortCount = this.abortCount
End Property

'Worksheet processed during the most recent Begin call.
Public Property Get LastWorksheetName() As String
    LastWorksheetName = this.lastWorksheet
End Property

'@section Helpers
'===============================================================================
Private Sub EnsureContextReady(ByVal context As IListBuildContext)
    If context Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Build context cannot be Nothing when preparing worksheets"
    End If

    If Trim$(context.WorksheetName) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Worksheet name is required when preparing worksheets"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Sub IListWorksheetPreparer_Begin(ByVal context As IListBuildContext)
    EnsureContextReady context

    If this.isBusy Then
        ThrowError CLng(ProjectError.ErrorUnexpectedState), _
                  "Worksheet preparer is already busy"
    End If

    this.isBusy = True
    this.beginCount = this.beginCount + 1
    this.lastWorksheet = Trim$(context.WorksheetName)
End Sub

Private Sub IListWorksheetPreparer_Complete(ByVal context As IListBuildContext)
    EnsureContextReady context

    If Not this.isBusy Then
        ThrowError CLng(ProjectError.ErrorUnexpectedState), _
                  "Worksheet preparer cannot complete without an active session"
    End If

    this.isBusy = False
    this.completeCount = this.completeCount + 1
End Sub

Private Sub IListWorksheetPreparer_Abort(ByVal context As IListBuildContext, _
                                         ByVal errNumber As Long, _
                                         ByVal messageText As String)
    'Abort should never surface additional failures; validation is intentionally lenient.
    On Error Resume Next
        If Not context Is Nothing Then this.lastWorksheet = Trim$(context.WorksheetName)
    On Error GoTo 0

    this.abortCount = this.abortCount + 1
    this.isBusy = False
End Sub

