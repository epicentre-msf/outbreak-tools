VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGraphExecutionResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Concrete container aggregating chart specifications emitted by AnalysisGraphEngine")
'@depends ProjectError, AnalysisArrayUtilities
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGraphExecutionResult

Private Type TExecutionResultState
    crossItems As BetterArray
    timeSeriesCharts As BetterArray
    allCharts As BetterArray
End Type

Private this As TExecutionResultState

Private Const CLASSNAME As String = "AnalysisGraphExecutionResult"

Private Sub Class_Initialize()
    Set this.crossItems = AnalysisArrayUtilities.CreateBuffer()
    Set this.timeSeriesCharts = AnalysisArrayUtilities.CreateBuffer()
    Set this.allCharts = AnalysisArrayUtilities.CreateBuffer()
End Sub

'@section Public API
'===============================================================================
Public Sub AddCrossTableItem(ByVal item As IAnalysisGraphExecutionItem)
    ValidateItem item, "item"
    this.crossItems.Push item
    AppendCharts item.Charts
End Sub

Public Sub SetTimeSeriesCharts(ByVal charts As BetterArray)
    ValidateCharts charts, "charts"
    Set this.timeSeriesCharts = CloneCharts(charts)
    AppendCharts charts
End Sub

Private Sub AppendCharts(ByVal charts As BetterArray)
    Dim clones As BetterArray
    Dim idx As Long

    Set clones = AnalysisArrayUtilities.CloneCloneable(charts)

    If clones Is Nothing Then Exit Sub
    If clones.Length = 0 Then Exit Sub

    For idx = clones.LowerBound To clones.UpperBound
        this.allCharts.Push clones.Item(idx)
    Next idx
End Sub

Private Function CloneCharts(ByVal charts As BetterArray) As BetterArray
    Set CloneCharts = AnalysisArrayUtilities.CloneCloneable(charts)
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidateItem(ByVal item As IAnalysisGraphExecutionItem, ByVal argumentName As String)
    If item Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub ValidateCharts(ByVal charts As BetterArray, ByVal argumentName As String)
    If charts Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisGraphExecutionResult_CrossTableItems() As BetterArray
    Set IAnalysisGraphExecutionResult_CrossTableItems = AnalysisArrayUtilities.CloneItems(this.crossItems)
End Property

Private Property Get IAnalysisGraphExecutionResult_TimeSeriesCharts() As BetterArray
    Set IAnalysisGraphExecutionResult_TimeSeriesCharts = CloneCharts(this.timeSeriesCharts)
End Property

Private Property Get IAnalysisGraphExecutionResult_AllCharts() As BetterArray
    Set IAnalysisGraphExecutionResult_AllCharts = CloneCharts(this.allCharts)
End Property

Private Property Get IAnalysisGraphExecutionResult_GraphCount() As Long
    If this.allCharts Is Nothing Then
        IAnalysisGraphExecutionResult_GraphCount = 0
    Else
        IAnalysisGraphExecutionResult_GraphCount = this.allCharts.Length
    End If
End Property

Private Property Get IAnalysisGraphExecutionResult_HasGraphs() As Boolean
    IAnalysisGraphExecutionResult_HasGraphs = (IAnalysisGraphExecutionResult_GraphCount() > 0)
End Property

Private Property Get IAnalysisGraphExecutionResult_HasTimeSeriesCharts() As Boolean
    If this.timeSeriesCharts Is Nothing Then
        IAnalysisGraphExecutionResult_HasTimeSeriesCharts = False
    Else
        IAnalysisGraphExecutionResult_HasTimeSeriesCharts = (this.timeSeriesCharts.Length > 0)
    End If
End Property
