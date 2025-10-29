VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGraphPlanResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Concrete container exposing classified graph plan items")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGraphPlanResult

Private Type TGraphPlanResultState
    items As BetterArray
    crossTableItems As BetterArray
    timeSeriesItems As BetterArray
End Type

Private this As TGraphPlanResultState

Private Const CLASSNAME As String = "AnalysisGraphPlanResult"

'@section Factory
'===============================================================================
'Create a graph plan result from the supplied collections.
Public Function Create(ByVal items As BetterArray, _
                       ByVal crossTableItems As BetterArray, _
                       ByVal timeSeriesItems As BetterArray) As IAnalysisGraphPlanResult
    ValidateArray items, "items"
    ValidateArray crossTableItems, "crossTableItems"
    ValidateArray timeSeriesItems, "timeSeriesItems"

    With New AnalysisGraphPlanResult
        Set .Items = items
        Set .CrossTableItems = crossTableItems
        Set .TimeSeriesItems = timeSeriesItems
        Set Create = .Self
    End With
End Function

'Expose the instance through its interface type.
Public Property Get Self() As IAnalysisGraphPlanResult
    Set Self = Me
End Property

'@section Properties
'===============================================================================
Public Property Set Items(ByVal items As BetterArray)
    ValidateArray items, "items"
    Set this.items = items
End Property

Private Property Get ItemsInternal() As BetterArray
    Set ItemsInternal = this.items
End Property

Public Property Set CrossTableItems(ByVal crossTableItems As BetterArray)
    ValidateArray crossTableItems, "crossTableItems"
    Set this.crossTableItems = crossTableItems
End Property

Private Property Get CrossTableItemsInternal() As BetterArray
    Set CrossTableItemsInternal = this.crossTableItems
End Property

Public Property Set TimeSeriesItems(ByVal timeSeriesItems As BetterArray)
    ValidateArray timeSeriesItems, "timeSeriesItems"
    Set this.timeSeriesItems = timeSeriesItems
End Property

Private Property Get TimeSeriesItemsInternal() As BetterArray
    Set TimeSeriesItemsInternal = this.timeSeriesItems
End Property

'@section Metrics
'===============================================================================
Public Property Get ItemCount() As Long
    Dim items As BetterArray

    Set items = ItemsInternal
    If items Is Nothing Then
        ItemCount = 0
    Else
        ItemCount = items.Length
    End If
End Property

Public Property Get TimeSeriesCount() As Long
    Dim items As BetterArray

    Set items = TimeSeriesItemsInternal
    If items Is Nothing Then
        TimeSeriesCount = 0
    Else
        TimeSeriesCount = items.Length
    End If
End Property

'@section Validation helpers
'===============================================================================
Private Sub ValidateArray(ByVal candidate As BetterArray, ByVal argumentName As String)
    If candidate Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get IAnalysisGraphPlanResult_Items() As BetterArray
    Set IAnalysisGraphPlanResult_Items = ItemsInternal()
End Property

Private Property Get IAnalysisGraphPlanResult_CrossTableItems() As BetterArray
    Set IAnalysisGraphPlanResult_CrossTableItems = CrossTableItemsInternal()
End Property

Private Property Get IAnalysisGraphPlanResult_TimeSeriesItems() As BetterArray
    Set IAnalysisGraphPlanResult_TimeSeriesItems = TimeSeriesItemsInternal()
End Property

Private Property Get IAnalysisGraphPlanResult_GraphCount() As Long
    IAnalysisGraphPlanResult_GraphCount = ItemCount()
End Property

Private Property Get IAnalysisGraphPlanResult_HasGraphs() As Boolean
    IAnalysisGraphPlanResult_HasGraphs = (ItemCount() > 0)
End Property

Private Property Get IAnalysisGraphPlanResult_HasTimeSeriesGraphs() As Boolean
    IAnalysisGraphPlanResult_HasTimeSeriesGraphs = (TimeSeriesCount() > 0)
End Property
