VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTablePlanResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Concrete container for analysis table plan items and navigation labels")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTablePlanResult

Private Type TPlanState
    items As BetterArray
    sectionLabels As BetterArray
    headerLabels As BetterArray
    sectionPrefix As String
    headerPrefix As String
End Type

Private this As TPlanState
Private Const CLASSNAME As String = "AnalysisTablePlanResult"

'@section Factory
'===============================================================================
'Create a plan result using the supplied collections.
Public Function Create(ByVal items As BetterArray, _
                       ByVal sectionLabels As BetterArray, _
                       ByVal headerLabels As BetterArray, _
                       Optional ByVal sectionPrefix As String = vbNullString, _
                       Optional ByVal headerPrefix As String = vbNullString) As IAnalysisTablePlanResult
    ValidateArray items, "items"
    ValidateArray sectionLabels, "sectionLabels"
    ValidateArray headerLabels, "headerLabels"

    With New AnalysisTablePlanResult
        Set .Items = items
        Set .SectionLabels = sectionLabels
        Set .HeaderLabels = headerLabels
        .SectionPrefix = sectionPrefix
        .HeaderPrefix = headerPrefix
        Set Create = .Self
    End With
End Function

'Expose the current instance as the plan interface.
Public Property Get Self() As IAnalysisTablePlanResult
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

Public Property Set SectionLabels(ByVal sectionLabels As BetterArray)
    ValidateArray sectionLabels, "sectionLabels"
    Set this.sectionLabels = sectionLabels
End Property

Private Property Get SectionLabelsInternal() As BetterArray
    Set SectionLabelsInternal = this.sectionLabels
End Property

Public Property Set HeaderLabels(ByVal headerLabels As BetterArray)
    ValidateArray headerLabels, "headerLabels"
    Set this.headerLabels = headerLabels
End Property

Private Property Get HeaderLabelsInternal() As BetterArray
    Set HeaderLabelsInternal = this.headerLabels
End Property

Public Property Let SectionPrefix(ByVal sectionPrefix As String)
    this.sectionPrefix = sectionPrefix
End Property

Private Property Get SectionPrefixInternal() As String
    SectionPrefixInternal = this.sectionPrefix
End Property

Public Property Let HeaderPrefix(ByVal headerPrefix As String)
    this.headerPrefix = headerPrefix
End Property

Private Property Get HeaderPrefixInternal() As String
    HeaderPrefixInternal = this.headerPrefix
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
Private Property Get IAnalysisTablePlanResult_Items() As BetterArray
    Set IAnalysisTablePlanResult_Items = ItemsInternal()
End Property

Private Property Get IAnalysisTablePlanResult_SectionLabels() As BetterArray
    Set IAnalysisTablePlanResult_SectionLabels = SectionLabelsInternal()
End Property

Private Property Get IAnalysisTablePlanResult_HeaderLabels() As BetterArray
    Set IAnalysisTablePlanResult_HeaderLabels = HeaderLabelsInternal()
End Property

Private Property Get IAnalysisTablePlanResult_SectionPrefix() As String
    IAnalysisTablePlanResult_SectionPrefix = SectionPrefixInternal()
End Property

Private Property Get IAnalysisTablePlanResult_HeaderPrefix() As String
    IAnalysisTablePlanResult_HeaderPrefix = HeaderPrefixInternal()
End Property
