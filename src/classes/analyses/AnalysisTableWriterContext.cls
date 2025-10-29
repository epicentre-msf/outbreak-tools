VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableWriterContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Baseline implementation of IAnalysisTableWriterContext capturing plan execution events")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTableWriterContext

Private Type TContextState
    planActive As Boolean
    writtenItems As BetterArray
    sectionLabels As BetterArray
    headerLabels As BetterArray
End Type

Private this As TContextState
Private Const CLASSNAME As String = "AnalysisTableWriterContext"

Private Sub Class_Initialize()
    Set this.writtenItems = New BetterArray
    this.writtenItems.LowerBound = 1
    Set this.sectionLabels = New BetterArray
    this.sectionLabels.LowerBound = 1
    Set this.headerLabels = New BetterArray
    this.headerLabels.LowerBound = 1
End Sub

'@section Public helpers
'===============================================================================
Public Property Get WrittenItems() As BetterArray
    Set WrittenItems = this.writtenItems
End Property

Public Property Get SectionLabels() As BetterArray
    Set SectionLabels = this.sectionLabels
End Property

Public Property Get HeaderLabels() As BetterArray
    Set HeaderLabels = this.headerLabels
End Property

'@section Validation helpers
'===============================================================================
Private Sub ValidatePlanItem(ByVal planItem As IAnalysisTablePlanItem)
    If planItem Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "planItem cannot be Nothing"
    End If
End Sub

Private Sub ValidateOutputSheet(ByVal outputSheet As Worksheet)
    If outputSheet Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "outputSheet cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Sub IAnalysisTableWriterContext_BeginPlan(ByVal outputSheet As Worksheet)
    ValidateOutputSheet outputSheet
    this.planActive = True
    this.writtenItems.Clear
    this.sectionLabels.Clear
    this.headerLabels.Clear
End Sub

Private Sub IAnalysisTableWriterContext_WriteTable(ByVal planItem As IAnalysisTablePlanItem, ByVal outputSheet As Worksheet)
    ValidateOutputSheet outputSheet
    ValidatePlanItem planItem

    If Not this.planActive Then
        RaiseError CLng(ProjectError.ErrorUnexpectedState), "BeginPlan must be called before WriteTable"
    End If

    this.writtenItems.Push planItem
End Sub

Private Sub IAnalysisTableWriterContext_ApplyNavigation(ByVal sectionLabels As BetterArray, _
                                                        ByVal headerLabels As BetterArray, _
                                                        ByVal outputSheet As Worksheet)
    ValidateOutputSheet outputSheet

    If Not sectionLabels Is Nothing Then Set this.sectionLabels = sectionLabels
    If Not headerLabels Is Nothing Then Set this.headerLabels = headerLabels
End Sub

Private Sub IAnalysisTableWriterContext_CompletePlan(ByVal outputSheet As Worksheet)
    ValidateOutputSheet outputSheet
    this.planActive = False
End Sub
