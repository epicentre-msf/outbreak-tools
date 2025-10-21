VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTablePlanBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Transforms policy-evaluated table results into a write plan with navigation metadata")
'@depends ProjectError, AnalysisTablePlanResult, AnalysisTablePlanItem
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTablePlanBuilder

Private Const CLASSNAME As String = "AnalysisTablePlanBuilder"

'@section Factory
'===============================================================================
'Create a plan result using the supplied policy evaluation output.
Public Function Build(ByVal policyResults As BetterArray, _
                      Optional ByVal sectionLabelPrefix As String = vbNullString, _
                      Optional ByVal headerLabelPrefix As String = vbNullString) As IAnalysisTablePlanResult

    ValidatePolicyResults policyResults

    Dim items As BetterArray
    Dim sectionLabels As BetterArray
    Dim headerLabels As BetterArray
    Dim resultIndex As Long
    Dim policyResult As IAnalysisTablePolicyResult
    Dim iterationItem As IAnalysisTableIterationItem
    Dim sectionLabel As String
    Dim headerLabel As String
    Dim spec As ITablesSpecs

    Set items = CreateBuffer()
    Set sectionLabels = CreateBuffer()
    Set headerLabels = CreateBuffer()

    For resultIndex = policyResults.LowerBound To policyResults.UpperBound
        Set policyResult = policyResults.Item(resultIndex)
        If policyResult Is Nothing Then GoTo ContinueLoop
        If Not policyResult.IsValid Then GoTo ContinueLoop

        Set iterationItem = policyResult.IterationItem
        Set spec = iterationItem.Specification

        items.Push AnalysisTablePlanItem.Create(policyResult, resultIndex - policyResults.LowerBound)

        If iterationItem.IsNewSection Then
            sectionLabel = Trim$(spec.Value("section"))
            If sectionLabel <> vbNullString Then
                sectionLabels.Push sectionLabelPrefix & sectionLabel
            End If
        End If

        headerLabel = Trim$(spec.Value("label"))
        If headerLabel <> vbNullString Then
            headerLabels.Push headerLabelPrefix & headerLabel
        End If

ContinueLoop:
        Set policyResult = Nothing
        Set iterationItem = Nothing
        Set spec = Nothing
    Next resultIndex

    Set Build = AnalysisTablePlanResult.Create(items, sectionLabels, headerLabels, sectionLabelPrefix, headerLabelPrefix)
End Function

'@section Helpers
'===============================================================================
Private Function CreateBuffer() As BetterArray
    Dim buffer As BetterArray
    Set buffer = New BetterArray
    buffer.LowerBound = 1
    Set CreateBuffer = buffer
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidatePolicyResults(ByVal policyResults As BetterArray)
    If policyResults Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "policyResults cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IAnalysisTablePlanBuilder_Build(ByVal policyResults As BetterArray, _
                                                 Optional ByVal sectionLabelPrefix As String = vbNullString, _
                                                 Optional ByVal headerLabelPrefix As String = vbNullString) As IAnalysisTablePlanResult
    Set IAnalysisTablePlanBuilder_Build = Build(policyResults, sectionLabelPrefix, headerLabelPrefix)
End Function
