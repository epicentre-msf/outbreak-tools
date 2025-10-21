VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGoToStackBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Computes section/header navigation labels from plan items using prefix rules")
'@depends ProjectError, BetterArray, AnalysisGoToStackContext, AnalysisGoToEntry
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGoToStackBuilder

Private Type TStackBuilderState
    sectionPrefix As String
    headerPrefix As String
End Type

Private this As TStackBuilderState
Private Const CLASSNAME As String = "AnalysisGoToStackBuilder"

'@section Configuration
'===============================================================================
Public Property Let SectionPrefix(ByVal prefix As String)
    this.sectionPrefix = prefix
End Property

Public Property Get SectionPrefix() As String
    SectionPrefix = this.sectionPrefix
End Property

Public Property Let HeaderPrefix(ByVal prefix As String)
    this.headerPrefix = prefix
End Property

Public Property Get HeaderPrefix() As String
    HeaderPrefix = this.headerPrefix
End Property

'@section Factory
'===============================================================================
Public Function Create(Optional ByVal sectionPrefix As String = vbNullString, _
                       Optional ByVal headerPrefix As String = vbNullString) As AnalysisGoToStackBuilder
    Dim builder As AnalysisGoToStackBuilder
    Set builder = New AnalysisGoToStackBuilder
    builder.SectionPrefix = sectionPrefix
    builder.HeaderPrefix = headerPrefix
    Set Create = builder
End Function

'@section Build API
'===============================================================================
Public Function Build(ByVal plan As IAnalysisTablePlanResult) As AnalysisGoToStackContext
    Dim context As AnalysisGoToStackContext
    Dim items As BetterArray
    Dim idx As Long
    Dim planItem As IAnalysisTablePlanItem
    Dim policyResult As IAnalysisTablePolicyResult
    Dim iterationItem As IAnalysisTableIterationItem
    Dim specs As ITablesSpecs

    ValidatePlan plan

    Set context = New AnalysisGoToStackContext
    context.Reset

    Set items = plan.Items
    If items Is Nothing Then
        Set Build = context
        Exit Function
    End If

    For idx = items.LowerBound To items.UpperBound
        Set planItem = items.Item(idx)
        If planItem Is Nothing Then GoTo ContinueLoop

        Set policyResult = planItem.PolicyResult
        If policyResult Is Nothing Then GoTo ContinueLoop

        Set iterationItem = policyResult.IterationItem
        If iterationItem Is Nothing Then GoTo ContinueLoop

        Set specs = iterationItem.Specification
        If specs Is Nothing Then GoTo ContinueLoop

        If iterationItem.IsNewSection Then
            AddSectionEntry context, specs
        End If

        AddHeaderEntry context, specs

ContinueLoop:
        Set planItem = Nothing
        Set policyResult = Nothing
        Set iterationItem = Nothing
        Set specs = Nothing
    Next idx

    Set Build = context
End Function

Private Sub AddSectionEntry(ByVal context As AnalysisGoToStackContext, ByVal specs As ITablesSpecs)
    Dim labelText As String
    Dim entry As IAnalysisGoToEntry

    labelText = Trim$(CStr(specs.Value("section")))
    If LenB(labelText) = 0 Then Exit Sub

    Set entry = AnalysisGoToEntry.Create("section", labelText, this.sectionPrefix, Trim$(CStr(specs.TableId)))
    context.AddSectionEntry entry
End Sub

Private Sub AddHeaderEntry(ByVal context As AnalysisGoToStackContext, ByVal specs As ITablesSpecs)
    Dim labelText As String
    Dim entry As IAnalysisGoToEntry

    labelText = Trim$(CStr(specs.Value("label")))
    If LenB(labelText) = 0 Then Exit Sub

    Set entry = AnalysisGoToEntry.Create("header", labelText, this.headerPrefix, Trim$(CStr(specs.TableId)))
    context.AddHeaderEntry entry
End Sub

Private Sub ValidatePlan(ByVal plan As IAnalysisTablePlanResult)
    If plan Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "plan cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Let IAnalysisGoToStackBuilder_SectionPrefix(ByVal prefix As String)
    SectionPrefix = prefix
End Property

Private Property Let IAnalysisGoToStackBuilder_HeaderPrefix(ByVal prefix As String)
    HeaderPrefix = prefix
End Property

Private Function IAnalysisGoToStackBuilder_Build(ByVal plan As IAnalysisTablePlanResult) As AnalysisGoToStackContext
    Set IAnalysisGoToStackBuilder_Build = Build(plan)
End Function
