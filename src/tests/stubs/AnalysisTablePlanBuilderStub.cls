VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTablePlanBuilderStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisTablePlanBuilder stub returning a configurable plan result")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisTablePlanBuilder

Private Type TPlanBuilderStubState
    planResult As IAnalysisTablePlanResult
    lastPolicyResults As BetterArray
    lastSectionPrefix As String
    lastHeaderPrefix As String
    buildCount As Long
End Type

Private this As TPlanBuilderStubState

Public Sub SetPlanResult(ByVal planResult As IAnalysisTablePlanResult)
    Set this.planResult = planResult
End Sub

Public Property Get LastPolicyResults() As BetterArray
    Set LastPolicyResults = this.lastPolicyResults
End Property

Public Property Get LastSectionPrefix() As String
    LastSectionPrefix = this.lastSectionPrefix
End Property

Public Property Get LastHeaderPrefix() As String
    LastHeaderPrefix = this.lastHeaderPrefix
End Property

Public Property Get BuildCount() As Long
    BuildCount = this.buildCount
End Property

Private Function IAnalysisTablePlanBuilder_Build(ByVal policyResults As BetterArray, _
                                                 Optional ByVal sectionLabelPrefix As String = vbNullString, _
                                                 Optional ByVal headerLabelPrefix As String = vbNullString) As IAnalysisTablePlanResult
    Set this.lastPolicyResults = policyResults
    this.lastSectionPrefix = sectionLabelPrefix
    this.lastHeaderPrefix = headerLabelPrefix
    this.buildCount = this.buildCount + 1
    Set IAnalysisTablePlanBuilder_Build = this.planResult
End Function
