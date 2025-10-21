VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableWriterContextStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("IAnalysisTableWriterContext stub capturing call order for verification")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisTableWriterContext

Public BeginPlanCount As Long
Public CompletePlanCount As Long
Public ApplyNavigationCount As Long
Public LastSectionLabels As BetterArray
Public LastHeaderLabels As BetterArray
Public WrittenItems As BetterArray

Private Sub Class_Initialize()
    Set WrittenItems = New BetterArray
    WrittenItems.LowerBound = 1
End Sub

Private Sub IAnalysisTableWriterContext_BeginPlan(ByVal outputSheet As Worksheet)
    BeginPlanCount = BeginPlanCount + 1
End Sub

Private Sub IAnalysisTableWriterContext_WriteTable(ByVal planItem As IAnalysisTablePlanItem, ByVal outputSheet As Worksheet)
    WrittenItems.Push planItem
End Sub

Private Sub IAnalysisTableWriterContext_ApplyNavigation(ByVal sectionLabels As BetterArray, _
                                                        ByVal headerLabels As BetterArray, _
                                                        ByVal outputSheet As Worksheet)
    ApplyNavigationCount = ApplyNavigationCount + 1
    Set LastSectionLabels = sectionLabels
    Set LastHeaderLabels = headerLabels
End Sub

Private Sub IAnalysisTableWriterContext_CompletePlan(ByVal outputSheet As Worksheet)
    CompletePlanCount = CompletePlanCount + 1
End Sub
