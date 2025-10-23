Attribute VB_Name = "TestShowHideState"

Option Explicit

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private Assert As ICustomTest
Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ShowHideState"


'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestShowHideState"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Assert.BeginTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@TestMethod("ShowHide")
Public Sub TestShowHideRuleForceHiddenAppliesToCRFPrinted()
    CustomTestSetTitles Assert, TESTMODULE, "TestShowHideRuleForceHiddenAppliesToCRFPrinted"

    Dim rule As IShowHideRule
    Set rule = ShowHideRule.Create("field_a", "Field A")
    rule.MarkLayerApplicability ShowHideLayerCRF, True
    rule.MarkLayerApplicability ShowHideLayerPrinted, True
    rule.ForceHidden = True

    Assert.IsTrue rule.EffectiveHidden(ShowHideLayerCRF), _
                  "ForceHidden should hide the column on CRF worksheets"
    Assert.IsTrue rule.EffectiveHidden(ShowHideLayerPrinted), _
                  "ForceHidden should hide the column on printed worksheets"
    Assert.IsFalse rule.EffectiveHidden(ShowHideLayerHList), _
                   "ForceHidden should not hide the column on HList by default"
End Sub


'@TestMethod("ShowHide")
Public Sub TestShowHideStateEvaluateHiddenUsesUserChoice()
    CustomTestSetTitles Assert, TESTMODULE, "TestShowHideStateEvaluateHiddenUsesUserChoice"

    Dim state As IShowHideState
    Dim rule As IShowHideRule
    Set state = ShowHideState.Create()
    Set rule = ShowHideRule.Create("field_b", "Field B")

    rule.MarkLayerApplicability ShowHideLayerHList, True
    state.AddRule rule
    state.UpdateLayerVisibility "field_b", ShowHideLayerHList, True

    Assert.IsTrue state.EvaluateHidden("field_b", ShowHideLayerHList), _
                  "EvaluateHidden should return the stored HList visibility"
    Assert.IsFalse state.EvaluateHidden("field_b", ShowHideLayerCRF), _
                   "EvaluateHidden should default to visible on other layers"
End Sub


'@TestMethod("ShowHide")
Public Sub TestShowHidePlanPreventsDuplicateEntries()
    CustomTestSetTitles Assert, TESTMODULE, "TestShowHidePlanPreventsDuplicateEntries"

    Dim plan As IShowHidePlan
    Dim errNumber As Long
    Dim hiddenState As Boolean
    Set plan = ShowHidePlan.Create(ShowHideLayerHList)

    plan.AddVisibility "field_a", "Field A", True

    On Error Resume Next
        plan.AddVisibility "field_a", "Field A", False
        errNumber = Err.Number
    On Error GoTo 0

    hiddenState = plan.VisibilityFor("field_a")

    Assert.AreEqual ProjectError.ElementShouldNotExists, errNumber, _
                    "Adding the same field twice should raise an ElementShouldNotExists error"
    Assert.IsTrue hiddenState, "VisibilityFor should return the first registered state"
End Sub


'@TestMethod("ShowHide")
Public Sub TestHListStrategyBuildsPlanWithRegisteredRules()
    CustomTestSetTitles Assert, TESTMODULE, "TestHListStrategyBuildsPlanWithRegisteredRules"

    Dim state As IShowHideState
    Dim ruleA As IShowHideRule
    Dim ruleB As IShowHideRule
    Dim strategy As IShowHideStrategy
    Dim plan As IShowHidePlan
    Dim actions As BetterArray
    Dim firstAction As Variant

    Set state = ShowHideState.Create()
    Set ruleA = ShowHideRule.Create("field_a", "Field A")
    ruleA.MarkLayerApplicability ShowHideLayerHList, True
    ruleA.SetLayerVisibility ShowHideLayerHList, False
    state.AddRule ruleA

    Set ruleB = ShowHideRule.Create("field_b", "Field B")
    ruleB.MarkLayerApplicability ShowHideLayerHList, True
    ruleB.SetLayerVisibility ShowHideLayerHList, True
    state.AddRule ruleB

    Set strategy = ShowHideHListStrategy.Create()
    Set plan = strategy.BuildPlan(state, Nothing)
    Set actions = plan.Actions
    firstAction = actions.Item(actions.LowerBound)

    Assert.AreEqual ShowHideLayerHList, plan.TargetLayer, _
                    "Plan should target the HList layer"
    Assert.AreEqual 2, actions.Length, "All registered rules should be represented in the plan"
    Assert.AreEqual "field_a", CStr(firstAction(0)), _
                    "Actions should preserve insertion order when building the plan"
End Sub


'@TestMethod("ShowHide")
Public Sub TestDictionaryStrategyIncludesDictionaryHiddenRules()
    CustomTestSetTitles Assert, TESTMODULE, "TestDictionaryStrategyIncludesDictionaryHiddenRules"

    Dim state As IShowHideState
    Dim rule As IShowHideRule
    Dim strategy As IShowHideStrategy
    Dim plan As IShowHidePlan
    Dim actions As BetterArray

    Set state = ShowHideState.Create()
    Set rule = ShowHideRule.Create("field_dict", "Field Dictionary")
    rule.DictionaryHidden = True
    state.AddRule rule

    Set strategy = ShowHideDictionaryStrategy.Create()
    Set plan = strategy.BuildPlan(state, Nothing)
    Set actions = plan.Actions

    Assert.AreEqual 1, actions.Length, _
                    "Dictionary strategy should include dictionary-hidden rules even without applicability flags"
    Assert.AreEqual True, CBool(actions.Item(actions.LowerBound)(2)), _
                    "Dictionary-hidden rule should be marked as hidden in the plan"
End Sub
