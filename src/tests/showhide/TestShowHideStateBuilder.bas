Attribute VB_Name = "TestShowHideStateBuilder"

Option Explicit

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private DictionarySheet As Worksheet
Private Dict As ILLdictionary
Private Variables As ILLVariables

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ShowHideStateBuilder"
Private Const DICTIONARY_SHEET As String = "DictionaryFixture"


'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestShowHideStateBuilder"
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
    BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, FixtureWorkbook
    Set DictionarySheet = FixtureWorkbook.Worksheets(DICTIONARY_SHEET)
    Set Dict = LLdictionary.Create(DictionarySheet, 1, 1)
    Dict.Prepare
    Set Variables = LLVariables.Create(Dict)
    Assert.BeginTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then
            TestHelpers.DeleteWorkbook FixtureWorkbook
        End If
    On Error GoTo 0

    Set Variables = Nothing
    Set Dict = Nothing
    Set DictionarySheet = Nothing
    Set FixtureWorkbook = Nothing
End Sub


'@TestMethod("ShowHide")
Public Sub TestBuildCreatesRulesForVList()
    CustomTestSetTitles Assert, TESTMODULE, "TestBuildCreatesRulesForVList"

    Dim builder As IShowHideStateBuilder
    Dim state As IShowHideState

    Set builder = ShowHideStateBuilder.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1", Variables)
    Set state = builder.Build

    Assert.IsTrue state.HasRule("hid_beg_v1"), _
                  "State should contain variables assigned to the target sheet"
    Assert.IsTrue state.EvaluateHidden("hid_beg_v1", ShowHideLayerVList), _
                  "Builder should mark dictionary-hidden variables as hidden"
    Assert.IsFalse state.EvaluateHidden("opt_vis_v1", ShowHideLayerVList), _
                   "Optional visible variables should default to visible"
    Assert.IsTrue state.Rule("mand_v1").Mandatory, _
                 "Mandatory dictionary entries must be flagged on the rule"
End Sub


'@TestMethod("ShowHide")
Public Sub TestForceHiddenMarksFormulaControlsOnCRF()
    CustomTestSetTitles Assert, TESTMODULE, "TestForceHiddenMarksFormulaControlsOnCRF"

    Dim builder As IShowHideStateBuilder
    Dim state As IShowHideState
    Dim rule As IShowHideRule

    Set builder = ShowHideStateBuilder.Create(Dict, ShowHideLayerCRF, "hlist2D-sheet2", Variables)
    Set state = builder.Build
    Set rule = state.Rule("val_of_text_h2")

    Assert.IsTrue rule.ForceHidden, _
                  "Formula-based variables must be force-hidden on CRF layer"
End Sub


'@TestMethod("ShowHide")
Public Sub TestPlanOverridesDictionaryVisibility()
    CustomTestSetTitles Assert, TESTMODULE, "TestPlanOverridesDictionaryVisibility"

    Dim builder As IShowHideStateBuilder
    Dim plan As IShowHidePlan
    Dim state As IShowHideState

    Set plan = ShowHidePlan.Create(ShowHideLayerVList)
    plan.AddVisibility "hid_beg_v1", "Hidden variable at the begining", False

    Set builder = ShowHideStateBuilder.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1", Variables)
    builder.ApplyPlan plan

    Set state = builder.Build

    Assert.IsFalse state.EvaluateHidden("hid_beg_v1", ShowHideLayerVList), _
                   "Persisted plans should override dictionary defaults when allowed"
End Sub


'@TestMethod("ShowHide")
Public Sub TestPlanCannotHideMandatoryVariables()
    CustomTestSetTitles Assert, TESTMODULE, "TestPlanCannotHideMandatoryVariables"

    Dim builder As IShowHideStateBuilder
    Dim plan As IShowHidePlan
    Dim state As IShowHideState

    Set plan = ShowHidePlan.Create(ShowHideLayerVList)
    plan.AddVisibility "mand_v1", "Mandatory variable on vlist1D", True

    Set builder = ShowHideStateBuilder.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1", Variables)
    builder.ApplyPlan plan

    Set state = builder.Build

    Assert.IsFalse state.EvaluateHidden("mand_v1", ShowHideLayerVList), _
                   "Mandatory variables must remain visible despite persisted plan requests"
End Sub
