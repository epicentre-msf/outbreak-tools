Attribute VB_Name = "TestAnalysisGoToStackBuilder"
Attribute VB_Description = "Tests verifying AnalysisGoToStackBuilder behaviour"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests verifying AnalysisGoToStackBuilder behaviour")

Private Assert As ICustomTest
Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisGoToStackBuilder"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

Private Function CreatePlanWithSpecs(ByVal sectionName As String, ByVal labelName As String, ByVal isNewSection As Boolean) As IAnalysisTablePolicyResult
    Dim specStub As GraphTablesSpecsStub
    Dim iterationStub As AnalysisTableIterationItemStub
    Dim policyStub As AnalysisTablePolicyResultStub

    Set specStub = New GraphTablesSpecsStub
    specStub.Configure TypeUnivariate, "table" & sectionName
    specStub.SetValue "section", sectionName
    specStub.SetValue "label", labelName

    Set iterationStub = New AnalysisTableIterationItemStub
    iterationStub.Configure specStub.Self, isNewSection

    Set policyStub = New AnalysisTablePolicyResultStub
    policyStub.Configure iterationStub, True
    Set CreatePlanWithSpecs = policyStub
End Function

Private Function BuildPlanResultStub() As IAnalysisTablePlanResult
    Dim items As BetterArray
    Dim resultA As IAnalysisTablePolicyResult
    Dim resultB As IAnalysisTablePolicyResult
    Dim emptySections As BetterArray
    Dim emptyHeaders As BetterArray

    Set items = New BetterArray
    items.LowerBound = 1

    Set resultA = CreatePlanWithSpecs("Section A", "Label 1", True)
    Set resultB = CreatePlanWithSpecs("Section B", "Label 2", False)

    items.Push AnalysisTablePlanItem.Create(resultA, 0)
    items.Push AnalysisTablePlanItem.Create(resultB, 1)

    Set emptySections = New BetterArray
    emptySections.LowerBound = 1
    Set emptyHeaders = New BetterArray
    emptyHeaders.LowerBound = 1

    Set BuildPlanResultStub = AnalysisTablePlanResult.Create(items, emptySections, emptyHeaders, "sec: ", "hdr: ")
End Function

'@TestMethod("AnalysisGoToStackBuilder")
Public Sub TestBuildAggregatesLabels()
    CustomTestSetTitles Assert, "AnalysisGoToStackBuilder", "TestBuildAggregatesLabels"
    Dim builder As AnalysisGoToStackBuilder
    Dim plan As IAnalysisTablePlanResult
    Dim stacks As AnalysisGoToStackContext
    Dim sectionEntry As IAnalysisGoToEntry
    Dim headerEntry As IAnalysisGoToEntry

    Set builder = AnalysisGoToStackBuilder.Create("sec: ", "hdr: ")
    Set plan = BuildPlanResultStub()

    Set stacks = builder.Build(plan)

    Assert.AreEqual 1&, stacks.SectionLabels.Length, "Only new sections should generate section labels"

    Set sectionEntry = stacks.SectionLabels.Item(1)
    Assert.AreEqual "section", sectionEntry.Scope, "Section entry should expose the correct scope"
    Assert.AreEqual "Section A", sectionEntry.Label, "Section entry should preserve the original label"
    Assert.AreEqual "sec: Section A", sectionEntry.DisplayText, "Section entry should include prefix in display text"
    Assert.AreEqual "tableSection A", sectionEntry.Suffix, "Section entry should expose the table identifier as suffix"

    Assert.AreEqual 2&, stacks.HeaderLabels.Length, "Every table should contribute a header label"

    Set headerEntry = stacks.HeaderLabels.Item(1)
    Assert.AreEqual "header", headerEntry.Scope, "Header entry should expose the correct scope"
    Assert.AreEqual "Label 1", headerEntry.Label, "Header entry should preserve the original label"
    Assert.AreEqual "hdr: Label 1", headerEntry.DisplayText, "Header entry should include prefix in display text"
End Sub
