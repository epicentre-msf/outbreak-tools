Attribute VB_Name = "TestAnalysisTablePlanBuilder"
Attribute VB_Description = "Unit tests for AnalysisTablePlanBuilder navigation planning"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests for AnalysisTablePlanBuilder navigation planning")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SPEC_SHEET As String = "PlanBuilderSpecs"
Private Const SPEC_TABLE_NAME As String = "T_PlanBuilderSpecs"

Private Assert As ICustomTest
Private SpecSheet As Worksheet
Private SpecTable As ListObject

'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisTablePlanBuilder"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    DeleteWorksheet SPEC_SHEET
    RestoreApp

    Set Assert = Nothing
    Set SpecSheet = Nothing
    Set SpecTable = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set SpecSheet = EnsureWorksheet(SPEC_SHEET)
    ClearWorksheet SpecSheet
    BuildSpecificationTable
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    If Not SpecSheet Is Nothing Then
        ClearWorksheet SpecSheet
    End If

    Set SpecTable = Nothing
End Sub

'@section Helpers
'===============================================================================
Private Sub BuildSpecificationTable()
    Dim data As Variant
    Dim tableRange As Range

    data = Array( _
        Array("section", "table_id", "label"), _
        Array("Section A", "table_1", "Label 1"), _
        Array("Section B", "table_2", "Label 2") _
    )

    WriteMatrix SpecSheet.Range("A1"), RowsToMatrix(data)
    Set tableRange = SpecSheet.Range("A1").Resize(UBound(data) + 1, UBound(data, 2) + 1)
    Set SpecTable = SpecSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                              XlListObjectHasHeaders:=xlYes)
    SpecTable.Name = SPEC_TABLE_NAME
End Sub

Private Function CreateIterationItem(ByVal listRowIndex As Long, ByVal isNewSection As Boolean) As IAnalysisTableIterationItem
    Dim specStub As GraphTablesSpecsStub
    Dim listRow As ListRow

    Set specStub = New GraphTablesSpecsStub
    specStub.Configure TypeUnivariate, "table_" & CStr(listRowIndex), "section_" & CStr(listRowIndex)
    specStub.SetIsNewSection isNewSection
    specStub.SetValue "section", SpecTable.ListRows(listRowIndex).Range.Cells(1, 1).Value
    specStub.SetValue "label", SpecTable.ListRows(listRowIndex).Range.Cells(1, 3).Value

    Set listRow = SpecTable.ListRows(listRowIndex)

    Set CreateIterationItem = AnalysisTableIterationItem.Create(specStub.Self, _
                                                                SpecTable.HeaderRowRange, _
                                                                listRow.Range, _
                                                                listRow, _
                                                                isNewSection)
End Function

Private Function BuildPolicyResult(ByVal iterationItem As IAnalysisTableIterationItem) As IAnalysisTablePolicyResult
    Dim policyStub As AnalysisTypePolicyStub
    Dim contextStub As AnalysisPolicyContextStub

    Set policyStub = New AnalysisTypePolicyStub
    policyStub.Configure TypeUnivariate, True, True, True, False, False

    Set contextStub = New AnalysisPolicyContextStub
    contextStub.Configure TypeUnivariate

    Set BuildPolicyResult = AnalysisTablePolicyResult.Create(iterationItem, policyStub, contextStub, True, True, True, False, False)
End Function

Private Function PolicyResultsBuffer(ParamArray items() As Variant) As BetterArray
    Dim buffer As BetterArray
    Dim idx As Long

    Set buffer = New BetterArray
    buffer.LowerBound = 1

    For idx = LBound(items) To UBound(items)
        buffer.Push items(idx)
    Next idx

    Set PolicyResultsBuffer = buffer
End Function

'@section Tests
'===============================================================================
'@TestMethod("AnalysisTablePlanBuilder")
Public Sub TestBuildCreatesPlanItems()
    CustomTestSetTitles Assert, "AnalysisTablePlanBuilder", "TestBuildCreatesPlanItems"
    Dim iterationItem As IAnalysisTableIterationItem
    Dim policyResult As IAnalysisTablePolicyResult
    Dim builder As AnalysisTablePlanBuilder
    Dim plan As IAnalysisTablePlanResult

    On Error GoTo Fail

    Set iterationItem = CreateIterationItem(1, True)
    Set policyResult = BuildPolicyResult(iterationItem)

    Set builder = New AnalysisTablePlanBuilder
    Set plan = builder.Build(PolicyResultsBuffer(policyResult))

    Assert.AreEqual 1&, plan.Items.Length, "Plan should contain single item"
    Assert.AreSameObj policyResult, plan.Items.Item(1).PolicyResult, "Plan item should reference policy result"
    Assert.AreEqual vbNullString, plan.SectionPrefix, "Default section prefix should be empty"
    Assert.AreEqual vbNullString, plan.HeaderPrefix, "Default header prefix should be empty"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestBuildCreatesPlanItems"
End Sub

'@TestMethod("AnalysisTablePlanBuilder")
Public Sub TestBuildCollectsSectionLabels()
    CustomTestSetTitles Assert, "AnalysisTablePlanBuilder", "TestBuildCollectsSectionLabels"
    Dim iterationItem As IAnalysisTableIterationItem
    Dim policyResult As IAnalysisTablePolicyResult
    Dim builder As AnalysisTablePlanBuilder
    Dim plan As IAnalysisTablePlanResult

    On Error GoTo Fail

    Set iterationItem = CreateIterationItem(1, True)
    Set policyResult = BuildPolicyResult(iterationItem)

    Set builder = New AnalysisTablePlanBuilder
    Set plan = builder.Build(PolicyResultsBuffer(policyResult), "sec: ")

    Assert.AreEqual 1&, plan.SectionLabels.Length, "One section label expected"
    Assert.AreEqual "sec: Section A", CStr(plan.SectionLabels.Item(1)), "Section label should include prefix"
    Assert.AreEqual "sec: ", plan.SectionPrefix, "Section prefix should be stored"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestBuildCollectsSectionLabels"
End Sub

'@TestMethod("AnalysisTablePlanBuilder")
Public Sub TestBuildCollectsHeaderLabels()
    CustomTestSetTitles Assert, "AnalysisTablePlanBuilder", "TestBuildCollectsHeaderLabels"
    Dim iterationItem As IAnalysisTableIterationItem
    Dim policyResult As IAnalysisTablePolicyResult
    Dim builder As AnalysisTablePlanBuilder
    Dim plan As IAnalysisTablePlanResult

    On Error GoTo Fail

    Set iterationItem = CreateIterationItem(2, False)
    Set policyResult = BuildPolicyResult(iterationItem)

    Set builder = New AnalysisTablePlanBuilder
    Set plan = builder.Build(PolicyResultsBuffer(policyResult), vbNullString, "hdr: ")

    Assert.AreEqual 1&, plan.HeaderLabels.Length, "One header label expected"
    Assert.AreEqual "hdr: Label 2", CStr(plan.HeaderLabels.Item(1)), "Header label should include prefix"
    Assert.AreEqual "hdr: ", plan.HeaderPrefix, "Header prefix should be stored"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestBuildCollectsHeaderLabels"
End Sub
