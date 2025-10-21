Attribute VB_Name = "TestAnalysisTableWriter"
Attribute VB_Description = "Unit tests for AnalysisTableWriter orchestration"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests for AnalysisTableWriter orchestration")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SPEC_SHEET As String = "WriterSpecs"
Private Const SPEC_TABLE_NAME As String = "T_WriterSpecs"

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
    Assert.SetModuleName "TestAnalysisTableWriter"
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

Private Function BuildPlan(ByVal sectionLabelPrefix As String, ByVal headerLabelPrefix As String) As IAnalysisTablePlanResult
    Dim planBuilder As AnalysisTablePlanBuilder
    Dim policyResults As BetterArray

    Set planBuilder = New AnalysisTablePlanBuilder

    Set policyResults = New BetterArray
    policyResults.LowerBound = 1

    policyResults.Push BuildPolicyResult(CreateIterationItem(1, True))
    policyResults.Push BuildPolicyResult(CreateIterationItem(2, True))

    Set BuildPlan = planBuilder.Build(policyResults, sectionLabelPrefix, headerLabelPrefix)
End Function

'@section Tests
'===============================================================================
'@TestMethod("AnalysisTableWriter")
Public Sub TestWriteInvokesContextForEachItem()
    CustomTestSetTitles Assert, "AnalysisTableWriter", "TestWriteInvokesContextForEachItem"
    Dim writer As IAnalysisTableWriter
    Dim contextStub As AnalysisTableWriterContextStub
    Dim plan As IAnalysisTablePlanResult

    On Error GoTo Fail

    Set writer = AnalysisTableWriter.Create
    Set contextStub = New AnalysisTableWriterContextStub
    Set plan = BuildPlan("sec: ", "hdr: ")

    writer.Write plan, SpecSheet, contextStub

    Assert.AreEqual 1&, contextStub.BeginPlanCount, "BeginPlan should be called once"
    Assert.AreEqual plan.Items.Length, contextStub.WrittenItems.Length, "WriteTable should be called for each plan item"
    Assert.AreEqual 1&, contextStub.ApplyNavigationCount, "Navigation should be applied once"
    Assert.AreEqual 1&, contextStub.CompletePlanCount, "CompletePlan should be called once"
    Assert.AreEqual plan.SectionLabels.Length, contextStub.LastSectionLabels.Length, "Section labels should be forwarded"
    Assert.AreEqual plan.HeaderLabels.Length, contextStub.LastHeaderLabels.Length, "Header labels should be forwarded"
    Assert.AreEqual "sec: Section A", CStr(contextStub.LastSectionLabels.Item(1)), "Section label content should include prefix"
    Assert.AreEqual "hdr: Label 1", CStr(contextStub.LastHeaderLabels.Item(1)), "Header label content should include prefix"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestWriteInvokesContextForEachItem"
End Sub

'@TestMethod("AnalysisTableWriter")
Public Sub TestWriteValidatesInputs()
    CustomTestSetTitles Assert, "AnalysisTableWriter", "TestWriteValidatesInputs"
    Dim writer As IAnalysisTableWriter
    Dim contextStub As AnalysisTableWriterContextStub
    Dim raisedError As Boolean

    Set writer = AnalysisTableWriter.Create
    Set contextStub = New AnalysisTableWriterContextStub

    On Error Resume Next
        writer.Write Nothing, SpecSheet, contextStub
        raisedError = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Writer should validate the plan argument"
End Sub
