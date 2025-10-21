Attribute VB_Name = "TestAnalysisTableEngine"
Attribute VB_Description = "Unit tests for AnalysisTableEngine coordinator"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests for AnalysisTableEngine coordinator")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SPEC_SHEET As String = "EngineSpecs"
Private Const SPEC_TABLE_NAME As String = "T_EngineSpecs"

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
    Assert.SetModuleName "TestAnalysisTableEngine"
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

    Set SpecSheet = Nothing
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

Private Function CreatePolicyResultStub(ByVal tableId As String, ByVal sectionName As String) As IAnalysisTablePolicyResult
    Dim specStub As GraphTablesSpecsStub
    Dim iterationStub As AnalysisTableIterationItemStub
    Dim policyStub As AnalysisTablePolicyResultStub

    Set specStub = New GraphTablesSpecsStub
    specStub.Configure TypeUnivariate, tableId
    specStub.SetValue "section", sectionName
    specStub.SetValue "label", "Label " & Right$(tableId, 1)

    Set iterationStub = New AnalysisTableIterationItemStub
    iterationStub.Configure specStub.Self, True

    Set policyStub = New AnalysisTablePolicyResultStub
    policyStub.Configure iterationStub, True
    policyStub.SetFlags True, True, False, False

    Set CreatePolicyResultStub = policyStub
End Function

Private Function BuildPlanResult() As IAnalysisTablePlanResult
    Dim policyResults As BetterArray
    Dim items As BetterArray

    Set policyResults = New BetterArray
    policyResults.LowerBound = 1
    policyResults.Push CreatePolicyResultStub("table_1", "Section A")
    policyResults.Push CreatePolicyResultStub("table_2", "Section B")

    Set items = New BetterArray
    items.LowerBound = 1
    items.Push AnalysisTablePlanItem.Create(policyResults.Item(1), 0)
    items.Push AnalysisTablePlanItem.Create(policyResults.Item(2), 1)

    BuildPlanResult = AnalysisTablePlanResult.Create(items, _
                                                     RowsToBetterArray(Array("sec: Section A", "sec: Section B")), _
                                                     RowsToBetterArray(Array("hdr: Label 1", "hdr: Label 2")), _
                                                     "sec: ", "hdr: ")
End Function

Private Function RowsToBetterArray(ByVal values As Variant) As BetterArray
    Dim buffer As BetterArray
    Dim idx As Long

    Set buffer = New BetterArray
    buffer.LowerBound = 1

    For idx = LBound(values) To UBound(values)
        buffer.Push values(idx)
    Next idx

    Set RowsToBetterArray = buffer
End Function

'@section Tests
'===============================================================================
'@TestMethod("AnalysisTableEngine")
Public Sub TestRunInvokesPipeline()
    CustomTestSetTitles Assert, "AnalysisTableEngine", "TestRunInvokesPipeline"
    Dim sequenceStub As AnalysisTableSequenceBuilderStub
    Dim planStub As AnalysisTablePlanBuilderStub
    Dim writerStub As AnalysisTableWriterStub
    Dim engine As IAnalysisTableEngine
    Dim planResult As IAnalysisTablePlanResult
    Dim sequenceResults As BetterArray
    Dim linelistStub As TableSpecsLinelistStub
    Dim contextStub As AnalysisTableWriterContextStub

    On Error GoTo Fail

    Set sequenceStub = New AnalysisTableSequenceBuilderStub
    Set sequenceResults = New BetterArray
    sequenceResults.LowerBound = 1
    sequenceResults.Push CreatePolicyResultStub("table_1", "Section A")
    sequenceStub.SetResults sequenceResults

    Set planStub = New AnalysisTablePlanBuilderStub
    planResult = BuildPlanResult()
    planStub.SetPlanResult planResult

    Set writerStub = New AnalysisTableWriterStub
    Set engine = AnalysisTableEngine.Create(sequenceStub, planStub, writerStub)
    Set linelistStub = New TableSpecsLinelistStub
    Set contextStub = New AnalysisTableWriterContextStub

    engine.Run SpecTable, SpecSheet, linelistStub, TypeUnivariate, contextStub, "sec: ", "hdr: "

    Assert.AreEqual 1&, sequenceStub.BuildCount, "Sequence builder should be invoked"
    Assert.AreSameObj SpecTable, sequenceStub.LastSpecificationList, "Sequence builder should receive specification list"
    Assert.AreSameObj linelistStub, sequenceStub.LastLinelistSpecs, "Sequence builder should receive linelist specs"

    Assert.AreEqual 1&, planStub.BuildCount, "Plan builder should be invoked"
    Assert.AreSameObj sequenceResults, planStub.LastPolicyResults, "Plan builder should receive sequence results"
    Assert.AreEqual "sec: ", planStub.LastSectionPrefix, "Section prefix should propagate"
    Assert.AreEqual "hdr: ", planStub.LastHeaderPrefix, "Header prefix should propagate"

    Assert.AreEqual 1&, writerStub.WriteCount, "Writer should be invoked"
    Assert.AreSameObj planResult, writerStub.LastPlan, "Writer should receive plan result"
    Assert.AreSameObj contextStub, writerStub.LastContext, "Writer should receive context"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestRunInvokesPipeline"
End Sub

'@TestMethod("AnalysisTableEngine")
Public Sub TestRunValidatesContext()
    CustomTestSetTitles Assert, "AnalysisTableEngine", "TestRunValidatesContext"
    Dim engine As IAnalysisTableEngine
    Dim sequenceStub As AnalysisTableSequenceBuilderStub
    Dim planStub As AnalysisTablePlanBuilderStub
    Dim writerStub As AnalysisTableWriterStub
    Dim raisedError As Boolean

    Set sequenceStub = New AnalysisTableSequenceBuilderStub
    Set planStub = New AnalysisTablePlanBuilderStub
    planStub.SetPlanResult BuildPlanResult()
    Set writerStub = New AnalysisTableWriterStub
    Set engine = AnalysisTableEngine.Create(sequenceStub, planStub, writerStub)

    On Error Resume Next
        engine.Run SpecTable, SpecSheet, New TableSpecsLinelistStub, TypeUnivariate, Nothing
        raisedError = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Engine should validate writer context argument"
End Sub
