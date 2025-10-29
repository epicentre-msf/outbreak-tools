Attribute VB_Name = "TestAnalysisTableSequenceBuilder"
Attribute VB_Description = "Unit tests for AnalysisTableSequenceBuilder orchestration"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests for AnalysisTableSequenceBuilder orchestration")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SPEC_SHEET As String = "SequenceSpecs"
Private Const SPEC_TABLE_NAME As String = "T_SequenceSpecs"

Private Assert As ICustomTest
Private SpecSheet As Worksheet
Private SpecTable As ListObject
Private EnumeratorStub As AnalysisTableEnumeratorStub
Private PolicyResolverStub As AnalysisTablePolicyResolverStub
Private ItemsBuffer As BetterArray

'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisTableSequenceBuilder"
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
    Set EnumeratorStub = Nothing
    Set PolicyResolverStub = Nothing
    Set ItemsBuffer = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set SpecSheet = EnsureWorksheet(SPEC_SHEET)
    ClearWorksheet SpecSheet

    BuildSpecificationTable

    Set EnumeratorStub = New AnalysisTableEnumeratorStub
    Set PolicyResolverStub = New AnalysisTablePolicyResolverStub

    Set ItemsBuffer = New BetterArray
    ItemsBuffer.LowerBound = 1
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    If Not SpecSheet Is Nothing Then
        ClearWorksheet SpecSheet
    End If

    Set ItemsBuffer = Nothing
    Set SpecTable = Nothing
End Sub

'@section Helpers
'===============================================================================
Private Sub BuildSpecificationTable()
    Dim data As Variant
    Dim tableRange As Range

    data = Array( _
        Array("section", "table_id", "row"), _
        Array("Section A", "table_1", "age"), _
        Array("Section A", "table_2", "age") _
    )

    WriteMatrix SpecSheet.Range("A1"), RowsToMatrix(data)
    Set tableRange = SpecSheet.Range("A1").Resize(UBound(data) + 1, UBound(data, 2) + 1)
    Set SpecTable = SpecSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                              XlListObjectHasHeaders:=xlYes)
    SpecTable.Name = SPEC_TABLE_NAME
End Sub

Private Function CreateIterationItem(ByVal spec As ITablesSpecs, ByVal listRowIndex As Long) As IAnalysisTableIterationItem
    Dim headerRange As Range
    Dim rowRange As Range
    Dim listRow As ListRow

    Set headerRange = SpecTable.HeaderRowRange
    Set listRow = SpecTable.ListRows(listRowIndex)
    Set rowRange = listRow.Range

    Set CreateIterationItem = AnalysisTableIterationItem.Create(spec, headerRange, rowRange, listRow, spec.IsNewSection)
End Function

Private Function BuildSpecStub(ByVal tableId As String, ByVal isNewSection As Boolean) As ITablesSpecs
    Dim stub As GraphTablesSpecsStub

    Set stub = New GraphTablesSpecsStub
    stub.Configure TypeUnivariate, tableId
    stub.SetIsNewSection isNewSection
    stub.SetValue "section", "Section A"
    Set BuildSpecStub = stub.Self
End Function

Private Function BuildPolicyResult(ByVal iterationItem As IAnalysisTableIterationItem, _
                                   ByVal isValid As Boolean) As IAnalysisTablePolicyResult
    Dim policyStub As AnalysisTypePolicyStub
    Dim contextStub As AnalysisPolicyContextStub

    Set policyStub = New AnalysisTypePolicyStub
    policyStub.Configure TypeUnivariate, isValid, True, True, False, False

    Set contextStub = New AnalysisPolicyContextStub
    contextStub.Configure TypeUnivariate

    Set BuildPolicyResult = AnalysisTablePolicyResult.Create(iterationItem, policyStub, contextStub, _
                                                            isValid, True, True, False, False)
End Function

Private Function BuildSequenceBuilder() As IAnalysisTableSequenceBuilder
    EnumeratorStub.SetItems ItemsBuffer
    Set BuildSequenceBuilder = AnalysisTableSequenceBuilder.Create(New TableSpecsLinelistStub, EnumeratorStub, PolicyResolverStub)
End Function

'@section Tests
'===============================================================================
'@TestMethod("AnalysisTableSequenceBuilder")
Public Sub TestBuildFiltersInvalidPolicies()
    CustomTestSetTitles Assert, "AnalysisTableSequenceBuilder", "TestBuildFiltersInvalidPolicies"
    Dim itemValid As IAnalysisTableIterationItem
    Dim itemInvalid As IAnalysisTableIterationItem
    Dim builder As IAnalysisTableSequenceBuilder
    Dim results As BetterArray

    On Error GoTo Fail

    Set itemValid = CreateIterationItem(BuildSpecStub("table_1", True), 1)
    Set itemInvalid = CreateIterationItem(BuildSpecStub("table_2", False), 2)

    ItemsBuffer.Push itemValid
    ItemsBuffer.Push itemInvalid

    PolicyResolverStub.EnqueueResult BuildPolicyResult(itemValid, True)
    PolicyResolverStub.EnqueueResult BuildPolicyResult(itemInvalid, False)

    Set builder = BuildSequenceBuilder()
    Set results = builder.Build(SpecTable, New TableSpecsLinelistStub, TypeUnivariate)

    Assert.AreEqual 1&, results.Length, "Builder should filter out invalid policy results"
    Assert.IsTrue results.Item(1).IsValid, "Remaining result should be valid"
    Assert.AreSameObj itemValid, results.Item(1).IterationItem, "Result should reference valid iteration item"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestBuildFiltersInvalidPolicies"
End Sub

'@TestMethod("AnalysisTableSequenceBuilder")
Public Sub TestBuildInvokesEnumerator()
    CustomTestSetTitles Assert, "AnalysisTableSequenceBuilder", "TestBuildInvokesEnumerator"
    Dim itemValid As IAnalysisTableIterationItem
    Dim builder As IAnalysisTableSequenceBuilder

    On Error GoTo Fail

    Set itemValid = CreateIterationItem(BuildSpecStub("table_1", True), 1)
    ItemsBuffer.Push itemValid
    PolicyResolverStub.EnqueueResult BuildPolicyResult(itemValid, True)

    Set builder = BuildSequenceBuilder()
    builder.Build SpecTable, New TableSpecsLinelistStub, TypeUnivariate

    Assert.AreEqual 1&, EnumeratorStub.EnumerateCount, "Enumerator should be invoked once"
    Assert.AreSameObj SpecTable, EnumeratorStub.LastSpecificationList, "Enumerator should receive the specification list"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestBuildInvokesEnumerator"
End Sub
