Attribute VB_Name = "TestAnalysisTableEnumerator"
Attribute VB_Description = "Unit tests covering AnalysisTableEnumerator navigation logic"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests covering AnalysisTableEnumerator navigation logic")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SPEC_SHEET As String = "EnumeratorSpecs"
Private Const SPEC_TABLE_NAME As String = "T_EnumeratorSpecs"

Private Assert As ICustomTest
Private SpecSheet As Worksheet
Private SpecTable As ListObject
Private Enumerator As IAnalysisTableEnumerator
Private BuilderStub As AnalysisTableEnumeratorBuilderStub
Private Linelist As ILinelistSpecs

'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisTableEnumerator"
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
    Set Enumerator = Nothing
    Set BuilderStub = Nothing
    Set Linelist = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set SpecSheet = EnsureWorksheet(SPEC_SHEET)
    ClearWorksheet SpecSheet

    Set BuilderStub = New AnalysisTableEnumeratorBuilderStub
    Set Enumerator = AnalysisTableEnumerator.Create(BuilderStub.Self)
    Set Linelist = New TableSpecsLinelistStub
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
    Set Enumerator = Nothing
    Set BuilderStub = Nothing
    Set Linelist = Nothing
End Sub

'@section Helpers
'===============================================================================
Private Sub BuildSpecificationTable(ByVal sections As Variant)
    Dim rowCount As Long
    Dim idx As Long
    Dim tableRows() As Variant
    Dim matrix As Variant
    Dim tableRange As Range

    rowCount = UBound(sections) - LBound(sections) + 1
    ReDim tableRows(0 To rowCount)
    tableRows(0) = Array("section", "table_id", "label")

    For idx = 0 To rowCount - 1
        tableRows(idx + 1) = Array( _
            CStr(sections(LBound(sections) + idx)), _
            "table_" & CStr(idx + 1), _
            "Label " & CStr(idx + 1) _
        )
    Next idx

    matrix = RowsToMatrix(tableRows)
    WriteMatrix SpecSheet.Range("A1"), matrix

    Set tableRange = SpecSheet.Range("A1").Resize(UBound(matrix, 1), UBound(matrix, 2))
    Set SpecTable = SpecSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    SpecTable.Name = SPEC_TABLE_NAME
End Sub

Private Function EnumerateSpecs() As BetterArray
    Set EnumerateSpecs = Enumerator.Enumerate(SpecTable, Linelist, TypeUnivariate)
End Function

Private Function SectionArray(ParamArray sectionValues() As Variant) As Variant
    SectionArray = sectionValues
End Function

'@section Tests
'===============================================================================
'@TestMethod("AnalysisTableEnumerator")
Public Sub TestEnumerateSkipsInvalidRows()
    CustomTestSetTitles Assert, "AnalysisTableEnumerator", "TestEnumerateSkipsInvalidRows"
    Dim results As BetterArray
    Dim item As IAnalysisTableIterationItem

    On Error GoTo Fail

    BuildSpecificationTable SectionArray("Section A", "Section A", "Section A")

    BuilderStub.SetRowValidity SpecTable.ListRows(2).Range.Row, False

    Set results = EnumerateSpecs()

    Assert.AreEqual 2&, results.Length, "Enumerator should skip invalid rows"

    Set item = results.Item(1)
    Assert.AreEqual SpecTable.ListRows(1).Range.Row, item.RowRange.Row, "First item should map to first row"
    Assert.AreEqual SpecTable.HeaderRowRange.Row, item.HeaderRange.Row, "Header range should point to header row"

    Set item = results.Item(2)
    Assert.AreEqual SpecTable.ListRows(3).Range.Row, item.RowRange.Row, "Second item should map to third row"
    Assert.AreEqual SpecTable.HeaderRowRange.Column, item.HeaderRange.Column, "Header range should preserve header start column"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestEnumerateSkipsInvalidRows"
End Sub

'@TestMethod("AnalysisTableEnumerator")
Public Sub TestEnumerateComputesSectionBoundaries()
    CustomTestSetTitles Assert, "AnalysisTableEnumerator", "TestEnumerateComputesSectionBoundaries"
    Dim results As BetterArray
    Dim firstItem As IAnalysisTableIterationItem
    Dim secondItem As IAnalysisTableIterationItem
    Dim thirdItem As IAnalysisTableIterationItem

    On Error GoTo Fail

    BuildSpecificationTable SectionArray("Section A", "Section A", "Section B")

    Set results = EnumerateSpecs()

    Assert.AreEqual 3&, results.Length, "Enumerator should yield all valid rows"

    Set firstItem = results.Item(1)
    Set secondItem = results.Item(2)
    Set thirdItem = results.Item(3)

    Assert.IsTrue firstItem.IsNewSection, "First entry should start a new section"
    Assert.IsTrue (Not firstItem.HeaderRange Is Nothing), "Header range should be captured on first entry"
    Assert.IsFalse secondItem.IsNewSection, "Second entry should remain in the same section"
    Assert.IsTrue thirdItem.IsNewSection, "Section change should be detected on third entry"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestEnumerateComputesSectionBoundaries"
End Sub

'@TestMethod("AnalysisTableEnumerator")
Public Sub TestEnumerateResolvesAdjacentSpecifications()
    CustomTestSetTitles Assert, "AnalysisTableEnumerator", "TestEnumerateResolvesAdjacentSpecifications"
    Dim results As BetterArray
    Dim secondItem As IAnalysisTableIterationItem
    Dim lastItem As IAnalysisTableIterationItem
    Dim previousStub As TableSpecsNavigatorSpecStub
    Dim nextStub As TableSpecsNavigatorSpecStub

    On Error GoTo Fail

    BuildSpecificationTable SectionArray("Section A", "Section A", "Section A")

    Set results = EnumerateSpecs()

    Set secondItem = results.Item(2)
    Set lastItem = results.Item(3)

    Assert.IsFalse secondItem.IsNewSection, "Middle entry should not start a new section"
    Assert.IsTrue Not secondItem.PreviousSpecification Is Nothing, "Second entry should expose a previous spec"
    Assert.IsTrue Not secondItem.NextSpecification Is Nothing, "Second entry should expose a next spec"

    Set previousStub = secondItem.PreviousSpecification
    Assert.AreEqual SpecTable.ListRows(1).Range.Row, previousStub.RowNumber, "Previous spec should map to first row"

    Set nextStub = secondItem.NextSpecification
    Assert.AreEqual SpecTable.ListRows(3).Range.Row, nextStub.RowNumber, "Next spec should map to third row"

    Assert.IsTrue lastItem.NextSpecification Is Nothing, "Last entry should not expose a next spec"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestEnumerateResolvesAdjacentSpecifications"
End Sub
