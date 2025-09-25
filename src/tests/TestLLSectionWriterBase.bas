Attribute VB_Name = "TestLLSectionWriterBase"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object

Private Const WRITER_DATA_SHEET As String = "LLSectionWriterData"
Private Const WRITER_SHEET As String = "WriterSheet"

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    TestHelpers.DeleteWorksheet WRITER_DATA_SHEET
    TestHelpers.DeleteWorksheet WRITER_SHEET
End Sub

Private Function BuildDictionaryStub() As LLSectionDictionaryStub
    Dim sheet As Worksheet
    Dim headers As Variant
    Dim rows As Variant
    Dim stub As LLSectionDictionaryStub

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_DATA_SHEET)
    sheet.Cells.Clear

    headers = Array("main section", "sub section", "sheet name", "variable name", "column index", "crf index")
    rows = Array( _
        Array("Section Writer", "Sub Writer", WRITER_SHEET, "var_w1", 4, 14), _
        Array("Section Writer", "Sub Writer", WRITER_SHEET, "var_w2", 6, 14))

    TestHelpers.WriteRow sheet.Cells(1, 1), headers
    TestHelpers.WriteMatrix sheet.Cells(2, 1), TestHelpers.RowsToMatrix(rows)

    Set stub = New LLSectionDictionaryStub
    stub.Configure sheet, 2, 1, 3, 6

    Set BuildDictionaryStub = stub
End Function

Private Function BuildContext(ByVal startRow As Long) As ILLSectionContext
    Dim dictStub As LLSectionDictionaryStub
    Dim llStub As LLVarContextLinelistStub
    Dim specsStub As LLVarContextSpecsStub
    Dim context As ILLSectionContext

    Set dictStub = BuildDictionaryStub()

    Set llStub = New LLVarContextLinelistStub
    llStub.UseWorkbook ThisWorkbook

    Set specsStub = New LLVarContextSpecsStub
    llStub.UseSpecs specsStub
    llStub.UseDictionary dictStub

    TestHelpers.EnsureWorksheet WRITER_SHEET

    Set context = New LLSectionContext
    context.Initialise llStub, startRow

    Set BuildContext = context
End Function

'@TestMethod("LLSectionWriter")
Private Sub TestWriteSectionInvokesHooks()
    Dim context As ILLSectionContext
    Dim layout As ILLSectionLayout
    Dim hooks As LLSectionWriterHooksStub
    Dim writer As ILLSectionWriter

    Set context = BuildContext(2)
    Set layout = New LLSectionLayoutHorizontal
    Set hooks = New LLSectionWriterHooksStub
    Set writer = New LLSectionWriterBase

    writer.Initialise context, layout, hooks
    writer.WriteSection 2, 3

    Assert.AreEqual 1, hooks.HeaderCallCount, _
                     "Section writer should invoke the header hook exactly once"
    Assert.AreEqual 1, hooks.SubSectionCallCount, _
                     "Section writer should invoke subsection formatting at least once"
    Assert.AreEqual 1, hooks.SubSectionLayoutRequests, _
                     "Hooks should be asked to provide subsection layouts"
    Assert.AreEqual "$D$5:$F$5", hooks.LastHeaderAddress, _
                     "Layout header range should be provided to the hooks"
    Assert.AreEqual 2, hooks.VariableWriteCount, _
                     "Variable writers should be requested for each variable in the subsection"

    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets(WRITER_SHEET)

    Assert.AreEqual "Section Writer", sheet.Cells(5, 4).Value, _
                     "Generic formatting should populate the section header"
    Assert.AreEqual "Sub Writer", sheet.Cells(6, 4).Value, _
                     "Generic formatting should populate the subsection header"
End Sub
