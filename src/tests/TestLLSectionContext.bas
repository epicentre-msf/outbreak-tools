Attribute VB_Name = "TestLLSectionContext"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object

Private Const CONTEXT_SHEET As String = "LLSectionContextSheet"

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    TestHelpers.DeleteWorksheet CONTEXT_SHEET
    Set Assert = Nothing
End Sub

Private Function BuildDictionaryStub() As LLSectionDictionaryStub
    Dim sheet As Worksheet
    Dim headers As Variant
    Dim rows As Variant

    Set sheet = TestHelpers.EnsureWorksheet(CONTEXT_SHEET)
    sheet.Cells.Clear

    headers = Array("main section", "sub section", "sheet name", "variable name", "column index", "crf index")
    rows = Array( _
        Array("Section A", "Sub A1", "Sheet1", "var_a1", 3, 12), _
        Array("Section A", "Sub A1", "Sheet1", "var_a2", 4, 12), _
        Array("Section A", "Sub A2", "Sheet1", "var_a3", 5, 16), _
        Array("Section B", "Sub B1", "Sheet2", "var_b1", 2, 10))

    TestHelpers.WriteRow sheet.Cells(1, 1), headers
    TestHelpers.WriteMatrix sheet.Cells(2, 1), TestHelpers.RowsToMatrix(rows)

    Dim stub As LLSectionDictionaryStub
    Set stub = New LLSectionDictionaryStub
    stub.Configure sheet, 2, 1, 5, 6

    Set BuildDictionaryStub = stub
End Function

Private Function BuildLinelistStub(ByVal dictStub As LLSectionDictionaryStub) As ILinelist
    Dim llStub As LLVarContextLinelistStub
    Dim specsStub As LLVarContextSpecsStub

    Set llStub = New LLVarContextLinelistStub
    llStub.UseWorkbook ThisWorkbook

    Set specsStub = New LLVarContextSpecsStub
    llStub.UseSpecs specsStub
    llStub.UseDictionary dictStub

    Set BuildLinelistStub = llStub
End Function

Private Function BuildContext(ByVal startRow As Long) As ILLSectionContext
    Dim dictStub As LLSectionDictionaryStub
    Dim llStub As ILinelist
    Dim context As ILLSectionContext

    Set dictStub = BuildDictionaryStub()
    Set llStub = BuildLinelistStub(dictStub)

    Set context = New LLSectionContext
    context.Initialise llStub, startRow

    Set BuildContext = context
End Function

'@TestMethod("LLSectionContext")
Private Sub TestEndRowUsesColumnBoundaries()
    Dim context As ILLSectionContext
    Dim endRow As Long

    Set context = BuildContext(2)

    endRow = context.EndRow(Array("main section", "sheet name"))

    Assert.AreEqual 4, endRow, "EndRow should extend through rows with identical section/sheet values"
End Sub

'@TestMethod("LLSectionContext")
Private Sub TestHasNextAndNextStartRow()
    Dim context As ILLSectionContext

    Set context = BuildContext(2)

    Assert.IsTrue context.HasNext(Array("sheet name")), _
                  "HasNext should be true when additional rows exist with same sheet"
    Assert.AreEqual 5, context.NextStartRow(Array("sheet name")), _
                     "NextStartRow should point to the first row of the following section"
End Sub

'@TestMethod("LLSectionContext")
Private Sub TestValueReadsDictionary()
    Dim context As ILLSectionContext

    Set context = BuildContext(2)

    Assert.AreEqual "var_a1", context.Value("variable name"), _
                     "Value should read the dictionary at the configured start row"
    Assert.AreEqual "var_a2", context.Value("variable name", rowOffset:=1), _
                     "Value should support row offsets"
End Sub

'@TestMethod("LLSectionContext")
Private Sub TestVariableNamesReturnsSlice()
    Dim context As ILLSectionContext
    Dim names As BetterArray

    Set context = BuildContext(2)

    Set names = context.VariableNames(2, 3)

    Assert.AreEqual 2, names.Length, "VariableNames should return the requested number of entries"
    Assert.AreEqual "var_a1", names.Item(1), "First variable should match dictionary order"
    Assert.AreEqual "var_a2", names.Item(2), "Second variable should match dictionary order"
End Sub
