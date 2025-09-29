Attribute VB_Name = "TestLLSectionLayout"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object

Private Const LAYOUT_DATA_SHEET As String = "LLSectionLayoutData"
Private Const SHEET_HORIZONTAL As String = "LayoutH"
Private Const SHEET_VERTICAL As String = "LayoutV"

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    TestHelpers.DeleteWorksheet LAYOUT_DATA_SHEET
    TestHelpers.DeleteWorksheet SHEET_HORIZONTAL
    TestHelpers.DeleteWorksheet SHEET_VERTICAL
End Sub

Private Function BuildDictionaryStub() As LLSectionDictionaryStub
    Dim sheet As Worksheet
    Dim headers As Variant
    Dim rows As Variant
    Dim stub As LLSectionDictionaryStub

    Set sheet = TestHelpers.EnsureWorksheet(LAYOUT_DATA_SHEET)
    sheet.Cells.Clear

    headers = Array("main section", "sub section", "sheet name", "variable name", "column index", "crf index")
    rows = Array( _
        Array("Section H", "Sub H1", SHEET_HORIZONTAL, "var_h1", 3, 12), _
        Array("Section H", "Sub H1", SHEET_HORIZONTAL, "var_h2", 5, 12), _
        Array("Section V", "Sub V1", SHEET_VERTICAL, "var_v1", 10, 0), _
        Array("Section V", "Sub V1", SHEET_VERTICAL, "var_v2", 12, 0))

    TestHelpers.WriteRow sheet.Cells(1, 1), headers
    TestHelpers.WriteMatrix sheet.Cells(2, 1), TestHelpers.RowsToMatrix(rows)

    Set stub = New LLSectionDictionaryStub
    stub.Configure sheet, 2, 1, 5, 6

    Set BuildDictionaryStub = stub
End Function

Private Function BuildLinelistStub(ByVal dictStub As ILLdictionary) As ILinelist
    Dim llStub As LLVarContextLinelistStub
    Dim specsStub As LLVarContextSpecsStub

    Set llStub = New LLVarContextLinelistStub
    llStub.UseWorkbook ThisWorkbook

    Set specsStub = New LLVarContextSpecsStub
    llStub.UseSpecs specsStub
    llStub.UseDictionary dictStub

    Set BuildLinelistStub = llStub
End Function

Private Function BuildSectionContext(ByVal startRow As Long) As ILLSectionContext
    Dim dictStub As LLSectionDictionaryStub
    Dim llStub As ILinelist
    Dim context As ILLSectionContext

    Set dictStub = BuildDictionaryStub()
    Set llStub = BuildLinelistStub(dictStub)

    TestHelpers.EnsureWorksheet SHEET_HORIZONTAL
    TestHelpers.EnsureWorksheet SHEET_VERTICAL

    Set context = New LLSectionContext
    context.Initialise llStub, startRow

    Set BuildSectionContext = context
End Function

'@TestMethod("LLSectionLayout")
Private Sub TestHorizontalLayoutRanges()
    Dim context As ILLSectionContext
    Dim layout As ILLSectionLayout

    Set context = BuildSectionContext(2)
    Set layout = New LLSectionLayoutHorizontal

    layout.Configure context, 2, 3

    Assert.AreEqual "$C$5:$E$5", layout.HeaderRange.Address(False, False), _
                     "Horizontal header range should span the first and last variable columns"

    Assert.AreEqual "$C$5:$E$5", layout.PrintHeaderRange.Address(False, False), _
                     "Printed header should mirror horizontal layout"

    Assert.AreEqual "$A$8:$G$8", layout.CrfHeaderRange.Address(False, False), _
                     "CRF header should be positioned relative to the CRF index"
End Sub

'@TestMethod("LLSectionLayout")
Private Sub TestVerticalLayoutRanges()
    Dim context As ILLSectionContext
    Dim layout As ILLSectionLayout

    Set context = BuildSectionContext(4)
    Set layout = New LLSectionLayoutVertical

    layout.Configure context, 4, 5

    Assert.AreEqual "$B$10:$B$12", layout.HeaderRange.Address(False, False), _
                     "Vertical header range should span rows derived from the column index"

    Assert.IsTrue layout.PrintHeaderRange Is Nothing, _
                  "Vertical layout does not expose a printed range"

    Assert.IsTrue layout.CrfHeaderRange Is Nothing, _
                  "Vertical layout does not expose CRF ranges"
End Sub

'@TestMethod("LLSectionLayout")
Private Sub TestHorizontalSubSectionLayout()
    Dim context As ILLSectionContext
    Dim layout As ILLSubSectionLayout

    Set context = BuildSectionContext(2)
    Set layout = New LLSubSectionLayoutHorizontal

    layout.Configure context, 2, 3

    Assert.AreEqual "$C$6:$E$6", layout.HeaderRange.Address(False, False), _
                     "Horizontal subsection header should align with the subsection row"

    Assert.AreEqual "$C$6:$E$6", layout.PrintHeaderRange.Address(False, False), _
                     "Printed subsection header should mirror the base layout"

    Assert.AreEqual "$A$9:$G$9", layout.CrfHeaderRange.Address(False, False), _
                     "CRF subsection header should account for the CRF offset"
End Sub

'@TestMethod("LLSectionLayout")
Private Sub TestVerticalSubSectionLayout()
    Dim context As ILLSectionContext
    Dim layout As ILLSubSectionLayout

    Set context = BuildSectionContext(4)
    Set layout = New LLSubSectionLayoutVertical

    layout.Configure context, 4, 5

    Assert.AreEqual "$C$10:$C$12", layout.HeaderRange.Address(False, False), _
                     "Vertical subsection header should span the derived row range"

    Assert.IsTrue layout.PrintHeaderRange Is Nothing, _
                  "Vertical subsection layout does not expose a printed range"

    Assert.IsTrue layout.CrfHeaderRange Is Nothing, _
                  "Vertical subsection layout does not expose CRF ranges"
End Sub

