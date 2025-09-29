Attribute VB_Name = "TestDictionarySectionBuilder"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests for the DictionarySectionBuilder helper")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private Builder As IListSectionBuilder
Private Context As ListBuildContextStub
Private LinelistStub As ILinelist
Private DictionarySheet As Worksheet
Private Const DICT_SHEET As String = "ListSectionDict"

'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    PrepareDictionarySheet
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet DICT_SHEET
    Set Assert = Nothing
End Sub

'@section Test lifecycle
'===============================================================================
'@TestInitialize
Private Sub TestInitialize()
    Set Builder = DictionarySectionBuilder.Create
    Set Context = New ListBuildContextStub
    Context.Configure "Sheet_A", CByte(2)
    Set LinelistStub = CreateLinelistStub
    Set Context.Linelist = LinelistStub
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Builder = Nothing
    Set Context = Nothing
    Set LinelistStub = Nothing
End Sub

'@section Helpers
'===============================================================================
Private Sub PrepareDictionarySheet()
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(DICT_SHEET)
    sh.Cells.Clear

    sh.Range("A1:D1").Value = Array("sheet name", "sheet type", "table name", "variable name")
    sh.Range("A2:D4").Value = Array( _
        Array("Sheet_A", "hlist2D", "tbl_a", "var_a"), _
        Array("Sheet_A", "hlist2D", "tbl_a", "var_b"), _
        Array("Sheet_B", "vlist1D", "tbl_b", "var_c"))

    If sh.ListObjects.Count > 0 Then sh.ListObjects(1).Delete
    sh.ListObjects.Add xlSrcRange, sh.Range("A1:D4"), , xlYes

    Set DictionarySheet = sh
End Sub

Private Function CreateLinelistStub() As ILinelist
    Dim dict As ILLdictionary
    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)

    Dim stub As LinelistDictionaryStub
    Set stub = New LinelistDictionaryStub
    stub.Initialise dict
    Set CreateLinelistStub = stub
End Function

'@section Tests
'===============================================================================
'@TestMethod("DictionarySectionBuilder")
Private Sub TestBuildSectionsReturnsMainSection()
    Dim sections As BetterArray

    Set sections = Builder.BuildSections(Context)

    Assert.AreEqual 1&, sections.Length, "Builder should return single section for default implementation"

    Dim descriptor As IListSectionDescriptor
    Set descriptor = sections.Item(sections.LowerBound)
    Assert.AreEqual "Sheet_A__main", descriptor.SectionName
    Assert.IsTrue descriptor.AnchorRow > 0
    Assert.IsTrue descriptor.AnchorColumn > 0
End Sub

'@TestMethod("DictionarySectionBuilder")
Private Sub TestBuildSectionsRaisesWhenSheetMissing()
    On Error GoTo ExpectError

    Context.Configure "MissingSheet", CByte(2)
    Set Context.Linelist = LinelistStub
    Builder.BuildSections Context
    Assert.Fail "Should raise when worksheet missing"
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ElementNotFound), Err.Number
    Err.Clear
End Sub

'@TestMethod("DictionarySectionBuilder")
Private Sub TestInvalidateResetsCache()
    Builder.Invalidate
    Assert.Pass "Invalidate should not raise"
End Sub
