Attribute VB_Name = "TestListBuildService"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests for the ListBuildService high-level orchestrator")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private DictionarySheet As Worksheet
Private Const DICT_SHEET As String = "ListBuildServiceDict"

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

'@section Helpers
'===============================================================================
Private Sub PrepareDictionarySheet()
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(DICT_SHEET)
    sh.Cells.Clear

    sh.Range("A1:D1").Value = Array("sheet name", "sheet type", "table name", "variable name")
    sh.Range("A2:D3").Value = Array( _
        Array("Sheet_Main", "hlist2D", "tbl_main", "var_main"), _
        Array("Sheet_Main", "hlist2D", "tbl_main", "var_main2"))

    If sh.ListObjects.Count > 0 Then sh.ListObjects(1).Delete
    sh.ListObjects.Add xlSrcRange, sh.Range("A1:D3"), , xlYes

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
'@TestMethod("ListBuildService")
Private Sub TestBuildReturnsSections()
    Dim service As ListBuildService
    Dim linelist As ILinelist
    Dim sections As BetterArray

    Set service = ListBuildService.Create
    Set linelist = CreateLinelistStub

    Set sections = service.Build(linelist, "Sheet_Main")

    Assert.AreEqual 1&, sections.Length
    Assert.AreEqual "Sheet_Main__main", sections.Item(sections.LowerBound).SectionName
End Sub

