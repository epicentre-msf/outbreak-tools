Attribute VB_Name = "TestListBuilderFactory"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests verifying ListBuilderFactory selects appropriate collaborators")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private DictionarySheet As Worksheet
Private Const DICT_SHEET As String = "ListBuilderFactoryDict"

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
        Array("Sheet_H", "hlist2D", "tbl_h", "var_h"), _
        Array("Sheet_V", "vlist1D", "tbl_v", "var_v"))

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
'@TestMethod("ListBuilderFactory")
Private Sub TestCreateCoordinatorSelectsHorizontalStrategy()
    Dim factory As ListBuilderFactory
    Dim linelist As ILinelist
    Dim context As IListBuildContext
    Dim coordinator As IListBuilderCoordinator

    Set factory = ListBuilderFactory.Create
    Set linelist = CreateLinelistStub
    Set context = factory.CreateContext(linelist, "Sheet_H")
    Set coordinator = factory.CreateCoordinator(context)

    Dim concrete As ListBuilderCoordinator
    Set concrete = coordinator

    Assert.AreEqual "HorizontalListLayoutStrategy", TypeName(concrete.LayoutStrategy)
End Sub

'@TestMethod("ListBuilderFactory")
Private Sub TestCreateCoordinatorSelectsVerticalStrategy()
    Dim factory As ListBuilderFactory
    Dim linelist As ILinelist
    Dim context As IListBuildContext
    Dim coordinator As IListBuilderCoordinator

    Set factory = ListBuilderFactory.Create
    Set linelist = CreateLinelistStub
    Set context = factory.CreateContext(linelist, "Sheet_V")
    Set coordinator = factory.CreateCoordinator(context)

    Dim concrete As ListBuilderCoordinator
    Set concrete = coordinator

    Assert.AreEqual "VerticalListLayoutStrategy", TypeName(concrete.LayoutStrategy)
End Sub

