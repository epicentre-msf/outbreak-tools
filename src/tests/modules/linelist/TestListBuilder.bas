Attribute VB_Name = "TestListBuilder"
Attribute VB_Description = "Tests for ListBuilder class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for ListBuilder class")

Option Explicit

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private Dict As ILLdictionary
Private Specs As LinelistSpecsWorkbookStub
Private FakeLL As LinelistStub

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ListBuilder"
Private Const DICTIONARY_SHEET As String = "DictFixture"


'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestListBuilder"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, FixtureWorkbook
    Set Dict = LLdictionary.Create(FixtureWorkbook.Worksheets(DICTIONARY_SHEET), 1, 1)
    Dict.Prepare

    Set Specs = New LinelistSpecsWorkbookStub
    Specs.Initialise Dict, FixtureWorkbook

    Set FakeLL = New LinelistStub
    FakeLL.Initialise Specs, Dict

    Assert.BeginTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then TestHelpers.DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set Dict = Nothing
    Set Specs = Nothing
    Set FakeLL = Nothing
    Set FixtureWorkbook = Nothing
End Sub


'@section Factory tests
'===============================================================================

'@TestMethod("ListBuilder")
Public Sub TestCreateHListReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateHListReturnsInstance"
    On Error GoTo TestFail

    Dim sheetName As String
    Dim sheetsList As BetterArray

    Set sheetsList = Dict.UniqueValues("sheet name")
    If sheetsList.Length = 0 Then
        Assert.IsTrue True, "No sheets in fixture — skipping"
        Exit Sub
    End If

    sheetName = sheetsList.Item(sheetsList.LowerBound)
    Dim sut As IListBuilder
    Set sut = ListBuilder.Create(ListBuilderLayerHList, sheetName, FakeLL)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance for HList layer"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateHListReturnsInstance", Err.Number, Err.Description
End Sub

'@TestMethod("ListBuilder")
Public Sub TestCreateVListReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateVListReturnsInstance"
    On Error GoTo TestFail

    Dim sheetName As String
    Dim sheetsList As BetterArray

    Set sheetsList = Dict.UniqueValues("sheet name")
    If sheetsList.Length = 0 Then
        Assert.IsTrue True, "No sheets in fixture — skipping"
        Exit Sub
    End If

    sheetName = sheetsList.Item(sheetsList.LowerBound)
    Dim sut As IListBuilder
    Set sut = ListBuilder.Create(ListBuilderLayerVList, sheetName, FakeLL)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance for VList layer"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateVListReturnsInstance", Err.Number, Err.Description
End Sub

'@TestMethod("ListBuilder")
Public Sub TestCreateRejectsNothingLinelist()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingLinelist"
    On Error GoTo ExpectError

    Dim sut As IListBuilder
    Set sut = ListBuilder.Create(ListBuilderLayerHList, "SomeSheet", Nothing)

    CustomTestLogFailure Assert, "TestCreateRejectsNothingLinelist", , _
                         "Expected error when linelist is Nothing"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when linelist is Nothing"
End Sub

'@TestMethod("ListBuilder")
Public Sub TestCreateRejectsEmptySheetName()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsEmptySheetName"
    On Error GoTo ExpectError

    Dim sut As IListBuilder
    Set sut = ListBuilder.Create(ListBuilderLayerHList, vbNullString, FakeLL)

    CustomTestLogFailure Assert, "TestCreateRejectsEmptySheetName", , _
                         "Expected error when sheet name is empty"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when sheet name is empty"
End Sub

'@TestMethod("ListBuilder")
Public Sub TestCreateRejectsUnknownSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsUnknownSheet"
    On Error GoTo ExpectError

    Dim sut As IListBuilder
    Set sut = ListBuilder.Create(ListBuilderLayerHList, "NonExistentSheet__xyz", FakeLL)

    CustomTestLogFailure Assert, "TestCreateRejectsUnknownSheet", , _
                         "Expected error when sheet name is not in dictionary"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when sheet name is not in dictionary"
End Sub
