Attribute VB_Name = "TestLinelist"
Attribute VB_Description = "Tests for Linelist class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for Linelist class")

Option Explicit

Private Assert As ICustomTest
Private SpecsWkb As Workbook
Private Specs As LinelistSpecsWorkbookStub
Private Dict As ILLdictionary

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "Linelist"
Private Const DICTIONARY_SHEET As String = "DictFixture"


'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelist"
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
    Set SpecsWkb = TestHelpers.NewWorkbook
    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, SpecsWkb
    Set Dict = LLdictionary.Create(SpecsWkb.Worksheets(DICTIONARY_SHEET), 1, 1)
    Dict.Prepare

    Set Specs = New LinelistSpecsWorkbookStub
    Specs.Initialise Dict, SpecsWkb

    Assert.BeginTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not SpecsWkb Is Nothing Then TestHelpers.DeleteWorkbook SpecsWkb
    On Error GoTo 0

    Set Dict = Nothing
    Set Specs = Nothing
    Set SpecsWkb = Nothing
End Sub


'@section Factory tests
'===============================================================================

'@TestMethod("Linelist")
Public Sub TestCreateReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsInstance"
    On Error GoTo TestFail

    Dim sut As ILinelist
    Set sut = Linelist.Create(Specs)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsInstance", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestCreateRejectsNothingSpecs()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingSpecs"
    On Error GoTo ExpectError

    Dim sut As ILinelist
    Set sut = Linelist.Create(Nothing)

    CustomTestLogFailure Assert, "TestCreateRejectsNothingSpecs", , _
                         "Expected error when specs is Nothing"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when specs is Nothing"
End Sub


'@section Property tests
'===============================================================================

'@TestMethod("Linelist")
Public Sub TestLinelistDataReturnsSpecs()
    CustomTestSetTitles Assert, TESTMODULE, "TestLinelistDataReturnsSpecs"
    On Error GoTo TestFail

    Dim sut As ILinelist
    Set sut = Linelist.Create(Specs)

    Assert.IsTrue Not sut.LinelistData Is Nothing, _
                  "LinelistData should return the specifications object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLinelistDataReturnsSpecs", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestDictionaryReturnsDictionary()
    CustomTestSetTitles Assert, TESTMODULE, "TestDictionaryReturnsDictionary"
    On Error GoTo TestFail

    Dim sut As ILinelist
    Set sut = Linelist.Create(Specs)

    Assert.IsTrue Not sut.Dictionary Is Nothing, _
                  "Dictionary should return the dictionary from specs"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDictionaryReturnsDictionary", Err.Number, Err.Description
End Sub

'@TestMethod("Linelist")
Public Sub TestSheetExistsReturnsFalse()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetExistsReturnsFalse"
    On Error GoTo TestFail

    Dim sut As ILinelist
    Set sut = Linelist.Create(Specs)

    Assert.IsTrue Not sut.SheetExists("NonExistentSheet__xyz"), _
                  "SheetExists should return False for a non-existing sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetExistsReturnsFalse", Err.Number, Err.Description
End Sub
