Attribute VB_Name = "TestCodeTransfer"
Attribute VB_Description = "Tests for CodeTransfer class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for CodeTransfer class")

Option Explicit

Private Assert As ICustomTest
Private SourceWkb As Workbook
Private TargetWkb As Workbook
Private TempRepos As ITemporaryRepos

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "CodeTransfer"


'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestCodeTransfer"
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
    Set SourceWkb = TestHelpers.NewWorkbook
    Set TargetWkb = TestHelpers.NewWorkbook
    Set TempRepos = TemporaryRepos.Create()

    Assert.BeginTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not TempRepos Is Nothing Then TempRepos.Reset
        If Not SourceWkb Is Nothing Then TestHelpers.DeleteWorkbook SourceWkb
        If Not TargetWkb Is Nothing Then TestHelpers.DeleteWorkbook TargetWkb
    On Error GoTo 0

    Set SourceWkb = Nothing
    Set TargetWkb = Nothing
    Set TempRepos = Nothing
End Sub


'@section Factory tests
'===============================================================================

'@TestMethod("CodeTransfer")
Public Sub TestCreateReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsInstance"
    On Error GoTo TestFail

    Dim sut As ICodeTransfer
    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsInstance", Err.Number, Err.Description
End Sub

'@TestMethod("CodeTransfer")
Public Sub TestCreateRejectsNothingSource()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingSource"
    On Error GoTo ExpectError

    Dim sut As ICodeTransfer
    Set sut = CodeTransfer.Create(Nothing, TargetWkb, TempRepos)

    CustomTestLogFailure Assert, "TestCreateRejectsNothingSource", , _
                         "Expected error when source workbook is Nothing"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when source workbook is Nothing"
End Sub

'@TestMethod("CodeTransfer")
Public Sub TestCreateRejectsNothingTarget()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingTarget"
    On Error GoTo ExpectError

    Dim sut As ICodeTransfer
    Set sut = CodeTransfer.Create(SourceWkb, Nothing, TempRepos)

    CustomTestLogFailure Assert, "TestCreateRejectsNothingTarget", , _
                         "Expected error when target workbook is Nothing"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when target workbook is Nothing"
End Sub

'@TestMethod("CodeTransfer")
Public Sub TestCreateRejectsNothingTempRepos()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingTempRepos"
    On Error GoTo ExpectError

    Dim sut As ICodeTransfer
    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, Nothing)

    CustomTestLogFailure Assert, "TestCreateRejectsNothingTempRepos", , _
                         "Expected error when tempRepos is Nothing"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when tempRepos is Nothing"
End Sub
