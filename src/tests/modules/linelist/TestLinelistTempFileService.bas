Attribute VB_Name = "TestLinelistTempFileService"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Service As ILinelistTempFileService
Private BaseFolder As String


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelistTempFileService"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
    RestoreApp
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BaseFolder = Application.DefaultFilePath & Application.PathSeparator & "LLTempTest_"
    ServiceReset
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    ServiceReset True
End Sub

Private Sub ServiceReset(Optional ByVal removeFolder As Boolean)
    Dim folderPath As String

    Set Service = LinelistTempFileService.Create(BaseFolder)
    Service.EnsureReady

    folderPath = Service.RootPath

    If removeFolder Then
        On Error Resume Next
            Service.DeleteAll
            RmDir folderPath
        On Error GoTo 0
    End If
End Sub


'@section Tests
'===============================================================================
'@TestMethod("LinelistTempFileService")
Public Sub TestEnsureReadyCreatesFolder()
    CustomTestSetTitles Assert, "LinelistTempFileService", "EnsureReadyCreatesFolder"

    Dim folderPath As String
    folderPath = Service.RootPath

    On Error GoTo Missing
        GetAttr folderPath
        Assert.Pass "Temporary folder exists"
        Exit Sub
Missing:
    Assert.Fail "EnsureReady should create the temporary folder"
End Sub

'@TestMethod("LinelistTempFileService")
Public Sub TestCreateFilePathSanitisesName()
    CustomTestSetTitles Assert, "LinelistTempFileService", "CreateFilePathSanitisesName"

    Dim filePath As String
    filePath = Service.CreateFilePath("test:module?.bas")

    Assert.IsTrue InStr(1, filePath, ":", vbBinaryCompare) = 0, "Sanitised path should not contain colon"
    Assert.IsTrue InStr(1, filePath, "?", vbBinaryCompare) = 0, "Sanitised path should not contain question mark"
End Sub

'@TestMethod("LinelistTempFileService")
Public Sub TestDeleteAllRemovesFiles()
    CustomTestSetTitles Assert, "LinelistTempFileService", "DeleteAllRemovesFiles"

    Dim filePath As String
    filePath = Service.CreateFilePath("sample.bas")

    Open filePath For Output As #1
    Print #1, "data"
    Close #1

    Service.DeleteAll

    On Error GoTo FileMissing
        Open filePath For Input As #1
        Close #1
        Assert.Fail "DeleteAll should remove temporary files"
        Exit Sub
FileMissing:
    Err.Clear
End Sub

