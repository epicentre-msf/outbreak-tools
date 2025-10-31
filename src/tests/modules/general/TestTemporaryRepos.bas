Attribute VB_Name = "TestTemporaryRepos"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Service As ITemporaryRepos
Private BaseFolder As String


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestTemporaryRepos"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
    RestoreApp
End Sub

'@TestInitialize
'@sub-title Prepare a fresh temporary repository before each test execution
Private Sub TestInitialize()
    'Use the host workbook directory as the base to avoid dependency on environment defaults.
    BaseFolder = ThisWorkbook.Path 
    ServiceReset
End Sub

'@TestCleanup
'@sub-title Flush harness output and tear down repository assets
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    ServiceReset True
End Sub

'@label:ServiceReset
'@sub-title Helper that rebuilds the temporary repository for each test
'@param removeFolder Boolean True to delete the folder after disposing of the service.
Private Sub ServiceReset(Optional ByVal removeFolder As Boolean)
    Dim folderPath As String

    'Initialise a fresh repository rooted under the workbook directory.
    Set Service = TemporaryRepos.Create(BaseFolder)
    Service.EnsureReady

    folderPath = Service.RootPath

    If removeFolder Then
        On Error Resume Next
            Service.DeleteAll
            RmDir folderPath 'Remove the root folder itself to guarantee isolation between tests
        On Error GoTo 0
    End If
End Sub


'@section Tests
'===============================================================================
'@TestMethod("TemporaryRepos")
Public Sub TestEnsureReadyCreatesFolder()
    CustomTestSetTitles Assert, "TemporaryRepos", "EnsureReadyCreatesFolder"

    Dim folderPath As String
    folderPath = Service.RootPath

    On Error GoTo Missing
        GetAttr folderPath
        Assert.LogSuccesses "Temporary folder exists"
        Exit Sub
Missing:
    Assert.LogFailure "EnsureReady should create the temporary folder"
End Sub

'@TestMethod("TemporaryRepos")
Public Sub TestCreateFilePathSanitisesName()
    CustomTestSetTitles Assert, "TemporaryRepos", "CreateFilePathSanitisesName"

    Dim filePath As String
    Dim fileLength As Long
    
    filePath = "test:module?.bas"
    fileLength = LenB(filePath)

    filePath = Service.CreateFilePath(filePath)

    Assert.IsTrue (InStr(1, filePath, ":", vbTextCompare) < (LenB(filePath) - fileLength)), "Sanitised path should not contain colon"
    Assert.IsTrue (InStr(1, filePath, "?", vbTextCompare) = 0), "Sanitised path should not contain question mark"
End Sub

'@TestMethod("TemporaryRepos")
Public Sub TestDeleteAllRemovesFiles()
    CustomTestSetTitles Assert, "TemporaryRepos", "DeleteAllRemovesFiles"

    Dim filePath As String
    filePath = Service.CreateFilePath("sample.bas")

    Open filePath For Output As #1
    Print #1, "data"
    Close #1

    Service.DeleteAll

    On Error GoTo FileMissing
        Open filePath For Input As #1
        Close #1
        Assert.LogFailure "DeleteAll should remove temporary files"
        Exit Sub
FileMissing:
    Err.Clear
    Assert.LogSuccesses "DeleteAll removed temporary files"
End Sub
