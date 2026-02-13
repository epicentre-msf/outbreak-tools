Attribute VB_Name = "TestTemporaryRepos"

'@ModuleDescription
'Test module for the TemporaryRepos class (ITemporaryRepos interface).
'Validates the temporary folder lifecycle: creation via EnsureReady, path
'generation and filename sanitisation via CreateFilePath, and file cleanup
'via DeleteAll. Each test starts with a freshly constructed repository
'rooted under the host workbook directory and tears down all artefacts
'on completion to guarantee full isolation between test runs.
'
'@depends TemporaryRepos, ITemporaryRepos, CustomTest, TestHelpers

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Service As ITemporaryRepos
Private BaseFolder As String


'@section Module lifecycle
'===============================================================================
'@description
'Initialisation and cleanup routines that bracket the entire module and each
'individual test. ModuleInitialize/ModuleCleanup run once; TestInitialize and
'TestCleanup wrap every Public test method.

'@ModuleInitialize
'@sub-title Initialise the test harness and suppress Excel UI updates
'@details
'Creates the CustomTest harness bound to the shared test output worksheet and
'registers this module name for result grouping. BusyApp disables screen
'updating, alerts, and animations so that file system operations do not
'trigger Excel UI redraws.
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestTemporaryRepos"
End Sub

'@ModuleCleanup
'@sub-title Print accumulated results and release the harness
'@details
'Ensures the output worksheet exists, prints the collected pass/fail results
'to it, releases the harness reference, and restores the standard Excel UI
'state via RestoreApp.
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
'@details
'Sets BaseFolder to the host workbook directory so that the temporary folder
'is created in a known, writable location independent of environment defaults.
'Calls ServiceReset without the remove flag, which constructs a new
'TemporaryRepos instance and ensures the root folder exists on disk.
Private Sub TestInitialize()
    'Use the host workbook directory as the base to avoid dependency on environment defaults.
    BaseFolder = ThisWorkbook.Path
    ServiceReset
End Sub

'@TestCleanup
'@sub-title Flush harness output and tear down repository assets
'@details
'Flushes any pending assertion output to the harness so results are not lost,
'then calls ServiceReset with removeFolder=True to delete all temporary files
'and the repository root folder itself, ensuring complete isolation between
'successive tests.
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    ServiceReset True
End Sub

'@label:ServiceReset
'@sub-title Helper that rebuilds the temporary repository for each test
'@details
'Instantiates a fresh TemporaryRepos rooted under BaseFolder and calls
'EnsureReady so the on-disk folder is guaranteed to exist. When removeFolder
'is True the helper additionally deletes all files via Service.DeleteAll
'and removes the root directory itself with RmDir, wrapped in error
'suppression so that already-deleted or locked paths do not raise errors.
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
'@description
'Each Public Sub below is a self-contained test method that exercises one
'behaviour of the TemporaryRepos class. Tests follow the Arrange/Act/Assert
'pattern and are registered under the "TemporaryRepos" group via
'CustomTestSetTitles.

'@TestMethod("TemporaryRepos")
'@sub-title Verify that EnsureReady creates the repository folder on disk
'@details
'After TestInitialize has already called EnsureReady, this test retrieves the
'RootPath from the service and uses GetAttr to probe the file system. If
'GetAttr succeeds without error the folder exists and the test passes. If
'GetAttr raises an error (file-not-found), execution jumps to the Missing
'label and logs a failure. This confirms that the core lifecycle method
'EnsureReady actually materialises the temporary directory.
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
'@sub-title Verify that CreateFilePath strips invalid characters from filenames
'@details
'Arranges a filename containing a colon and a question mark, both of which
'are illegal on Windows file systems. Acts by passing this filename through
'Service.CreateFilePath, which delegates to the internal SanitizeFileName
'routine. Asserts that the returned path no longer contains either character
'within the filename portion. The colon check uses a positional comparison
'against the original filename length to exclude any drive-letter colon that
'may be part of the root path prefix on Windows.
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
'@sub-title Verify that DeleteAll removes all files from the repository
'@details
'Arranges by creating a physical file named "sample.bas" inside the temporary
'repository using low-level Open/Print/Close statements. Acts by calling
'Service.DeleteAll to wipe all repository contents. Asserts by attempting to
'reopen the file for input: if the file still exists the test fails; if the
'Open raises an error (file-not-found), execution falls through to the
'FileMissing label and logs success. This confirms that DeleteAll correctly
'removes files previously created through CreateFilePath.
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
