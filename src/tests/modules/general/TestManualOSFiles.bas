Attribute VB_Name = "TestManualOSFiles"
Option Explicit
Option Private Module

'@TestModule
'@Folder("CustomTests")
'@ModuleDescription("Manual regression checks for OSFiles pickers. Requires user interaction.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

'@description
'Manual regression tests for the OSFiles class that exercise the real
'platform file and folder picker dialogs. Each test calls LoadFile,
'LoadFiles, or LoadFolders which opens a native OS dialog, so these
'tests cannot run unattended. If the user cancels the dialog the test
'fails with a descriptive message rather than a false positive.
'The suite verifies three interaction paths: single-file selection via
'LoadFile, multi-file selection via LoadFiles with full iterator
'traversal, and multi-folder selection via LoadFolders with full
'iterator traversal. A shared SafeArrayLength helper is used to count
'picker results safely without risking an error on uninitialised arrays.
'Uses the CustomTest runner with results printed to the testsOutputs sheet.
'@depends OSFiles, IOSFiles, CustomTest, ICustomTest, TestHelpers

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
'@sub-title Prepare the test harness before the first test in this module runs.
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestManualOSFiles"
End Sub

'@ModuleCleanup
'@sub-title Print accumulated results and release the test harness after all tests finish.
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
'@sub-title Flush previous assertions and suppress screen updating before each test.
Private Sub TestInitialize()
    BusyApp
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@TestCleanup
'@sub-title Restore normal application state after each test.
Private Sub TestCleanup()
    RestoreApp
End Sub

'@section Helpers
'===============================================================================

'@sub-title Safely compute the number of elements in a Variant that may or may not be an initialised array.
'@details
'Uses On Error Resume Next to guard against uninitialised or empty
'arrays where LBound/UBound would raise an error. Returns 0 for
'non-array values, uninitialised arrays, or arrays whose UBound is
'below their LBound.
Private Function SafeArrayLength(ByVal candidate As Variant) As Long
    Dim lowerBound As Long
    Dim upperBound As Long

    On Error Resume Next
    If IsArray(candidate) Then
        lowerBound = LBound(candidate)
        upperBound = UBound(candidate)
        If Err.Number = 0 Then
            If upperBound >= lowerBound Then
                SafeArrayLength = (upperBound - lowerBound) + 1
            End If
        Else
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Function

'@section Tests
'===============================================================================

'@TestMethod("ManualOSFiles")
'@sub-title Verify single-file selection via the native file picker dialog.
'@details
'Arranges by creating a fresh OSFiles instance and calling LoadFile with
'a "*.xlsb" filter, which opens the platform file picker for the user
'to choose one file. If the user cancels, the test fails with a
'descriptive skip message. When a file is selected, asserts that
'HasValidFile returns True and that the scalar File() accessor agrees
'with the first element of the Files() array. This guards against
'regressions where the single-selection path diverges from the
'collection path.
Public Sub TestManualSelectSingleFile()
    CustomTestSetTitles Assert, "ManualOSFiles", "SelectSingleFile"
    On Error GoTo Fail

    Dim os As IOSFiles
    Dim selected As Variant

    Set os = OSFiles.Create()
    os.LoadFile "*.xlsb"

    If Not os.HasValidFile() Then
        Assert.Fail "No file was selected. Re-run and choose a file to verify behaviour."
        Exit Sub
    End If

    selected = os.Files()

    Assert.IsTrue os.HasValidFile(), "Single selection should mark HasValidFile True"
    Assert.AreEqual os.File(), CStr(selected(LBound(selected))), "File() should match first entry in Files()"

    Exit Sub
Fail:
    FailUnexpectedError Assert, "TestManualSelectSingleFile"
End Sub

'@TestMethod("ManualOSFiles")
'@sub-title Verify multi-file selection and iterator traversal via the native file picker dialog.
'@details
'Arranges by creating a fresh OSFiles instance and calling LoadFiles
'with a combined "*.xlsb, *.xlsx" filter. The user is expected to
'select one or more files from the picker. If the user cancels, the
'test fails with a descriptive skip message. When files are selected,
'asserts that HasValidFiles and HasNextFile both return True, then
'iterates through all files using HasNextFile/NextFile, counting
'non-empty paths. Finally compares the iterator count against the
'SafeArrayLength of Files() to ensure the iterator visits every
'selected file exactly once.
Public Sub TestManualSelectMultipleFiles()
    CustomTestSetTitles Assert, "ManualOSFiles", "SelectMultipleFiles"
    On Error GoTo Fail

    Dim os As IOSFiles
    Dim files As Variant
    Dim count As Long

    Set os = OSFiles.Create()
    os.LoadFiles "*.xlsb, *.xlsx"

    files = os.Files()
    count = SafeArrayLength(files)

    If count = 0 Then
        Assert.Fail "No files were selected. Re-run and choose one or more files to continue testing."
        Exit Sub
    End If

    Assert.IsTrue os.HasValidFiles(), "HasValidFiles should be true when selections are provided"
    Assert.IsTrue os.HasNextFile(), "Iterator should detect at least one file"

    Dim seen As Long
    Dim path As String

    Do While os.HasNextFile()
        path = os.NextFile()
        If path <> vbNullString Then
            seen = seen + 1
        End If
    Loop

    Assert.AreEqual count, seen, "Iterator should traverse all selected files"

    Exit Sub
Fail:
    FailUnexpectedError Assert, "TestManualSelectMultipleFiles"
End Sub

'@TestMethod("ManualOSFiles")
'@sub-title Verify folder selection and iterator traversal via the native folder picker dialog.
'@details
'Arranges by creating a fresh OSFiles instance and calling LoadFolders,
'which opens the platform folder picker for the user to choose one or
'more directories. If the user cancels, the test fails with a
'descriptive skip message. When folders are selected, asserts that
'HasValidFolders and HasNextFolder both return True, then iterates
'through all folders using HasNextFolder/NextFolder, counting non-empty
'paths. Compares the iterator count against the SafeArrayLength of
'Folders() to ensure complete traversal without duplication or omission.
Public Sub TestManualSelectFolders()
    CustomTestSetTitles Assert, "ManualOSFiles", "SelectFolders"
    On Error GoTo Fail

    Dim os As IOSFiles
    Dim folders As Variant
    Dim count As Long

    Set os = OSFiles.Create()
    os.LoadFolders

    folders = os.Folders()
    count = SafeArrayLength(folders)

    If count = 0 Then
        Assert.Fail "No folders were selected. Re-run and choose at least one folder to verify the picker."
        Exit Sub
    End If

    Assert.IsTrue os.HasValidFolders(), "HasValidFolders should be true when selections are provided"
    Assert.IsTrue os.HasNextFolder(), "Iterator should recognise the first folder"

    Dim seen As Long
    Dim folderPath As String

    Do While os.HasNextFolder()
        folderPath = os.NextFolder()
        If folderPath <> vbNullString Then
            seen = seen + 1
        End If
    Loop

    Assert.AreEqual count, seen, "Iterator should traverse all selected folders"

    Exit Sub
Fail:
    FailUnexpectedError Assert, "TestManualSelectFolders"
End Sub
