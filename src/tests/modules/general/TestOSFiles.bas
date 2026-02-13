Attribute VB_Name = "TestOSFiles"
Option Explicit
Option Private Module

'@TestModule
'@Folder("CustomTests")
'@ModuleDescription("Unit tests validating OSFiles default state and unsupported-platform fallbacks")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

'@description
'Automated unit tests for the OSFiles class that run without user
'interaction. The suite validates two key areas: (1) the default state
'of a freshly-created OSFiles instance, ensuring all validity flags are
'false and all collections are empty; and (2) the behaviour when the OS
'property is set to an unrecognised platform string, verifying that
'LoadFiles and LoadFolders silently no-op rather than raising errors.
'Additionally, two tests exercise the file and folder iterators by
'injecting test data via AssignFilesForTesting/AssignFoldersForTesting,
'verifying sequential traversal, exhaustion semantics (vbNullString at
'end), and correct restart after calling ResetFilesIterator or
'ResetFoldersIterator.
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
    Assert.SetModuleName "TestOSFiles"
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

'@section Tests -- Initial State
'===============================================================================

'@TestMethod("OSFiles")
'@sub-title Verify that a freshly-created OSFiles instance has no file selections and empty collections.
'@details
'Arranges by creating a new OSFiles via the factory method. Acts by
'querying HasValidFiles, HasValidFile, Files(), and File(). Asserts
'that both validity flags return False, that Files() is a Variant array
'with zero elements (via SafeArrayLength), and that the scalar File()
'accessor returns vbNullString. This ensures no stale data leaks from
'previous uses or from uninitialised internal state.
Public Sub TestFilesCollectionInitialState()
    CustomTestSetTitles Assert, "OSFiles", "FilesCollectionInitialState"
    On Error GoTo Fail

    Dim os As IOSFiles
    Dim files As Variant

    Set os = OSFiles.Create()
    files = os.Files()

    Assert.IsFalse os.HasValidFiles(), "Fresh instance should not report file selections"
    Assert.IsFalse os.HasValidFile(), "Legacy HasValidFile should remain false initially"
    Assert.IsTrue IsArray(files), "Files() should expose a variant array"
    Assert.AreEqual 0&, SafeArrayLength(files), "Files() should be empty after creation"
    Assert.AreEqual vbNullString, os.File(), "File() should return an empty string when nothing is selected"

    Exit Sub
Fail:
    FailUnexpectedError Assert, "TestFilesCollectionInitialState"
End Sub

'@TestMethod("OSFiles")
'@sub-title Verify that a freshly-created OSFiles instance has no folder selections and empty collections.
'@details
'Arranges by creating a new OSFiles via the factory method. Acts by
'querying HasValidFolders, HasValidFolder, Folders(), and Folder().
'Asserts that both validity flags return False, that Folders() is a
'Variant array with zero elements, and that the scalar Folder()
'accessor returns vbNullString. This mirrors the file-side initial
'state test to confirm symmetry across the file and folder APIs.
Public Sub TestFoldersCollectionInitialState()
    CustomTestSetTitles Assert, "OSFiles", "FoldersCollectionInitialState"
    On Error GoTo Fail

    Dim os As IOSFiles
    Dim folders As Variant

    Set os = OSFiles.Create()
    folders = os.Folders()

    Assert.IsFalse os.HasValidFolders(), "Fresh instance should not report folder selections"
    Assert.IsFalse os.HasValidFolder(), "Legacy HasValidFolder should remain false initially"
    Assert.IsTrue IsArray(folders), "Folders() should expose a variant array"
    Assert.AreEqual 0&, SafeArrayLength(folders), "Folders() should be empty after creation"
    Assert.AreEqual vbNullString, os.Folder(), "Folder() should return an empty string when nothing is selected"

    Exit Sub
Fail:
    FailUnexpectedError Assert, "TestFoldersCollectionInitialState"
End Sub

'@section Tests -- Unsupported platform fallback
'===============================================================================

'@TestMethod("OSFiles")
'@sub-title Verify that LoadFiles is a no-op when the OS property is set to an unrecognised platform.
'@details
'Arranges by instantiating OSFiles directly (not via Create) and
'setting the OS property to "Unsupported" to simulate an unknown
'platform. Acts by calling LoadFiles with a "*.xlsx" filter and
'inspecting the results. Asserts that HasValidFiles remains False,
'File() returns vbNullString, and Files() contains zero elements.
'This confirms the class gracefully degrades rather than raising
'runtime errors on platforms it does not recognise.
Public Sub TestLoadFilesUnsupportedSystemDoesNotPopulate()
    CustomTestSetTitles Assert, "OSFiles", "LoadFilesUnsupportedSystemDoesNotPopulate"
    On Error GoTo Fail

    Dim raw As OSFiles
    Dim sut As IOSFiles
    Dim files As Variant

    Set raw = New OSFiles
    raw.OS = "Unsupported"
    Set sut = raw.Self

    sut.LoadFiles "*.xlsx"
    files = sut.Files()

    Assert.IsFalse sut.HasValidFiles(), "Unsupported OS should not record file selections"
    Assert.AreEqual vbNullString, sut.File(), "File() should remain empty when nothing is selected"
    Assert.AreEqual 0&, SafeArrayLength(files), "Files() should remain empty when the picker does not run"

    Exit Sub
Fail:
    FailUnexpectedError Assert, "TestLoadFilesUnsupportedSystemDoesNotPopulate"
End Sub

'@TestMethod("OSFiles")
'@sub-title Verify that LoadFolders is a no-op when the OS property is set to an unrecognised platform.
'@details
'Arranges by instantiating OSFiles directly (not via Create) and
'setting the OS property to "Unsupported" to simulate an unknown
'platform. Acts by calling LoadFolders and inspecting the results.
'Asserts that HasValidFolders remains False, Folder() returns
'vbNullString, and Folders() contains zero elements. This mirrors
'the file-side unsupported platform test to confirm both picker
'paths degrade gracefully.
Public Sub TestLoadFoldersUnsupportedSystemDoesNotPopulate()
    CustomTestSetTitles Assert, "OSFiles", "LoadFoldersUnsupportedSystemDoesNotPopulate"
    On Error GoTo Fail

    Dim raw As OSFiles
    Dim sut As IOSFiles
    Dim folders As Variant

    Set raw = New OSFiles
    raw.OS = "Unsupported"
    Set sut = raw.Self

    sut.LoadFolders
    folders = sut.Folders()

    Assert.IsFalse sut.HasValidFolders(), "Unsupported OS should not record folder selections"
    Assert.AreEqual vbNullString, sut.Folder(), "Folder() should remain empty when nothing is selected"
    Assert.AreEqual 0&, SafeArrayLength(folders), "Folders() should remain empty when the picker does not run"

    Exit Sub
Fail:
    FailUnexpectedError Assert, "TestLoadFoldersUnsupportedSystemDoesNotPopulate"
End Sub

'@section Tests -- Iterator traversal
'===============================================================================

'@TestMethod("OSFiles")
'@sub-title Verify file iterator traversal, exhaustion, and reset using injected test data.
'@details
'Arranges by instantiating OSFiles directly and calling
'AssignFilesForTesting with a two-element array ("fileA", "fileB") to
'bypass the native picker dialog. Acts by stepping through the iterator
'with HasNextFile/NextFile and verifying each returned value in order.
'Asserts that HasValidFiles becomes True after assignment, that
'NextFile returns "fileA" then "fileB" in sequence, that HasNextFile
'becomes False once exhausted, and that NextFile returns vbNullString
'past the end. Then calls ResetFilesIterator and asserts that the
'iterator restarts from "fileA", confirming reusability.
Public Sub TestFileIteratorTraversesAssignedSelection()
    CustomTestSetTitles Assert, "OSFiles", "FileIteratorTraversesAssignedSelection"
    On Error GoTo Fail

    Dim raw As OSFiles
    Dim sut As IOSFiles
    Dim results As Variant

    Set raw = New OSFiles
    raw.AssignFilesForTesting Array("fileA", "fileB")
    Set sut = raw.Self

    Assert.IsTrue sut.HasValidFiles(), "AssignFilesForTesting should seed selections"
    Assert.IsTrue sut.HasNextFile(), "Iterator should detect first element"
    Assert.AreEqual "fileA", sut.NextFile(), "First NextFile call should return first element"
    Assert.IsTrue sut.HasNextFile(), "Iterator should detect remaining elements"
    Assert.AreEqual "fileB", sut.NextFile(), "Second NextFile call should return second element"
    Assert.IsFalse sut.HasNextFile(), "Iterator should report completion"
    Assert.AreEqual vbNullString, sut.NextFile(), "NextFile should return vbNullString when exhausted"

    sut.ResetFilesIterator
    Assert.IsTrue sut.HasNextFile(), "Iterator reset should allow iteration again"
    results = Array(sut.NextFile(), sut.NextFile())
    Assert.AreEqual "fileA", CStr(results(0)), "Iterator should restart from first element"
    Assert.AreEqual "fileB", CStr(results(1)), "Iterator should continue in order"

    Exit Sub
Fail:
    FailUnexpectedError Assert, "TestFileIteratorTraversesAssignedSelection"
End Sub

'@TestMethod("OSFiles")
'@sub-title Verify folder iterator traversal, exhaustion, and reset using injected test data.
'@details
'Arranges by instantiating OSFiles directly and calling
'AssignFoldersForTesting with a two-element array ("/tmp", "/var") to
'bypass the native picker dialog. Acts by stepping through the iterator
'with HasNextFolder/NextFolder and verifying each returned value in
'order. Asserts that HasValidFolders becomes True after assignment,
'that NextFolder returns "/tmp" then "/var" in sequence, that
'HasNextFolder becomes False once exhausted, and that NextFolder
'returns vbNullString past the end. Then calls ResetFoldersIterator
'and verifies the iterator restarts from "/tmp", confirming
'reusability of the folder enumeration.
Public Sub TestFolderIteratorTraversesAssignedSelection()
    CustomTestSetTitles Assert, "OSFiles", "FolderIteratorTraversesAssignedSelection"
    On Error GoTo Fail

    Dim raw As OSFiles
    Dim sut As IOSFiles

    Set raw = New OSFiles
    raw.AssignFoldersForTesting Array("/tmp", "/var")
    Set sut = raw.Self

    Assert.IsTrue sut.HasValidFolders(), "AssignFoldersForTesting should seed folder selections"
    Assert.IsTrue sut.HasNextFolder(), "Iterator should detect first folder"
    Assert.AreEqual "/tmp", sut.NextFolder(), "NextFolder should return first folder"
    Assert.IsTrue sut.HasNextFolder(), "Iterator should detect second folder"
    Assert.AreEqual "/var", sut.NextFolder(), "NextFolder should return second folder"
    Assert.IsFalse sut.HasNextFolder(), "Iterator should report completion"
    Assert.AreEqual vbNullString, sut.NextFolder(), "NextFolder should return vbNullString when exhausted"

    sut.ResetFoldersIterator
    Assert.IsTrue sut.HasNextFolder(), "ResetFoldersIterator should restart enumeration"
    Assert.AreEqual "/tmp", sut.NextFolder(), "Enumeration should restart from first folder"

    Exit Sub
Fail:
    FailUnexpectedError Assert, "TestFolderIteratorTraversesAssignedSelection"
End Sub
