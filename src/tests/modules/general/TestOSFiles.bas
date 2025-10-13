Attribute VB_Name = "TestOSFiles"
Option Explicit
Option Private Module

'@TestModule
'@Folder("CustomTests")
'@ModuleDescription("Unit tests validating OSFiles default state and unsupported-platform fallbacks")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestOSFiles"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@TestCleanup
Private Sub TestCleanup()
    RestoreApp
End Sub

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

'@TestMethod("OSFiles")
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

'@TestMethod("OSFiles")
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

'@TestMethod("OSFiles")
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
