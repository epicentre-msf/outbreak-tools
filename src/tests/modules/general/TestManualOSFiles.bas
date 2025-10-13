Attribute VB_Name = "TestManualOSFiles"
Option Explicit
Option Private Module

'Dependencies: OSFiles.cls, IOSFiles.cls, CustomTest.cls, ICustomTest.cls, Checking.cls, IChecking.cls, CheckingOutput.cls, ICheckingOutput.cls, BetterArray.cls, TestHelpers.bas

'@TestModule
'@Folder("CustomTests")
'@ModuleDescription("Manual regression checks for OSFiles pickers. Requires user interaction.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestManualOSFiles"
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

'@TestMethod("ManualOSFiles")
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
