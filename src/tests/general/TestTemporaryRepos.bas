Attribute VB_Name = "TestTemporaryRepos"
Attribute VB_Description = "Tests for the TemporaryRepos class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the TemporaryRepos class")

'@description
'Drives TemporaryRepos, the folder LinelistSpecs.Prepare writes its working
'files into. Prepare saves the template copy of the output workbook there, and
'a failed Prepare drops the instance so the working files go with it.
'
'Every test hands the factory an explicit base folder under the test workbook's
'own directory. The default base is Application.DefaultFilePath, which sits
'outside the folder this run is allowed to write to.
'
'The suite this replaced had three tests and its module hooks were Private. Its
'DeleteAll test wrote one file, and one file is the case that passes whatever
'the walk does.
'@depends TemporaryRepos, CustomTest

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TEST_FOLDER_NAME As String = "obt_temprepos_tests"

'What the class appends to the base folder it is given.
Private Const REPO_FOLDER_NAME As String = "OBTApp_"

Private Assert As CustomTest
Private BaseFolder As String

'@section Module lifecycle
'===============================================================================

'@sub-title Set up the assertion harness and a writable base folder.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestTemporaryRepos"

    BaseFolder = BuildTempFolder(Nothing, TEST_FOLDER_NAME)
End Sub

'@sub-title Print results and remove everything the suite wrote.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RemoveRepoFolder
    On Error Resume Next
        RmDir BaseFolder
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Remove the repository folder before each test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    RemoveRepoFolder
End Sub

'@sub-title Flush assert state after each test.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section The root path
'===============================================================================

'@sub-title The root sits under the base folder and ends with a separator.
'@details
'CreateFilePath joins the root and the file name with no separator of its own,
'so a root with no trailing separator would run the two names together.
'@TestMethod("TemporaryRepos")
Public Sub TestRootPathEndsWithASeparator()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestRootPathEndsWithASeparator"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos
    Dim root As String

    Set repos = TemporaryRepos.Create(BaseFolder)
    root = repos.RootPath()

    Assert.AreEqual Application.PathSeparator, Right$(root, 1), _
                    "The root path ends with the path separator"
    Assert.AreEqual CLng(1), CLng(InStr(1, root, BaseFolder, vbTextCompare)), _
                    "The root path opens with the base folder it was given"
    Assert.IsTrue (InStr(1, root, REPO_FOLDER_NAME, vbTextCompare) > 0), _
                  "The root path carries the repository folder name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRootPathEndsWithASeparator", Err.Number, Err.Description
End Sub

'@sub-title A base folder already ending with a separator gives one separator.
'@TestMethod("TemporaryRepos")
Public Sub TestBaseFolderSeparatorIsNotDoubled()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestBaseFolderSeparatorIsNotDoubled"
    On Error GoTo TestFail

    Dim withSeparator As TemporaryRepos
    Dim withoutSeparator As TemporaryRepos

    Set withSeparator = TemporaryRepos.Create(BaseFolder & Application.PathSeparator)
    Set withoutSeparator = TemporaryRepos.Create(BaseFolder)

    Assert.AreEqual withoutSeparator.RootPath(), withSeparator.RootPath(), _
                    "A trailing separator on the base folder changes nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBaseFolderSeparatorIsNotDoubled", _
                         Err.Number, Err.Description
End Sub

'@section Creating the folder
'===============================================================================

'@sub-title The folder appears on disk once EnsureReady has run.
'@details
'Create on its own writes nothing, which is what lets a caller build the
'object and decide later whether it needs the folder at all.
'@TestMethod("TemporaryRepos")
Public Sub TestEnsureReadyCreatesTheFolder()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestEnsureReadyCreatesTheFolder"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos

    Set repos = TemporaryRepos.Create(BaseFolder)

    Assert.IsFalse FolderExists(repos.RootPath()), _
                   "Create on its own writes nothing to disk"

    repos.EnsureReady

    Assert.IsTrue FolderExists(repos.RootPath()), "EnsureReady creates the folder"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEnsureReadyCreatesTheFolder", _
                         Err.Number, Err.Description
End Sub

'@sub-title EnsureReady over a folder that is already there succeeds.
'@details
'Two generations in one session build a second instance over the same folder,
'so the second EnsureReady meets a folder that exists.
'@TestMethod("TemporaryRepos")
Public Sub TestEnsureReadyIsSafeToRepeat()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestEnsureReadyIsSafeToRepeat"
    On Error GoTo TestFail

    Dim first As TemporaryRepos
    Dim second As TemporaryRepos

    Set first = TemporaryRepos.Create(BaseFolder)
    first.EnsureReady

    Set second = TemporaryRepos.Create(BaseFolder)
    second.EnsureReady

    Assert.IsTrue FolderExists(second.RootPath()), _
                  "A second instance over the same folder is ready too"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEnsureReadyIsSafeToRepeat", Err.Number, Err.Description
End Sub

'@section File paths
'===============================================================================

'@sub-title CreateFilePath builds a path inside the repository and readies it.
'@TestMethod("TemporaryRepos")
Public Sub TestCreateFilePathJoinsTheRoot()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestCreateFilePathJoinsTheRoot"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos
    Dim filePath As String

    Set repos = TemporaryRepos.Create(BaseFolder)
    filePath = repos.CreateFilePath("__temp.xlsb")

    Assert.AreEqual repos.RootPath() & "__temp.xlsb", filePath, _
                    "The path is the root followed by the file name"
    Assert.IsTrue FolderExists(repos.RootPath()), _
                  "Asking for a file path readies the folder that holds it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateFilePathJoinsTheRoot", Err.Number, Err.Description
End Sub

'@sub-title Characters the file systems refuse are replaced.
'@details
'The separator is on that list, so a name carrying one stays inside the
'repository folder.
'@TestMethod("TemporaryRepos")
Public Sub TestCreateFilePathSanitisesTheName()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestCreateFilePathSanitisesTheName"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos
    Dim filePath As String

    Set repos = TemporaryRepos.Create(BaseFolder)
    filePath = repos.CreateFilePath("a/b:c*d?e|f<g>h,i.xlsb")

    Assert.AreEqual repos.RootPath() & "a_b_c_d_e_f_g_h_i.xlsb", filePath, _
                    "Every character the file systems refuse becomes an underscore"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateFilePathSanitisesTheName", _
                         Err.Number, Err.Description
End Sub

'@sub-title A file name of spaces is refused.
'@TestMethod("TemporaryRepos")
Public Sub TestCreateFilePathRejectsAnEmptyName()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestCreateFilePathRejectsAnEmptyName"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos
    Dim errNumber As Long
    Dim answer As String

    Set repos = TemporaryRepos.Create(BaseFolder)

    On Error Resume Next
        answer = repos.CreateFilePath("   ")
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A file name of spaces is refused"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateFilePathRejectsAnEmptyName", _
                         Err.Number, Err.Description
End Sub

'@section Deleting
'===============================================================================

'@sub-title DeleteFile removes one file and ignores a missing one.
'@TestMethod("TemporaryRepos")
Public Sub TestDeleteFileRemovesOneFile()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestDeleteFileRemovesOneFile"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos
    Dim filePath As String

    Set repos = TemporaryRepos.Create(BaseFolder)
    filePath = repos.CreateFilePath("one.txt")
    WriteTextFile filePath, "one"

    Assert.IsTrue FileExists(filePath), "The file was written"

    repos.DeleteFile filePath
    Assert.IsFalse FileExists(filePath), "DeleteFile removed it"

    repos.DeleteFile filePath
    Assert.IsFalse FileExists(filePath), "Deleting it a second time raises nothing"

    repos.DeleteFile vbNullString
    Assert.IsFalse FileExists(filePath), "An empty path is ignored"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDeleteFileRemovesOneFile", Err.Number, Err.Description
End Sub

'@sub-title DeleteAll empties the folder and leaves the folder standing.
'@details
'Several files at once is the case that matters. The walk used to delete each
'entry as it read the listing, and Dir$ walks a folder through state the host
'holds, so deleting during the walk left files behind. One file passes either
'way, which is why the old suite never showed it.
'@TestMethod("TemporaryRepos")
Public Sub TestDeleteAllEmptiesTheFolder()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestDeleteAllEmptiesTheFolder"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos
    Dim idx As Long

    Set repos = TemporaryRepos.Create(BaseFolder)

    For idx = 1 To 6
        WriteTextFile repos.CreateFilePath("file" & CStr(idx) & ".txt"), "x"
    Next idx

    Assert.AreEqual CLng(6), FileCountIn(repos.RootPath()), "Six files were written"

    repos.DeleteAll

    Assert.AreEqual CLng(0), FileCountIn(repos.RootPath()), _
                    "DeleteAll leaves no file behind"
    Assert.IsTrue FolderExists(repos.RootPath()), "The folder itself is still there"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDeleteAllEmptiesTheFolder", Err.Number, Err.Description
End Sub

'@sub-title Reset empties the folder and the repository is usable afterwards.
'@TestMethod("TemporaryRepos")
Public Sub TestResetEmptiesAndReopens()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestResetEmptiesAndReopens"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos
    Dim filePath As String

    Set repos = TemporaryRepos.Create(BaseFolder)
    WriteTextFile repos.CreateFilePath("before.txt"), "x"
    WriteTextFile repos.CreateFilePath("before2.txt"), "x"

    repos.Reset

    Assert.AreEqual CLng(0), FileCountIn(repos.RootPath()), "Reset leaves no file behind"

    filePath = repos.CreateFilePath("after.txt")
    WriteTextFile filePath, "y"

    Assert.IsTrue FileExists(filePath), "The repository is usable again after a reset"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestResetEmptiesAndReopens", Err.Number, Err.Description
End Sub

'@sub-title Dropping the instance empties the folder.
'@details
'This is what a failed Prepare relies on: it closes the output workbook, drops
'the repository, and the working files go with it.
'@TestMethod("TemporaryRepos")
Public Sub TestDroppingTheInstanceEmptiesTheFolder()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestDroppingTheInstanceEmptiesTheFolder"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos
    Dim root As String

    Set repos = TemporaryRepos.Create(BaseFolder)
    root = repos.RootPath()
    WriteTextFile repos.CreateFilePath("leftover.txt"), "x"
    WriteTextFile repos.CreateFilePath("leftover2.txt"), "x"

    Assert.AreEqual CLng(2), FileCountIn(root), "Two working files are on disk"

    Set repos = Nothing

    Assert.AreEqual CLng(0), FileCountIn(root), _
                    "Dropping the instance takes the working files with it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDroppingTheInstanceEmptiesTheFolder", _
                         Err.Number, Err.Description
End Sub

'@sub-title A file KeepFile marked survives the drop of the instance.
'@details
'This is what a failed build relies on: it saves the unfinished linelist
'as __temp.xlsb, marks it, and drops the repository with the rest of the
'working files.
'@TestMethod("TemporaryRepos")
Public Sub TestKeepFileSurvivesTheInstanceDrop()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestKeepFileSurvivesTheInstanceDrop"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos
    Dim root As String
    Dim keptPath As String

    Set repos = TemporaryRepos.Create(BaseFolder)
    root = repos.RootPath()
    keptPath = repos.CreateFilePath("__temp.xlsb")
    WriteTextFile keptPath, "x"
    WriteTextFile repos.CreateFilePath("component.bas"), "x"

    repos.KeepFile "__temp.xlsb"
    Set repos = Nothing

    Assert.AreEqual CLng(1), FileCountIn(root), _
                    "Dropping the instance takes every file but the kept one"
    Assert.IsTrue FileExists(keptPath), "The kept file is still on disk"

    'Leave nothing behind for the next test
    On Error Resume Next
    Kill keptPath
    On Error GoTo TestFail

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestKeepFileSurvivesTheInstanceDrop", _
                         Err.Number, Err.Description
End Sub

'@sub-title Reset wipes the kept file and forgets the mark.
'@details
'Reset runs after a linelist is saved, so the folder has to start clean:
'the file goes, and a later drop of the instance takes a new file of the
'same name with it.
'@TestMethod("TemporaryRepos")
Public Sub TestResetForgetsTheKeptFile()
    CustomTestSetTitles Assert, "TemporaryRepos", "TestResetForgetsTheKeptFile"
    On Error GoTo TestFail

    Dim repos As TemporaryRepos
    Dim root As String

    Set repos = TemporaryRepos.Create(BaseFolder)
    root = repos.RootPath()
    WriteTextFile repos.CreateFilePath("__temp.xlsb"), "x"
    repos.KeepFile "__temp.xlsb"

    repos.Reset

    Assert.AreEqual CLng(0), FileCountIn(root), "Reset takes the kept file too"

    WriteTextFile repos.CreateFilePath("__temp.xlsb"), "y"
    Set repos = Nothing

    Assert.AreEqual CLng(0), FileCountIn(root), _
                    "After a reset the mark is gone and the drop takes the new file"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestResetForgetsTheKeptFile", Err.Number, Err.Description
End Sub

'@section Fixture helpers
'===============================================================================

'@sub-title Whether a folder exists on disk.
'@param folderPath String. The path to test.
'@return Boolean. True when the path is a folder.
Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim attr As Long
    Dim readError As Long

    If LenB(folderPath) = 0 Then Exit Function

    On Error Resume Next
        Err.Clear
        attr = GetAttr(folderPath)
        readError = Err.Number
    On Error GoTo 0

    If readError <> 0 Then Exit Function
    FolderExists = ((attr And vbDirectory) = vbDirectory)
End Function

'@sub-title Whether a file exists on disk.
'@param filePath String. The path to test.
'@return Boolean. True when the file is there.
Private Function FileExists(ByVal filePath As String) As Boolean
    Dim found As String

    If LenB(filePath) = 0 Then Exit Function

    On Error Resume Next
        found = Dir$(filePath)
    On Error GoTo 0

    FileExists = (LenB(found) > 0)
End Function

'@sub-title How many files sit in a folder.
'@details
'The whole listing is read before anything is counted, the same way the class
'reads it before deleting.
'@param folderPath String. The folder to count.
'@return Long. The number of files.
Private Function FileCountIn(ByVal folderPath As String) As Long
    Dim fileName As String
    Dim total As Long

    If Not FolderExists(folderPath) Then Exit Function

    On Error Resume Next
        fileName = Dir$(folderPath & "*")
        Do While LenB(fileName) > 0
            If fileName <> "." And fileName <> ".." Then total = total + 1
            fileName = Dir$
        Loop
    On Error GoTo 0

    FileCountIn = total
End Function

'@sub-title Write a small text file.
'@param filePath String. Where to write it.
'@param content String. What to write.
Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim handle As Integer

    handle = FreeFile
    Open filePath For Output As #handle
    Print #handle, content
    Close #handle
End Sub

'@sub-title Delete every file of the repository folder and the folder itself.
Private Sub RemoveRepoFolder()
    Dim root As String
    Dim fileName As String
    Dim names As Collection
    Dim entry As Variant

    If LenB(BaseFolder) = 0 Then Exit Sub

    root = BaseFolder & Application.PathSeparator & REPO_FOLDER_NAME & _
           Application.PathSeparator
    If Not FolderExists(root) Then Exit Sub

    Set names = New Collection

    On Error Resume Next
        fileName = Dir$(root & "*")
        Do While LenB(fileName) > 0
            If fileName <> "." And fileName <> ".." Then names.Add fileName
            fileName = Dir$
        Loop
    On Error GoTo 0

    For Each entry In names
        On Error Resume Next
            Kill root & CStr(entry)
        On Error GoTo 0
    Next entry

    On Error Resume Next
        RmDir root
    On Error GoTo 0
End Sub
