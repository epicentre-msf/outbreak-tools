Attribute VB_Name = "TestSetupRibbon"
Attribute VB_Description = "Tests for the three setup entry points a script can call"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the three setup entry points a script can call")

'@description
'Drives RunSetupExport, RunSetupImportFile and SetupLastSummary, the entry
'points Session 112 put at the foot of SetupRibbon so a script outside Excel can
'reach the setup without an IRibbonControl and without a picker.
'
'THIS WORKBOOK IS THE TEST DRIVER, NOT A SETUP
'-------------------------------------------------------------------------------
'All three wrappers work on ThisWorkbook, and here that is the driver: it holds
'no Dictionary, no Choices, no Analysis and no Translations sheet. So the suite
'drives the paths that are the same on any workbook -- the guards, the answer
'shape, the disarm, the summary -- and one export, which runs end to end and
'writes a workbook carrying no setup sheet.
'
'WHAT THIS SUITE DOES NOT DRIVE, AND WHY
'-------------------------------------------------------------------------------
'RunSetupImportFile is driven on its guard alone. Handed a real file it reads
'that file INTO ThisWorkbook, which here would rewrite the driver's own sheets
'mid-run.
'
'RunSetupTags is not driven at all. It opens with
'SetupPreparation.ResetUpdatedRegistry on ThisWorkbook, which drops and rebuilds
'the update registry of the workbook the run is happening in.
'
'Both are proved instead by the R package calling them, which is what the block
'these sessions belong to is for. A wrapper that opens a box hangs that call.
'
'EVERY CALL IN HERE ANSWERS A STRING
'-------------------------------------------------------------------------------
'That is the contract, and it is the one thing every test below checks: no path
'of a wrapper opens a box, and no path leaves the messenger armed for whoever
'calls next.
'@depends SetupRibbon, Messenger, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'A folder no machine running this has, so the export guard refuses it.
Private Const MISSING_FOLDER As String = "obt-no-such-folder-112"

'A file no machine running this has, so the import guard refuses it.
Private Const MISSING_FILE As String = "obt-no-such-file-112.xlsx"

'Where the one real export writes. Made under the run dir and emptied after.
Private Const EXPORT_FOLDER_NAME As String = "setup-ribbon-exports"

'What a wrapper answers when the run went through.
Private Const OUTCOME_OK As String = "OK"

'What an outcome that failed opens with.
Private Const OUTCOME_ERROR_LEAD As String = "ERROR "

'The marker the free text of the summary starts after.
Private Const REPORT_MARKER As String = "--report--"

'What the summary file is called, after the name of the file the run touched.
Private Const SUMMARY_SUFFIX As String = "-obt-summary.txt"

Private Assert As CustomTest


'@section Module lifecycle
'===============================================================================

'@sub-title Set up the assertion harness.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSetupRibbon"
End Sub

'@sub-title Print results and take away every file the suite wrote.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    ClearExportFolder
    RemoveWorkbookSummary
    Messenger.Reset
    Set Assert = Nothing
    RestoreApp
End Sub

'@sub-title Start every test with the boxes on.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Messenger.Reset
End Sub

'@sub-title Flush assert state and leave nothing armed.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    Messenger.Reset
End Sub


'@section Helper routines
'===============================================================================

'@sub-title A path under the run dir that nothing has made.
'@param leafName String. The name to hang off the run dir.
'@return String. The full path.
Private Function PathUnderRunDir(ByVal leafName As String) As String
    PathUnderRunDir = ThisWorkbook.Path & Application.PathSeparator & leafName
End Function

'@sub-title The folder the one real export writes into.
'@details
'This makes the folder itself rather than calling TestHelpersLite.BuildTempFolder.
'That helper decides whether to MkDir with Dir$(path, vbDirectory), which does not
'answer reliably for a folder on Mac Excel: the folder was never created, the
'wrapper refused a path that was not there, and three tests here failed together.
'The check below is GetAttr, the idiom HeadlessBuild and TemporaryRepos use.
'@return String. The folder path, made if it was not there. Empty when it could
'not be made, which fails the test that asked for it rather than passing quietly.
Private Function ExportFolder() As String
    Dim folderPath As String

    folderPath = PathUnderRunDir(EXPORT_FOLDER_NAME)

    If Not FolderIsThere(folderPath) Then
        On Error Resume Next
            MkDir folderPath
            Err.Clear
        On Error GoTo 0
    End If

    If FolderIsThere(folderPath) Then ExportFolder = folderPath
End Function

'@sub-title Whether a folder sits at that path.
'@param folderPath String. The path to look at.
'@return Boolean. True when the path names a folder that is there.
Private Function FolderIsThere(ByVal folderPath As String) As Boolean
    Dim folderAttributes As Long

    If LenB(folderPath) = 0 Then Exit Function

    On Error Resume Next
        folderAttributes = GetAttr(folderPath)
        If Err.Number = 0 Then
            FolderIsThere = ((folderAttributes And vbDirectory) = vbDirectory)
        End If
        Err.Clear
    On Error GoTo 0
End Function

'@sub-title Take every file the export folder holds, then the folder.
Private Sub ClearExportFolder()
    Dim folderPath As String
    Dim fileName As String

    folderPath = PathUnderRunDir(EXPORT_FOLDER_NAME)

    On Error Resume Next
        fileName = Dir$(folderPath & Application.PathSeparator & "*")
        Do While LenB(fileName) > 0
            Kill folderPath & Application.PathSeparator & fileName
            fileName = Dir$()
        Loop
        RmDir folderPath
    On Error GoTo 0
End Sub

'@sub-title Take away the summary a refused run wrote beside this workbook.
Private Sub RemoveWorkbookSummary()
    On Error Resume Next
        Kill PathUnderRunDir(BareName(ThisWorkbook.Name) & SUMMARY_SUFFIX)
    On Error GoTo 0
End Sub

'@sub-title The name of a file with its folder and its extension taken off.
'@param filePath String. A file name or a full path.
'@return String. The bare name.
Private Function BareName(ByVal filePath As String) As String
    Dim answer As String
    Dim cutAt As Long

    answer = filePath

    cutAt = InStrRev(answer, Application.PathSeparator)
    If cutAt > 0 Then answer = Mid$(answer, cutAt + 1)

    cutAt = InStrRev(answer, ".")
    If cutAt > 1 Then answer = Left$(answer, cutAt - 1)

    BareName = answer
End Function

'@sub-title Read one key off the summary of the last run.
'@param keyName String. The key, with no equals sign.
'@return String. What that key answers, empty when the key is not there.
Private Function SummaryValue(ByVal keyName As String) As String
    Dim lines As Variant
    Dim index As Long
    Dim lead As String

    lead = keyName & "="
    lines = Split(SetupRibbon.SetupLastSummary(), vbLf)

    For index = LBound(lines) To UBound(lines)
        If CStr(lines(index)) = REPORT_MARKER Then Exit Function
        If InStr(1, CStr(lines(index)), lead, vbBinaryCompare) = 1 Then
            SummaryValue = Mid$(CStr(lines(index)), Len(lead) + 1)
            Exit Function
        End If
    Next index
End Function

'@sub-title Everything the summary holds after the marker.
'@return String. The free text, empty when the run swallowed nothing.
Private Function SummaryReport() As String
    Dim summaryText As String
    Dim markerAt As Long

    summaryText = SetupRibbon.SetupLastSummary()
    markerAt = InStr(1, summaryText, vbLf & REPORT_MARKER & vbLf, vbBinaryCompare)
    If markerAt = 0 Then Exit Function

    SummaryReport = Mid$(summaryText, markerAt + Len(REPORT_MARKER) + 2)
End Function

'@sub-title Whether a file sits at that path.
'@param filePath String. The path to look at.
'@return Boolean. True when the file is there.
Private Function FileIsThere(ByVal filePath As String) As Boolean
    If LenB(filePath) = 0 Then Exit Function

    On Error Resume Next
        FileIsThere = (LenB(Dir$(filePath)) > 0)
    On Error GoTo 0
End Function

'@sub-title The first line of a text file.
'@param filePath String. The file to read.
'@return String. The first line, empty when the file cannot be read.
Private Function FirstLineOf(ByVal filePath As String) As String
    Dim fileNumber As Long
    Dim lineText As String

    fileNumber = FreeFile

    On Error GoTo CloseFile
    Open filePath For Input As #fileNumber
    If Not EOF(fileNumber) Then Line Input #fileNumber, lineText
    Close #fileNumber

    FirstLineOf = lineText
    Exit Function

CloseFile:
    On Error Resume Next
        Close #fileNumber
    On Error GoTo 0
End Function


'@section The guards
'===============================================================================
'A picker proved the file was there. A path from a script has proved nothing.

'@sub-title The export refuses a folder that is not there and starts no work.
'@TestMethod("SetupRibbon")
Public Sub TestExportRefusesAFolderThatIsNotThere()
    Dim missingFolder As String
    Dim outcome As String

    CustomTestSetTitles Assert, "SetupRibbon", "TestExportRefusesAFolderThatIsNotThere"
    On Error GoTo TestFail

    missingFolder = PathUnderRunDir(MISSING_FOLDER)
    outcome = SetupRibbon.RunSetupExport(missingFolder)

    Assert.AreEqual "ERROR 0: no folder at " & missingFolder, outcome, _
                    "The export names the folder it was given and the number 0"
    Assert.AreEqual vbNullString, SummaryValue("export"), _
                    "A refused export wrote no file"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestExportRefusesAFolderThatIsNotThere", Err.Number, Err.Description
End Sub

'@sub-title The import refuses a path that is not there and starts no work.
'@TestMethod("SetupRibbon")
Public Sub TestImportRefusesAPathThatIsNotThere()
    Dim missingFile As String
    Dim outcome As String

    CustomTestSetTitles Assert, "SetupRibbon", "TestImportRefusesAPathThatIsNotThere"
    On Error GoTo TestFail

    missingFile = PathUnderRunDir(MISSING_FILE)
    outcome = SetupRibbon.RunSetupImportFile(missingFile)

    Assert.AreEqual "ERROR 0: no file at " & missingFile, outcome, _
                    "The import names the file it was given and the number 0"
    Assert.AreEqual vbNullString, SummaryValue("imported"), _
                    "A refused import read no file"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestImportRefusesAPathThatIsNotThere", Err.Number, Err.Description
End Sub

'@sub-title The import refuses an empty path.
'@details
'A script that forgot the argument reaches this, and the answer has to say so
'rather than opening the picker the button opens.
'@TestMethod("SetupRibbon")
Public Sub TestImportRefusesAnEmptyPath()
    Dim outcome As String

    CustomTestSetTitles Assert, "SetupRibbon", "TestImportRefusesAnEmptyPath"
    On Error GoTo TestFail

    outcome = SetupRibbon.RunSetupImportFile(vbNullString)

    Assert.AreEqual "ERROR 0: no file at ", outcome, _
                    "An empty path is refused as a file that is not there"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestImportRefusesAnEmptyPath", Err.Number, Err.Description
End Sub


'@section The disarm
'===============================================================================
'A wrapper that left the messenger armed would swallow the boxes of whoever ran
'next, and that person is a human clicking a button.

'@sub-title A refused export leaves the boxes on.
'@TestMethod("SetupRibbon")
Public Sub TestARefusedExportLeavesTheBoxesOn()
    CustomTestSetTitles Assert, "SetupRibbon", "TestARefusedExportLeavesTheBoxesOn"
    On Error GoTo TestFail

    SetupRibbon.RunSetupExport PathUnderRunDir(MISSING_FOLDER)

    Assert.IsFalse Messenger.Armed, "The export disarmed on its guard path"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARefusedExportLeavesTheBoxesOn", Err.Number, Err.Description
End Sub

'@sub-title A refused import leaves the boxes on.
'@TestMethod("SetupRibbon")
Public Sub TestARefusedImportLeavesTheBoxesOn()
    CustomTestSetTitles Assert, "SetupRibbon", "TestARefusedImportLeavesTheBoxesOn"
    On Error GoTo TestFail

    SetupRibbon.RunSetupImportFile PathUnderRunDir(MISSING_FILE)

    Assert.IsFalse Messenger.Armed, "The import disarmed on its guard path"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARefusedImportLeavesTheBoxesOn", Err.Number, Err.Description
End Sub


'@section The summary
'===============================================================================

'@sub-title The summary answers its four keys and its marker, in order.
'@TestMethod("SetupRibbon")
Public Sub TestTheSummaryAnswersItsKeysAndItsMarker()
    Dim lines As Variant

    CustomTestSetTitles Assert, "SetupRibbon", "TestTheSummaryAnswersItsKeysAndItsMarker"
    On Error GoTo TestFail

    SetupRibbon.RunSetupExport PathUnderRunDir(MISSING_FOLDER)

    lines = Split(SetupRibbon.SetupLastSummary(), vbLf)

    Assert.IsTrue (UBound(lines) - LBound(lines) >= 4), _
                  "The summary holds four keys and the marker at least"
    Assert.AreEqual 1&, CLng(InStr(1, CStr(lines(LBound(lines))), "outcome=")), _
                    "outcome leads the summary, because it is what the lost answer held"
    Assert.AreEqual 1&, CLng(InStr(1, CStr(lines(LBound(lines) + 1)), "export=")), _
                    "export is the second key"
    Assert.AreEqual 1&, CLng(InStr(1, CStr(lines(LBound(lines) + 2)), "imported=")), _
                    "imported is the third key"
    Assert.AreEqual REPORT_MARKER, CStr(lines(LBound(lines) + 3)), _
                    "The marker closes the keys"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSummaryAnswersItsKeysAndItsMarker", Err.Number, Err.Description
End Sub

'@sub-title The summary carries the outcome the wrapper answered.
'@details
'This is the whole reason outcome leads the block: the answer of Application.Run
'is what the transport loses, and the outcome is the first thing that reading
'loses with it.
'@TestMethod("SetupRibbon")
Public Sub TestTheSummaryCarriesTheOutcome()
    Dim outcome As String

    CustomTestSetTitles Assert, "SetupRibbon", "TestTheSummaryCarriesTheOutcome"
    On Error GoTo TestFail

    outcome = SetupRibbon.RunSetupImportFile(PathUnderRunDir(MISSING_FILE))

    Assert.AreEqual outcome, SummaryValue("outcome"), _
                    "The summary answers the same outcome the call answered"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSummaryCarriesTheOutcome", Err.Number, Err.Description
End Sub

'@sub-title A run forgets the run before it.
'@details
'A script reading the summary after its second call must never be handed the
'paths of its first.
'@TestMethod("SetupRibbon")
Public Sub TestARunForgetsTheRunBeforeIt()
    Dim firstOutcome As String

    CustomTestSetTitles Assert, "SetupRibbon", "TestARunForgetsTheRunBeforeIt"
    On Error GoTo TestFail

    firstOutcome = SetupRibbon.RunSetupImportFile(PathUnderRunDir(MISSING_FILE))
    SetupRibbon.RunSetupExport PathUnderRunDir(MISSING_FOLDER)

    Assert.AreEqual vbNullString, SummaryValue("imported"), _
                    "The second run answers nothing for the file the first one named"
    Assert.IsFalse (SummaryValue("outcome") = firstOutcome), _
                   "The second run answers its own outcome"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARunForgetsTheRunBeforeIt", Err.Number, Err.Description
End Sub

'@sub-title A refused run still writes its summary beside this workbook.
'@details
'The file is the reading that survives a transport that gave up, so a run that
'failed needs it more than a run that worked.
'@TestMethod("SetupRibbon")
Public Sub TestARefusedRunStillWritesItsSummaryFile()
    Dim summaryPath As String
    Dim outcome As String

    CustomTestSetTitles Assert, "SetupRibbon", "TestARefusedRunStillWritesItsSummaryFile"
    On Error GoTo TestFail

    summaryPath = PathUnderRunDir(BareName(ThisWorkbook.Name) & SUMMARY_SUFFIX)
    RemoveWorkbookSummary

    outcome = SetupRibbon.RunSetupExport(PathUnderRunDir(MISSING_FOLDER))

    Assert.IsTrue FileIsThere(summaryPath), _
                  "A refused export left a summary file beside this workbook"
    Assert.AreEqual "outcome=" & outcome, FirstLineOf(summaryPath), _
                    "The file opens on the outcome the call answered"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestARefusedRunStillWritesItsSummaryFile", Err.Number, Err.Description
End Sub


'@section One run end to end
'===============================================================================
'The export is the one wrapper that can run whole on a workbook that is not a
'setup: it reads the setup worksheets it finds, and here it finds none, so it
'writes a workbook carrying nothing and saves it where it was told.

'@sub-title The export runs with no box and names the file it wrote.
'@TestMethod("SetupRibbon")
Public Sub TestTheExportRunsAndNamesItsFile()
    Dim targetFolder As String
    Dim outcome As String

    CustomTestSetTitles Assert, "SetupRibbon", "TestTheExportRunsAndNamesItsFile"
    On Error GoTo TestFail

    ClearExportFolder
    targetFolder = ExportFolder()
    Assert.IsTrue (LenB(targetFolder) > 0), "The export folder was made"

    outcome = SetupRibbon.RunSetupExport(targetFolder)

    Assert.AreEqual OUTCOME_OK, outcome, "The export ran through with no dialog"
    Assert.IsFalse Messenger.Armed, "The export disarmed on its way out"
    Assert.IsTrue (LenB(SummaryValue("export")) > 0), _
                  "The summary names the file the export wrote"
    Assert.IsTrue FileIsThere(SummaryValue("export")), _
                  "The file the summary names is on disk"

    ClearExportFolder
    Exit Sub
TestFail:
    ClearExportFolder
    CustomTestLogFailure Assert, "TestTheExportRunsAndNamesItsFile", Err.Number, Err.Description
End Sub

'@sub-title The box the export would have shown is readable in the summary.
'@details
'A run that ate a message and said nothing is worse than the box it replaced.
'@TestMethod("SetupRibbon")
Public Sub TestTheSwallowedBoxIsReadableAfterTheExport()
    Dim targetFolder As String

    CustomTestSetTitles Assert, "SetupRibbon", "TestTheSwallowedBoxIsReadableAfterTheExport"
    On Error GoTo TestFail

    ClearExportFolder
    targetFolder = ExportFolder()
    Assert.IsTrue (LenB(targetFolder) > 0), "The export folder was made"

    SetupRibbon.RunSetupExport targetFolder

    Assert.IsTrue (InStr(1, SummaryReport(), "Setup exported to:", vbTextCompare) > 0), _
                  "The box the button shows is written down under the marker"

    ClearExportFolder
    Exit Sub
TestFail:
    ClearExportFolder
    CustomTestLogFailure Assert, "TestTheSwallowedBoxIsReadableAfterTheExport", Err.Number, Err.Description
End Sub

'@sub-title The export writes its summary beside the file it wrote.
'@TestMethod("SetupRibbon")
Public Sub TestTheExportWritesItsSummaryBesideTheExport()
    Dim targetFolder As String
    Dim summaryPath As String

    CustomTestSetTitles Assert, "SetupRibbon", "TestTheExportWritesItsSummaryBesideTheExport"
    On Error GoTo TestFail

    ClearExportFolder
    targetFolder = ExportFolder()
    Assert.IsTrue (LenB(targetFolder) > 0), "The export folder was made"

    SetupRibbon.RunSetupExport targetFolder

    summaryPath = targetFolder & Application.PathSeparator & _
                  BareName(SummaryValue("export")) & SUMMARY_SUFFIX

    Assert.IsTrue FileIsThere(summaryPath), _
                  "The summary sits beside the export, named after it"
    Assert.AreEqual "outcome=" & OUTCOME_OK, FirstLineOf(summaryPath), _
                    "The file opens on the outcome"

    ClearExportFolder
    Exit Sub
TestFail:
    ClearExportFolder
    CustomTestLogFailure Assert, "TestTheExportWritesItsSummaryBesideTheExport", Err.Number, Err.Description
End Sub

'@sub-title Every answer is an outcome string, never a raise.
'@details
'The contract of the whole block: a script reads one string and decides. A
'wrapper that let a raise out would give Application.Run nothing to read.
'@TestMethod("SetupRibbon")
Public Sub TestEveryWrapperAnswersAnOutcomeString()
    Dim outcomes As Variant
    Dim index As Long
    Dim answer As String

    CustomTestSetTitles Assert, "SetupRibbon", "TestEveryWrapperAnswersAnOutcomeString"
    On Error GoTo TestFail

    outcomes = Array(SetupRibbon.RunSetupExport(PathUnderRunDir(MISSING_FOLDER)), _
                     SetupRibbon.RunSetupImportFile(PathUnderRunDir(MISSING_FILE)), _
                     SetupRibbon.RunSetupImportFile(vbNullString))

    For index = LBound(outcomes) To UBound(outcomes)
        answer = CStr(outcomes(index))
        Assert.IsTrue (answer = OUTCOME_OK Or _
                       InStr(1, answer, OUTCOME_ERROR_LEAD, vbBinaryCompare) = 1), _
                      "Answer " & CStr(index + 1) & " reads OK or ERROR"
    Next index

    Assert.IsFalse Messenger.Armed, "Three calls in a row left the boxes on"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEveryWrapperAnswersAnOutcomeString", Err.Number, Err.Description
End Sub
