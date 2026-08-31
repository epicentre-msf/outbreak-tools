Attribute VB_Name = "TestLinelistRun"
Attribute VB_Description = "Tests for the entry points a script can call on a running linelist"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the entry points a script can call on a running linelist")

'@description
'Drives RunImportGeobase, RunImportData, RunExport and LinelistLastSummary, the
'entry points Sessions 115 and 116 put in LinelistRun so a script outside Excel
'can reach an import or an export with no picker, no prompt and no box.
'
'THIS WORKBOOK IS THE TEST DRIVER, NOT A LINELIST
'-------------------------------------------------------------------------------
'It holds no dictionary, no translation sheet, no data entry table and no
'Exports sheet. So the suite drives the paths that are the same on any workbook
'-- the guards, the refusal when the linelist is not one, the disarm, the answer
'shape and the summary. The walks themselves are proved by the R package calling
'them against a real linelist.
'
'EVERY CALL IN HERE ANSWERS A STRING
'-------------------------------------------------------------------------------
'That is the contract, and it is the one thing every test below checks: no path
'of a wrapper opens a box, no path lets a raise out, and no path leaves the
'messenger armed for whoever calls next.
'
'AND EVERY CALL IN HERE FAILS, WHICH IS WHAT KEEPS THIS SUITE ALIVE
'-------------------------------------------------------------------------------
'A run that worked closes the workbook it ran in, and closing a workbook ends
'the code running inside it. That workbook here is the test driver holding this
'suite. Every call below is refused before it reaches the work -- a path or a
'folder that is not there, a word the wrapper does not know, a number that names
'no active export, or a workbook with no translation sheet -- so none of them
'reaches the close. A test written here that ever gets an OK out of a wrapper
'would take the whole run down with it.
'
'THE SERVICE IS PUT BACK AFTER THE MODULE
'-------------------------------------------------------------------------------
'LinelistRun reaches the linelist through LinelistEventsManager, which holds ONE
'EventLinelist for the whole session and builds it over ThisWorkbook on first
'read. This suite is what causes that build here, so ModuleCleanup calls
'DisposeEventLinelist and the suites after it start from nothing.
'@depends LinelistRun, LinelistEventsManager, Messenger, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'A file no machine running this has, so the two path guards refuse it.
Private Const MISSING_FILE As String = "obt-no-such-file-115.xlsx"

'Where the one real file this suite needs is made, and emptied after.
Private Const WORK_FOLDER_NAME As String = "linelist-run-files"

'A file that is really there, so a test reaches the guards past the path check.
Private Const PRESENT_FILE As String = "obt-present-115.xlsx"

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
    Assert.SetModuleName "TestLinelistRun"
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

    ClearWorkFolder
    RemoveWorkbookSummary
    Messenger.Reset
    LinelistEventsManager.DisposeEventLinelist
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

'@sub-title A path under the run dir.
'@param leafName String. The name to hang off the run dir.
'@return String. The full path.
Private Function PathUnderRunDir(ByVal leafName As String) As String
    PathUnderRunDir = ThisWorkbook.Path & Application.PathSeparator & leafName
End Function

'@sub-title The folder the one real file of this suite sits in.
'@details
'This makes the folder itself rather than calling TestHelpersLite.BuildTempFolder.
'That helper decides whether to MkDir with Dir$(path, vbDirectory), which does
'not answer reliably for a folder on Mac Excel. The check below is GetAttr, the
'idiom HeadlessBuild and TemporaryRepos use.
'@return String. The folder path, made if it was not there. Empty when it could
'not be made, which fails the test that asked for it rather than passing quietly.
Private Function WorkFolder() As String
    Dim folderPath As String

    folderPath = PathUnderRunDir(WORK_FOLDER_NAME)

    If Not FolderIsThere(folderPath) Then
        On Error Resume Next
            MkDir folderPath
            Err.Clear
        On Error GoTo 0
    End If

    If FolderIsThere(folderPath) Then WorkFolder = folderPath
End Function

'@sub-title A file that is really on disk, so a test reaches past the path guard.
'@details
'The wrappers never open it -- both stop at the translator, because this
'workbook carries no translation sheet -- so an empty file named .xlsx is enough.
'@return String. The full path, empty when the file could not be made.
Private Function PresentFile() As String
    Dim folderPath As String
    Dim filePath As String
    Dim fileNumber As Long

    folderPath = WorkFolder()
    If LenB(folderPath) = 0 Then Exit Function

    filePath = folderPath & Application.PathSeparator & PRESENT_FILE

    If Not FileIsThere(filePath) Then
        fileNumber = FreeFile
        On Error Resume Next
            Open filePath For Output As #fileNumber
            Close #fileNumber
            Err.Clear
        On Error GoTo 0
    End If

    If FileIsThere(filePath) Then PresentFile = filePath
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

'@sub-title Whether a file sits at that path.
'@param filePath String. The path to look at.
'@return Boolean. True when the file is there.
Private Function FileIsThere(ByVal filePath As String) As Boolean
    If LenB(filePath) = 0 Then Exit Function

    On Error Resume Next
        FileIsThere = (LenB(Dir$(filePath)) > 0)
    On Error GoTo 0
End Function

'@sub-title Take every file the work folder holds, then the folder.
Private Sub ClearWorkFolder()
    Dim folderPath As String
    Dim fileName As String

    folderPath = PathUnderRunDir(WORK_FOLDER_NAME)

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
    lines = Split(LinelistRun.LinelistLastSummary(), vbLf)

    For index = LBound(lines) To UBound(lines)
        If CStr(lines(index)) = REPORT_MARKER Then Exit Function
        If InStr(1, CStr(lines(index)), lead, vbBinaryCompare) = 1 Then
            SummaryValue = Mid$(CStr(lines(index)), Len(lead) + 1)
            Exit Function
        End If
    Next index
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

'@sub-title Whether an answer reads as a refusal.
'@param answer String. What a wrapper answered.
'@return Boolean. True when the answer opens with the error lead.
Private Function IsRefusal(ByVal answer As String) As Boolean
    IsRefusal = (InStr(1, answer, OUTCOME_ERROR_LEAD, vbBinaryCompare) = 1)
End Function


'@section The guards on the paths
'===============================================================================
'A picker proved the file was there. A path from a script has proved nothing.

'@sub-title The geobase import refuses a file that is not there and starts no work.
'@TestMethod("LinelistRun")
Public Sub TestGeobaseRefusesAFileThatIsNotThere()
    Dim missingFile As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestGeobaseRefusesAFileThatIsNotThere"
    On Error GoTo TestFail

    missingFile = PathUnderRunDir(MISSING_FILE)
    outcome = LinelistRun.RunImportGeobase(missingFile)

    Assert.AreEqual "ERROR 0: no file at " & missingFile, outcome, _
                    "The geobase import names the file it was given and the number 0"
    Assert.AreEqual vbNullString, SummaryValue("geobase"), _
                    "A refused geobase import read no file"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestGeobaseRefusesAFileThatIsNotThere", Err.Number, Err.Description
End Sub

'@sub-title The data import refuses a file that is not there and starts no work.
'@TestMethod("LinelistRun")
Public Sub TestImportRefusesAFileThatIsNotThere()
    Dim missingFile As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestImportRefusesAFileThatIsNotThere"
    On Error GoTo TestFail

    missingFile = PathUnderRunDir(MISSING_FILE)
    outcome = LinelistRun.RunImportData(missingFile, "append")

    Assert.AreEqual "ERROR 0: no file at " & missingFile, outcome, _
                    "The data import names the file it was given and the number 0"
    Assert.AreEqual vbNullString, SummaryValue("import"), _
                    "A refused import read no file"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestImportRefusesAFileThatIsNotThere", Err.Number, Err.Description
End Sub

'@sub-title Both wrappers refuse an empty path.
'@details
'A script that forgot the argument reaches this, and the answer has to say so
'rather than opening the picker the button opens.
'@TestMethod("LinelistRun")
Public Sub TestBothWrappersRefuseAnEmptyPath()
    CustomTestSetTitles Assert, "LinelistRun", "TestBothWrappersRefuseAnEmptyPath"
    On Error GoTo TestFail

    Assert.AreEqual "ERROR 0: no file at ", LinelistRun.RunImportGeobase(vbNullString), _
                    "An empty geobase path is refused as a file that is not there"
    Assert.AreEqual "ERROR 0: no file at ", LinelistRun.RunImportData(vbNullString, "append"), _
                    "An empty import path is refused as a file that is not there"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestBothWrappersRefuseAnEmptyPath", Err.Number, Err.Description
End Sub


'@section The guards on the two words
'===============================================================================
'A word the wrapper does not know is refused rather than guessed at. The pasting
'rule decides whether the rows a user has entered are wiped, and force decides
'whether a warning nobody can read is pushed past.

'@sub-title A pasting rule that is neither word is refused, and the answer says so.
'@TestMethod("LinelistRun")
Public Sub TestImportRefusesAPastingRuleItDoesNotKnow()
    Dim filePath As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestImportRefusesAPastingRuleItDoesNotKnow"
    On Error GoTo TestFail

    filePath = PresentFile()
    Assert.IsTrue (LenB(filePath) > 0), "The file the guard reads past was made"

    outcome = LinelistRun.RunImportData(filePath, "top")

    Assert.AreEqual "ERROR 0: pasting rule ""top"" is neither replace nor append", _
                    outcome, "The refusal names the word it was given and both words it takes"
    Assert.AreEqual vbNullString, SummaryValue("import"), _
                    "A refused rule imported nothing"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestImportRefusesAPastingRuleItDoesNotKnow", Err.Number, Err.Description
End Sub

'@sub-title An empty pasting rule means append, so it reads past the rule guard.
'@details
'Append is the reading that keeps the rows the linelist already holds, which is
'why it is what an absent word means.
'@TestMethod("LinelistRun")
Public Sub TestAnEmptyPastingRuleIsAppend()
    Dim filePath As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestAnEmptyPastingRuleIsAppend"
    On Error GoTo TestFail

    filePath = PresentFile()
    Assert.IsTrue (LenB(filePath) > 0), "The file the guard reads past was made"

    outcome = LinelistRun.RunImportData(filePath, vbNullString)

    Assert.AreEqual 0&, CLng(InStr(1, outcome, "pasting rule", vbTextCompare)), _
                    "An empty rule is never refused as a rule"
    Assert.IsTrue IsRefusal(outcome), _
                  "It stops further on, because this workbook is not a linelist"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestAnEmptyPastingRuleIsAppend", Err.Number, Err.Description
End Sub

'@sub-title A force word that is neither Yes nor No is refused.
'@details
'A typo must never quietly read as No. The caller has to ask for force by name,
'and a word the wrapper does not know means it cannot tell whether they did.
'@TestMethod("LinelistRun")
Public Sub TestImportRefusesAForceWordItDoesNotKnow()
    Dim filePath As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestImportRefusesAForceWordItDoesNotKnow"
    On Error GoTo TestFail

    filePath = PresentFile()
    Assert.IsTrue (LenB(filePath) > 0), "The file the guard reads past was made"

    outcome = LinelistRun.RunImportData(filePath, "append", "maybe")

    Assert.AreEqual "ERROR 0: force ""maybe"" is neither Yes nor No", outcome, _
                    "The refusal names the word it was given and both words it takes"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestImportRefusesAForceWordItDoesNotKnow", Err.Number, Err.Description
End Sub

'@sub-title Both force words are read, and neither is refused.
'@TestMethod("LinelistRun")
Public Sub TestBothForceWordsAreRead()
    Dim filePath As String

    CustomTestSetTitles Assert, "LinelistRun", "TestBothForceWordsAreRead"
    On Error GoTo TestFail

    filePath = PresentFile()
    Assert.IsTrue (LenB(filePath) > 0), "The file the guard reads past was made"

    Assert.AreEqual 0&, _
                    CLng(InStr(1, LinelistRun.RunImportData(filePath, "append", "Yes"), _
                               "force", vbTextCompare)), _
                    "Yes reads past the force guard"
    Assert.AreEqual 0&, _
                    CLng(InStr(1, LinelistRun.RunImportData(filePath, "append", "No"), _
                               "force", vbTextCompare)), _
                    "No reads past the force guard"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestBothForceWordsAreRead", Err.Number, Err.Description
End Sub


'@section A workbook that is not a linelist
'===============================================================================
'The driver holds no translation sheet, so both wrappers refuse it. What matters
'is that they refuse it with a string.

'@sub-title A workbook that is not a linelist is refused, with no box and no raise.
'@TestMethod("LinelistRun")
Public Sub TestAWorkbookThatIsNotALinelistIsRefused()
    Dim filePath As String
    Dim importOutcome As String
    Dim geobaseOutcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestAWorkbookThatIsNotALinelistIsRefused"
    On Error GoTo TestFail

    filePath = PresentFile()
    Assert.IsTrue (LenB(filePath) > 0), "The file the guard reads past was made"

    importOutcome = LinelistRun.RunImportData(filePath, "append")
    geobaseOutcome = LinelistRun.RunImportGeobase(filePath)

    Assert.IsTrue IsRefusal(importOutcome), "The data import refused this workbook"
    Assert.IsTrue IsRefusal(geobaseOutcome), "The geobase import refused this workbook"
    Assert.IsFalse Messenger.Armed, "Neither refusal left the boxes off"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestAWorkbookThatIsNotALinelistIsRefused", Err.Number, Err.Description
End Sub


'@section The disarm
'===============================================================================
'A wrapper that left the messenger armed would swallow the boxes of whoever ran
'next, and that person is a human clicking a button.

'@sub-title A refused run leaves the workbook open.
'@details
'A run that imported saves this workbook and closes it, and closing a workbook
'ends the code inside it. A refused run must not, or a script could never look
'at a linelist whose import it got wrong -- and this suite could never run a
'second test.
'@TestMethod("LinelistRun")
Public Sub TestARefusedRunLeavesTheWorkbookOpen()
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestARefusedRunLeavesTheWorkbookOpen"
    On Error GoTo TestFail

    outcome = LinelistRun.RunImportData(PathUnderRunDir(MISSING_FILE), "append")

    Assert.IsTrue IsRefusal(outcome), "The run was refused"
    Assert.IsTrue (LenB(ThisWorkbook.Name) > 0), _
                  "The workbook is still open, which is why this line runs at all"
    Assert.IsFalse Messenger.Armed, "And the refusal disarmed on its way out"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestARefusedRunLeavesTheWorkbookOpen", Err.Number, Err.Description
End Sub

'@sub-title A refused geobase import leaves the boxes on.
'@TestMethod("LinelistRun")
Public Sub TestARefusedGeobaseLeavesTheBoxesOn()
    CustomTestSetTitles Assert, "LinelistRun", "TestARefusedGeobaseLeavesTheBoxesOn"
    On Error GoTo TestFail

    LinelistRun.RunImportGeobase PathUnderRunDir(MISSING_FILE)

    Assert.IsFalse Messenger.Armed, "The geobase import disarmed on its guard path"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestARefusedGeobaseLeavesTheBoxesOn", Err.Number, Err.Description
End Sub

'@sub-title A refused data import leaves the boxes on, force or no force.
'@details
'Disarm drops the force state with the silence, so CarryOn answers vbNo again
'after a forced run. A run that left force on would push the NEXT caller past
'the three warnings it is meant to stop at.
'@TestMethod("LinelistRun")
Public Sub TestARefusedImportLeavesTheBoxesOnAndTheForceOff()
    CustomTestSetTitles Assert, "LinelistRun", "TestARefusedImportLeavesTheBoxesOnAndTheForceOff"
    On Error GoTo TestFail

    LinelistRun.RunImportData PathUnderRunDir(MISSING_FILE), "append", "Yes"

    Assert.IsFalse Messenger.Armed, "The data import disarmed on its guard path"
    Assert.AreEqual CLng(vbNo), CLng(Messenger.CarryOn()), _
                    "The force a run asked for went with the disarm"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestARefusedImportLeavesTheBoxesOnAndTheForceOff", Err.Number, Err.Description
End Sub


'@section The summary
'===============================================================================

'@sub-title The summary answers its four keys and its marker, in order.
'@TestMethod("LinelistRun")
Public Sub TestTheSummaryAnswersItsKeysAndItsMarker()
    Dim lines As Variant

    CustomTestSetTitles Assert, "LinelistRun", "TestTheSummaryAnswersItsKeysAndItsMarker"
    On Error GoTo TestFail

    LinelistRun.RunImportGeobase PathUnderRunDir(MISSING_FILE)

    lines = Split(LinelistRun.LinelistLastSummary(), vbLf)

    Assert.IsTrue (UBound(lines) - LBound(lines) >= 4), _
                  "The summary holds four keys and the marker at least"
    Assert.AreEqual 1&, CLng(InStr(1, CStr(lines(LBound(lines))), "outcome=")), _
                    "outcome leads the summary, because it is what the lost answer held"
    Assert.AreEqual 1&, CLng(InStr(1, CStr(lines(LBound(lines) + 1)), "geobase=")), _
                    "geobase is the second key"
    Assert.AreEqual 1&, CLng(InStr(1, CStr(lines(LBound(lines) + 2)), "import=")), _
                    "import is the third key"
    Assert.AreEqual 1&, CLng(InStr(1, CStr(lines(LBound(lines) + 3)), "export=")), _
                    "export is the fourth key"
    Assert.AreEqual REPORT_MARKER, CStr(lines(LBound(lines) + 4)), _
                    "The marker closes the keys"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestTheSummaryAnswersItsKeysAndItsMarker", Err.Number, Err.Description
End Sub

'@sub-title The summary carries the outcome the wrapper answered.
'@details
'This is the whole reason outcome leads the block: the answer of Application.Run
'is what the transport loses, and the outcome is the first thing that reading
'loses with it.
'@TestMethod("LinelistRun")
Public Sub TestTheSummaryCarriesTheOutcome()
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestTheSummaryCarriesTheOutcome"
    On Error GoTo TestFail

    outcome = LinelistRun.RunImportData(PathUnderRunDir(MISSING_FILE), "append")

    Assert.AreEqual outcome, SummaryValue("outcome"), _
                    "The summary answers the same outcome the call answered"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestTheSummaryCarriesTheOutcome", Err.Number, Err.Description
End Sub

'@sub-title A run forgets the run before it.
'@details
'A script reading the summary after its second call must never be handed the
'answer of its first.
'@TestMethod("LinelistRun")
Public Sub TestARunForgetsTheRunBeforeIt()
    Dim firstOutcome As String
    Dim filePath As String

    CustomTestSetTitles Assert, "LinelistRun", "TestARunForgetsTheRunBeforeIt"
    On Error GoTo TestFail

    filePath = PresentFile()
    Assert.IsTrue (LenB(filePath) > 0), "The file the guard reads past was made"

    firstOutcome = LinelistRun.RunImportData(filePath, "top")
    LinelistRun.RunImportGeobase PathUnderRunDir(MISSING_FILE)

    Assert.IsFalse (SummaryValue("outcome") = firstOutcome), _
                   "The second run answers its own outcome"
    Assert.AreEqual vbNullString, SummaryValue("import"), _
                    "The second run answers nothing for the import of the first"

    ClearWorkFolder
    RemoveWorkbookSummary
    Exit Sub
TestFail:
    ClearWorkFolder
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestARunForgetsTheRunBeforeIt", Err.Number, Err.Description
End Sub

'@sub-title A refused run still writes its summary beside this workbook.
'@details
'The file is the reading that survives a transport that gave up, so a run that
'failed needs it more than a run that worked.
'@TestMethod("LinelistRun")
Public Sub TestARefusedRunStillWritesItsSummaryFile()
    Dim summaryPath As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestARefusedRunStillWritesItsSummaryFile"
    On Error GoTo TestFail

    summaryPath = PathUnderRunDir(BareName(ThisWorkbook.Name) & SUMMARY_SUFFIX)
    RemoveWorkbookSummary

    outcome = LinelistRun.RunImportGeobase(PathUnderRunDir(MISSING_FILE))

    Assert.IsTrue FileIsThere(summaryPath), _
                  "A refused geobase import left a summary file beside this workbook"
    Assert.AreEqual "outcome=" & outcome, FirstLineOf(summaryPath), _
                    "The file opens on the outcome the call answered"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestARefusedRunStillWritesItsSummaryFile", Err.Number, Err.Description
End Sub

'@sub-title The summary answers paths and no counts.
'@details
'How many variables and how many sheets a linelist holds belongs to the
'generation run, which keeps its own record and answers it through
'DesignerLastSummary. This block has four keys and it keeps four.
'@TestMethod("LinelistRun")
Public Sub TestTheSummaryAnswersPathsAndNoCounts()
    Dim lines As Variant
    Dim index As Long
    Dim keyCount As Long

    CustomTestSetTitles Assert, "LinelistRun", "TestTheSummaryAnswersPathsAndNoCounts"
    On Error GoTo TestFail

    LinelistRun.RunImportGeobase PathUnderRunDir(MISSING_FILE)

    lines = Split(LinelistRun.LinelistLastSummary(), vbLf)

    For index = LBound(lines) To UBound(lines)
        If CStr(lines(index)) = REPORT_MARKER Then Exit For
        keyCount = keyCount + 1
    Next index

    Assert.AreEqual 4&, keyCount, "The summary holds four keys and no fifth"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestTheSummaryAnswersPathsAndNoCounts", Err.Number, Err.Description
End Sub

'@sub-title Every answer is an outcome string, never a raise.
'@details
'The contract of the whole block: a script reads one string and decides. A
'wrapper that let a raise out would give Application.Run nothing to read.
'@TestMethod("LinelistRun")
Public Sub TestEveryWrapperAnswersAnOutcomeString()
    Dim outcomes As Variant
    Dim index As Long
    Dim answer As String
    Dim filePath As String

    CustomTestSetTitles Assert, "LinelistRun", "TestEveryWrapperAnswersAnOutcomeString"
    On Error GoTo TestFail

    filePath = PresentFile()
    Assert.IsTrue (LenB(filePath) > 0), "The file the guard reads past was made"

    outcomes = Array(LinelistRun.RunImportGeobase(PathUnderRunDir(MISSING_FILE)), _
                     LinelistRun.RunImportGeobase(vbNullString), _
                     LinelistRun.RunImportGeobase(filePath), _
                     LinelistRun.RunImportData(PathUnderRunDir(MISSING_FILE), "append"), _
                     LinelistRun.RunImportData(filePath, "top"), _
                     LinelistRun.RunImportData(filePath, "append", "maybe"), _
                     LinelistRun.RunImportData(filePath, "replace", "Yes"))

    For index = LBound(outcomes) To UBound(outcomes)
        answer = CStr(outcomes(index))
        Assert.IsTrue (answer = OUTCOME_OK Or IsRefusal(answer)), _
                      "Answer " & CStr(index + 1) & " reads OK or ERROR"
    Next index

    Assert.IsFalse Messenger.Armed, "Seven calls in a row left the boxes on"

    ClearWorkFolder
    RemoveWorkbookSummary
    Exit Sub
TestFail:
    ClearWorkFolder
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestEveryWrapperAnswersAnOutcomeString", Err.Number, Err.Description
End Sub

'@section The guards of the export wrapper
'===============================================================================
'RunExport takes a word rather than five checkboxes, and a folder rather than a
'picker. Every guard below refuses before the export starts, which is what keeps
'the driver workbook open and this suite running.

'@sub-title The export refuses a folder that is not there.
'@details
'The button picked the folder, so it was there. A folder from a script has
'proved nothing, and an export writing into a folder Excel cannot reach fails
'deep inside LLExporter instead of here.
'@TestMethod("LinelistRun")
Public Sub TestExportRefusesAFolderThatIsNotThere()
    Dim missingFolder As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestExportRefusesAFolderThatIsNotThere"
    On Error GoTo TestFail

    missingFolder = PathUnderRunDir("obt-no-such-folder-116")
    outcome = LinelistRun.RunExport(vbNullString, missingFolder)

    Assert.AreEqual "ERROR 0: no folder at " & missingFolder, outcome, _
                    "The export names the folder it was given and the number 0"
    Assert.AreEqual vbNullString, SummaryValue("export"), _
                    "A refused export wrote no file"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestExportRefusesAFolderThatIsNotThere", Err.Number, Err.Description
End Sub

'@sub-title The export refuses an empty folder.
'@TestMethod("LinelistRun")
Public Sub TestExportRefusesAnEmptyFolder()
    CustomTestSetTitles Assert, "LinelistRun", "TestExportRefusesAnEmptyFolder"
    On Error GoTo TestFail

    Assert.AreEqual "ERROR 0: no folder at ", LinelistRun.RunExport("migration", vbNullString), _
                    "An empty folder is refused as a folder that is not there"
    Assert.IsFalse Messenger.Armed, "The refused export left the boxes on"

    RemoveWorkbookSummary
    Exit Sub
TestFail:
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestExportRefusesAnEmptyFolder", Err.Number, Err.Description
End Sub

'@sub-title A word that names no export is refused and read back to the caller.
'@details
'The five checkboxes of F_ExportMig are a word here, so a typo has to be named
'rather than run as the migration export the empty word means.
'@TestMethod("LinelistRun")
Public Sub TestExportRefusesAWordItDoesNotKnow()
    Dim folderPath As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestExportRefusesAWordItDoesNotKnow"
    On Error GoTo TestFail

    folderPath = WorkFolder()
    Assert.IsTrue (LenB(folderPath) > 0), "The folder the guard reads past was made"

    outcome = LinelistRun.RunExport("geobase", folderPath)

    Assert.AreEqual "ERROR 0: export ""geobase"" is no export this linelist runs", outcome, _
                    "The refusal reads the word back to the caller"
    Assert.IsFalse Messenger.Armed, "The refused export left the boxes on"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestExportRefusesAWordItDoesNotKnow", Err.Number, Err.Description
End Sub

'@sub-title The four words and a number are all read, and no other word is.
'@details
'The empty word means the migration export, so it is read like the four. Each
'of these gets past the word guard and stops at the guard after it, which is
'what the refusal text says.
'@TestMethod("LinelistRun")
Public Sub TestTheFourWordsAndANumberAreRead()
    Dim folderPath As String
    Dim words As Variant
    Dim index As Long
    Dim answer As String

    CustomTestSetTitles Assert, "LinelistRun", "TestTheFourWordsAndANumberAreRead"
    On Error GoTo TestFail

    folderPath = WorkFolder()
    Assert.IsTrue (LenB(folderPath) > 0), "The folder the guard reads past was made"

    words = Array(vbNullString, "migration", "geo", "historic", "analysis", "7")

    For index = LBound(words) To UBound(words)
        answer = LinelistRun.RunExport(CStr(words(index)), folderPath)
        Assert.IsFalse (InStr(1, answer, "is no export this linelist runs", vbBinaryCompare) > 0), _
                       "Word " & CStr(index + 1) & " got past the word guard"
    Next index

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestTheFourWordsAndANumberAreRead", Err.Number, Err.Description
End Sub

'@sub-title A number that names no active export is refused by its number.
'@details
'This workbook carries no Exports sheet, so every number is refused here. The
'refusal has to name the number: running export one in place of an export seven
'the caller asked for would write the wrong file and answer OK.
'@TestMethod("LinelistRun")
Public Sub TestExportRefusesANumberThatIsNotActive()
    Dim folderPath As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestExportRefusesANumberThatIsNotActive"
    On Error GoTo TestFail

    folderPath = WorkFolder()
    Assert.IsTrue (LenB(folderPath) > 0), "The folder the guard reads past was made"

    outcome = LinelistRun.RunExport("7", folderPath)

    Assert.AreEqual "ERROR 0: export number 7 is not an active export on the Exports sheet", _
                    outcome, "The refusal names the number the caller asked for"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestExportRefusesANumberThatIsNotActive", Err.Number, Err.Description
End Sub

'@sub-title A word with a comma in it is no export number.
'@details
'This box reads as en_FR, where the decimal separator is a COMMA. IsNumeric
'answers True for "3,5" here and Val reads 3 out of it, so a caller with a typo
'would have got export number three. Every character is checked against the
'digits instead.
'@TestMethod("LinelistRun")
Public Sub TestOnlyDigitsReadAsAnExportNumber()
    Dim folderPath As String

    CustomTestSetTitles Assert, "LinelistRun", "TestOnlyDigitsReadAsAnExportNumber"
    On Error GoTo TestFail

    folderPath = WorkFolder()
    Assert.IsTrue (LenB(folderPath) > 0), "The folder the guard reads past was made"

    Assert.AreEqual "ERROR 0: export ""3,5"" is no export this linelist runs", _
                    LinelistRun.RunExport("3,5", folderPath), _
                    "A comma makes the word no export number at all"
    Assert.AreEqual "ERROR 0: export ""0"" is no export this linelist runs", _
                    LinelistRun.RunExport("0", folderPath), _
                    "Export numbers start at one, so zero names none"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestOnlyDigitsReadAsAnExportNumber", Err.Number, Err.Description
End Sub

'@sub-title Only the three migration files take another linelist.
'@details
'The analysis export and the custom exports read the running linelist and
'nothing else. A caller pairing one of them with a file is told so, rather than
'having the file quietly ignored and the wrong workbook exported.
'@TestMethod("LinelistRun")
Public Sub TestOnlyTheMigrationWordsTakeAnotherLinelist()
    Dim folderPath As String
    Dim otherFile As String

    CustomTestSetTitles Assert, "LinelistRun", "TestOnlyTheMigrationWordsTakeAnotherLinelist"
    On Error GoTo TestFail

    folderPath = WorkFolder()
    otherFile = PresentFile()
    Assert.IsTrue (LenB(otherFile) > 0), "The file the test hands in was made"

    Assert.AreEqual "ERROR 0: the analysis export reads this linelist only, " & _
                    "so it takes no other linelist", _
                    LinelistRun.RunExport("analysis", folderPath, vbNullString, otherFile), _
                    "The analysis export refuses another linelist"
    Assert.AreEqual "ERROR 0: the custom export reads this linelist only, " & _
                    "so it takes no other linelist", _
                    LinelistRun.RunExport("2", folderPath, vbNullString, otherFile), _
                    "A custom export refuses another linelist"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestOnlyTheMigrationWordsTakeAnotherLinelist", Err.Number, Err.Description
End Sub

'@sub-title A workbook that is not a linelist cannot be exported either.
'@details
'The three import wrappers stop at the translator and so does this one. The
'workbook holding this suite has no translation sheet, which is what makes the
'test possible: a run that reached the export would close the workbook.
'@TestMethod("LinelistRun")
Public Sub TestTheExportRefusesAWorkbookThatIsNotALinelist()
    Dim folderPath As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestTheExportRefusesAWorkbookThatIsNotALinelist"
    On Error GoTo TestFail

    folderPath = WorkFolder()
    Assert.IsTrue (LenB(folderPath) > 0), "The folder the guard reads past was made"

    outcome = LinelistRun.RunExport("migration", folderPath)

    Assert.AreEqual "ERROR 0: this linelist carries no usable translation sheet", outcome, _
                    "A workbook with no translation sheet is refused by name"
    Assert.IsTrue (LenB(ThisWorkbook.Name) > 0), "The workbook is still open"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestTheExportRefusesAWorkbookThatIsNotALinelist", Err.Number, Err.Description
End Sub

'@sub-title A refused export leaves the boxes on for whoever calls next.
'@details
'The caller after a wrapper is a person clicking a button, and a linelist whose
'boxes stayed swallowed answers none of their questions.
'@TestMethod("LinelistRun")
Public Sub TestARefusedExportLeavesTheBoxesOn()
    Dim folderPath As String
    Dim ignored As String

    CustomTestSetTitles Assert, "LinelistRun", "TestARefusedExportLeavesTheBoxesOn"
    On Error GoTo TestFail

    folderPath = WorkFolder()

    ignored = LinelistRun.RunExport("migration", folderPath)
    Assert.IsFalse Messenger.Armed, "The export that reached the translator disarmed"

    ignored = LinelistRun.RunExport("nonsense", folderPath)
    Assert.IsFalse Messenger.Armed, "The export refused at its word guard disarmed"

    ignored = LinelistRun.RunExport(vbNullString, vbNullString)
    Assert.IsFalse Messenger.Armed, "The export refused at its folder guard disarmed"

    ClearWorkFolder
    RemoveWorkbookSummary
    Exit Sub
TestFail:
    ClearWorkFolder
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestARefusedExportLeavesTheBoxesOn", Err.Number, Err.Description
End Sub

'@sub-title A refused export writes its summary into the folder it was given.
'@details
'The answer of Application.Run is lost whenever the Apple Event transport gives
'up, so the file is the reading that survives. It is named after this linelist
'rather than after a file the run touched, because one export word can write
'three files and no one of them is the file.
'@TestMethod("LinelistRun")
Public Sub TestARefusedExportWritesItsSummaryFile()
    Dim folderPath As String
    Dim summaryPath As String
    Dim outcome As String

    CustomTestSetTitles Assert, "LinelistRun", "TestARefusedExportWritesItsSummaryFile"
    On Error GoTo TestFail

    folderPath = WorkFolder()
    Assert.IsTrue (LenB(folderPath) > 0), "The folder the summary goes into was made"

    outcome = LinelistRun.RunExport("migration", folderPath)

    summaryPath = folderPath & Application.PathSeparator & _
                  BareName(ThisWorkbook.Name) & SUMMARY_SUFFIX

    Assert.IsTrue FileIsThere(summaryPath), "The summary file is beside the export folder"
    Assert.AreEqual "outcome=" & outcome, FirstLineOf(summaryPath), _
                    "The first line of the file carries the outcome"

    ClearWorkFolder
    Exit Sub
TestFail:
    ClearWorkFolder
    CustomTestLogFailure Assert, "TestARefusedExportWritesItsSummaryFile", Err.Number, Err.Description
End Sub

'@sub-title The export wrapper answers an outcome string on every path.
'@TestMethod("LinelistRun")
Public Sub TestTheExportWrapperAnswersAnOutcomeString()
    Dim folderPath As String
    Dim otherFile As String
    Dim outcomes As Variant
    Dim index As Long
    Dim answer As String

    CustomTestSetTitles Assert, "LinelistRun", "TestTheExportWrapperAnswersAnOutcomeString"
    On Error GoTo TestFail

    folderPath = WorkFolder()
    otherFile = PresentFile()
    Assert.IsTrue (LenB(otherFile) > 0), "The file the test hands in was made"

    outcomes = Array(LinelistRun.RunExport(vbNullString, vbNullString), _
                     LinelistRun.RunExport(vbNullString, PathUnderRunDir("obt-no-such-folder-116")), _
                     LinelistRun.RunExport("nonsense", folderPath), _
                     LinelistRun.RunExport("9", folderPath), _
                     LinelistRun.RunExport("analysis", folderPath, vbNullString, otherFile), _
                     LinelistRun.RunExport("geo", folderPath), _
                     LinelistRun.RunExport("historic", folderPath, "a password", otherFile))

    For index = LBound(outcomes) To UBound(outcomes)
        answer = CStr(outcomes(index))
        Assert.IsTrue (answer = OUTCOME_OK Or IsRefusal(answer)), _
                      "Answer " & CStr(index + 1) & " reads OK or ERROR"
    Next index

    Assert.IsFalse Messenger.Armed, "Seven exports in a row left the boxes on"
    Assert.AreEqual vbNullString, SummaryValue("export"), _
                    "No refused export wrote a path into the summary"

    ClearWorkFolder
    RemoveWorkbookSummary
    Exit Sub
TestFail:
    ClearWorkFolder
    RemoveWorkbookSummary
    CustomTestLogFailure Assert, "TestTheExportWrapperAnswersAnOutcomeString", Err.Number, Err.Description
End Sub
