Attribute VB_Name = "TestHeadlessLinelistBuild"
Attribute VB_Description = "Builds one real linelist from a filled setup, headless, and reads back what landed"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Builds one real linelist from a filled setup, headless, and reads back what landed")

Option Explicit

'@description
'Step two of the headless workflow, under test: a whole linelist is generated
'from a filled setup with no designer machinery -- no ribbon press, no
'DesignerEntry, no form and no dialog.
'
'WHY THIS SUITE EXISTS
'-------------------------------------------------------------------------------
'Linelist.Prepare had never been run end to end by anything. Every suite that
'touched it stopped short of the code transfer, because a transfer needs a
'source project carrying the ten forms and the driver workbook carries none.
'So the one step that decides whether a delivered linelist can do anything at
'all -- moving 55 VBA components into it -- was measured by nothing, and a
'green run said nothing about it.
'
'This module closes that. It drives HeadlessBuild over a designer copy whose
'code has been replaced from src/, and then REOPENS THE BUILT FILE and counts
'what is inside its project. A build that writes a beautiful workbook holding
'no code passes every other check and fails here.
'
'THE FOUR FILES IT NEEDS
'-------------------------------------------------------------------------------
'  .mock/designer_mock.xlsb            the designer-shaped worksheets
'  src/bin/setup/setup_dev.xlsb        the setup to fill, copied per run
'  src/bin/test-files/generic-test-setup.xlsb   what fills it
'  ribbons/_ribbontemplate_dev.xlsb    the ribbon template of the output
'
'plus the merged forms in .test-runner/forms/merged, which are NOT a checked-in
'artefact: they are written by
'
'    Rscript scripts/headless/merge-form-code.R
'
'which puts the current FormLogic* code into the exported .frm files. Run it
'before the harness. A missing folder is reported as a failure naming that
'command rather than a raise, because the forms carry stale code without it and
'a linelist built from stale forms is worse than no linelist.
'
'THE FILLED SETUP IS BUILT HERE WHEN IT IS ABSENT
'-------------------------------------------------------------------------------
'TestHeadlessSetupFill writes the same file, and the harness gives no ordering
'between modules. So this one fills the setup itself when nothing is on disk,
'and reuses what is there otherwise. Neither suite depends on the other having
'run.
'
'THE PATHS ARE ABSOLUTE, AND THAT IS A KNOWN LIMIT
'-------------------------------------------------------------------------------
'A test module runs from the staging copy of the driver workbook, so a relative
'path resolves against the run directory rather than the repository. REPO_ROOT
'is the one place that assumption lives; HeadlessBuild itself takes every path
'as a parameter and holds none.
'@depends HeadlessBuild, CustomTest

Private Assert As CustomTest

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Const REPO_ROOT As String = _
    "/Users/komlaviamevoin/Unsync-Working-Folders/outbreak-tools"

Private Const DESIGNER_FILE As String = REPO_ROOT & "/.mock/designer_mock.xlsb"
Private Const EMPTY_SETUP As String = REPO_ROOT & "/src/bin/setup/setup_dev.xlsb"
Private Const GENERIC_SETUP As String = REPO_ROOT & "/src/bin/test-files/generic-test-setup.xlsb"
Private Const INJECTED_MODULE As String = REPO_ROOT & "/scripts/headless/vba/OBTSetupImportHeadless.bas"
Private Const FILLED_SETUP As String = REPO_ROOT & "/.obt/draft/demo_setup_filled.xlsb"
Private Const MERGED_FORMS As String = REPO_ROOT & "/.test-runner/forms/merged"
Private Const RIBBON_TEMPLATE As String = REPO_ROOT & "/ribbons/_ribbontemplate_dev.xlsb"

Private Const OUTPUT_FOLDER As String = REPO_ROOT & "/.obt/draft"
Private Const OUTPUT_NAME As String = "headless_linelist"

'Three of the 55 components the transfer moves, one of each kind. A workbook
'holding all three held the loop that moved them.
Private Const PROBE_CLASS As String = "LLdictionary"
Private Const PROBE_MODULE As String = "EventsLinelistButtons"
Private Const PROBE_FORM As String = "F_Advanced"

Private Outcome As String
Private ProbeNarrative As String
Private BuiltPath As String
Private LogFilePath As String
Private BuildReport As String
Private SheetsBuilt As Long
Private VariablesBuilt As Long
Private ComponentsImported As Long
Private ComponentsDelivered As Long
Private DeliveredNames As String
Private SetupOutcome As String
Private SetupError As Long
Private SetupMessage As String
Private CurrentStep As String


'@section Lifecycle
'===============================================================================

'@sub-title Build the linelist once, and read back what landed.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestHeadlessLinelistBuild"

    SetupError = 0
    SetupMessage = vbNullString
    CurrentStep = "start"

    'One grant for everything this suite and the build touch. The repo root
    'covers src/, .mock, ribbons, .obt and .test-runner in a single dialog on
    'the first run of a machine, and no dialog after that.
    HeadlessBuild.EnsureFileAccess Array(REPO_ROOT)

    On Error Resume Next
        BuildTheLinelist
        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Hand the screen back and print.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    'The build opens and closes four workbooks, so the screen goes back to the
    'driver before PrintResults writes into one of its own sheets.
    On Error Resume Next
        ThisWorkbook.Activate
    On Error GoTo 0

    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section The build
'===============================================================================

'@sub-title Fill a setup when needed, generate from it, then read the file back.
Private Sub BuildTheLinelist()
    CurrentStep = "check the merged forms folder"
    If Not FolderIsThere(MERGED_FORMS) Then
        Outcome = "MISSING: " & MERGED_FORMS & _
                  " - run 'Rscript scripts/headless/merge-form-code.R' first"
        Exit Sub
    End If

    CurrentStep = "check the designer file"
    If Len(Dir$(DESIGNER_FILE)) = 0 Then
        Outcome = "MISSING: " & DESIGNER_FILE
        Exit Sub
    End If

    CurrentStep = "check the ribbon template"
    If Len(Dir$(RIBBON_TEMPLATE)) = 0 Then
        Outcome = "MISSING: " & RIBBON_TEMPLATE
        Exit Sub
    End If

    CurrentStep = "fill the setup"
    SetupOutcome = EnsureFilledSetup()
    If SetupOutcome <> "OK" Then
        Outcome = "the setup could not be filled - " & SetupOutcome
        Exit Sub
    End If

    CurrentStep = "generate the linelist"
    DropPreviousOutput

    Outcome = HeadlessBuild.BuildLinelistFromSetup( _
                  DESIGNER_FILE, FILLED_SETUP, REPO_ROOT, MERGED_FORMS, _
                  OUTPUT_FOLDER, OUTPUT_NAME, BuildOptions())

    CurrentStep = "read what the build recorded"
    BuildReport = HeadlessBuild.LastReport()
    BuiltPath = HeadlessBuild.LastLinelistPath()
    LogFilePath = HeadlessBuild.LastLogPath()
    SheetsBuilt = HeadlessBuild.LastSheetCount()
    VariablesBuilt = HeadlessBuild.LastVariableCount()
    ComponentsImported = HeadlessBuild.LastComponentCount()

    CurrentStep = "reopen the linelist and count its components"
    ReadDeliveredComponents

    'The bisection, only when the build refused: the same specification phase,
    'replayed one public call at a time, so the step that failed raises its own
    'error instead of the boundary-mangled one the build reports.
    If Outcome <> "OK" Then
        CurrentStep = "replay the specification phase"
        ReplayPrepareSteps
    End If
End Sub

'@sub-title Replay the opening steps of LinelistSpecs.Prepare, one at a time.
'@details
'Prepare runs behind one error boundary, and a fault inside it reaches the
'caller as "Method 'Prepare' of object 'LinelistSpecs' failed" with the
'project as its source -- everything that says WHERE is gone. Every step of
'its opening sequence is public API, so this replays them under a
'per-step handler and records the first raw error. Diagnostic only: whatever
'this builds is dropped, and a green replay proves nothing about the build.
Private Sub ReplayPrepareSteps()
    Dim repos As TemporaryRepos
    Dim tplWkb As Workbook
    Dim desWkb As Workbook
    Dim tempPath As String
    Dim previousEvents As Boolean

    'Events off for the whole replay: the designer copy carries the full
    're-imported code, and opening it with events live runs its Workbook_Open
    'with no screen for whatever that puts up.
    previousEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error Resume Next

    Err.Clear
    Set repos = TemporaryRepos.Create(OUTPUT_FOLDER)
    repos.EnsureReady
    tempPath = repos.CreateFilePath("__probe.xlsb")
    If Not ProbeStep("temporary repository") Then GoTo ProbeDone

    Err.Clear
    Set tplWkb = Application.Workbooks.Open(fileName:=RIBBON_TEMPLATE)
    If Not ProbeStep("open the ribbon template") Then GoTo ProbeDone

    Err.Clear
    tplWkb.SaveAs fileName:=tempPath, fileFormat:=xlExcel12
    If Not ProbeStep("save the template copy") Then GoTo ProbeDone

    Err.Clear
    Set desWkb = Application.Workbooks.Open( _
                     fileName:=OUTPUT_FOLDER & "/" & OUTPUT_NAME & "-designer.xlsb")
    If Not ProbeStep("open the designer copy") Then GoTo ProbeDone

    Err.Clear
    InitTransfer.TransferToLinelist desWkb, FILLED_SETUP, tplWkb
    If Not ProbeStep("transfer the setup and designer components") Then GoTo ProbeDone

    Dim dictProbe As LLdictionary
    Dim tradProbe As LLTranslation

    Err.Clear
    Set dictProbe = InitTransfer.LinelistDictionary(tplWkb)
    Set tradProbe = InitTransfer.LinelistTranslations(tplWkb)
    ProbeStep "build the domain managers"

ProbeDone:
    If Not tplWkb Is Nothing Then tplWkb.Close SaveChanges:=False
    If Not desWkb Is Nothing Then desWkb.Close SaveChanges:=False
    Kill tempPath
    Err.Clear
    On Error GoTo 0

    Application.EnableEvents = previousEvents
End Sub

'@fun-title Record one replay step, and say whether the replay goes on.
'@details
'The narrative is also flushed to a text file per step, because a step that
'HANGS leaves no CSV at all -- the trace on disk is then the only record of
'how far the replay got.
'@param stepName String. The step that just ran.
'@return Boolean. True when the step raised nothing.
Private Function ProbeStep(ByVal stepName As String) As Boolean
    If Err.Number = 0 Then
        ProbeNarrative = ProbeNarrative & stepName & ": ok / "
        ProbeStep = True
    Else
        ProbeNarrative = ProbeNarrative & stepName & ": ERROR " & _
                         CStr(Err.Number) & " (" & Err.Source & ") " & _
                         Err.Description & " / "
    End If

    WriteTrace ProbeNarrative
End Function

'@sub-title Flush the replay narrative to a file beside the outputs.
'@param traceText String. The narrative so far.
Private Sub WriteTrace(ByVal traceText As String)
    Dim fileNumber As Long

    On Error Resume Next
        fileNumber = FreeFile
        Open OUTPUT_FOLDER & "/replay-trace.txt" For Output As #fileNumber
        Print #fileNumber, traceText
        Close #fileNumber
        Err.Clear
    On Error GoTo 0
End Sub

'@fun-title The options string of the run.
'@details
'The template is named, so this is the RIBBON build: no action buttons on the
'sheets and no Admin sheet. The geobase is left empty, which is the common
'case and the one the build has to tolerate. The open password is empty on
'purpose, so the file opens on a double-click and the count below can read it.
'@return String. Pipe-separated key=value pairs.
Private Function BuildOptions() As String
    BuildOptions = "temppath=" & RIBBON_TEMPLATE & "|" & _
                   "geopath=" & "|" & _
                   "llpassword="
End Function

'@fun-title Fill a real setup from the generic one, unless one is already there.
'@details
'TestHeadlessSetupFill writes the same file and the harness gives no order
'between modules, so this reuses what is on disk rather than filling twice.
'@return String. "OK", or what went wrong.
Private Function EnsureFilledSetup() As String
    If Len(Dir$(FILLED_SETUP)) > 0 Then
        EnsureFilledSetup = "OK"
        Exit Function
    End If

    If Len(Dir$(EMPTY_SETUP)) = 0 Then
        EnsureFilledSetup = "MISSING: " & EMPTY_SETUP
        Exit Function
    End If

    If Len(Dir$(GENERIC_SETUP)) = 0 Then
        EnsureFilledSetup = "MISSING: " & GENERIC_SETUP
        Exit Function
    End If

    FileCopy EMPTY_SETUP, FILLED_SETUP

    EnsureFilledSetup = HeadlessBuild.ImportSetupFromWorkbook( _
                            FILLED_SETUP, GENERIC_SETUP, INJECTED_MODULE)
End Function

'@sub-title Drop what a previous run of this suite wrote.
'@details
'The three files the build writes. A stale linelist left behind would let a
'failed run report a file on disk and pass.
Private Sub DropPreviousOutput()
    On Error Resume Next
        Kill OUTPUT_FOLDER & "/" & OUTPUT_NAME & ".xlsb"
        Kill OUTPUT_FOLDER & "/" & OUTPUT_NAME & "-generation.txt"
        Kill OUTPUT_FOLDER & "/" & OUTPUT_NAME & "-designer.xlsb"
        Err.Clear
    On Error GoTo 0
End Sub

'@sub-title Reopen the built linelist and count the components of its project.
'@details
'The proof that the transfer ran. Events are off for the open, because the
'workbook now carries the linelist's own Workbook_Open handler and this run
'has no screen for whatever it would put there.
Private Sub ReadDeliveredComponents()
    Dim builtBook As Workbook
    Dim proj As Object
    Dim comp As Object
    Dim previousEvents As Boolean

    ComponentsDelivered = -1
    If Len(Dir$(BuiltPath)) = 0 Then Exit Sub

    previousEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo CleanExit
    Set builtBook = Application.Workbooks.Open(fileName:=BuiltPath, ReadOnly:=True)
    Set proj = builtBook.VBProject

    ComponentsDelivered = proj.VBComponents.Count

    For Each comp In proj.VBComponents
        If comp.Name = PROBE_CLASS Or comp.Name = PROBE_MODULE Or _
           comp.Name = PROBE_FORM Then
            DeliveredNames = DeliveredNames & comp.Name & " "
        End If
    Next

CleanExit:
    On Error Resume Next
        If Not builtBook Is Nothing Then builtBook.Close SaveChanges:=False
        Application.EnableEvents = previousEvents
        Err.Clear
    On Error GoTo 0
End Sub

'@fun-title True when a path names a folder that is there.
'@param folderPath String. The path to judge.
'@return Boolean. True for an existing folder.
Private Function FolderIsThere(ByVal folderPath As String) As Boolean
    Dim folderAttributes As Long

    On Error Resume Next
        folderAttributes = GetAttr(folderPath)
        If Err.Number = 0 Then
            FolderIsThere = ((folderAttributes And vbDirectory) = vbDirectory)
        End If
        Err.Clear
    On Error GoTo 0
End Function


'@section What the run reports
'===============================================================================

'@sub-title The build answered OK.
'@TestMethod("headless")
Private Sub TestTheHeadlessBuildAnswersOK()
    Const TESTNAME As String = "TestTheHeadlessBuildAnswersOK"

    If Not RanAtAll(TESTNAME) Then Exit Sub

    On Error GoTo ErrHandler
    CustomTestSetTitles Assert, TESTNAME, "The headless linelist build"
    Assert.AreEqual "OK", Outcome, _
                    "the build answers OK, and names its fault otherwise"
    Assert.IsTrue Len(BuildReport) > 0, "the build recorded what it did"
    Exit Sub

ErrHandler:
    CustomTestLogFailure Assert, TESTNAME, Err.Number, Err.Description
End Sub

'@sub-title The linelist and its log are on disk.
'@TestMethod("headless")
Private Sub TestTheLinelistIsWritten()
    Const TESTNAME As String = "TestTheLinelistIsWritten"

    If Not RanAtAll(TESTNAME) Then Exit Sub

    On Error GoTo ErrHandler
    CustomTestSetTitles Assert, TESTNAME, "The written files"
    Assert.IsTrue Len(BuiltPath) > 0, "the build answered the path it wrote"
    Assert.IsTrue Len(Dir$(BuiltPath)) > 0, _
                  "the linelist file is on disk (" & BuiltPath & ")"
    Assert.IsTrue Len(Dir$(LogFilePath)) > 0, _
                  "the generation log is on disk (" & LogFilePath & ")"
    Exit Sub

ErrHandler:
    CustomTestLogFailure Assert, TESTNAME, Err.Number, Err.Description
End Sub

'@sub-title The build wrote the sheets and variables the dictionary asked for.
'@TestMethod("headless")
Private Sub TestTheBuildWroteTheDictionary()
    Const TESTNAME As String = "TestTheBuildWroteTheDictionary"

    If Not RanAtAll(TESTNAME) Then Exit Sub

    On Error GoTo ErrHandler
    CustomTestSetTitles Assert, TESTNAME, "What the build wrote"
    Assert.IsTrue SheetsBuilt > 0, _
                  "the build wrote data entry sheets (" & CStr(SheetsBuilt) & ")"
    Assert.IsTrue VariablesBuilt > 0, _
                  "the build wrote variables (" & CStr(VariablesBuilt) & ")"
    Exit Sub

ErrHandler:
    CustomTestLogFailure Assert, TESTNAME, Err.Number, Err.Description
End Sub

'@sub-title Every component was re-imported from source before the transfer.
'@details
'The whole point of copying a designer rather than driving one: the components
'the transfer moves are the ones in src/, not the ones somebody last pasted
'into a binary.
'@TestMethod("headless")
Private Sub TestTheDesignerCarriedTheCurrentSource()
    Const TESTNAME As String = "TestTheDesignerCarriedTheCurrentSource"

    If Not RanAtAll(TESTNAME) Then Exit Sub

    On Error GoTo ErrHandler
    CustomTestSetTitles Assert, TESTNAME, "The designer copy"
    Assert.IsTrue ComponentsImported > 55, _
                  "every component of src/ and every merged form was " & _
                  "re-imported (" & CStr(ComponentsImported) & ")"
    Assert.IsTrue Len(Dir$(OUTPUT_FOLDER & "/" & OUTPUT_NAME & "-designer.xlsb")) > 0, _
                  "the designer the run used was kept beside the linelist"
    Exit Sub

ErrHandler:
    CustomTestLogFailure Assert, TESTNAME, Err.Number, Err.Description
End Sub

'@sub-title The delivered linelist carries the transferred code.
'@details
'The assertion this suite was written for. Everything else can pass over a
'workbook holding no VBA at all.
'@TestMethod("headless")
Private Sub TestTheLinelistCarriesItsCode()
    Const TESTNAME As String = "TestTheLinelistCarriesItsCode"

    If Not RanAtAll(TESTNAME) Then Exit Sub

    On Error GoTo ErrHandler
    CustomTestSetTitles Assert, TESTNAME, "The delivered code"
    Assert.IsTrue ComponentsDelivered > 55, _
                  "the linelist project holds the transferred components (" & _
                  CStr(ComponentsDelivered) & ")"
    Assert.AreEqual PROBE_CLASS & " " & PROBE_MODULE & " " & PROBE_FORM & " ", _
                    DeliveredNames, _
                    "a class, a module and a form all arrived"
    Exit Sub

ErrHandler:
    CustomTestLogFailure Assert, TESTNAME, Err.Number, Err.Description
End Sub

'@sub-title The narrative of the run, for a reader.
'@details
'Not a judgement: one entry carrying what the build recorded, so a failure
'above can be read without opening the log file.
'@TestMethod("headless")
Private Sub TestTheBuildNarrative()
    Const TESTNAME As String = "TestTheBuildNarrative"

    On Error GoTo ErrHandler
    CustomTestSetTitles Assert, TESTNAME, "The narrative"
    Assert.IsTrue True, "steps -> " & Replace(BuildReport, vbLf, " / ")

    If LenB(ProbeNarrative) > 0 Then
        Assert.IsTrue True, "replay -> " & ProbeNarrative
    End If
    Exit Sub

ErrHandler:
    CustomTestLogFailure Assert, TESTNAME, Err.Number, Err.Description
End Sub


'@section Helpers
'===============================================================================

'@fun-title Report a build that could not run, once per test.
'@param testName String. The test asking.
'@return Boolean. True when the build ran.
Private Function RanAtAll(ByVal testName As String) As Boolean
    If SetupError = 0 Then
        RanAtAll = True
        Exit Function
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The build could not run at [" & CurrentStep & "] - " & _
                         SetupMessage
End Function
