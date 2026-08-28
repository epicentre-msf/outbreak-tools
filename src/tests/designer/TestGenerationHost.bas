Attribute VB_Name = "TestGenerationHost"
Attribute VB_Description = "Unit tests for the GenerationHost class"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates GenerationHost: the path read off the designer flag and the choice, the file-name rules, Run in place on a step that answers and on one that raises, and on Windows the hidden instance: its start, the copy opened read-write, Run across the processes, the instance gone, and ReleaseInstance from a call and from the drop of the object.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'THE PROBE STEPS LIVE HERE
'-------------------------------------------------------------------------------
'Run reaches a public function of the workbook the steps run inside. On the
'instance path that workbook is a copy of this driver, so the three probe
'steps below travel with it: HostProbeOk answers "OK", HostProbeEcho
'answers its argument, HostProbeFails traps its own raise and answers the
'error outcome, the way every BuildSteps step does. They carry no test
'annotation, so the runner leaves them alone.
'
'No probe step raises. Application.Run opens a fresh call stack, so an
'untrapped raise inside a step reaches no handler of the caller on either
'path: in place it shows the runtime error box, across the processes the
'hidden instance shows the VBE box and the visible side blocks. Both wedge
'a headless run with nobody there to click.
'
'THE PROCESS COUNT
'-------------------------------------------------------------------------------
'The instance tests count the Excel processes of the machine before and
'after, through WMI, so a leaked instance fails the test. Another Excel
'starting or stopping on the machine during one of these tests moves the
'count and fails it too; run them again when that happens.
'@depends GenerationHost, DesignerPreparation, CustomTest, TestHelpersLite

Private Assert As CustomTest
Private FixtureWorkbook As Workbook

'The host of the running test, released by TestCleanup whatever the exit
Private heldHost As GenerationHost

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'The folder the copy is written in, beside this driver
Private Const COPY_FOLDER_NAME As String = "host_copy"
Private Const COPY_FILE_NAME As String = "__designer_copy.xlsb"

'The flag of the designer ribbon the host reads on Windows
Private Const FLAG_BUILD_IN_PLACE As String = "chkBuildInPlace"

'The probe steps, qualified the way a driver names them
Private Const STEP_OK As String = "TestGenerationHost.HostProbeOk"
Private Const STEP_ECHO As String = "TestGenerationHost.HostProbeEcho"
Private Const STEP_FAILS As String = "TestGenerationHost.HostProbeFails"
Private Const STEP_MISSING As String = "TestGenerationHost.HostProbeNowhere"

Private Const PROBE_RAISE_TEXT As String = "the probe step raised on purpose"

'How long a process count is waited for, in seconds
Private Const PROCESS_WAIT_SECONDS As Long = 30


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestGenerationHost"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        ReleaseHeldHost
        DeleteCopyFile
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    BusyApp

    Set FixtureWorkbook = NewWorkbook
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        ReleaseHeldHost
        DeleteCopyFile
        DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set FixtureWorkbook = Nothing

    'A mid-test workbook close can hand the screen to another workbook
    ThisWorkbook.Activate

    RestoreApp
End Sub


'@section The probe steps
'===============================================================================

'@Description("A step that answers OK.")
Public Function HostProbeOk() As String
    HostProbeOk = "OK"
End Function

'@Description("A step that answers its argument.")
'@param text String. What to answer.
Public Function HostProbeEcho(ByVal text As String) As String
    HostProbeEcho = "echo:" & text
End Function

'@Description("A step that fails the way a build step fails: it traps its own raise and answers the error outcome.")
Public Function HostProbeFails() As String
    On Error GoTo Failed
    Err.Raise ProjectError.SomethingWentWrong, "TestGenerationHost", PROBE_RAISE_TEXT
    Exit Function

Failed:
    HostProbeFails = "ERROR " & CStr(Err.Number) & " (" & Err.Source & "): " & Err.Description
End Function


'@section Helpers
'===============================================================================

'@sub-title Whether this Excel can start a second instance
'@return Boolean. True on Windows.
Private Function InstancePathAvailable() As Boolean
    #If Mac Then
        InstancePathAvailable = False
    #Else
        InstancePathAvailable = True
    #End If
End Function

'@sub-title The number of Excel processes on the machine
'@return Long. The count, or -1 when it cannot be read.
Private Function ExcelProcessCount() As Long
    #If Mac Then
        ExcelProcessCount = -1
    #Else
        Dim management As Object
        Dim processes As Object

        On Error GoTo Unreadable
        Set management = GetObject("winmgmts:\\.\root\cimv2")
        Set processes = management.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
        ExcelProcessCount = processes.Count
        On Error GoTo 0
        Exit Function

Unreadable:
        ExcelProcessCount = -1
    #End If
End Function

'@sub-title Wait for the Excel process count to reach a value
'@details
'A quit instance takes a moment to leave the process table, so the count
'is read again until it answers the expected value or the wait runs out.
'@param expected Long. The count waited for.
'@return Long. The last count read.
Private Function WaitForProcessCount(ByVal expected As Long) As Long
    Dim startedAt As Single
    Dim observed As Long

    startedAt = Timer
    Do
        observed = ExcelProcessCount()
        If observed = expected Then Exit Do
        DoEvents
    Loop While Timer - startedAt < PROCESS_WAIT_SECONDS And Timer >= startedAt

    WaitForProcessCount = observed
End Function

'@sub-title The folder the copy is written in
'@return String. An existing folder beside this driver.
Private Function CopyFolder() As String
    CopyFolder = BuildTempFolder(ThisWorkbook, COPY_FOLDER_NAME)
End Function

'@sub-title The path the copy takes inside the copy folder
'@return String.
Private Function ExpectedCopyPath() As String
    ExpectedCopyPath = CopyFolder() & Application.PathSeparator & COPY_FILE_NAME
End Function

'@sub-title Remove a copy file a test left behind
Private Sub DeleteCopyFile()
    On Error Resume Next
    If LenB(Dir$(ExpectedCopyPath())) > 0 Then Kill ExpectedCopyPath()
    On Error GoTo 0
End Sub

'@sub-title ReleaseInstance the host of the running test
Private Sub ReleaseHeldHost()
    On Error Resume Next
    If Not heldHost Is Nothing Then heldHost.ReleaseInstance
    Set heldHost = Nothing
    On Error GoTo 0
End Sub

'@sub-title Write the build-in-place flag on the fixture designer
'@details
'SetFlag updates a hidden name and raises when the name is missing; it
'creates the two default flags alone. The name is made first, the way a
'prepared designer carries it.
'@param enabled Boolean. The value.
Private Sub SetFixtureFlag(ByVal enabled As Boolean)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(FixtureWorkbook)
    store.EnsureName FLAG_BUILD_IN_PLACE, "No", HiddenNameTypeString

    DesignerPreparation.Create(FixtureWorkbook).SetFlag FLAG_BUILD_IN_PLACE, enabled
End Sub

'@sub-title Whether an outcome carries the error lead
'@param outcome String.
'@return Boolean.
Private Function IsErrorOutcome(ByVal outcome As String) As Boolean
    IsErrorOutcome = (Left$(outcome, 6) = "ERROR ")
End Function


'@section InPlace Tests
'===============================================================================
'@TestMethod("GenerationHost.InPlace")
Public Sub TestInPlaceReadsTheDesignerFlag()
    CustomTestSetTitles Assert, "GenerationHost", "TestInPlaceReadsTheDesignerFlag"
    On Error GoTo Fail

    'Arrange and act: the flag on, then off
    SetFixtureFlag True
    Dim onHost As GenerationHost
    Set onHost = GenerationHost.Create(FixtureWorkbook)

    SetFixtureFlag False
    Dim offHost As GenerationHost
    Set offHost = GenerationHost.Create(FixtureWorkbook)

    'Assert: the flag decides on Windows; Mac is in place either way
    Assert.IsTrue onHost.InPlace, _
                  "The flag checked should put the build in place."
    Assert.AreEqual Not InstancePathAvailable(), offHost.InPlace, _
                    "The flag unchecked should take the instance path on Windows and stay in place on Mac."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestInPlaceReadsTheDesignerFlag", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationHost.InPlace")
Public Sub TestPathChoiceOverridesTheFlag()
    CustomTestSetTitles Assert, "GenerationHost", "TestPathChoiceOverridesTheFlag"
    On Error GoTo Fail

    'Arrange and act: the flag off with the in-place choice
    SetFixtureFlag False
    Dim forcedInPlace As GenerationHost
    Set forcedInPlace = GenerationHost.Create(FixtureWorkbook, HostPathInPlace)

    'The flag on with the instance choice
    SetFixtureFlag True
    Dim forcedInstance As GenerationHost
    Set forcedInstance = GenerationHost.Create(FixtureWorkbook, HostPathInstance)

    'Assert
    Assert.IsTrue forcedInPlace.InPlace, _
                  "The in-place choice should win over the flag."
    Assert.AreEqual Not InstancePathAvailable(), forcedInstance.InPlace, _
                    "The instance choice should win over the flag on Windows and stay in place on Mac."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestPathChoiceOverridesTheFlag", Err.Number, Err.Description
End Sub


'@section CheckOpenNames Tests
'===============================================================================
'@TestMethod("GenerationHost.CheckOpenNames")
Public Sub TestCheckOpenNamesRefusesTwoPathsSharingAName()
    CustomTestSetTitles Assert, "GenerationHost", "TestCheckOpenNamesRefusesTwoPathsSharingAName"
    On Error GoTo Fail

    'Arrange: a host that never acquires, so the check reads this Excel
    Dim host As GenerationHost
    Set host = GenerationHost.Create(FixtureWorkbook, HostPathInPlace)

    'Act: a setup and a template sharing a name, in different folders and cases
    Dim clashText As String
    clashText = host.CheckOpenNames(BetterArrayFromList("C:\one\setup.xlsb", _
                                                        "C:\two\SETUP.xlsb", _
                                                        "C:\three\geo.xlsx"))

    'Assert: both paths are named
    Assert.IsTrue InStr(1, clashText, "share the name") > 0, _
                  "Two files of one name should be refused."
    Assert.IsTrue InStr(1, clashText, "C:\one\setup.xlsb") > 0, _
                  "The refusal should name the first path."
    Assert.IsTrue InStr(1, clashText, "C:\two\SETUP.xlsb") > 0, _
                  "The refusal should name the second path."

    'Act and assert: distinct names, empty entries and no list pass
    Assert.AreEqual vbNullString, _
                    host.CheckOpenNames(BetterArrayFromList("C:\one\setup.xlsb", "", _
                                                            "C:\one\template.xlsb", "   ")), _
                    "Distinct names should answer no refusal, empty entries skipped."
    Assert.AreEqual vbNullString, host.CheckOpenNames(Nothing), _
                    "No list should answer no refusal."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestCheckOpenNamesRefusesTwoPathsSharingAName", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationHost.CheckOpenNames")
Public Sub TestCheckOpenNamesRefusesAFileOpenInTheInstance()
    CustomTestSetTitles Assert, "GenerationHost", "TestCheckOpenNamesRefusesAFileOpenInTheInstance"
    On Error GoTo Fail

    'Arrange: the fixture and this driver are both open in this Excel
    Dim host As GenerationHost
    Set host = GenerationHost.Create(FixtureWorkbook, HostPathInPlace)

    'Act: a file elsewhere carrying the fixture's name, and the driver's own name
    Dim clashText As String
    clashText = host.CheckOpenNames(BetterArrayFromList("C:\elsewhere\" & FixtureWorkbook.Name, _
                                                        ThisWorkbook.Name))

    'Assert: both are refused, the open file named
    Assert.IsTrue InStr(1, clashText, "already open") > 0, _
                  "A name open in the instance should be refused."
    Assert.IsTrue InStr(1, clashText, FixtureWorkbook.FullName) > 0, _
                  "The refusal should name the open fixture."
    Assert.IsTrue InStr(1, clashText, ThisWorkbook.FullName) > 0, _
                  "The designer itself should count as an open name."
    Assert.AreEqual CLng(2), UBound(Split(clashText, vbLf)) + 1, _
                    "Each refusal should take one line."

    'Act and assert: a name nobody has open passes
    Assert.AreEqual vbNullString, _
                    host.CheckOpenNames(BetterArrayFromList("C:\elsewhere\nobody_has_this_open.xlsb")), _
                    "A name nobody has open should answer no refusal."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestCheckOpenNamesRefusesAFileOpenInTheInstance", Err.Number, Err.Description
End Sub


'@section Run in place Tests
'===============================================================================
'@TestMethod("GenerationHost.Run")
Public Sub TestRunInPlaceAnswersTheStep()
    CustomTestSetTitles Assert, "GenerationHost", "TestRunInPlaceAnswersTheStep"
    On Error GoTo Fail

    'Arrange: in place over this driver, which carries the probe steps
    Set heldHost = GenerationHost.Create(ThisWorkbook, HostPathInPlace)
    heldHost.Acquire

    Assert.IsTrue heldHost.IsAcquired, "Acquire should mark the host acquired."
    Assert.IsTrue heldHost.HostApplication Is Application, _
                  "In place the host application should be this Excel."
    Assert.AreEqual CLng(0), heldHost.Hwnd, "In place no window handle is taken."

    'Act: the copy is the designer itself
    Dim copyBook As Workbook
    Set copyBook = heldHost.OpenDesignerCopy()

    Assert.IsTrue copyBook Is ThisWorkbook, _
                  "In place the designer itself is the workbook the steps run inside."
    Assert.AreEqual vbNullString, heldHost.CopyPath, "In place no copy file is written."

    'Act and assert: a step that answers OK, and one that answers its argument
    Assert.AreEqual "OK", heldHost.Run(STEP_OK), _
                    "Run should answer what the step answered."
    Assert.AreEqual "echo:abc", heldHost.Run(STEP_ECHO, "abc"), _
                    "Run should forward an argument to the step."
    Assert.AreEqual STEP_ECHO, heldHost.LastStep, "LastStep should name the last step run."

    'Act and assert: ReleaseInstance is a no-op that still answers OK
    Assert.AreEqual "OK", heldHost.ReleaseInstance(), "ReleaseInstance in place should answer OK."
    Assert.IsFalse heldHost.IsAcquired, "ReleaseInstance should clear the acquired mark."
    Assert.AreEqual "OK", heldHost.ReleaseOutcome, "ReleaseOutcome should keep the answer."
    Assert.IsNothing heldHost.DesignerCopy, "ReleaseInstance should drop the copy reference."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestRunInPlaceAnswersTheStep", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationHost.Run")
Public Sub TestRunInPlacePassesAnErrorOutcomeThrough()
    CustomTestSetTitles Assert, "GenerationHost", "TestRunInPlacePassesAnErrorOutcomeThrough"
    On Error GoTo Fail

    'Arrange
    Set heldHost = GenerationHost.Create(ThisWorkbook, HostPathInPlace)
    heldHost.Acquire
    heldHost.OpenDesignerCopy

    'Act: a step that failed answers its error outcome, and Run passes it on
    Dim outcome As String
    outcome = heldHost.Run(STEP_FAILS)

    'Assert
    Assert.IsTrue IsErrorOutcome(outcome), _
                  "A step that failed should answer an error outcome: " & outcome
    Assert.IsTrue InStr(1, outcome, PROBE_RAISE_TEXT) > 0, _
                  "The outcome should carry the description the step wrote: " & outcome
    Assert.IsFalse heldHost.InstanceStopped, _
                   "A failed step should leave the instance marked alive."

    'Act and assert: a step the workbook does not carry, then the host still runs
    outcome = heldHost.Run(STEP_MISSING)
    Assert.IsTrue IsErrorOutcome(outcome), _
                  "A missing step should answer an error outcome: " & outcome
    Assert.IsFalse heldHost.InstanceStopped, _
                   "A missing step should leave the instance marked alive."
    Assert.AreEqual "OK", heldHost.Run(STEP_OK), _
                    "The host should still run a step after an error outcome."

    Assert.AreEqual "OK", heldHost.ReleaseInstance(), "ReleaseInstance should answer OK."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestRunInPlacePassesAnErrorOutcomeThrough", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationHost.Run")
Public Sub TestMembersOutOfOrderRaise()
    CustomTestSetTitles Assert, "GenerationHost", "TestMembersOutOfOrderRaise"
    On Error GoTo Fail

    Dim host As GenerationHost
    Set host = GenerationHost.Create(FixtureWorkbook, HostPathInPlace)

    'Act and assert: OpenDesignerCopy before Acquire
    Dim errNumber As Long
    On Error Resume Next
    host.OpenDesignerCopy
    errNumber = Err.Number
    On Error GoTo Fail
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), errNumber, _
                    "OpenDesignerCopy before Acquire should raise the unexpected-state error."

    'Act and assert: Run before OpenDesignerCopy
    host.Acquire
    On Error Resume Next
    host.Run STEP_OK
    errNumber = Err.Number
    On Error GoTo Fail
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), errNumber, _
                    "Run before OpenDesignerCopy should raise the unexpected-state error."

    'Act and assert: Acquire twice
    On Error Resume Next
    host.Acquire
    errNumber = Err.Number
    On Error GoTo Fail
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), errNumber, _
                    "A second Acquire should raise the unexpected-state error."

    host.ReleaseInstance

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestMembersOutOfOrderRaise", Err.Number, Err.Description
End Sub


'@section Instance Tests (Windows)
'===============================================================================
'@TestMethod("GenerationHost.Instance")
Public Sub TestAcquireStartsAHiddenInstanceAndReleaseQuitsIt()
    CustomTestSetTitles Assert, "GenerationHost", "TestAcquireStartsAHiddenInstanceAndReleaseQuitsIt"
    On Error GoTo Fail

    If Not InstancePathAvailable() Then
        Assert.IsTrue True, "The instance path exists on Windows alone; nothing to check here."
        Exit Sub
    End If

    'Arrange
    Dim before As Long
    before = ExcelProcessCount()

    'Act: a hidden instance
    Set heldHost = GenerationHost.Create(ThisWorkbook, HostPathInstance)
    heldHost.Acquire

    'Assert: one more Excel, hidden and quiet, its handle recorded
    Assert.AreEqual before + 1, WaitForProcessCount(before + 1), _
                    "Acquire should start one Excel process."
    Assert.IsTrue heldHost.Hwnd <> 0, "Acquire should record the window handle."
    Assert.IsFalse heldHost.HostApplication Is Application, _
                   "The host application should be another instance."
    Assert.IsFalse heldHost.HostApplication.Visible, "The instance should be hidden."
    Assert.IsFalse heldHost.HostApplication.DisplayAlerts, "The instance should show no alert."
    Assert.IsFalse heldHost.HostApplication.EnableEvents, "The instance should fire no event."
    Assert.IsFalse heldHost.HostApplication.AskToUpdateLinks, "The instance should ask about no link."
    Assert.AreEqual CLng(msoAutomationSecurityLow), CLng(heldHost.HostApplication.AutomationSecurity), _
                    "The instance should run the macros of the copy."
    Assert.IsTrue heldHost.HostApplication.ScreenUpdating, _
                  "ScreenUpdating stays on in the instance, or the output workbook refuses its freeze."

    'Act: ReleaseInstance quits it
    Dim recordedHandle As Long
    recordedHandle = heldHost.Hwnd
    Assert.AreEqual "OK", heldHost.ReleaseInstance(), "ReleaseInstance should answer OK."

    'Assert: the process is gone, the handle stays on record
    Assert.AreEqual before, WaitForProcessCount(before), _
                    "ReleaseInstance should quit the instance it started."
    Assert.IsFalse heldHost.IsAcquired, "ReleaseInstance should clear the acquired mark."
    Assert.AreEqual recordedHandle, heldHost.Hwnd, "The handle should stay readable after ReleaseInstance."
    Assert.IsNothing heldHost.HostApplication, "ReleaseInstance should drop the instance reference."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestAcquireStartsAHiddenInstanceAndReleaseQuitsIt", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationHost.Instance")
Public Sub TestOpenDesignerCopyOpensTheCopyReadWriteInTheInstance()
    CustomTestSetTitles Assert, "GenerationHost", "TestOpenDesignerCopyOpensTheCopyReadWriteInTheInstance"
    On Error GoTo Fail

    If Not InstancePathAvailable() Then
        Assert.IsTrue True, "The instance path exists on Windows alone; nothing to check here."
        Exit Sub
    End If

    'Arrange
    Dim before As Long
    before = ExcelProcessCount()

    Set heldHost = GenerationHost.Create(ThisWorkbook, HostPathInstance)
    heldHost.Acquire

    'Act
    Dim copyBook As Workbook
    Set copyBook = heldHost.OpenDesignerCopy(CopyFolder())

    'Assert: the copy carries its name, sits in the instance, opened read-write
    Assert.AreEqual COPY_FILE_NAME, copyBook.Name, "The copy should carry the reserved name."
    Assert.AreEqual ExpectedCopyPath(), heldHost.CopyPath, "CopyPath should name the file written."
    Assert.IsTrue LenB(Dir$(heldHost.CopyPath)) > 0, "The copy file should be on disk."
    Assert.IsFalse copyBook.ReadOnly, "The copy should open read-write."
    Assert.IsFalse copyBook.Application Is Application, "The copy should be open in the other instance."
    Assert.AreEqual CLng(xlCalculationManual), CLng(heldHost.HostApplication.Calculation), _
                    "Calculation should go manual once the copy is open."
    Assert.IsTrue copyBook Is heldHost.DesignerCopy, "DesignerCopy should answer the copy."

    'Assert: the copy's name is one more open name in the instance
    Dim clashText As String
    clashText = heldHost.CheckOpenNames(BetterArrayFromList("C:\user\" & COPY_FILE_NAME))
    Assert.IsTrue InStr(1, clashText, "the build instance") > 0, _
                  "A user file carrying the copy's name should be refused: " & clashText

    'Assert: this Excel is untouched, so the visible-side check passes the same name
    Assert.AreEqual vbNullString, _
                    heldHost.CheckOpenNames(BetterArrayFromList("C:\user\" & COPY_FILE_NAME), Application), _
                    "Against this Excel the copy's name is open nowhere."

    'Act and assert: ReleaseInstance closes the copy, quits, deletes the file
    Assert.AreEqual "OK", heldHost.ReleaseInstance(), "ReleaseInstance should answer OK."
    Assert.AreEqual before, WaitForProcessCount(before), "ReleaseInstance should quit the instance."
    Assert.AreEqual vbNullString, Dir$(ExpectedCopyPath()), "ReleaseInstance should delete the copy file."
    Assert.IsNothing heldHost.DesignerCopy, "ReleaseInstance should drop the copy reference."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestOpenDesignerCopyOpensTheCopyReadWriteInTheInstance", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationHost.Instance")
Public Sub TestOpenDesignerCopyRefusesAMissingFolder()
    CustomTestSetTitles Assert, "GenerationHost", "TestOpenDesignerCopyRefusesAMissingFolder"
    On Error GoTo Fail

    If Not InstancePathAvailable() Then
        Assert.IsTrue True, "The instance path exists on Windows alone; nothing to check here."
        Exit Sub
    End If

    'Arrange
    Dim before As Long
    before = ExcelProcessCount()

    Set heldHost = GenerationHost.Create(ThisWorkbook, HostPathInstance)
    heldHost.Acquire

    'Act and assert: an empty folder, then a folder that is not there
    Dim errNumber As Long
    On Error Resume Next
    heldHost.OpenDesignerCopy ""
    errNumber = Err.Number
    On Error GoTo Fail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "An empty folder should raise the invalid-argument error."

    On Error Resume Next
    heldHost.OpenDesignerCopy CopyFolder() & Application.PathSeparator & "nowhere_at_all"
    errNumber = Err.Number
    On Error GoTo Fail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A missing folder should raise the invalid-argument error."
    Assert.IsNothing heldHost.DesignerCopy, "No copy should be open after a refusal."

    'Assert: the instance is still there to release
    Assert.AreEqual "OK", heldHost.ReleaseInstance(), "ReleaseInstance should still answer OK."
    Assert.AreEqual before, WaitForProcessCount(before), "ReleaseInstance should quit the instance."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestOpenDesignerCopyRefusesAMissingFolder", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationHost.Instance")
Public Sub TestRunInTheInstanceAnswersTheStep()
    CustomTestSetTitles Assert, "GenerationHost", "TestRunInTheInstanceAnswersTheStep"
    On Error GoTo Fail

    If Not InstancePathAvailable() Then
        Assert.IsTrue True, "The instance path exists on Windows alone; nothing to check here."
        Exit Sub
    End If

    'Arrange
    Dim before As Long
    before = ExcelProcessCount()

    Set heldHost = GenerationHost.Create(ThisWorkbook, HostPathInstance)
    heldHost.Acquire
    heldHost.OpenDesignerCopy CopyFolder()

    'Act and assert: a step that answers OK, one that answers its argument
    Assert.AreEqual "OK", heldHost.Run(STEP_OK), _
                    "Run across the processes should answer what the step answered."
    Assert.AreEqual "echo:abc", heldHost.Run(STEP_ECHO, "abc"), _
                    "Run across the processes should forward a string argument."

    'Act and assert: a step the copy does not carry raises on this side at once
    Dim outcome As String
    outcome = heldHost.Run(STEP_MISSING)
    Assert.IsTrue IsErrorOutcome(outcome), _
                  "A missing step should answer an error outcome: " & outcome
    Assert.IsFalse heldHost.InstanceStopped, _
                   "A missing step should leave the instance marked alive."
    Assert.AreEqual "OK", heldHost.Run(STEP_OK), _
                    "The instance should still answer after a missing step."

    'Act and assert: ReleaseInstance
    Assert.AreEqual "OK", heldHost.ReleaseInstance(), "ReleaseInstance should answer OK."
    Assert.AreEqual before, WaitForProcessCount(before), "ReleaseInstance should quit the instance."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestRunInTheInstanceAnswersTheStep", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationHost.Instance")
Public Sub TestRunOnAnInstanceGoneAnswersStoppedAnswering()
    CustomTestSetTitles Assert, "GenerationHost", "TestRunOnAnInstanceGoneAnswersStoppedAnswering"
    On Error GoTo Fail

    If Not InstancePathAvailable() Then
        Assert.IsTrue True, "The instance path exists on Windows alone; nothing to check here."
        Exit Sub
    End If

    'Arrange: the instance is quit behind the host's back, through the
    'host's own reference. That is the shape the spike measured: a clean
    'Quit then a Run answers 462, or the disconnected object on the first
    'call. The process itself stays in the table until the host drops its
    'references, which ReleaseInstance does, so the count is read after it.
    Dim before As Long
    before = ExcelProcessCount()

    Set heldHost = GenerationHost.Create(ThisWorkbook, HostPathInstance)
    heldHost.Acquire
    heldHost.OpenDesignerCopy CopyFolder()

    Dim recordedHandle As Long
    recordedHandle = heldHost.Hwnd

    heldHost.DesignerCopy.Close SaveChanges:=False
    heldHost.HostApplication.Quit

    'Act
    Dim outcome As String
    outcome = heldHost.Run(STEP_OK)

    'Assert: the stopped outcome names the step, and the mark is set
    Assert.IsTrue IsErrorOutcome(outcome), _
                  "A Run on an instance gone should answer an error outcome: " & outcome
    Assert.IsTrue InStr(1, outcome, "stopped answering after " & STEP_OK) > 0, _
                  "The outcome should say the instance stopped answering after the step: " & outcome
    Assert.IsTrue heldHost.InstanceStopped, "The instance should be marked stopped."

    'Act and assert: a later Run answers the same without a call
    outcome = heldHost.Run(STEP_ECHO, "abc")
    Assert.IsTrue InStr(1, outcome, "stopped answering after " & STEP_ECHO) > 0, _
                  "Every later Run should answer the stopped outcome: " & outcome

    'Act and assert: ReleaseInstance quits nothing and names the handle
    outcome = heldHost.ReleaseInstance()
    Assert.IsTrue IsErrorOutcome(outcome), _
                  "ReleaseInstance after the instance is gone should answer an error outcome: " & outcome
    Assert.IsTrue InStr(1, outcome, CStr(recordedHandle)) > 0, _
                  "ReleaseInstance should name the window handle of the instance: " & outcome
    Assert.AreEqual outcome, heldHost.ReleaseOutcome, "ReleaseOutcome should keep the answer."
    Assert.AreEqual vbNullString, Dir$(ExpectedCopyPath()), _
                    "ReleaseInstance should still delete the copy file."
    Assert.AreEqual before, WaitForProcessCount(before), _
                    "The process should leave once the host drops its references."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestRunOnAnInstanceGoneAnswersStoppedAnswering", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationHost.Instance")
Public Sub TestDroppingTheHostReleasesTheInstance()
    CustomTestSetTitles Assert, "GenerationHost", "TestDroppingTheHostReleasesTheInstance"
    On Error GoTo Fail

    If Not InstancePathAvailable() Then
        Assert.IsTrue True, "The instance path exists on Windows alone; nothing to check here."
        Exit Sub
    End If

    'Arrange: a host nobody releases
    Dim before As Long
    before = ExcelProcessCount()

    Dim host As GenerationHost
    Set host = GenerationHost.Create(ThisWorkbook, HostPathInstance)
    host.Acquire
    host.OpenDesignerCopy CopyFolder()
    Assert.AreEqual before + 1, WaitForProcessCount(before + 1), "Acquire should start one Excel process."

    'Act: the object is dropped
    Set host = Nothing

    'Assert: the instance is gone and the copy file with it
    Assert.AreEqual before, WaitForProcessCount(before), _
                    "Dropping the host should quit the instance it started."
    Assert.AreEqual vbNullString, Dir$(ExpectedCopyPath()), _
                    "Dropping the host should delete the copy file."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestDroppingTheHostReleasesTheInstance", Err.Number, Err.Description
End Sub
