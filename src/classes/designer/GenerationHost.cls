VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenerationHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "The Excel instance a linelist build runs in, and the way a driver reaches the build steps there"

'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("The Excel instance a linelist build runs in, and the way a driver reaches the build steps there")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@depends ProjectError, DesignerPreparation, BetterArray

Option Explicit

'@class GenerationHost
'@description
'Carries the whole platform difference of a linelist build in one object.
'A build runs as the steps of BuildSteps, and this class owns the Excel
'instance those steps run in. On the instance path (Windows, the
'chkBuildInPlace box unchecked) Acquire starts a hidden Excel,
'OpenDesignerCopy writes a copy of the designer and opens it there, Run
'reaches a step through Application.Run across the two processes, and
'ReleaseInstance closes the copy and quits the instance this object started. On
'the in-place path (Mac always, Windows when the box is checked) the
'instance is the designer's own, the copy is the designer itself, Run is
'an in-process Application.Run and ReleaseInstance is a no-op.
'
'GenerationHost.cls without interface: single implementation, no test fake,
'no compile cycle.
'
'THE HIDDEN INSTANCE SHOWS NOTHING
'-------------------------------------------------------------------------------
'The instance is started with Visible False, DisplayAlerts False,
'EnableEvents False, AskToUpdateLinks False and AutomationSecurity Low. Low
'is the level the steps need: ForceDisable turns off every macro of the
'opened copy and Application.Run answers 1004. The events stay off through
'EnableEvents False, so the copy's Workbook_Open and ribbon load stay quiet.
'Calculation is set to manual once the copy is open, because the write
'raises 1004 on an instance holding no workbook.
'
'SCREENUPDATING IS OFF IN THE HIDDEN INSTANCE
'-------------------------------------------------------------------------------
'An invisible Excel still repaints when ScreenUpdating is on, and Excel
'16.0.20430 keeps shared graphics memory on every one of those repaints:
'about 0.7 MB per cell write, measured on 2026-10-01. On a laptop that memory
'is system RAM. The Info sheet of the ebola setup, 7 variables, cost 168 MB
'and ran six times slower, and the Linelist sheet of 511 variables took
'219 seconds against 93 with ScreenUpdating off. So Acquire switches
'ScreenUpdating off there and Run writes it off again before every step.
'
'THE OUTPUT WORKBOOK IS PAINTED ONCE, FROM THIS SIDE
'-------------------------------------------------------------------------------
'A window that was never painted refuses Window.FreezePanes with 1004, and
'with ScreenUpdating off for the whole build the window of the output
'workbook is never painted: every data entry sheet came out unfrozen.
'Nothing a macro of the instance does paints it. Switching ScreenUpdating
'on inside the step, with the sheet activated, the zoom written, a scroll,
'a cell rewritten, Application.Wait or DoEvents after it, was refused every
'time (tried 2026-10-01). What paints it is ScreenUpdating switched on from
'THIS side while the instance sits idle between two steps, and off again
'right after. PaintInstance does that. The driver calls it once the output
'workbook exists, after the linelist step, and every freeze of the build
'is accepted from then on: Info and the 511 variable Linelist sheet of the
'ebola setup both froze, the shared graphics memory of the instance stayed
'at 18 MB, and the sheet took the same time with and without it.
'
'A CALL THIS EXCEL STOPPED WAITING FOR
'-------------------------------------------------------------------------------
'A step that runs long makes the visible Excel raise its box "waiting for
'another application to complete an OLE action", and the pending call can
'end cancelled: Run then sees -2147418110 while the step carries on in the
'instance and finishes. Its answer is still there. BuildSteps keeps what its
'last step answered, and after a dropped call Run asks for it through
'LastAnswer, waiting for the instance to come free. The step named by
'LastAnswerStep is the one asked; RecoveredCalls counts the answers read
'back that way, for the driver to file in its report.
'
'THE COPY OPENS READ-WRITE
'-------------------------------------------------------------------------------
'The copy is a private temp file named __designer_copy.xlsb that is never
'saved, so it opens read-write. A read-only open of the same file costs 9 to
'20 seconds against 0.6; that difference is what makes the hidden build run
'at the in-place speed. A user file never carries the copy's name.
'
'A RAISE INSIDE A STEP IS A HANG
'-------------------------------------------------------------------------------
'Application.Run opens a fresh call stack, so an untrapped raise inside a
'step reaches no handler here, on either path. In place it shows the
'runtime error box; across the processes the hidden instance shows the VBE
'box, un-hides itself, and the visible side blocks. Every step therefore
'keeps its own handler and answers a string, and Run passes that string
'through. A raise Run sees on the visible side is one of three things. The
'instance gone (462, RPC call failed, RPC server unavailable) is answered
'as "the build instance stopped answering after <step>" with the instance
'marked stopped so ReleaseInstance quits nothing. A call dropped on this
'side while the instance was at work is answered with what the step
'answered, read back from the copy. A plain fault, such as a step name the
'copy does not carry, is answered as "ERROR <n> (<source>): <desc>".
'
'RELEASE RUNS FROM EVERY EXIT
'-------------------------------------------------------------------------------
'ReleaseInstance closes the copy without saving, quits the instance this object
'owns and deletes the copy file. It runs from Class_Terminate too, under On
'Error Resume Next, so a driver that raises on its way out still frees the
'instance. A Quit that fails leaves the instance and answers its window
'handle, taken at Acquire, for the driver to log. The instance is never
'found by process name.
'@depends ProjectError, DesignerPreparation, BetterArray
'@version 1.0 (2026-08-28)


'@section Enumerations
'===============================================================================

'How Create picks the build path. Parameters take the value as Byte.
Public Enum GenerationHostPath
    HostPathFromFlag = 0    'Mac in place; Windows reads chkBuildInPlace
    HostPathInPlace = 1     'in place on both platforms
    HostPathInstance = 2    'the hidden instance on Windows; still in place on Mac
End Enum


'@section State and constants
'===============================================================================
Private Type TGenerationHost
    designerBook As Workbook       'The designer the build reads its entries from
    hostApp As Application         'The instance the build runs in
    ownsInstance As Boolean        'True when Acquire started hostApp
    designerCopy As Workbook       'The workbook the steps run inside: the copy, or the designer in place
    copyPath As String             'The path of the copy file; empty in place
    windowHandle As Long           'The Hwnd of the owned instance, taken at Acquire
    inPlace As Boolean             'The path this object took
    acquired As Boolean            'True between Acquire and ReleaseInstance
    instanceStopped As Boolean     'True once a Run saw the instance gone
    lastStep As String             'The step the last Run was asked for
    releaseOutcome As String       'What the last ReleaseInstance answered
    lastAnswerStep As String       'The step that answers what the last step answered
    recoveredCalls As Long         'The answers read back after a dropped call
End Type

Private this As TGenerationHost

Private Const CLASS_NAME As String = "GenerationHost"

'The name of the designer copy in the instance. Run reaches the steps
'through it, and CheckOpenNames refuses a user file carrying it.
Private Const COPY_FILE_NAME As String = "__designer_copy.xlsb"

'The flag of the designer ribbon that picks the path on Windows
Private Const FLAG_BUILD_IN_PLACE As String = "chkBuildInPlace"

'The outcome shapes, the ones BuildSteps answers
Private Const OUTCOME_OK As String = "OK"
Private Const OUTCOME_ERROR_LEAD As String = "ERROR "

'The errors a Run sees on the visible side when the instance is gone: the
'remote server missing, the RPC call failed, the RPC server unavailable,
'and the object disconnected from its clients, which is what a reference
'answers right after the instance quit under it.
Private Const ERR_REMOTE_SERVER_GONE As Long = 462
Private Const ERR_RPC_CALL_FAILED As Long = -2147023170
Private Const ERR_RPC_SERVER_UNAVAILABLE As Long = -2147023174
Private Const ERR_RPC_DISCONNECTED As Long = -2147417848

'The errors a Run sees when the call was dropped on this side while the
'instance was at work: cancelled by the message filter of this Excel,
'rejected by the instance, or told to come back later.
Private Const ERR_RPC_CALL_CANCELED As Long = -2147418110
Private Const ERR_RPC_CALL_REJECTED As Long = -2147418111
Private Const ERR_RPC_RETRY_LATER As Long = -2147417846

'The step that answers "<step>=<answer>" for the last step the copy ran,
'and the mark between the two parts
Private Const DEFAULT_LAST_ANSWER_STEP As String = "BuildSteps.BuildLastAnswer"
Private Const LAST_ANSWER_MARK As String = "="

'How long the answer of a dropped call is waited for, and the pause
'between two tries, in seconds
Private Const RECOVERY_SECONDS As Long = 1800
Private Const RECOVERY_PAUSE_SECONDS As Long = 2

'What the wait for a last answer ended on
Private Const AWAIT_ANSWERED As Byte = 1
Private Const AWAIT_INSTANCE_GONE As Byte = 2
Private Const AWAIT_UNREADABLE As Byte = 3

'The most arguments Run forwards to a step
Private Const MAX_RUN_ARGUMENTS As Long = 3


'@section Factory
'===============================================================================

'@label:create
'@sub-title Build a host over a designer workbook and pick its path
'@details
'The path is settled here and never changes for the life of the object:
'in place on Mac whatever the choice, and on Windows the chkBuildInPlace
'flag of the designer unless the choice forces a path. Nothing is started
'until Acquire.
'@param designerBook Optional Workbook. The designer the build reads. Nothing takes ThisWorkbook.
'@param pathChoice Optional Byte. A GenerationHostPath member. Defaults to HostPathFromFlag.
'@return GenerationHost. Ready for Acquire.
'@export
Public Function Create(Optional ByVal designerBook As Workbook = Nothing, _
                       Optional ByVal pathChoice As Byte = HostPathFromFlag) As GenerationHost
    Dim host As GenerationHost

    Set host = New GenerationHost
    host.Initialise designerBook, pathChoice
    Set Create = host
End Function

'@label:initialise
'@sub-title Bind the designer and settle the path
'@param designerBook Workbook. The designer, or Nothing for ThisWorkbook.
'@param pathChoice Byte. A GenerationHostPath member.
Friend Sub Initialise(ByVal designerBook As Workbook, ByVal pathChoice As Byte)
    If designerBook Is Nothing Then
        Set this.designerBook = ThisWorkbook
    Else
        Set this.designerBook = designerBook
    End If

    this.inPlace = ResolveInPlace(pathChoice)
    this.releaseOutcome = vbNullString
    this.lastAnswerStep = DEFAULT_LAST_ANSWER_STEP
    this.recoveredCalls = 0
End Sub

'@label:resolve-in-place
'@sub-title Decide the path from the platform, the choice and the flag
'@details
'Mac Excel is single-instance and has no COM, so the answer is True there
'whatever the choice. On Windows a forced choice wins; otherwise the
'chkBuildInPlace flag of the designer decides, False when it was never set.
'@param pathChoice Byte. A GenerationHostPath member.
'@return Boolean. True for the in-place path.
Private Function ResolveInPlace(ByVal pathChoice As Byte) As Boolean
    #If Mac Then
        ResolveInPlace = True
    #Else
        Dim preparation As DesignerPreparation

        Select Case pathChoice
            Case HostPathInPlace
                ResolveInPlace = True
            Case HostPathInstance
                ResolveInPlace = False
            Case Else
                Set preparation = DesignerPreparation.Create(this.designerBook)
                ResolveInPlace = preparation.GetFlag(FLAG_BUILD_IN_PLACE, False)
        End Select
    #End If
End Function


'@section Introspection
'===============================================================================

'@label:in-place
'@prop-title Whether the build runs in the designer's own instance
'@return Boolean. True on Mac, and on Windows when the flag or the choice says so.
'@export
Public Property Get InPlace() As Boolean
    InPlace = this.inPlace
End Property

'@label:is-acquired
'@prop-title Whether Acquire has run and ReleaseInstance has not
'@return Boolean.
'@export
Public Property Get IsAcquired() As Boolean
    IsAcquired = this.acquired
End Property

'@label:instance-stopped
'@prop-title Whether a Run found the instance gone
'@return Boolean. True once a Run answered "stopped answering".
'@export
Public Property Get InstanceStopped() As Boolean
    InstanceStopped = this.instanceStopped
End Property

'@label:hwnd
'@prop-title The window handle of the owned instance, taken at Acquire
'@return Long. 0 in place, and before Acquire.
'@export
Public Property Get Hwnd() As Long
    Hwnd = this.windowHandle
End Property

'@label:host-application
'@prop-title The instance the build runs in
'@return Application. The owned instance, or the designer's own in place. Nothing before Acquire.
'@export
Public Property Get HostApplication() As Application
    Set HostApplication = this.hostApp
End Property

'@label:designer-copy
'@prop-title The workbook the steps run inside
'@return Workbook. The copy in the instance, or the designer in place. Nothing before OpenDesignerCopy.
'@export
Public Property Get DesignerCopy() As Workbook
    Set DesignerCopy = this.designerCopy
End Property

'@label:copy-path
'@prop-title The path of the copy file
'@return String. Empty in place, and before OpenDesignerCopy.
'@export
Public Property Get CopyPath() As String
    CopyPath = this.copyPath
End Property

'@label:last-step
'@prop-title The step the last Run was asked for
'@return String. Empty before the first Run.
'@export
Public Property Get LastStep() As String
    LastStep = this.lastStep
End Property

'@label:release-outcome
'@prop-title What the last ReleaseInstance answered
'@return String. "OK", or "ERROR <n> (<source>): <desc>" naming the window handle. Empty before ReleaseInstance.
'@export
Public Property Get ReleaseOutcome() As String
    ReleaseOutcome = this.releaseOutcome
End Property

'@label:last-answer-step
'@prop-title The step that answers what the last step answered
'@details
'A public function of the workbook the steps run inside, answering
'"<step>=<answer>" for the last step that workbook ran. BuildSteps carries
'BuildLastAnswer, which is the default; a driver whose steps live in another
'module names its own.
'@return String. The step name, with its module.
'@export
Public Property Get LastAnswerStep() As String
    LastAnswerStep = this.lastAnswerStep
End Property

'@label:last-answer-step-let
'@prop-title Name the step that answers what the last step answered
'@param stepName String. A public function of the copy. Empty puts the default back.
'@export
Public Property Let LastAnswerStep(ByVal stepName As String)
    If LenB(Trim$(stepName)) = 0 Then
        this.lastAnswerStep = DEFAULT_LAST_ANSWER_STEP
    Else
        this.lastAnswerStep = Trim$(stepName)
    End If
End Property

'@label:recovered-calls
'@prop-title How many answers were read back after a dropped call
'@return Long. Zero when every Run came back on its own call.
'@export
Public Property Get RecoveredCalls() As Long
    RecoveredCalls = this.recoveredCalls
End Property


'@section The instance
'===============================================================================

'@label:acquire
'@sub-title Start the instance the build runs in
'@details
'In place, the instance is the designer's own Application and nothing
'changes on it. On the instance path a new Excel is started hidden, with
'alerts, events and link prompts off and automation security Low, and its
'window handle is recorded. Calculation waits for OpenDesignerCopy.
'@throws ProjectError.ErrorUnexpectedState When Acquire already ran, or the instance would not start.
'@export
Public Sub Acquire()
    If this.acquired Then
        ThrowError ProjectError.ErrorUnexpectedState, "Acquire already ran; call ReleaseInstance first"
    End If

    this.instanceStopped = False
    this.lastStep = vbNullString
    this.releaseOutcome = vbNullString

    If this.inPlace Then
        Set this.hostApp = this.designerBook.Application
        this.ownsInstance = False
        this.windowHandle = 0
    Else
        StartHiddenInstance
    End If

    this.acquired = True
End Sub

'@label:start-hidden-instance
'@sub-title Start a hidden Excel and set it up for a build nobody watches
'@details
'Windows only; the in-place path never reaches this. A start that fails
'leaves nothing behind.
'@throws ProjectError.ErrorUnexpectedState When the instance would not start or take a setting.
Private Sub StartHiddenInstance()
    Dim failMessage As String

    #If Mac Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "A build instance cannot be started on Mac Excel"
    #Else
        On Error GoTo StartFail
        Set this.hostApp = New Excel.Application
        this.ownsInstance = True

        With this.hostApp
            .Visible = False
            .DisplayAlerts = False
            .EnableEvents = False
            .AskToUpdateLinks = False
            .AutomationSecurity = msoAutomationSecurityLow
            'Off for the whole build: an invisible instance that repaints
            'keeps shared graphics memory on every repaint.
            .ScreenUpdating = False
            this.windowHandle = .Hwnd
        End With
        On Error GoTo 0
        Exit Sub

StartFail:
        failMessage = Err.Description
        On Error Resume Next
        If Not this.hostApp Is Nothing Then this.hostApp.Quit
        Set this.hostApp = Nothing
        this.ownsInstance = False
        this.windowHandle = 0
        On Error GoTo 0
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "The build instance could not be started: " & failMessage
    #End If
End Sub

'@label:open-designer-copy
'@sub-title Put the designer into the instance
'@details
'In place the designer itself is the workbook the steps run inside, and
'the folder is ignored. On the instance path the designer is written as
'__designer_copy.xlsb in the folder through SaveCopyAs, which takes the
'unsaved entries, formatter, passwords and translations with it and leaves
'the open designer alone, then the copy is opened read-write in the
'instance with links left alone, and Calculation goes manual there.
'@param copyFolder String. The folder the copy is written in, on the instance path. Must exist.
'@return Workbook. The workbook the steps run inside.
'@throws ProjectError.ErrorUnexpectedState When Acquire has not run, or the copy would not open.
'@throws ProjectError.InvalidArgument When the folder is empty or missing, on the instance path.
'@export
Public Function OpenDesignerCopy(Optional ByVal copyFolder As String = vbNullString) As Workbook
    Dim failMessage As String
    Dim failNumber As Long

    RequireAcquired "OpenDesignerCopy"

    If this.inPlace Then
        Set this.designerCopy = this.designerBook
        this.copyPath = vbNullString
        Set OpenDesignerCopy = this.designerCopy
        Exit Function
    End If

    If Not this.designerCopy Is Nothing Then
        ThrowError ProjectError.ErrorUnexpectedState, "The designer copy is already open"
    End If

    this.copyPath = ResolveCopyPath(copyFolder)

    On Error GoTo OpenFail
    'An earlier copy left behind by a run that never released is scratch
    DeleteCopyFile

    this.designerBook.SaveCopyAs this.copyPath

    Set this.designerCopy = this.hostApp.Workbooks.Open(FileName:=this.copyPath, _
                                                        UpdateLinks:=0, _
                                                        ReadOnly:=False)

    'The write raises on an instance holding no workbook, so it waits for
    'the copy
    this.hostApp.Calculation = xlCalculationManual
    On Error GoTo 0

    Set OpenDesignerCopy = this.designerCopy
    Exit Function

OpenFail:
    failNumber = Err.Number
    failMessage = Err.Description
    On Error Resume Next
    If Not this.designerCopy Is Nothing Then this.designerCopy.Close SaveChanges:=False
    Set this.designerCopy = Nothing
    DeleteCopyFile
    On Error GoTo 0
    ThrowError ProjectError.ErrorUnexpectedState, _
               "The designer copy could not be opened in the build instance (" & _
               CStr(failNumber) & "): " & failMessage
End Function

'@label:resolve-copy-path
'@sub-title The path of the copy inside the folder
'@param copyFolder String. The folder, with or without its trailing separator.
'@return String. The full path of the copy file.
'@throws ProjectError.InvalidArgument When the folder is empty or missing.
Private Function ResolveCopyPath(ByVal copyFolder As String) As String
    Dim folder As String
    Dim separator As String

    folder = Trim$(copyFolder)
    If LenB(folder) = 0 Then
        ThrowError ProjectError.InvalidArgument, "The copy folder is required on the instance path"
    End If

    separator = this.designerBook.Application.PathSeparator
    If Right$(folder, 1) <> separator Then folder = folder & separator

    If Not FolderExists(folder) Then
        ThrowError ProjectError.InvalidArgument, "The copy folder does not exist: " & folder
    End If

    ResolveCopyPath = folder & COPY_FILE_NAME
End Function

'@label:folder-exists
'@sub-title Whether a path names an existing folder
'@param folderPath String. The path.
'@return Boolean.
Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim attributes As Long

    On Error Resume Next
    Err.Clear
    attributes = GetAttr(folderPath)
    If Err.Number = 0 Then FolderExists = ((attributes And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

'@label:delete-copy-file
'@sub-title Remove the copy file from disk, quietly
'@details
'A file the instance still holds open refuses the delete, and that is
'left as it is: the file is scratch and the next run overwrites it.
Private Sub DeleteCopyFile()
    If LenB(this.copyPath) = 0 Then Exit Sub

    On Error Resume Next
    If LenB(Dir$(this.copyPath)) > 0 Then Kill this.copyPath
    On Error GoTo 0
End Sub


'@section The file names
'===============================================================================

'@label:check-open-names
'@sub-title Apply the file-name rules to the files a build opens
'@details
'Excel refuses to open a file whose name matches one already open in the
'same instance, and to save onto a file open anywhere. Two rules: two
'files of the list sharing a name is a refusal naming both paths, and a
'name already open in the target instance is a refusal naming the open
'file. Names compare without regard to case; empty entries are skipped.
'The target is the instance the build runs in once Acquire ran, or the
'designer's own before, so on Mac the designer itself is one more open
'name and on Windows the copy is. A driver checks the visible side by
'passing its own Application.
'@param paths BetterArray. Full paths, or bare file names. Nothing checks nothing.
'@param target Optional Application. The instance to check against. Nothing takes the host's.
'@return String. Empty when nothing clashes; otherwise one refusal per line, vbLf between them.
'@export
Public Function CheckOpenNames(ByVal paths As BetterArray, _
                               Optional ByVal target As Application = Nothing) As String
    Dim checkedApp As Application
    Dim clashes As String
    Dim index As Long
    Dim earlier As Long
    Dim pathText As String
    Dim baseName As String
    Dim openBook As Workbook

    If paths Is Nothing Then Exit Function

    Set checkedApp = target
    If checkedApp Is Nothing Then Set checkedApp = this.hostApp
    If checkedApp Is Nothing Then Set checkedApp = this.designerBook.Application

    For index = paths.LowerBound To paths.UpperBound
        pathText = Trim$(CStr(paths.Item(index)))
        If LenB(pathText) > 0 Then
            baseName = FileNameOf(pathText)

            'Against the earlier entries of the list
            For earlier = paths.LowerBound To index - 1
                If SameName(baseName, FileNameOf(Trim$(CStr(paths.Item(earlier))))) Then
                    AppendLine clashes, "Two files of this build share the name " & baseName & _
                                        ": " & Trim$(CStr(paths.Item(earlier))) & " and " & _
                                        pathText & ". Excel opens one file of a name at a time."
                End If
            Next earlier

            'Against the workbooks open in the target instance
            For Each openBook In checkedApp.Workbooks
                If SameName(baseName, openBook.Name) Then
                    AppendLine clashes, "The file " & baseName & " is already open in " & _
                                        InstanceLabel(checkedApp) & ": " & openBook.FullName & _
                                        ". Close it before the build."
                End If
            Next openBook
        End If
    Next index

    CheckOpenNames = clashes
End Function

'@label:file-name-of
'@sub-title The file name at the end of a path
'@details
'Both separators are read, because a Mac path carries a slash and a
'Windows path may carry either.
'@param pathText String. A full path or a bare name.
'@return String. The part after the last separator.
Private Function FileNameOf(ByVal pathText As String) As String
    Dim cut As Long
    Dim slashAt As Long

    cut = InStrRev(pathText, "\")
    slashAt = InStrRev(pathText, "/")
    If slashAt > cut Then cut = slashAt

    FileNameOf = Mid$(pathText, cut + 1)
End Function

'@label:same-name
'@sub-title Whether two file names match without regard to case
'@param first String.
'@param second String.
'@return Boolean.
Private Function SameName(ByVal first As String, ByVal second As String) As Boolean
    If LenB(first) = 0 Or LenB(second) = 0 Then Exit Function
    SameName = (StrComp(first, second, vbTextCompare) = 0)
End Function

'@label:instance-label
'@sub-title Name an instance for a refusal message
'@param checkedApp Application. The instance.
'@return String. "the build instance" for the owned one, "this Excel" otherwise.
Private Function InstanceLabel(ByVal checkedApp As Application) As String
    If this.ownsInstance And (checkedApp Is this.hostApp) Then
        InstanceLabel = "the build instance"
    Else
        InstanceLabel = "this Excel"
    End If
End Function

'@label:append-line
'@sub-title Add a line to a text, vbLf between the lines
'@param text String. ByRef. The text so far.
'@param lineText String. The line.
Private Sub AppendLine(ByRef text As String, ByVal lineText As String)
    If LenB(text) > 0 Then text = text & vbLf
    text = text & lineText
End Sub


'@section The steps
'===============================================================================

'@label:run
'@sub-title Run one build step inside the workbook the steps run in
'@details
'Application.Run over the step name, qualified with the workbook the
'steps run inside: the copy in the instance, the designer in place. The
'name may carry its module, as in "BuildSteps.BuildSheet". Up to three
'arguments are forwarded; across the processes they have to be strings
'or numbers, since a workbook object cannot cross. The step's answer
'comes back as text. A raise on the visible side comes back as an
'outcome string too: the instance gone is "the build instance stopped
'answering after <step>" and marks the instance stopped; any other fault
'is the plain error outcome. Once the instance is marked stopped every
'later Run answers the stopped outcome without a call.
'
'A call this Excel stopped waiting for (cancelled, rejected, or told to
'come back later) is answered from the copy: the step carries on over
'there, so Run waits for the instance to come free and reads back what
'the step answered, through LastAnswer. The driver then sees the answer
'of the step itself. On the instance path ScreenUpdating is written off
'before the step, whatever the step before it left.
'@param stepName String. The public function to run.
'@param arguments ParamArray. Up to three arguments for the step.
'@return String. What the step answered, or an "ERROR ..." outcome.
'@throws ProjectError.ErrorUnexpectedState When OpenDesignerCopy has not run.
'@throws ProjectError.InvalidArgument When more than three arguments are handed over.
'@export
Public Function Run(ByVal stepName As String, ParamArray arguments() As Variant) As String
    Dim macroName As String
    Dim answer As Variant
    Dim argumentCount As Long
    Dim stepLabel As String

    RequireCopy "Run"
    this.lastStep = stepName

    If this.instanceStopped Then
        Run = StoppedOutcome(ERR_REMOTE_SERVER_GONE)
        Exit Function
    End If

    argumentCount = UBound(arguments) - LBound(arguments) + 1
    If argumentCount > MAX_RUN_ARGUMENTS Then
        ThrowError ProjectError.InvalidArgument, _
                   "Run forwards at most " & CStr(MAX_RUN_ARGUMENTS) & " arguments to a step"
    End If

    'The label the copy files this step under, for the read back of a
    'dropped call
    If argumentCount = 1 Then
        stepLabel = StepLabelOf(stepName, arguments(LBound(arguments)))
    Else
        stepLabel = StepLabelOf(stepName)
    End If

    'The handler is armed before the first call on the copy: reading its
    'name is already a call across the processes, and on an instance gone
    'it is the call that raises.
    On Error GoTo RunFail
    macroName = "'" & this.designerCopy.Name & "'!" & stepName

    'An invisible instance that repaints keeps graphics memory on every
    'repaint, so the screen is off before every step.
    If Not this.inPlace Then this.hostApp.ScreenUpdating = False

    Select Case argumentCount
        Case 0
            answer = this.hostApp.Run(macroName)
        Case 1
            answer = this.hostApp.Run(macroName, arguments(LBound(arguments)))
        Case 2
            answer = this.hostApp.Run(macroName, arguments(LBound(arguments)), _
                                      arguments(LBound(arguments) + 1))
        Case Else
            answer = this.hostApp.Run(macroName, arguments(LBound(arguments)), _
                                      arguments(LBound(arguments) + 1), _
                                      arguments(LBound(arguments) + 2))
    End Select
    On Error GoTo 0

    'A Sub answers Empty; a step answers its outcome text
    If IsEmpty(answer) Or IsNull(answer) Then
        Run = vbNullString
    Else
        Run = CStr(answer)
    End If
    Exit Function

RunFail:
    Run = FailureOutcome(Err.Number, Err.Source, Err.Description, stepLabel)
End Function

'@label:paint-instance
'@sub-title Let the hidden instance paint its windows once
'@details
'ScreenUpdating on, then off, as two calls from this side with the
'instance idle between them. That one paint is what makes the window of
'the output workbook take its freezes for the rest of the build; see the
'note at the top of this class. In place nothing is written: the designer's
'own Excel paints on its own and its windows take the freeze. The second
'write runs even when the first one raised, so the screen of the instance
'is off again on every exit.
'@return String. "OK", the stopped outcome when the instance is gone, or "ERROR <n> (GenerationHost.PaintInstance): <desc>".
'@throws ProjectError.ErrorUnexpectedState When OpenDesignerCopy has not run.
'@export
Public Function PaintInstance() As String
    Dim failNumber As Long
    Dim failDesc As String

    RequireCopy "PaintInstance"

    If this.inPlace Then
        PaintInstance = OUTCOME_OK
        Exit Function
    End If

    If this.instanceStopped Then
        PaintInstance = StoppedOutcome(ERR_REMOTE_SERVER_GONE)
        Exit Function
    End If

    On Error Resume Next
        Err.Clear
        this.hostApp.ScreenUpdating = True
        failNumber = Err.Number
        failDesc = Err.Description
        Err.Clear
        this.hostApp.ScreenUpdating = False
        If failNumber = 0 Then
            failNumber = Err.Number
            failDesc = Err.Description
        End If
        Err.Clear
    On Error GoTo 0

    If failNumber = 0 Then
        PaintInstance = OUTCOME_OK
    ElseIf IsInstanceGoneError(failNumber) Then
        this.instanceStopped = True
        PaintInstance = StoppedOutcome(failNumber)
    Else
        PaintInstance = OUTCOME_ERROR_LEAD & CStr(failNumber) & " (" & CLASS_NAME & _
                        ".PaintInstance): " & failDesc
    End If
End Function

'@label:last-answer
'@sub-title Read back what the last step of the copy answered
'@details
'Asks the step named by LastAnswerStep, which answers "<step>=<answer>"
'for the last step the copy ran, and answers the second part when the
'first names the step asked about. A step is filed under its own name
'with its module dropped, and under "<name> <n>" when it took one whole
'number, so "BuildSteps.BuildSheet" with 2 is read as "BuildSheet 2".
'While the instance is still at work the question waits, up to thirty
'minutes. Run calls this after a dropped call; a driver may call it too.
'@param stepName String. The step whose answer is wanted, as Run was given it.
'@param argument Optional Variant. The one argument of that step, when it took one.
'@return String. What the step answered, or an "ERROR ..." outcome when the answer could not be read back.
'@throws ProjectError.ErrorUnexpectedState When OpenDesignerCopy has not run.
'@export
Public Function LastAnswer(ByVal stepName As String, Optional ByVal argument As Variant) As String
    Dim answer As String

    RequireCopy "LastAnswer"

    Select Case AwaitLastAnswer(StepLabelOf(stepName, argument), answer)
        Case AWAIT_ANSWERED
            LastAnswer = answer
        Case AWAIT_INSTANCE_GONE
            this.instanceStopped = True
            LastAnswer = StoppedOutcome(ERR_REMOTE_SERVER_GONE)
        Case Else
            LastAnswer = OUTCOME_ERROR_LEAD & CStr(ProjectError.ErrorUnexpectedState) & " (" & _
                         CLASS_NAME & ".LastAnswer): the answer of " & stepName & _
                         " could not be read back: " & answer
    End Select
End Function

'@label:failure-outcome
'@sub-title Shape the outcome of a Run that raised on the visible side
'@details
'Three cases. The instance gone answers the stopped outcome. A call
'dropped on this side while the instance was at work is answered with
'what the step answered over there, read back from the copy. Any other
'fault answers the plain error outcome.
'@param errNumber Long. The error number.
'@param errSource String. The error source. Empty takes this class.
'@param errDesc String. The description.
'@param stepLabel String. The label the copy files the step under.
'@return String. The stopped outcome, the answer read back, or the plain error outcome.
Private Function FailureOutcome(ByVal errNumber As Long, _
                                ByVal errSource As String, _
                                ByVal errDesc As String, _
                                ByVal stepLabel As String) As String
    Dim sourceText As String

    If IsInstanceGoneError(errNumber) Then
        this.instanceStopped = True
        FailureOutcome = StoppedOutcome(errNumber)
        Exit Function
    End If

    If IsDroppedCallError(errNumber) And Not this.inPlace Then
        FailureOutcome = AnswerOfDroppedCall(errNumber, stepLabel)
        Exit Function
    End If

    sourceText = errSource
    If LenB(sourceText) = 0 Then sourceText = CLASS_NAME & ".Run"

    FailureOutcome = OUTCOME_ERROR_LEAD & CStr(errNumber) & " (" & sourceText & "): " & errDesc
End Function

'@label:answer-of-dropped-call
'@sub-title Answer a Run whose call was dropped while the instance was at work
'@details
'The step is still running over there, or has finished. Its answer is read
'back once the instance comes free and counted on RecoveredCalls. An
'instance that went away meanwhile answers the stopped outcome. An answer
'that cannot be read back answers an error outcome carrying the number of
'the dropped call and the reason.
'@param errNumber Long. The number the dropped call raised.
'@param stepLabel String. The label the copy files the step under.
'@return String. What the step answered, or an "ERROR ..." outcome.
Private Function AnswerOfDroppedCall(ByVal errNumber As Long, _
                                     ByVal stepLabel As String) As String
    Dim answer As String

    Select Case AwaitLastAnswer(stepLabel, answer)
        Case AWAIT_ANSWERED
            this.recoveredCalls = this.recoveredCalls + 1
            AnswerOfDroppedCall = answer
        Case AWAIT_INSTANCE_GONE
            this.instanceStopped = True
            AnswerOfDroppedCall = StoppedOutcome(errNumber)
        Case Else
            AnswerOfDroppedCall = OUTCOME_ERROR_LEAD & CStr(errNumber) & " (" & CLASS_NAME & _
                                  ".Run): this Excel stopped waiting for " & this.lastStep & _
                                  " and its answer could not be read back: " & answer
    End Select
End Function

'@label:await-last-answer
'@sub-title Ask the copy for its last answer until the instance answers
'@details
'One call on the step named by LastAnswerStep. While the instance is still
'inside the step, that call waits its turn and may be dropped like the
'first one, so it is asked again after a short pause, for thirty minutes
'at most. An instance gone ends the wait; so does any other fault, a copy
'that carries no such step being the common one.
'@param stepLabel String. The label the answer has to be filed under.
'@param answer String. Answers what the step answered, or why it could not be read.
'@return Byte. AWAIT_ANSWERED, AWAIT_INSTANCE_GONE or AWAIT_UNREADABLE.
Private Function AwaitLastAnswer(ByVal stepLabel As String, ByRef answer As String) As Byte
    Dim startedAt As Date
    Dim macroName As String
    Dim filed As Variant
    Dim callNumber As Long
    Dim callDesc As String

    startedAt = Now
    Do
        filed = Empty

        On Error Resume Next
            Err.Clear
            macroName = "'" & this.designerCopy.Name & "'!" & this.lastAnswerStep
            filed = this.hostApp.Run(macroName)
            callNumber = Err.Number
            callDesc = Err.Description
            Err.Clear
        On Error GoTo 0

        If callNumber = 0 Then
            AwaitLastAnswer = ReadFiledAnswer(filed, stepLabel, answer)
            Exit Function
        End If

        If IsInstanceGoneError(callNumber) Then
            AwaitLastAnswer = AWAIT_INSTANCE_GONE
            Exit Function
        End If

        If Not IsDroppedCallError(callNumber) Then
            answer = CStr(callNumber) & " " & callDesc
            AwaitLastAnswer = AWAIT_UNREADABLE
            Exit Function
        End If

        PauseThisExcel
    Loop While DateDiff("s", startedAt, Now) < RECOVERY_SECONDS

    answer = "the build instance stayed busy for " & CStr(RECOVERY_SECONDS) & " seconds"
    AwaitLastAnswer = AWAIT_UNREADABLE
End Function

'@label:read-filed-answer
'@sub-title Take the answer out of "<step>=<answer>" when the step is the one asked about
'@param filed Variant. What the last-answer step answered.
'@param stepLabel String. The label the answer has to be filed under.
'@param answer String. Answers what the step answered, or why it could not be read.
'@return Byte. AWAIT_ANSWERED or AWAIT_UNREADABLE.
Private Function ReadFiledAnswer(ByVal filed As Variant, _
                                 ByVal stepLabel As String, _
                                 ByRef answer As String) As Byte
    Dim filedText As String
    Dim markAt As Long
    Dim filedLabel As String

    If Not (IsEmpty(filed) Or IsNull(filed)) Then filedText = CStr(filed)

    markAt = InStr(filedText, LAST_ANSWER_MARK)
    If markAt = 0 Then
        answer = "the copy filed no step"
        ReadFiledAnswer = AWAIT_UNREADABLE
        Exit Function
    End If

    filedLabel = Left$(filedText, markAt - 1)
    If StrComp(filedLabel, stepLabel, vbTextCompare) <> 0 Then
        If LenB(filedLabel) = 0 Then filedLabel = "(none)"
        answer = "the last step the copy ran is " & filedLabel
        ReadFiledAnswer = AWAIT_UNREADABLE
        Exit Function
    End If

    answer = Mid$(filedText, markAt + Len(LAST_ANSWER_MARK))
    ReadFiledAnswer = AWAIT_ANSWERED
End Function

'@label:step-label-of
'@sub-title The label the copy files a step under
'@details
'The name of the step with its module dropped, and the argument after a
'space when the step took one whole number: "BuildSheet 2".
'@param stepName String. The step, with or without its module.
'@param argument Optional Variant. The one argument of the step, when it took one.
'@return String. The label.
Private Function StepLabelOf(ByVal stepName As String, Optional ByVal argument As Variant) As String
    Dim label As String

    label = Mid$(stepName, InStrRev(stepName, ".") + 1)

    If Not IsMissing(argument) Then
        Select Case VarType(argument)
            Case vbByte, vbInteger, vbLong
                label = label & " " & CStr(argument)
        End Select
    End If

    StepLabelOf = label
End Function

'@label:is-dropped-call-error
'@sub-title Whether an error number says this side dropped the call
'@details
'Cancelled by the message filter of this Excel, which is what a long step
'ends on once the "waiting for another application" box has been up;
'rejected by the instance; or told to come back later. In all three the
'instance is alive.
'@param errNumber Long.
'@return Boolean.
Private Function IsDroppedCallError(ByVal errNumber As Long) As Boolean
    Select Case errNumber
        Case ERR_RPC_CALL_CANCELED, ERR_RPC_CALL_REJECTED, ERR_RPC_RETRY_LATER
            IsDroppedCallError = True
        Case Else
            IsDroppedCallError = False
    End Select
End Function

'@label:pause-this-excel
'@sub-title Hold the visible Excel for the pause between two tries
Private Sub PauseThisExcel()
    On Error Resume Next
    this.designerBook.Application.Wait Now + TimeSerial(0, 0, RECOVERY_PAUSE_SECONDS)
    On Error GoTo 0
End Sub

'@label:is-instance-gone-error
'@sub-title Whether an error number says the instance is gone
'@details
'The three shapes the spike measured: 462 on every call after a quit or a
'kill, the RPC call failed on the call the kill interrupted, and the RPC
'server unavailable. Plus the one the suite met: the object disconnected
'from its clients, on the first call after a quit.
'@param errNumber Long.
'@return Boolean.
Private Function IsInstanceGoneError(ByVal errNumber As Long) As Boolean
    Select Case errNumber
        Case ERR_REMOTE_SERVER_GONE, ERR_RPC_CALL_FAILED, ERR_RPC_SERVER_UNAVAILABLE, _
             ERR_RPC_DISCONNECTED
            IsInstanceGoneError = True
        Case Else
            IsInstanceGoneError = False
    End Select
End Function

'@label:stopped-outcome
'@sub-title The outcome of a Run whose instance is gone
'@param errNumber Long. The number that said so.
'@return String. "ERROR <n> (GenerationHost.Run): the build instance stopped answering after <step>".
Private Function StoppedOutcome(ByVal errNumber As Long) As String
    StoppedOutcome = OUTCOME_ERROR_LEAD & CStr(errNumber) & " (" & CLASS_NAME & ".Run): " & _
                     "the build instance stopped answering after " & this.lastStep
End Function


'@section ReleaseInstance
'===============================================================================

'@label:release
'@sub-title Close the copy, quit the owned instance, delete the copy file
'@details
'In place nothing is closed and nothing quits: the designer stays where
'it is. On the instance path the copy is closed without saving and the
'instance this object started is quit; an instance a Run found gone is
'left alone, since there is nothing left to quit. The whole body runs
'under On Error Resume Next, because it runs from the failure exits and
'from Class_Terminate. A Quit that fails answers the window handle for
'the driver to log; the answer is kept on ReleaseOutcome too. Calling
'ReleaseInstance twice is one release.
'@return String. "OK", or "ERROR <n> (GenerationHost.ReleaseInstance): ..." naming the window handle.
'@export
Public Function ReleaseInstance() As String
    Dim outcome As String
    Dim quitNumber As Long
    Dim quitMessage As String

    outcome = OUTCOME_OK

    On Error Resume Next
    If this.inPlace Then
        Set this.designerCopy = Nothing
    Else
        If Not this.designerCopy Is Nothing Then
            If Not this.instanceStopped Then this.designerCopy.Close SaveChanges:=False
        End If
        Set this.designerCopy = Nothing

        If this.ownsInstance And Not this.hostApp Is Nothing Then
            If this.instanceStopped Then
                outcome = OUTCOME_ERROR_LEAD & CStr(ERR_REMOTE_SERVER_GONE) & " (" & CLASS_NAME & _
                          ".ReleaseInstance): the build instance " & CStr(this.windowHandle) & _
                          " stopped answering after " & this.lastStep & " and was left alone"
            Else
                Err.Clear
                this.hostApp.Quit
                quitNumber = Err.Number
                quitMessage = Err.Description
                If quitNumber <> 0 Then
                    outcome = OUTCOME_ERROR_LEAD & CStr(quitNumber) & " (" & CLASS_NAME & _
                              ".ReleaseInstance): the build instance " & CStr(this.windowHandle) & _
                              " could not be quit: " & quitMessage
                End If
            End If
        End If

        DeleteCopyFile
    End If

    Set this.hostApp = Nothing
    this.ownsInstance = False
    this.acquired = False
    this.copyPath = vbNullString
    On Error GoTo 0

    this.releaseOutcome = outcome
    ReleaseInstance = outcome
End Function

'@label:class-terminate
'@sub-title ReleaseInstance the instance when the object is dropped
'@details
'A driver that raises on its way out drops this object without calling
'ReleaseInstance, and the hidden instance would live on with nobody to quit it.
Private Sub Class_Terminate()
    On Error Resume Next
    If this.acquired Then ReleaseInstance
    On Error GoTo 0
End Sub


'@section Guards
'===============================================================================

'@label:require-acquired
'@sub-title Raise when Acquire has not run
'@param memberName String. The member asking.
Private Sub RequireAcquired(ByVal memberName As String)
    If Not this.acquired Then
        ThrowError ProjectError.ErrorUnexpectedState, memberName & " needs Acquire to have run first"
    End If
End Sub

'@label:require-copy
'@sub-title Raise when OpenDesignerCopy has not run
'@param memberName String. The member asking.
Private Sub RequireCopy(ByVal memberName As String)
    RequireAcquired memberName
    If this.designerCopy Is Nothing Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                   memberName & " needs OpenDesignerCopy to have run first"
    End If
End Sub

'@label:throw-error
'@sub-title Raise a typed project error
'@param errNumber ProjectError. The error code.
'@param message String. Descriptive error message.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASS_NAME, message
End Sub
