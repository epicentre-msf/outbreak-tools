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
'ScreenUpdating stays True there: the output workbook refused its freeze
'under ScreenUpdating False in an invisible instance, and nothing is painted
'in an invisible instance anyway. Calculation is set to manual once the copy
'is open, because the write raises 1004 on an instance holding no workbook.
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
'through. A raise Run sees on the visible side is either
'the instance gone (462, RPC call failed, RPC server unavailable), answered
'as "the build instance stopped answering after <step>" with the instance
'marked stopped so ReleaseInstance quits nothing, or a plain fault such as a step
'name the copy does not carry, answered as "ERROR <n> (<source>): <desc>".
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

    'The handler is armed before the first call on the copy: reading its
    'name is already a call across the processes, and on an instance gone
    'it is the call that raises.
    On Error GoTo RunFail
    macroName = "'" & this.designerCopy.Name & "'!" & stepName

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
    Run = FailureOutcome(Err.Number, Err.Source, Err.Description)
End Function

'@label:failure-outcome
'@sub-title Shape the outcome of a Run that raised on the visible side
'@param errNumber Long. The error number.
'@param errSource String. The error source. Empty takes this class.
'@param errDesc String. The description.
'@return String. The stopped outcome for the instance gone, the plain error outcome otherwise.
Private Function FailureOutcome(ByVal errNumber As Long, _
                                ByVal errSource As String, _
                                ByVal errDesc As String) As String
    Dim sourceText As String

    If IsInstanceGoneError(errNumber) Then
        this.instanceStopped = True
        FailureOutcome = StoppedOutcome(errNumber)
        Exit Function
    End If

    sourceText = errSource
    If LenB(sourceText) = 0 Then sourceText = CLASS_NAME & ".Run"

    FailureOutcome = OUTCOME_ERROR_LEAD & CStr(errNumber) & " (" & sourceText & "): " & errDesc
End Function

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
