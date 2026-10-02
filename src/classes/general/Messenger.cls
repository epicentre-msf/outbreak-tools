VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Messenger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "The one place a message box is shown from"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("The one place a message box is shown from, so a script can drive the workbook with no box waiting for a click")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName
'@depends ProjectError, BetterArray, HiddenNames

Option Explicit

'@class Messenger
'@description
'Every message box the R package can walk into is shown from here. While the
'messenger is armed it shows nothing: it writes the text down and answers the
'caller with the answer that caller named. While it is disarmed it calls MsgBox
'with exactly what it was given and answers the click. A person opening the
'workbook and pressing buttons sees the boxes they have always seen.
'
'THE RECORD LIVES ON THE DEFAULT INSTANCE
'-------------------------------------------------------------------------------
'Callers write Messenger.Show and this class has no Create. One Create per call
'site would give each caller its own instance and its own record, and Messages
'would come back empty at the end of the run. One instance, one record.
'
'ARM AND DISARM COME IN PAIRS
'-------------------------------------------------------------------------------
'A wrapper arms the messenger as its first act and disarms it in its exit path,
'in the same Excel session. Nothing is written to the file, so a run that dies
'part way cannot leave a delivered workbook mute.
'
'Arm starts a run: it empties the record. Disarm ends the silence and keeps the
'record, because the summary a wrapper answers is read after the run is over.
'
'THE FORCE STATE AND CarryOn
'-------------------------------------------------------------------------------
'Some questions ask whether to push past a warning. A call site writes
'Messenger.CarryOn() as its silent answer to one of those: vbYes when the run
'was armed with force, vbNo otherwise. Force is the only thing CarryOn reads,
'and Disarm drops it.
'
'THE STORED SWITCH COVERS THE OPEN PATH
'-------------------------------------------------------------------------------
'A box that fires from Workbook_Open fires before any Application.Run can reach
'a wrapper, so nothing in process can arm the messenger in time.
'__OBT__SILENT_OPERATIONS__ is read off the workbook itself, and the open path
'arms on what it answers:
'
'    If Messenger.ReadStoredSwitch(ThisWorkbook) Then Messenger.Arm ThisWorkbook
'
'The read goes through HiddenNames.QuickValue, three COM crossings with no
'instance behind them. HiddenNames.Create walks the whole Names collection of
'the host, which is the largest cost that class has, and the open path reads
'one value once.
'
'A missing name reads as No, so a workbook built before the switch existed
'behaves exactly as it did.
'
'WHY A CLASS
'-------------------------------------------------------------------------------
'Six classes show a box today, and a class calling a standard module is out of
'scope (conventions.md section 9). A class serves the modules and the classes
'both.
'@depends ProjectError, BetterArray, HiddenNames
'@version 1.0 (2026-08-31)


Private Type TMessenger
    isArmed As Boolean
    isForced As Boolean
    hostBook As Workbook
    record As BetterArray
End Type

Private this As TMessenger

Private Const CLASSNAME As String = "Messenger"

'The switch is stored at workbook level as a string holding Yes or No. The
'name is settled on both sides of the wire and this is where the VBA side
'keeps it: the three workbooks seed the flag through Messenger.SwitchName.
Private Const SWITCH_NAME As String = "__OBT__SILENT_OPERATIONS__"
Private Const SWITCH_ON As String = "Yes"
Private Const SWITCH_OFF As String = "No"

'What separates a title from its message on a recorded line.
Private Const TITLE_JOIN As String = ": "


'@section Lifecycle
'===============================================================================
'The default instance is the only instance. There is no Create.

'@label:class-initialize
'@sub-title Start with an empty record and the boxes on
Private Sub Class_Initialize()
    Set this.record = New BetterArray
End Sub

'@label:reset
'@sub-title Put the messenger back to its opening state
'@details
'Empties the record, ends the silence, drops the force state and forgets the
'host. Arm calls this, so a new run never reads the messages of the run before.
'@export
Public Sub Reset()
    Set this.record = New BetterArray
    this.isArmed = False
    this.isForced = False
    Set this.hostBook = Nothing
End Sub


'@section Arming
'===============================================================================
'The pair a wrapper calls around the work it does.

'@label:arm
'@sub-title Turn the boxes off for the length of this call
'@details
'Reads nothing from the file. The wrapper decides, and the wrapper disarms in
'its exit path. The record of any run before is emptied here.
'@param hostBook Workbook. The workbook this run is about.
'@param force Boolean. Optional, default False. True makes CarryOn answer vbYes,
'which is what pushes a run past a warning that would otherwise stop it.
'@export
Public Sub Arm(ByVal hostBook As Workbook, Optional ByVal force As Boolean = False)
    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Arm needs the workbook this run is about."
    End If

    Reset
    Set this.hostBook = hostBook
    this.isForced = force
    this.isArmed = True
End Sub

'@label:disarm
'@sub-title Turn the boxes back on
'@details
'The record is kept. A wrapper disarms and then reads Messages for the summary
'it answers.
'@export
Public Sub Disarm()
    this.isArmed = False
    this.isForced = False
End Sub

'@label:armed
'@prop-title Whether the boxes are off right now
'@return Boolean. True while the messenger is armed.
'@export
Public Property Get Armed() As Boolean
    Armed = this.isArmed
End Property

'@label:host-book
'@prop-title The workbook the current run is about
'@return Workbook. What Arm was given, or Nothing.
'@export
Public Property Get HostBook() As Workbook
    Set HostBook = this.hostBook
End Property

'@label:carry-on
'@sub-title The silent answer to a question that asks whether to push on
'@details
'Written at a call site as the silent answer of a warning the run may pass
'through: Messenger.Show(text, Messenger.CarryOn(), vbYesNo). A run armed with
'force carries on. Every other run stops, which is the safe reading of a
'warning nobody is there to read.
'@return VbMsgBoxResult. vbYes when the run was forced, vbNo otherwise.
'@export
Public Function CarryOn() As VbMsgBoxResult
    CarryOn = vbNo
    If this.isForced Then CarryOn = vbYes
End Function


'@section The stored switch
'===============================================================================
'What the open path reads, before any wrapper can be called.

'@label:switch-name
'@prop-title The name the switch is stored under
'@details
'The three workbooks seed their flag through this, so the string lives in one
'place on the VBA side.
'@return String. __OBT__SILENT_OPERATIONS__
'@export
Public Property Get SwitchName() As String
    SwitchName = SWITCH_NAME
End Property

'@label:switch-on-value
'@prop-title The stored value that means silent
'@return String. Yes
'@export
Public Property Get SwitchOnValue() As String
    SwitchOnValue = SWITCH_ON
End Property

'@label:switch-off-value
'@prop-title The stored value that means the workbook behaves as it always has
'@return String. No
'@export
Public Property Get SwitchOffValue() As String
    SwitchOffValue = SWITCH_OFF
End Property

'@label:read-stored-switch
'@sub-title Read the silence switch off a workbook
'@details
'The read costs three COM crossings through HiddenNames.QuickValue. A workbook
'that holds no such name answers False, so every workbook built before the
'switch existed behaves as it did.
'
'This answers the value and changes nothing. The open path arms on it:
'
'    If Messenger.ReadStoredSwitch(ThisWorkbook) Then Messenger.Arm ThisWorkbook
'@param hostBook Workbook. The workbook holding the name.
'@return Boolean. True when the stored value reads Yes.
'@export
Public Function ReadStoredSwitch(ByVal hostBook As Workbook) As Boolean
    Dim stored As String

    If hostBook Is Nothing Then Exit Function

    stored = HiddenNames.QuickValue(hostBook, SWITCH_NAME, SWITCH_OFF)
    ReadStoredSwitch = (StrComp(Trim$(stored), SWITCH_ON, vbTextCompare) = 0)
End Function


'@section Showing
'===============================================================================

'@label:show
'@sub-title Show one box, or write it down and answer for it
'@details
'Armed: nothing is shown, the text goes into the record and silentAnswer comes
'back. Disarmed: MsgBox is called with what this was given and the click comes
'back unchanged.
'
'silentAnswer is the second parameter and it carries no default. A box that
'guards a destructive step and one that merely asks "carry on" want different
'answers, and only the call site knows which. VBA needs every optional
'parameter last, so boxStyle and title follow it, and a call site that forgets
'the answer fails to compile.
'@param message String. The text of the box.
'@param silentAnswer Long. What this answers while armed. vbOK for a box with
'one button, and for a question the answer that call site wants with nobody
'there. Messenger.CarryOn() for a warning a forced run may pass through.
'@param boxStyle Long. Optional, default vbOKOnly. The MsgBox style.
'@param title String. Optional. The box title.
'@return VbMsgBoxResult. The click while disarmed, silentAnswer while armed.
'@export
Public Function Show(ByVal message As String, _
                     ByVal silentAnswer As Long, _
                     Optional ByVal boxStyle As Long = vbOKOnly, _
                     Optional ByVal title As String = vbNullString) As VbMsgBoxResult
    If this.isArmed Then
        RecordMessage message, title
        Show = silentAnswer
        Exit Function
    End If

    If LenB(title) = 0 Then
        Show = MsgBox(message, boxStyle)
        Exit Function
    End If

    Show = MsgBox(message, boxStyle, title)
End Function


'@section The record
'===============================================================================
'A run that ate a message and said nothing is worse than the box it replaced.

'@label:messages
'@sub-title Everything the run swallowed, one line each
'@return String. The recorded lines joined by a line break. Empty when nothing
'was swallowed.
'@export
Public Function Messages() As String
    Dim counter As Long
    Dim answer As String

    If this.record Is Nothing Then Exit Function
    If this.record.Length = 0 Then Exit Function

    For counter = this.record.LowerBound To this.record.UpperBound
        If LenB(answer) > 0 Then answer = answer & vbNewLine
        answer = answer & CStr(this.record.Item(counter))
    Next counter

    Messages = answer
End Function

'@label:has-messages
'@sub-title Whether the run swallowed anything
'@return Boolean. True when the record holds at least one line.
'@export
Public Function HasMessages() As Boolean
    If this.record Is Nothing Then Exit Function
    HasMessages = (this.record.Length > 0)
End Function


'@section Private helpers
'===============================================================================

'@label:record-message
'@sub-title Write one swallowed box into the record
'@param message String. The text of the box.
'@param title String. The box title. May be empty.
Private Sub RecordMessage(ByVal message As String, ByVal title As String)
    Dim recorded As String

    recorded = SingleLine(message)
    If LenB(Trim$(title)) > 0 Then recorded = SingleLine(title) & TITLE_JOIN & recorded

    this.record.Push recorded
End Sub

'@label:single-line
'@sub-title Flatten a box text onto one line
'@details
'A box text often carries line breaks. The record answers one line per message,
'so a reader can count what a run swallowed, and every break becomes a space.
'@param boxText String. The text to flatten.
'@return String. The same text with no line break in it.
Private Function SingleLine(ByVal boxText As String) As String
    Dim flattened As String

    flattened = Replace$(boxText, vbCrLf, " ")
    flattened = Replace$(flattened, vbCr, " ")
    flattened = Replace$(flattened, vbLf, " ")

    SingleLine = Trim$(flattened)
End Function

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber Long. A ProjectError value.
'@param message String. What went wrong.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub
