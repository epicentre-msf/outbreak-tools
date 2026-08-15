VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Dated record of user actions on the __log worksheet"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Dated record of user actions on the __log worksheet")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@depends BetterArray, Checking, CheckingOutput, HiddenNames

Option Explicit

'@class LLLog
'@description
'Keeps the dated record of the important user actions of a running
'linelist on the very hidden __log worksheet. Each event is one Checking
'bundle: the bundle title is the SECTION the action belongs to, the entry
'label carries the timestamp, the action code and a short detail, and the
'scope carries the outcome. Every bundle is flushed through
'CheckingOutput.PrintOutput, which appends across renders, so the sheet
'grows one line per event and the built-in status and title dropdowns of
'the output sheet filter the log by outcome and by section.
'
'THE THREE SECTIONS
'-------------------------------------------------------------------------------
'The log is read by three sections, and the title of an entry is the
'section the action belongs to. The writer runs in compact mode, so a
'section title is written the first time the sheet carries it and never
'again: a session of a hundred actions reads as a hundred lines under at
'most three titles.
'
'  open/close          the workbook open, the workbook close, a log restart
'  data input/output   every import, every export, and clearing the data
'  linelist lifecycle  everything else the user does inside the workbook
'
'An action code no case names lands in linelist lifecycle, so a new logged
'action reaches the sheet without touching this class. The action code
'itself moves down to the entry line, beside the timestamp, where it stays
'readable without heading a block of its own.
'
'A title heads a block of its own kind. An action coming back to a section
'written earlier lands at the foot of THAT block, not at the foot of the
'sheet, so the three blocks stay whole however far apart in time their
'actions were taken, and a blank row stands between them. Inside a block
'the lines keep the order they happened. Each line also carries its section
'in the hidden title column, and that is what the title dropdown of the
'output sheet filters on, so reading the log one section at a time is a
'dropdown away.
'
'The last thing that happened is therefore no longer the last line of the
'sheet: a close lands under the open it belongs with, near the top, while
'the lifecycle block below carries the bulk of the session.
'
'PROVISIONING
'-------------------------------------------------------------------------------
'Create(wkb) binds to the __log worksheet of the given workbook and builds
'it very hidden when the workbook has none. An old linelist grows the
'sheet on its first logged event, so the generation tree stays untouched.
'Building a sheet wants the workbook structure unprotected, which is the
'caller''s duty.
'
'ROTATION
'-------------------------------------------------------------------------------
'LOG_MAX_ENTRIES is read as a row cap. A flush that finds the last
'written row past the cap clears the sheet and removes the row marker
'CheckingOutput keeps, so the next render re-initialises the sheet, and
'the fresh log opens with one info line ("restarted").
'
'FILTER DROPDOWNS
'-------------------------------------------------------------------------------
'The interactive row hiding of the output sheet wants a Worksheet_Change
'handler injected through the VBE, and injecting code wants VBE trust,
'which field machines lack. LLLog leaves the worksheet code module alone;
'the dropdowns stay plain validation cells and the filtering is reachable
'through CheckingOutput.FilterWorksheet.
'@version 1.0 (2026-08-12)


'@section State
'===============================================================================
'@description Backing UDT and class constants.

Private Const CLASSNAME As String = "LLLog"
Private Const LOG_SHEET_NAME As String = "__log"
Private Const LOG_MAX_ENTRIES As Long = 10000
Private Const OUTPUT_TITLE As String = "log"

'The separator CheckingOutput splits a label on when writing a row. The
'timestamp and the detail sit on each side of it, so they land in two
'columns of the log sheet.
Private Const DETAIL_SEPARATOR As String = "--"

'The column CheckingOutput writes its first visible value into. The last
'written row of that column is what the rotation reads against the cap.
Private Const OUTPUT_COLUMN As Long = 3

'The name mirrors the private row marker CheckingOutput keeps on the log
'sheet. Removing that marker is what makes CheckingOutput re-initialise
'the sheet on the next render.
Private Const ROW_MARKER_NAME As String = "CheckingOutputStartingRow"

Private Const RESTART_ACTION As String = "log"
Private Const RESTART_DETAIL As String = "restarted"

'The three sections the log is read by. Each is the title of every entry
'that belongs to it, so the title dropdown of the output sheet filters the
'log down to one of the three.
Private Const SECTION_OPENCLOSE As String = "open/close"
Private Const SECTION_DATAIO As String = "data input/output"
Private Const SECTION_LIFECYCLE As String = "linelist lifecycle"

'What sits between the action code and the detail on an entry line. The
'two share the last written column, so they need a mark of their own.
Private Const ACTION_SEPARATOR As String = ": "

Private Type TLLLog
    targetSheet As Worksheet
    writer As CheckingOutput
    store As HiddenNames
    sealed As Boolean
    platformTag As String
End Type

Private this As TLLLog

'@section Factory
'===============================================================================
'@description Entry point for creating LLLog instances.

'@label:create
'@sub-title Create a log store bound to the __log worksheet of a workbook
'@details
'Binds to the __log worksheet of the given workbook. When the workbook
'has none, the sheet is appended after the last worksheet and made very
'hidden, so an old linelist grows its log on first use.
'@param wkb Workbook. The linelist workbook carrying the log.
'@return LLLog. Fully initialised instance.
'@throws ProjectError.ObjectNotInitialized When wkb is Nothing.
'@export
Public Function Create(ByVal wkb As Workbook) As LLLog
    If wkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
         "Unable to create an instance of the class, please provide a workbook"
    End If

    Dim newLog As LLLog
    Set newLog = New LLLog

    Set newLog.Wksh = ResolveLogSheet(wkb)
    newLog.Seal
    Set Create = newLog
End Function

'@section PublicAccessors
'===============================================================================
'@description Properties that expose internal state.

'@label:wksh
'@prop-title Retrieve the log worksheet backing this store
'@return Worksheet. Reference to the log worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.targetSheet
End Property

'@label:wksh-set
'@prop-title Define the worksheet that will receive the log
'@param sh Worksheet. Target worksheet to use.
Public Property Set Wksh(ByVal sh As Worksheet)
    GuardNotSealed "Wksh"
    Set this.targetSheet = sh
End Property

'@label:sheet-name
'@prop-title Name of the log worksheet
'@details
'Answers the internal sheet name so the viewing surface and the close
'guard reach the sheet through the class.
'@return String. The log worksheet name.
'@export
Public Property Get SheetName() As String
    SheetName = LOG_SHEET_NAME
End Property

'@label:section-of
'@sub-title The section an action code is logged under
'@details
'The three sections are the whole title vocabulary of the log sheet. The
'workbook open, the workbook close and a log restart are the session
'boundaries and read as open/close; everything that moves data in or out
'of the workbook reads as data input/output; every other action the user
'takes inside the workbook reads as linelist lifecycle, which is also
'where an action code named by no case below lands.
'@param action String. The action code of the event.
'@return String. One of the three section titles.
'@export
Public Function SectionOf(ByVal action As String) As String
    Select Case LCase$(Trim$(action))

    Case "open", "close", RESTART_ACTION
        SectionOf = SECTION_OPENCLOSE

    Case "import-data", "import-geobase", "clear-data", _
         "export", "export-migration", "export-other", "export-analysis"
        SectionOf = SECTION_DATAIO

    Case Else
        SectionOf = SECTION_LIFECYCLE

    End Select
End Function

'@label:max-entries
'@prop-title Row cap of the log sheet
'@details
'Answers the cap the rotation reads, so a test can seed a row past it
'without writing ten thousand lines.
'@return Long. The row cap.
'@export
Public Property Get MaxEntries() As Long
    MaxEntries = LOG_MAX_ENTRIES
End Property

'@section Logging
'===============================================================================
'@description One method per outcome, each flushing one Checking bundle.

'@label:log-success
'@sub-title Log an action that ended well
'@param action String. The action code, used as the bundle title.
'@param detail Optional String. Short free text beside the timestamp.
'@export
Public Sub LogSuccess(ByVal action As String, Optional ByVal detail As String = vbNullString)
    WriteEntry action, detail, checkingSuccess
End Sub

'@label:log-warning
'@sub-title Log a refusal or a degraded run
'@param action String. The action code, used as the bundle title.
'@param detail Optional String. Short free text beside the timestamp.
'@export
Public Sub LogWarning(ByVal action As String, Optional ByVal detail As String = vbNullString)
    WriteEntry action, detail, checkingWarning
End Sub

'@label:log-failure
'@sub-title Log an action that failed
'@param action String. The action code, used as the bundle title.
'@param detail Optional String. Short free text beside the timestamp.
'@export
Public Sub LogFailure(ByVal action As String, Optional ByVal detail As String = vbNullString)
    WriteEntry action, detail, checkingError
End Sub

'@label:log-info
'@sub-title Log a lifecycle line
'@param action String. The action code, used as the bundle title.
'@param detail Optional String. Short free text beside the timestamp.
'@export
Public Sub LogInfo(ByVal action As String, Optional ByVal detail As String = vbNullString)
    WriteEntry action, detail, checkingInfo
End Sub

'@section Internals
'===============================================================================
'@description Sheet resolution, entry building, and rotation.

'@label:resolve-log-sheet
'@sub-title Resolve the log worksheet, building it very hidden when absent
'@details
'The sheet is appended after the last worksheet, because a bare Add puts
'a new sheet in front of whatever is active. An existing sheet keeps its
'visibility, so the viewing surface can hold it open across a bind.
'@param wkb Workbook. The linelist workbook.
'@return Worksheet. The resolved or newly created worksheet.
Private Function ResolveLogSheet(ByVal wkb As Workbook) As Worksheet
    Dim sh As Worksheet

    'The worksheets collection raises 9 on a missing name; the miss is the
    'signal to build the sheet.
    On Error Resume Next
        Set sh = wkb.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = wkb.Worksheets.Add(After:=wkb.Worksheets(wkb.Worksheets.Count))
        sh.Name = LOG_SHEET_NAME
        sh.Visible = xlSheetVeryHidden
    End If

    Set ResolveLogSheet = sh
End Function

'@label:write-entry
'@sub-title Rotate when the cap is passed, then flush one dated entry
'@param action String. The action code, used as the bundle title.
'@param detail String. Short free text beside the timestamp.
'@param scope Byte. CheckingScope value carrying the outcome.
'@throws ProjectError.InvalidArgument When the action code is empty.
Private Sub WriteEntry(ByVal action As String, ByVal detail As String, ByVal scope As Byte)
    If LenB(Trim$(action)) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
         "An action code is required to log an event"
    End If

    RotateWhenPastCap
    FlushEntry action, detail, scope
End Sub

'@label:flush-entry
'@sub-title Build one Checking bundle and flush it through the writer
'@details
'The bundle title is the section of the action, not the action itself, so
'the sheet carries three titles rather than one per action code. The
'bundle carries one key, so the key uniqueness of Checking binds inside
'the bundle and every event stays a plain append.
'@param action String. The action code, written on the entry line.
'@param detail String. Short free text beside the timestamp.
'@param scope Byte. CheckingScope value carrying the outcome.
Private Sub FlushEntry(ByVal action As String, ByVal detail As String, ByVal scope As Byte)
    Dim checks As Checking
    Dim batch As BetterArray

    Set checks = Checking.Create(SectionOf(action))
    checks.Add "entry", EntryLabel(action, detail), scope

    Set batch = New BetterArray
    batch.LowerBound = 1
    batch.Push checks

    Writer().PrintOutput batch
End Sub

'@label:entry-label
'@sub-title Compose the dated label of one entry
'@details
'The label opens with the timestamp and the platform; the action code and
'the detail follow behind the separator, so the two land in their own
'columns of the log sheet. The action code moved here when the title became
'the section: the sheet has three titles now, and the line has to say which
'action it records. A separator inside the detail would push text past the
'last written column, so it is softened to a single dash first.
'The platform rides with the timestamp rather than with the detail, because
'the detail column is what a reader searches and what several tests read
'whole. It is on every line on purpose: a log rotation clears the sheet, so
'a platform written once at the top of a session would not survive one.
'@param action String. The action code of the event.
'@param detail String. Short free text beside the timestamp.
'@return String. The composed label.
Private Function EntryLabel(ByVal action As String, ByVal detail As String) As String
    Dim label As String

    label = Format$(Now, "yyyy-mm-dd hh:mm:ss") & " " & PlatformTag() & _
            DETAIL_SEPARATOR & Trim$(action)

    If LenB(Trim$(detail)) > 0 Then
        label = label & ACTION_SEPARATOR & _
                Replace(Trim$(detail), DETAIL_SEPARATOR, "-")
    End If

    EntryLabel = label
End Function

'@label:platform-tag
'@sub-title The operating system, the bitness and the Excel version, on one tag
'@details
'A log read months later says nothing about where it was written, and the
'same code does not behave the same on the two platforms: a geobase export
'that Windows accepts has refused on this Mac. Every entry therefore names
'its platform.
'The name and the bitness come from the compile constants, so they are the
'build that is running and cannot be misread. The Excel version comes from
'the application and is read once, because it is the same for every line of
'a session and the log writes a line per user action.
'@return String. Something of the shape "mac-64 excel-16.90".
Private Function PlatformTag() As String
    Dim osName As String
    Dim bits As String

    If LenB(this.platformTag) > 0 Then
        PlatformTag = this.platformTag
        Exit Function
    End If

    #If Mac Then
        osName = "mac"
    #Else
        osName = "win"
    #End If

    #If Win64 Then
        bits = "64"
    #Else
        bits = "32"
    #End If

    this.platformTag = osName & "-" & bits & " excel-" & Application.Version
    PlatformTag = this.platformTag
End Function

'@label:rotate-when-past-cap
'@sub-title Clear the sheet and restart the log once the cap is passed
'@details
'Reads the last written row of the output column. Past the cap, the
'sheet is cleared and the row marker removed, so CheckingOutput
're-initialises on the next render, and the fresh log opens with one
'info line.
Private Sub RotateWhenPastCap()
    Dim sh As Worksheet
    Dim lastRow As Long

    Set sh = Wksh()
    lastRow = sh.Cells(sh.Rows.Count, OUTPUT_COLUMN).End(xlUp).Row
    If lastRow <= LOG_MAX_ENTRIES Then Exit Sub

    sh.Cells.Clear
    If HiddenStore().HasName(ROW_MARKER_NAME) Then _
        HiddenStore().RemoveName ROW_MARKER_NAME

    FlushEntry RESTART_ACTION, RESTART_DETAIL, checkingInfo
End Sub

'@label:writer
'@sub-title Lazily create and return the CheckingOutput writer
'@details
'The writer is asked for compact appending, which is the mode a log wants:
'a section title is written the first time the sheet carries it and never
'again, and no blank rows stand between the entries.
'@return CheckingOutput. Writer bound to the log worksheet.
Private Function Writer() As CheckingOutput
    If this.writer Is Nothing Then
        Set this.writer = CheckingOutput.Create(Wksh(), OUTPUT_TITLE, compactAppend:=True)
    End If
    Set Writer = this.writer
End Function

'@label:hidden-store
'@sub-title Lazily create and return the HiddenNames store of the log sheet
'@return HiddenNames. Hidden name store bound to the log worksheet.
Private Function HiddenStore() As HiddenNames
    If this.store Is Nothing Then
        Set this.store = HiddenNames.Create(Wksh())
    End If
    Set HiddenStore = this.store
End Function

'@section ErrorHandling
'===============================================================================
'@description Centralised error-raising pattern and the set-once seal.

'@label:throw-error
'@sub-title Raise a ProjectError-coded runtime error
'@param errNumb Long. ProjectError enumeration value.
'@param errorMessage String. Human-readable description.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub

'@label:seal
'@sub-title Seal the instance so setup setters can no longer be written
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against writes after sealing
'@param propName String. Name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub
