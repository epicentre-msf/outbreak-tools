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
'WHAT ONE LINE SAYS
'-------------------------------------------------------------------------------
'An entry line reads as the timestamp and the platform, then the procedure
'that raised the event, then the action code and the detail:
'
'  2026-08-15 13:55:45 mac-64 excel-16.112  ClickShowHideSection > showhide-section: Demographics hidden on sheet1
'
'The procedure is named by the caller. VBA carries no call stack to read a
'name from, and one button can raise the same refusal from four places, so
'a line that names only the action leaves the reader guessing which press
'wrote it. The name is optional: a caller that passes none writes the same
'line as before, opening at the action code.
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
'THE STOPWATCH OF A LONG WALK
'-------------------------------------------------------------------------------
'StartWalk, MarkStep, StepReport and LogSteps time a walk that runs in
'steps -- an export, an import -- and leave the whole walk on one info
'line, each step with its seconds and a total at the end. MarkStep
'answers the name it was given, so a walk that already carries the step
'it is on keeps one line per step and the name still reaches its failure
'message.
'The stopwatch is state on the instance, so one instance times one walk.
'One line rather than one per step: a migration export names a dozen
'steps, and a line each would push the row cap twelve times faster for
'the same reading.
'
'FILTER DROPDOWNS
'-------------------------------------------------------------------------------
'The interactive row hiding of the output sheet wants a Worksheet_Change
'handler injected through the VBE, and injecting code wants VBE trust,
'which field machines lack. LLLog leaves the worksheet code module alone;
'the dropdowns stay plain validation cells and the filtering is reachable
'through CheckingOutput.FilterWorksheet.
'@version 1.2 (2026-08-31)


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

'The two columns beside it. CheckingOutput splits a label on the separator
'and lays the pieces across three columns: the outcome of an entry or the
'title of a section, then the timestamp, then the action and the detail.
'The text export reads the sheet back through these.
Private Const STAMP_COLUMN As Long = 4
Private Const ENTRY_COLUMN As Long = 5

'What the text export is called and where it reads the metadata from
Private Const EXPORT_SUFFIX As String = "-log.txt"
Private Const METADATA_SHEET_NAME As String = "Metadata"

'The three lines a reader of the text export finds its way by
Private Const HEADING_TITLE As String = "OutbreakTools linelist log"
Private Const METADATA_HEADING As String = "-- metadata --"
Private Const ENTRIES_HEADING As String = "-- log --"

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

'What sits between the procedure that raised the event and the action code.
'The three share the last written column, and this mark is what tells the
'name of a procedure apart from the action it took.
Private Const SOURCE_SEPARATOR As String = " > "

'What a step timing line opens with, and what stands between two steps on
'it. The opening words are what a reader searches the log for, and neither
'mark is one of the two separators above, so a timing line lands in the
'detail column whole.
Private Const STEPS_LEAD As String = "step times | "
Private Const STEPS_SEPARATOR As String = "; "
Private Const STEPS_TOTAL_LEAD As String = "total "

'Timer answers the seconds since midnight, so a walk running across
'midnight reads as a negative difference. A day is added back on.
Private Const SECONDS_IN_A_DAY As Double = 86400

Private Type TLLLog
    targetSheet As Worksheet
    writer As CheckingOutput
    store As HiddenNames
    sealed As Boolean
    platformTag As String
    stepNames As BetterArray
    stepSeconds As BetterArray
    openStepName As String
    openStepAt As Double
    walkStartedAt As Double
    walking As Boolean
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
         "export", "export-migration", "export-other", "export-analysis", _
         "export-geo", "export-log"
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
'@param source Optional String. The procedure that raised the event.
'@export
Public Sub LogSuccess(ByVal action As String, _
                      Optional ByVal detail As String = vbNullString, _
                      Optional ByVal source As String = vbNullString)
    WriteEntry action, detail, checkingSuccess, source
End Sub

'@label:log-warning
'@sub-title Log a refusal or a degraded run
'@param action String. The action code, used as the bundle title.
'@param detail Optional String. Short free text beside the timestamp.
'@param source Optional String. The procedure that raised the event.
'@export
Public Sub LogWarning(ByVal action As String, _
                      Optional ByVal detail As String = vbNullString, _
                      Optional ByVal source As String = vbNullString)
    WriteEntry action, detail, checkingWarning, source
End Sub

'@label:log-failure
'@sub-title Log an action that failed
'@param action String. The action code, used as the bundle title.
'@param detail Optional String. Short free text beside the timestamp.
'@param source Optional String. The procedure that raised the event.
'@export
Public Sub LogFailure(ByVal action As String, _
                      Optional ByVal detail As String = vbNullString, _
                      Optional ByVal source As String = vbNullString)
    WriteEntry action, detail, checkingError, source
End Sub

'@label:log-info
'@sub-title Log a lifecycle line
'@param action String. The action code, used as the bundle title.
'@param detail Optional String. Short free text beside the timestamp.
'@param source Optional String. The procedure that raised the event.
'@export
Public Sub LogInfo(ByVal action As String, _
                   Optional ByVal detail As String = vbNullString, _
                   Optional ByVal source As String = vbNullString)
    WriteEntry action, detail, checkingInfo, source
End Sub

'@section StepTiming
'===============================================================================
'@description The stopwatch of a long walk, and the one line it leaves behind.

'@label:start-walk
'@sub-title Open a fresh stopwatch on this log
'@details
'Called at the top of a long walk -- an export, an import -- before the
'first step is named. Anything an earlier walk left behind is dropped, so
'a walk that failed before it wrote its line never leaks its steps into
'the next one.
'The stopwatch is state on the instance, so one instance times one walk.
'LLExporter holds its own log and LinelistRun reads the one the event
'service holds, which is why the export walk and the import walk never
'meet in here.
'@export
Public Sub StartWalk()
    Set this.stepNames = New BetterArray
    Set this.stepSeconds = New BetterArray
    this.stepNames.LowerBound = 1
    this.stepSeconds.LowerBound = 1

    this.openStepName = vbNullString
    this.walkStartedAt = TimerNow()
    this.openStepAt = this.walkStartedAt
    this.walking = True
End Sub

'@label:mark-step
'@sub-title Close the step that was running and open the one named here
'@details
'Answers the name it was given, so a walk already carrying the step it is
'on in a variable keeps one line per step:
'
'  atStep = logStore.MarkStep("reading the dictionary")
'
'The name still reaches the failure message, and the stopwatch rides
'along. A walk with no stopwatch open is left alone and the name comes
'straight back, so a caller whose workbook takes no log writes the same
'code.
'@param stepName String. What the walk is about to do.
'@return String. The name it was given.
'@export
Public Function MarkStep(ByVal stepName As String) As String
    MarkStep = stepName
    If Not this.walking Then Exit Function

    CloseOpenStep
    this.openStepName = stepName
    this.openStepAt = TimerNow()
End Function

'@label:step-report
'@sub-title The whole walk on one line, each step with its seconds
'@details
'Reads the pairs and adds the total, and reads them without touching
'them: a step still running is reported with the seconds it has spent so
'far and stays open. So this answers the same walk whether it is called
'at the end or at a failure label.
'The seconds are written through Str, not Format, because Format follows
'the decimal separator of the machine and this box writes a comma. A
'timing line is read beside the timing line of another machine, so the
'two have to be written the same way.
'@return String. The line, empty when no walk is open.
'@export
Public Function StepReport() As String
    Dim index As Long
    Dim report As String
    Dim atNow As Double

    If Not this.walking Then Exit Function
    atNow = TimerNow()

    If Not this.stepNames Is Nothing Then
        If this.stepNames.Length > 0 Then
            For index = this.stepNames.LowerBound To this.stepNames.UpperBound
                report = report & CStr(this.stepNames.Item(index)) & " " & _
                         SecondsText(CDbl(this.stepSeconds.Item(index))) & _
                         STEPS_SEPARATOR
            Next index
        End If
    End If

    'The step still running. A failure label reaches this with the step
    'that refused still open, which is the one a reader wants most.
    If LenB(this.openStepName) > 0 Then
        report = report & this.openStepName & " " & _
                 SecondsText(ElapsedBetween(this.openStepAt, atNow)) & _
                 STEPS_SEPARATOR
    End If

    StepReport = report & STEPS_TOTAL_LEAD & _
                 SecondsText(ElapsedBetween(this.walkStartedAt, atNow))
End Function

'@label:log-steps
'@sub-title Write the walk''s step times as one info line, and close the walk
'@details
'One line per walk rather than one per step: a migration export names a
'dozen steps, and a line each would push the row cap of the sheet twelve
'times faster for the same reading.
'The action code is the walk''s own, so the timing line lands in the same
'section block as the success or failure line it belongs with.
'A walk with no stopwatch open writes nothing, so a caller can call this
'on every path.
'@param action String. The action code of the walk.
'@param source Optional String. The procedure that ran the walk.
'@export
Public Sub LogSteps(ByVal action As String, _
                    Optional ByVal source As String = vbNullString)
    Dim detail As String

    If Not this.walking Then Exit Sub

    detail = StepReport()
    EndWalk

    If LenB(detail) = 0 Then Exit Sub
    LogInfo action, STEPS_LEAD & detail, source
End Sub

'@section Export
'===============================================================================
'@description Writing the log out as a file a user can send on.

'@label:report-lines
'@sub-title The whole log as plain text, one line per element
'@details
'The text a user sends on when something goes wrong, so it stands on its
'own: the workbook it came from and the platform it is written on, then
'every row of the Metadata worksheet, then the log itself in the order it
'was written.
'The log SHEET is the record here, unlike GenerationLog which keeps its run
'in memory, so the rows are read back off the sheet.
'The text is built here rather than inside ExportText so what a file will
'say can be read without a file being written. A log that cannot reach its
'metadata still answers its entries.
'@return BetterArray. The lines, lower bound 1.
'@throws ProjectError.ObjectNotInitialized When the log has no worksheet.
'@export
Public Function ReportLines() As BetterArray
    Dim lines As BetterArray
    Dim sh As Worksheet

    Set sh = Wksh()
    If sh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
         "This log has no worksheet to write out"
    End If

    Set lines = New BetterArray
    lines.LowerBound = 1

    AppendHeadingTo lines, sh
    AppendMetadataTo lines, sh.Parent
    AppendEntriesTo lines, sh

    Set ReportLines = lines
End Function

'@label:export-text
'@sub-title Write the whole log out as a plain text file
'@details
'The file lands in the given folder as <baseName>-log.txt. A second export
'of the same linelist overwrites the first, which is what a person sending
'"the log" wants.
'What goes into it is ReportLines. This routine only opens the file, writes
'the lines and closes it.
'@param folderPath String. The folder taking the file.
'@param baseName String. The file name without its suffix. The workbook name when empty.
'@return String. The full path of the written file.
'@throws ProjectError.ObjectNotInitialized When the log has no worksheet.
'@throws ProjectError.InvalidArgument When no folder is given.
'@export
Public Function ExportText(ByVal folderPath As String, _
                           Optional ByVal baseName As String = vbNullString) As String
    Dim filePath As String
    Dim fileNumber As Long
    Dim lines As BetterArray
    Dim index As Long

    If LenB(Trim$(folderPath)) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
         "A folder is required to write the log out"
    End If

    'Built before the file is opened, so a fault while reading the sheets
    'leaves no half-written file behind.
    Set lines = ReportLines()

    If LenB(Trim$(baseName)) = 0 Then baseName = Wksh().Parent.Name

    filePath = folderPath
    If Right$(filePath, 1) <> Application.PathSeparator Then
        filePath = filePath & Application.PathSeparator
    End If
    filePath = filePath & baseName & EXPORT_SUFFIX

    fileNumber = FreeFile

    'The handler closes the file on a mid-write fault; a locked or missing
    'folder raises to the caller with the file released.
    On Error GoTo CloseFile
    Open filePath For Output As #fileNumber
    For index = lines.LowerBound To lines.UpperBound
        Print #fileNumber, CStr(lines.Item(index))
    Next index
    Close #fileNumber
    On Error GoTo 0

    ExportText = filePath
    Exit Function

CloseFile:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    Close #fileNumber
    On Error GoTo 0

    ThrowError errNumber, errDesc
End Function

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

'@label:write-heading-to
'@sub-title The first block of the report: which workbook, and on what
'@details
'The platform of the export and the platform of an entry can differ, since
'a linelist travels between machines, and both are worth having.
'@param lines BetterArray. The lines built so far.
'@param sh Worksheet. The log worksheet.
Private Sub AppendHeadingTo(ByVal lines As BetterArray, ByVal sh As Worksheet)
    Dim savedIn As String

    savedIn = sh.Parent.Path
    If LenB(savedIn) = 0 Then savedIn = "(never saved)"

    lines.Push HEADING_TITLE
    lines.Push "written  " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & _
               " on " & PlatformTag()
    lines.Push "workbook " & sh.Parent.Name
    lines.Push "saved in " & savedIn
    lines.Push vbNullString
End Sub

'@label:append-metadata-to
'@sub-title Every row of the Metadata worksheet, into the report
'@details
'The whole variable and value block goes out, not a chosen few rows, so a
'linelist built by a later designer carries whatever that designer wrote
'without this routine being touched again. The block below the creation rows
'is the geobase metadata, and it goes out with the rest.
'A linelist with no Metadata worksheet says so and the report goes on: a log
'is worth having even when the metadata is missing.
'@param lines BetterArray. The lines built so far.
'@param wkb Workbook. The linelist workbook.
Private Sub AppendMetadataTo(ByVal lines As BetterArray, ByVal wkb As Workbook)
    Dim metash As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim keyText As String
    Dim valueText As String

    lines.Push METADATA_HEADING

    'The worksheets collection raises 9 on a missing name, and a linelist
    'without the sheet is an ordinary state rather than a fault.
    Set metash = Nothing
    On Error Resume Next
    Set metash = wkb.Worksheets(METADATA_SHEET_NAME)
    On Error GoTo 0

    If metash Is Nothing Then
        lines.Push "this linelist carries no " & METADATA_SHEET_NAME & " worksheet"
        lines.Push vbNullString
        Exit Sub
    End If

    lastRow = metash.Cells(metash.Rows.Count, 1).End(xlUp).Row

    For rowIndex = 1 To lastRow
        keyText = Trim$(CellText(metash.Cells(rowIndex, 1).Value))
        valueText = Trim$(CellText(metash.Cells(rowIndex, 2).Value))
        If LenB(keyText) > 0 Then lines.Push keyText & " = " & valueText
    Next rowIndex

    lines.Push vbNullString
End Sub

'@label:append-entries-to
'@sub-title Every row of the log sheet, into the report
'@details
'A row carrying the title alone opens a section and reads as a heading in
'brackets; every other written row is one entry, and it reads as the
'timestamp, the outcome in brackets, then the action and the detail. Blank
'rows are dropped, because CheckingOutput lays them out for the eye and they
'say nothing.
'@param lines BetterArray. The lines built so far.
'@param sh Worksheet. The log worksheet.
Private Sub AppendEntriesTo(ByVal lines As BetterArray, ByVal sh As Worksheet)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim titleText As String
    Dim stampText As String
    Dim entryText As String

    lines.Push ENTRIES_HEADING

    lastRow = sh.Cells(sh.Rows.Count, OUTPUT_COLUMN).End(xlUp).Row

    For rowIndex = 1 To lastRow
        titleText = PlainOutcome(Trim$(CellText(sh.Cells(rowIndex, OUTPUT_COLUMN).Value)))
        stampText = Trim$(CellText(sh.Cells(rowIndex, STAMP_COLUMN).Value))
        entryText = Trim$(CellText(sh.Cells(rowIndex, ENTRY_COLUMN).Value))

        If LenB(stampText) = 0 And LenB(entryText) = 0 Then
            If LenB(titleText) > 0 Then
                lines.Push vbNullString
                lines.Push "[" & titleText & "]"
            End If
        Else
            lines.Push stampText & "  [" & titleText & "]  " & entryText
        End If
    Next rowIndex
End Sub

'@label:plain-outcome
'@sub-title The outcome word of a row, without the picture in front of it
'@details
'Checking paints an outcome with a symbol before the word -- a cross for an
'error, a tick for a success, an i in a circle for an info line -- and that
'symbol is what the sheet shows.
'The text export cannot carry it. Print # writes one byte per character, so
'every one of those symbols reaches the file as a single stray mark, and a
'log a user sends on reads "[_ Error]" and "[_ Info]" throughout. The word
'alone is what the file wants.
'A section title carries no symbol and comes back untouched: only a first
'character outside plain text is taken off.
'@param text String. The cell value, already trimmed.
'@return String. The word alone.
Private Function PlainOutcome(ByVal text As String) As String
    PlainOutcome = text
    If LenB(text) = 0 Then Exit Function
    If AscW(Left$(text, 1)) < 128 Then Exit Function

    PlainOutcome = Trim$(Mid$(text, 2))
End Function

'@label:cell-text
'@sub-title One cell as text, with an error value reading as empty
'@details
'A cell holding #N/A or #REF! raises 13 on CStr, and one such cell in a log
'or on a metadata sheet must not take down the export of everything else.
'@param cellValue Variant. The value read from a cell.
'@return String. The value as text, or an empty string when it is an error.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@label:write-entry
'@sub-title Rotate when the cap is passed, then flush one dated entry
'@param action String. The action code, used as the bundle title.
'@param detail String. Short free text beside the timestamp.
'@param scope Byte. CheckingScope value carrying the outcome.
'@param source String. The procedure that raised the event. May be empty.
'@throws ProjectError.InvalidArgument When the action code is empty.
Private Sub WriteEntry(ByVal action As String, ByVal detail As String, _
                       ByVal scope As Byte, _
                       Optional ByVal source As String = vbNullString)
    If LenB(Trim$(action)) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
         "An action code is required to log an event"
    End If

    RotateWhenPastCap
    FlushEntry action, detail, scope, source
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
'@param source String. The procedure that raised the event. May be empty.
Private Sub FlushEntry(ByVal action As String, ByVal detail As String, _
                       ByVal scope As Byte, _
                       Optional ByVal source As String = vbNullString)
    Dim checks As Checking
    Dim batch As BetterArray

    Set checks = Checking.Create(SectionOf(action))
    checks.Add "entry", EntryLabel(action, detail, source), scope

    Set batch = New BetterArray
    batch.LowerBound = 1
    batch.Push checks

    Writer().PrintOutput batch
End Sub

'@label:entry-label
'@sub-title Compose the dated label of one entry
'@details
'The label opens with the timestamp and the platform; the procedure, the
'action code and the detail follow behind the separator, so the two halves
'land in their own columns of the log sheet. The action code moved here when
'the title became the section: the sheet has three titles now, and the line
'has to say which action it records. A separator inside the free text would
'push it past the last written column, so it is softened to a single dash.
'The procedure sits in front of the action because that is the order a
'reader wants: who acted, what they did, and on what.
'The platform rides with the timestamp rather than with the detail, because
'the detail column is what a reader searches and what several tests read
'whole. It is on every line on purpose: a log rotation clears the sheet, so
'a platform written once at the top of a session would not survive one.
'@param action String. The action code of the event.
'@param detail String. Short free text beside the timestamp.
'@param source String. The procedure that raised the event. May be empty.
'@return String. The composed label.
Private Function EntryLabel(ByVal action As String, ByVal detail As String, _
                            Optional ByVal source As String = vbNullString) As String
    Dim label As String

    label = Format$(Now, "yyyy-mm-dd hh:mm:ss") & " " & PlatformTag() & _
            DETAIL_SEPARATOR

    If LenB(Trim$(source)) > 0 Then
        label = label & Soften(source) & SOURCE_SEPARATOR
    End If

    label = label & Trim$(action)

    If LenB(Trim$(detail)) > 0 Then
        label = label & ACTION_SEPARATOR & Soften(detail)
    End If

    EntryLabel = label
End Function

'@label:soften
'@sub-title Take the column separator out of a piece of free text
'@details
'CheckingOutput splits the label on the separator and lays the pieces across
'three columns, so a separator inside free text would push the rest of the
'line past the last written column and out of the sheet.
'@param text String. The free text.
'@return String. The text with every separator turned into a single dash.
Private Function Soften(ByVal text As String) As String
    Soften = Replace(Trim$(text), DETAIL_SEPARATOR, "-")
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

'@label:close-open-step
'@sub-title Keep the step that was running, with the seconds it took
Private Sub CloseOpenStep()
    Dim atNow As Double

    If LenB(this.openStepName) = 0 Then Exit Sub

    atNow = TimerNow()
    this.stepNames.Push this.openStepName
    this.stepSeconds.Push ElapsedBetween(this.openStepAt, atNow)

    this.openStepName = vbNullString
    this.openStepAt = atNow
End Sub

'@label:end-walk
'@sub-title Close the stopwatch and drop what it held
Private Sub EndWalk()
    this.walking = False
    this.openStepName = vbNullString
    Set this.stepNames = Nothing
    Set this.stepSeconds = Nothing
End Sub

'@label:timer-now
'@sub-title The stopwatch reading, as a Double
'@details
'Timer answers a Single, which carries about seven digits: past midday
'the seconds since midnight already use six of them, so a Single reads
'hundredths at best. The widening happens here, once, so every
'subtraction below is done in Double.
'@return Double. Seconds since midnight.
Private Function TimerNow() As Double
    TimerNow = CDbl(Timer)
End Function

'@label:elapsed-between
'@sub-title The seconds between two stopwatch readings
'@details
'A walk that runs across midnight reads as a negative difference, because
'Timer restarts at zero. A day is added back on.
'@param fromAt Double. The earlier reading.
'@param toAt Double. The later reading.
'@return Double. The seconds between them.
Private Function ElapsedBetween(ByVal fromAt As Double, _
                                ByVal toAt As Double) As Double
    ElapsedBetween = toAt - fromAt
    If ElapsedBetween < 0 Then _
        ElapsedBetween = ElapsedBetween + SECONDS_IN_A_DAY
End Function

'@label:seconds-text
'@sub-title A number of seconds as text, the same way on every machine
'@details
'Str is used rather than Format because Format follows the decimal
'separator of the machine: this box is en_FR and writes 2,15 where a
'Windows box in English writes 2.15. Two timing lines that are read side
'by side have to be written the same way.
'Str drops the leading zero of a value under one, so it is put back.
'@param seconds Double. The seconds to write.
'@return String. Something of the shape "2.15s".
Private Function SecondsText(ByVal seconds As Double) As String
    Dim text As String

    text = Trim$(Str$(Round(seconds, 2)))

    If Left$(text, 1) = "." Then
        text = "0" & text
    ElseIf Left$(text, 2) = "-." Then
        text = "-0" & Mid$(text, 2)
    End If

    SecondsText = text & "s"
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
