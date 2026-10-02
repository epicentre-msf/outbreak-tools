VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MasterSetupLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Dated record of user actions on the __log worksheet of a master setup"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Dated record of user actions on the __log worksheet of a master setup")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@depends BetterArray, Checking, CheckingOutput, HiddenNames

Option Explicit

'@class MasterSetupLog
'@description
'Keeps the dated record of the important user actions of a master setup
'on the very hidden __log worksheet. It is the master setup copy of LLLog.
'Each event is one Checking bundle: the bundle title is the SECTION the
'action belongs to, the entry label carries the timestamp, the action code
'and a short detail, and the scope carries the outcome. Every bundle is
'flushed through CheckingOutput.PrintOutput in compact mode, so the sheet
'grows one line per event and the status and title dropdowns of the output
'sheet filter the log by outcome and by section.
'
'WHAT ONE LINE SAYS
'-------------------------------------------------------------------------------
'An entry line reads as the timestamp and the platform, then the procedure
'that raised the event, then the action code and the detail:
'
'  2026-08-26 13:55:45 win-64 excel-16.0  clickAddSheet > add-disease: Measles
'
'The procedure is named by the caller. VBA carries no call stack to read a
'name from, so a line that names only the action leaves the reader guessing
'which button wrote it. The name is optional: a caller that passes none
'writes a line opening at the action code.
'
'THE THREE SECTIONS
'-------------------------------------------------------------------------------
'The log is read by three sections, and the title of an entry is the
'section the action belongs to. The writer runs in compact mode, so a
'section title is written the first time the sheet carries it and never
'again.
'
'  open/close             the workbook open, the workbook close, a log restart
'  data input/output      every export, every import, the disease comparison
'  master setup lifecycle everything else the user does in the file
'
'An action code no case names lands in master setup lifecycle, so a new
'logged action reaches the sheet without touching this class. A code that
'opens with export- or import- lands in data input/output for the same
'reason.
'
'PROVISIONING
'-------------------------------------------------------------------------------
'Create(wkb) binds to the __log worksheet of the given workbook and builds
'it very hidden when the workbook has none, so an older master setup grows
'its log on its first logged event. Building a sheet wants the workbook
'structure unprotected, which is the duty of the caller.
'
'ROTATION
'-------------------------------------------------------------------------------
'LOG_MAX_ENTRIES is read as a row cap. A flush that finds the last written
'row past the cap clears the sheet and removes the row marker
'CheckingOutput keeps, so the next render re-initialises the sheet, and the
'fresh log opens with one info line ("restarted").
'
'TEXT EXPORT
'-------------------------------------------------------------------------------
'A master setup has no Metadata sheet, so the header of the text export
'carries the file name, the date, the platform and the count of disease
'worksheets, then the entries in the order they were written.
'
'FILTER DROPDOWNS
'-------------------------------------------------------------------------------
'The interactive row hiding of the output sheet wants a Worksheet_Change
'handler injected through the VBE, and injecting code wants VBE trust,
'which field machines lack. The worksheet code module is left alone; the
'dropdowns stay plain validation cells and the filtering is reachable
'through CheckingOutput.FilterWorksheet.
'@version 1.0 (2026-08-26)


'@section State
'===============================================================================
'@description Backing UDT and class constants.

Private Const CLASSNAME As String = "MasterSetupLog"
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

'What the text export is called
Private Const EXPORT_SUFFIX As String = "-log.txt"

'The two lines a reader of the text export finds its way by
Private Const HEADING_TITLE As String = "OutbreakTools master setup log"
Private Const ENTRIES_HEADING As String = "-- log --"

'The hidden name a disease worksheet carries, and its value. The text
'export counts the disease worksheets through it. The two mirror the
'constants of MasterSetupHelpers, since a class calls no module.
Private Const SHEET_TAG_NAME As String = "sheetTag"
Private Const DISEASE_TAG_VALUE As String = "disease"

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
Private Const SECTION_LIFECYCLE As String = "master setup lifecycle"

'The two prefixes that send an action code to data input/output on their
'own, so a new export or import reaches the right block untouched.
Private Const EXPORT_PREFIX As String = "export-"
Private Const IMPORT_PREFIX As String = "import-"

'What sits between the action code and the detail on an entry line. The
'two share the last written column, so they need a mark of their own.
Private Const ACTION_SEPARATOR As String = ": "

'What sits between the procedure that raised the event and the action code.
Private Const SOURCE_SEPARATOR As String = " > "

Private Type TMasterSetupLog
    targetSheet As Worksheet
    writer As CheckingOutput
    store As HiddenNames
    sealed As Boolean
    platformTag As String
End Type

Private this As TMasterSetupLog

'@section Factory
'===============================================================================
'@description Entry point for creating MasterSetupLog instances.

'@label:create
'@sub-title Create a log store bound to the __log worksheet of a workbook
'@details
'Binds to the __log worksheet of the given workbook. When the workbook
'has none, the sheet is appended after the last worksheet and made very
'hidden, so an older master setup grows its log on first use.
'@param wkb Workbook. The master setup workbook carrying the log.
'@return MasterSetupLog. Fully initialised instance.
'@throws ProjectError.ObjectNotInitialized When wkb is Nothing.
'@export
Public Function Create(ByVal wkb As Workbook) As MasterSetupLog
    If wkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
         "Unable to create an instance of the class, please provide a workbook"
    End If

    Dim newLog As MasterSetupLog
    Set newLog = New MasterSetupLog

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
'Answers the internal sheet name so the Open Log door and the deploy list
'reach the sheet through the class.
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
'boundaries and read as open/close; every export, every import and the
'disease comparison move data in or out of the file and read as data
'input/output; every other action the user takes in the file reads as
'master setup lifecycle, which is also where an action code named by no
'case below lands.
'@param action String. The action code of the event.
'@return String. One of the three section titles.
'@export
Public Function SectionOf(ByVal action As String) As String
    Dim code As String

    code = LCase$(Trim$(action))

    Select Case code

    Case "open", "close", "workbook-open", "workbook-close", RESTART_ACTION
        SectionOf = SECTION_OPENCLOSE

    Case "export-setup", "export-migration", "import-migration", _
         "import-setup", "import-passwords", "compare-diseases", _
         "export-log", "export", "import"
        SectionOf = SECTION_DATAIO

    Case Else
        If Left$(code, Len(EXPORT_PREFIX)) = EXPORT_PREFIX Or _
           Left$(code, Len(IMPORT_PREFIX)) = IMPORT_PREFIX Then
            SectionOf = SECTION_DATAIO
        Else
            SectionOf = SECTION_LIFECYCLE
        End If

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

'@label:disease-sheet-count
'@prop-title How many disease worksheets the workbook carries
'@details
'A disease worksheet is one whose hidden sheetTag name reads disease, the
'same reading MasterSetupHelpers.IsMasterDiseaseSheet makes. The count
'heads the text export.
'@return Long. The count of disease worksheets.
'@export
Public Property Get DiseaseSheetCount() As Long
    Dim sh As Worksheet
    Dim total As Long

    For Each sh In Wksh().Parent.Worksheets
        If IsDiseaseSheet(sh) Then total = total + 1
    Next sh

    DiseaseSheetCount = total
End Property

'@section Logging
'===============================================================================
'@description One method per outcome, each flushing one Checking bundle.

'@label:log-success
'@sub-title Log an action that ended well
'@param action String. The action code, written on the entry line.
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
'@param action String. The action code, written on the entry line.
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
'@param action String. The action code, written on the entry line.
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
'@param action String. The action code, written on the entry line.
'@param detail Optional String. Short free text beside the timestamp.
'@param source Optional String. The procedure that raised the event.
'@export
Public Sub LogInfo(ByVal action As String, _
                   Optional ByVal detail As String = vbNullString, _
                   Optional ByVal source As String = vbNullString)
    WriteEntry action, detail, checkingInfo, source
End Sub

'@section Export
'===============================================================================
'@description Writing the log out as a file a user can send on.

'@label:report-lines
'@sub-title The whole log as plain text, one line per element
'@details
'The text a user sends on when something goes wrong, so it stands on its
'own: the workbook it came from, when and on what platform it was written,
'how many disease worksheets the file carries, then the log itself in the
'order it was written. The text is built here so what a file will say can
'be read without a file being written.
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
    AppendEntriesTo lines, sh

    Set ReportLines = lines
End Function

'@label:export-text
'@sub-title Write the whole log out as a plain text file
'@details
'The file lands in the given folder as <baseName>-log.txt. A second export
'of the same master setup overwrites the first, which is what a person
'sending "the log" wants. What goes into it is ReportLines. This routine
'only opens the file, writes the lines and closes it.
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

    'Built before the file is opened, so a fault while reading the sheet
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
'visibility, so the Open Log door can hold it open across a bind.
'@param wkb Workbook. The master setup workbook.
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

'@label:is-disease-sheet
'@sub-title Whether a worksheet carries the disease tag
'@details
'HiddenNames.Create raises on a sheet it cannot bind, and a sheet with no
'tag answers an empty string; both read as a plain worksheet.
'@param sh Worksheet. The worksheet to read.
'@return Boolean. True when the sheetTag name reads disease.
Private Function IsDiseaseSheet(ByVal sh As Worksheet) As Boolean
    Dim tagStore As HiddenNames

    If sh Is Nothing Then Exit Function

    On Error Resume Next
        Set tagStore = HiddenNames.Create(sh)
    On Error GoTo 0
    If tagStore Is Nothing Then Exit Function

    IsDiseaseSheet = (StrComp(Trim$(tagStore.ValueAsString(SHEET_TAG_NAME)), _
                              DISEASE_TAG_VALUE, vbTextCompare) = 0)
End Function

'@label:append-heading-to
'@sub-title The first block of the report: which workbook, when, on what
'@details
'A master setup has no Metadata worksheet, so the header says what a
'reader of the file wants to know first: the file, the date and the
'platform of the export, where the file is saved and how many disease
'worksheets it carries.
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
    lines.Push "disease worksheets " & CStr(DiseaseSheetCount())
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
'Checking paints an outcome with a symbol before the word, and that symbol
'is what the sheet shows. Print # writes one byte per character, so every
'one of those symbols reaches the file as a single stray mark. The word
'alone is what the file wants. A section title carries no symbol and comes
'back untouched: only a first character outside plain text is taken off.
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
'A cell holding #N/A or #REF! raises 13 on CStr, and one such cell must
'keep the export of everything else alive.
'@param cellValue Variant. The value read from a cell.
'@return String. The value as text, or an empty string when it is an error.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@label:write-entry
'@sub-title Rotate when the cap is passed, then flush one dated entry
'@param action String. The action code, written on the entry line.
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
'The bundle title is the section of the action, so the sheet carries three
'titles. The bundle carries one key, so the key uniqueness of Checking
'binds inside the bundle and every event stays a plain append.
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
'land in their own columns of the log sheet. A separator inside the free
'text would push it past the last written column, so it is softened to a
'single dash. The platform is on every line on purpose: a log rotation
'clears the sheet, so a platform written once at the top of a session
'would be lost with it.
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
'A log read months later has to say where it was written, since the same
'code behaves differently on the two platforms. The name and the bitness
'come from the compile constants; the Excel version comes from the
'application and is read once, because it is the same for every line of a
'session.
'@return String. Something of the shape "win-64 excel-16.0".
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
'again, and no blank rows stand between the entries of one block.
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
