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
'bundle: the action code is the bundle title, the entry label carries the
'timestamp and a short detail, and the scope carries the outcome. Every
'bundle is flushed through CheckingOutput.PrintOutput, which appends
'across renders, so the sheet grows one block per event and the built-in
'status and title dropdowns of the output sheet filter the log by outcome
'and by action.
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

Private Type TLLLog
    targetSheet As Worksheet
    writer As CheckingOutput
    store As HiddenNames
    sealed As Boolean
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
'The bundle carries one key, so the key uniqueness of Checking binds
'inside the bundle and every event stays a plain append.
'@param action String. The action code, used as the bundle title.
'@param detail String. Short free text beside the timestamp.
'@param scope Byte. CheckingScope value carrying the outcome.
Private Sub FlushEntry(ByVal action As String, ByVal detail As String, ByVal scope As Byte)
    Dim checks As Checking
    Dim batch As BetterArray

    Set checks = Checking.Create(action)
    checks.Add "entry", EntryLabel(detail), scope

    Set batch = New BetterArray
    batch.LowerBound = 1
    batch.Push checks

    Writer().PrintOutput batch
End Sub

'@label:entry-label
'@sub-title Compose the dated label of one entry
'@details
'The label opens with the timestamp; the detail follows behind the
'separator, so the two land in their own columns of the log sheet. A
'separator inside the detail would push text past the last written
'column, so it is softened to a single dash first.
'@param detail String. Short free text beside the timestamp.
'@return String. The composed label.
Private Function EntryLabel(ByVal detail As String) As String
    Dim label As String

    label = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    If LenB(Trim$(detail)) > 0 Then
        label = label & DETAIL_SEPARATOR & _
                Replace(Trim$(detail), DETAIL_SEPARATOR, "-")
    End If

    EntryLabel = label
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
'@return CheckingOutput. Writer bound to the log worksheet.
Private Function Writer() As CheckingOutput
    If this.writer Is Nothing Then
        Set this.writer = CheckingOutput.Create(Wksh(), OUTPUT_TITLE)
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
