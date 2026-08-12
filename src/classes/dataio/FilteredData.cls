VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilteredData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Syncs the filtered companion sheets ahead of a filtered export"

'@PredeclaredId
'@Folder("DataIO")
'@ModuleDescription("Syncs the filtered companion sheets ahead of a filtered export")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ProcedureNotUsed

Option Explicit

'@class FilteredData
'@description
'Copies every HList table of the source workbook into its filtered companion
'sheet, so a filtered export reads what the user can see on the screen. The
'companion keeps the rows that are visible and filled: a row leaves the copy
'when its source row is hidden or when the copied row is empty.
'
'Sync walks every worksheet, rewrites each companion, and answers the number
'of sheets it brought over. A sheet that fails is skipped and recorded, and
'the walk continues, so one broken sheet leaves the others synced.
'FailedSheets names each skipped sheet with the reason. A run that skipped a
'sheet also writes one failure line to the user log of the workbook. The
'caller reads FailedSheets and decides what a failure aborts; ExportButton
'aborts the export.
'
'THE DELETE WORKS RUN BY RUN
'-------------------------------------------------------------------------------
'Hidden and empty rows sit in contiguous runs, and each run leaves the
'companion with one delete. A per-row delete range grows one area per row,
'and Excel slows sharply past a few hundred areas, so on a large table the
'run shape is what keeps the sync usable.
'@depends HiddenNames, LLLog, BetterArray
'@version 1.0 (2026-08-12)


'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "FilteredData"

'The worksheet-level hidden names the sync reads. One store per sheet
'answers both.
Private Const SHEET_TYPE_KEY As String = "sheet_type"
Private Const FILTERED_SHEET_KEY As String = "filtered_sheet"

'The sheet_type value of a data entry sheet with a filtered companion
Private Const HLIST_TAG As String = "HList"

'The action code of the failure line in the user log. It is the code the
'linelist event surface logs a filter-refresh failure under, so the two
'paths read as one action in the log.
Private Const LOG_ACTION As String = "MSG_ErrUpdate"


'@section State
'===============================================================================

Private Type TFilteredData
    sourceWkb As Workbook
    failedSheets As BetterArray
    userLog As LLLog
    triedLog As Boolean
    sealed As Boolean
End Type

Private this As TFilteredData


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a filtered data sync over a workbook
'@details
'Factory method that validates the workbook and hands back a ready instance.
'Called on the predeclared instance.
'@param sourceWkb Workbook. The linelist workbook holding the HList sheets.
'@return FilteredData. A ready instance.
'@throws ProjectError.ObjectNotInitialized When sourceWkb is Nothing.
'@export
Public Function Create(ByVal sourceWkb As Workbook) As FilteredData

    If sourceWkb Is Nothing Then
        ThrowError ObjectNotInitialized, "Workbook required"
    End If

    Dim newSync As FilteredData
    Set newSync = New FilteredData
    Set newSync.SourceWorkbook = sourceWkb
    newSync.Seal
    Set Create = newSync
End Function


'@section Public Accessors
'===============================================================================

'@label:source-workbook-set
'@prop-title Assign the source linelist workbook
'@param wkb Workbook. The workbook to store.
Public Property Set SourceWorkbook(ByVal wkb As Workbook)
    GuardNotSealed "SourceWorkbook"
    Set this.sourceWkb = wkb
End Property

'@label:seal
'@sub-title Prevent further changes to setup-only properties
Public Sub Seal()
    this.sealed = True
End Sub

'@label:failed-sheets
'@prop-title The sheets the last sync skipped, each with the reason
'@details
'One entry per skipped sheet, shaped "sheet name: reason". The list is
'rebuilt at the top of every Sync, so it always describes the last run.
'Before the first run it answers an empty list.
'@return BetterArray. The skipped sheets of the last sync.
'@export
Public Property Get FailedSheets() As BetterArray
    If this.failedSheets Is Nothing Then Set this.failedSheets = New BetterArray
    Set FailedSheets = this.failedSheets
End Property


'@section Sync
'===============================================================================

'@label:sync
'@sub-title Rewrite every filtered companion from its live table
'@details
'Walks the worksheets of the source workbook. Each sheet whose sheet_type is
'HList has its table copied into the filtered companion, and the hidden and
'empty rows are then deleted from the copy. A sheet whose table holds no data
'rows is left as it is. A sheet that raises is recorded in FailedSheets with
'the reason and the walk continues.
'@return Long. The number of sheets whose companion was rewritten.
'@export
Public Function Sync() As Long

    Dim sh As Worksheet
    Dim store As HiddenNames
    Dim syncedCount As Long

    Set this.failedSheets = New BetterArray

    'The handler records the sheet and resumes the walk, so one broken
    'sheet leaves the others synced. The caller reads FailedSheets.
    On Error GoTo SheetFail

    For Each sh In this.sourceWkb.Worksheets
        Set store = HiddenNames.Create(sh)
        If store.ValueAsString(SHEET_TYPE_KEY) = HLIST_TAG Then
            If SyncSheet(sh, store) Then syncedCount = syncedCount + 1
        End If
NextSheet:
    Next sh

    On Error GoTo 0
    LogFailures
    Sync = syncedCount
    Exit Function

SheetFail:
    RecordFailure sh.Name, Err.Description
    Resume NextSheet
End Function

'@label:sync-sheet
'@sub-title Copy one HList table into its filtered companion
'@details
'Reads the companion name off the store, clears the companion table, sizes
'it to the live table, copies the values across, and deletes the hidden and
'empty rows from the copy. Raises when the filtered_sheet name is missing,
'and lets any worksheet failure travel up to the per-sheet handler in Sync.
'@param sh Worksheet. The HList sheet to copy.
'@param store HiddenNames. The store of that sheet, already built by Sync.
'@return Boolean. True when the companion was rewritten.
'@throws ProjectError.ElementNotFound When the filtered_sheet name is missing.
Private Function SyncSheet(ByVal sh As Worksheet, _
                           ByVal store As HiddenNames) As Boolean

    Dim lo As ListObject
    Dim loRng As Range
    Dim filtsh As Worksheet
    Dim filtLo As ListObject
    Dim destRng As Range
    Dim filteredName As String

    Set lo = sh.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function

    filteredName = store.ValueAsString(FILTERED_SHEET_KEY)
    If LenB(filteredName) = 0 Then
        ThrowError ElementNotFound, "the filtered_sheet name is missing"
    End If

    Set loRng = lo.DataBodyRange
    Set filtsh = this.sourceWkb.Worksheets(filteredName)
    Set filtLo = filtsh.ListObjects(1)

    If Not filtLo.DataBodyRange Is Nothing Then filtLo.DataBodyRange.Delete

    filtLo.Resize filtsh.Range(lo.Range.Address)
    Set destRng = filtsh.Range(loRng.Address)

    'Copy values to the companion, then take the hidden and empty rows off
    destRng.Value = loRng.Value
    DeleteFilteredRows loRng, filtLo.HeaderRowRange

    SyncSheet = True
End Function

'@label:delete-filtered-rows
'@sub-title Delete the hidden and empty rows from the companion copy
'@details
'Walks the rows bottom-up, so a delete leaves the rows above it in place. A
'row is marked when its source row is hidden or when its copy is empty.
'Marked rows are gathered into contiguous runs and each run is deleted
'whole, one delete per run, which keeps the area count at the run count.
'The hidden state is read on the source sheet and the emptiness on the copy.
'@param sourceRng Range. The data body of the live table.
'@param headerRng Range. The header row of the companion table.
Private Sub DeleteFilteredRows(ByVal sourceRng As Range, _
                               ByVal headerRng As Range)

    Dim rowCounter As Long
    Dim runLength As Long
    Dim wsFunc As WorksheetFunction

    Set wsFunc = Application.WorksheetFunction

    runLength = 0
    For rowCounter = sourceRng.Rows.Count To 1 Step -1
        If sourceRng.Cells(rowCounter, 1).EntireRow.Hidden Or _
           wsFunc.CountA(headerRng.Offset(rowCounter)) = 0 Then
            runLength = runLength + 1
        Else
            If runLength > 0 Then
                headerRng.Offset(rowCounter + 1).Resize(runLength).Delete
                runLength = 0
            End If
        End If
    Next rowCounter

    'A run that reaches the first data row closes here
    If runLength > 0 Then headerRng.Offset(1).Resize(runLength).Delete
End Sub


'@section Private Helpers
'===============================================================================

'@label:record-failure
'@sub-title Record one skipped sheet with the reason
'@param sheetName String. The sheet the sync skipped.
'@param reason String. What the failure said.
Private Sub RecordFailure(ByVal sheetName As String, ByVal reason As String)
    FailedSheets.Push sheetName & ": " & reason
End Sub

'@label:log-failures
'@sub-title Write the skipped sheets of a run to the user log
'@details
'One failure line per run, with the skipped sheets joined in the detail. A
'run that skipped nothing writes nothing. A log write must never take down
'the sync it records, so a write that raises is swallowed here.
Private Sub LogFailures()

    Dim heldLog As LLLog
    Dim detail As String
    Dim counter As Long

    If this.failedSheets.Length = 0 Then Exit Sub

    Set heldLog = UserLog()
    If heldLog Is Nothing Then Exit Sub

    For counter = this.failedSheets.LowerBound To this.failedSheets.UpperBound
        If LenB(detail) > 0 Then detail = detail & "; "
        detail = detail & this.failedSheets.Item(counter)
    Next counter

    'The log is a bystander of the sync (see the details above)
    On Error Resume Next
    heldLog.LogFailure LOG_ACTION, detail
    On Error GoTo 0
End Sub

'@label:user-log
'@sub-title The user log of the source workbook, built once and held
'@details
'A workbook whose log cannot be built answers Nothing and the sync keeps
'its result, so the log stays a bystander. The tried flag keeps a failed
'build from being paid again on every run.
'@return LLLog. The held log, or Nothing when the build failed.
Private Function UserLog() As LLLog
    If Not this.triedLog Then
        this.triedLog = True
        'The build failure is the state the Nothing answer carries
        On Error Resume Next
        Set this.userLog = LLLog.Create(this.sourceWkb)
        On Error GoTo 0
    End If
    Set UserLog = this.userLog
End Function

'@label:guard-not-sealed
'@sub-title Raise when a setup-only property is written after sealing
'@param propName String. The name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber Long. A ProjectError enum value (e.g. ObjectNotInitialized).
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
