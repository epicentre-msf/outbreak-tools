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
'sheet, so a filtered export and the analysis formulas read what the user can
'see on the screen. The companion keeps the rows that are visible and filled:
'a row stays out of the copy when its source row is hidden or when the row
'holds no value at all.
'
'Sync walks every worksheet, rewrites each companion, and answers the number
'of sheets it brought over. A sheet that fails is skipped and recorded, and
'the walk continues, so one broken sheet leaves the others synced.
'FailedSheets names each skipped sheet with the reason. A run that skipped a
'sheet also writes one failure line to the user log of the workbook. The
'caller reads FailedSheets and decides what a failure costs: ExportButton
'aborts the export, and the linelist filter refresh shows one box.
'
'THE WHOLE ROW TEST RUNS IN MEMORY
'-------------------------------------------------------------------------------
'Each companion is written once. The source block is read in one crossing,
'visibility comes off one SpecialCells call over the source rows, the blank
'test runs on the array, and the kept rows go down in one write. A shape that
'copies the whole block over and then deletes the rejects pays two crossings
'per row, on tables that ClickAddRows pads 199 blank rows at a time.
'
'THE SHEET TAG IS READ THE CHEAP WAY
'-------------------------------------------------------------------------------
'The walk reads sheet_type through HiddenNames.QuickValue, which reads one
'name, and the full store is built for the HList sheets alone. A full store
'costs a walk of a sheet's whole Names collection, and the linelist filter
'refresh runs far more often than an export.
'@depends HiddenNames, LLLog, BetterArray
'@version 1.1 (2026-08-12)


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
'HList has its visible filled rows written into the filtered companion. A
'sheet whose table holds no data rows is left as it is. A sheet that raises is
'recorded in FailedSheets with the reason and the walk continues.
'@return Long. The number of sheets whose companion was rewritten.
'@export
Public Function Sync() As Long

    Dim sh As Worksheet
    Dim syncedCount As Long

    Set this.failedSheets = New BetterArray

    'The handler records the sheet and resumes the walk, so one broken
    'sheet leaves the others synced. The caller reads FailedSheets.
    On Error GoTo SheetFail

    For Each sh In this.sourceWkb.Worksheets
        If HiddenNames.QuickValue(sh, SHEET_TYPE_KEY) = HLIST_TAG Then
            If SyncSheet(sh) Then syncedCount = syncedCount + 1
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
'@sub-title Rewrite one filtered companion from its HList table
'@details
'Reads the companion name off the sheet store and hands the table and the
'companion sheet to the rewrite. Raises when the filtered_sheet name is
'missing, and lets any worksheet failure travel up to the per-sheet handler
'in Sync.
'@param sh Worksheet. The HList sheet to copy.
'@return Boolean. True when the companion was rewritten.
'@throws ProjectError.ElementNotFound When the filtered_sheet name is missing.
Private Function SyncSheet(ByVal sh As Worksheet) As Boolean

    Dim lo As ListObject
    Dim store As HiddenNames
    Dim filtsh As Worksheet
    Dim filteredName As String

    'Every lookup below used to be written straight, so a workbook missing any
    'one of these pieces answered a bare Excel number -- 9 for an absent sheet,
    '91 for an absent table -- and the box the user got said "Object variable or
    'With block variable not set" with nothing to act on. Each piece now names
    'itself, and the walk carries on with the other sheets.
    On Error Resume Next
    Set lo = sh.ListObjects(1)
    On Error GoTo 0

    If lo Is Nothing Then
        ThrowError ElementNotFound, "the sheet carries no table to read"
    End If

    If lo.DataBodyRange Is Nothing Then Exit Function

    Set store = HiddenNames.Create(sh)
    filteredName = store.ValueAsString(FILTERED_SHEET_KEY)
    If LenB(filteredName) = 0 Then
        ThrowError ElementNotFound, "the filtered_sheet name is missing"
    End If

    On Error Resume Next
    Set filtsh = this.sourceWkb.Worksheets(filteredName)
    On Error GoTo 0

    If filtsh Is Nothing Then
        ThrowError ElementNotFound, _
                   "the workbook carries no worksheet named " & filteredName
    End If

    RewriteCompanion lo, filtsh

    SyncSheet = True
End Function

'@label:rewrite-companion
'@sub-title Write the visible filled rows of one table into its companion
'@details
'A row is kept when it is visible and holds at least one value. Visibility is
'read off the source rows in one SpecialCells call. The read spans the whole
'row on purpose: a column hidden by show/hide keeps its values in the
'companion, and a row leaves when a filter or the user hides it. The values
'travel through memory -- one read of the whole block, the blank test on the
'array, and one write of the kept rows. The two tables sit at the same address
'on their two sheets, which is the layout the builder writes.
'@param sourceTable ListObject. The data entry table the rows come from.
'@param filtsh Worksheet. The sheet holding the filtered companion table.
Private Sub RewriteCompanion(ByVal sourceTable As ListObject, _
                             ByVal filtsh As Worksheet)

    Dim bodyRng As Range
    Dim visibleCells As Range
    Dim visibleArea As Range
    Dim filtLo As ListObject
    Dim sourceValues As Variant
    Dim singleValue As Variant
    Dim keptValues() As Variant
    Dim keepRow() As Boolean
    Dim rowCount As Long
    Dim columnCount As Long
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim keptCount As Long
    Dim firstRow As Long
    Dim keptIdx As Long

    Set bodyRng = sourceTable.DataBodyRange
    rowCount = bodyRng.Rows.Count
    columnCount = bodyRng.Columns.Count

    sourceValues = bodyRng.Value
    If (rowCount = 1) And (columnCount = 1) Then
        'A single cell answers a scalar, and the walk below wants a block.
        singleValue = sourceValues
        ReDim sourceValues(1 To 1, 1 To 1)
        sourceValues(1, 1) = singleValue
    End If

    'SpecialCells raises 1004 when every row is hidden, and an empty answer is
    'what that state means. The suppression covers that one call.
    ReDim keepRow(1 To rowCount)
    On Error Resume Next
    Set visibleCells = bodyRng.EntireRow.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleCells Is Nothing Then
        For Each visibleArea In visibleCells.Areas
            firstRow = visibleArea.Row - bodyRng.Row + 1
            For rowIdx = firstRow To firstRow + visibleArea.Rows.Count - 1
                If (rowIdx >= 1) And (rowIdx <= rowCount) Then keepRow(rowIdx) = True
            Next rowIdx
        Next visibleArea
    End If

    For rowIdx = 1 To rowCount
        If keepRow(rowIdx) Then
            If RowHoldsAValue(sourceValues, rowIdx, columnCount) Then
                keptCount = keptCount + 1
            Else
                keepRow(rowIdx) = False
            End If
        End If
    Next rowIdx

    If keptCount > 0 Then
        ReDim keptValues(1 To keptCount, 1 To columnCount)
        For rowIdx = 1 To rowCount
            If keepRow(rowIdx) Then
                keptIdx = keptIdx + 1
                For colIdx = 1 To columnCount
                    keptValues(keptIdx, colIdx) = sourceValues(rowIdx, colIdx)
                Next colIdx
            End If
        Next rowIdx
    End If

    On Error Resume Next
    Set filtLo = filtsh.ListObjects(1)
    On Error GoTo 0

    If filtLo Is Nothing Then
        ThrowError ElementNotFound, _
                   "the companion sheet " & filtsh.Name & " carries no table"
    End If

    'A table already down to its header row has nothing to delete.
    On Error Resume Next
    filtLo.DataBodyRange.Delete
    On Error GoTo 0

    If keptCount = 0 Then Exit Sub

    filtLo.Resize filtsh.Range( _
        sourceTable.HeaderRowRange.Resize(keptCount + 1).Address)

    'Resize is what gives the table its body back, and a resize the sheet
    'refused leaves the table on its header row alone. Writing to the body then
    'raised 91 with nothing naming the sheet it happened on.
    If filtLo.DataBodyRange Is Nothing Then
        ThrowError ElementNotFound, _
                   "the companion table on " & filtsh.Name & _
                   " could not be grown to " & CStr(keptCount) & " rows"
    End If

    filtLo.DataBodyRange.Value = keptValues
End Sub

'@label:row-holds-a-value
'@sub-title Say whether one row of a value block holds anything
'@details
'An empty cell and an empty string count as nothing, and every other value --
'a zero, a date, an error -- counts as data. This is what CountA answers on a
'worksheet row, read off the array instead.
'@param blockValues Variant. The value block, 1-based on both sides.
'@param rowIdx Long. The row to read.
'@param columnCount Long. The width of the block.
'@return Boolean. True when at least one cell of the row holds a value.
Private Function RowHoldsAValue(ByRef blockValues As Variant, _
                                ByVal rowIdx As Long, _
                                ByVal columnCount As Long) As Boolean

    Dim colIdx As Long
    Dim cellValue As Variant

    For colIdx = 1 To columnCount
        cellValue = blockValues(rowIdx, colIdx)
        If Not IsEmpty(cellValue) Then
            If VarType(cellValue) = vbString Then
                If LenB(cellValue) > 0 Then
                    RowHoldsAValue = True
                    Exit Function
                End If
            Else
                RowHoldsAValue = True
                Exit Function
            End If
        End If
    Next colIdx
End Function


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
