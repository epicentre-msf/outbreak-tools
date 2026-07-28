VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Wrapper for Listobjects manipulations"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Wrapper for Listobjects manipulations")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, MissingAnnotationArgument, ReadOnlyPropertyAssignment, UseMeaningfulName, HungarianNotation

Option Explicit


'@class CustomTable
'@description
'Wraps an Excel ListObject with convenience methods for column lookup,
'row insertion and deletion, sorting, validation, import/export, and
'diagnostic logging. Column lookups read the header row once as a block
'and match in memory, so nothing about the table shape is kept on the
'instance: most callers throw the wrapper away after one call, and the
'three that keep one alive see the sheet change behind their back. The
'class also supports snapshot-based rollback during import and tracks
'hidden-column state across import and export. An optional ID column
'provides dictionary-style keyed access to individual rows.
'@depends BetterArray, Checking, HiddenNames, DropdownLists
'@version 1.1 (2026-07-28)


Private Type TCustomTable
    lo As ListObject
    idCol As String
    idPrefix As String
    hasImport As Boolean
    hasChecks As Boolean
    backupHeaders As BetterArray
    backupRows As BetterArray
    hasBackup As Boolean
    hasHiddenList As Boolean
    hasShiftSetup As Boolean
    notImportedColumns As BetterArray
    hiddenColumnsList As BetterArray
    checkCounter As Long
    internalChecks As Checking
    sealed As Boolean
End Type

Private this As TCustomTable
Private Const TAG_INSERT_SHIFT_USED As String = "InsertShiftUsed"


'@section StateManagement
'===============================================================================
'@description Backup and guard helpers that keep internal state consistent.

'@label:reset-caches
'@sub-title Drop the state a structural change can make wrong
'@details
'Column lookups read the header row from the sheet on every call, so there
'is no lookup cache to drop here any more. The shift flag is the one piece
'of state a resize can leave saying the wrong thing.
Private Sub ResetCaches()
    this.hasShiftSetup = False
End Sub

'@label:reset-import-state
'@sub-title Reset import state so that imports start fresh
'@details
'Clears the import flag, hidden-column list, not-imported-columns list,
'and any saved backup so the next import begins from a clean slate.
Private Sub ResetImportState()
    this.hasImport = False
    this.hasHiddenList = False
    If this.notImportedColumns Is Nothing Then Set this.notImportedColumns = New BetterArray
    this.notImportedColumns.Clear
    If this.hiddenColumnsList Is Nothing Then Set this.hiddenColumnsList = New BetterArray
    this.hiddenColumnsList.Clear
    ClearBackup
End Sub

'@label:clear-backup
'@sub-title Release backup buffers
'@details
'Sets the backup header and row arrays to Nothing and clears the
'hasBackup flag once the operation completes successfully.
Private Sub ClearBackup()
    'Release backup buffers once the operation completes successfully.
    Set this.backupHeaders = Nothing
    Set this.backupRows = Nothing
    this.hasBackup = False
End Sub

'@label:save-table-snapshot
'@sub-title Save a snapshot of the current table state
'@details
'Captures the current headers and data rows into BetterArray buffers
'so RestoreTableSnapshot can revert the table if an import fails.
'Unhides hidden columns before reading, and leaves them unhidden: the
'caller puts them back once the whole operation is over.
Private Sub SaveTableSnapshot()
    'Capture the current table layout so we can restore it if an import fails.
    Dim lo As ListObject
    Dim bodyRng As Range

    UnhideHiddenColumns
    Set lo = Table()

    Set this.backupHeaders = New BetterArray
    this.backupHeaders.LowerBound = 1
    this.backupHeaders.FromExcelRange lo.HeaderRowRange, DetectLastRow:=False, DetectLastColumn:=False

    Set this.backupRows = New BetterArray
    this.backupRows.LowerBound = 1
    Set bodyRng = lo.DataBodyRange
    If Not bodyRng Is Nothing Then
        this.backupRows.FromExcelRange bodyRng, DetectLastRow:=False, DetectLastColumn:=False
    End If

    this.hasBackup = True
End Sub

'@label:restore-table-snapshot
'@sub-title Restore the table from a previously saved snapshot
'@details
'Reapplies the cached headers and rows when an import needs to roll
'back. Resizes the ListObject, writes the backup data, resets state,
'and re-hides previously hidden columns.
'@export
Public Sub RestoreTableSnapshot()
    'Reapply the cached headers and rows when an import needs to roll back.
    Dim lo As ListObject
    Dim headerRng As Range
    Dim bodyRng As Range
    Dim totalCols As Long
    Dim targetRows As Long
    Dim hasRows As Boolean
    Dim newRange As Range

    If Not this.hasBackup Then Exit Sub

    Set lo = Table()

    'One fetch: the row and the column below were read from Excel twice.
    Set headerRng = lo.HeaderRowRange
    totalCols = this.backupHeaders.Length

    UnhideHiddenColumns

    'SaveTableSnapshot always assigns a fresh BetterArray, so the only
    'question left is whether it holds any row.
    hasRows = (this.backupRows.Length > 0)
    If hasRows Then
        targetRows = this.backupRows.Length
    Else
        targetRows = 1
    End If

    Set newRange = SheetBlock(headerRng.Row, headerRng.Column, _
                              headerRng.Row + targetRows, _
                              headerRng.Column + totalCols - 1)
    lo.Resize newRange

    If Not this.backupHeaders Is Nothing Then
        this.backupHeaders.ToExcelRange lo.HeaderRowRange.Cells(1, 1), TransposeValues:=True
    End If

    If hasRows Then
        this.backupRows.ToExcelRange lo.DataBodyRange
    Else
        Set bodyRng = lo.DataBodyRange
        If Not bodyRng Is Nothing Then bodyRng.ClearContents
    End If

    ResetCaches
    ClearBackup

    ReturnBackHiddenColumns
End Sub


'@label:ensure-table-initialized
'@sub-title Guard against calls before Create has run
'@details
'Raises an error when the ListObject reference is Nothing, preventing
'all table operations from executing on an uninitialised instance.
'@throws ProjectError.ObjectNotInitialized When the ListObject is Nothing.
Private Sub EnsureTableInitialized()
    'Guard all table operations to avoid accidental calls before Create() runs.
    If this.lo Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "The ListObject reference is missing. Initialise the table first"
    End If
End Sub

'@section Factory
'===============================================================================
'@description Entry point for creating CustomTable instances.

'@label:create
'@sub-title Create a CustomTable wrapper around a ListObject
'@details
'Wraps the supplied ListObject with convenience methods for column
'lookup, row manipulation, import/export, and diagnostic logging. An
'optional ID column enables dictionary-style keyed row access. When
'adding rows, the ID column is automatically numbered in sequence,
'optionally prefixed (e.g. "line 1", "line 2").
'@param Lo ListObject. The Excel table to wrap.
'@param idCol Optional String. Name of the ID column. Defaults to vbNullString.
'@param idPrefix Optional String. Prefix for auto-generated ID values. Defaults to vbNullString.
'@return CustomTable. A fully initialised CustomTable instance.
'@throws ProjectError.ObjectNotInitialized When Lo is Nothing.
'@depends BetterArray, Checking
'@export
Public Function Create(ByVal Lo As ListObject, _
                       Optional ByVal idCol As String = vbNullString, _
                       Optional ByVal idPrefix As String = vbNullString) As CustomTable

    If Lo Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Unable to initialize the class, no listobject provided"

    Dim newTable As CustomTable
    Set newTable = New CustomTable

    Set newTable.Table = Lo
    newTable.IdValue = idCol
    newTable.PrefixValue = idPrefix
    newTable.Seal
    Set Create = newTable
End Function


'@section Elements
'===============================================================================
'@description Properties that expose the table structure and cell access.

'@label:wksh
'@prop-title Worksheet hosting the ListObject
'@details
'Returns the parent worksheet of the underlying ListObject.
'@return Worksheet. The host worksheet.
Private Property Get Wksh() As Worksheet
    Set Wksh = Table().Parent
End Property

'@label:sheet-block
'@sub-title Build a Range on the host worksheet from four numbers
'@details
'The "With sh ... .Range(.Cells(a, b), .Cells(c, d))" shape was written out
'in five places. One small builder replaces all of them. It is deliberately
'bound to this table's own worksheet, so it must never be used to build a
'range on an export target sheet.
'@param firstRow Long. Top row.
'@param firstColumn Long. Left column.
'@param lastRow Long. Bottom row.
'@param lastColumn Long. Right column.
'@return Range. The block.
Private Function SheetBlock(ByVal firstRow As Long, ByVal firstColumn As Long, _
                            ByVal lastRow As Long, ByVal lastColumn As Long) As Range
    With Wksh()
        Set SheetBlock = .Range(.Cells(firstRow, firstColumn), .Cells(lastRow, lastColumn))
    End With
End Function

'@label:row-values
'@sub-title Read a one-row Range as a 1-based list of values
'@details
'Range.Value hands back a plain value when the Range is one cell, and a
'two-dimensional array otherwise. This flattens both cases so callers see
'one shape and do not repeat the test.
'@param blockRng Range. A single-row block.
'@return Variant. A 1-based one-dimensional array of cell values.
Private Function RowValues(ByVal blockRng As Range) As Variant
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long
    Dim nbColumns As Long

    nbColumns = blockRng.Columns.Count
    ReDim flat(1 To nbColumns)

    If nbColumns = 1 Then
        flat(1) = blockRng.Cells(1, 1).Value
    Else
        rawValues = blockRng.Value
        For counter = 1 To nbColumns
            flat(counter) = rawValues(1, counter)
        Next
    End If

    RowValues = flat
End Function

'@label:column-values
'@sub-title Read a one-column Range as a 1-based list of values
'@details
'The column twin of RowValues, and for the same reason: a one-cell Range
'hands back a plain value where a taller one hands back a two-dimensional
'array.
'@param blockRng Range. A single-column block.
'@return Variant. A 1-based one-dimensional array of cell values.
Private Function ColumnValues(ByVal blockRng As Range) As Variant
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long
    Dim nbRows As Long

    nbRows = blockRng.Rows.Count
    ReDim flat(1 To nbRows)

    If nbRows = 1 Then
        flat(1) = blockRng.Cells(1, 1).Value
    Else
        rawValues = blockRng.Value
        For counter = 1 To nbRows
            flat(counter) = rawValues(counter, 1)
        Next
    End If

    ColumnValues = flat
End Function

'@label:cell-text
'@sub-title Read one cell value as text without raising
'@details
'A cell holding an error value (#N/A and the like) makes CStr raise a
'type mismatch. Such a header can never match a column name, so it reads
'as empty text instead of stopping the caller.
'@param cellValue Variant. The raw cell value.
'@return String. The value as text, or empty text.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@label:table
'@prop-title Underlying ListObject reference
'@details
'Returns the ListObject wrapped by this CustomTable. Raises an error
'when the reference has not been set.
'@return ListObject. The wrapped table.
'@throws ProjectError.ObjectNotInitialized When the ListObject is Nothing.
Public Property Get Table() As ListObject
    EnsureTableInitialized
    Set Table = this.lo
End Property

'@label:table-set
'@prop-title Assign the ListObject reference
'@details
'Stores the ListObject and resets state and backup. Raises an error when
'Lo is Nothing.
'@param Lo ListObject. The table to assign.
'@throws ProjectError.ObjectNotInitialized When Lo is Nothing.
Public Property Set Table(ByVal Lo As ListObject)
    GuardNotSealed "Table"
    If Lo Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Cannot assign an empty ListObject reference to CustomTable"
    End If
    Set this.lo = Lo
    ResetCaches
    ClearBackup
End Property

'@label:id-value
'@prop-title ID column name
'@details
'Returns the name of the column used as a unique row identifier.
'@return String. The ID column name.
'@export
Public Property Get IdValue() As String
    IdValue = this.idCol
End Property

'@label:prefix-value
'@prop-title ID prefix string
'@details
'Returns the prefix prepended to sequential ID values.
'@return String. The prefix text.
Public Property Get PrefixValue() As String
    PrefixValue = this.idPrefix
End Property

'@label:id-value-set
'@prop-title Assign the ID column name
'@param idCol String. Column name to use as the row key.
Public Property Let IdValue(ByVal idCol As String)
    GuardNotSealed "IdValue"
    this.idCol = idCol
End Property

'@label:prefix-value-set
'@prop-title Assign the ID prefix string
'@param idPrefix String. Prefix to prepend to sequential ID values.
Public Property Let PrefixValue(ByVal idPrefix As String)
    GuardNotSealed "PrefixValue"
    this.idPrefix = idPrefix
End Property

'@label:header-range
'@prop-title Header row Range of the CustomTable
'@details
'Returns the Range covering the header row of the underlying ListObject.
'@return Range. The header row range.
'@export
Public Property Get HeaderRange() As Range
    Set HeaderRange = Table().HeaderRowRange()
End Property

'@label:find-column
'@sub-title Find one column in the header row, reading that row once
'@details
'This is the single lookup behind ColumnIndex, and so behind every column
'name this class is given. The header row is read once as a block and
'matched in memory: Range.Find used to cost a worksheet search plus five
'property reads, and a name that matched nothing repeated the whole search
'on every call because misses were never remembered.
'The two match rules are kept exactly as Range.Find applied them:
'strictSearch True is whole-string equality (xlWhole) and False is a
'substring test (xlPart); matchCase True compares binary and False compares
'text. The leftmost match wins, which several callers rely on.
'Two Range.Find behaviours are deliberately gone. It read "*", "?" and "~"
'in the search text as wildcards even under xlWhole, and it reused whatever
'LookIn and SearchOrder the last search anywhere in Excel had used,
'including the user's own Ctrl+F. Both were faults.
'The block is NOT kept on the instance: callers write to these sheets
'without going through this object, so it is read again on the next call.
'@param headerRng Range. The header row to search.
'@param colName String. Column header to look for.
'@param strictSearch Boolean. True for whole-string matching.
'@param matchCase Boolean. True for case-sensitive matching.
'@return Long. The 1-based position within the header row, or 0 when there is none.
Private Function FindColumn(ByVal headerRng As Range, _
                            ByVal colName As String, _
                            ByVal strictSearch As Boolean, _
                            ByVal matchCase As Boolean) As Long

    Dim headers As Variant
    Dim currentHeader As String
    Dim compareMethod As Long
    Dim isMatch As Boolean
    Dim counter As Long

    'An empty name never matches, so answer it without touching the sheet.
    If LenB(colName) = 0 Then Exit Function

    compareMethod = IIf(matchCase, vbBinaryCompare, vbTextCompare)
    headers = RowValues(headerRng)

    For counter = 1 To UBound(headers)
        currentHeader = CellText(headers(counter))

        If strictSearch Then
            isMatch = (StrComp(currentHeader, colName, compareMethod) = 0)
        Else
            isMatch = (InStr(1, currentHeader, colName, compareMethod) > 0)
        End If

        If isMatch Then
            FindColumn = counter
            Exit Function
        End If
    Next
End Function

'@label:column-index
'@sub-title Resolve the 1-based column index for a header name
'@details
'Searches the header row for colName and returns its position within that
'row. Returns 0 when not found. When required is True, raises an error
'instead of returning 0.
'The 0-for-not-found answer is deliberate and stays: DataSheet answers -1,
'and lining the two up would make a "< 0" test written for a DataSheet
'silently pass for a CustomTable miss.
'@param colName String. Column name to look up.
'@param strictSearch Optional Boolean. When True, uses exact whole-string matching. Defaults to False.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
'@param required Optional Boolean. When True, raises an error if not found. Defaults to False.
'@return Long. The 1-based column index, or 0 when not found.
'@throws ProjectError.InvalidArgument When colName is empty and required is True.
'@throws ProjectError.ObjectNotInitialized When the header range is Nothing and required is True.
'@throws ProjectError.ElementNotFound When the column is not found and required is True.
Private Function ColumnIndex(ByVal colName As String, _
                             Optional ByVal strictSearch As Boolean = False, _
                             Optional ByVal matchCase As Boolean = False, _
                             Optional ByVal required As Boolean = False) As Long
    Dim hRng As Range

    If LenB(colName) = 0 Then
        If required Then
            ThrowError ProjectError.InvalidArgument, "Column name cannot be empty"
        End If
        Exit Function
    End If

    'One header fetch for the whole lookup. HeaderRange reaches the table
    'through Table(), which carries the initialisation guard.
    Set hRng = HeaderRange()
    If hRng Is Nothing Then
        If required Then
            ThrowError ProjectError.ObjectNotInitialized, _
                       "Unable to read headers for table '" & Name() & "'"
        End If
        Exit Function
    End If

    ColumnIndex = FindColumn(hRng, colName, strictSearch, matchCase)

    If ColumnIndex = 0 And required Then
        ThrowError ProjectError.ElementNotFound, _
                   "Column '" & colName & "' not found in table '" & Name() & "'"
    End If
End Function


'@label:column-exists
'@sub-title Check whether a column exists in the table
'@details
'Delegates to ColumnIndex and returns True when the index is positive.
'@param colName String. Column name to search.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to False.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
'@return Boolean. True when the column is found.
'@export
Public Function ColumnExists(ByVal colName As String, _
                             Optional ByVal strictSearch As Boolean = False, _
                             Optional ByVal matchCase As Boolean = False) As Boolean
    ColumnExists = (ColumnIndex(colName, strictSearch, matchCase) > 0)
End Function

'@label:resolve-list-column
'@sub-title Retrieve the ListColumn object for a header name
'@details
'Centralises ListColumn retrieval so every caller benefits from the
'same guards. Returns Nothing when not found unless required is True.
'@param colName String. Column header to resolve.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to False.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
'@param required Optional Boolean. When True, raises an error if not found. Defaults to False.
'@return ListColumn. The resolved column, or Nothing.
'@throws ProjectError.ElementNotFound When the column is not found and required is True.
Private Function ResolveListColumn(ByVal colName As String, _
                                   Optional ByVal strictSearch As Boolean = False, _
                                   Optional ByVal matchCase As Boolean = False, _
                                   Optional ByVal required As Boolean = False) As ListColumn
    'Centralise ListColumn retrieval so every caller benefits from the same guards.
    Dim columnIndexValue As Long

    columnIndexValue = ColumnIndex(colName, strictSearch, matchCase, required)
    If columnIndexValue <= 0 Then Exit Function

    'Guards error 9: the header row and ListColumns can disagree for a moment
    'while the table is being resized. Answer Nothing rather than raise.
    On Error Resume Next
        Set ResolveListColumn = Table().ListColumns(columnIndexValue)
    On Error GoTo 0

    If ResolveListColumn Is Nothing And required Then
        ThrowError ProjectError.ElementNotFound, _
                   "Unable to retrieve ListColumn '" & colName & "' on table '" & Name() & "'"
    End If
End Function


'@label:data-range
'@prop-title Range of a column or the entire data body
'@details
'Returns a data Range for the specified column, or the entire data body
'when colName is "__all__". When strictSearch is False, the search uses
'partial matching; the first partial match is returned. Use strictSearch
'when column names share common text to avoid unexpected output. Returns
'Nothing when the column has no data rows or is not found.
'Answering Nothing is part of the contract: four callers read it as "this
'table does not have that column", so it can never become a raise or an
'empty range.
'@param colName Optional String. Column header to retrieve. Defaults to "__all__".
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@param strictSearch Optional Boolean. When True, uses exact whole-string matching. Defaults to False.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
'@return Range. The requested data range, or Nothing when unavailable.
'@export
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False, _
                              Optional ByVal strictSearch As Boolean = False, _
                              Optional ByVal matchCase As Boolean = False) As Range
    Dim lo As ListObject
    Dim targetColumn As ListColumn
    Dim rng As Range

    Set lo = Table()

    If colName = "__all__" Then
        If includeHeaders Then
            Set rng = lo.Range
        ElseIf Not lo.DataBodyRange Is Nothing Then
            Set rng = lo.DataBodyRange
        End If
    Else
        Set targetColumn = ResolveListColumn(colName, strictSearch, matchCase)
        If targetColumn Is Nothing Then
            'A miss is a normal answer here and it happens often, so this
            'path stays a plain log with no debug trace behind it.
            LogInfo "Unable to find the column name " & colName & " of " & Name()
            Exit Property
        End If

        If includeHeaders Then
            Set rng = targetColumn.Range
        ElseIf Not targetColumn.DataBodyRange Is Nothing Then
            Set rng = targetColumn.DataBodyRange
        End If
    End If

    Set DataRange = rng
End Property

'@label:name
'@prop-title Name of the attached ListObject
'@details
'Reads the name from the live ListObject. It is deliberately not kept on
'the instance: the user can rename a table from the Excel interface at any
'time, and a name stored at creation would then be wrong for the rest of
'the object's life.
'@return String. The table name.
'@export
Public Property Get Name() As String
    Name = Table().Name
End Property

'@label:cell-range
'@prop-title A specific cell in the CustomTable
'@details
'Returns the cell at the intersection of the given column and worksheet
'row number. Returns Nothing when the column does not exist or the row
'is out of bounds.
'@param colName String. Column header identifying the column.
'@param lineNum Long. Worksheet row number (1-based).
'@return Range. The target cell, or Nothing when not found.
'@export
Public Property Get CellRange(ByVal colName As String, _
                              ByVal lineNum As Long) As Range
    Dim cellPos As Long
    Dim rng As Range

    'One lookup. DataRange already answers Nothing when the column is
    'missing, so the ColumnExists call that used to sit here paid for the
    'whole header search a second time.
    Set rng = DataRange(colName)
    If rng Is Nothing Then Exit Property
    cellPos = lineNum - rng.Row + 1
    If cellPos < 1 Or cellPos > rng.Rows.Count Then Exit Property
    Set CellRange = rng.Cells(cellPos, 1)
End Property


'@label:value-range
'@sub-title Locate the cell Range by column and key
'@details
'Returns the Range at the intersection of the specified column and the
'row matching keyName in the ID column. Returns Nothing when the key or
'column is not found, or when no ID column is configured.
'@param colName String. Column header identifying the column.
'@param keyName String. Key value identifying the row.
'@return Range. The target cell range, or Nothing when not found.
'@export
Public Function ValueRange(ByVal colName As String, ByVal keyName As String) As Range

    Dim idRng As Range
    Dim colRng As Range
    Dim keyColumnValues As Variant
    Dim rowCount As Long
    Dim counter As Long

    If IdValue() = vbNullString Then Exit Function

    'One lookup per column. The pair of ColumnExists calls that used to open
    'this routine made it pay for four header searches instead of two.
    Set idRng = DataRange(IdValue())
    If idRng Is Nothing Then Exit Function
    Set colRng = DataRange(colName)
    If colRng Is Nothing Then Exit Function

    rowCount = idRng.Rows.Count
    If rowCount = 0 Then Exit Function
    If colRng.Rows.Count = 0 Then Exit Function

    'Read the ID column once and match in memory. The old loop re-read its
    'bound, the cell and the row from the worksheet on every pass, about six
    'round trips a row, and Value / SetValue / CellRange all come through
    'here.
    keyColumnValues = ColumnValues(idRng)

    For counter = 1 To rowCount
        'Compare the raw value. Wrapping it in CStr would change the answer
        'for a numeric ID column.
        If keyColumnValues(counter) = keyName Then
            Set ValueRange = colRng.Cells(counter, 1)
            Exit Function
        End If
    Next
End Function

'@label:value
'@prop-title Retrieve a cell value by column and key
'@details
'For tables with a key column, returns the cell value at the intersection
'of the specified column and the row matching keyName. Returns an empty
'string when the key or column is not found.
'@param colName String. Column header to read from.
'@param keyName String. Key value identifying the row.
'@return String. The cell value converted to string, or vbNullString.
'@export
Public Property Get Value(ByVal colName As String, _
                          ByVal keyName As String) As String
    Dim retValue As String
    Dim cellRetRng As Range

    'Get the value Range and retrieve it. When there is no such cell the
    'returned value is the null string.
    Set cellRetRng = ValueRange(colName, keyName)
    If Not cellRetRng Is Nothing Then
        retValue = CStr(cellRetRng.Value)
    End If
    Value = retValue
End Property


'@section Modify
'===============================================================================
'@description Methods that add, remove, or reorganise table rows and columns.

'@label:add-shift-tracker
'@sub-title Record that worksheet row insertion was used
'@details
'Persists a hidden name on the worksheet so future operations know to
'delete entire worksheet rows instead of just ListRows.
'@param value Optional String. Tracker value ("Yes" or "No"). Defaults to "Yes".
Private Sub AddShiftTracker(Optional ByVal value As String = "Yes")

    Dim store As HiddenNames
    Dim wanted As Boolean

    wanted = (value = "Yes")

    'HiddenNames.Create walks every Name on the sheet and reads several
    'properties off each one. When the flag already says Yes the sheet
    'already says Yes too, so nothing has to be built.
    If wanted And this.hasShiftSetup Then Exit Sub

    'The flag has to follow the value both ways. It used to be set on "Yes"
    'and never cleared on "No", so one "No" left it saying Yes for the rest
    'of the object's life.
    this.hasShiftSetup = wanted

    Set store = HiddenNames.Create(Wksh())
    If Not store.HasName(TAG_INSERT_SHIFT_USED) Then
        store.EnsureName TAG_INSERT_SHIFT_USED, value, HiddenNameTypeString
    Else
        store.SetValue TAG_INSERT_SHIFT_USED, value
    End If

End Sub


'@label:has-shift-tracker
'@sub-title Check whether worksheet row insertion was previously used
'@details
'Reads the hidden name from the worksheet to determine whether prior
'operations used worksheet row insertion.
'@return Boolean. True when the shift tracker is active.
Private Function HasShiftTracker() As Boolean

    Dim store As HiddenNames

    'Only build the store when the answer is not already known: Create walks
    'every Name on the sheet, and it used to be built and thrown away unused
    'on every call where the flag already said Yes.
    If Not this.hasShiftSetup Then
        Set store = HiddenNames.Create(Wksh())
        this.hasShiftSetup = (store.ValueAsString(TAG_INSERT_SHIFT_USED, "No") = "Yes")
    End If

    HasShiftTracker = this.hasShiftSetup

End Function

'@label:resize
'@sub-title Resize the table by adding or removing rows
'@details
'Dispatches to AddTableRows or RemoveEmptyDataRows depending on the
'shouldAdd flag, then optionally renumbers the ID column.
'@param shouldAdd Optional Boolean. When True, adds rows; when False, removes empty rows. Defaults to False.
'@param insertShift Optional Boolean. When True, inserts worksheet rows. Defaults to True.
'@param totalRowCount Optional Long. Non-empty cell threshold for row removal. Defaults to 0.
'@param nbRows Optional Long. Number of rows to add. Defaults to 5.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion. Defaults to False.
Private Sub Resize(Optional ByVal shouldAdd As Boolean = False, _
                   Optional ByVal insertShift As Boolean = True, _
                   Optional ByVal totalRowCount As Long = 0, _
                   Optional ByVal nbRows As Long = 5, _
                   Optional ByVal includeIds As Boolean = True, _
                   Optional ByVal forceShift As Boolean = False)
    Dim lo As ListObject

    Set lo = Table()

    If shouldAdd Then
        AddTableRows lo, nbRows, insertShift
    Else
        RemoveEmptyDataRows lo, totalRowCount, forceShift
    End If

    If includeIds Then AddIds
End Sub

'@label:add-table-rows
'@sub-title Physically add rows to the ListObject
'@details
'Inserts worksheet rows below the table when insertShift is True
'(protecting stacked tables), otherwise extends the ListObject range
'directly. Resets state after the operation.
'@param lo ListObject. The table to extend.
'@param nbRows Long. Number of rows to add.
'@param insertShift Boolean. When True, inserts worksheet rows.
Private Sub AddTableRows(ByVal lo As ListObject, ByVal nbRows As Long, ByVal insertShift As Boolean)
    'ListRows.Add preserves formulas and formats; inserting worksheet rows prevents overlaps when tables stack.
    Dim tableRng As Range
    Dim headerRow As Long
    Dim headerCol As Long
    Dim totalCols As Long
    Dim totalRows As Long
    Dim insertStart As Long
    Dim newRange As Range

    If nbRows <= 0 Then Exit Sub

    'One fetch of the table range: the four numbers below used to be read
    'from Excel four separate times.
    Set tableRng = lo.Range
    headerRow = tableRng.Row
    headerCol = tableRng.Column
    totalCols = tableRng.Columns.Count
    totalRows = tableRng.Rows.Count

    If insertShift Then
        'Insert a block of rows below the current table so a table stacked
        'underneath is pushed down instead of being written over.
        insertStart = headerRow + totalRows
        Wksh().Rows(insertStart).Resize(nbRows).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Set newRange = SheetBlock(headerRow, headerCol, _
                                  insertStart + nbRows - 1, headerCol + totalCols - 1)
        AddShiftTracker "Yes"
    Else
        'Nothing is stacked below, so one resize covers it.
        Set newRange = SheetBlock(headerRow, headerCol, _
                                  headerRow + totalRows + nbRows - 1, headerCol + totalCols - 1)
    End If

    lo.Resize newRange

    ResetCaches
End Sub

'@label:ensure-selection-on-table-sheet
'@sub-title Ensure the provided range belongs to the table worksheet
'@details
'Validates that targetCell is not Nothing and resides on the same
'worksheet as the wrapped ListObject.
'@param targetCell Range. The range to validate.
'@throws ProjectError.InvalidArgument When targetCell is Nothing or belongs to a different sheet.
Private Sub EnsureSelectionOnTableSheet(ByVal targetCell As Range)
    Dim targetSheet As Worksheet

    If targetCell Is Nothing Then
        ThrowError ProjectError.InvalidArgument, _
                   "Target range cannot be empty when modifying table '" & Name() & "'"
    End If

    'A live Range always answers Worksheet, so the error trap that used to
    'wrap this line could never fire.
    Set targetSheet = targetCell.Worksheet

    If targetSheet Is Nothing Then
        ThrowError ProjectError.InvalidArgument, _
                   "Unable to resolve the worksheet hosting the supplied range"
    End If

    If Not targetSheet Is Wksh() Then
        ThrowError ProjectError.InvalidArgument, _
                   "Target range must belong to worksheet '" & Wksh().Name & "'"
    End If
End Sub

'@label:selection-row-numbers
'@sub-title Resolve worksheet row numbers intersecting the table
'@details
'Returns a BetterArray of worksheet row numbers where targetCell
'intersects the table scope. When restrictToDataRows is True, only
'data rows are considered; otherwise the full table range is used.
'@param targetCell Range. The selection to intersect.
'@param restrictToDataRows Boolean. When True, limits to data rows only.
'@return BetterArray. Row numbers, or Nothing when no intersection.
Private Function SelectionRowNumbers(ByVal targetCell As Range, _
                                     ByVal restrictToDataRows As Boolean) As BetterArray
    Dim lo As ListObject
    Dim scopeRange As Range
    Dim intersectRange As Range
    Dim areaRange As Range
    Dim rowsList As BetterArray
    Dim firstRow As Long
    Dim counter As Long

    Set lo = Table()

    If restrictToDataRows Then
        Set scopeRange = lo.DataBodyRange
    Else
        Set scopeRange = lo.Range
    End If

    If scopeRange Is Nothing Then Exit Function

    'One intersection for the whole selection. Intersecting once per selected
    'row instead means that a user who clicks a column header and then the
    'delete-rows button makes Excel do it over a million times.
    Set intersectRange = Intersect(targetCell, scopeRange)
    If intersectRange Is Nothing Then Exit Function

    Set rowsList = New BetterArray
    rowsList.LowerBound = 1

    'The areas come out in selection order and both callers walk the list
    'backwards, so the order the numbers land in must not change.
    For Each areaRange In intersectRange.Areas
        firstRow = areaRange.Row
        For counter = 0 To areaRange.Rows.Count - 1
            rowsList.Push firstRow + counter
        Next
    Next

    If rowsList.Length = 0 Then Exit Function

    Set SelectionRowNumbers = rowsList
End Function

'@label:insert-rows-at
'@sub-title Insert rows at the selected position
'@details
'Inserts rows matching the height of targetCell. Uses worksheet row
'insertion when insertShift is True to protect stacked tables, otherwise
'adds ListRows at the selection anchor. Renumbers the ID column when
'includeIds is True.
'@param targetCell Range. Selection anchoring the insertion point.
'@param insertShift Optional Boolean. When True, inserts worksheet rows. Defaults to False.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@export
Public Sub InsertRowsAt(ByVal targetCell As Range, _
                        Optional ByVal insertShift As Boolean = False, _
                        Optional ByVal includeIds As Boolean = True)
    Dim lo As ListObject
    Dim sheet As Worksheet
    Dim rowNumbers As BetterArray
    Dim rowCount As Long
    Dim headerRow As Long
    Dim idx As Long
    Dim targetRow As Long
    Dim relativePosition As Long

    EnsureSelectionOnTableSheet targetCell
    Set lo = Table()
    Set sheet = Wksh()

    'SelectionRowNumbers answers Nothing when nothing intersected, so the
    'empty-length test that used to follow could never fire.
    Set rowNumbers = SelectionRowNumbers(targetCell, restrictToDataRows:=False)
    If rowNumbers Is Nothing Then Exit Sub

    rowCount = rowNumbers.Length
    headerRow = lo.HeaderRowRange.Row

    If lo.DataBodyRange Is Nothing Then
        AddTableRows lo, rowCount, insertShift
    ElseIf insertShift Then
        AddShiftTracker "Yes"
        'Walk from the bottom so an insertion does not move the rows that
        'are still to be handled.
        For idx = rowNumbers.UpperBound To rowNumbers.LowerBound Step -1
            targetRow = CLng(rowNumbers.Item(idx))
            If targetRow <= headerRow Then targetRow = headerRow + 1
            sheet.Rows(targetRow).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next idx
    Else
        For idx = rowNumbers.UpperBound To rowNumbers.LowerBound Step -1
            targetRow = CLng(rowNumbers.Item(idx))
            If targetRow <= headerRow Then targetRow = headerRow + 1
            relativePosition = targetRow - headerRow
            If relativePosition < 1 Then relativePosition = 1

            If relativePosition > lo.ListRows.Count Then
                lo.ListRows.Add
            Else
                lo.ListRows.Add Position:=relativePosition
            End If
        Next idx
    End If

    ResetCaches
    If includeIds Then AddIds
End Sub



'@label:get-total-formula-columns
'@sub-title Count the number of columns containing formulas
'@details
'Scans the first data row and counts cells that have formulas. Used
'to determine the baseline non-empty cell count for row removal.
'This used to be a Sub with a ByRef argument, and its only caller wrapped
'that argument in brackets, which passes a copy: the answer never came
'back and the count was always zero. A Function cannot be called that way.
'@return Long. The number of columns whose first data cell holds a formula.
Private Function GetTotalFormulaColumns() As Long
    Dim rng As Range
    Dim rowHasFormula As Variant
    Dim counter As Long

    Set rng = Table().DataBodyRange
    If rng Is Nothing Then Exit Function

    'HasFormula on the whole first row answers all three cases in one round
    'trip: False when no cell holds a formula, True when they all do, and
    'Null when it is mixed. Only the mixed case needs the per-column walk.
    rowHasFormula = rng.Rows(1).HasFormula

    If IsNull(rowHasFormula) Then
        For counter = 1 To rng.Columns.Count
            If rng.Cells(1, counter).HasFormula Then
                GetTotalFormulaColumns = GetTotalFormulaColumns + 1
            End If
        Next
    ElseIf rowHasFormula Then
        GetTotalFormulaColumns = rng.Columns.Count
    End If
End Function


'@label:remove-empty-data-rows
'@sub-title Delete data rows that are effectively empty
'@details
'Scans from the bottom so deletions do not disturb subsequent indexes.
'A row is considered empty when its non-empty cell count is at or below
'totalRowCount. Deletes entire worksheet rows when the shift tracker is
'active or forceShift is True; otherwise removes ListRows directly.
'Always preserves the first data row to keep formulas and templates.
'@param lo ListObject. The table to clean.
'@param totalRowCount Optional Long. Non-empty cell threshold.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion.
Private Sub RemoveEmptyDataRows(ByVal lo As ListObject, _
                                Optional ByVal totalRowCount As Long, _
                                Optional ByVal forceShift As Boolean)
    'Scan from the bottom so deletions do not disturb subsequent indexes.
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim filledCells As Long
    Dim tabRowCount As Long
    Dim tabColumnCount As Long
    Dim rowCount As Long
    Dim bodyRng As Range
    Dim bodyValues As Variant
    Dim sh As Worksheet
    Dim deleteWholeRow As Boolean
    Dim deleteError As Long

    If totalRowCount = 0 Then
        rowCount = GetTotalFormulaColumns()
    Else
        rowCount = totalRowCount
    End If

    deleteWholeRow = (HasShiftTracker() Or forceShift)

    Set bodyRng = lo.DataBodyRange
    If bodyRng Is Nothing Then Exit Sub

    tabRowCount = bodyRng.Rows.Count
    'Keep the first data row: it carries the formulas and the formats every
    'new row is built from.
    If tabRowCount <= 1 Then Exit Sub

    tabColumnCount = bodyRng.Columns.Count

    'Read the body once. The old loop re-fetched DataBodyRange and ran CountA
    'per row, about three round trips a row before anything was deleted.
    bodyValues = bodyRng.Value

    'Hoisted out of the loop: the sheet was resolved again on every deletion.
    Set sh = Wksh()

    For rowIndex = tabRowCount To 2 Step -1
        filledCells = 0
        For columnIndex = 1 To tabColumnCount
            'This counts exactly what CountA counted: a formula that returns
            '"" reads back from a block as "", not as Empty, so it counts on
            'both sides.
            If Not IsEmpty(bodyValues(rowIndex, columnIndex)) Then
                filledCells = filledCells + 1
            End If
        Next

        If filledCells <= rowCount Then
            If deleteWholeRow Then
                deleteError = 0
                'Deleting the worksheet row already removes the ListRow and
                'pulls everything below up by one. The ListRows delete that
                'used to follow unconditionally therefore hit the row below,
                'which may hold data. It is now only a fallback.
                On Error Resume Next
                    sh.Rows(lo.ListRows(rowIndex).Range().Row).Delete
                    deleteError = Err.Number
                    Err.Clear
                On Error GoTo 0

                If deleteError <> 0 Then
                    On Error Resume Next
                        lo.ListRows(rowIndex).Delete
                        deleteError = Err.Number
                        Err.Clear
                    On Error GoTo 0
                End If

                'Raise outside the protected block. Raised inside it, VBA
                'resumes at the next statement and the caller never sees it.
                If deleteError <> 0 Then
                    ThrowError ProjectError.SomethingWentWrong, _
                               "Unable to resize the table (maybe tables are stacked?)"
                End If
            Else
                'Either delete the entire worksheet row or the listRow
                lo.ListRows(rowIndex).Delete
            End If
        End If
    Next rowIndex
End Sub

'@label:delete-rows-at
'@sub-title Delete rows intersecting the selection
'@details
'Removes data rows that intersect targetCell. Uses worksheet row
'deletion when forceShift or the stored shift tracker is active,
'otherwise removes ListRows directly. Adds a placeholder row when
'all rows are deleted. Renumbers the ID column when includeIds is True.
'@param targetCell Range. Selection identifying which rows to delete.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion. Defaults to False.
'@export
Public Sub DeleteRowsAt(ByVal targetCell As Range, _
                        Optional ByVal includeIds As Boolean = True, _
                        Optional ByVal forceShift As Boolean = False)
    Dim lo As ListObject
    Dim rowNumbers As BetterArray
    Dim counter As Long
    Dim headerRow As Long
    Dim relativeIndex As Long
    Dim targetRow As Long
    Dim shouldShift As Boolean
    Dim sheet As Worksheet
    Dim pendingRowCount As Long

    EnsureSelectionOnTableSheet targetCell
    Set lo = Table()

    If lo.DataBodyRange Is Nothing Then Exit Sub

    'SelectionRowNumbers answers Nothing when nothing intersected, so the
    'empty-length test that used to follow could never fire.
    Set rowNumbers = SelectionRowNumbers(targetCell, restrictToDataRows:=True)
    If rowNumbers Is Nothing Then Exit Sub

    shouldShift = forceShift Or HasShiftTracker()
    Set sheet = Wksh()
    headerRow = lo.HeaderRowRange.Row
    pendingRowCount = lo.ListRows.Count

    'Walk from the bottom so a deletion does not move the rows still to go.
    If shouldShift Then
        For counter = rowNumbers.UpperBound To rowNumbers.LowerBound Step -1
            targetRow = CLng(rowNumbers.Item(counter))
            sheet.Rows(targetRow).Delete
            pendingRowCount = pendingRowCount - 1
        Next counter
    Else
        For counter = rowNumbers.UpperBound To rowNumbers.LowerBound Step -1
            targetRow = CLng(rowNumbers.Item(counter))
            relativeIndex = targetRow - headerRow
            If relativeIndex >= 1 And relativeIndex <= lo.ListRows.Count Then
                lo.ListRows(relativeIndex).Delete
                pendingRowCount = pendingRowCount - 1
            End If
        Next counter
    End If

    'A ListObject cannot live with no data row, so put a blank one back.
    If pendingRowCount <= 0 Then
        lo.ListRows.Add
    End If

    ResetCaches
    If includeIds Then AddIds
End Sub

'@label:data-row-count
'@sub-title Count the number of data rows in the table
'@return Long. The data row count (excluding headers), or 0 when empty.
Private Function DataRowCount() As Long
    'Return the number of data rows currently stored in the table (without headers).
    Dim body As Range

    Set body = Table().DataBodyRange
    If body Is Nothing Then Exit Function

    DataRowCount = body.Rows.Count
End Function

'@label:data-column-count
'@sub-title Count the number of columns in the table
'@return Long. The column count.
Private Function DataColumnCount() As Long
    DataColumnCount = Table().ListColumns.Count
End Function

'@label:ensure-row-capacity
'@sub-title Guarantee the table has at least the requested data rows
'@param requiredRows Long. Minimum number of data rows needed.
'@param insertShift Boolean. When True, inserts worksheet rows.
Private Sub EnsureRowCapacity(ByVal requiredRows As Long, ByVal insertShift As Boolean)
    'Guarantee the table contains at least the requested number of data rows.
    Dim currentRows As Long
    Dim trimedRows As Long

    trimedRows = Application.WorksheetFunction.Max(1, requiredRows)
    'DataRowCount never answers below zero, so the old Max(0, ...) clamp
    'around it could not bind.
    currentRows = DataRowCount()

    If currentRows >= trimedRows Then Exit Sub

    AddTableRows Table(), trimedRows - currentRows, insertShift
End Sub

'@label:ensure-column-capacity
'@sub-title Guarantee the table has exactly the requested column count
'@details
'Adds or removes columns until the table matches requiredColumns. Trims
'excess columns using a descending loop so index shifts do not skip columns.
'@param requiredColumns Long. Target number of columns.
Private Sub EnsureColumnCapacity(ByVal requiredColumns As Long)

    Dim currCols As Long
    Dim lo As ListObject

    Set lo = Table()
    'A ListObject always carries at least one column and the caller always
    'asks for at least one, so the two Max() clamps that used to sit here
    'could not bind.
    currCols = DataColumnCount()

    'Add the missing columns before copying incoming headers so we avoid
    'Excel-provided defaults like "Column1".
    Do While currCols < requiredColumns
        lo.ListColumns.Add
        currCols = currCols + 1
    Loop

    'Trim excess columns using a descending loop so index shifts do not skip columns.
    Do While currCols > requiredColumns
        'A header-only table has no data body to clear.
        On Error Resume Next
            lo.ListColumns(currCols).DataBodyRange.Cells.Clear
        On Error GoTo 0
        lo.ListColumns(currCols).Delete
        currCols = currCols - 1
    Loop
End Sub

'@label:trim-table-rows
'@sub-title Shrink the table to the requested number of data rows
'@details
'Resizes the ListObject so only targetRows data rows remain.
'Exits when the table already has fewer rows than the target.
'@param targetRows Long. Desired data row count.
Private Sub TrimTableRows(ByVal targetRows As Long)
    'Shrink the table so only the requested number of data rows remain.
    Dim lo As ListObject
    Dim headerRng As Range
    Dim newRange As Range
    Dim normalizedTarget As Long

    Set lo = Table()

    normalizedTarget = Application.WorksheetFunction.Max(1, targetRows)
    If DataRowCount() <= normalizedTarget Then Exit Sub

    'One fetch: the three numbers below were read from Excel three times.
    Set headerRng = lo.HeaderRowRange
    Set newRange = SheetBlock(headerRng.Row, headerRng.Column, _
                              headerRng.Row + normalizedTarget, _
                              headerRng.Column + headerRng.Columns.Count - 1)
    lo.Resize newRange
    ResetCaches
End Sub

'@label:source-row-count
'@sub-title Derive the data row count from an import source
'@details
'Reads the DataRange from the supplied CustomTable or DataSheet and
'returns the number of data rows. Returns 0 when the source is empty.
'@param impTab Object. A CustomTable or DataSheet to measure.
'@return Long. The number of source data rows.
Private Function SourceRowCount(ByVal impTab As Object) As Long
    'Derive the number of data rows provided by the import source (CustomTable or DataSheet).
    Dim srcRange As Range

    'Guards error 438 and 91: the source is reached late-bound, and a sheet
    'with no data can refuse the call outright.
    On Error Resume Next
        Set srcRange = impTab.DataRange(colName:="__all__", includeHeaders:=False)
    On Error GoTo 0

    If srcRange Is Nothing Then Exit Function

    SourceRowCount = srcRange.Rows.Count
End Function

'@label:add-ids
'@sub-title Fill the ID column with sequential values
'@details
'Writes sequential numbers (optionally prefixed) to every cell in the
'ID column. Silently exits when no ID column is configured or when the
'column does not exist. The ID column is typically locked and not
'modified by the user directly.
'@export
Public Sub AddIds()
    Dim rng As Range
    Dim counter As Long
    Dim rowCount As Long
    Dim idColumnName As String
    Dim prefixText As String
    Dim idColumn As ListColumn
    Dim newIdValues() As Variant

    idColumnName = IdValue()
    If idColumnName = vbNullString Then Exit Sub

    'One lookup. ResolveListColumn answers Nothing when the column is
    'missing, and its whole-string match is narrower than the loose
    'ColumnExists test that used to run first, so that test could only ever
    'pay for the same search twice.
    Set idColumn = ResolveListColumn(idColumnName, strictSearch:=True)
    If idColumn Is Nothing Then Exit Sub
    Set rng = idColumn.DataBodyRange
    If rng Is Nothing Then Exit Sub

    prefixText = Trim$(PrefixValue())
    rowCount = rng.Rows.Count

    'Build the whole column in memory and write it in one statement. Cell by
    'cell this cost two round trips a row, and it is the tail of Resize,
    'InsertRowsAt, DeleteRowsAt and Import: one AddRows click on a 5000-row
    'linelist spent 10000 round trips only renumbering.
    ReDim newIdValues(1 To rowCount, 1 To 1)
    For counter = 1 To rowCount
        newIdValues(counter, 1) = FormatIdValue(prefixText, counter)
    Next

    rng.Value = newIdValues
End Sub

'@label:format-id-value
'@sub-title Build an ID string from a prefix and numeric part
'@param prefixText String. Prefix text, or empty for numeric-only IDs.
'@param numericPart Long. The sequential number.
'@return String. The formatted ID value.
Private Function FormatIdValue(ByVal prefixText As String, ByVal numericPart As Long) As String
    'Build the ID value ensuring spacing is consistent when a prefix is supplied.
    If LenB(prefixText) = 0 Then
        FormatIdValue = CStr(numericPart)
    Else
        FormatIdValue = prefixText & " " & CStr(numericPart)
    End If
End Function


'@label:add-rows
'@sub-title Add rows to the CustomTable
'@details
'Appends the requested number of rows, either by inserting worksheet
'rows (to protect stacked tables) or by extending the ListObject range
'directly. Insertion is useful when multiple ListObjects share a
'worksheet and overlaps must be avoided. Optionally renumbers the ID
'column afterwards.
'@param nbRows Optional Long. Number of rows to add. Defaults to 5.
'@param insertShift Optional Boolean. When True, inserts worksheet rows. Defaults to False.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@export
Public Sub AddRows(Optional ByVal nbRows As Long = 5, _
                   Optional ByVal insertShift As Boolean = False, _
                   Optional ByVal includeIds As Boolean = True)

    Resize shouldAdd:=True, nbRows:=nbRows, insertShift:=insertShift, _
           includeIds:=includeIds
End Sub


'@label:remove-rows
'@sub-title Remove empty rows from the CustomTable
'@details
'Deletes rows whose non-empty cell count falls at or below totalCount.
'In the linelist, formula cells count as non-empty, so totalCount
'should match the number of formula columns. When totalCount is left at
'0 the class counts the formula columns itself. Optionally renumbers the
'ID column afterwards.
'@param totalCount Optional Long. Maximum non-empty cells for a row to be considered empty. Defaults to 0.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion. Defaults to False.
'@export
Public Sub RemoveRows(Optional ByVal totalCount As Long = 0, _
                      Optional ByVal includeIds As Boolean = True, _
                      Optional ByVal forceShift As Boolean = False)

    Resize shouldAdd:=False, totalRowCount:=totalCount, includeIds:=includeIds, forceShift:=forceShift
End Sub

'@label:set-validation
'@sub-title Set dropdown validation on a column
'@details
'Applies data validation from the supplied DropdownLists to every cell
'in the specified column. Silently exits when the column does not exist
'or has no data range.
'@param colName String. Column header to validate.
'@param drop DropdownLists. Dropdown source providing the validation list.
'@param dropName String. Name of the dropdown list within the source.
'@param alertType Optional String. Alert severity ("info", "warning", "error"). Defaults to "info".
'@param message Optional String. Message shown when validation fails. Defaults to vbNullString.
'@export
Public Sub SetValidation(ByVal colName As String, _
                         ByVal drop As DropdownLists, _
                         ByVal dropName As String, _
                         Optional ByVal alertType As String = "info", _
                         Optional ByVal message As String = vbNullString)

    Dim rng As Range

    'One lookup. DataRange already answers Nothing when the column is
    'missing, so the ColumnExists test that used to open this routine paid
    'for the same header search twice.
    Set rng = DataRange(colName)
    If rng Is Nothing Then Exit Sub
    'set validation for each cell in the datarange
    drop.SetValidation cellRng:=rng, listName:=dropName, _
                       alertType:=alertType, message:=message
End Sub


'@label:set-value
'@sub-title Change the value of one specific cell
'@details
'Writes newValue to the cell identified by column name and key.
'Silently exits when the key or column is not found. Use with caution
'as it modifies data inside the ListObject.
'@param colName String. Column header identifying the cell.
'@param keyName String. Key value identifying the row.
'@param newValue String. The new value to assign.
'@export
Public Sub SetValue(ByVal colName As String, ByVal keyName As String, ByVal newValue As String)

    Dim cellRetRng As Range

    'Get the value Range and assign it the newValue. When there is no such
    'cell nothing is done.
    Set cellRetRng = ValueRange(colName, keyName)
    If cellRetRng Is Nothing Then Exit Sub
    cellRetRng.Value = newValue
End Sub

'@label:rename-column
'@sub-title Rename a column header within the table
'@details
'Replaces the header text of the matched column with newName.
'Raises an error when newName is empty.
'The strict default here is the opposite of the loose default on
'ColumnIndex and DataRange, and that is on purpose: it is what stops a
'caller asking for "Label" from renaming "Formula Label".
'@param colName String. Current column header to find.
'@param newName String. New header text to assign.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
'@throws ProjectError.InvalidArgument When newName is empty.
'@export
Public Sub RenameColumn(ByVal colName As String, _
                        ByVal newName As String, _
                        Optional ByVal strictSearch As Boolean = True, _
                        Optional ByVal matchCase As Boolean = False)

    Dim targetColumn As ListColumn

    If LenB(newName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "New column name cannot be empty"
    End If

    Set targetColumn = ResolveListColumn(colName, strictSearch, matchCase)
    If targetColumn Is Nothing Then Exit Sub

    targetColumn.Name = newName
    ResetCaches
End Sub

'@label:clean
'@sub-title Clear all non-formula, non-key columns
'@details
'Iterates every column in the table and clears the contents of those
'that do not contain formulas. The key column is preserved: it holds
'plain values, and clearing it wiped every ID whenever Clean was called
'outside Import.
'Each column is taken from its own position in the loop. Looking the
'header text back up cost about nine round trips and a worksheet search
'for a column whose place the loop already knew.
'@export
Public Sub Clean()
    Dim lo As ListObject
    Dim rng As Range
    Dim headers As Variant
    Dim idColumnName As String
    Dim keepColumn As Boolean
    Dim headerCount As Long
    Dim counter As Long

    Set lo = Table()
    idColumnName = IdValue()

    'The header row is only needed to spot the key column.
    If LenB(idColumnName) > 0 Then
        headers = RowValues(HeaderRange())
        headerCount = UBound(headers)
    End If

    'Clean only columns without formulas that are not keys.
    For counter = 1 To lo.ListColumns.Count
        keepColumn = False
        If counter <= headerCount Then
            'Matched the same way AddIds resolves the key column.
            keepColumn = (StrComp(CellText(headers(counter)), idColumnName, vbTextCompare) = 0)
        End If

        If Not keepColumn Then
            Set rng = lo.ListColumns(counter).DataBodyRange
            If Not rng Is Nothing Then
                'The first cell decides for the whole column: a generated
                'column has to keep its formulas across an import.
                If Not rng.Cells(1, 1).HasFormula Then rng.ClearContents
            End If
        End If
    Next
End Sub

'@label:sort
'@sub-title Sort the table on one or more columns
'@details
'colList holds the sort keys in DECREASING significance: its first item ends
'up the primary key, its last item the weakest of the list, and colName is
'the final tiebreak of them all.
'That reads backwards until you see how it works. Each pass is a full
'single-key Range.Sort over the whole table, so the LAST pass applied is the
'one that decides the order. colName is therefore sorted first and colList is
'walked from its last item to its first.
'Because the keys are applied one at a time, the precedence between them
'rests on Excel leaving rows that tie on a key in the order it found them.
'Excel behaves that way, but it does not promise to. A caller that cannot
'live with that should issue one Range.Sort carrying key1/key2/key3 instead,
'the way LLChoices does.
'@param colName Optional String. Column header sorted first, so it ends up the weakest key. Defaults to vbNullString.
'@param colList Optional Object. BetterArray of column headers, most significant first. Defaults to Nothing.
'@param directSort Optional Boolean. When True, sorts directly; when False, groups rows by first occurrence of each distinct value. Defaults to True.
'@param strictSearch Optional Boolean. When True, uses exact column matching. Defaults to True.
'@export
Public Sub Sort(Optional ByVal colName As String = vbNullString, _
                Optional ByVal colList As Object = Nothing, _
                Optional ByVal directSort As Boolean = True, _
                Optional ByVal strictSearch As Boolean = True)

    Dim counter As Long
    Dim cName As String

    'Kept here on purpose: on the early-exit path below nothing else reaches
    'the table, so this is the only guard.
    EnsureTableInitialized

    'Exit if there is no need to sort on
    If (colName = vbNullString) And (colList Is Nothing) Then Exit Sub

    'colName goes first so that it ends up the weakest key: every later pass
    'sorts the whole table again and demotes it to a tiebreak.
    SortHelper colName:=colName, directSort:=directSort, strictSearch:=strictSearch

    'Sort multiple column names (in a Betterarray).
    'Step -1 is load-bearing. The caller writes the list most significant
    'first, and the last column sorted is the one that wins, so the first item
    'has to be applied last. Without it the loop counts upward from the high
    'number, runs zero times for any list of more than one item, and every
    'extra column is dropped in silence.
    'strictSearch travels with every key. The caller asked for one matching
    'rule, and a key that cannot find its column is skipped without an error,
    'so a mismatch here shows up as a half-sorted table and nothing else.
    If (TypeName(colList) = "BetterArray") Then
        If (colList.Length > 0) Then
            For counter = colList.UpperBound To colList.LowerBound Step -1
                cName = colList.Item(counter)
                SortHelper colName:=cName, directSort:=directSort, strictSearch:=strictSearch
            Next
        End If
    End If
End Sub

'@label:sort-helper
'@sub-title Dispatch to the appropriate sort strategy
'@param colName String. Column header to sort on.
'@param directSort Optional Boolean. When True, sorts directly. Defaults to True.
'@param strictSearch Optional Boolean. When True, uses exact column matching. Defaults to True.
Private Sub SortHelper(ByVal colName As String, _
                       Optional ByVal directSort As Boolean = True, _
                       Optional ByVal strictSearch As Boolean = True)
    If (colName <> vbNullString) Then
        If directSort Then
            SortSimple colName, strictSearch
        Else
            SortOnFirst colName, strictSearch
        End If
    End If
End Sub

'@label:sort-on-first
'@sub-title Sort by grouping rows by first occurrence of each value
'@details
'Adds a temporary "__number" column, assigns ascending numbers based on
'first-seen order of each distinct value, sorts the table on that
'column, then removes it. This groups identical values together while
'preserving the original order of first appearance.
'@param colName String. Column header to sort on.
'@param strictSearch Optional Boolean. When True, uses exact column matching. Defaults to True.
Private Sub SortOnFirst(ByVal colName As String, Optional ByVal strictSearch As Boolean = True)

    Dim numRng As Range
    Dim colRng As Range
    Dim loRng As Range
    Dim hRng As Range
    Dim resRng As Range
    Dim lo As ListObject
    Dim trackData As BetterArray
    Dim sortedValue As String
    Dim sortValues As Variant
    Dim numbers() As Variant
    Dim helperColumn As Long
    Dim rowCount As Long
    Dim counter As Long
    Dim failedNumber As Long
    Dim failedSource As String
    Dim failedDescription As String

    Set lo = Table()

    'Sort on values only if there is data. Each object test stands on its
    'own: VBA does not short-circuit Or, so one combined test would read the
    'default member of a Nothing range and raise 91 - the guard written to
    'stop a crash was the thing that crashed.
    Set loRng = DataRange(includeHeaders:=False)
    If loRng Is Nothing Then Exit Sub
    Set colRng = DataRange(colName, strictSearch:=strictSearch)
    If colRng Is Nothing Then Exit Sub

    Set hRng = HeaderRange()
    'One local: this position used to be read from Excel four times.
    helperColumn = hRng.Cells.Count + 1

    'Write the helper header first, then widen the table over it.
    hRng.Cells(1, helperColumn).Value = "__number"

    'From here the sheet carries a column that is not part of the table, so
    'every way out has to take it away again. Without this the table could be
    'left permanently one column wider: the resize, the write and the sort
    'below can all raise.
    On Error GoTo RemoveHelperAndRaise

    Set resRng = Wksh().Range(hRng.Cells(1, 1), _
                              loRng.Cells(loRng.Rows.Count, loRng.Columns.Count + 1))
    lo.Resize resRng

    Set numRng = lo.ListColumns(helperColumn).DataBodyRange
    'Sometimes the data body of a brand new column comes back empty; there
    'is nothing to number then.
    If numRng Is Nothing Then
        On Error GoTo 0
        RemoveHelperColumn lo, helperColumn
        Exit Sub
    End If

    rowCount = colRng.Rows.Count
    If numRng.Rows.Count <> rowCount Then
        On Error GoTo 0
        RemoveHelperColumn lo, helperColumn
        Exit Sub
    End If

    'Read the sort column once and build the numbers in memory. Cell by cell
    'this was four round trips a row, and it runs for all nine analysis
    'tables on every setup import.
    sortValues = ColumnValues(colRng)
    ReDim numbers(1 To rowCount, 1 To 1)

    Set trackData = New BetterArray
    trackData.LowerBound = 1

    For counter = 1 To rowCount
        'CellText never raises, where CStr on an error cell did - and left
        'the previous row's value in place, grouping that row with the wrong
        'value and saying nothing.
        sortedValue = CellText(sortValues(counter))
        If Not trackData.Includes(sortedValue) Then trackData.Push sortedValue
        'The position of the first sighting IS the number to sort on, so the
        'second BetterArray the old code kept alongside held nothing new -
        'and reading it copied the whole internal array on every row.
        numbers(counter, 1) = trackData.IndexOf(sortedValue)
    Next

    numRng.Value = numbers

    'Now sort the listObject on the __number column
    resRng.Sort key1:=numRng, Order1:=xlAscending, Header:=xlYes

    On Error GoTo 0
    RemoveHelperColumn lo, helperColumn
    Exit Sub

RemoveHelperAndRaise:
    'Copy the error before cleaning up: the cleanup runs its own On Error
    'and that clears Err.
    failedNumber = Err.Number
    failedSource = Err.Source
    failedDescription = Err.Description
    RemoveHelperColumn lo, helperColumn
    Err.Raise failedNumber, failedSource, failedDescription
End Sub

'@label:remove-helper-column
'@sub-title Drop the temporary sort column whatever happened
'@details
'The helper column must go on every path out of SortOnFirst. A failure to
'remove it must not replace the error that brought us here, so the attempt
'is protected.
'@param lo ListObject. The table carrying the helper column.
'@param helperColumn Long. Position of the helper column.
Private Sub RemoveHelperColumn(ByVal lo As ListObject, ByVal helperColumn As Long)
    On Error Resume Next
        lo.ListColumns(helperColumn).Delete
    On Error GoTo 0
End Sub

'@label:sort-simple
'@sub-title Sort the table on a single column in ascending order
'@param colName String. Column header to sort on.
'@param strictSearch Optional Boolean. When True, uses exact column matching. Defaults to True.
Private Sub SortSimple(ByVal colName As String, Optional ByVal strictSearch As Boolean = True)

    Dim colRng As Range
    Dim loRng As Range

    'One object test per If: VBA does not short-circuit Or.
    Set loRng = DataRange(includeHeaders:=True)
    If loRng Is Nothing Then Exit Sub
    Set colRng = DataRange(colName, strictSearch:=strictSearch)
    If colRng Is Nothing Then Exit Sub

    loRng.Sort key1:=colRng, Order1:=xlAscending, Header:=xlYes
End Sub


'@section DataExchange
'===============================================================================
'@description Import, export, and snapshot operations.

'@label:import
'@sub-title Import data from a DataSheet or CustomTable
'@details
'Copies column-matched data from the source object into this table.
'When keepSourceHeaders is True, replaces all headers and data entirely.
'Creates a snapshot before import and rolls back on error. Unhides
'hidden columns during the operation and restores their state afterwards.
'Note on size: with keepSourceHeaders False and the default insertShift
'True, this GROWS the table to fit the source but never shrinks it. A
'three-row table importing two rows stays three rows, with the third row
'merely blanked by Clean. Only the insertShift False branch trims.
'@param impTab Object. A DataSheet or CustomTable to import from.
'@param pasteAtBottom Optional Boolean. When True, appends below existing data. Defaults to False.
'@param strictColumnSearch Optional Boolean. When True, uses exact column matching. Defaults to False.
'@param insertShift Optional Boolean. When True, inserts worksheet rows during resize. Defaults to True.
'@param formatHeaders Optional Object. BetterArray of headers whose formatting to preserve. Defaults to Nothing.
'@param keepSourceHeaders Optional Boolean. When True, replaces destination headers. Defaults to False.
'@export
Public Sub Import(ByVal impTab As Object, _
                  Optional ByVal pasteAtBottom As Boolean = False, _
                  Optional ByVal strictColumnSearch As Boolean = False, _
                  Optional ByVal insertShift As Boolean = True, _
                  Optional ByVal formatHeaders As Object = Nothing, _
                  Optional ByVal keepSourceHeaders As Boolean = False)

    Dim failedNumber As Long
    Dim failedSource As String
    Dim failedDescription As String

    On Error GoTo RollbackImport

    If ((TypeName(impTab) <> "DataSheet") And (TypeName(impTab) <> "CustomTable")) Then
        LogAndTrace "Import not done: imported elements are not tables or datasheet", checkingWarning
        Exit Sub
    End If

    ResetImportState
    'SaveTableSnapshot unhides the columns and records them, so the table is
    'already fully visible from here.
    SaveTableSnapshot

    If keepSourceHeaders Then
        Clean
        ImportAll impTab
    Else

        If Not pasteAtBottom Then Clean
        ImportObject impTab:=impTab, pasteAtBottom:=pasteAtBottom, _
                     strictColumnSearch:=strictColumnSearch, _
                     insertShift:=insertShift, _
                     formatHeaders:=formatHeaders
    End If

    this.hasImport = True
    RestoreHiddenAndReset
    AddIds
    Exit Sub

RollbackImport:
    'Copy the error as the very first thing. RestoreTableSnapshot, LogInfo
    'and the hidden-column restore each run an Exit or an On Error, and both
    'clear Err - so the re-raise below used to raise nothing at all.
    failedNumber = Err.Number
    failedSource = Err.Source
    failedDescription = Err.Description

    RestoreTableSnapshot
    LogInfo "Import aborted: " & failedDescription, checkingError
    RestoreHiddenAndReset
    Err.Raise failedNumber, failedSource, failedDescription
End Sub

'@label:import-object
'@sub-title Import column-matched data from a source object
'@details
'Iterates the source headers, finds matching columns in this table,
'and copies data column by column. Columns not found are recorded and
'read back through ImportColumnsNotFound. Optionally preserves
'formatting for specified headers. Trims the table when insertShift is
'False, and only then: on the default True branch the table grows to fit
'the source but never shrinks.
'@param impTab Object. A DataSheet or CustomTable to import from.
'@param pasteAtBottom Optional Boolean. When True, appends below existing data. Defaults to False.
'@param strictColumnSearch Optional Boolean. When True, uses exact column matching. Defaults to False.
'@param insertShift Optional Boolean. When True, inserts worksheet rows during resize. Defaults to True.
'@param formatHeaders Optional Object. BetterArray of headers whose formatting to preserve. Defaults to Nothing.
Private Sub ImportObject(ByVal impTab As Object, _
                         Optional ByVal pasteAtBottom As Boolean = False, _
                         Optional ByVal strictColumnSearch As Boolean = False, _
                         Optional ByVal insertShift As Boolean = True, _
                         Optional ByVal formatHeaders As Object = Nothing)

    Dim imphRng As Range
    Dim sourceHeaders As Variant
    Dim headerValue As String
    Dim counter As Long
    Dim impDataRng As Range
    Dim currDataRng As Range
    Dim bodyRng As Range
    Dim rng As Range
    Dim startPastingRow As Long
    Dim currColumn As ListColumn
    Dim existingRows As Long
    Dim incomingRows As Long
    Dim targetRowCount As Long

    Set imphRng = impTab.HeaderRange()
    'One block read of the source header row instead of one read per column.
    sourceHeaders = RowValues(imphRng)

    existingRows = DataRowCount()
    incomingRows = SourceRowCount(impTab)

    If pasteAtBottom Then
        startPastingRow = existingRows + 1
        targetRowCount = existingRows + incomingRows
    Else
        startPastingRow = 1
        targetRowCount = incomingRows
    End If

    targetRowCount = Application.WorksheetFunction.Max(1, targetRowCount)

    'EnsureRowCapacity carries its own "already big enough" test, so the row
    'count that used to be read again here bought nothing.
    EnsureRowCapacity targetRowCount, insertShift

    For counter = 1 To UBound(sourceHeaders)
        headerValue = CellText(sourceHeaders(counter))

        'One lookup per source column. ColumnExists then ResolveListColumn
        'searched the header row twice for the same name.
        Set currColumn = ResolveListColumn(headerValue, _
                                           strictSearch:=strictColumnSearch, _
                                           matchCase:=strictColumnSearch)

        If currColumn Is Nothing Then
            this.notImportedColumns.Push headerValue
            GoTo ContinueImport
        End If

        If (TypeName(impTab) = "CustomTable") Then
            Set impDataRng = impTab.DataRange(colName:=headerValue, strictSearch:=strictColumnSearch, matchCase:=strictColumnSearch)
        Else
            Set impDataRng = impTab.DataRange(headerValue)
        End If

        If impDataRng Is Nothing Then GoTo ContinueImport
        If impDataRng.Rows.Count = 0 Then GoTo ContinueImport

        Set currDataRng = currColumn.DataBodyRange
        If currDataRng Is Nothing Then GoTo ContinueImport

        Set rng = currDataRng.Cells(startPastingRow, 1).Resize(impDataRng.Rows.Count, 1)

        'A generated column keeps its formulas across an import.
        If Not rng.Cells(1, 1).HasFormula Then rng.Value = impDataRng.Value

        If (TypeName(formatHeaders) = "BetterArray") Then
            If (formatHeaders.Length > 0) Then
                If formatHeaders.Includes(headerValue) Then ImportFormat impRng:=impDataRng, currRng:=rng
            End If
        End If
ContinueImport:
    Next

    'An empty table has no data body. Asking it to recalculate raised 91,
    'which landed in the rollback and turned a harmless empty import into an
    'aborted one.
    Set bodyRng = DataRange(includeHeaders:=False)
    If Not bodyRng Is Nothing Then bodyRng.Calculate

    If (Not insertShift) Then TrimTableRows targetRowCount 'Trim the table to the targetRowCount
End Sub

'@label:import-format
'@sub-title Copy formatting from a source range to a destination range
'@details
'Transfers comments, comment threads, interior colour, font colour,
'bold, and italic from each cell in impRng to the corresponding cell
'in currRng. Uses error suppression for Excel versions that do not
'support CommentThreaded.
'This is deliberately NOT a straight format copy: the interior colour is
'only carried over when the source is not white and the font colour only
'when the source is not black. xlPasteFormats would be faster and would
'not do the same thing, so the routine stays a per-cell copy. What it
'stopped doing is paying for each cell more than once.
'@param impRng Range. Source range with formatting to copy.
'@param currRng Range. Destination range in this table.
Private Sub ImportFormat(ByVal impRng As Range, ByVal currRng As Range)

    Dim formatCounter As Long
    Dim rowCount As Long
    Dim sourceCell As Range
    Dim targetCell As Range
    Dim hasComment As Boolean
    Dim hasCommentThreaded As Boolean
    Dim commentText As String
    Dim commentThreadedText As String
    Dim sourceInterior As Long
    Dim sourceFontColor As Long

    'Clear all eventual comments on the range
    currRng.ClearComments

    rowCount = impRng.Rows.Count

    For formatCounter = 1 To rowCount
        'Hold both cells once. The old body rebuilt the source cell about ten
        'times a row and the target about seven.
        Set sourceCell = impRng.Cells(formatCounter, 1)
        Set targetCell = currRng.Cells(formatCounter, 1)

        'keep the comment if there is one
        hasComment = False
        hasCommentThreaded = False
        commentText = vbNullString
        commentThreadedText = vbNullString

        'On some excel version there is no CommentThread. So it would
        'result in Error.
        On Error Resume Next
            commentThreadedText = sourceCell.CommentThreaded.Text
            hasCommentThreaded = (commentThreadedText <> vbNullString)
            commentText = sourceCell.Comment.Text
            hasComment = (commentText <> vbNullString)
        On Error GoTo 0

        'Comments and comment threads both land as plain comments. No thread
        'is added here.
        If hasComment Then
            On Error Resume Next
                targetCell.AddComment
                targetCell.Comment.Text Text:=commentText
                targetCell.Comment.Visible = False
            On Error GoTo 0
        End If

        If hasCommentThreaded Then
            On Error Resume Next
                targetCell.AddComment
                targetCell.Comment.Text Text:=commentThreadedText
            On Error GoTo 0
        End If

        'Add interior color
        sourceInterior = sourceCell.Interior.Color
        If sourceInterior <> vbWhite Then targetCell.Interior.Color = sourceInterior

        'Add Font color
        sourceFontColor = sourceCell.Font.Color
        If sourceFontColor <> vbBlack Then targetCell.Font.Color = sourceFontColor

        'Add Font weight and decoration
        targetCell.Font.Bold = sourceCell.Font.Bold
        targetCell.Font.Italic = sourceCell.Font.Italic
    Next
End Sub


'@label:import-all
'@sub-title Replace the entire table with imported data
'@details
'Replaces all headers and data with the content of the supplied object,
'discarding the current structure entirely. Trims the table to one row,
'adjusts column count to match the source, then overwrites headers and
'data body.
'@param impTab Object. A DataSheet or CustomTable to import from.
Private Sub ImportAll(ByVal impTab As Object)

    Dim imphRng As Range
    Dim currhRng As Range
    Dim currLo As ListObject
    Dim sh As Worksheet
    Dim bodyRng As Range
    Dim incomingRows As Long
    Dim sourceColumns As Long
    Dim headerRow As Long
    Dim headerCol As Long

    Set currLo = Table()
    Set imphRng = impTab.HeaderRange()
    Set sh = Wksh()

    incomingRows = Application.WorksheetFunction.Max(1, SourceRowCount(impTab))
    sourceColumns = imphRng.Columns.Count

    TrimTableRows 1
    EnsureColumnCapacity sourceColumns
    Set currhRng = HeaderRange()

    headerRow = currhRng.Row
    headerCol = currhRng.Column
    Set bodyRng = SheetBlock(headerRow + 1, headerCol, _
                             headerRow + incomingRows, headerCol + sourceColumns - 1)

    currhRng.Value = imphRng.Value
    bodyRng.Value = impTab.DataRange().Value

    'Widen the block back up over the header row and resize the table on it.
    'This one is anchored on a Range, not on four numbers, so it does not go
    'through SheetBlock.
    Set bodyRng = sh.Range(bodyRng.Cells(0, 1), _
                           bodyRng.Cells(bodyRng.Rows.Count, bodyRng.Columns.Count))
    currLo.Resize bodyRng
End Sub


'@label:import-columns-not-found
'@prop-title Columns not matched during the last import
'@details
'Returns a clone of the BetterArray containing column headers from the
'source that had no match in this table. Only available after a
'successful import; logs a diagnostic and exits when no import has
'occurred.
'@return BetterArray. Unmatched column headers, or Nothing.
'@export
Public Property Get ImportColumnsNotFound() As BetterArray

    If Not this.hasImport Then
        LogAndTrace "The current table " & Name() & "has no imports"
        Exit Property
    End If

    If this.notImportedColumns Is Nothing Then Exit Property
    Set ImportColumnsNotFound = this.notImportedColumns.Clone()
End Property

'@label:has-columns-not-imported
'@prop-title Whether the last import had unmatched columns
'@details
'Returns True when the most recent import left some source columns
'unmatched against this table.
'@return Boolean. True when unmatched columns exist.
'@export
Public Property Get HasColumnsNotImported() As Boolean
    If Not this.hasImport Then Exit Property
    If this.notImportedColumns Is Nothing Then Exit Property
    HasColumnsNotImported = (this.notImportedColumns.Length >= 1)
End Property

'@label:unhide-hidden-columns
'@sub-title Unhide all hidden columns before import/export
'@details
'Import and export both write across the whole table, and a hidden column
'has to take its share, so every column is made visible first. The indexes
'of the columns that were hidden are recorded so ReturnBackHiddenColumns
'can put them back afterwards.
'The whole header band is tested first: EntireColumn.Hidden answers False
'when nothing is hidden, True when everything is and Null when it is
'mixed, so the common case costs one round trip instead of three per
'column.
Private Sub UnhideHiddenColumns()
    Dim counter As Long
    Dim hRng As Range
    Dim blockHidden As Variant

    'A list already recorded is not built a second time.
    If this.hasHiddenList Then Exit Sub

    Set this.hiddenColumnsList = New BetterArray

    Set hRng = HeaderRange()
    blockHidden = hRng.EntireColumn.Hidden

    If Not IsNull(blockHidden) Then
        If Not blockHidden Then
            this.hasHiddenList = False
            Exit Sub
        End If
    End If

    For counter = 1 To hRng.Columns.Count
        If hRng.Cells(1, counter).EntireColumn.Hidden Then
            'A protected sheet can refuse the unhide. The column then stays
            'hidden rather than stopping the whole import.
            On Error Resume Next
                hRng.Cells(1, counter).EntireColumn.Hidden = False
            On Error GoTo 0
            this.hiddenColumnsList.Push counter
        End If
    Next

    this.hasHiddenList = (this.hiddenColumnsList.Length > 0)
End Sub

'@label:return-back-hidden-columns
'@sub-title Restore previously hidden columns after import/export
'@details
'Re-hides columns that were recorded by UnhideHiddenColumns. Uses error
'suppression as a safeguard against stale column indexes.
Private Sub ReturnBackHiddenColumns()
    Dim counter As Long
    Dim hRng As Range
    Dim columnPosition As Long
    Dim lastColumn As Long

    If Not this.hasHiddenList Then Exit Sub

    Set hRng = HeaderRange()
    lastColumn = hRng.Columns.Count

    'The list holds positions, not names. An import can leave the table
    'narrower than it was, and hRng.Cells(1, oldPosition) then resolves
    'happily past the table edge and would hide a worksheet column that is
    'not ours - so the position is clamped to the table.
    'A returnBack should always come AFTER an unhide. For safeguard purposes
    'errors are ignored here.
    On Error Resume Next
        For counter = this.hiddenColumnsList.LowerBound To this.hiddenColumnsList.UpperBound
            columnPosition = CLng(this.hiddenColumnsList.Item(counter))
            If columnPosition >= 1 And columnPosition <= lastColumn Then
                hRng.Cells(1, columnPosition).EntireColumn.Hidden = True
            End If
        Next
        this.hiddenColumnsList.Clear
    On Error GoTo 0

    this.hasHiddenList = False
End Sub

'@label:restore-hidden-and-reset
'@sub-title Put the hidden columns back and drop the transient state
'@details
'The pair was written out three times, on both exits of Import and at the
'end of Export.
Private Sub RestoreHiddenAndReset()
    ReturnBackHiddenColumns
    ResetCaches
End Sub

'@label:export
'@sub-title Export the table to a worksheet
'@details
'Writes the requested columns (or all columns when headersTable is not
'a BetterArray) to the target worksheet, optionally creating a
'ListObject on the output range. Unhides hidden columns during the
'operation and restores their state afterwards, on the failure path too.
'@param sh Worksheet. Target worksheet to write to.
'@param headersTable Optional Object. BetterArray of column headers to export. Defaults to Nothing (all columns).
'@param startLine Optional Long. Row to begin writing at. Defaults to 1.
'@param startColumn Optional Long. Column to begin writing at. Defaults to 1.
'@param addListObject Optional Boolean. When True, creates a ListObject on the output. Defaults to False.
'@param clearSheet Optional Boolean. When True, clears the target worksheet first. Defaults to True.
'@export
Public Sub Export(ByVal sh As Worksheet, _
                  Optional ByVal headersTable As Object = Nothing, _
                  Optional ByVal startLine As Long = 1, _
                  Optional ByVal startColumn As Long = 1, _
                  Optional ByVal addListObject As Boolean = False, _
                  Optional ByVal clearSheet As Boolean = True)

    Dim counter As Long
    Dim inpRng As Range
    Dim outRng As Range
    Dim outLo As ListObject
    Dim targetColumn As ListColumn
    Dim headerName As String
    Dim headerCounter As Long
    Dim headersList As BetterArray
    Dim lastLine As Long
    Dim failedNumber As Long
    Dim failedSource As String
    Dim failedDescription As String

    If sh Is Nothing Then
        LogAndTrace "Unable to export CustomTable, sheet not initialized", checkingError
        Exit Sub
    End If

    Set headersList = New BetterArray

    If (TypeName(headersTable) = "BetterArray") Then
        Set headersList = headersTable.Clone()
    Else
        LogInfo "Provided headersTable for export not found, all the headers are exported"
        headersList.FromExcelRange HeaderRange(), DetectLastRow:=False, DetectLastColumn:=False
    End If

    If clearSheet Then sh.Cells.Clear

    'This will unhide all hidden columns and keep track of them
    UnhideHiddenColumns

    'From here the table is showing columns the user had hidden, so every way
    'out has to put them back.
    On Error GoTo RestoreAndRaise

    'One recalculation for the whole table. It covers every column the loop
    'below reads, where the per-column call inside the loop repeated the work
    'once per exported column.
    Table().Range.Calculate

    'column where to start writing values
    headerCounter = startColumn

    For counter = headersList.LowerBound To headersList.UpperBound

        headerName = headersList.Item(counter)

        'One lookup per column. A header with no match is skipped in silence,
        'so the "Unable to find the column name" message still does not
        'appear on this path.
        Set targetColumn = ResolveListColumn(headerName, strictSearch:=True)
        If Not targetColumn Is Nothing Then
            Set inpRng = targetColumn.Range

            With sh
                Set outRng = .Range(.Cells(startLine, headerCounter), _
                                    .Cells(inpRng.Rows.Count + startLine - 1, headerCounter))
            End With
            outRng.Value = inpRng.Value

            headerCounter = headerCounter + 1
        End If

    Next

    'Add Listobject if required on the export. The headerCounter test matters:
    'when no header matched it still sits on startColumn, and .Cells(lastLine,
    'startColumn - 1) is column 0 for the default start, which raises 1004.
    If addListObject And headerCounter > startColumn Then
        lastLine = startLine + Table().ListColumns(1).Range.Rows.Count - 1
        With sh
            Set outRng = .Range(.Cells(startLine, startColumn), .Cells(lastLine, headerCounter - 1))
            Set outLo = .ListObjects.Add(SourceType:=xlSrcRange, Source:=outRng, XlListObjectHasHeaders:=xlYes)
            outLo.Name = Name()
            outLo.TableStyle = Table().TableStyle
        End With
    End If

    On Error GoTo 0
    'Return back the previously hidden columns
    RestoreHiddenAndReset
    Exit Sub

RestoreAndRaise:
    'Copy the error before cleaning up: the restore runs its own On Error and
    'that clears Err. Without this the columns stayed visible on every raise.
    failedNumber = Err.Number
    failedSource = Err.Source
    failedDescription = Err.Description
    RestoreHiddenAndReset
    Err.Raise failedNumber, failedSource, failedDescription
End Sub

'@section Checkings
'===============================================================================
'@description Diagnostic logging and error handling.

'@label:log-info
'@sub-title Log a diagnostic message
'@details
'Adds a message to the internal Checking object, creating it on first
'use. The checkCounter provides a unique key for each entry.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level from CheckingScope. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)

    'First, instanciante the internalChecks object. The table name is only
    'needed for that, so it is read once and not on every message.
    If this.internalChecks Is Nothing Then
        Set this.internalChecks = Checking.Create("CustomTable " & Name() & " object trace")
    End If

    'Add new informations to the checking object
    this.internalChecks.Add CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:log-and-trace
'@sub-title Log a message and print the debug trace
'@details
'The "log, show the trace, then leave" tail was written out on the failure
'paths of Import, ImportColumnsNotFound and Export.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level from CheckingScope. Defaults to checkingNote.
Private Sub LogAndTrace(ByVal label As String, _
                        Optional ByVal scope As Byte = checkingNote)
    LogInfo label, scope
    ShowDebug
End Sub

'@label:has-checkings
'@prop-title Whether the table has logged diagnostic messages
'@details
'Returns True when at least one diagnostic entry has been recorded.
'@return Boolean. True when diagnostic entries exist.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@details
'Returns the internal Checking object containing all logged messages.
'Returns Nothing when no entries have been recorded.
'@return Object. An Checking instance, or Nothing.
'@export
Public Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then Exit Property

    Set CheckingValues = this.internalChecks
End Property


'@label:show-debug
'@sub-title Print a debug trace to the Immediate window
Private Sub ShowDebug()
    Debug.Print "Check the informations of custom table object " & Name()
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to "CustomTable"
'for consistent stack traces.
'@param errNumb Long. ProjectError code.
'@param errorMessage String. Descriptive message.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "CustomTable", "Error: " & errorMessage
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
