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

Implements ICustomTable

'@class CustomTable
'@description
'Wraps an Excel ListObject with convenience methods for column lookup,
'row insertion and deletion, sorting, validation, import/export, and
'diagnostic logging. The class maintains a header-index cache for fast
'repeated column lookups, supports snapshot-based rollback during import,
'and tracks hidden-column state so that Find-based operations see all
'columns. Consumers interact through the ICustomTable interface. An
'optional ID column provides dictionary-style keyed access to individual
'rows.
'@depends BetterArray, Checking, HiddenNames, IDropdownLists
'@version 1.0 (2026-02-09)
'@author Yves Amevoin

Private Type TCustomTable
  Lo As ListObject
  idCol As String
  idPrefix As String
  hasImport As Boolean
  hasChecks As Boolean
  headerIndexCache As BetterArray
  backupHeaders As BetterArray
  backupRows As BetterArray
  hasBackup As Boolean
  csName As String
  hasHiddenList As Boolean
  hasShiftSetup As Boolean
End Type

Private this As TCustomTable
Private Const CACHE_SEPARATOR As String = "|"
Private Const TAG_INSERT_SHIFT_USED As String = "InsertShiftUsed"
Private NOTIMPORTEDCOLUMNS As BetterArray
Private HIDDENCOLUMNSLIST As BetterArray

'checking objects
Private checkCounter As Long
Private internalChecks As IChecking


'@section StateManagement
'===============================================================================
'@description Cache, backup, and guard helpers that keep internal state consistent.

'@label:reset-caches
'@sub-title Reset transient caches after structural changes
'@details
'Rebuilds the header index cache so column lookups remain accurate
'after columns are added, removed, or renamed.
Private Sub ResetCaches()
  'Always rebuild the header index cache after structural changes to the table.
  Set this.headerIndexCache = New BetterArray
  this.headerIndexCache.LowerBound = 1
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
  If NOTIMPORTEDCOLUMNS Is Nothing Then Set NOTIMPORTEDCOLUMNS = New BetterArray
  NOTIMPORTEDCOLUMNS.Clear
  If HIDDENCOLUMNSLIST Is Nothing Then Set HIDDENCOLUMNSLIST = New BetterArray
  HIDDENCOLUMNSLIST.Clear
  ClearBackUp
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



'@label:cache-key-for-header
'@sub-title Generate a unique cache key for header lookups
'@details
'Combines the column name with the search mode flags so that different
'search configurations stay isolated in the cache.
'@param colName String. The column name to look up.
'@param strictSearch Boolean. Whether to use exact matching.
'@param matchCase Boolean. Whether to match case.
'@return String. The generated cache key.
Private Function CacheKeyForHeader(ByVal colName As String, _
                                   ByVal strictSearch As Boolean, _
                                   ByVal matchCase As Boolean) As String
  'Cache key combines match settings so different search modes stay isolated.
  Dim headerToken As String

  headerToken = colName
  If Not matchCase Then headerToken = LCase$(headerToken)
  CacheKeyForHeader = CStr(strictSearch) & CACHE_SEPARATOR & _
                      CStr(matchCase) & CACHE_SEPARATOR & headerToken
End Function

'@label:try-get-cached-index
'@sub-title Attempt to retrieve a cached column index
'@details
'Searches the lightweight BetterArray cache for a matching key. When
'found, populates cachedValue with the stored column index.
'@param cacheKey String. The cache key to look up.
'@param cachedValue ByRef Long. Populated with the column index on hit.
'@return Boolean. True when found in cache, False otherwise.
Private Function TryGetCachedIndex(ByVal cacheKey As String, ByRef cachedValue As Long) As Boolean
  'Search the lightweight cache implemented with BetterArray pairs key, index.
  Dim idx As Long
  Dim entry As Variant

  TryGetCachedIndex = False
  cachedValue = 0

  If this.headerIndexCache Is Nothing Then Exit Function
  If this.headerIndexCache.Length = 0 Then Exit Function

  For idx = this.headerIndexCache.LowerBound To this.headerIndexCache.UpperBound
    entry = this.headerIndexCache.Item(idx)
    If IsArray(entry) Then
      If StrComp(CStr(entry(LBound(entry))), cacheKey, vbBinaryCompare) = 0 Then
        cachedValue = CLng(entry(UBound(entry)))
        TryGetCachedIndex = True
        Exit Function
      End If
    End If
  Next idx
End Function

'@label:store-header-index
'@sub-title Store a successful column lookup in the cache
'@details
'Persists a successful lookup so subsequent queries avoid a Range.Find call.
'Ignores zero or negative column indexes.
'@param cacheKey String. The cache key to store.
'@param columnIndexValue Long. The column index to cache.
Private Sub StoreHeaderIndex(ByVal cacheKey As String, ByVal columnIndexValue As Long)

  'Persist successful lookups so subsequent queries avoid a Range.Find call.
  Dim entry As Variant

  If columnIndexValue <= 0 Then Exit Sub
  If this.headerIndexCache Is Nothing Then ResetCaches

  entry = Array(cacheKey, columnIndexValue)
  this.headerIndexCache.Push entry
End Sub

'@label:save-table-snapshot
'@sub-title Save a snapshot of the current table state
'@details
'Captures the current headers and data rows into BetterArray buffers
'so RestoreTableSnapshot can revert the table if an import fails.
'Unhides hidden columns before reading to ensure full coverage.
'@param returnHidden Optional Boolean. When True, re-hides columns after capture. Defaults to True.
'@export
Private Sub SaveTableSnapshot(Optional ByVal returnHidden As Boolean = True)
  'Capture the current table layout so we can restore it if an import fails.
  Dim lo As ListObject

  EnsureTableInitialized
  UnhideHiddenColumns

  Set lo = Table()

  Set this.backupHeaders = New BetterArray
  this.backupHeaders.LowerBound = 1
  this.backupHeaders.FromExcelRange lo.HeaderRowRange, DetectLastRow:=False, DetectLastColumn:=False

  Set this.backupRows = New BetterArray
  this.backupRows.LowerBound = 1
  If Not (lo.DataBodyRange Is Nothing) Then
    this.backupRows.FromExcelRange lo.DataBodyRange, DetectLastRow:=False, DetectLastColumn:=False
  End If

  this.hasBackup = True
  If returnHidden Then ReturnBackHiddenColumns
End Sub

'@label:restore-table-snapshot
'@sub-title Restore the table from a previously saved snapshot
'@details
'Reapplies the cached headers and rows when an import needs to roll
'back. Resizes the ListObject, writes the backup data, resets caches,
'and re-hides previously hidden columns.
'@export
Private Sub RestoreTableSnapshot()
  'Reapply the cached headers and rows when an import needs to roll back.
  Dim lo As ListObject
  Dim sheet As Worksheet
  Dim headerRow As Long
  Dim headerCol As Long
  Dim totalCols As Long
  Dim targetRows As Long
  Dim newRange As Range

  If Not this.hasBackup Then Exit Sub

  EnsureTableInitialized
  Set lo = Table()
  Set sheet = lo.Parent

  headerRow = lo.HeaderRowRange.Row
  headerCol = lo.HeaderRowRange.Column
  totalCols = this.backupHeaders.Length

  UnhideHiddenColumns

  If (this.backupRows Is Nothing) Then
    targetRows = 1
  ElseIf this.backupRows.Length = 0 Then
    targetRows = 1
  Else
    targetRows = Application.WorksheetFunction.Max(1, this.backupRows.Length)
  End If

  Set newRange = sheet.Range(sheet.Cells(headerRow, headerCol), _
                             sheet.Cells(headerRow + targetRows, headerCol + totalCols - 1))
  lo.Resize newRange

  If Not (this.backupHeaders Is Nothing) Then
    this.backupHeaders.ToExcelRange lo.HeaderRowRange.Cells(1, 1), TransposeValues:=True
  End If

  If (this.backupRows Is Nothing) Or (this.backupRows.Length = 0) Then
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
  Else
    this.backupRows.ToExcelRange lo.DataBodyRange
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
  If this.Lo Is Nothing Then
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
'@return ICustomTable. A fully initialised CustomTable instance.
'@throws ProjectError.ObjectNotInitialized When Lo is Nothing.
'@depends BetterArray, Checking
'@export
Public Function Create(ByVal Lo As ListObject, _
                      Optional ByVal idCol As String = vbNullString, _
                      Optional ByVal idPrefix As String = vbNullString) As ICustomTable
  
  If Lo Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Unable to initialize the class, no listobject provided"

  With New CustomTable
    Set .Table = Lo
    .IdValue = idCol
    .PrefixValue = idPrefix
    Set Create = .Self()
  End With
End Function

'@label:self
'@prop-title Current object instance
'@details
'Returns the current instance cast to ICustomTable. Used internally
'by the factory method to return an interface reference.
'@return ICustomTable. Reference to the current object.
Public Property Get Self() As ICustomTable
  Set Self = Me
End Property


'@section Elements
'===============================================================================
'@description Properties that expose the table structure and cell access.

'@label:wksh
'@prop-title Worksheet hosting the ListObject
'@details
'Returns the parent worksheet of the underlying ListObject.
'@return Worksheet. The host worksheet.
Private Property Get Wksh() As Worksheet
  Dim Lo As ListObject
  Set Lo = Table()
  Set Wksh = Lo.Parent
End Property

'@label:table
'@prop-title Underlying ListObject reference
'@details
'Returns the ListObject wrapped by this CustomTable. Raises an error
'when the reference has not been set.
'@return ListObject. The wrapped table.
'@throws ProjectError.ObjectNotInitialized When the ListObject is Nothing.
Public Property Get Table() As ListObject
  EnsureTableInitialized
  Set Table = this.Lo
End Property

'@label:table-set
'@prop-title Assign the ListObject reference
'@details
'Stores the ListObject, resets caches and backup, and caches the table
'name. Raises an error when Lo is Nothing.
'@param Lo ListObject. The table to assign.
'@throws ProjectError.ObjectNotInitialized When Lo is Nothing.
Public Property Set Table(ByVal Lo As ListObject)
  If Lo Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, _
               "Cannot assign an empty ListObject reference to CustomTable"
  End If
  Set this.Lo = Lo
  ResetCaches
  ClearBackup
  this.csName = Lo.Name
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
  this.idCol = idCol
End Property

'@label:prefix-value-set
'@prop-title Assign the ID prefix string
'@param idPrefix String. Prefix to prepend to sequential ID values.
Public Property Let PrefixValue(ByVal idPrefix As String)
  this.idPrefix = idPrefix
End Property

'@label:header-range
'@prop-title Header row Range of the CustomTable
'@details
'Returns the Range covering the header row of the underlying ListObject.
'@return Range. The header row range.
'@export
Private Property Get HeaderRange() As Range
  Dim Lo As ListObject
  Set Lo = Table()
  Set HeaderRange = Lo.HeaderRowRange()
End Property

'@label:column-index
'@sub-title Resolve the 1-based column index for a header name
'@details
'Searches the header row for colName using Range.Find, consulting the
'cache first for speed. Returns 0 when not found. When required is
'True, raises an error instead of returning 0.
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
  Dim foundCell As Range
  Dim cacheKey As String

  If LenB(colName) = 0 Then
    If required Then
      ThrowError ProjectError.InvalidArgument, "Column name cannot be empty"
    End If
    Exit Function
  End If

  EnsureTableInitialized
  Set hRng = HeaderRange()
  If hRng Is Nothing Then
    If required Then
      ThrowError ProjectError.ObjectNotInitialized, _
                 "Unable to read headers for table '" & Name() & "'"
    End If
    Exit Function
  End If

  cacheKey = CacheKeyForHeader(colName, strictSearch, matchCase)
  If TryGetCachedIndex(cacheKey, ColumnIndex) Then
    If ColumnIndex = 0 And required Then
      ThrowError ProjectError.ElementNotFound, _
                 "Column '" & colName & "' not found in table '" & Name() & "'"
    End If
    Exit Function
  End If

  On Error Resume Next
    Set foundCell = hRng.Find(What:=colName, _
                              LookAt:=IIf(strictSearch, xlWhole, xlPart), _
                              MatchCase:=matchCase, _
                              After:=hRng.Cells(1, hRng.Columns.Count))
  On Error GoTo 0

  If Not (foundCell Is Nothing) Then
    ColumnIndex = foundCell.Column - hRng.Column + 1
  Else
    ColumnIndex = 0
    If required Then
      ThrowError ProjectError.ElementNotFound, _
                 "Column '" & colName & "' not found in table '" & Name() & "'"
    End If
  End If

  StoreHeaderIndex cacheKey, ColumnIndex
End Function


'@label:column-exists
'@sub-title Check whether a column exists in the table
'@details
'Delegates to ColumnIndex and returns True when the index is positive.
'@param colName String. Column name to search.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to False.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
'@return Boolean. True when the column is found.
Private Function ColumnExists(ByVal colName As String, _
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

  On Error Resume Next
    Set ResolveListColumn = Table.ListColumns(columnIndexValue)
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
'@param colName Optional String. Column header to retrieve. Defaults to "__all__".
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@param strictSearch Optional Boolean. When True, uses exact whole-string matching. Defaults to False.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
'@return Range. The requested data range, or Nothing when unavailable.
'@export
Private Property Get DataRange(Optional ByVal colName As String = "__all__", _
                               Optional ByVal includeHeaders As Boolean = False, _
                               Optional ByVal strictSearch As Boolean = False, _
                               Optional ByVal matchCase As Boolean = False) As Range
  Dim lo As ListObject
  Dim targetColumn As ListColumn
  Dim rng As Range

  EnsureTableInitialized
  Set lo = Table()

  If colName = "__all__" Then
    If includeHeaders Then
      Set rng = lo.Range
    ElseIf Not (lo.DataBodyRange Is Nothing) Then
      Set rng = lo.DataBodyRange
    End If
  Else
    Set targetColumn = ResolveListColumn(colName, strictSearch, matchCase)
    If targetColumn Is Nothing Then
      LogInfo "Unable to find the column name " & colName & " of " & Name()
      Exit Property
    End If

    If includeHeaders Then
      Set rng = targetColumn.Range
    ElseIf Not (targetColumn.DataBodyRange Is Nothing) Then
      Set rng = targetColumn.DataBodyRange
    End If
  End If

  Set DataRange = rng
End Property

'@label:name
'@prop-title Name of the attached ListObject
'@details
'Returns the cached name of the ListObject assigned during creation.
'@return String. The table name.
'@export
Private Property Get Name() As String
  Name = this.csName
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
Private Property Get CellRange(ByVal colName As String, _
                               ByVal lineNum As Long) As Range
  Dim cellPos As Long
  Dim rng As Range

  If Not ColumnExists(colName) Then Exit Property
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
Private Function ValueRange(ByVal colName As String, ByVal keyName As String) As Range

  Dim idRng As Range
  Dim colRng As Range

  'Iterations for id and return ranges
  Dim cellidRng As Range
  Dim cellRetRng As Range

  If IdValue = vbNullString Then Exit Function

  If (Not ColumnExists(colName)) Or (Not ColumnExists(IdValue)) Then Exit Function
  Set idRng = DataRange(IdValue)
  Set colRng = DataRange(colName)
  If (idRng Is Nothing) Or (colRng Is Nothing) Then Exit Function
  If idRng.Rows.Count = 0 Or colRng.Rows.Count = 0 Then Exit Function

  Set cellidRng = idRng.Cells(1, 1)
  Set cellRetRng = colRng.Cells(1, 1)

  Do While (cellidRng.Value <> keyName) And _
           (cellidRng.Row <= idRng.Cells(idRng.Rows.Count, 1).Row)
    Set cellidRng = cellidRng.Offset(1)
    Set cellRetRng = cellRetRng.Offset(1)
  Loop

  If cellidRng.Row <= idRng.Cells(idRng.Rows.Count, 1).Row Then
    Set ValueRange = cellRetRng
  End If
End Function

'@label:value
'@prop-title Retrieve a cell value by column and key
'@details
'For tables with a key column, returns the cell value at the intersection
'of the specified column and the row matching keyName. Returns an empty
'string when the key or column is not found, or when the cell value
'cannot be converted to string.
'@param colName String. Column header to read from.
'@param keyName String. Key value identifying the row.
'@return String. The cell value converted to string, or vbNullString.
'@export
Private Property Get Value(ByVal colName As String, _
                           ByVal keyName As String) As String
  Dim retValue As String
  Dim cellRetRng As Range

  'Get the value Range and retrieve (try) it value
  'On error (unable to convert to string for ex) nothing is done, returned value is null string
  Set cellRetRng = ValueRange(colName, keyName)
  If Not (cellRetRng Is Nothing) Then
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

  Dim sh As Worksheet
  Dim store As IHiddenNames

  Set sh = Wksh()
  Set store = HiddenNames.Create(sh)
  If value = "Yes" Then this.hasShiftSetup = True

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

  Dim store As IHiddenNames
  
  Set store = HiddenNames.Create(Wksh())

  If Not this.hasShiftSetup Then
    this.hasShiftSetup = (store.ValueAsString(TAG_INSERT_SHIFT_USED, "No") = "Yes")
  End If

  HasShiftTracker = this.hasShiftSetup

End Function

'@label:resize
'@sub-title Resize the table by adding or removing rows
'@details
'Dispatches to AddTableRows or RemoveEmptyDataRows depending on the
'addRows flag, then optionally renumbers the ID column.
'@param addRows Optional Boolean. When True, adds rows; when False, removes empty rows. Defaults to False.
'@param insertShift Optional Boolean. When True, inserts worksheet rows. Defaults to True.
'@param totalRowCount Optional Long. Non-empty cell threshold for row removal. Defaults to 0.
'@param nbRows Optional Long. Number of rows to add. Defaults to 5.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion. Defaults to False.
Private Sub Resize(Optional ByVal addRows As Boolean = False, _
                   Optional ByVal insertShift As Boolean = True, _
                   Optional ByVal totalRowCount As Long = 0, _
                   Optional ByVal nbRows As Long = 5, _
                   Optional ByVal includeIds As Boolean = True, _ 
                   Optional ByVal forceShift As Boolean = False)
  Dim lo As ListObject

  EnsureTableInitialized
  Set lo = Table()

  If addRows Then
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
'directly. Resets caches after the operation.
'@param lo ListObject. The table to extend.
'@param nbRows Long. Number of rows to add.
'@param insertShift Boolean. When True, inserts worksheet rows.
Private Sub AddTableRows(ByVal lo As ListObject, ByVal nbRows As Long, ByVal insertShift As Boolean)
  'ListRows.Add preserves formulas and formats; AlwaysInsert prevents overlaps when tables stack.
  Dim sheet As Worksheet
  Dim headerRow As Long
  Dim headerCol As Long
  Dim totalCols As Long
  Dim totalRows As Long
  Dim insertStart As Long
  Dim newRange As Range

  If nbRows <= 0 Then Exit Sub

  Set sheet = lo.Parent
  headerRow = lo.Range.Row
  headerCol = lo.Range.Column
  totalCols = lo.Range.Columns.Count
  totalRows = lo.Range.Rows.Count

  If (insertShift) Then
    'Insert a block of rows below the current table to avoid overlapping adjacent tables.
    insertStart = headerRow + totalRows
    sheet.Rows(insertStart).Resize(nbRows).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newRange = sheet.Range(sheet.Cells(headerRow, headerCol), _
                               sheet.Cells(insertStart + nbRows - 1, headerCol + totalCols - 1))
    AddShiftTracker "Yes"
    Application.ScreenUpdating = False                         
    lo.Resize newRange
  Else
    'Extend the listobject in a single resize call when overlapping is not a concern.
    Set newRange = sheet.Range(sheet.Cells(headerRow, headerCol), _
                               sheet.Cells(headerRow + totalRows + nbRows - 1, headerCol + totalCols - 1))
    Application.ScreenUpdating = False
    lo.Resize newRange
  End If

  ResetCaches
End Sub

'@label:ensure-selection-on-table-sheet
'@sub-title Ensure the provided range belongs to the table worksheet
'@details
'Validates that targetCell is not Nothing and resides on the same
'worksheet as the wrapped ListObject.
'@param targetCell Range. The range to validate.
'@throws ProjectError.InvalidArgument When targetCell is Nothing, has no worksheet, or belongs to a different sheet.
Private Sub EnsureSelectionOnTableSheet(ByVal targetCell As Range)
  Dim targetSheet As Worksheet

  If targetCell Is Nothing Then
    ThrowError ProjectError.InvalidArgument, _
               "Target range cannot be empty when modifying table '" & Name() & "'"
  End If

  On Error Resume Next
    Set targetSheet = targetCell.Worksheet
  On Error GoTo 0

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
  Dim rowRange As Range
  Dim intersectRange As Range
  Dim rowsList As BetterArray

  EnsureTableInitialized
  Set lo = Table()

  If targetCell Is Nothing Then Exit Function

  If restrictToDataRows Then
    Set scopeRange = lo.DataBodyRange
  Else
    Set scopeRange = lo.Range
    If scopeRange Is Nothing Then
      Set scopeRange = lo.HeaderRowRange
    End If
  End If

  If scopeRange Is Nothing Then Exit Function
  If Intersect(targetCell, scopeRange) Is Nothing Then Exit Function

  Set rowsList = New BetterArray
  rowsList.LowerBound = 1

  For Each rowRange In targetCell.Rows
    Set intersectRange = Intersect(rowRange, scopeRange)
    If Not intersectRange Is Nothing Then rowsList.Push intersectRange.Row
  Next rowRange

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
Private Sub InsertRowsAt(ByVal targetCell As Range, _
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

  EnsureTableInitialized
  EnsureSelectionOnTableSheet targetCell
  Set lo = Table()
  Set sheet = Wksh()

  Set rowNumbers = SelectionRowNumbers(targetCell, restrictToDataRows:=False)
  If rowNumbers Is Nothing Then Exit Sub
  If rowNumbers.Length = 0 Then Exit Sub

  rowCount = rowNumbers.Length
  headerRow = lo.HeaderRowRange.Row

  If lo.DataBodyRange Is Nothing Then
    AddTableRows lo, rowCount, insertShift
  ElseIf insertShift Then
    AddShiftTracker "Yes"
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
'@param rowCount ByRef Long. Populated with the formula column count.
Private Sub GetTotalFormulaColumns(ByRef rowCount As Long)
  Dim rng As Range
  Dim counter As Long

  rowCount = 0
  Set rng = Table.DataBodyRange()
  If rng Is Nothing Then Exit Sub
  For counter = 1 To rng.Columns.Count
    If rng.Cells(1, counter).HasFormula Then
      rowCount = rowCount + 1
    End If
  Next

End Sub


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
  Dim targetRow As Range
  Dim tabRowCount As Long
  Dim rowCount As Long
  Dim sh As Worksheet
  Dim hasTracker As Boolean

  If totalRowCount = 0 Then
    GetTotalFormulaColumns(rowCount)
  Else
    rowCount = totalRowCount
  End If

  hasTracker = HasShiftTracker()

  If lo.DataBodyRange Is Nothing Then Exit Sub

  tabRowCount = lo.DataBodyRange.Rows.Count
  If tabRowCount <= 1 Then Exit Sub   'Keep the first data row to preserve formulas/templates.

  For rowIndex = tabRowCount To 2 Step -1
    Set targetRow = lo.DataBodyRange.Rows(rowIndex)
    If WorksheetFunction.CountA(targetRow) <= rowCount Then
      If hasTracker Or forceShift Then
        Set sh = Wksh()
        'Delete the entire worksheet Row, if it does not work, force listObject deletion.
        'in case everything fail
        On Error Resume Next
        sh.Rows(lo.ListRows(rowIndex).Range().Row).Delete
        Application.ScreenUpdating = False
        lo.ListRows(rowIndex).Delete
        Application.ScreenUpdating = False
        If Err.Number <> 0 Then
          Err.Clear
          ThrowError ProjectError.SomethingWentWrong, "Unable to resize the table (maybe tables are stacked?)"
        End If
        On Error GoTo 0
      Else
        Application.ScreenUpdating = False
        'Either delele the entire worksheet row or the the listRow
        lo.ListRows(rowIndex).Delete
        Application.ScreenUpdating = False
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
Private Sub DeleteRowsAt(ByVal targetCell As Range, _
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

  EnsureTableInitialized
  EnsureSelectionOnTableSheet targetCell
  Set lo = Table()

  If lo.DataBodyRange Is Nothing Then Exit Sub

  Set rowNumbers = SelectionRowNumbers(targetCell, restrictToDataRows:=True)
  If rowNumbers Is Nothing Then Exit Sub
  If rowNumbers.Length = 0 Then Exit Sub

  shouldShift = forceShift Or HasShiftTracker()
  Set sheet = Wksh()
  headerRow = lo.HeaderRowRange.Row
  pendingRowCount = lo.ListRows.Count

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

  EnsureTableInitialized
  Set body = Table().DataBodyRange
  If body Is Nothing Then Exit Function

  DataRowCount = body.Rows.Count
End Function

'@label:data-column-count
'@sub-title Count the number of columns in the table
'@return Long. The column count.
Private Function DataColumnCount() As Long
  EnsureTableInitialized
  DataColumnCount = Table.ListColumns.Count
End Function

'@label:ensure-row-capacity
'@sub-title Guarantee the table has at least the requested data rows
'@param requiredRows Long. Minimum number of data rows needed.
'@param insertShift Boolean. When True, inserts worksheet rows.
Private Sub EnsureRowCapacity(ByVal requiredRows As Long, ByVal insertShift As Boolean)
  'Guarantee the table contains at least the requested number of data rows.
  Dim currentRows As Long
  Dim missingRows As Long
  Dim trimedRows As Long

  trimedRows = Application.WorksheetFunction.Max(1, requiredRows)
  currentRows = Application.WorksheetFunction.Max(0, DataRowCount())

  If currentRows >= trimedRows Then Exit Sub

  missingRows = trimedRows - currentRows
  AddTableRows Table(), missingRows, insertShift
End Sub

'@label:ensure-column-capacity
'@sub-title Guarantee the table has exactly the requested column count
'@details
'Adds or removes columns until the table matches requiredColumns. Trims
'excess columns using a descending loop so index shifts do not skip columns.
'@param requiredColumns Long. Target number of columns.
Private Sub EnsureColumnCapacity(ByVal requiredColumns As Long)

  Dim currCols As Long
  Dim Lo As ListObject
  Dim trimedColumns As Long

  EnsureTableInitialized

  trimedColumns = Application.WorksheetFunction.Max(0, requiredColumns)
  currCols = Application.WorksheetFunction.Max(1, DataColumnCount())

  Set Lo = Table()

  'Add the missing columns before copying incoming headers so we avoid
  'Excel-provided defaults like "Column1".
  Do While currCols < trimedColumns
    Lo.ListColumns.Add
    currCols = currCols + 1
  Loop

  'Trim excess columns using a descending loop so index shifts do not skip columns.
  Do While currCols > trimedColumns
    On Error Resume Next
      Lo.ListColumns(currCols).DataBodyRange.Cells.Clear
    On Error GoTo 0
    Lo.ListColumns(currCols).Delete
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
  Dim sheet As Worksheet
  Dim headerRow As Long
  Dim headerCol As Long
  Dim totalCols As Long
  Dim newRange As Range
  Dim normalizedTarget As Long

  EnsureTableInitialized
  Set lo = Table()
  Set sheet = lo.Parent

  normalizedTarget = Application.WorksheetFunction.Max(1, targetRows)
  If DataRowCount() <= normalizedTarget Then Exit Sub

  headerRow = lo.HeaderRowRange.Row
  headerCol = lo.HeaderRowRange.Column
  totalCols = lo.HeaderRowRange.Columns.Count

  Set newRange = sheet.Range(sheet.Cells(headerRow, headerCol), _
                             sheet.Cells(headerRow + normalizedTarget, headerCol + totalCols - 1))
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

  SourceRowCount = 0

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
Private Sub AddIds()
  Dim rng As Range
  Dim counter As Long
  Dim idPref As String
  Dim ID As String
  Dim nextIndex As Long
  Dim prefixText As String
  Dim idColumn As ListColumn

  ID = IdValue()
  If (ID = vbNullString) Then Exit Sub
  If Not (ColumnExists(ID)) Then Exit Sub
  idPref = PrefixValue()
  Set idColumn = ResolveListColumn(ID, strictSearch:=True)
  If idColumn Is Nothing Then Exit Sub
  If idColumn.DataBodyRange Is Nothing Then Exit Sub
  Set rng = idColumn.DataBodyRange

  nextIndex = 1
  prefixText = Trim$(idPref)

  For counter = 1 To rng.Rows.Count
    rng.Cells(counter, 1).Value = FormatIdValue(prefixText, nextIndex)
    nextIndex = nextIndex + 1
  Next
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
Private Sub AddRows(Optional ByVal nbRows As Long = 5, _ 
                    Optional Byval insertShift As Boolean = False, _ 
                    Optional ByVal includeIds As Boolean = True)
  
  Resize AddRows:=True, nbRows:=nbRows, insertShift:=insertShift, _ 
         includeIds:=includeIds
End Sub


'@label:remove-rows
'@sub-title Remove empty rows from the CustomTable
'@details
'Deletes rows whose non-empty cell count falls at or below totalCount.
'In the linelist, formula cells count as non-empty, so totalCount
'should match the number of formula columns. Optionally renumbers the
'ID column afterwards.
'@param totalCount Optional Long. Maximum non-empty cells for a row to be considered empty. Defaults to 0.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion. Defaults to False.
'@export
Private Sub RemoveRows(Optional ByVal totalCount As Long = 0, _ 
                       Optional ByVal includeIds As Boolean = True, _
                       Optional ByVal forceShift As Boolean = False)

  Resize AddRows:=False, totalRowCount:=totalCount, includeIds:=includeIds, forceShift:=forceShift
End Sub

'@label:set-validation
'@sub-title Set dropdown validation on a column
'@details
'Applies data validation from the supplied IDropdownLists to every cell
'in the specified column. Silently exits when the column does not exist
'or has no data range.
'@param colName String. Column header to validate.
'@param drop IDropdownLists. Dropdown source providing the validation list.
'@param dropName String. Name of the dropdown list within the source.
'@param alertType Optional String. Alert severity ("info", "warning", "error"). Defaults to "info".
'@param message Optional String. Message shown when validation fails. Defaults to vbNullString.
'@export
Private Sub SetValidation(ByVal colName As String, _
                          ByVal drop As IDropdownLists, _
                          ByVal dropName As String, _
                          Optional ByVal alertType As String = "info", _
                          Optional ByVal message As String = vbNullString)

    Dim rng As Range

    'Test if the colname exists before seting the validation
    If Not ColumnExists(colName) Then Exit Sub
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
Private Sub SetValue(ByVal colName As String, ByVal keyName As String, ByVal newValue As String)

  Dim cellRetRng As Range

  'Get the value Range and assign (try) it the newValue
  'On error (cellRetRng is Nothing for ex, nothing is done)
  'also if for some reasons it is not possible to change the values of the cell,
  'nothing is done
  Set cellRetRng = ValueRange(colName, keyName)
  If cellRetRng Is Nothing Then Exit Sub
  cellRetRng.Value = newValue
End Sub

'@label:rename-column
'@sub-title Rename a column header within the table
'@details
'Replaces the header text of the matched column with newName and resets
'the lookup cache. Raises an error when newName is empty.
'@param colName String. Current column header to find.
'@param newName String. New header text to assign.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
'@throws ProjectError.InvalidArgument When newName is empty.
'@export
Private Sub RenameColumn(ByVal colName As String, _
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
'that do not contain formulas. Key columns are preserved since they
'are identified by their header, not by content.
'@export
Private Sub Clean()
  Dim rng As Range
  Dim hRng As Range
  Dim headerValue As String
  Dim counter As Long

  Set hRng = HeaderRange()

  'Clean only columns without formulas that are not keys (key columns are kept as it).
  For counter = 1 To hRng.Columns.Count
    headerValue = hRng.Cells(1, counter).Value

      Set rng = DataRange(colName:=headerValue, strictSearch:=True, includeHeaders:=False)
      If Not (rng Is Nothing) Then
        'On formula columns, or when the column is empty, do not clear formulas.
        If Not rng.Cells(1, 1).HasFormula Then rng.ClearContents
      End If
  Next
End Sub

'@label:sort
'@sub-title Sort the table on one or more columns
'@details
'Sorts in ascending order on the specified column, then on each column
'in colList. When directSort is True, uses a simple Range.Sort; when
'False, groups rows by first occurrence of each distinct value.
'@param colName Optional String. Column header to sort on. Defaults to vbNullString.
'@param colList Optional Object. BetterArray of additional column headers. Defaults to Nothing.
'@param directSort Optional Boolean. When True, sorts directly. Defaults to True.
'@param strictSearch Optional Boolean. When True, uses exact column matching. Defaults to True.
'@export
Private Sub Sort(Optional ByVal colName As String = vbNullString, _ 
                 Optional ByVal colList As Object = Nothing, _
                 Optional ByVal directSort As Boolean = True, _ 
                 Optional ByVal strictSearch As Boolean = True)
  
  Dim counter As Long
  Dim cName As String

  EnsureTableInitialized

  'Exit if there is no need to sort on
  If (colName = vbNullString) And (colList Is Nothing) Then Exit Sub

  'Sort only one colName
  SortHelper colName:=colName, directSort:=directSort, strictSearch:=strictSearch

  'Sort multiple column names (in a Betterarray)
  If (TypeName(colList) = "BetterArray") Then
    If (colList.Length > 0) Then
      For counter = colList.UpperBound To colList.LowerBound
        cName = colList.Item(counter)
        SortHelper colName:=cName, directSort:=directSort
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
  Dim LoRng As Range 
  Dim hRng As Range
  Dim counter As Long
  Dim trackData As BetterArray
  Dim sh As Worksheet
  Dim sortedValue As String
  Dim num As Long
  Dim numData As BetterArray
  Dim resRng As Range
  Dim Lo As ListObject

  'Sort on values only If there is data
  Set Lo = Table()
  Set LoRng = DataRange(includeHeaders:=False)
  Set colRng = DataRange(colName, strictSearch:=strictSearch)

  'Exit if there is nothing
  If (LoRng Is Nothing) Or IsEmpty(LoRng) Or (colRng Is Nothing) Or _
     (IsEmpty(colRng)) Then Exit Sub

  Set trackData = New BetterArray
  Set numData = New BetterArray
  Set sh = Wksh()
  Set hRng = HeaderRange()

  'Test values progressively
  'If the trackData does not contains the values,
  'Update the number (where you will sort)
  hRng.Cells(1, hRng.Cells.Count + 1).Value = "__number"
  'Resize the listObject to be sure it includes the number added
  With sh
    Set resRng = .Range(hRng.Cells(1, 1), _ 
                        LoRng.Cells(LoRng.Rows.Count, LoRng.Columns.Count + 1))
    Lo.Resize resRng
  End With

  Set numRng = Lo.ListColumns(hRng.Cells.Count + 1).DataBodyRange
  'Seems like sometimes the databodyrange of a new column
  'could be empty in that case exit
  If numRng Is Nothing Then
    Lo.ListColumns(hRng.Cells.Count + 1).Delete
    Exit Sub
  End If
  num = 1 'Number to sort later on

  For counter = 1 To colRng.Cells.Count
     On Error Resume Next
     sortedValue = Cstr(colRng.Cells(counter, 1).Value)
     On Error GoTo 0
    If Not trackData.Includes(sortedValue) Then
      trackData.Push sortedValue
      numData.Push num
      numRng.Cells(counter, 1).Value = num
      num = num + 1
    Else
      numRng.Cells(counter, 1) = numData.Items(trackData.IndexOf(sortedValue))
    End If
  Next
  'Now sort the listObject on the __number column
  resRng.Sort key1:=numRng, Order1:=xlAscending, Header:=xlYes
  'Remove the __number column
  Lo.ListColumns(hRng.Cells.Count + 1).Delete
End Sub

'@label:sort-simple
'@sub-title Sort the table on a single column in ascending order
'@param colName String. Column header to sort on.
'@param strictSearch Optional Boolean. When True, uses exact column matching. Defaults to True.
Private Sub SortSimple(ByVal colName As String, Optional ByVal strictSearch As Boolean = True)

  Dim colRng As Range
  Dim LoRng As Range

  Set LoRng = DataRange(includeHeaders:=True)
  Set colRng = DataRange(colName, strictSearch:=strictSearch)

  If (LoRng Is Nothing) Or IsEmpty(LoRng) Or (colRng Is Nothing) Or _
      (IsEmpty(colRng)) Then Exit Sub

  LoRng.Sort key1:=colRng, Order1:=xlAscending, Header:=xlYes
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
'@param impTab Object. A DataSheet or CustomTable to import from.
'@param pasteAtBottom Optional Boolean. When True, appends below existing data. Defaults to False.
'@param strictColumnSearch Optional Boolean. When True, uses exact column matching. Defaults to False.
'@param insertShift Optional Boolean. When True, inserts worksheet rows during resize. Defaults to True.
'@param formatHeaders Optional Object. BetterArray of headers whose formatting to preserve. Defaults to Nothing.
'@param keepSourceHeaders Optional Boolean. When True, replaces destination headers. Defaults to False.
'@export
Private Sub Import(ByVal impTab As Object, _
                   Optional ByVal pasteAtBottom As Boolean = False, _
                   Optional ByVal strictColumnSearch As Boolean = False, _ 
                   Optional ByVal insertShift As Boolean = True, _ 
                   Optional ByVal formatHeaders As Object = Nothing, _ 
                   Optional ByVal keepSourceHeaders As Boolean = False)

  On Error GoTo RollbackImport
  
  If ((TypeName(impTab) <> "DataSheet") And (TypeName(impTab) <> "CustomTable")) Then
    LogInfo "Import not done: imported elements are not tables or datasheet", checkingWarning
    ShowDebug
    Exit Sub
  End If

  ResetImportState
  SaveTableSnapshot returnHidden:=False
  UnhideHiddenColumns

  If keepSourceHeaders Then
    Clean
    ImportAll impTab, insertShift:=insertShift
  Else

    If Not pasteAtBottom Then Clean
    ImportObject impTab:=impTab, pasteAtBottom:=pasteAtBottom, _ 
                 strictColumnSearch:=strictColumnSearch, _
                 insertShift:=insertShift, _
                 formatHeaders:=formatHeaders
  End If

  EnsureTableInitialized
  this.hasImport = True
  ReturnBackHiddenColumns
  ResetCaches
  AddIds
  Exit Sub

RollbackImport:
  RestoreTableSnapshot
  LogInfo "Import aborted: " & Err.Description, checkingError
  ReturnBackHiddenColumns
  ResetCaches
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'@label:import-object
'@sub-title Import column-matched data from a source object
'@details
'Iterates the source headers, finds matching columns in this table,
'and copies data cell by cell. Columns not found are tracked in
'NOTIMPORTEDCOLUMNS. Optionally preserves formatting for specified
'headers. Trims the table when insertShift is False.
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
  Dim headerValue As String
  Dim counter As Long
  Dim impDataRng As Range
  Dim currDataRng As Range
  Dim rng As Range
  Dim startPastingRow As Long
  Dim currColumn As ListColumn
  Dim existingRows As Long
  Dim incomingRows As Long
  Dim targetRowCount As Long


  Set imphRng = impTab.HeaderRange()
  existingRows = DataRowCount()
  incomingRows = SourceRowCount(impTab)

  If pasteAtBottom Then
    startPastingRow = existingRows + 1
    targetRowCount = existingRows + incomingRows
  Else
    startPastingRow = 1
    targetRowCount = incomingRows
  End If

  '@Ignore AssignmentNotUsed
  targetRowCount = Application.WorksheetFunction.Max(1, targetRowCount)

  If DataRowCount() < targetRowCount Then
    EnsureRowCapacity targetRowCount, insertShift
  End If

  For counter = 1 To imphRng.Columns.Count
    headerValue = imphRng.Cells(1, counter).Value

    If ColumnExists(headerValue, strictSearch:=strictColumnSearch, matchCase:=strictColumnSearch) Then
      If (TypeName(impTab) = "CustomTable") Then
        Set impDataRng = impTab.DataRange(colName:=headerValue, strictSearch:=strictColumnSearch, matchCase:=strictColumnSearch)
      Else
        Set impDataRng = impTab.DataRange(headerValue)
      End If

      If impDataRng Is Nothing Then GoTo ContinueImport
      If impDataRng.Rows.Count = 0 Then GoTo ContinueImport

      Set currColumn = ResolveListColumn(headerValue, strictSearch:=strictColumnSearch, matchCase:=strictColumnSearch, required:=True)
      If currColumn Is Nothing Then GoTo ContinueImport

      Set currDataRng = currColumn.DataBodyRange
      If currDataRng Is Nothing Then GoTo ContinueImport

      Set rng = currDataRng.Cells(startPastingRow, 1).Resize(impDataRng.Rows.Count, 1)

      If Not rng.Cells(1, 1).HasFormula Then rng.Value = impDataRng.Value

      If (TypeName(formatHeaders) = "BetterArray") Then
        If (formatHeaders.Length > 0) Then
          If formatHeaders.Includes(headerValue) Then ImportFormat impRng:=impDataRng, currRng:=rng
        End If
      End If
    Else
      NOTIMPORTEDCOLUMNS.Push headerValue
    End If
ContinueImport:
  Next
  DataRange.Calculate
  If (Not InsertShift) Then TrimTableRows targetRowCount 'Trim the table to the targetRowCount
End Sub

'@label:import-format
'@sub-title Copy formatting from a source range to a destination range
'@details
'Transfers comments, comment threads, interior colour, font colour,
'bold, and italic from each cell in impRng to the corresponding cell
'in currRng. Uses error suppression for Excel versions that do not
'support CommentThreaded.
'@param impRng Range. Source range with formatting to copy.
'@param currRng Range. Destination range in this table.
Private Sub ImportFormat(ByVal impRng As Range, ByVal currRng As Range)

  Dim formatCounter As Long
  Dim hasComment As Boolean
  Dim hasCommentThreaded As Boolean
  Dim commentText As String
  Dim commentThreadedText As String
      
  'Clear all eventual comments on the range
  currRng.ClearComments

  For formatCounter = 1 To impRng.Rows.Count
    'keep the comment if there is one
    hasComment = False
    hasCommentThreaded = False
        
    'On some excel version there is no CommentThread. So it would
    'result in Error.
    On Error Resume Next
      hasCommentThreaded = (impRng.Cells(formatCounter, 1).CommentThreaded.Text <> vbNullString)
      hasComment = (impRng.Cells(formatCounter, 1).Comment.Text <> vbNullString)
      commentThreadedText = impRng.Cells(formatCounter, 1).CommentThreaded.Text
      commentText = impRng.Cells(formatCounter, 1).Comment.Text
    On Error GoTo 0

    'Add comment and commentThread are converted to comments. No threads is 
    'added here

    If hasComment Then
      On Error Resume Next
      currRng.Cells(formatCounter, 1).AddComment
      currRng.Cells(formatCounter, 1).Comment.Text Text:=commentText
      currRng.Cells(formatCounter, 1).Comment.Visible = False
      On Error GoTo 0
    End If

    If hasCommentThreaded Then
      On Error Resume Next
        currRng.Cells(formatCounter, 1).AddComment
        currRng.Cells(formatCounter, 1).Comment.Text Text:=commentThreadedText
      On Error GoTo 0
    End If
        'Add interiror color
    If (impRng.Cells(formatCounter, 1).Interior.color <> vbWhite) Then 
      currRng.Cells(formatCounter, 1).Interior.color = impRng.Cells(formatCounter, 1).Interior.color
    End If
        
    'Add Font color
    If (impRng.Cells(formatCounter, 1).Font.Color <> vbBlack) Then
      currRng.Cells(formatCounter, 1).Font.color = impRng.Cells(formatCounter, 1).Font.color
    End If
    'Add Font weight and decoration
    currRng.Cells(formatCounter, 1).Font.Bold = impRng.Cells(formatCounter, 1).Font.Bold
    currRng.Cells(formatCounter, 1).Font.Italic = impRng.Cells(formatCounter, 1).Font.Italic
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
'@param insertShift Optional Boolean. When True, inserts worksheet rows. Defaults to False.
'@export
Private Sub ImportAll(ByVal impTab As Object, Optional ByVal insertShift As Boolean = False)

  Dim imphRng As Range
  Dim currhRng As Range
  Dim currLo As listObject
  Dim sh As Worksheet
  Dim bodyRng As Range
  Dim incomingRows As Long
  Dim headerRow As Long
  Dim headerCol As Long

  Set currLo = Table()
  Set imphRng = impTab.HeaderRange()

  Set sh = Wksh()

  incomingRows = SourceRowCount(impTab)
  incomingRows = Application.WorksheetFunction.Max(1, incomingRows)

  TrimTableRows 1
  EnsureColumnCapacity imphRng.Columns.count
  Set currhRng = HeaderRange()

  headerRow = currhRng.Row
  headerCol = currhRng.Column
  Set bodyRng = sh.Range(sh.Cells(headerRow + 1, headerCol), _
                           sh.Cells(headerRow + incomingRows, headerCol + imphRng.Columns.count - 1))
  
  currhRng.Value = imphRng.Value
  bodyRng.Value = impTab.DataRange().Value

  Set bodyRng = sh.Range(bodyRng.Cells(0, 1), bodyRng.Cells(bodyRng.Rows.Count, bodyRng.Columns.Count))
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
Private Property Get ImportColumnsNotFound() As BetterArray

  If Not this.hasImport Then
    LogInfo "The current table " & Name() & "has no imports"
    ShowDebug
    Exit Property
  End If

  If NOTIMPORTEDCOLUMNS Is Nothing Then Exit Property
  Set ImportColumnsNotFound = NOTIMPORTEDCOLUMNS.Clone()
End Property

'@label:has-columns-not-imported
'@prop-title Whether the last import had unmatched columns
'@details
'Returns True when the most recent import left some source columns
'unmatched against this table.
'@return Boolean. True when unmatched columns exist.
'@export
Private Property Get HasColumnsNotImported() As Boolean
  If Not this.hasImport Then Exit Property
  If NOTIMPORTEDCOLUMNS Is Nothing Then Exit Property
  HasColumnsNotImported = (NOTIMPORTEDCOLUMNS.Length >= 1)
End Property

'@label:unhide-hidden-columns
'@sub-title Unhide all hidden columns before import/export
'@details
'Range.Find does not include hidden columns, so all columns must be
'made visible before import or export operations. Records the indexes
'of previously hidden columns so ReturnBackHiddenColumns can restore
'them afterwards.
Private Sub UnhideHiddenColumns()
  Dim counter As Long
  Dim hRng As Range

  Set hRng = HeaderRange()
  
  If Not this.hasHiddenList Then 
    Set HIDDENCOLUMNSLIST = New BetterArray
    HIDDENCOLUMNSLIST.Clear
    For counter = 1 To hRng.Columns.Count
      If hRng.Cells(1, counter).EntireColumn.HIDDEN Then
        On Error Resume Next
        hRng.Cells(1, counter).EntireColumn.HIDDEN = False
        On Error GoTo 0
        HIDDENCOLUMNSLIST.Push counter
      End If
    Next
  End If

  this.hasHiddenList = (HIDDENCOLUMNSLIST.Length > 0)
End Sub

'@label:return-back-hidden-columns
'@sub-title Restore previously hidden columns after import/export
'@details
'Re-hides columns that were recorded by UnhideHiddenColumns. Uses error
'suppression as a safeguard against stale column indexes.
Private Sub ReturnBackHiddenColumns()
  Dim counter As Long
  Dim hRng As Range

  If Not this.hasHiddenList Then Exit Sub

  Set hRng = HeaderRange()
  'We should keep in mind that a returnBack shoud be AFTER a unhide.
  'For safeguard purposes, I can work toward ignoring errors (variable not set for ex.)
  On Error Resume Next
    For counter = HIDDENCOLUMNSLIST.LowerBound To HIDDENCOLUMNSLIST.UpperBound
      hRng.Cells(1, HIDDENCOLUMNSLIST.Item(counter)).EntireColumn.HIDDEN = True
    Next
    HIDDENCOLUMNSLIST.Clear
  On Error GoTo 0

  this.hasHiddenList = False
End Sub

'@label:export
'@sub-title Export the table to a worksheet
'@details
'Writes the requested columns (or all columns when headersTable is not
'a BetterArray) to the target worksheet, optionally creating a
'ListObject on the output range. Unhides hidden columns during the
'operation and restores their state afterwards.
'@param sh Worksheet. Target worksheet to write to.
'@param headersTable Optional Object. BetterArray of column headers to export. Defaults to Nothing (all columns).
'@param startLine Optional Long. Row to begin writing at. Defaults to 1.
'@param startColumn Optional Long. Column to begin writing at. Defaults to 1.
'@param addListObject Optional Boolean. When True, creates a ListObject on the output. Defaults to False.
'@export
Private Sub Export(ByVal sh As Worksheet, _ 
           Optional ByVal headersTable As Object = Nothing, _ 
           Optional ByVal startLine As Long = 1, _ 
           Optional ByVal startColumn As Long = 1, _ 
           Optional ByVal addListObject As Boolean = False)

  Dim counter As Long
  Dim inpRng As Range
  Dim outrng As Range
  Dim headerName As String
  Dim headerCounter As Long
  Dim headersList As BetterArray
  Dim lastLine As Long

  If sh Is Nothing Then 
    LogInfo "Unable to export CustomTable, sheet not initialized", checkingError
    Exit Sub
  End If

  Set headersList = New BetterArray

  If (TypeName(headersTable) = "BetterArray") Then
    Set headersList = headersTable.Clone()
  Else
    LogInfo "Provided headersTable for export not found, all the headers are exported"
    headersList.FromExcelRange HeaderRange(), DetectLastRow:=False, DetectLastColumn:=False
  End If

  sh.Cells.Clear

  'column where to start writing values
  headerCounter = startColumn
  lastLine = startLine + Table.ListColumns(1).Range.Rows.Count - 1

  'This will unhide all hidden columns and keep track of them
  UnhideHiddenColumns

  For counter = headersList.LowerBound To headersList.UpperBound

    headerName = headersList.Item(counter)
    
    If ColumnExists(headerName, strictSearch:=True) Then
      Set inpRng = DataRange(headerName, strictSearch:=True, includeHeaders:=True)
      inpRng.Calculate
      
      With sh
        Set outrng = .Range(.Cells(startLine, headerCounter), .Cells(inpRng.Rows.Count + startLine - 1, headerCounter))
        outrng.Value = inpRng.Value
      End With

      headerCounter = headerCounter + 1
    End If
    
  Next

  'Add Listobjec if required on the export
  If addListObject Then
    Dim outLo As ListObject
    With sh
      Set outrng = .Range(.Cells(startLine, startColumn), .Cells(lastLine, headerCounter - 1))
      Set outLo = .ListObjects.Add(xlSrcRange, outrng, ,xlYes)
      outLo.Name = Name()
      outLo.TableStyle = Table.TableStyle
    End with
  End If

  'Return back the previously hidden columns
  ReturnBackHiddenColumns
  ResetCaches
End Sub

'@section Checkings
'===============================================================================
'@description Diagnostic logging and error handling.

'@label:log-info
'@sub-title Log a diagnostic message
'@details
'Adds a message to the internal IChecking object, creating it on first
'use. The checkCounter provides a unique key for each entry.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level from CheckingScope. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _ 
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "CustomTable " & Name() & " object trace"

    'First, instanciante the internalCheks Objects
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)
    
    'Add new informations to the checking object
    internalChecks.Add CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:has-checkings
'@prop-title Whether the table has logged diagnostic messages
'@details
'Returns True when at least one diagnostic entry has been recorded.
'@return Boolean. True when diagnostic entries exist.
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@details
'Returns the internal IChecking object containing all logged messages.
'Returns Nothing when no entries have been recorded.
'@return Object. An IChecking instance, or Nothing.
'@export
Private Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
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


'@section Interface
'===============================================================================
'@description Delegated members satisfying the ICustomTable contract.

Private Property Get ICustomTable_HeaderRange() As Range
  Set ICustomTable_HeaderRange = HeaderRange()
End Property

Private Property Get ICustomTable_IdValue() As String
  ICustomTable_IdValue = IdValue()
End Property

Private Property Get ICustomTable_DataRange( _
          Optional ByVal colName As String = "__all__", _
          Optional ByVal includeHeaders As Boolean = False, _
          Optional ByVal strictSearch As Boolean = False, _
          Optional ByVal matchCase As Boolean = False) As Range
  Set ICustomTable_DataRange = DataRange(colName, includeHeaders, strictSearch, matchCase)
End Property

Private Property Get ICustomTable_Name() As String
  ICustomTable_Name = Name()
End Property

Private Property Get ICustomTable_CellRange(ByVal colName As String, _ 
                                            ByVal lineNum As Long) As Range
  Set ICustomTable_CellRange = CellRange(colName, lineNum)
End Property

Private Property Get ICustomTable_Value(ByVal colName As String, _ 
                                        ByVal keyName As String) As String
  ICustomTable_Value = Value(colName, keyName)
End Property

Private Property Get ICustomTable_ValueRange(ByVal colName As String, _ 
                                             ByVal keyName As String) As Range

    Set ICustomTable_ValueRange = ValueRange(colName, keyName)
End Property

Private Sub ICustomTable_AddIds()
    AddIds
End Sub


Private Sub ICustomTable_AddRows(Optional ByVal nbRows As Long = 5, _ 
                                 Optional Byval insertShift As Boolean = False, _ 
                                 Optional ByVal includeIds As Boolean = True)

    AddRows nbRows:=nbRows, _ 
            insertShift:=insertShift, _ 
            includeIds:=includeIds
End Sub

Private Sub ICustomTable_RemoveRows(Optional ByVal totalCount As Long = 0, _ 
                                    Optional ByVal includeIds As Boolean = True, _ 
                                    Optional ByVal forceShift As Boolean = False)
  RemoveRows totalCount:=totalCount, includeIds:=includeIds, forceShift:=forceShift
End Sub


Private Sub ICustomTable_SetValidation(ByVal colName As String, _ 
     ByVal drop As IDropdownLists, _
     ByVal dropName As String, _ 
     Optional ByVal alertType As String = "info", _ 
     Optional ByVal message As String = vbNullString _ 
    )
    SetValidation colName:=colName, drop:=drop, dropName:=dropName, _ 
                  alertType:=alertType, message:=message
End Sub


Private Sub ICustomTable_SetValue(ByVal colName As String, _ 
        ByVal keyName As String, _ 
        ByVal newValue As String)
    SetValue colName, keyName, newValue
End Sub

Private Sub ICustomTable_RenameColumn(ByVal colName As String, _
                                      ByVal newName As String, _
                                      Optional ByVal strictSearch As Boolean = True, _
                                      Optional ByVal matchCase As Boolean = False)
    RenameColumn colName, newName, strictSearch, matchCase
End Sub


Private Sub ICustomTable_Clean()
  Clean
End Sub


Private Sub ICustomTable_Sort(Optional ByVal colName As String = vbNullString, _ 
                              Optional ByVal colList As Object = Nothing, _ 
                              Optional ByVal directSort As Boolean = True, _ 
                              Optional ByVal strictSearch As Boolean = True)
  Sort colName:=colName, colList:=colList, directSort:=directSort, strictSearch:=strictSearch
End Sub


Private Sub ICustomTable_InsertRowsAt(ByVal targetCell As Range, _ 
                                      Optional ByVal insertShift As Boolean = False, _
                                      Optional ByVal includeIds As Boolean = True)
  InsertRowsAt targetCell:=targetCell, _
               insertShift:=insertShift, _
               includeIds:=includeIds
End Sub

Private Sub ICustomTable_DeleteRowsAt(ByVal targetCell As Range, _ 
                                     Optional ByVal includeIds As Boolean = True, _
                                     Optional ByVal forceShift As Boolean = False)
  DeleteRowsAt targetCell:=targetCell, _
               includeIds:=includeIds, _
               forceShift:=forceShift
End Sub

Private Sub ICustomTable_Import( _
              ByVal impTab As Object, _
              Optional ByVal pasteAtBottom As Boolean = False, _
              Optional ByVal strictColumnSearch As Boolean = False, _
              Optional ByVal insertShift As Boolean = True, _ 
              Optional ByVal formatHeaders As Object = Nothing, _ 
              Optional ByVal keepSourceHeaders As Boolean = False)

  Import impTab:=impTab, pasteAtBottom:=pasteAtBottom, _ 
         strictColumnSearch:=strictColumnSearch, _ 
         insertShift:=insertShift, _ 
         formatHeaders:=formatHeaders, _
         keepSourceHeaders:=keepSourceHeaders

End Sub


Private Sub ICustomTable_Export(ByVal sh As Worksheet, _
                               Optional ByVal headersTable As Object = Nothing, _
                               Optional ByVal startLine As Long = 1, _ 
                               Optional ByVal startColumn As Long = 1, _ 
                               Optional ByVal addListObject As Boolean = False)
  Export sh, headersTable, startLine, startColumn, addListObject
End Sub


Private Sub ICustomTable_ImportAll(ByVal impTab As Object)
  ImportAll impTab
End Sub

Private Sub ICustomTable_RestoreTableSnapshot()
  RestoreTableSnapshot
End Sub

Private Sub ICustomTable_SaveTableSnapshot()
  SaveTableSnapshot
End Sub

Private Property Get ICustomTable_ImportColumnsNotFound() As BetterArray
  Set ICustomTable_ImportColumnsNotFound = ImportColumnsNotFound()
End Property

Private Property Get ICustomTable_HasColumnsNotImported() As Boolean
    ICustomTable_HasColumnsNotImported = HasColumnsNotImported()
End Property


Private Property Get ICustomTable_HasCheckings() As Boolean
    ICustomTable_HasCheckings = HasCheckings()
End Property

Private Property Get ICustomTable_CheckingValues() As Object
     Set ICustomTable_CheckingValues = CheckingValues()
End Property
