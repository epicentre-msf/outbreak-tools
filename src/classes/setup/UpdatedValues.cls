VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpdatedValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Watch updated values in the setup"

'@PredeclaredId
'@Folder("Setup")
'@ModuleDescription("Tracks watched setup columns and exposes update status to events")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation, ParameterNotUsed
Option Explicit

'@class UpdatedValues
'@description
'Tracks the setup columns that are watched for a change, and answers whether
'one of them has been edited.
'
'A setup column is watched when one of the three rows above its header carries
'a known tag: "watch for update", "translate as text" or "translate as
'formula". AddColumns reads those tags and writes one row per watched column
'into the registry table. CheckUpdate flips that row's flag when the user edits
'the column, and IsUpdated answers the flag. AddSheet registers a worksheet,
'RemoveLo drops one registry table, ClearUp and DeleteUp empty the registry,
'and SwitchTagsToNo and SwitchTagsToYes set every flag at once.
'
'ON THE SHEET
'-------------------------------------------------------------------------------
'The registry worksheet holds ONE registry ListObject, named UpLo_<sheetkey>
'where sheetkey is the normalised worksheet name. It carries four columns in
'this order, and both the names and the order are a published format that
'SetupTranslationsTable reads:
'
'    colname      | rngname     | updated | headername
'    <table>-<col>| RNG_...     | no      | watch for update
'
'Each row also has a workbook-level defined name, RNG_<table>_<column>_<sheet>,
'pointing at the watched column. That name is how a registry row turns back
'into a Range for the translations run. It is built at build time, and the
'change handler leaves it alone.
'
'When the registry empties, the table is unlisted and its cells are cleared.
'The sheet is then free of tables, which is what SetupTranslationsTable reads:
'it treats every table it finds on that sheet as a registry.
'
'IN MEMORY
'-------------------------------------------------------------------------------
'The registry is read once, on first use, in a single .Value2 call, into keyed
'Collections: the composite key "<tablekey>|<columnkey>" points at a row index
'and a flag, an alias map answers IsUpdated, and a table set lets the change
'handler leave an unwatched table without reading anything. Scripting.Dictionary
'is missing on Mac Excel, so these are keyed Collections.
'
'An edit inside a watched table costs four to five COM crossings whatever the
'registry holds, and the flag is written the first time a column flips within a
'reset cycle.
'
'TWO READERS, ONE FLAG COLUMN
'-------------------------------------------------------------------------------
'The translations run resets the flags itself through SwitchTagsToNo. The
'analysis dropdown refresh resets nothing, so it calls AcknowledgeUpdates:
'every flag reading "yes" at that moment goes into an in-memory set, and
'IsUpdated answers False for those keys until the user edits the column again.
'The set lives in memory alone, so opening the workbook starts it empty.
'
'CREATE WRITES NOTHING
'-------------------------------------------------------------------------------
'Building an instance binds a worksheet and stops there. Two helper factories
'build a fresh instance on every ribbon click.


'@section ConstantsAndState
'===============================================================================
'RANGE_PREFIX           Prefix used for workbook-level named ranges.
'LISTOBJECT_PREFIX      Prefix used when creating the update registry table.
'STATUS_DEFAULT         Default status flag stored in the registry (no changes detected).
'STATUS_UPDATED         Updated status flag stored in the registry after a change is detected.
'KEY_SEPARATOR          Separator between the table key and the column key of an index key.
'LABEL_SEPARATOR        Separator between the table name and the column name of a registry label.
'TUpValues              Internal state: the worksheet, the registry table and the in-memory index.
Private Const CLASS_NAME As String = "UpdatedValues"
Private Const RANGE_PREFIX As String = "RNG_"
Private Const LISTOBJECT_PREFIX As String = "UpLo_"
Private Const STATUS_DEFAULT As String = "no"
Private Const STATUS_UPDATED As String = "yes"
Private Const HEADER_COLUMN_NAME As String = "colname"
Private Const HEADER_RANGE_NAME As String = "rngname"
Private Const HEADER_STATUS As String = "updated"
Private Const HEADER_TAG As String = "headername"
Private Const TAG_TRANSLATE_TEXT As String = "translate as text"
Private Const TAG_TRANSLATE_FORMULA As String = "translate as formula"
Private Const TAG_WATCH_UPDATE As String = "watch for update"
Private Const MAX_TAG_LOOKUP_DEPTH As Long = 3
Private Const REGISTRY_COLUMN_COUNT As Long = 4
Private Const KEY_SEPARATOR As String = "|"
Private Const LABEL_SEPARATOR As String = "-"
Private Const FALLBACK_IDENTIFIER As String = "field"

Private Enum RegistryColumnIndex
  RegistryColumnLabel = 1
  RegistryRangeName = 2
  RegistryStatus = 3
  RegistryHeaderTag = 4
End Enum

Private Enum RegistryEntryField
  EntryLabel = 0
  EntryRangeName = 1
  EntryStatus = 2
  EntryTag = 3
  EntryHeader = 4
End Enum

Private Type TUpValues
  sh As Worksheet                 'Worksheet hosting the registry table
  sheetKey As String              'Normalised worksheet name, read once
  registry As ListObject          'The registry table, resolved lazily
  statusColumn As Range           'DataBodyRange of the "updated" column
  rowCount As Long                'Number of registry rows the index describes
  keys As Collection              'Composite keys in row order
  rowsByKey As Collection         '"<table>|<column>" -> row index
  statusByKey As Collection       '"<table>|<column>" -> "yes" / "no"
  keyByAlias As Collection        'column, label or range name -> composite key
  watchedTables As Collection     'table key -> True, the early exit of CheckUpdate
  handled As Collection           'Keys the in-session reader has acknowledged
  pending As Collection           'Registry rows being built during a build pass
  batching As Boolean             'True while AddSheet groups many tables into one write
  indexLoaded As Boolean          'True once the in-memory index describes the sheet
  sealed As Boolean               'When True, setup-only setters are blocked
End Type

Private this As TUpValues

'@section Factory
'===============================================================================
'@label:create
'@sub-title Build a new watcher bound to the specified registry worksheet.
'@details
'Binds the worksheet and nothing else. No cell is read and no table is created
'until a member that needs the registry runs.
'@param sh Worksheet. Registry worksheet hosting the watcher table.
'@param ignoredId Optional String. Reserved for future use (ignored).
'@return UpdatedValues. Configured watcher instance.
'@export
Public Function Create(ByVal sh As Worksheet, Optional ByVal ignoredId As String = vbNullString) As UpdatedValues
  Dim instance As UpdatedValues

  If sh Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "Worksheet reference is required to create UpdatedValues"
  End If

  Set instance = New UpdatedValues
  Set instance.Wksh = sh
  instance.Seal
  Set Create = instance
End Function

'@label:seal
'@sub-title Seal the instance so setup-only setters can no longer be used.
Public Sub Seal()
  this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Raise when a setup-only setter is used after sealing.
Private Sub GuardNotSealed(ByVal propName As String)
  If this.sealed Then _
    ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@section PublicState
'===============================================================================
'@label:worksheet
'@prop-title Worksheet hosting the watcher registry table.
'@return Worksheet. The registry worksheet.
'@export
Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

'@label:worksheet-set
'@prop-title Bind the watcher to a registry worksheet.
'@param sh Worksheet. The worksheet to assign.
Public Property Set Wksh(ByVal sh As Worksheet)
  GuardNotSealed "Wksh"
  If sh Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "Worksheet reference cannot be Nothing"
  End If

  Set this.sh = sh
  this.sheetKey = vbNullString
  ResetIndex
End Property

'@label:ensure-worksheet
'@sub-title Guard that the registry worksheet reference is initialised before use.
Private Function EnsureWorksheet() As Worksheet
  If this.sh Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "UpdatedValues worksheet was not initialised"
  End If
  Set EnsureWorksheet = this.sh
End Function

'@section ColumnRegistration
'===============================================================================
'@label:add-columns
'@sub-title Register or refresh watched columns for the supplied ListObject.
'@details
'Reads the three rows above the header in one call and the header row in one
'call, keeps the columns whose tag is one of the three known tags, and rewrites
'the registry rows for this table. A table with no tagged column creates
'nothing. The flag of a column that was already registered is carried over, so
'a rebuild does not lose a pending change.
'@param source ListObject. Table to inspect for update tags.
Public Sub AddColumns(ByVal source As ListObject)
  Dim headerRow As Range
  Dim headers As Variant
  Dim tagBlock As Variant
  Dim tagRows As Long
  Dim columnCount As Long
  Dim idx As Long
  Dim tableName As String
  Dim tableKey As String
  Dim headerText As String
  Dim tagValue As String
  Dim columnKey As String
  Dim wanted As Collection
  Dim standalone As Boolean

  If source Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "ListObject reference is required when registering updated values"
  End If

  Set headerRow = source.HeaderRowRange
  If headerRow Is Nothing Then Exit Sub

  tableName = source.Name
  tableKey = NormalizeIdentifier(tableName)
  columnCount = headerRow.Cells.Count
  headers = headerRow.Value2
  ReadTagBlock headerRow, tagBlock, tagRows

  Set wanted = New Collection
  For idx = 1 To columnCount
    headerText = CellText(headers, 1, idx, 1, columnCount)
    If LenB(headerText) > 0 Then
      tagValue = TagForColumn(tagBlock, tagRows, columnCount, idx)
      If LenB(tagValue) > 0 Then
        columnKey = NormalizeIdentifier(headerText)
        wanted.Add Array(BuildColumnLabel(tableName, headerText), _
                         BuildRangeName(tableKey, columnKey), _
                         STATUS_DEFAULT, _
                         tagValue, _
                         headerText)
      End If
    End If
  Next idx

  standalone = Not this.batching
  If standalone Then BeginBatch

  MergeTableEntries tableName, wanted, source

  If standalone Then EndBatch
End Sub

'@label:add-sheet
'@sub-title Register tagged columns for every ListObject on the provided worksheet.
'@details
'Every table on the sheet is read, and the registry is written once at the end
'rather than once per table.
Public Sub AddSheet(ByVal targetSheet As Worksheet)
  Dim lo As ListObject
  Dim errNumber As Long
  Dim errSource As String
  Dim errText As String

  If targetSheet Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "Worksheet reference is required when registering all updated values"
  End If

  If targetSheet.ListObjects.Count = 0 Then Exit Sub

  BeginBatch
  On Error GoTo CleanFail

  For Each lo In targetSheet.ListObjects
    AddColumns lo
  Next lo

  EndBatch
  Exit Sub

CleanFail:
  'The error is captured before the clean-up, because clearing the batch runs
  'under its own error trap and would overwrite Err.
  errNumber = Err.Number
  errSource = Err.Source
  errText = Err.Description
  this.batching = False
  Set this.pending = Nothing
  ResetIndex
  Err.Raise errNumber, errSource, errText
End Sub

'@label:remove-lo
'@sub-title Remove registry entries associated with the supplied ListObject.
Public Sub RemoveLo(ByVal target As ListObject)
  Dim standalone As Boolean

  If target Is Nothing Then Exit Sub

  standalone = Not this.batching
  If standalone Then BeginBatch

  MergeTableEntries target.Name, New Collection, Nothing

  If standalone Then EndBatch
End Sub

'@label:merge-table-entries
'@sub-title Replace the registry rows belonging to one table with the supplied rows.
'@param tableName String. Source table whose rows are being replaced.
'@param wanted Collection. Rows to keep for that table, empty to drop them all.
'@param source ListObject. Table the defined names point at, Nothing when dropping.
Private Sub MergeTableEntries(ByVal tableName As String, _
                              ByVal wanted As Collection, _
                              ByVal source As ListObject)
  Dim kept As Collection
  Dim heldStatus As Collection
  Dim obsoleteNames As Collection
  Dim entry As Variant
  Dim item As Variant
  Dim rangeName As String
  Dim status As String

  Set kept = New Collection
  Set heldStatus = New Collection
  Set obsoleteNames = New Collection

  'Rows from the other tables are kept as they are. The rows of this table are
  'set aside, so their flags can be carried over and their defined names can be
  'deleted when the column is no longer tagged.
  For Each item In this.pending
    If ColumnLabelMatchesTable(CStr(item(EntryLabel)), tableName) Then
      rangeName = CStr(item(EntryRangeName))
      KeyedAdd heldStatus, CStr(item(EntryStatus)), NormalizeValue(rangeName)
      KeyedAdd obsoleteNames, rangeName, NormalizeValue(rangeName)
    Else
      kept.Add item
    End If
  Next item

  For Each entry In wanted
    rangeName = CStr(entry(EntryRangeName))
    status = KeyedText(heldStatus, NormalizeValue(rangeName))
    If status <> STATUS_UPDATED Then status = STATUS_DEFAULT

    kept.Add Array(CStr(entry(EntryLabel)), rangeName, status, CStr(entry(EntryTag)))
    KeyedRemove obsoleteNames, NormalizeValue(rangeName)

    If Not source Is Nothing Then
      EnsureNamedRange source, CStr(entry(EntryHeader)), rangeName
    End If
  Next entry

  For Each item In obsoleteNames
    RemoveNamedRange CStr(item)
  Next item

  Set this.pending = kept
End Sub

'@label:column-label-matches-table
'@sub-title Determine whether a registry column label belongs to the provided ListObject name.
Private Function ColumnLabelMatchesTable(ByVal columnLabel As String, _
                                         ByVal tableName As String) As Boolean
  Dim expectedPrefix As String

  expectedPrefix = tableName & LABEL_SEPARATOR
  ColumnLabelMatchesTable = (StrComp(Left$(columnLabel, Len(expectedPrefix)), expectedPrefix, vbTextCompare) = 0)
End Function

'@label:read-tag-block
'@sub-title Read the rows above the header row in one call.
'@param headerRow Range. Header row of the source table.
'@param block Variant. Output, the values of the rows above the header.
'@param tagRows Long. Output, how many rows the block holds.
Private Sub ReadTagBlock(ByVal headerRow As Range, ByRef block As Variant, ByRef tagRows As Long)
  Dim sh As Worksheet
  Dim firstRow As Long
  Dim topRow As Long
  Dim firstColumn As Long
  Dim lastColumn As Long

  block = Empty
  tagRows = 0

  firstRow = headerRow.Row
  If firstRow <= 1 Then Exit Sub

  topRow = firstRow - MAX_TAG_LOOKUP_DEPTH
  If topRow < 1 Then topRow = 1
  tagRows = firstRow - topRow

  Set sh = headerRow.Worksheet
  firstColumn = headerRow.Column
  lastColumn = firstColumn + headerRow.Cells.Count - 1

  block = sh.Range(sh.Cells(topRow, firstColumn), sh.Cells(firstRow - 1, lastColumn)).Value2
End Sub

'@label:tag-for-column
'@sub-title Read the metadata tag of a column, walking up from the header row.
'@details
'The walk answers with the first cell holding one of the three known tags. A
'section title or a note sitting between the tag row and the header is stepped
'over, and a value that is not a tag never wins. Analysis tables are stacked,
'so the row three above a header can belong to the table above it; looking for
'a known tag removes that risk with it.
Private Function TagForColumn(ByVal block As Variant, _
                              ByVal tagRows As Long, _
                              ByVal columnCount As Long, _
                              ByVal columnIndex As Long) As String
  Dim rowIndex As Long
  Dim candidate As String

  If tagRows = 0 Then Exit Function

  For rowIndex = tagRows To 1 Step -1
    candidate = NormalizeValue(CellText(block, rowIndex, columnIndex, tagRows, columnCount))
    If ShouldWatchColumn(candidate) Then
      TagForColumn = candidate
      Exit Function
    End If
  Next rowIndex
End Function

'@label:should-watch-column
'@sub-title Determine whether a column should be monitored based on its tag value.
'@param tagValue String. Tag retrieved from worksheet metadata.
Private Function ShouldWatchColumn(ByVal tagValue As String) As Boolean
  Select Case tagValue
    Case TAG_TRANSLATE_TEXT, TAG_TRANSLATE_FORMULA, TAG_WATCH_UPDATE
      ShouldWatchColumn = True
  End Select
End Function

'@section RegistryWriting
'===============================================================================
'@label:begin-batch
'@sub-title Load the registry rows into memory so a build pass can change them.
Private Sub BeginBatch()
  Set this.pending = ReadRegistryEntries()
  this.batching = True
End Sub

'@label:end-batch
'@sub-title Write the registry rows back in one call and drop the in-memory index.
Private Sub EndBatch()
  WriteRegistryEntries this.pending
  this.batching = False
  Set this.pending = Nothing
  ResetIndex
End Sub

'@label:read-registry-entries
'@sub-title Read every registry row into a collection, in one call.
Private Function ReadRegistryEntries() As Collection
  Dim entries As Collection
  Dim lo As ListObject
  Dim dataRange As Range
  Dim values As Variant
  Dim rowIndex As Long
  Dim labelColumn As Long
  Dim nameColumn As Long
  Dim statusColumn As Long
  Dim tagColumn As Long
  Dim rangeName As String

  Set entries = New Collection
  Set ReadRegistryEntries = entries

  Set lo = ResolveRegistry(False)
  If lo Is Nothing Then Exit Function

  Set dataRange = lo.DataBodyRange
  If dataRange Is Nothing Then Exit Function

  labelColumn = RegistryColumnPosition(lo, HEADER_COLUMN_NAME, RegistryColumnLabel)
  nameColumn = RegistryColumnPosition(lo, HEADER_RANGE_NAME, RegistryRangeName)
  statusColumn = RegistryColumnPosition(lo, HEADER_STATUS, RegistryStatus)
  tagColumn = RegistryColumnPosition(lo, HEADER_TAG, RegistryHeaderTag)

  values = dataRange.Value2
  If Not IsArray(values) Then Exit Function

  For rowIndex = LBound(values, 1) To UBound(values, 1)
    rangeName = SafeText(values(rowIndex, nameColumn))
    If LenB(rangeName) > 0 Then
      entries.Add Array(SafeText(values(rowIndex, labelColumn)), _
                        rangeName, _
                        ReadStatus(SafeText(values(rowIndex, statusColumn))), _
                        SafeText(values(rowIndex, tagColumn)))
    End If
  Next rowIndex
End Function

'@label:write-registry-entries
'@sub-title Write the supplied rows over the registry table in one call.
'@details
'An empty registry leaves no table behind: the ListObject is unlisted and its
'cells are cleared. A table always keeps one data row, so an empty table would
'mean a blank row every reader has to step over.
Private Sub WriteRegistryEntries(ByVal entries As Collection)
  Dim lo As ListObject
  Dim buffer As Variant
  Dim entry As Variant
  Dim rowIndex As Long
  Dim rowTotal As Long
  Dim previousRows As Long
  Dim leftover As Range

  If entries Is Nothing Then Exit Sub

  rowTotal = entries.Count
  If rowTotal = 0 Then
    RemoveRegistryTable
    Exit Sub
  End If

  Set lo = ResolveRegistry(True)
  If lo Is Nothing Then Exit Sub

  ReDim buffer(1 To rowTotal, 1 To REGISTRY_COLUMN_COUNT)
  rowIndex = 0
  For Each entry In entries
    rowIndex = rowIndex + 1
    buffer(rowIndex, RegistryColumnLabel) = entry(EntryLabel)
    buffer(rowIndex, RegistryRangeName) = entry(EntryRangeName)
    buffer(rowIndex, RegistryStatus) = entry(EntryStatus)
    buffer(rowIndex, RegistryHeaderTag) = entry(EntryTag)
  Next entry

  previousRows = 0
  If Not lo.DataBodyRange Is Nothing Then previousRows = lo.DataBodyRange.Rows.Count

  'The rows a shrinking table leaves behind keep their content, so they are
  'held now and cleared once the table has been resized.
  Set leftover = Nothing
  If previousRows > rowTotal Then
    Set leftover = lo.Range.Rows(rowTotal + 2).Resize(previousRows - rowTotal)
  End If

  'Resize grows the table OVER whatever sits under it, and the write that
  'follows overwrites those cells. The registry sheet holds this one table and
  'nothing else, which is what makes the grow safe here.
  lo.Resize lo.Range.Resize(rowTotal + 1)
  lo.DataBodyRange.Value2 = buffer

  If Not leftover Is Nothing Then leftover.ClearContents
End Sub

'@label:resolve-registry
'@sub-title Locate the registry table, building it when asked.
Private Function ResolveRegistry(ByVal createWhenMissing As Boolean) As ListObject
  Dim sh As Worksheet
  Dim lo As ListObject

  'The held table is dropped by ResetIndex and by RemoveRegistryTable, which are
  'the two things that can make it stale.
  If Not this.registry Is Nothing Then
    Set ResolveRegistry = this.registry
    Exit Function
  End If

  Set sh = EnsureWorksheet()

  On Error Resume Next
    Set lo = sh.ListObjects(RegistryName())
  On Error GoTo 0

  If lo Is Nothing And createWhenMissing Then Set lo = CreateRegistryTable(sh)

  Set this.registry = lo
  Set ResolveRegistry = lo
End Function

'@label:create-registry-table
'@sub-title Build an empty registry ListObject with the expected headers.
Private Function CreateRegistryTable(ByVal sh As Worksheet) As ListObject
  Dim headerStart As Range
  Dim tableRange As Range
  Dim lo As ListObject

  Set headerStart = sh.Cells(1, 1)
  headerStart.Value = HEADER_COLUMN_NAME
  headerStart.Offset(0, 1).Value = HEADER_RANGE_NAME
  headerStart.Offset(0, 2).Value = HEADER_STATUS
  headerStart.Offset(0, 3).Value = HEADER_TAG

  Set tableRange = sh.Range(headerStart, headerStart.Offset(1, REGISTRY_COLUMN_COUNT - 1))
  Set lo = sh.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
  lo.Name = RegistryName()
  EnsureRegistryHeaders lo

  Set CreateRegistryTable = lo
End Function

'@label:ensure-registry-headers
'@sub-title Guarantee the registry ListObject exposes the expected header names.
Private Sub EnsureRegistryHeaders(ByVal lo As ListObject)
  Dim expected As Variant
  Dim idx As Long

  expected = Array(HEADER_COLUMN_NAME, HEADER_RANGE_NAME, HEADER_STATUS, HEADER_TAG)

  Do While lo.ListColumns.Count < REGISTRY_COLUMN_COUNT
    lo.ListColumns.Add
  Loop

  For idx = LBound(expected) To UBound(expected)
    lo.ListColumns(idx + 1).Name = CStr(expected(idx))
    lo.HeaderRowRange.Cells(1, idx + 1).Value = expected(idx)
  Next idx
End Sub

'@label:remove-registry-table
'@sub-title Take the registry table off the sheet and clear the cells it held.
Private Sub RemoveRegistryTable()
  Dim lo As ListObject
  Dim held As Range

  Set lo = ResolveRegistry(False)
  If lo Is Nothing Then Exit Sub

  Set held = lo.Range
  lo.Unlist
  held.Cells.Clear
  Set this.registry = Nothing
End Sub

'@section NamedRanges
'===============================================================================
'@label:ensure-named-range
'@sub-title Create or refresh the workbook defined name pointing to the watched column.
Private Sub EnsureNamedRange(ByVal source As ListObject, _
                             ByVal columnName As String, _
                             ByVal rangeName As String)
  Dim wb As Workbook
  Dim definedName As Name
  Dim refersTo As String

  If NormalizeValue(rangeName) = vbNullString Then Exit Sub

  refersTo = "=" & source.Name & "[" & columnName & "]"
  Set wb = source.Parent.Parent

  On Error Resume Next
    Set definedName = wb.Names(rangeName)
  On Error GoTo 0

  If definedName Is Nothing Then
    wb.Names.Add Name:=rangeName, RefersToR1C1:=refersTo, Visible:=False
  ElseIf StrComp(definedName.RefersTo, refersTo, vbBinaryCompare) <> 0 Then
    definedName.RefersTo = refersTo
  End If
End Sub

'@label:remove-named-range
'@sub-title Delete a single defined name if it exists in the workbook.
Private Sub RemoveNamedRange(ByVal rangeName As String)
  Dim wb As Workbook

  If NormalizeValue(rangeName) = vbNullString Then Exit Sub

  Set wb = EnsureWorksheet().Parent

  On Error Resume Next
    wb.Names(rangeName).Delete
  On Error GoTo 0
End Sub

'@label:remove-sheet-named-ranges
'@sub-title Delete every watcher defined name belonging to the registry sheet, in one pass.
Private Sub RemoveSheetNamedRanges()
  Dim wb As Workbook
  Dim definedName As Name
  Dim idx As Long
  Dim nameText As String
  Dim suffix As String

  Set wb = EnsureWorksheet().Parent
  suffix = "_" & SheetKey()

  For idx = wb.Names.Count To 1 Step -1
    Set definedName = wb.Names(idx)

    nameText = vbNullString
    On Error Resume Next
      nameText = definedName.Name
    On Error GoTo 0

    If StrComp(Left$(nameText, Len(RANGE_PREFIX)), RANGE_PREFIX, vbTextCompare) = 0 Then
      If Len(nameText) > Len(suffix) Then
        If StrComp(Right$(nameText, Len(suffix)), suffix, vbTextCompare) = 0 Then
          On Error Resume Next
            definedName.Delete
          On Error GoTo 0
        End If
      End If
    End If
  Next idx
End Sub

'@section StatusTracking
'===============================================================================
'@label:is-updated
'@prop-title Determine whether the tracked column has been marked as updated.
'@details
'The identifier is matched against the column name, the whole registry label
'and the defined name, all normalised the same way. So "variable name" and
'"variable_name" ask the same question, which is what the Dictionary headers
'need. The answer comes from memory and costs no worksheet call.
'@param checkColName String. Column name, registry label or defined name.
'@return Boolean. True when the column carries a change nobody has acted on.
'@export
Public Property Get IsUpdated(ByVal checkColName As String) As Boolean
  Dim key As String

  If NormalizeValue(checkColName) = vbNullString Then Exit Property

  EnsureIndex
  key = KeyedText(this.keyByAlias, NormalizeIdentifier(checkColName))
  If LenB(key) = 0 Then Exit Property

  IsUpdated = (KeyedText(this.statusByKey, key) = STATUS_UPDATED) And Not KeyedHas(this.handled, key)
End Property

'@label:acknowledge-updates
'@sub-title Record every flag that currently reads "yes" as already acted on.
'@details
'The reader that never resets the flag column calls this after it has rebuilt
'what it had to rebuild. The set is memory only, so a workbook open starts it
'empty and the first Analysis activation of a session still rebuilds.
'@export
Public Sub AcknowledgeUpdates()
  Dim key As Variant

  EnsureIndex
  If this.keys Is Nothing Then Exit Sub
  If this.handled Is Nothing Then Set this.handled = New Collection

  For Each key In this.keys
    If KeyedText(this.statusByKey, CStr(key)) = STATUS_UPDATED Then
      KeyedAdd this.handled, True, CStr(key)
    End If
  Next key
End Sub

'@label:check-update
'@sub-title Evaluate a range change against the registered watchers and flag updates.
'@details
'Every watched column the change covers is flagged, so a paste across two
'watched columns flips two flags.
'
'The handler only ever writes a flag. Cleanup belongs to AddColumns and
'AddSheet, which run at build time: a defined name can fail to resolve for a
'moment while a source table resizes or a sheet is renamed, and deleting a row
'on that would take the registration away for good.
'
'The cost is four to five worksheet calls whatever the registry holds. A table
'nobody watches is left after two of them, and the flag write happens only the
'first time a column flips within a reset cycle.
'@param tagReference Variant. Kept for the call shape; the table is read from Target.
'@param Target Range. The changed cells to evaluate.
Public Sub CheckUpdate(ByVal tagReference As Variant, ByVal Target As Range)
  Dim lo As ListObject

  If Target Is Nothing Then Exit Sub

  EnsureIndex
  If this.rowCount = 0 Then Exit Sub

  Set lo = Target.ListObject
  If lo Is Nothing Then
    MarkWideTarget Target
    Exit Sub
  End If

  MarkColumns lo, Target
End Sub

'@label:mark-wide-target
'@sub-title Flag the watched tables a whole-row or whole-column change reaches.
'@details
'Selecting a worksheet row or column gives a target with no ListObject. Only
'the watched tables on that sheet are asked, one Intersect each.
Private Sub MarkWideTarget(ByVal changed As Range)
  Dim lo As ListObject
  Dim body As Range
  Dim hit As Range

  For Each lo In changed.Worksheet.ListObjects
    If KeyedHas(this.watchedTables, NormalizeIdentifier(lo.Name)) Then
      Set body = lo.DataBodyRange
      If Not body Is Nothing Then
        Set hit = Intersect(changed, body)
        If Not hit Is Nothing Then MarkColumns lo, hit
      End If
    End If
  Next lo
End Sub

'@label:mark-columns
'@sub-title Flag every watched column of one table that the change covers.
'@details
'The header row is read on every edit rather than cached. It is one worksheet
'call, and it stays right when the user inserts or deletes a column in the
'source table; a cached position would need an invalidation signal that costs a
'call to read anyway.
Private Sub MarkColumns(ByVal lo As ListObject, ByVal changed As Range)
  Dim tableKey As String
  Dim headers As Variant
  Dim firstColumn As Long
  Dim headerCount As Long
  Dim area As Range
  Dim columnIndex As Long
  Dim relative As Long

  tableKey = NormalizeIdentifier(lo.Name)
  If Not KeyedHas(this.watchedTables, tableKey) Then Exit Sub

  headers = lo.HeaderRowRange.Value2
  firstColumn = lo.Range.Column

  If IsArray(headers) Then
    headerCount = UBound(headers, 2)
  Else
    headerCount = 1
  End If

  For Each area In changed.Areas
    For columnIndex = area.Column To area.Column + area.Columns.Count - 1
      relative = columnIndex - firstColumn + 1
      If relative >= 1 And relative <= headerCount Then
        MarkKey tableKey, CellText(headers, 1, relative, 1, headerCount)
      End If
    Next columnIndex
  Next area
End Sub

'@label:mark-key
'@sub-title Write the updated flag for one watched column, the first time it changes.
Private Sub MarkKey(ByVal tableKey As String, ByVal headerText As String)
  Dim key As String
  Dim rowIndex As Long

  If NormalizeValue(headerText) = vbNullString Then Exit Sub

  key = tableKey & KEY_SEPARATOR & NormalizeIdentifier(headerText)
  rowIndex = KeyedNumber(this.rowsByKey, key)
  If rowIndex = 0 Then Exit Sub
  If rowIndex > this.rowCount Then Exit Sub

  'The acknowledgement goes first, and it goes even when the flag already reads
  '"yes". A reader that has acted on this column has to see the next edit too.
  KeyedRemove this.handled, key

  If KeyedText(this.statusByKey, key) = STATUS_UPDATED Then Exit Sub
  If this.statusColumn Is Nothing Then Exit Sub

  this.statusColumn.Cells(rowIndex, 1).Value = STATUS_UPDATED
  SetStatus key, STATUS_UPDATED
End Sub

'@section CleanupHelpers
'===============================================================================
'@label:clear-up
'@sub-title Reset all tracked columns to the default (not updated) status.
Public Sub ClearUp()
  SwitchTags STATUS_DEFAULT
End Sub

'@label:delete-up
'@sub-title Remove every registry table on the sheet and the defined names they used.
'@details
'Every ListObject on the registry sheet goes, including the ones this class
'never created. A workbook built before the single-table layout carries one
'registry per source table and a name index table beside them; leaving any of
'them behind would make the translations run read the old rows and the new rows
'side by side and add every label twice. SetupPreparation calls this before it
'rebuilds, so this sweep is the whole upgrade.
'@export
Public Sub DeleteUp()
  Dim sh As Worksheet
  Dim lo As ListObject
  Dim held As Range
  Dim idx As Long

  Set sh = EnsureWorksheet()

  For idx = sh.ListObjects.Count To 1 Step -1
    Set lo = sh.ListObjects(idx)
    Set held = lo.Range
    lo.Unlist
    held.Cells.Clear
  Next idx

  RemoveSheetNamedRanges
  sh.Cells.Clear

  ResetIndex
  Set this.handled = Nothing
End Sub

'@label:switch-tags
'@sub-title Set every registry status flag on the sheet to the supplied value.
'@details
'Every table on the sheet carrying an "updated" column is reset, including the
'ones this class never built. The master setup registry is hand-built and a
'workbook made before the single-table layout carries its own registries; both
'have to reset when the translations button is pressed.
Private Sub SwitchTags(ByVal value As String)
  Dim sh As Worksheet
  Dim lo As ListObject
  Dim statusColumn As Range

  'Both resets start a fresh acknowledgement cycle.
  Set this.handled = Nothing

  Set sh = EnsureWorksheet()

  For Each lo In sh.ListObjects
    Set statusColumn = StatusColumnOf(lo)
    If Not statusColumn Is Nothing Then statusColumn.Value = value
  Next lo

  RefreshStatusIndex value
End Sub

'@label:switch-tags-to-no
'@sub-title Force every registry status flag to the default value.
'@export
Public Sub SwitchTagsToNo()
  SwitchTags STATUS_DEFAULT
End Sub

'@label:switch-tags-to-yes
'@sub-title Force every registry status flag to the updated value.
'@export
Public Sub SwitchTagsToYes()
  SwitchTags STATUS_UPDATED
End Sub

'@label:status-column-of
'@sub-title Return the flag column of a table that carries one.
Private Function StatusColumnOf(ByVal lo As ListObject) As Range
  Dim listCol As ListColumn

  If lo Is Nothing Then Exit Function

  On Error Resume Next
    Set listCol = lo.ListColumns(HEADER_STATUS)
  On Error GoTo 0

  If listCol Is Nothing Then Exit Function
  Set StatusColumnOf = listCol.DataBodyRange
End Function

'@section InMemoryIndex
'===============================================================================
'@label:ensure-index
'@sub-title Build the in-memory index the first time something needs it.
Private Sub EnsureIndex()
  If this.indexLoaded Then Exit Sub
  LoadIndex
End Sub

'@label:reset-index
'@sub-title Drop the in-memory index so the next read rebuilds it from the sheet.
'@details
'The acknowledged set is left alone. It answers for the session, and the index
'is dropped whenever the registry is written.
Private Sub ResetIndex()
  this.indexLoaded = False
  this.rowCount = 0
  Set this.registry = Nothing
  Set this.statusColumn = Nothing
  Set this.keys = Nothing
  Set this.rowsByKey = Nothing
  Set this.statusByKey = Nothing
  Set this.keyByAlias = Nothing
  Set this.watchedTables = Nothing
End Sub

'@label:load-index
'@sub-title Read the whole registry in one call and build the lookups from it.
Private Sub LoadIndex()
  Dim lo As ListObject
  Dim dataRange As Range
  Dim values As Variant
  Dim rowIndex As Long
  Dim labelColumn As Long
  Dim nameColumn As Long
  Dim statusColumn As Long
  Dim label As String
  Dim rangeName As String
  Dim tableKey As String
  Dim columnKey As String
  Dim key As String

  Set this.keys = New Collection
  Set this.rowsByKey = New Collection
  Set this.statusByKey = New Collection
  Set this.keyByAlias = New Collection
  Set this.watchedTables = New Collection
  Set this.statusColumn = Nothing
  this.rowCount = 0
  this.indexLoaded = True

  Set lo = ResolveRegistry(False)
  If lo Is Nothing Then Exit Sub

  labelColumn = RegistryColumnPosition(lo, HEADER_COLUMN_NAME, RegistryColumnLabel)
  nameColumn = RegistryColumnPosition(lo, HEADER_RANGE_NAME, RegistryRangeName)
  statusColumn = RegistryColumnPosition(lo, HEADER_STATUS, RegistryStatus)

  Set this.statusColumn = ColumnData(lo, statusColumn)

  Set dataRange = lo.DataBodyRange
  If dataRange Is Nothing Then Exit Sub

  values = dataRange.Value2
  If Not IsArray(values) Then Exit Sub

  For rowIndex = LBound(values, 1) To UBound(values, 1)
    rangeName = SafeText(values(rowIndex, nameColumn))
    If LenB(rangeName) > 0 Then
      label = SafeText(values(rowIndex, labelColumn))
      tableKey = NormalizeIdentifier(TableNameFromLabel(label))
      columnKey = NormalizeIdentifier(ExtractColumnNameFromLabel(label))
      key = tableKey & KEY_SEPARATOR & columnKey

      this.keys.Add key
      KeyedAdd this.rowsByKey, rowIndex, key
      SetStatus key, ReadStatus(SafeText(values(rowIndex, statusColumn)))
      KeyedAdd this.watchedTables, True, tableKey

      'Three ways to name the same column, so a caller can ask with the column
      'header, the whole registry label or the defined name.
      KeyedAdd this.keyByAlias, key, columnKey
      KeyedAdd this.keyByAlias, key, NormalizeIdentifier(label)
      KeyedAdd this.keyByAlias, key, NormalizeIdentifier(rangeName)
    End If
  Next rowIndex

  this.rowCount = this.keys.Count
End Sub

'@label:refresh-status-index
'@sub-title Put the same flag on every key the index holds.
Private Sub RefreshStatusIndex(ByVal value As String)
  Dim key As Variant

  If Not this.indexLoaded Then Exit Sub
  If this.keys Is Nothing Then Exit Sub

  For Each key In this.keys
    SetStatus CStr(key), value
  Next key
End Sub

'@label:set-status
'@sub-title Record the flag of one key.
Private Sub SetStatus(ByVal key As String, ByVal value As String)
  KeyedRemove this.statusByKey, key
  KeyedAdd this.statusByKey, value, key
End Sub

'@label:read-status
'@sub-title Read a flag cell, answering the default for anything that is not "yes".
Private Function ReadStatus(ByVal valueText As String) As String
  If NormalizeValue(valueText) = STATUS_UPDATED Then
    ReadStatus = STATUS_UPDATED
  Else
    ReadStatus = STATUS_DEFAULT
  End If
End Function

'@label:registry-column-position
'@sub-title Find a registry column by header name, falling back to its published position.
Private Function RegistryColumnPosition(ByVal lo As ListObject, _
                                        ByVal headerName As String, _
                                        ByVal fallbackColumn As Long) As Long
  Dim listCol As ListColumn

  RegistryColumnPosition = fallbackColumn
  If lo Is Nothing Then Exit Function

  For Each listCol In lo.ListColumns
    If StrComp(Trim$(listCol.Name), headerName, vbTextCompare) = 0 Then
      RegistryColumnPosition = listCol.Index
      Exit Function
    End If
  Next listCol
End Function

'@label:column-data
'@sub-title Return the data body of one registry column by position.
Private Function ColumnData(ByVal lo As ListObject, ByVal position As Long) As Range
  If lo Is Nothing Then Exit Function
  If position < 1 Or position > lo.ListColumns.Count Then Exit Function
  Set ColumnData = lo.ListColumns(position).DataBodyRange
End Function

'@section Utilities
'===============================================================================
'@label:registry-name
'@sub-title Compose the name of the registry table hosted on this sheet.
Private Function RegistryName() As String
  RegistryName = LISTOBJECT_PREFIX & SheetKey()
End Function

'@label:sheet-key
'@sub-title Return the normalised worksheet name, read and normalised once.
Private Function SheetKey() As String
  If LenB(this.sheetKey) = 0 Then
    this.sheetKey = NormalizeIdentifier(EnsureWorksheet().Name)
  End If
  SheetKey = this.sheetKey
End Function

'@label:build-column-label
'@sub-title Compose the descriptive label stored in the registry for a watched column.
Private Function BuildColumnLabel(ByVal tableName As String, ByVal columnName As String) As String
  BuildColumnLabel = tableName & LABEL_SEPARATOR & columnName
End Function

'@label:table-name-from-label
'@sub-title Retrieve the table portion from a table-column label.
Private Function TableNameFromLabel(ByVal columnLabel As String) As String
  Dim separatorPos As Long

  separatorPos = InStr(columnLabel, LABEL_SEPARATOR)
  If separatorPos = 0 Then
    TableNameFromLabel = Trim$(columnLabel)
  Else
    TableNameFromLabel = Trim$(Left$(columnLabel, separatorPos - 1))
  End If
End Function

'@label:extract-column-name-from-label
'@sub-title Retrieve the column portion from a table-column label.
Private Function ExtractColumnNameFromLabel(ByVal columnLabel As String) As String
  Dim separatorPos As Long

  separatorPos = InStr(columnLabel, LABEL_SEPARATOR)
  If separatorPos = 0 Then
    ExtractColumnNameFromLabel = Trim$(columnLabel)
  Else
    ExtractColumnNameFromLabel = Trim$(Mid$(columnLabel, separatorPos + 1))
  End If
End Function

'@label:build-range-name
'@sub-title Generate the workbook defined name for a watched column, including table context.
Private Function BuildRangeName(ByVal tableKey As String, ByVal columnKey As String) As String
  Dim normalizedTable As String

  normalizedTable = tableKey
  If LenB(normalizedTable) = 0 Then normalizedTable = "table"

  BuildRangeName = RANGE_PREFIX & normalizedTable & "_" & columnKey & "_" & SheetKey()
End Function

'@label:normalize-identifier
'@sub-title Produce a safe workbook-level identifier from free-form text.
'@details
'Letters are lower-cased, digits are kept, and everything else becomes a single
'underscore, with leading and trailing underscores dropped. One pass writes the
'answer in place with Mid$, the way HiddenNames.SanitizeNameId does it.
Private Function NormalizeIdentifier(ByVal valueText As String) As String
  Dim source As String
  Dim result As String
  Dim idx As Long
  Dim outLen As Long
  Dim code As Long

  source = Trim$(Replace(valueText, ChrW$(160), " "))
  If LenB(source) = 0 Then
    NormalizeIdentifier = FALLBACK_IDENTIFIER
    Exit Function
  End If

  result = Space$(Len(source))
  outLen = 0

  For idx = 1 To Len(source)
    code = AscW(Mid$(source, idx, 1))

    If code >= 65 And code <= 90 Then
      outLen = outLen + 1
      Mid$(result, outLen, 1) = ChrW$(code + 32)
    ElseIf (code >= 97 And code <= 122) Or (code >= 48 And code <= 57) Then
      outLen = outLen + 1
      Mid$(result, outLen, 1) = ChrW$(code)
    ElseIf outLen > 0 Then
      'One underscore stands for any run of other characters, and a run at the
      'start of the text is dropped.
      If Mid$(result, outLen, 1) <> "_" Then
        outLen = outLen + 1
        Mid$(result, outLen, 1) = "_"
      End If
    End If
  Next idx

  Do While outLen > 0
    If Mid$(result, outLen, 1) <> "_" Then Exit Do
    outLen = outLen - 1
  Loop

  If outLen = 0 Then
    NormalizeIdentifier = FALLBACK_IDENTIFIER
  Else
    NormalizeIdentifier = Left$(result, outLen)
  End If
End Function

'@label:normalize-value
'@sub-title Normalise text by trimming spaces and applying lower case.
Private Function NormalizeValue(ByVal valueText As String) As String
  NormalizeValue = LCase$(Trim$(valueText))
End Function

'@label:safe-text
'@sub-title Read a cell value as text, answering an empty string for an error value.
Private Function SafeText(ByVal value As Variant) As String
  If IsError(value) Then Exit Function
  If IsNull(value) Then Exit Function

  On Error Resume Next
    SafeText = Trim$(CStr(value))
  On Error GoTo 0
End Function

'@label:cell-text
'@sub-title Read one cell out of a bulk .Value2 result.
'@details
'A single cell read through .Value2 answers a scalar rather than a 2D array, so
'both shapes are handled here.
Private Function CellText(ByVal block As Variant, _
                          ByVal rowIndex As Long, _
                          ByVal columnIndex As Long, _
                          ByVal rowTotal As Long, _
                          ByVal columnTotal As Long) As String
  If IsEmpty(block) Then Exit Function

  If IsArray(block) Then
    CellText = SafeText(block(rowIndex, columnIndex))
  ElseIf rowTotal = 1 And columnTotal = 1 Then
    CellText = SafeText(block)
  End If
End Function

'@label:keyed-add
'@sub-title Add an item under a key, leaving the first entry in place.
Private Sub KeyedAdd(ByVal target As Collection, ByVal item As Variant, ByVal key As String)
  If target Is Nothing Then Exit Sub
  If LenB(key) = 0 Then Exit Sub

  'A key already taken raises 457, and the first entry is the one to keep.
  On Error Resume Next
    target.Add item, key
  On Error GoTo 0
End Sub

'@label:keyed-remove
'@sub-title Remove the item held under a key, when there is one.
Private Sub KeyedRemove(ByVal target As Collection, ByVal key As String)
  If target Is Nothing Then Exit Sub
  If LenB(key) = 0 Then Exit Sub

  'A key that is not there raises 5, and there is nothing to remove.
  On Error Resume Next
    target.Remove key
  On Error GoTo 0
End Sub

'@label:keyed-has
'@sub-title Report whether a key holds an item.
Private Function KeyedHas(ByVal source As Collection, ByVal key As String) As Boolean
  Dim held As Variant

  If source Is Nothing Then Exit Function
  If LenB(key) = 0 Then Exit Function

  On Error Resume Next
    held = source(key)
  On Error GoTo 0

  KeyedHas = Not IsEmpty(held)
End Function

'@label:keyed-text
'@sub-title Read the text held under a key, answering an empty string when there is none.
Private Function KeyedText(ByVal source As Collection, ByVal key As String) As String
  If source Is Nothing Then Exit Function
  If LenB(key) = 0 Then Exit Function

  On Error Resume Next
    KeyedText = CStr(source(key))
  On Error GoTo 0
End Function

'@label:keyed-number
'@sub-title Read the number held under a key, answering zero when there is none.
Private Function KeyedNumber(ByVal source As Collection, ByVal key As String) As Long
  If source Is Nothing Then Exit Function
  If LenB(key) = 0 Then Exit Function

  On Error Resume Next
    KeyedNumber = CLng(source(key))
  On Error GoTo 0
End Function

'@label:throw-error
'@sub-title Raise a ProjectError aligned error from the current class.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errorMessage As String)
  Err.Raise CLng(errNumber), CLASS_NAME, "Error: " & errorMessage
End Sub
