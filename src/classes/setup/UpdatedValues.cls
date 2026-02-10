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
'@Depends(ProjectError, CustomTable, ICustomTable, BetterArray)
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
Option Explicit

Implements IUpdatedValues

'@class UpdatedValues
'@description
'Tracks watched setup columns and exposes update status to events. Maintains
'a registry of ListObjects on a dedicated worksheet, backed by workbook-level
'named ranges, to detect column-level changes and report update flags.


'@section ConstantsAndState
'===============================================================================
'RANGE_PREFIX           Prefix used for workbook-level named ranges.
'LISTOBJECT_PREFIX      Prefix used when creating the update registry table.
'STATUS_DEFAULT         Default status flag stored in the registry (no changes detected).
'STATUS_UPDATED         Updated status flag stored in the registry after a change is detected.
'TUpValues              Internal state cache storing worksheet, registry and identifier references.
Private Const CLASS_NAME As String = "UpdatedValues"
Private Const RANGE_PREFIX As String = "RNG_"
Private Const LISTOBJECT_PREFIX As String = "UpLo_"
Private Const STATUS_DEFAULT As String = "no"
Private Const STATUS_UPDATED As String = "yes"
Private Const HEADER_COLUMN_NAME As String = "colname"
Private Const HEADER_RANGE_NAME As String = "rngname"
Private Const HEADER_STATUS As String = "updated"
Private Const HEADER_TAG As String = "headername"
Private Const DEFAULT_REGISTRY_START_COLUMN As Long = 1
Private Const TAG_TRANSLATE_TEXT As String = "translate as text"
Private Const TAG_TRANSLATE_FORMULA As String = "translate as formula"
Private Const TAG_WATCH_UPDATE As String = "watch for update"
Private Const MAX_TAG_LOOKUP_DEPTH As Long = 3
Private Const REGISTRY_COLUMN_COUNT As Long = 4
Private Const NAMES_TABLE_NAME As String = "__UpLo__Names__"
Private Const NAME_HEADER_SHEET As String = "sheet"
Private Const NAME_HEADER_TABLE As String = "listobject"
Private Const NAME_HEADER_REGISTRY As String = "registry"
Private Const REGISTRY_NAME_SEPARATOR As String = "_"

Private Enum RegistryColumnIndex
  RegistryColumnLabel = 1
  RegistryRangeName = 2
  RegistryStatus = 3
  RegistryHeaderTag = 4
End Enum

Private Enum NameIndexColumn
  NameIndexSheet = 1
  NameIndexTable = 2
  NameIndexRegistry = 3
End Enum

Private Type TUpValues
  sh As Worksheet                     'Worksheet hosting the registry tables
  lastColumn As Long                  'Tracks the last column used by the registry
  lastRegistryName As String          'Tracks the most recently accessed registry ListObject name
  nameCache As Collection             'In-memory cache of registry names keyed by listobject name
End Type

Private this As TUpValues

'@section Factory
'===============================================================================
'@label:create
'@sub-title Build a new watcher bound to the specified registry worksheet.
'@param sh Worksheet. Registry worksheet hosting the watcher tables.
'@param ignoredId Optional String. Reserved for future use (ignored).
'@return IUpdatedValues. Configured watcher instance.
'@export
Public Function Create(ByVal sh As Worksheet, Optional ByVal ignoredId As String = vbNullString) As IUpdatedValues
  If sh Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "Worksheet reference is required to create UpdatedValues"
  End If

  With New UpdatedValues
    Set .Wksh = sh
    Set Create = .Self
  End With
End Function

'@label:self
'@prop-title Current object instance cast to the interface.
'@return IUpdatedValues. The current instance.
Public Property Get Self() As IUpdatedValues
  Set Self = Me
End Property

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
  If sh Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "Worksheet reference cannot be Nothing"
  End If

  Set this.sh = sh
  ResetRegistryCache
  EnsureNameIndexTable
End Property

'@section PrivateAccessors
'===============================================================================
'@label:ensure-worksheet
'@sub-title Guard that the registry worksheet reference is initialised before use.
Private Function EnsureWorksheet() As Worksheet
  If this.sh Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "UpdatedValues worksheet was not initialised"
  End If
  Set EnsureWorksheet = this.sh
End Function

'@label:registry-list-object-for-table
'@sub-title Retrieve or create the registry ListObject that stores watchers for the supplied source table.
Private Function RegistryListObjectForTable(ByVal tableName As String, _
                                            ByVal createWhenMissing As Boolean) As ListObject
  Dim registryName As String
  Dim lo As ListObject
  Dim sh As Worksheet

  registryName = RegistryListNameForTable(tableName, createWhenMissing)
  If NormalizeValue(registryName) = vbNullString Then Exit Function

  Set lo = RegistryListObjectByName(registryName)

  If lo Is Nothing Then
    If Not createWhenMissing Then Exit Function
    Set sh = EnsureWorksheet()
    Set lo = CreateRegistryListObject(sh, registryName)
    EnsureNameIndexEntry tableName, registryName
    ResetRegistryNameCache
  Else
    EnsureRegistryHeaders lo
    If createWhenMissing Then
      EnsureNameIndexEntry tableName, registryName
    End If
  End If

  TrackLastRegistryName registryName
  Set RegistryListObjectForTable = lo
End Function

'@label:registry-list-object-by-name
'@sub-title Locate a registry ListObject on the worksheet by its name.
Private Function RegistryListObjectByName(ByVal listName As String) As ListObject
  Dim sh As Worksheet
  If NormalizeValue(listName) = vbNullString Then Exit Function

  Set sh = EnsureWorksheet()
  On Error Resume Next
    Set RegistryListObjectByName = sh.ListObjects(listName)
  On Error GoTo 0
End Function

'@label:registry-table
'@sub-title Return a CustomTable helper for the specified registry ListObject name or the last accessed one.
Private Property Get RegistryTable(Optional ByVal tableName As String = vbNullString) As ICustomTable
  Dim targetName As String
  Dim lo As ListObject

  targetName = ResolveRegistryListName(tableName)
  If NormalizeValue(targetName) = vbNullString Then Exit Property

  Set lo = RegistryListObjectByName(targetName)
  If lo Is Nothing Then Exit Property

  TrackLastRegistryName targetName
  Set RegistryTable = CustomTable.Create(lo, HEADER_RANGE_NAME)
End Property

'@label:reset-registry-cache
'@sub-title Clear cached references so subsequent calls rebuild registry helpers.
Private Sub ResetRegistryCache()
  this.lastColumn = 0
  this.lastRegistryName = vbNullString
  ResetRegistryNameCache
End Sub

'@label:track-last-registry-name
'@sub-title Remember the most recently accessed registry ListObject name.
Private Sub TrackLastRegistryName(ByVal listName As String)
  If NormalizeValue(listName) = vbNullString Then Exit Sub
  this.lastRegistryName = listName
End Sub

'@label:resolve-registry-list-name
'@sub-title Determine the registry ListObject name to use for the provided identifier or the last accessed registry.
Private Function ResolveRegistryListName(ByVal candidate As String) As String
  Dim normalized As String

  normalized = NormalizeValue(candidate)
  If normalized <> vbNullString Then
    If Left$(candidate, Len(LISTOBJECT_PREFIX)) = LISTOBJECT_PREFIX Then
      ResolveRegistryListName = candidate
    Else
      ResolveRegistryListName = RegistryListNameForTable(candidate, False)
    End If
  ElseIf NormalizeValue(this.lastRegistryName) <> vbNullString Then
    ResolveRegistryListName = this.lastRegistryName
  Else
    ResolveRegistryListName = LastRegistryListName()
  End If
End Function

'@label:last-registry-list-name
'@sub-title Retrieve the most recently added registry ListObject name on the worksheet.
Private Function LastRegistryListName() As String
  Dim registries As Collection
  Dim registryName As Variant

  Set registries = RegistryNameCache()
  If registries Is Nothing Then Exit Function
  If registries.Count = 0 Then Exit Function

  registryName = registries(registries.Count)
  LastRegistryListName = CStr(registryName)
End Function

'@label:registry-list-objects
'@sub-title Enumerate every registry ListObject hosted on the watcher worksheet.
Private Function RegistryListObjects(Optional ByVal preferredName As String = vbNullString) As Collection
  Dim registries As Collection
  Dim registryName As Variant
  Dim lo As ListObject
  Dim nameList As Collection

  Set registries = New Collection

  If NormalizeValue(preferredName) <> vbNullString Then
    Set lo = RegistryListObjectByName(preferredName)
    If Not lo Is Nothing Then
      registries.Add lo
    End If
    Set RegistryListObjects = registries
    Exit Function
  End If

  Set nameList = RegistryNameCache()
  If nameList Is Nothing Then
    Set RegistryListObjects = registries
    Exit Function
  End If

  For Each registryName In nameList
    Set lo = RegistryListObjectByName(CStr(registryName))
    If Not lo Is Nothing Then
      registries.Add lo
    End If
  Next registryName

  Set RegistryListObjects = registries
End Function

'@section NameIndex
'===============================================================================
'@label:registry-list-name-for-table
'@sub-title Resolve the registry ListObject name for the supplied table, creating a new identifier when requested.
Private Function RegistryListNameForTable(ByVal tableName As String, _
                                          ByVal allowCreateEntry As Boolean) As String
  Dim cachedName As String
  Dim computedName As String

  cachedName = LookupRegistryName(tableName)
  If NormalizeValue(cachedName) <> vbNullString Then
    RegistryListNameForTable = cachedName
    Exit Function
  End If

  If Not allowCreateEntry Then Exit Function

  computedName = BuildRegistryListName(tableName)
  RegistryListNameForTable = computedName
End Function

'@label:build-registry-list-name
'@sub-title Compose the registry ListObject name for the supplied table.
Private Function BuildRegistryListName(ByVal tableName As String) As String
  BuildRegistryListName = LISTOBJECT_PREFIX _
                          & NormalizeIdentifier(tableName) _
                          & REGISTRY_NAME_SEPARATOR _
                          & NormalizeIdentifier(EnsureWorksheet().Name)
End Function

'@label:ensure-name-index-table
'@sub-title Guarantee the registry name index ListObject exists on the worksheet.
Private Function EnsureNameIndexTable() As ListObject
  Dim sh As Worksheet
  Dim lo As ListObject
  Dim headerStart As Range
  Dim tableRange As Range
  Dim startColumn As Long

  Set sh = EnsureWorksheet()
  On Error Resume Next
    Set lo = sh.ListObjects(NAMES_TABLE_NAME)
  On Error GoTo 0

  If lo Is Nothing Then
    startColumn = NextAvailableColumn()
    Set headerStart = sh.Cells(1, startColumn)
    headerStart.Value = NAME_HEADER_SHEET
    headerStart.Offset(0, 1).Value = NAME_HEADER_TABLE
    headerStart.Offset(0, 2).Value = NAME_HEADER_REGISTRY

    Set tableRange = sh.Range(headerStart, headerStart.Offset(1, 2))
    Set lo = sh.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    lo.Name = NAMES_TABLE_NAME
    EnsureNameIndexHeaders lo
    lo.DataBodyRange.ClearContents
    this.lastColumn = startColumn + lo.ListColumns.Count - 1
  Else
    EnsureNameIndexHeaders lo
  End If

  Set EnsureNameIndexTable = lo
End Function

'@label:ensure-name-index-headers
'@sub-title Apply the expected headers to the name index ListObject.
Private Sub EnsureNameIndexHeaders(ByVal lo As ListObject)
  lo.ListColumns(NameIndexSheet).Name = NAME_HEADER_SHEET
  lo.ListColumns(NameIndexTable).Name = NAME_HEADER_TABLE
  lo.ListColumns(NameIndexRegistry).Name = NAME_HEADER_REGISTRY
  lo.HeaderRowRange.Cells(1, NameIndexSheet).Value = NAME_HEADER_SHEET
  lo.HeaderRowRange.Cells(1, NameIndexTable).Value = NAME_HEADER_TABLE
  lo.HeaderRowRange.Cells(1, NameIndexRegistry).Value = NAME_HEADER_REGISTRY
End Sub

'@label:ensure-name-index-entry
'@sub-title Insert or refresh the name index entry for the supplied table.
Private Sub EnsureNameIndexEntry(ByVal tableName As String, ByVal registryName As String)
  Dim lo As ListObject
  Dim listRow As ListRow
  Dim rowSheet As String
  Dim rowTable As String
  Dim sheetName As String

  sheetName = EnsureWorksheet().Name
  Set lo = EnsureNameIndexTable()

  For Each listRow In lo.ListRows
    rowSheet = CStr(listRow.Range.Cells(1, NameIndexSheet).Value)
    rowTable = CStr(listRow.Range.Cells(1, NameIndexTable).Value)
    If StrComp(NormalizeValue(rowSheet), NormalizeValue(sheetName), vbBinaryCompare) = 0 _
       And StrComp(NormalizeValue(rowTable), NormalizeValue(tableName), vbBinaryCompare) = 0 Then
      listRow.Range.Cells(1, NameIndexRegistry).Value = registryName
      ResetRegistryNameCache
      Exit Sub
    End If
  Next listRow

  Set listRow = lo.ListRows.Add
  listRow.Range.Cells(1, NameIndexSheet).Value = sheetName
  listRow.Range.Cells(1, NameIndexTable).Value = tableName
  listRow.Range.Cells(1, NameIndexRegistry).Value = registryName
  ResetRegistryNameCache
End Sub

'@label:remove-name-index-entry
'@sub-title Remove the name index entry associated with the supplied registry.
Private Sub RemoveNameIndexEntry(ByVal tableName As String, ByVal registryName As String)
  Dim lo As ListObject
  Dim listRow As ListRow
  Dim sheetName As String

  If NormalizeValue(registryName) = vbNullString Then Exit Sub

  sheetName = EnsureWorksheet().Name
  Set lo = EnsureNameIndexTable()

  For Each listRow In lo.ListRows
    If StrComp(NormalizeValue(CStr(listRow.Range.Cells(1, NameIndexSheet).Value)), NormalizeValue(sheetName), vbBinaryCompare) = 0 _
       And StrComp(NormalizeValue(CStr(listRow.Range.Cells(1, NameIndexTable).Value)), NormalizeValue(tableName), vbBinaryCompare) = 0 _
       And StrComp(NormalizeValue(CStr(listRow.Range.Cells(1, NameIndexRegistry).Value)), NormalizeValue(registryName), vbBinaryCompare) = 0 Then
      listRow.Delete
      Exit For
    End If
  Next listRow

  ResetRegistryNameCache
End Sub

'@label:reset-registry-name-cache
'@sub-title Clear the cached registry name lookup.
Private Sub ResetRegistryNameCache()
  Set this.nameCache = Nothing
End Sub

'@label:registry-name-cache
'@sub-title Load or retrieve the cached registry names for the current worksheet.
Private Function RegistryNameCache() As Collection
  Dim lo As ListObject
  Dim listRow As ListRow
  Dim sheetName As String
  Dim rowSheet As String
  Dim tableKey As String
  Dim registryName As String

  sheetName = EnsureWorksheet().Name

  If this.nameCache Is Nothing Then
    Set this.nameCache = New Collection
    Set lo = EnsureNameIndexTable()
    For Each listRow In lo.ListRows
      rowSheet = CStr(listRow.Range.Cells(1, NameIndexSheet).Value)
      If StrComp(NormalizeValue(rowSheet), NormalizeValue(sheetName), vbBinaryCompare) = 0 Then
        tableKey = NormalizeIdentifier(CStr(listRow.Range.Cells(1, NameIndexTable).Value))
        registryName = CStr(listRow.Range.Cells(1, NameIndexRegistry).Value)
        On Error Resume Next
          this.nameCache.Add registryName, tableKey
        On Error GoTo 0
      End If
    Next listRow
  End If

  Set RegistryNameCache = this.nameCache
End Function

'@label:lookup-registry-name
'@sub-title Retrieve the registry ListObject name associated with the supplied table.
Private Function LookupRegistryName(ByVal tableName As String) As String
  Dim cache As Collection
  Dim key As String

  key = NormalizeIdentifier(tableName)
  Set cache = RegistryNameCache()

  On Error Resume Next
    LookupRegistryName = CStr(cache(key))
  On Error GoTo 0
End Function

'@label:lookup-table-name-from-registry
'@sub-title Attempt to determine the original table name from a registry ListObject.
Private Function LookupTableNameFromRegistry(ByVal registryName As String) As String
  Dim lo As ListObject
  Dim listRow As ListRow

  If NormalizeValue(registryName) = vbNullString Then Exit Function

  Set lo = EnsureNameIndexTable()
  For Each listRow In lo.ListRows
    If StrComp(NormalizeValue(CStr(listRow.Range.Cells(1, NameIndexRegistry).Value)), NormalizeValue(registryName), vbBinaryCompare) = 0 Then
      LookupTableNameFromRegistry = CStr(listRow.Range.Cells(1, NameIndexTable).Value)
      Exit Function
    End If
  Next listRow
End Function

'@label:clear-name-index-entries
'@sub-title Remove every entry from the registry name index.
Private Sub ClearNameIndexEntries()
  Dim lo As ListObject

  Set lo = EnsureNameIndexTable()
  Do While lo.ListRows.Count > 0
    lo.ListRows(lo.ListRows.Count).Delete
  Loop
  ResetRegistryNameCache
End Sub

'@label:clamp-start-column
'@sub-title Ensure a proposed registry column start fits within the worksheet boundaries.
Private Function ClampStartColumn(ByVal startColumn As Long, _
                                  Optional ByVal targetSh As Worksheet) As Long
  Dim effectiveSheet As Worksheet
  Dim maxStart As Long

  If targetSh Is Nothing Then
    Set effectiveSheet = EnsureWorksheet()
  Else
    Set effectiveSheet = targetSh
  End If

  If startColumn < DEFAULT_REGISTRY_START_COLUMN Then
    startColumn = DEFAULT_REGISTRY_START_COLUMN
  End If

  maxStart = effectiveSheet.Columns.Count - REGISTRY_COLUMN_COUNT + 1
  If maxStart < DEFAULT_REGISTRY_START_COLUMN Then
    maxStart = DEFAULT_REGISTRY_START_COLUMN
  End If

  If startColumn > maxStart Then
    ClampStartColumn = maxStart
  Else
    ClampStartColumn = startColumn
  End If
End Function

'@label:worksheet-last-used-column
'@sub-title Determine the last used column on the registry worksheet using End(xlToLeft).
Private Function WorksheetLastUsedColumn(ByVal sh As Worksheet) As Long
  Dim lastCell As Range
  Dim lastCol As Long

  On Error Resume Next
    Set lastCell = sh.Cells(1, sh.Columns.Count).End(xlToLeft)
  On Error GoTo 0

  If lastCell Is Nothing Then Exit Function

  If lastCell.Column = 1 And LenB(CStr(lastCell.Value)) = 0 Then
    WorksheetLastUsedColumn = 0
  Else
    WorksheetLastUsedColumn = lastCell.Column
  End If
End Function

'@label:next-available-column
'@sub-title Determine the next column available for placing the registry table.
Private Function NextAvailableColumn() As Long
  Dim sh As Worksheet
  Dim sheetLast As Long
  Dim candidate As Long

  Set sh = EnsureWorksheet()

  sheetLast = WorksheetLastUsedColumn(sh)

  If sheetLast <> this.lastColumn Then
    this.lastColumn = sheetLast
  End If

  candidate = IIf(this.lastColumn = 1, 1, this.lastColumn + 2)
  candidate = ClampStartColumn(candidate, sh)

  NextAvailableColumn = candidate
End Function

'@label:create-registry-list-object
'@sub-title Build an empty registry ListObject with the expected headers.
Private Function CreateRegistryListObject(ByVal sh As Worksheet, ByVal loName As String) As ListObject
  Dim startColumn As Long
  Dim headerStart As Range
  Dim tableRange As Range

  startColumn = NextAvailableColumn()
  Set headerStart = sh.Cells(1, startColumn)

  headerStart.Value = HEADER_COLUMN_NAME
  headerStart.Offset(0, 1).Value = HEADER_RANGE_NAME
  headerStart.Offset(0, 2).Value = HEADER_STATUS
  headerStart.Offset(0, 3).Value = HEADER_TAG

  Set tableRange = sh.Range(headerStart, headerStart.Offset(1, REGISTRY_COLUMN_COUNT - 1))

  Set CreateRegistryListObject = sh.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
  CreateRegistryListObject.Name = loName
  EnsureRegistryHeaders CreateRegistryListObject
  this.lastColumn = headerStart.Column + REGISTRY_COLUMN_COUNT - 1
End Function

'@label:ensure-registry-headers
'@sub-title Guarantee the registry ListObject exposes the expected header names.
Private Sub EnsureRegistryHeaders(ByVal lo As ListObject)
  Dim expected As BetterArray
  Dim idx As Long

  Set expected = New BetterArray
  expected.Push HEADER_COLUMN_NAME, HEADER_RANGE_NAME, HEADER_STATUS, HEADER_TAG

  Do While lo.ListColumns.Count < REGISTRY_COLUMN_COUNT
    lo.ListColumns.Add
  Loop

  For idx = expected.LowerBound To expected.UpperBound
    lo.ListColumns(idx + 1).Name = expected(idx)
    lo.HeaderRowRange.Cells(1, idx + 1).Value = expected.Item(idx)
  Next idx
End Sub

'@section ColumnRegistration
'===============================================================================
'@label:add-columns
'@sub-title Register or refresh watched columns for the supplied ListObject.
'@details
'Scans the source ListObject header row for metadata tags and synchronises
'the registry table so named ranges remain aligned with the current layout.
'@param source ListObject. Table to inspect for update tags.
Private Sub AddColumns(ByVal source As ListObject)
  Dim headerCell As Range
  Dim expected As BetterArray
  Dim registryIndex As BetterArray
  Dim registry As ListObject
  Dim normalizedRange As String
  Dim columnLabel As String
  Dim tagValue As String
  Dim tableKey As String
  Dim rowIndex As Long
  Dim isNew As Boolean

  If source Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "ListObject reference is required when registering updated values"
  End If

  If source.HeaderRowRange Is Nothing Then Exit Sub

  Set expected = New BetterArray
  expected.LowerBound = 1

  tableKey = source.Name
  Set registry = RegistryListObjectForTable(tableKey, createWhenMissing:=True)
  If registry Is Nothing Then Exit Sub

  Set registryIndex = BuildRegistryIndex(registry)

  For Each headerCell In source.HeaderRowRange.Cells
    tagValue = ColumnTag(headerCell)
    If ShouldWatchColumn(tagValue) Then
      normalizedRange = BuildRangeName(tableKey, NormalizeColumnKey(headerCell.Value))
      columnLabel = BuildColumnLabel(source.Name, headerCell.Value)
      rowIndex = EnsureRegistryRowFromIndex(registry, registryIndex, normalizedRange, isNew)
      PopulateRegistryRow registry, rowIndex, columnLabel, normalizedRange, tagValue, isNew
      EnsureNamedRange source, headerCell.Value, normalizedRange
      AddExpectedKey expected, normalizedRange
    End If
  Next headerCell

  PruneObsoleteEntries registry, expected, tableKey
End Sub

'@label:add-sheet
'@sub-title Register tagged columns for every ListObject on the provided worksheet.
Private Sub AddSheet(ByVal targetSheet As Worksheet)
  Dim lo As ListObject

  If targetSheet Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "Worksheet reference is required when registering all updated values"
  End If

  If targetSheet.ListObjects.Count = 0 Then Exit Sub

  For Each lo In targetSheet.ListObjects
    AddColumns lo
  Next lo
End Sub

'@label:remove-lo
'@sub-title Remove registry entries associated with the supplied ListObject.
Private Sub RemoveListObject(ByVal target As ListObject)
  Dim registry As ListObject

  If target Is Nothing Then Exit Sub

  Set registry = RegistryListObjectForTable(target.Name, createWhenMissing:=False)
  If registry Is Nothing Then Exit Sub

  RemoveRegistryTable registry, target.Name
End Sub

'@label:column-label-matches-table
'@sub-title Determine whether a registry column label belongs to the provided ListObject name.
Private Function ColumnLabelMatchesTable(ByVal columnLabel As String, _
                                         ByVal tableName As String) As Boolean
  Dim expectedPrefix As String

  expectedPrefix = tableName & "-"
  ColumnLabelMatchesTable = (StrComp(Left$(columnLabel, Len(expectedPrefix)), expectedPrefix, vbTextCompare) = 0)
End Function

'@label:column-tag
'@sub-title Read the metadata tag associated with a column header, searching above the header row.
'@param headerCell Range. The target header cell.
Private Function ColumnTag(ByVal headerCell As Range) As String
  Dim depth As Long
  Dim candidate As Range
  Dim valueText As String

  If headerCell Is Nothing Then Exit Function

  For depth = 1 To MAX_TAG_LOOKUP_DEPTH
    If headerCell.Row <= depth Then Exit For

    Set candidate = headerCell.Offset(-depth, 0)
    valueText = NormalizeValue(CStr(candidate.Value))
    If valueText <> vbNullString Then
      ColumnTag = valueText
      Exit Function
    End If
  Next depth
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

'@label:build-registry-index
'@sub-title Create a lightweight lookup of existing registry rows keyed by normalised range name.
Private Function BuildRegistryIndex(ByVal registry As ListObject) As BetterArray
  Dim index As BetterArray
  Dim listRow As ListRow
  Dim rangeName As String
  Dim key As String
  Dim entry As Variant

  Set index = New BetterArray
  index.LowerBound = 1

  If registry Is Nothing Then
    Set BuildRegistryIndex = index
    Exit Function
  End If

  For Each listRow In registry.ListRows
    rangeName = CStr(listRow.Range.Cells(1, RegistryRangeName).Value)
    key = NormalizeValue(rangeName)
    If key <> vbNullString Then
      If FindRegistryRowInIndex(index, key) = 0 Then
        entry = Array(key, listRow.Index)
        index.Push entry
      End If
    End If
  Next listRow

  Set BuildRegistryIndex = index
End Function

'@label:ensure-registry-row-from-index
'@sub-title Locate or append a registry row for the provided range name.
'@param registryIndex BetterArray. Index built from existing entries.
'@param rangeName String. Normalised range identifier.
'@param isNew Boolean. Output flag indicating whether a new row was added.
Private Function EnsureRegistryRowFromIndex(ByVal registry As ListObject, _
                                            ByVal registryIndex As BetterArray, _
                                            ByVal rangeName As String, _
                                            ByRef isNew As Boolean) As Long
  Dim key As String
  Dim existingRowIndex As Long
  Dim listRow As ListRow
  Dim entry As Variant

  key = NormalizeValue(rangeName)
  existingRowIndex = FindRegistryRowInIndex(registryIndex, key)
  isNew = False

  If existingRowIndex = 0 Then
    Set listRow = registry.ListRows.Add
    EnsureRegistryRowFromIndex = listRow.Index
    entry = Array(key, listRow.Index)
    registryIndex.Push entry
    isNew = True
  Else
    EnsureRegistryRowFromIndex = existingRowIndex
  End If
End Function

'@label:find-registry-row-in-index
'@sub-title Locate a registry row index in the cached index by its normalised key.
Private Function FindRegistryRowInIndex(ByVal registryIndex As BetterArray, _
                                        ByVal key As String) As Long
  Dim idx As Long
  Dim entry As Variant

  If registryIndex Is Nothing Then Exit Function
  If registryIndex.Length = 0 Then Exit Function

  For idx = registryIndex.LowerBound To registryIndex.UpperBound
    entry = registryIndex.Item(idx)
    If IsArray(entry) Then
      If StrComp(CStr(entry(LBound(entry))), key, vbBinaryCompare) = 0 Then
        If UBound(entry) >= 1 Then
          FindRegistryRowInIndex = CLng(entry(UBound(entry)))
        End If
        Exit Function
      End If
    End If
  Next idx
End Function

'@label:populate-registry-row
'@sub-title Write watcher metadata to the specified registry row and initialise status flags.
Private Sub PopulateRegistryRow(ByVal registry As ListObject, _
                                ByVal rowIndex As Long, _
                                ByVal columnLabel As String, _
                                ByVal rangeName As String, _
                                ByVal tagValue As String, _
                                ByVal isNew As Boolean)
  Dim rowRange As Range

  If rowIndex <= 0 Then Exit Sub
  If registry Is Nothing Then Exit Sub
  If rowIndex < 1 Or rowIndex > registry.ListRows.Count Then Exit Sub

  Set rowRange = registry.ListRows(rowIndex).Range

  rowRange.Cells(1, RegistryColumnLabel).Value = columnLabel
  rowRange.Cells(1, RegistryRangeName).Value = rangeName

  If isNew Or NormalizeValue(CStr(rowRange.Cells(1, RegistryStatus).Value)) = vbNullString Then
    rowRange.Cells(1, RegistryStatus).Value = STATUS_DEFAULT
  End If

  rowRange.Cells(1, RegistryHeaderTag).Value = tagValue
End Sub

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

'@label:add-expected-key
'@sub-title Track range names detected during scanning to support registry pruning.
Private Sub AddExpectedKey(ByVal expected As BetterArray, ByVal rangeName As String)
  Dim key As String

  If expected Is Nothing Then Exit Sub

  key = NormalizeValue(rangeName)
  If key = vbNullString Then Exit Sub

  If expected.Includes(key) Then Exit Sub
  expected.Push key
End Sub

'@label:expected-contains
'@sub-title Test whether the expected list already includes the provided range name.
Private Function ExpectedContains(ByVal expected As BetterArray, ByVal rangeName As String) As Boolean
  Dim key As String

  ExpectedContains = False
  If expected Is Nothing Then Exit Function
  If expected.Length = 0 Then Exit Function

  key = NormalizeValue(rangeName)
  If key = vbNullString Then Exit Function

  ExpectedContains = expected.Includes(key)
End Function

'@label:prune-obsolete-entries
'@sub-title Remove registry rows and named ranges for columns no longer tagged for watching.
Private Sub PruneObsoleteEntries(ByVal registry As ListObject, _
                                 ByVal expected As BetterArray, _
                                 ByVal tableName As String)
  Dim rowIndex As Long
  Dim rangeName As String

  If registry Is Nothing Then Exit Sub

  If expected Is Nothing Then
    RemoveRegistryTable registry, tableName
    Exit Sub
  End If

  If expected.Length = 0 Then
    RemoveRegistryTable registry, tableName
    Exit Sub
  End If

  For rowIndex = registry.ListRows.Count To 1 Step -1
    rangeName = CStr(registry.ListRows(rowIndex).Range.Cells(1, RegistryRangeName).Value)
    If NormalizeValue(rangeName) = vbNullString Then
      registry.ListRows(rowIndex).Delete
      GoTo ContinueLoop
    End If

    If Not ExpectedContains(expected, rangeName) Then
      RemoveNamedRange rangeName
      registry.ListRows(rowIndex).Delete
    End If
ContinueLoop:
  Next rowIndex

  If registry.ListRows.Count = 0 Then
    RemoveRegistryTable registry, tableName
  End If
End Sub

'@section StatusTracking
'===============================================================================
'@label:is-updated
'@prop-title Determine whether the tracked column has been marked as updated.
'@details
'Searches across all registry ListObjects for a matching entry and returns
'True when the column is flagged as updated since the last reset.
Private Property Get IsUpdated(ByVal checkColName As String) As Boolean
  Dim registries As Collection
  Dim registry As ListObject
  Dim rangeColumn As Range
  Dim statusColumn As Range
  Dim labelColumn As Range
  Dim idx As Long
  Dim targetKey As String
  Dim columnLabel As String
  Dim rangeName As String
  Dim statusValue As String

  targetKey = NormalizeValue(checkColName)
  If targetKey = vbNullString Then Exit Property

  Set registries = RegistryListObjects()
  If registries Is Nothing Then Exit Property

  For Each registry In registries
    Set rangeColumn = RegistryColumnData(registry, HEADER_RANGE_NAME)
    Set statusColumn = RegistryColumnData(registry, HEADER_STATUS)
    Set labelColumn = RegistryColumnData(registry, HEADER_COLUMN_NAME)

    If rangeColumn Is Nothing Then GoTo NextRegistry
    If statusColumn Is Nothing Then GoTo NextRegistry
    If labelColumn Is Nothing Then GoTo NextRegistry

    For idx = 1 To rangeColumn.Rows.Count
      rangeName = CStr(rangeColumn.Cells(idx, 1).Value)
      columnLabel = CStr(labelColumn.Cells(idx, 1).Value)
      statusValue = NormalizeValue(CStr(statusColumn.Cells(idx, 1).Value))

      If ColumnMatchesIdentifier(columnLabel, rangeName, targetKey) Then
        IsUpdated = (statusValue = STATUS_UPDATED)
        Exit Property
      End If
    Next idx

NextRegistry:
  Next registry
End Property

'@label:check-update
'@sub-title Evaluate a range change against the registered watchers and flag updates.
'@details
'Compares the changed range with named ranges tracked in the registry and
'marks the first matching entry as updated.
'@param tagReference Variant. Identifier selecting the registry scope.
'@param Target Range. The changed cells to evaluate.
Private Sub CheckStatus(ByVal tagReference As Variant, ByVal Target As Range)
  Dim registries As Collection
  Dim registry As ListObject
  Dim rangeColumn As Range
  Dim statusColumn As Range
  Dim idx As Long
  Dim rangeName As String
  Dim statusCell As Range
  Dim namedRange As Range
  Dim registryName As String
  Dim rowRemoved As Boolean
  Dim newCount As Long
  Dim registryTag As String

  If Target Is Nothing Then Exit Sub

  registryTag = RegistryTagFromVariant(tagReference)
  If NormalizeValue(registryTag) <> vbNullString Then
    registryName = ResolveRegistryListName(registryTag)
  Else
    registryName = vbNullString
  End If

  Set registries = RegistryListObjects(registryName)
  If registries Is Nothing Then Exit Sub
  If registries.Count = 0 Then Exit Sub

  For Each registry In registries
    Set rangeColumn = RegistryColumnData(registry, HEADER_RANGE_NAME)
    Set statusColumn = RegistryColumnData(registry, HEADER_STATUS)

    If rangeColumn Is Nothing Then GoTo NextRegistry
    If statusColumn Is Nothing Then GoTo NextRegistry

    idx = RowCount(rangeColumn)

    Do While idx >= 1
      rangeName = CStr(rangeColumn.Cells(idx, 1).Value)
      Set statusCell = statusColumn.Cells(idx, 1)
      rowRemoved = False

      If NormalizeValue(rangeName) = vbNullString Then
        rowRemoved = True
        If Not RemoveOutdatedEntry(registry, rangeColumn.Cells(idx, 1)) Then
          Exit Do
        End If
        GoTo NextIteration
      End If

      If NormalizeValue(CStr(statusCell.Value)) = STATUS_UPDATED Then
        GoTo NextIteration
      End If

      Set namedRange = ResolveNamedRange(rangeName)
      If namedRange Is Nothing Then
        rowRemoved = True
        If Not RemoveOutdatedEntry(registry, rangeColumn.Cells(idx, 1)) Then
          Exit Do
        End If
        GoTo NextIteration
      End If

      If Intersect(Target, namedRange) Is Nothing Then
        GoTo NextIteration
      End If

      statusCell.Value = STATUS_UPDATED
      TrackLastRegistryName registry.Name
      Exit Sub

NextIteration:
      If Not rowRemoved Then
        idx = idx - 1
      Else
        Set rangeColumn = RegistryColumnData(registry, HEADER_RANGE_NAME)
        Set statusColumn = RegistryColumnData(registry, HEADER_STATUS)
        newCount = RowCount(rangeColumn)
        If newCount = 0 Then Exit Do
        If idx - 1 > newCount Then
          idx = newCount
        Else
          idx = idx - 1
        End If
      End If
    Loop

NextRegistry:
  Next registry
End Sub

'@label:resolve-named-range
'@sub-title Attempt to resolve the workbook defined name into a concrete Range reference.
Private Function ResolveNamedRange(ByVal rangeName As String) As Range
  Dim wb As Workbook
  Dim definedName As Name

  If NormalizeValue(rangeName) = vbNullString Then Exit Function

  Set wb = EnsureWorksheet().Parent

  On Error Resume Next
    Set definedName = wb.Names(rangeName)
    If Not definedName Is Nothing Then
      Set ResolveNamedRange = definedName.RefersToRange
    End If
  On Error GoTo 0
End Function

'@label:update-status-for-cell
'@sub-title Write the supplied status flag into the registry row matching the range cell.
Private Sub UpdateStatusForCell(ByVal rangeCell As Range, ByVal statusValue As String)
  Dim statusCell As Range

  If rangeCell Is Nothing Then Exit Sub

  Set statusCell = rangeCell.Offset(0, RegistryStatus - RegistryRangeName)
  statusCell.Value = statusValue
End Sub

'@label:remove-outdated-entry
'@sub-title Clean registry rows that point to missing or invalid named ranges.
Private Function RemoveOutdatedEntry(ByVal registry As ListObject, _
                                     ByVal rangeCell As Range) As Boolean
  Dim rowIndex As Long
  Dim rangeName As String
  Dim tableName As String

  If rangeCell Is Nothing Then Exit Function
  If registry Is Nothing Then Exit Function

  rowIndex = rangeCell.Row - registry.HeaderRowRange.Row

  If rowIndex < 1 Or rowIndex > registry.ListRows.Count Then Exit Function

  rangeName = CStr(rangeCell.Value)
  RemoveNamedRange rangeName
  registry.ListRows(rowIndex).Delete

  If registry.ListRows.Count = 0 Then
    tableName = LookupTableNameFromRegistry(registry.Name)
    RemoveRegistryTable registry, tableName
     Exit Function
  End If

  RemoveOutdatedEntry = True
End Function

'@label:reset-update-status
'@sub-title Reset every registry row to the default (non-updated) status value.
Private Sub ResetUpdateStatus()
  Dim statusColumn As Range
  Dim cell As Range
  Dim registries As Collection
  Dim registry As ListObject

  Set registries = RegistryListObjects()
  If registries Is Nothing Then Exit Sub

  For Each registry In registries
    Set statusColumn = RegistryColumnData(registry, HEADER_STATUS)
    If Not statusColumn Is Nothing Then
      For Each cell In statusColumn.Cells
        cell.Value = STATUS_DEFAULT
      Next cell
    End If
  Next registry
End Sub

'@section CleanupHelpers
'===============================================================================
'@label:clear-up
'@sub-title Reset all tracked columns to the default (not updated) status.
Private Sub ClearUp()
  ResetUpdateStatus
End Sub

'@label:delete-up
'@sub-title Remove the watcher registry and associated workbook named ranges.
Private Sub DeleteUp()
  Dim registries As Collection
  Dim registry As ListObject
  Dim sheetLast As Long
  Dim idx As Long

  Set registries = RegistryListObjects()
  If registries Is Nothing Then GoTo ResetPlacement

  For idx = registries.Count To 1 Step -1
    Set registry = registries(idx)
    RemoveRegistryTable registry
  Next idx

  ClearNameIndexEntries

ResetPlacement:
  ResetRegistryCache
  sheetLast = WorksheetLastUsedColumn(EnsureWorksheet())
  this.lastColumn = sheetLast
End Sub

'@label:remove-registry-table
'@sub-title Remove an entire registry ListObject and associated named ranges from the worksheet.
Private Sub RemoveRegistryTable(ByVal registry As ListObject, _
                                Optional ByVal tableName As String = vbNullString)
  Dim registryRange As Range
  Dim rowIndex As Long
  Dim rangeName As String
  Dim registryName As String
  Dim resolvedTable As String

  If registry Is Nothing Then Exit Sub

  registryName = registry.Name
  For rowIndex = registry.ListRows.Count To 1 Step -1
    rangeName = CStr(registry.ListRows(rowIndex).Range.Cells(1, RegistryRangeName).Value)
    RemoveNamedRange rangeName
  Next rowIndex

  resolvedTable = tableName
  If NormalizeValue(resolvedTable) = vbNullString Then
    resolvedTable = LookupTableNameFromRegistry(registryName)
  End If
  If NormalizeValue(resolvedTable) <> vbNullString Then
    RemoveNameIndexEntry resolvedTable, registryName
  End If

  Set registryRange = registry.Range
  registry.Unlist
  registryRange.Cells.Clear

  this.lastColumn = WorksheetLastUsedColumn(EnsureWorksheet())
  If StrComp(this.lastRegistryName, registryName, vbBinaryCompare) = 0 Then
    this.lastRegistryName = vbNullString
  End If
End Sub

'@label:remove-all-named-ranges
'@sub-title Delete all workbook defined names created for watcher columns.
Private Sub RemoveAllNamedRanges()
  Dim registries As Collection
  Dim registry As ListObject
  Dim idx As Long
  Dim rowIndex As Long
  Dim rangeName As String

  Set registries = RegistryListObjects()
  If registries Is Nothing Then Exit Sub

  For idx = 1 To registries.Count
    Set registry = registries(idx)
    For rowIndex = 1 To registry.ListRows.Count
      rangeName = CStr(registry.ListRows(rowIndex).Range.Cells(1, RegistryRangeName).Value)
      RemoveNamedRange rangeName
    Next rowIndex
  Next idx
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

'@label:switch-tags
'@sub-title Set every registry status flag to the supplied value.
Private Sub SwitchTags(ByVal value As String)
    Dim registries As Collection
    Dim registry As ListObject
    Dim rng As Range
    Dim counter As Long

    Set registries = RegistryListObjects()
    If registries Is Nothing Then Exit Sub

    For Each registry In registries
      Set rng = RegistryColumnData(registry, HEADER_STATUS)
      If Not rng Is Nothing Then
        For counter = 1 To rng.Rows.Count
          If NormalizeValue(CStr(rng.Cells(counter, 1).Value)) <> NormalizeValue(value) Then
            rng.Cells(counter, 1).Value = value
          End If
        Next counter
      End If
    Next registry
End Sub

'@section Utilities
'===============================================================================
'@label:registry-column-data
'@sub-title Return the DataBodyRange for the specified registry column if available.
Private Function RegistryColumnData(ByVal registry As ListObject, _
                                    ByVal headerName As String) As Range
  Dim listColumn As ListColumn

  If registry Is Nothing Then Exit Function

  On Error Resume Next
    Set listColumn = registry.ListColumns(headerName)
  On Error GoTo 0

  If listColumn Is Nothing Then Exit Function
  Set RegistryColumnData = listColumn.DataBodyRange
End Function

'@label:registry-tag-from-variant
'@sub-title Normalise a tag reference into a registry identifier string.
Private Function RegistryTagFromVariant(ByVal tagReference As Variant) As String
  If IsObject(tagReference) Then
    Select Case TypeName(tagReference)
      Case "ListObject"
        RegistryTagFromVariant = CStr(tagReference.Name)
      Case "Worksheet"
        RegistryTagFromVariant = vbNullString
      Case Else
        RegistryTagFromVariant = vbNullString
    End Select
  Else
    On Error Resume Next
      RegistryTagFromVariant = CStr(tagReference)
    On Error GoTo 0
  End If
End Function

'@label:row-count
'@sub-title Return the number of rows contained in the supplied range.
Private Function RowCount(ByVal target As Range) As Long
  If target Is Nothing Then Exit Function
  RowCount = target.Rows.Count
End Function

'@label:current-lo-name
'@sub-title Compose the ListObject name used when storing the registry table.
'@param tableName String. Source ListObject name associated with the registry.
Private Function CurrentLoName(ByVal tableName As String) As String
  CurrentLoName = BuildRegistryListName(tableName)
End Function

'@label:build-column-label
'@sub-title Compose the descriptive label stored in the registry for a watched column.
Private Function BuildColumnLabel(ByVal tableName As String, ByVal columnName As String) As String
  BuildColumnLabel = tableName & "-" & columnName
End Function

'@label:extract-column-name-from-label
'@sub-title Retrieve the column portion from a table-column label.
Private Function ExtractColumnNameFromLabel(ByVal columnLabel As String) As String
  Dim separatorPos As Long
  Dim extracted As String

  separatorPos = InStr(columnLabel, "-")
  If separatorPos = 0 Then
    extracted = columnLabel
  Else
    extracted = Mid$(columnLabel, separatorPos + 1)
  End If

  ExtractColumnNameFromLabel = Trim$(extracted)
End Function

'@label:column-matches-identifier
'@sub-title Determine whether the supplied registry label or range name matches the target identifier.
Private Function ColumnMatchesIdentifier(ByVal columnLabel As String, _
                                         ByVal rangeName As String, _
                                         ByVal targetKey As String) As Boolean
  Dim normalizedLabel As String
  Dim normalizedColumn As String
  Dim normalizedRange As String

  normalizedLabel = NormalizeValue(columnLabel)
  normalizedColumn = NormalizeValue(ExtractColumnNameFromLabel(columnLabel))
  normalizedRange = NormalizeValue(rangeName)

  ColumnMatchesIdentifier = (normalizedLabel = targetKey) _
                            Or (normalizedColumn = targetKey) _
                            Or (normalizedRange = targetKey)
End Function

'@label:build-range-name
'@sub-title Generate the workbook defined name for a watched column, including table context.
Private Function BuildRangeName(ByVal tableKey As String, ByVal normalizedColumnKey As String) As String
  Dim normalizedTable As String
  Dim normalizedSheet As String

  normalizedTable = NormalizeColumnKey(tableKey)
  If NormalizeValue(normalizedTable) = vbNullString Then
    normalizedTable = "table"
  End If
  normalizedSheet = NormalizeIdentifier(EnsureWorksheet().Name)

  BuildRangeName = RANGE_PREFIX & normalizedTable & "_" & normalizedColumnKey & "_" & normalizedSheet
End Function

'@label:normalize-column-key
'@sub-title Produce a workbook-friendly identifier from a column header.
Private Function NormalizeColumnKey(ByVal columnName As String) As String
  NormalizeColumnKey = NormalizeIdentifier(columnName)
End Function

'@label:normalize-identifier
'@sub-title Produce a safe workbook-level identifier from free-form text.
Private Function NormalizeIdentifier(ByVal valueText As String) As String
  Dim idx As Long
  Dim ch As String
  Dim buffer As String

  valueText = Replace(valueText, Chr$(160), " ")
  valueText = Trim$(valueText)

  For idx = 1 To Len(valueText)
    ch = Mid$(valueText, idx, 1)

    Select Case ch
      Case "A" To "Z", "a" To "z", "0" To "9"
        buffer = buffer & LCase$(ch)
      Case "_"
        buffer = buffer & "_"
      Case Else
        buffer = buffer & "_"
    End Select
  Next idx

  buffer = ReplaceRepeatedUnderscores(buffer)
  buffer = TrimUnderscores(buffer)

  If LenB(buffer) = 0 Then buffer = "field"

  NormalizeIdentifier = buffer
End Function

'@label:replace-repeated-underscores
'@sub-title Collapse multiple underscores into a single instance.
Private Function ReplaceRepeatedUnderscores(ByVal valueText As String) As String
  Do While InStr(valueText, "__") > 0
    valueText = Replace(valueText, "__", "_")
  Loop
  ReplaceRepeatedUnderscores = valueText
End Function

'@label:trim-underscores
'@sub-title Remove leading and trailing underscore characters.
Private Function TrimUnderscores(ByVal valueText As String) As String
  Do While Len(valueText) > 0 And Left$(valueText, 1) = "_"
    valueText = Mid$(valueText, 2)
  Loop

  Do While Len(valueText) > 0 And Right$(valueText, 1) = "_"
    valueText = Left$(valueText, Len(valueText) - 1)
  Loop

  TrimUnderscores = valueText
End Function

'@label:normalize-value
'@sub-title Normalise text by trimming spaces and applying lower case.
Private Function NormalizeValue(ByVal valueText As String) As String
  NormalizeValue = LCase$(Trim$(valueText))
End Function


'@label:throw-error
'@sub-title Raise a ProjectError aligned error from the current class.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errorMessage As String)
  Err.Raise CLng(errNumber), CLASS_NAME, "Error: " & errorMessage
End Sub

'@section Interface implementation
'@description Thin delegation stubs forwarding each IUpdatedValues member
'to the corresponding internal implementation above.
'===============================================================================
Private Sub IUpdatedValues_AddColumns(ByVal Lo As ListObject)
  AddColumns Lo
End Sub

Private Sub IUpdatedValues_AddSheet(ByVal sh As Worksheet)
  AddSheet sh
End Sub

Private Sub IUpdatedValues_RemoveLo(ByVal Lo As ListObject)
  RemoveListObject Lo
End Sub

Private Property Get IUpdatedValues_IsUpdated(ByVal colName As String) As Boolean
  IUpdatedValues_IsUpdated = IsUpdated(colName)
End Property

Private Sub IUpdatedValues_CheckUpdate(ByVal tagReference As Variant, ByVal Target As Range)
  CheckStatus tagReference, Target
End Sub

Private Sub IUpdatedValues_ClearUp()
  ClearUp
End Sub

Private Sub IUpdatedValues_DeleteUp()
  DeleteUp
End Sub

'@label:switch-tags-to-no
Private Sub IUpdatedValues_SwitchTagsToNo()
  SwitchTags STATUS_DEFAULT
End Sub

'@label:switch-tags-to-yes
Private Sub IUpdatedValues_SwitchTagsToYes()
  SwitchTags STATUS_UPDATED
End Sub

'@label:exists
'@sub-title Determine whether a registry ListObject already exists for the supplied identifier.
'@param listName String. Table or registry identifier to look up.
'@return Boolean. True when a matching registry ListObject exists.
'@export
Public Function Exists(ByVal listName As String) As Boolean
  Dim sh As Worksheet
  Dim loName As String
  Dim lo As ListObject

  Set sh = EnsureWorksheet()
  loName = ResolveRegistryListName(listName)
  If NormalizeValue(loName) = vbNullString Then Exit Function

  On Error Resume Next
    Set lo = sh.ListObjects(loName)
  On Error GoTo 0

  Exists = Not lo Is Nothing
End Function
