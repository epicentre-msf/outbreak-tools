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

'@section ConstantsAndState
'===============================================================================
'@constant RANGE_PREFIX Prefix used for workbook-level named ranges.
'@constant LISTOBJECT_PREFIX Prefix used when creating the update registry table.
'@constant STATUS_DEFAULT Default status flag stored in the registry (no changes detected).
'@constant STATUS_UPDATED Updated status flag stored in the registry after a change is detected.
'@type TUpValues Internal state cache storing worksheet, registry and identifier references.
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

Private Enum RegistryColumnIndex
  RegistryColumnLabel = 1
  RegistryRangeName = 2
  RegistryStatus = 3
  RegistryHeaderTag = 4
End Enum

Private Type TUpValues
  registryTable As ICustomTable   'Cached helper for interacting with the registry ListObject
  registryLo As ListObject        'Cached ListObject that stores watched columns
  sh As Worksheet                 'Worksheet hosting the registry table
  idValue As String               'Identifier suffix used for registry naming (e.g., UpLo_dict)
End Type

Private this As TUpValues

'@section Factory
'===============================================================================
'@method Create
'@description Build a new watcher bound to the specified registry worksheet and identifier.
Public Function Create(ByVal sh As Worksheet, ByVal upName As String) As IUpdatedValues
  If sh Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "Worksheet reference is required to create UpdatedValues"
  End If

  If NormalizeValue(upName) = vbNullString Then
    ThrowError ProjectError.InvalidArgument, "Update identifier cannot be empty"
  End If

  With New UpdatedValues
    Set .Wksh = sh
    .ID = upName
    Set Create = .Self
  End With
End Function

'@label:Self
'@description Return the current object instance implementing the interface.
Public Property Get Self() As IUpdatedValues
  Set Self = Me
End Property

'@section PublicState
'===============================================================================
'@property Wksh
'@description Worksheet hosting the watcher registry table.
Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  If sh Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "Worksheet reference cannot be Nothing"
  End If

  Set this.sh = sh
  ResetRegistryCache
End Property

'@property ID
'@description Identifier suffix used for named ranges and registry ListObject.
Public Property Get ID() As String
  ID = this.idValue
End Property

Public Property Let ID(ByVal upName As String)
  If NormalizeValue(upName) = vbNullString Then
    ThrowError ProjectError.InvalidArgument, "Update identifier cannot be empty"
  End If

  If StrComp(this.idValue, upName, vbBinaryCompare) = 0 Then Exit Property

  this.idValue = Trim$(upName)
  ResetRegistryCache
End Property

'@section PrivateAccessors
'===============================================================================
'@label:EnsureWorksheet
'@description Guard that the registry worksheet reference is initialised before use.
Private Function EnsureWorksheet() As Worksheet
  If this.sh Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "UpdatedValues worksheet was not initialised"
  End If
  Set EnsureWorksheet = this.sh
End Function

'@label:RegistryListObject
'@description Lazily create or retrieve the ListObject storing watcher metadata.
Private Property Get RegistryListObject() As ListObject
  If this.registryLo Is Nothing Then
    Set this.registryLo = FetchRegistryListObject(createWhenMissing:=True)
  End If
  Set RegistryListObject = this.registryLo
End Property

'@label:FetchRegistryListObject
'@description Return the registry ListObject, optionally creating it when missing.
'@param createWhenMissing Boolean instructing whether to create the ListObject.
Private Function FetchRegistryListObject(ByVal createWhenMissing As Boolean) As ListObject
  Dim sh As Worksheet
  Dim loName As String
  Dim lo As ListObject

  Set sh = EnsureWorksheet()
  loName = CurrentLoName(ID)

  On Error Resume Next
    Set lo = sh.ListObjects(loName)
  On Error GoTo 0

  If lo Is Nothing Then
    If Not createWhenMissing Then Exit Function
    Set lo = CreateRegistryListObject(sh, loName)
  Else
    EnsureRegistryHeaders lo
  End If

  Set FetchRegistryListObject = lo
End Function

'@label:RegistryTable
'@description Cached CustomTable helper over the registry ListObject keyed by range name.
Private Property Get RegistryTable() As ICustomTable
  If this.registryTable Is Nothing Then
    Set this.registryTable = CustomTable.Create(RegistryListObject, HEADER_RANGE_NAME)
  End If
  Set RegistryTable = this.registryTable
End Property

'@label:ResetRegistryCache
'@description Clear cached references so subsequent calls rebuild registry helpers.
Private Sub ResetRegistryCache()
  Set this.registryLo = Nothing
  Set this.registryTable = Nothing
End Sub

'@label:InvalidateRegistryTable
'@description Drop only the CustomTable cache when the underlying ListObject changes.
Private Sub InvalidateRegistryTable()
  Set this.registryTable = Nothing
End Sub

'@label:CurrentLoName
'@description Compose the ListObject name for the supplied identifier.
Private Function CurrentLoName(ByVal listName As String) As String
  CurrentLoName = LISTOBJECT_PREFIX & NormalizeColumnKey(listName)
End Function

'@label:NormalizedId
'@description Return the normalised identifier used for named range suffixes.
Private Function NormalizedId() As String
  NormalizedId = NormalizeColumnKey(this.idValue)
End Function

'@label:CreateRegistryListObject
'@description Build an empty registry ListObject with the expected headers.
Private Function CreateRegistryListObject(ByVal sh As Worksheet, ByVal loName As String) As ListObject
  Dim startColumn As Long
  Dim headerStart As Range
  Dim tableRange As Range

  startColumn = NextAvailableColumn(sh)
  Set headerStart = sh.Cells(1, startColumn)

  headerStart.Value = HEADER_COLUMN_NAME
  headerStart.Offset(0, 1).Value = HEADER_RANGE_NAME
  headerStart.Offset(0, 2).Value = HEADER_STATUS
  headerStart.Offset(0, 3).Value = HEADER_TAG

  Set tableRange = sh.Range(headerStart, headerStart.Offset(1, REGISTRY_COLUMN_COUNT - 1))

  Set CreateRegistryListObject = sh.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
  CreateRegistryListObject.Name = loName
  EnsureRegistryHeaders CreateRegistryListObject

End Function

'@label:EnsureRegistryHeaders
'@description Guarantee the registry ListObject exposes the expected header names.
Private Sub EnsureRegistryHeaders(ByVal lo As ListObject)
  Dim expected() As String
  Dim idx As Long

  expected = Array(HEADER_COLUMN_NAME, HEADER_RANGE_NAME, HEADER_STATUS, HEADER_TAG)

  Do While lo.ListColumns.Count < REGISTRY_COLUMN_COUNT
    lo.ListColumns.Add
  Loop

  For idx = LBound(expected) To UBound(expected)
    lo.ListColumns(idx + 1).Name = expected(idx)
    lo.HeaderRowRange.Cells(1, idx + 1).Value = expected(idx)
  Next idx
End Sub

'@label:NextAvailableColumn
'@description Find the next column available to place the registry ListObject.
Private Function NextAvailableColumn(ByVal sh As Worksheet) As Long
  Dim lastUsed As Range

  On Error Resume Next
    Set lastUsed = sh.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
  On Error GoTo 0

  If lastUsed Is Nothing Then
    NextAvailableColumn = 1
  Else
    NextAvailableColumn = Application.WorksheetFunction.Min(lastUsed.Column + 2, sh.Columns.Count - REGISTRY_COLUMN_COUNT + 1)
  End If
End Function

'@section ColumnRegistration
'===============================================================================
'@label:AddColumns
'@details Scan a source ListObject for tagged watcher columns, synchronising the registry table
'so named ranges remain aligned with the current worksheet layout.
'@params
'- source ListObject. Table to inspect for update tags.
Private Sub AddColumns(ByVal source As ListObject)
  Dim headerCell As Range
  Dim expected As BetterArray
  Dim registryIndex As BetterArray
  Dim normalizedRange As String
  Dim columnLabel As String
  Dim tagValue As String
  Dim rowRange As Range
  Dim isNew As Boolean

  If source Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, "ListObject reference is required when registering updated values"
  End If

  If source.HeaderRowRange Is Nothing Then Exit Sub

  Set expected = New BetterArray
  expected.LowerBound = 1
  Set registryIndex = BuildRegistryIndex()

  For Each headerCell In source.HeaderRowRange.Cells
    tagValue = ColumnTag(headerCell)
    If ShouldWatchColumn(tagValue) Then
      normalizedRange = BuildRangeName(NormalizeColumnKey(headerCell.Value))
      columnLabel = BuildColumnLabel(source.Name, headerCell.Value)
      Set rowRange = EnsureRegistryRowFromIndex(registryIndex, normalizedRange, isNew)
      PopulateRegistryRow rowRange, columnLabel, normalizedRange, tagValue, isNew
      EnsureNamedRange source, headerCell.Value, normalizedRange
      AddExpectedKey expected, normalizedRange
    End If
  Next headerCell

  PruneObsoleteEntries expected
End Sub

'@label:ColumnTag
'@description Read the metadata tag associated with a column header, searching above the header row.
'@param headerCell Range representing the target header.
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

'@label:ShouldWatchColumn
'@description Determine whether a column should be monitored based on its tag value.
'@param tagValue String tag retrieved from worksheet metadata.
Private Function ShouldWatchColumn(ByVal tagValue As String) As Boolean
  Select Case tagValue
    Case TAG_TRANSLATE_TEXT, TAG_TRANSLATE_FORMULA, TAG_WATCH_UPDATE
      ShouldWatchColumn = True
  End Select
End Function

'@label:BuildRegistryIndex
'@description Create a lightweight lookup of existing registry rows keyed by normalised range name.
Private Function BuildRegistryIndex() As BetterArray
  Dim index As BetterArray
  Dim lo As ListObject
  Dim listRow As ListRow
  Dim rangeName As String
  Dim key As String
  Dim entry As Variant

  Set index = New BetterArray
  index.LowerBound = 1
  Set lo = FetchRegistryListObject(createWhenMissing:=False)

  If lo Is Nothing Then
    Set BuildRegistryIndex = index
    Exit Function
  End If

  For Each listRow In lo.ListRows
    rangeName = CStr(listRow.Range.Cells(1, RegistryRangeName).Value)
    key = NormalizeValue(rangeName)
    If key <> vbNullString Then
      If FindRegistryRowInIndex(index, key) Is Nothing Then
        entry = Array(key, listRow.Range)
        index.Push entry
      End If
    End If
  Next listRow

  Set BuildRegistryIndex = index
End Function

'@label:EnsureRegistryRowFromIndex
'@description Locate or append a registry row for the provided range name.
'@params
'- registryIndex BetterArray index built from existing entries.
'- rangeName String normalised range identifier.
'- isNew Boolean flag indicating whether a new row was added.
Private Function EnsureRegistryRowFromIndex(ByVal registryIndex As BetterArray, _
                                            ByVal rangeName As String, _
                                            ByRef isNew As Boolean) As Range
  Dim key As String
  Dim existing As Range
  Dim lo As ListObject
  Dim listRow As ListRow
  Dim entry As Variant

  key = NormalizeValue(rangeName)
  Set existing = FindRegistryRowInIndex(registryIndex, key)
  isNew = False

  If existing Is Nothing Then
    Set lo = RegistryListObject
    Set listRow = lo.ListRows.Add
    Set EnsureRegistryRowFromIndex = listRow.Range
    entry = Array(key, listRow.Range)
    registryIndex.Push entry
    isNew = True
    InvalidateRegistryTable
  Else
    Set EnsureRegistryRowFromIndex = existing
  End If
End Function

'@label:FindRegistryRowInIndex
'@description Locate a registry row range in the cached index by its normalised key.
Private Function FindRegistryRowInIndex(ByVal registryIndex As BetterArray, _
                                        ByVal key As String) As Range
  Dim idx As Long
  Dim entry As Variant

  If registryIndex Is Nothing Then Exit Function
  If registryIndex.Length = 0 Then Exit Function

  For idx = registryIndex.LowerBound To registryIndex.UpperBound
    entry = registryIndex.Item(idx)
    If IsArray(entry) Then
      If StrComp(CStr(entry(LBound(entry))), key, vbBinaryCompare) = 0 Then
        Set FindRegistryRowInIndex = entry(UBound(entry))
        Exit Function
      End If
    End If
  Next idx
End Function

'@label:PopulateRegistryRow
'@description Write watcher metadata to the specified registry row and initialise status flags.
Private Sub PopulateRegistryRow(ByVal rowRange As Range, _
                                ByVal columnLabel As String, _
                                ByVal rangeName As String, _
                                ByVal tagValue As String, _
                                ByVal isNew As Boolean)
  If rowRange Is Nothing Then Exit Sub

  rowRange.Cells(1, RegistryColumnLabel).Value = columnLabel
  rowRange.Cells(1, RegistryRangeName).Value = rangeName

  If isNew Or NormalizeValue(CStr(rowRange.Cells(1, RegistryStatus).Value)) = vbNullString Then
    rowRange.Cells(1, RegistryStatus).Value = STATUS_DEFAULT
  End If

  rowRange.Cells(1, RegistryHeaderTag).Value = tagValue
End Sub

'@label:EnsureNamedRange
'@description Create or refresh the workbook defined name pointing to the watched column.
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

'@label:AddExpectedKey
'@description Track range names detected during scanning to support registry pruning.
Private Sub AddExpectedKey(ByVal expected As BetterArray, ByVal rangeName As String)
  Dim key As String

  If expected Is Nothing Then Exit Sub

  key = NormalizeValue(rangeName)
  If key = vbNullString Then Exit Sub

  If expected.Includes(key) Then Exit Sub
  expected.Push key
End Sub

'@label:ExpectedContains
'@description Test whether the expected list already includes the provided range name.
Private Function ExpectedContains(ByVal expected As BetterArray, ByVal rangeName As String) As Boolean
  Dim key As String

  ExpectedContains = False
  If expected Is Nothing Then Exit Function
  If expected.Length = 0 Then Exit Function

  key = NormalizeValue(rangeName)
  If key = vbNullString Then Exit Function

  ExpectedContains = expected.Includes(key)
End Function

'@label:PruneObsoleteEntries
'@description Remove registry rows and named ranges for columns no longer tagged for watching.
Private Sub PruneObsoleteEntries(ByVal expected As BetterArray)
  Dim lo As ListObject
  Dim rowIndex As Long
  Dim rangeName As String
  Dim totalRows As Long

  Set lo = RegistryListObject
  totalRows = lo.ListRows.Count

  If expected Is Nothing Then
    RemoveAllNamedRanges
    DeleteAllRegistryRows lo
    Exit Sub
  End If

  If expected.Length = 0 Then
    RemoveAllNamedRanges
    DeleteAllRegistryRows lo
    Exit Sub
  End If

  For rowIndex = totalRows To 1 Step -1
    rangeName = CStr(lo.ListRows(rowIndex).Range.Cells(1, RegistryRangeName).Value)
    If NormalizeValue(rangeName) = vbNullString Then
      lo.ListRows(rowIndex).Delete
      InvalidateRegistryTable
      GoTo ContinueLoop
    End If

    If Not ExpectedContains(expected, rangeName) Then
      RemoveNamedRange rangeName
      lo.ListRows(rowIndex).Delete
      InvalidateRegistryTable
    End If
ContinueLoop:
  Next rowIndex
End Sub

'@label:DeleteAllRegistryRows
'@description Remove each data row from the registry ListObject and reset caches.
Private Sub DeleteAllRegistryRows(ByVal lo As ListObject)
  Do While lo.ListRows.Count > 0
    lo.ListRows(lo.ListRows.Count).Delete
  Loop
  InvalidateRegistryTable
End Sub

'@section StatusTracking
'===============================================================================
'@label:IsUpdated
'@description Check the registry status flag for the supplied column name.
Private Property Get IsUpdated(ByVal checkColName As String) As Boolean
  Dim normalizedRange As String
  Dim statusValue As String

  normalizedRange = BuildRangeName(NormalizeColumnKey(checkColName))
  statusValue = RegistryTable.Value(HEADER_STATUS, normalizedRange)
  IsUpdated = (NormalizeValue(statusValue) = STATUS_UPDATED)
End Property

'@label:CheckStatus
'@description Compare the changed range with registered watch ranges and flag the first match as updated.
'@params
'- sh Worksheet hosting the change event.
'- Target Range that triggered the change.
Private Sub CheckStatus(ByVal sh As Worksheet, ByVal Target As Range)
  Dim rangeColumn As Range
  Dim rangeCell As Range
  Dim namedRange As Range
  Dim rangeName As String

  If sh Is Nothing Then Exit Sub
  If Target Is Nothing Then Exit Sub

  Set rangeColumn = RegistryTable.DataRange(HEADER_RANGE_NAME, includeHeaders:=False, strictSearch:=True)
  If rangeColumn Is Nothing Then Exit Sub

  For Each rangeCell In rangeColumn.Cells
    rangeName = CStr(rangeCell.Value)
    If NormalizeValue(rangeName) = vbNullString Then
      RemoveOutdatedEntry rangeCell
      GoTo ContinueLoop
    End If

    Set namedRange = ResolveNamedRange(rangeName)
    If namedRange Is Nothing Then
      RemoveOutdatedEntry rangeCell
      GoTo ContinueLoop
    End If

    If Not Intersect(Target, namedRange) Is Nothing Then
      UpdateStatusForCell rangeCell, STATUS_UPDATED
      Exit Sub
    End If
ContinueLoop:
  Next rangeCell
End Sub

'@label:ResolveNamedRange
'@description Attempt to resolve the workbook defined name into a concrete Range reference.
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

'@label:UpdateStatusForCell
'@description Write the supplied status flag into the registry row matching the range cell.
Private Sub UpdateStatusForCell(ByVal rangeCell As Range, ByVal statusValue As String)
  Dim statusCell As Range

  If rangeCell Is Nothing Then Exit Sub

  Set statusCell = rangeCell.Offset(0, RegistryStatus - RegistryRangeName)
  statusCell.Value = statusValue
End Sub

'@label:RemoveOutdatedEntry
'@description Clean registry rows that point to missing or invalid named ranges.
Private Sub RemoveOutdatedEntry(ByVal rangeCell As Range)
  Dim lo As ListObject
  Dim rowIndex As Long
  Dim rangeName As String

  If rangeCell Is Nothing Then Exit Sub

  Set lo = RegistryListObject
  rowIndex = rangeCell.Row - lo.HeaderRowRange.Row

  If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Sub

  rangeName = CStr(rangeCell.Value)
  RemoveNamedRange rangeName
  lo.ListRows(rowIndex).Delete
  InvalidateRegistryTable
End Sub

'@label:ResetUpdateStatus
'@label:ResetUpdateStatus
'@description Reset every registry row to the default (non-updated) status value.
Private Sub ResetUpdateStatus()
  Dim statusColumn As Range
  Dim cell As Range

  Set statusColumn = RegistryTable.DataRange(HEADER_STATUS, includeHeaders:=False, strictSearch:=True)
  If statusColumn Is Nothing Then Exit Sub

  For Each cell In statusColumn.Cells
    cell.Value = STATUS_DEFAULT
  Next cell
End Sub

'@section CleanupHelpers
'===============================================================================
'@label:ClearUp
'@description Reset the watcher registry statuses to their default values.
Private Sub ClearUp()
  ResetUpdateStatus
End Sub

'@label:DeleteUp
'@description Remove the registry ListObject entirely and clean any associated named ranges.
Private Sub DeleteUp()
  Dim lo As ListObject
  Dim registryRange As Range

  Set lo = FetchRegistryListObject(createWhenMissing:=False)
  If lo Is Nothing Then Exit Sub

  RemoveAllNamedRanges

  Set registryRange = lo.Range
  lo.Unlist
  registryRange.Cells.Clear

  ResetRegistryCache
End Sub

'@label:RemoveAllNamedRanges
'@description Delete all workbook defined names created for watcher columns.
Private Sub RemoveAllNamedRanges()
  Dim lo As ListObject
  Dim rowIndex As Long
  Dim rangeName As String

  Set lo = FetchRegistryListObject(createWhenMissing:=False)
  If lo Is Nothing Then Exit Sub

  For rowIndex = 1 To lo.ListRows.Count
    rangeName = CStr(lo.ListRows(rowIndex).Range.Cells(1, RegistryRangeName).Value)
    RemoveNamedRange rangeName
  Next rowIndex
End Sub

'@label:RemoveNamedRange
'@description Delete a single defined name if it exists in the workbook.
Private Sub RemoveNamedRange(ByVal rangeName As String)
  Dim wb As Workbook

  If NormalizeValue(rangeName) = vbNullString Then Exit Sub

  Set wb = EnsureWorksheet().Parent

  On Error Resume Next
    wb.Names(rangeName).Delete
  On Error GoTo 0
End Sub

'@section Utilities
'===============================================================================
'@label:BuildColumnLabel
'@description Compose the descriptive label stored in the registry for a watched column.
Private Function BuildColumnLabel(ByVal tableName As String, ByVal columnName As String) As String
  BuildColumnLabel = tableName & "-" & columnName
End Function

'@label:BuildRangeName
'@description Generate the workbook defined name for a watched column.
Private Function BuildRangeName(ByVal normalizedColumnKey As String) As String
  BuildRangeName = RANGE_PREFIX & normalizedColumnKey & "_" & NormalizedId
End Function

'@label:NormalizeColumnKey
'@description Produce a workbook-friendly identifier from a column header.
Private Function NormalizeColumnKey(ByVal columnName As String) As String
  Dim idx As Long
  Dim ch As String
  Dim buffer As String

  columnName = Replace(columnName, Chr$(160), " ")
  columnName = Trim$(columnName)

  For idx = 1 To Len(columnName)
    ch = Mid$(columnName, idx, 1)

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

  NormalizeColumnKey = buffer
End Function

'@label:ReplaceRepeatedUnderscores
'@description Collapse multiple underscores into a single instance.
Private Function ReplaceRepeatedUnderscores(ByVal valueText As String) As String
  Do While InStr(valueText, "__") > 0
    valueText = Replace(valueText, "__", "_")
  Loop
  ReplaceRepeatedUnderscores = valueText
End Function

'@label:TrimUnderscores
'@description Remove leading and trailing underscore characters.
Private Function TrimUnderscores(ByVal valueText As String) As String
  Do While Len(valueText) > 0 And Left$(valueText, 1) = "_"
    valueText = Mid$(valueText, 2)
  Loop

  Do While Len(valueText) > 0 And Right$(valueText, 1) = "_"
    valueText = Left$(valueText, Len(valueText) - 1)
  Loop

  TrimUnderscores = valueText
End Function

'@label:NormalizeValue
'@description Normalise text by trimming spaces and applying lower case.
Private Function NormalizeValue(ByVal valueText As String) As String
  NormalizeValue = LCase$(Trim$(valueText))
End Function

'@label:ThrowError
'@description Raise a ProjectError aligned error from the current class.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errorMessage As String)
  Err.Raise CLng(errNumber), CLASS_NAME, "Error: " & errorMessage
End Sub

'@section InterfaceImplementation
'===============================================================================
'@label:IUpdatedValues_AddColumns
'@description Interface proxy that delegates to the internal AddColumns routine.
Private Sub IUpdatedValues_AddColumns(ByVal Lo As ListObject)
  AddColumns Lo
End Sub

'@label:IUpdatedValues_IsUpdated
'@description Interface proxy exposing the IsUpdated status flag.
Private Property Get IUpdatedValues_IsUpdated(ByVal colName As String) As Boolean
  IUpdatedValues_IsUpdated = IsUpdated(colName)
End Property

'@label:IUpdatedValues_CheckUpdate
'@description Interface proxy invoking the CheckStatus routine.
Private Sub IUpdatedValues_CheckUpdate(ByVal sh As Worksheet, ByVal Target As Range)
  CheckStatus sh, Target
End Sub

'@label:IUpdatedValues_ClearUp
'@description Interface proxy resetting registry statuses.
Private Sub IUpdatedValues_ClearUp()
  ClearUp
End Sub

'@label:IUpdatedValues_DeleteUp
'@description Interface proxy removing the registry table and workbook names.
Private Sub IUpdatedValues_DeleteUp()
  DeleteUp
End Sub

'@label:Exists
'@description Determine whether a registry ListObject already exists for the supplied identifier.
Public Function Exists(ByVal listName As String) As Boolean
  Dim sh As Worksheet
  Dim loName As String
  Dim lo As ListObject

  Set sh = EnsureWorksheet()
  loName = CurrentLoName(listName)

  On Error Resume Next
    Set lo = sh.ListObjects(loName)
  On Error GoTo 0

  Exists = Not lo Is Nothing
End Function
