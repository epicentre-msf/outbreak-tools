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

'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument,  ExcelMemberMayReturnNothing,  SuperfluousAnnotationArgument, MissingAnnotationArgument, ReadOnlyPropertyAssignment, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Wrapper for Listobjects manipulations")


Option Explicit

Implements ICustomTable
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
'@sub-title ResetCaches
'Reset transient caches so column lookups remain accurate after structural changes.
Private Sub ResetCaches()
  'Always rebuild the header index cache after structural changes to the table.
  Set this.headerIndexCache = New BetterArray
  this.headerIndexCache.LowerBound = 1
  this.hasShiftSetup = False
End Sub

'@sub-title ResetImportState
'@details Reset the caches of import state so that imports start fresh
Private Sub ResetImportState()
  this.hasImport = False
  this.hasHiddenList = False
  If NOTIMPORTEDCOLUMNS Is Nothing Then Set NOTIMPORTEDCOLUMNS = New BetterArray
  NOTIMPORTEDCOLUMNS.Clear
  If HIDDENCOLUMNSLIST Is Nothing Then Set HIDDENCOLUMNSLIST = New BetterArray
  HIDDENCOLUMNSLIST.Clear
  ClearBackUp
End Sub

'@label:ClearBackup
Private Sub ClearBackup()
  'Release backup buffers once the operation completes successfully.
  Set this.backupHeaders = Nothing
  Set this.backupRows = Nothing
  this.hasBackup = False
End Sub



'@sub-title CacheKeyForHeader
'@description Generates a unique cache key for header lookups combining search settings
'@param String colName - The column name to look up
'@param Boolean strictSearch - Whether to use exact matching
'@param Boolean matchCase - Whether to match case
'@return String The generated cache key
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

'@sub-title TryGetCachedIndex
'@description Attempts to retrieve a cached column index
'@param String cacheKey - The cache key to look up
'@param Long cachedValue - Output parameter for the found index
'@return Boolean True if found in cache, False otherwise
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

'@sub-title StoreHeaderIndex
'@description Stores a successful column lookup in the cache
'@param String cacheKey - The cache key to store
'@param Long columnIndexValue - The column index to cache
'@private
Private Sub StoreHeaderIndex(ByVal cacheKey As String, ByVal columnIndexValue As Long)

  'Persist successful lookups so subsequent queries avoid a Range.Find call.
  Dim entry As Variant

  If columnIndexValue <= 0 Then Exit Sub
  If this.headerIndexCache Is Nothing Then ResetCaches

  entry = Array(cacheKey, columnIndexValue)
  this.headerIndexCache.Push entry
End Sub

'@label:SaveTableSnapshot
'@sub-title SaveTableSnapshot
'@description Creates a backup of the current table state
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

'@label:RestoreTableSnapshot
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


'@label:EnsureTableInitialized
Private Sub EnsureTableInitialized()
  'Guard all table operations to avoid accidental calls before Create() runs.
  If this.Lo Is Nothing Then
    ThrowError ProjectError.ObjectNotInitialized, _
               "The ListObject reference is missing. Initialise the table first"
  End If
End Sub

'@section Instanciation
'===============================================================================

'@label:Create
'@prop-title Create a CustomTable object

'@details
'The custom data table is created just with one listObject. It is a wrapper arround
'the listobject class. Which basically add some functionalities to ease the
'data-manipulation process. A custom data table can have a "id" column. This 
'is a column where values are not duplicated, used to locate lines. When adding
'new rows to a table, one can decide to automatically number the id column in 
'a sequential way (1, 2, 3, ...) or a a prefix before those numbers. ie "line 1", "line 2".
'The prefix and the id column are provided at instanciation.
'
'
'@param Lo A ListObject
'@param idCol A String, the ID column of the CustomTable. Default is empty string
'@param idPrefix A String. A prefix to add to each values of the ID column when adding rows.
'       Default is empty string.
'@return A CustomTable Object
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

Public Property Get Self() As ICustomTable
  Set Self = Me
End Property

'@label:elements

'@section CustomTable elements
'===============================================================================

'@prop-title Worksheet of the listObject
Private Property Get Wksh() As Worksheet
  Dim Lo As ListObject
  Set Lo = Table()
  Set Wksh = Lo.Parent
End Property

'@prop-title Table listObject
Public Property Get Table() As ListObject
  EnsureTableInitialized
  Set Table = this.Lo
End Property

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

'@label:idvalue
'@prop-title IDs
Public Property Get IdValue() As String
  IdValue = this.idCol
End Property

Public Property Get PrefixValue() As String
  PrefixValue = this.idPrefix
End Property

Public Property Let IdValue(ByVal idCol As String)
  this.idCol = idCol
End Property

Public Property Let PrefixValue(ByVal idPrefix As String)
  this.idPrefix = idPrefix
End Property

'@label:HeaderRange
'@prop-title Header row Range of the CustomTable
'@return A Range object
'@export
Private Property Get HeaderRange() As Range
  Dim Lo As ListObject
  Set Lo = Table()
  Set HeaderRange = Lo.HeaderRowRange()
End Property

'@label:ColumnIndex
'@prop-title Column index of a Range
'@details
'If the column is not found, returns -1.
'@param colName A String. Column name
'@param stricSearch Boolean. Should you search column strictly? Default is False.
'@return A Long
'@export
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


'@label:ColumnExists
'@prop-title Check if a column exists in the CustomTable
'@param colName String. Column name to search
'@param strictSearch Boolean. Whether to search strictly with matchase, and the whole string or not. Default is False
'@return A Boolean
Private Function ColumnExists(ByVal colName As String, _
                              Optional ByVal strictSearch As Boolean = False, _
                              Optional ByVal matchCase As Boolean = False) As Boolean
    ColumnExists = (ColumnIndex(colName, strictSearch, matchCase) > 0)
End Function

'@label:ResolveListColumn
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


'@label:DataRange
'@prop-title Get the Range of a column
'
'@details
'Return a Data Range given a column name. When strictSearch is false, the 
'search for the column name is done using partial match. The first partial match
'found is the returned column. If the column names of the CustomTable are quite
'unique, the partial match returns unique values. In case some colunm names share
'some text, make sure to use strictSearch = True to avoid unexpected output.
'
'
'@param colName String, the name of the column to look for
'@param includeHeaders Boolean. Whether to return the colunmn plus the headers, default is false
'@param strictSearch Boolean. Look for partial values when searching the column name,
'                    without case sensitive search. Default is False.
'@return A Range object
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

'@label:Name
'@prop-title Name of the attached ListObject
'@return A String
'@export
Private Property Get Name() As String
  Name = this.csName
End Property

'@label:CellRange
'@prop-title A specific Cell of the CustomTable
'@details
'using the line number in a worksheet and a column name, get the corresponding
'cell at the intersection of the line and the column name.
'@param colName As String, the column of the cell
'@param lineNum Long, the line number in the excel worksheet
'@return A Range object
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


'@label:ValueRange
'@prop-title The range of a value
'@param colName String. The name of the column of the table
'@param keyName String. Key. Keys should be unique.
'@return Range
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

'@label:Value
'@prop-title Get one specific value on a table with key
'@details
'
'For tables with a key column, extract one specific value for a column name
'and given a keyvalue. If the table does not have a key, it returns an empty
'String. An empty string is returned in case it is impossible to convert the cell
'value to string
'
'@param colName String, the column name to extract the value from
'@param keyName String, the key
'@export
'@return A String. The value of the cell converted
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


'@label:modify
'@section Modify the table
'===============================================================================

'@label:InsertShiftTracker
'@sub-title Shift tracker
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


Private Function HasShiftTracker() As Boolean

  Dim store As IHiddenNames
  
  Set store = HiddenNames.Create(Wksh())

  If Not this.hasShiftSetup Then
    this.hasShiftSetup = (store.ValueAsString(TAG_INSERT_SHIFT_USED, "No") = "Yes")
  End If

  HasShiftTracker = this.hasShiftSetup

End Function

'@label:Resize
'@sub-title Resize a CustomTable
'@param addRows Optional, Boolean. Whether to Resize by adding rows. Default is False.
'@param insertShift Optional, Boolean. When adding rows, insert new line. Default is True
'@param totalRowCount Optional, Long. Row count threshold for deleting a line. Default is 0
'@param nbRows Optional, Long. Number or fows to add when adding new rows. Default is 5.
'@includeIds Optional, Boolean. Whether to number in a sequential way the ID column after resizing. Default is True.
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

'@label:AddTableRows Helper
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
    lo.Resize newRange
  Else
    'Extend the listobject in a single resize call when overlapping is not a concern.
    Set newRange = sheet.Range(sheet.Cells(headerRow, headerCol), _
                               sheet.Cells(headerRow + totalRows + nbRows - 1, headerCol + totalCols - 1))
    lo.Resize newRange
  End If

  ResetCaches
End Sub

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


'@label:RemoveEmptyDataRows
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
        'Delete the entire worksheet Row
        sh.Rows(lo.ListRows(rowIndex).Range().Row).Delete
      Else
        'Either delele the entire worksheet row or the the listRow
        lo.ListRows(rowIndex).Delete
      End If
    End If
  Next rowIndex
End Sub

'@label:DataRowCount
Private Function DataRowCount() As Long
  'Return the number of data rows currently stored in the table (without headers).
  Dim body As Range

  EnsureTableInitialized
  Set body = Table().DataBodyRange
  If body Is Nothing Then Exit Function

  DataRowCount = body.Rows.Count
End Function

Private Function DataColumnCount() As Long
   Dim body As Range

  EnsureTableInitialized
  Set body = Table().DataBodyRange
  If body Is Nothing Then Exit Function

  DataColumnCount = body.Columns.Count

End Function

'@label:EnsureRowCapacity
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

'@label: EnsureColumnCapacity
Private Sub EnsureColumnCapacity(ByVal requiredColumns As Long)

  Dim currCols As Long
  Dim Lo As ListObject
  Dim counter As Long
  Dim trimedColumns As Long

  EnsureTableInitialized

  trimedColumns = Application.WorksheetFunction.Max(0, requiredColumns)
  currCols = Application.WorksheetFunction.Max(1, DataColumnCount())

  If currCols <= trimedColumns Then Exit Sub
  Set Lo = Table()

  For counter = trimedColumns + 1 To currCols
    Lo.ListColumns(counter).Range.Cells.Clear
    Lo.ListColumns(counter).Delete
  Next
End Sub

'@label:TrimTableRows
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

'@label:SourceRowCount
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

'@label:AddIds
'@sub-title Add Ids to the ID column in a sequential way
'@details
'For some tables with key columns when resizing it is possible to add automatically 
'a sequential number (with prefix) to the ID column which is not modified by 
'the user, and in general is locked.
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

'@label:FormatIdValue
Private Function FormatIdValue(ByVal prefixText As String, ByVal numericPart As Long) As String
  'Build the ID value ensuring spacing is consistent when a prefix is supplied.
  If LenB(prefixText) = 0 Then
    FormatIdValue = CStr(numericPart)
  Else
    FormatIdValue = prefixText & " " & CStr(numericPart)
  End If
End Function


'@label:AddRows
'@sub-title Add Rows to a CustomTable
'
'@details
'Add rows to a CustomTable, either by inserting rows before resizing the 
'corresponding listObject or just by resizing the listObject directly. 
'There are different ways to add new rows: insert rows in the worksheet, or just
'expand the range of the listobject. The first option is useful when there are
'multiple list objects on the same worksheet and avoids overlaps. 
'On Linelist HList worksheet, there are no stacked listObjects on one worksheet.
'So there is no need to insert shifts like in analysis. This is time consuming.
'
'@param nbRows Optional, Long. number of rows to add. Default is 5
'@param insertShift Optional, Boolean. When adding rows, insert new line. Default is True
'@includeIds Optional, Boolean. Whether to number in a sequential way the ID column after
'            resizing. Default is True.
'
'@export
Private Sub AddRows(Optional ByVal nbRows As Long = 5, _ 
                    Optional Byval insertShift As Boolean = False, _ 
                    Optional ByVal includeIds As Boolean = True)
  
  Resize AddRows:=True, nbRows:=nbRows, insertShift:=insertShift, _ 
         includeIds:=includeIds
End Sub


'@label:RemoveRows
'@sub-title Remove empty rows from a CustomTable
'
'@details
'Delete some  empty rows in the CustomTable, and resize the listObject afterwards.
'Empty rows are defined as rows where the total row count is less
'than a given value provided by the user. The default value here is 0. In the
'linelist HList, the total row count is number of  formulas on the line. Formulas
'are considered as non empty cells.
'
'@param totalCount Optional, Long. total row count on the line
'@param includeIds Optional Boolean. Whether to number in a sequential way the ID
'       column after resizing. Default is true
'
'@export
Private Sub RemoveRows(Optional ByVal totalCount As Long = 0, _ 
                       Optional ByVal includeIds As Boolean = True, _
                       Optional ByVal forceShift As Boolean = False)

  Resize AddRows:=False, totalRowCount:=totalCount, includeIds:=includeIds, forceShift:=forceShift
End Sub

'@label:SetValidation
'@sub-title Set some validations on the listObject cell
'@param colName String. Name of column to set the validation on
'@param drop A Dropdown Object.
'@param dropName String. The name of the dropdown to select from the dropdown object
'@param alertType Optional, String. Type of alert to raise if values are not in dropdown. Could take "info", "warning", "error". Default is "info".
'@param message Optional, String. Message to raise if values are not in the dropdown. Default is vbNullString
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


'@label:SetValue
'@sub-title Change the value of one specific cell in the CustomTable
'@details
'
'Given one specific cell in the CustomTable, change value of the cell. The
'cell is identified using the key of the row and the column name of the column.
'Use with caution as it modifies data inside the ListObject.
'
'@param colName String. Column name of the cell to modify
'@param keyName String. They key of the cell to modify
'@param newValue String. The new value to add
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

'@label:Clean
'@sub-title Clean all the columns of the CustomTable except the key column
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

'@label:Sort
'@sub-title Sort The Current table on some columns
'@details
'
'Sort the columns in Ascending order. One can either sort one column or 
'sort multiple columns provided in a BetterArray Object. The sort is done
'on the provided column first and then on the BetterArray if there is one.
'
'@param colName Optional String. The column to sort the table on. Default is vbNullString
'@param colList Optional, an Object, expecting a BetterArray of colNames. Default is Nothing
'@param directSort Optional, Boolean. Whether to sort on first element or to sort directly. Default is True, sort directly.
'@param sortOrder Optional, String. Sort
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

'@label:SortHelper
'@sub-title Helper for the sort sub
'@colName String. The column to sort on
'@directSort Boolean. Sort directly or on first element? Default is True.
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

'@label:SortOnFirst
'@sub-title Sort a column name by bringing together first values found
'@param colName A String. The name of the column
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
     sortedValue = colRng.Cells(counter, 1).Value
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

Private Sub SortSimple(ByVal colName As String, Optional ByVal strictSearch As Boolean = True)

  Dim colRng As Range
  Dim LoRng As Range

  Set LoRng = DataRange(includeHeaders:=True)
  Set colRng = DataRange(colName, strictSearch:=strictSearch)

  If (LoRng Is Nothing) Or IsEmpty(LoRng) Or (colRng Is Nothing) Or _
      (IsEmpty(colRng)) Then Exit Sub

  LoRng.Sort key1:=colRng, Order1:=xlAscending, Header:=xlYes
End Sub


'@label:ImportExport
'@section Import Exports
'
'@label:Import
'@sub-title: Import a DataSheet or a CustomTable
'
'@details
'
'If keepSourceHeaders = True, the import process will import the incoming data, 
'by keeping the imported data as "Master". All the the headers of the
'incoming data are kept as is and all the headers of the current data are 
'removed. keepSourceHeaders = True will completly clean the data in the current
'CustomTable and import ALL data in the incoming object.
'
'
'@param impTab DataSheet of CustomTable Object. The table to import
'@param pasteAtBottom Optional, Boolean. Paste the data at the bottom of the current table. Default is False
'@param strictColumnSearch Optional, Boolean. Use strict rules when searching column names in the current table. Default is False
'@param keepSourceHeaders Optional, Boolean. Replace the headers of the current table by the imported ones. Default is False
'@param insertShift Optional, Boolean. Insert some lines when extending the table. Default is True.
'@param formatHeaders A BetterArray of headers names where column format should be kept intact
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

'@sub-title Import a DataSheet or a CustomTable
'@param impTab DataSheet of CustomTable Object. The table to import
'@param pasteAtBottom Optional, Boolean. Paste the data at the bottom of the current table. Default is False
'@param strictColumnSearch Optional, Boolean. Use strict rules when searching column names in the current table. Default is False
'@param insertShift Optional, Boolean. Insert some lines when extending the table. Default is True.
'@param formatHeaders A BetterArray of headers names where column format should be kept
'@export
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
  TrimTableRows targetRowCount 'Trim the table to the targetRowCount
End Sub

'@sub-title Import the formatting of a range into a currrent range
'@param impRng The range to  importing range
'@param currRng the current range in the current CustomTable
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


'@label:ImportAll
'@sub-title Import a CustomTable and a DataSheet
'@details
'
'All the columns of the imported Object ARE  imported, and
'replace all the current columns WITH all the imported ones
'
'@param impTab A CustomTable or DataSheet Object
'@export
Private Sub ImportAll(ByVal impTab As Object, Optional ByVal insertShift As Boolean = False)

  Dim currhRng As Range
  Dim imphRng As Range
  Dim currLo As listObject
  Dim sh As Worksheet
  Dim bodyRng As Range
  Dim incomingRows As Long
  Dim headerRow As Long
  Dim headerCol As Long

  Set currLo = Table()
  Set currhRng = HeaderRange()
  Set imphRng = impTab.HeaderRange()

  Set sh = Wksh()

  incomingRows = SourceRowCount(impTab)
  incomingRows = Application.WorksheetFunction.Max(1, incomingRows)

  TrimTableRows 1
  EnsureColumnCapacity imphRng.Columns.count

  headerRow = currhRng.Row
  headerCol = currhRng.Column
  Set bodyRng = sh.Range(sh.Cells(headerRow + 1, headerCol), _
                           sh.Cells(headerRow + incomingRows, headerCol + imphRng.Columns.count - 1))
  
  currhRng.Value = imphRng.Value
  bodyRng.Value = impTab.DataRange().Value

  Set bodyRng = sh.Range(bodyRng.Cells(0, 1), bodyRng.Cells(bodyRng.Rows.Count, bodyRng.Columns.Count))
  currLo.Resize bodyRng
End Sub


'@label:ImportColumnsNotFound
'@prop-title After one import, get the columns not imported
'@export
'return A BetterArray object
Private Property Get ImportColumnsNotFound() As BetterArray

  If Not this.hasImport Then
    LogInfo "The current table " & Name() & "has no imports"
    ShowDebug
    Exit Property
  End If

  If NOTIMPORTEDCOLUMNS Is Nothing Then Exit Property
  Set ImportColumnsNotFound = NOTIMPORTEDCOLUMNS.Clone()
End Property

'@label:HasColumnsNotImported
'@prop-title Are there imports in the CustomTable?
'@export
'@return A Boolean
Private Property Get HasColumnsNotImported() As Boolean
  If Not this.hasImport Then Exit Property
  If NOTIMPORTEDCOLUMNS Is Nothing Then Exit Property
  HasColumnsNotImported = (NOTIMPORTEDCOLUMNS.Length >= 1)
End Property

'Those functions are for Unhiding and hiding back columns before Import or Exports
'Sometimes CustomTable contains hidden columns, range.Find method does not include
'hidden rows or column when searching, so we need to unhide all columns of
'the worksheet before Import/Export and to return back the state of hidden columns
'after import/export.

'@sub-title UnhideHiddenColumns
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

'@sub-title ReturnBackHiddenColumns
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

'@label:Export
'@sub-title Export a CustomTable to a Worksheet
'@details
'
'Export the current CustomTable to a worksheet. The startLine is where to paste
'the data in the current worksheet.
'
'@param hedersTable BetterArray Object, the list of the columns to export
'@param startLine Optional, Long. The line where to start exporting in the worksheet. Default is 1.
Private Sub Export(ByVal sh As Worksheet, _ 
           Optional ByVal headersTable As Object = Nothing, _ 
           Optional ByVal startLine As Long = 1)

  Dim counter As Long
  Dim inpRng As Range
  Dim outrng As Range
  Dim headerName As String
  Dim headerCounter As Long
  Dim headersList As BetterArray

  If sh Is Nothing Then 
    LogInfo "Unable to export CustomTable, sheet not initialized", checkingError
    Exit Sub
  End If

  Set headersList = New BetterArray

  If (TypeName(headersTable) = "BetterArray") Then
    Set headersList = headersTable.Clone()
  Else
    LogInfo "Provided headersTable for export not found, all the headers are exported"
    ShowDebug
    headersList.FromExcelRange HeaderRange(), DetectLastRow:=False, DetectLastColumn:=False
  End If

  sh.Cells.Clear
  headerCounter = 1

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

  'Return back the previously hidden columns
  ReturnBackHiddenColumns
  ResetCaches
End Sub

'@label:Checkings
'@section Checkings
'===============================================================================

'@sub-title Add elements to the checking object for future notes
'@param Label String, the message to add.
'@param scope Byte. The scope of the checking object.
'@export
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

'@label:HasChecking
'@prop-title Check if the object has checkings
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:CheckingValues
'@prop-title Get the cheking values
'@export
Private Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
End Property


Private Sub ShowDebug()
    Debug.Print "Check the informations of custom table object " & Name()
End Sub


Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "CustomTable", "Error: " & errorMessage
End Sub


'@jump:Create

'@jump:elements
'CustomTable elements ----------------------------------------------------------

'@jump:HeaderRange
'Header row range
Private Property Get ICustomTable_HeaderRange() As Range
  Set ICustomTable_HeaderRange = HeaderRange()
End Property

'@jump:idvalue
Private Property Get ICustomTable_IdValue() As String
  ICustomTable_IdValue = IdValue()
End Property

'@jump:DataRange
'Range of a column
Private Property Get ICustomTable_DataRange( _
          Optional ByVal colName As String = "__all__", _
          Optional ByVal includeHeaders As Boolean = False, _
          Optional ByVal strictSearch As Boolean = False, _
          Optional ByVal matchCase As Boolean = False) As Range
  Set ICustomTable_DataRange = DataRange(colName, includeHeaders, strictSearch, matchCase)
End Property

'@jump:Name
'Name of attached listObject
Private Property Get ICustomTable_Name() As String
  ICustomTable_Name = Name()
End Property

'@jump:CellRange
'Specific cell in the customtable
Private Property Get ICustomTable_CellRange(ByVal colName As String, _ 
                                            ByVal lineNum As Long) As Range
  Set ICustomTable_CellRange = CellRange(colName, lineNum)
End Property

'@jump:Value
'Specific value of a customtable with ID column
Private Property Get ICustomTable_Value(ByVal colName As String, _ 
                                        ByVal keyName As String) As String
  ICustomTable_Value = Value(colName, keyName)
End Property

Private Property Get ICustomTable_ValueRange(ByVal colName As String, _ 
                                             ByVal keyName As String) As Range

    Set ICustomTable_ValueRange = ValueRange(colName, keyName)
End Property

'@jump:modify
'Modify the CustomTable -------------------------------------------------------

'@jump:AddIds
'AddIds to the CustomTable
Private Sub ICustomTable_AddIds()
    AddIds
End Sub


'@jump:AddRows
'AddRows to a CustomTable
Private Sub ICustomTable_AddRows(Optional ByVal nbRows As Long = 5, _ 
                                 Optional Byval insertShift As Boolean = False, _ 
                                 Optional ByVal includeIds As Boolean = True)

    AddRows nbRows:=nbRows, _ 
            insertShift:=insertShift, _ 
            includeIds:=includeIds
End Sub

'@jump:RemoveRows
'Remove rows to a CustomTable
Private Sub ICustomTable_RemoveRows(Optional ByVal totalCount As Long = 0, _ 
                                    Optional ByVal includeIds As Boolean = True, _ 
                                    Optional ByVal forceShift As Boolean = False)
  RemoveRows totalCount:=totalCount, includeIds:=includeIds, forceShift:=forceShift
End Sub


'@jump:SetValidation
'Set validations on one specific column of the listObject
Private Sub ICustomTable_SetValidation(ByVal colName As String, _ 
     ByVal drop As IDropdownLists, _
     ByVal dropName As String, _ 
     Optional ByVal alertType As String = "info", _ 
     Optional ByVal message As String = vbNullString _ 
    )
    SetValidation colName:=colName, drop:=drop, dropName:=dropName, _ 
                  alertType:=alertType, message:=message
End Sub


'@jump:SetValue
'change the values of one cell
Private Sub ICustomTable_SetValue(ByVal colName As String, _ 
        ByVal keyName As String, _ 
        ByVal newValue As String)
    SetValue colName, keyName, newValue
End Sub


'@jump:Clean
'Clean the CustomTable
Private Sub ICustomTable_Clean()
  Clean
End Sub


'@jump:Sort
'Sort a CustomTable
Private Sub ICustomTable_Sort(Optional ByVal colName As String = vbNullString, _ 
                              Optional ByVal colList As Object = Nothing, _ 
                              Optional ByVal directSort As Boolean = True, _ 
                              Optional ByVal strictSearch As Boolean = True)
  Sort colName:=colName, colList:=colList, directSort:=directSort, strictSearch:=strictSearch
End Sub


'@jump:ImportExport
'Import Export Functions -------------------------------------------------------

'@jump:Import
'Imports an object
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


'@jump:Export
'Export to a worksheet
Private Sub ICustomTable_Export(ByVal sh As Worksheet, _
                               Optional ByVal headersTable As Object = Nothing, _
                               Optional ByVal startLine As Long = 1)
  Export sh, headersTable, startLine
End Sub


'@jump:ImportAll
'import with imported as master
Private Sub ICustomTable_ImportAll(ByVal impTab As Object)
  ImportAll impTab
End Sub

'@jump:RestoreTableSnapshot
Private Sub ICustomTable_RestoreTableSnapshot()
  RestoreTableSnapshot
End Sub

'@jump:SaveTableSnapshot
Private Sub ICustomTable_SaveTableSnapshot()
  SaveTableSnapshot
End Sub

'@jump:ImportColumnsNotFound
'import columns not found
Private Property Get ICustomTable_ImportColumnsNotFound() As BetterArray
  Set ICustomTable_ImportColumnsNotFound = ImportColumnsNotFound()
End Property

'@jump:HasColumnsNotImported
'Test if there are imports in the CustomTable
Private Property Get ICustomTable_HasColumnsNotImported() As Boolean
    ICustomTable_HasColumnsNotImported = HasColumnsNotImported()
End Property


'@jump:Checkings
'Errors and checkings----------------------------------------------------------

'@jump:HasCheckings
Private Property Get ICustomTable_HasCheckings() As Boolean
    ICustomTable_HasCheckings = HasCheckings()
End Property

'@jump:CheckingValues
Private Property Get ICustomTable_CheckingValues() As Object
     Set ICustomTable_CheckingValues = CheckingValues()
End Property