VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLdictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist dictionary utilities"
Option Explicit


'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing,  UseMeaningfulName, HungarianNotation, ExcelMemberMayReturnNothing
'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Linelist dictionary utilities")
'@class LLdictionary
'@description
'Manages the "Dictionary" worksheet, the sheet that describes every variable of
'a linelist. Create binds the worksheet through a DataSheet. ColumnExists,
'VariableExists, UniqueValues and DataRange read it. AddColumn, InsertColumn,
'RemoveColumn and RenameColumn work on its columns, and ManageRows, InsertRows
'and DeleteRows on its rows. Sort orders the sheet, Clean empties it, Translate
'rewrites its labels, and Import and Export move it between workbooks.
'SpecialVars, ChoicesVars, GeoVars and TimeVars answer the variables of one
'kind.
'
'PREPARATION
'-------------------------------------------------------------------------------
'Prepare adds the helper columns a build reads: table names, column indices,
'CRF metadata, visibility and the geo rows. It runs once per sheet and Prepared
'says whether it has run. TotalNumberOfExports carries the number of export
'columns the sheet holds. The geo object Prepare is handed is typed Object and
'reached at run time, so LLGeo stays outside the compile dependencies of this
'class.
'
'HOW THE SHEET IS READ
'-------------------------------------------------------------------------------
'A row loop reads its block into a Variant once, works in memory, and writes
'the block back once.
'
'WHAT IS REPORTED
'-------------------------------------------------------------------------------
'Entries filed while a member runs leave through HasCheckings and
'CheckingValues.
'@depends BetterArray, Checking, CustomTable, DataSheet, HiddenNames, TranslationObject

Private Type TDictionary
    dictDataSheet As DataSheet
    numberOfExports As Long
    hasChecks As Boolean
    hiddenValues As HiddenNames
    sealed As Boolean
    allowedColumns As BetterArray      'Headers the dictionary accepts
    addedColumns As BetterArray        'Headers preparation creates
    preparedValue As Boolean           'True once the dictionary was seen prepared
    checkCounter As Long
    internalChecks As Checking
    collaboratorsMerged As Boolean     'True once the data sheet's entries were taken
End Type

Private this As TDictionary

Private Const CLASSNAME As String = "LLdictionary"
Private Const SHEETTYPEVERTICAL As String = "vlist1D"
Private Const SHEETTYPEHORIZONTAL As String = "hlist2D"
Private Const DEFAULTNUMBEROFEXPORTS As Long = 20
Private Const DEFAULTDICTIONARYCOLUMNS As String = "variable name|main label|dev comments|editable label|sub label|note|sheet name|sheet type|main section|sub section|status|register book|personal identifier|variable type|variable format|control|control details|unique|min|max|alert|message|formatting condition|formatting values|lock cells"
Private Const DEFAULTCREATEDCOLUMNS As String = "table name|column index|visibility|crf index|crf choices|crf status|list auto"
Private Const TOTAL_EXPORTS_NAME As String = "__ll_exports_total__"
Private Const MAXREPORTEDROWS As Long = 20

'@section Instantiation
'===============================================================================

'@label:create
'@prop-title Create a dictionary object
'@details
'Creates a linelist dictionary wrapper around the supplied worksheet. The
'routine initialises a backing `DataSheet` so all subsequent access goes through
'a consistent abstraction and callers interact with a predeclared instance.
'
'`numberOfExports` defaults to 0, which means "keep whatever the worksheet
'already stores". Fourteen of the twenty production call sites omit the
'argument, and a default of 20 wrote 20 over the real count and persisted it,
'so a workbook holding 3 exports came back claiming 20. `SetupErrors` then
'looped 1 To 20 over a three-row status table and reported notes about exports
'that are not there. The four call sites that pass a number still decide.
'@params
'- dictWksh Worksheet hosting the dictionary data
'- dictStartRow Long. Header row of the dictionary
'- dictStartColumn Long. First column of the dictionary
'- numberOfExports Optional Long. Expected number of export columns. Zero reads
'  the stored value.
'@return An LLdictionary instance
'@depends DataSheet, Checking
'@export
Public Function Create(ByVal dictWksh As Worksheet, ByVal dictStartRow As Long, _
                       ByVal dictStartColumn As Long, _
                       Optional ByVal numberOfExports As Long = 0) As LLdictionary

    Dim customDataSheet As DataSheet

    Set customDataSheet = DataSheet.Create(dictWksh, dictStartRow, dictStartColumn)

    Dim newDictionary As LLdictionary
    Set newDictionary = New LLdictionary
    Set newDictionary.Data = customDataSheet
    newDictionary.InitialiseTotalExports numberOfExports
    newDictionary.Seal
    Set Create = newDictionary
End Function

'@label:init-exports
'@sub-title Initialise total export counter
'@details
'Loads the stored number of export columns from the worksheet when available.
'When no persisted value exists, the routine falls back to the requested input
'and persists it so subsequent sessions pick up the same configuration.
'A requested total of zero means "use what is stored".
Public Sub InitialiseTotalExports(ByVal requestedTotal As Long)
    Dim storedTotal As Long

    storedTotal = StoredTotalExports()

    If requestedTotal > 0 Then
        If (storedTotal = 0) Or (storedTotal <> requestedTotal) Then
            ApplyTotalExports requestedTotal
        Else
            ApplyTotalExports storedTotal, persist:=False, suppressInvalidLog:=True
        End If
    ElseIf storedTotal > 0 Then
        ApplyTotalExports storedTotal, persist:=False, suppressInvalidLog:=True
    Else
        ApplyTotalExports DEFAULTNUMBEROFEXPORTS
    End If
End Sub

'@section Data Elements
'===============================================================================

'@label:data
'@prop-title DataSheet backing store
'@details
'Expose the lazily initialised `DataSheet` wrapper used for all worksheet
'interactions. Consumers rely on this to perform advanced queries without
'directly manipulating Excel ranges.
'@return DataSheet
Public Property Get Data() As DataSheet
    Set Data = this.dictDataSheet
End Property

Public Property Set Data(ByVal customDataSheet As DataSheet)
    GuardNotSealed "Data"
    Set this.dictDataSheet = customDataSheet
    Set this.hiddenValues = Nothing
    ClearPreparedCache
End Property

'@label:seal
'@sub-title Prevent further changes to creation-only setters
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard creation-only setters after sealing
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:exports
'@prop-title Number of export columns
'@details
'Number of export slots currently configured for the dictionary. This mirrors
'the legacy behaviour where designer workflows expect a bounded set of export
'columns ready for mapping.
'@return Long
Public Property Get TotalNumberOfExports() As Long
    If this.numberOfExports <= 0 Then
        ApplyTotalExports DEFAULTNUMBEROFEXPORTS, persist:=False, suppressInvalidLog:=True
    End If

    TotalNumberOfExports = this.numberOfExports
End Property

'@label:exports-set
'@details
'Persist the number of configured export columns while defaulting back to the
'class constant when an invalid value is provided.
'@params
'- numberOfExports Long. Requested number of export columns.
Public Property Let TotalNumberOfExports(ByVal numberOfExports As Long)
    ApplyTotalExports numberOfExports
End Property

'@section ExportCounterManagement
'===============================================================================

'@label:apply-exports
'@sub-title Normalise and persist export counters
'@details
'Centralises validation, caching and persistence of the export counter so every
'entry point (creation, import, manual setter) behaves consistently. The helper
'optionally suppresses warnings for legacy workbooks that do not yet store the
'hidden counter.
Private Sub ApplyTotalExports(ByVal requestedTotal As Long, _
                              Optional ByVal persist As Boolean = True, _
                              Optional ByVal suppressInvalidLog As Boolean = False)

    Dim normalisedTotal As Long

    normalisedTotal = NormaliseTotalExports(requestedTotal, suppressInvalidLog)
    this.numberOfExports = normalisedTotal

    If persist Then PersistTotalExports normalisedTotal
End Sub

'@label:normalise-exports
Private Function NormaliseTotalExports(ByVal requestedTotal As Long, ByVal suppressInvalidLog As Boolean) As Long
    If requestedTotal <= 0 Then
        NormaliseTotalExports = DEFAULTNUMBEROFEXPORTS
        If Not suppressInvalidLog Then
            LogInfo "Number of exports requested was not valid. Falling back to default", checkingWarning
        End If
    Else
        NormaliseTotalExports = requestedTotal
    End If
End Function

'@label:persist-exports
'@details
'Persist the export counter into the hidden name store of the dictionary
'worksheet.
Private Sub PersistTotalExports(ByVal totalCount As Long)
    Dim store As HiddenNames

    Set store = HiddenNameStore()
    If store Is Nothing Then Exit Sub

    WriteTotalExports store, totalCount
End Sub

'@label:persist-exports-on
'@details
'Persist the export counter onto a worksheet this object does not own. Used
'when the counter is written into an exported workbook.
Private Sub PersistTotalExportsOn(ByVal targetSheet As Worksheet, ByVal totalCount As Long)
    If targetSheet Is Nothing Then Exit Sub

    WriteTotalExports HiddenNames.Create(targetSheet), totalCount
End Sub

'@label:write-exports
Private Sub WriteTotalExports(ByVal store As HiddenNames, ByVal totalCount As Long)
    If store Is Nothing Then Exit Sub
    If totalCount < 0 Then totalCount = 0

    store.EnsureName TOTAL_EXPORTS_NAME, totalCount, HiddenNameTypeLong
    store.SetValue TOTAL_EXPORTS_NAME, totalCount
End Sub

'@label:stored-exports
'@sub-title Export counter held on the dictionary worksheet
Private Function StoredTotalExports() As Long
    Dim store As HiddenNames

    Set store = HiddenNameStore()
    If store Is Nothing Then Exit Function

    StoredTotalExports = store.ValueAsLong(TOTAL_EXPORTS_NAME, 0)
End Function

'@label:stored-exports-from
'@sub-title Export counter held on a foreign worksheet
Private Function StoredTotalExportsFromSheet(ByVal sourceSheet As Worksheet) As Long
    If sourceSheet Is Nothing Then Exit Function

    StoredTotalExportsFromSheet = HiddenNames.Create(sourceSheet).ValueAsLong(TOTAL_EXPORTS_NAME, 0)
End Function

'@label:hidden-name-store
'@sub-title Hidden name store of the dictionary worksheet
'@details
'The store is built once and kept. The host worksheet is fixed at creation, so
'the second test only guards the case where the backing DataSheet was replaced
'before the object was sealed.
Private Function HiddenNameStore() As HiddenNames
    Dim target As Worksheet

    If this.dictDataSheet Is Nothing Then Exit Function

    Set target = Wksh()
    If target Is Nothing Then Exit Function

    If this.hiddenValues Is Nothing Then
        Set this.hiddenValues = HiddenNames.Create(target)
    ElseIf Not this.hiddenValues.Sheet Is target Then
        Set this.hiddenValues = HiddenNames.Create(target)
    End If

    Set HiddenNameStore = this.hiddenValues
End Function

'@label:wksh
'@prop-title Host worksheet
'@details
'Expose the worksheet currently backing this dictionary. Primarily used by
'helpers that need to perform sheet-level operations such as adding columns or
'listobjects.
'@return Worksheet
Public Property Get Wksh() As Worksheet
    Set Wksh = Data().Wksh
End Property

'@label:startrow
'@prop-title Header row index
'@return Long
Private Property Get StartRow() As Long
    StartRow = Data().DataStartRow
End Property

'@label:startcol
'@prop-title First column index
'@return Long
Private Property Get StartColumn() As Long
    StartColumn = Data().DataStartColumn
End Property

'@label:endrow
'@prop-title Last data row
'@details
'One `End(xlUp)` scan per call, so callers read it once and hold the number
'rather than asking inside a loop.
'@return Long
Private Property Get DictEndRow() As Long
    DictEndRow = Data().DataEndRow
End Property

'@label:endcol
'@prop-title Last data column
'@details
'One `End(xlToLeft)` scan per call. Same rule as `DictEndRow`.
'@return Long
Private Property Get DictEndColumn() As Long
    DictEndColumn = Data().DataEndColumn
End Property

'@label:datarange
'@prop-title Dictionary range accessor
'@details
'Central entry point to retrieve ranges from the dictionary, either by column
'name or the full table. The helper delegates to the underlying `DataSheet`
'so filtering or header inclusion behaves consistently everywhere.
'@params
'- colName Optional column header name
'- includeHeaders Optional flag to include header row
'@return Range
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False) As Range
    Set DataRange = Data.DataRange(colName, includeHeaders, matchCase:=False)
End Property


'@section Block Readers And Writers
'===============================================================================
'@description
'Small helpers that move a whole column between the worksheet and a plain
'Variant array. They exist so every row loop in this class costs one read and
'one write instead of one crossing per cell.

'@label:column-values
'@sub-title Read a single-column Range as a 1-based list of values
'@details
'`Range.Value` hands back a plain value for one cell and a two-dimensional
'array otherwise. This flattens both cases so callers see one shape.
'@param target Range. A single-column block.
'@return Variant. A 1-based one-dimensional array of cell values.
Private Function ColumnValues(ByVal target As Range) As Variant
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long
    Dim nbRows As Long

    If target Is Nothing Then Exit Function

    nbRows = target.Rows.Count
    ReDim flat(1 To nbRows)

    If nbRows = 1 Then
        flat(1) = target.Cells(1, 1).Value
    Else
        rawValues = target.Columns(1).Value
        For counter = 1 To nbRows
            flat(counter) = rawValues(counter, 1)
        Next
    End If

    ColumnValues = flat
End Function

'@label:row-values
'@sub-title Read a single-row Range as a 1-based list of values
'@param target Range. A single-row block.
'@return Variant. A 1-based one-dimensional array of cell values.
Private Function RowValues(ByVal target As Range) As Variant
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long
    Dim nbColumns As Long

    If target Is Nothing Then Exit Function

    nbColumns = target.Columns.Count
    ReDim flat(1 To nbColumns)

    If nbColumns = 1 Then
        flat(1) = target.Cells(1, 1).Value
    Else
        rawValues = target.Rows(1).Value
        For counter = 1 To nbColumns
            flat(counter) = rawValues(1, counter)
        Next
    End If

    RowValues = flat
End Function

'@label:sheet-column-values
'@sub-title Read one worksheet column between two rows
'@param columnNumber Long. Worksheet column index.
'@param firstRow Long. Top row of the block.
'@param lastRow Long. Bottom row of the block.
'@return Variant. A 1-based one-dimensional array of cell values.
Private Function SheetColumnValues(ByVal columnNumber As Long, _
                                   ByVal firstRow As Long, _
                                   ByVal lastRow As Long) As Variant
    Dim hostSheet As Worksheet

    Set hostSheet = Wksh()
    SheetColumnValues = ColumnValues(hostSheet.Range(hostSheet.Cells(firstRow, columnNumber), _
                                                     hostSheet.Cells(lastRow, columnNumber)))
End Function

'@label:write-sheet-column
'@sub-title Write a 1-based list of values back into one worksheet column
'@details
'An element left `Empty` clears its cell, which is how the routines below leave
'a row untouched without writing an empty string into it. An empty string would
'count as content for `End(xlUp)` and move the last data row.
'@param columnNumber Long. Worksheet column index.
'@param firstRow Long. Top row of the block.
'@param values Variant. A one-dimensional array of values.
Private Sub WriteSheetColumn(ByVal columnNumber As Long, ByVal firstRow As Long, ByRef values As Variant)
    Dim hostSheet As Worksheet
    Dim block() As Variant
    Dim counter As Long
    Dim nbRows As Long

    If Not IsArray(values) Then Exit Sub

    nbRows = UBound(values) - LBound(values) + 1
    If nbRows <= 0 Then Exit Sub

    ReDim block(1 To nbRows, 1 To 1)
    For counter = 1 To nbRows
        block(counter, 1) = values(LBound(values) + counter - 1)
    Next

    Set hostSheet = Wksh()
    hostSheet.Range(hostSheet.Cells(firstRow, columnNumber), _
                    hostSheet.Cells(firstRow + nbRows - 1, columnNumber)).Value = block
End Sub

'@label:text-of
'@sub-title Read one cell value as text without raising
'@details
'A cell holding an error value (#N/A and the like) or Null makes `CStr` raise a
'type mismatch. Such a value can never match a header or a control tag, so it
'reads as empty text.
'@param cellValue Variant. The raw cell value.
'@return String
Private Function TextOf(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function

    TextOf = CStr(cellValue)
End Function

'@label:list-position
'@sub-title Find a value in a plain String list
'@details
'`BetterArray.Includes` copies the whole internal array on every call, so a
'linear scan over a local array is cheaper inside a row loop.
'@param values String array. The list, filled from index 1.
'@param usedCount Long. How many entries are filled.
'@param searchValue String. The value to look for.
'@return Long. The 1-based position, or zero when the value is absent.
Private Function ListPosition(ByRef values() As String, ByVal usedCount As Long, _
                              ByVal searchValue As String) As Long
    Dim counter As Long

    For counter = 1 To usedCount
        If values(counter) = searchValue Then
            ListPosition = counter
            Exit Function
        End If
    Next
End Function


'@section Column Configuration
'===============================================================================

'@label:ensure-columns
'@sub-title Build the allowed and generated column lists
'@details
'Both lists come from constants that never move, so they are built once per
'object and never rebuilt. Export columns are answered by the "export" prefix
'test in `ColumnExists`; pushing `export 1` ... `export N` into the list as
'well changed no answer, and it made the list depend on the export counter.
Private Sub EnsureColumnsInitialized()
    If Not this.allowedColumns Is Nothing Then Exit Sub

    Set this.allowedColumns = New BetterArray
    Set this.addedColumns = New BetterArray

    this.allowedColumns.Items = Split(DEFAULTDICTIONARYCOLUMNS, "|")

    'Columns introduced during preparation (e.g. table name, CRF helpers) are
    'tracked separately so we can optionally strip them during cleaning.
    this.addedColumns.Items = Split(DEFAULTCREATEDCOLUMNS, "|")

    this.allowedColumns.Concat this.addedColumns.Items
End Sub

'@label:ensure-format-columns
'@sub-title Ensure the conditional formatting columns exist
Private Sub EnsureFormatColumns()
    Data().AddFormatsColumns False, True, _
                             "formatting condition", "formatting values", _
                             "variable name", "control", "main label", _
                             "lock cells", "dev comments"
End Sub

'@label:columnexists
'@prop-title Test if a column exists
'@details
'Checks whether a column header is present in the dictionary and optionally
'verifies it belongs to the known schema. Recording the attempt helps capture
'incomplete dictionaries during setup.
'
'The lookup lower-cases its argument and searches with `matchCase` False.
'`LLExport.ExportColumnName` builds "Export 3" with a capital and relies on
'that, so both sides have to keep it true.
'@params
'- colName Column header to look for
'- checkValidity Optional: ensure the column belongs to the expected list
'@return Boolean
Public Function ColumnExists(ByVal colName As String, Optional ByVal checkValidity As Boolean = False) As Boolean
    Dim exists As Boolean

    'Dictionary headers are historically case-insensitive, so the underlying
    'datasheet lookup is forced to ignore casing differences.
    colName = LCase$(colName)
    exists = Data.ColumnExists(colName, matchCase:=False)

    If exists And checkValidity Then
        EnsureColumnsInitialized
        exists = this.allowedColumns.Includes(colName) Or (Split(colName, " ")(0) = "export")
        If Not exists Then LogInfo "Column " & colName & " exists but is not part of the allowed dictionary columns", checkingInfo
    ElseIf (Not exists) And checkValidity Then
        LogInfo "Column " & colName & " is missing from dictionary", checkingError
    End If

    ColumnExists = exists
End Function

'@label:columnindex
'@prop-title Column index lookup
'@param colName Column header name
'@param inDataRange Optional flag to restrict index to the data range
'@return Long
Private Function ColumnIndex(ByVal colName As String, Optional ByVal inDataRange As Boolean = False) As Long
    ColumnIndex = Data.ColumnIndex(colName, inDataRange, shouldExist:=True, matchCase:=False)
End Function

'@label:uniquevalues
'@prop-title Unique values in a column
'@details
'Collects a distinct list of values for the requested column. The column is
'read in one crossing and the duplicates are dropped in memory.
'@params
'- colName Column header
'@return BetterArray
Public Function UniqueValues(ByVal colName As String) As BetterArray
    'Downstream components use this information to populate combo boxes or
    'validate user input against known dictionary metadata.

    Dim values As BetterArray
    Dim columnValues As Variant
    Dim seen() As String
    Dim seenCount As Long
    Dim counter As Long
    Dim nextValue As String
    Dim firstRow As Long
    Dim lastRow As Long

    Set values = New BetterArray

    If Not ColumnExists(colName) Then
        LogInfo "Unable to compute unique values. Column " & colName & " not found", checkingWarning
        Set UniqueValues = values.Clone()
        Exit Function
    End If

    firstRow = StartRow() + 1
    lastRow = DictEndRow()

    If lastRow < firstRow Then
        Set UniqueValues = values.Clone()
        Exit Function
    End If

    columnValues = SheetColumnValues(ColumnIndex(colName), firstRow, lastRow)
    ReDim seen(1 To UBound(columnValues))

    For counter = 1 To UBound(columnValues)
        nextValue = TextOf(columnValues(counter))
        If ListPosition(seen, seenCount, nextValue) = 0 Then
            seenCount = seenCount + 1
            seen(seenCount) = nextValue
            values.Push nextValue
        End If
    Next

    Set UniqueValues = values.Clone()
End Function

'@label:sanitize-search-term
'@sub-title Escape the characters Excel reads as a search pattern
'@details
'`Range.Find` treats *, ?, [, ] and # as patterns, and a tilde as the escape.
'`TableSpecs` asks about composed names such as "adm1_" & rowVar, so a user
'variable name carrying one of those characters would match the wrong row.
'`LLVariables` and `LLSheets` each hold their own copy of this routine; folding
'the three into one place belongs to the session that owns those two files.
'@param searchText String. The raw search term.
'@return String. The escaped term.
Private Function SanitizeSearchTerm(ByVal searchText As String) As String
    Dim sanitized As String

    sanitized = Replace(searchText, "~", "~~")
    sanitized = Replace(sanitized, "*", "~*")
    sanitized = Replace(sanitized, "?", "~?")
    sanitized = Replace(sanitized, "[", "~[")
    sanitized = Replace(sanitized, "]", "~]")
    sanitized = Replace(sanitized, "#", "~#")

    SanitizeSearchTerm = sanitized
End Function

'@label:variableexists
'@prop-title Variable presence test
'@details
'Tests the dictionary for a specific variable identifier by scanning the
'"variable name" column. Used to guard operations that may attempt to insert
'duplicate entries.
'
'`LookIn` and `SearchOrder` are named because `Range.Find` reuses whatever the
'last search anywhere in the application used, the user's own Ctrl+F included.
'The match stays case-sensitive: `LLVariables.Contains` answers the same way and
'`TableSpecs` reads both.
'@params
'- varName Variable identifier
'@return Boolean
Public Function VariableExists(ByVal varName As String) As Boolean
    Dim varRange As Range
    Dim searchValue As String

    Set varRange = DataRange("variable name")

    If varRange Is Nothing Then
        LogInfo "Unable to locate variable name column while checking for variable " & varName, checkingError
        Exit Function
    End If

    If LenB(varName) = 0 Then Exit Function

    searchValue = SanitizeSearchTerm(varName)

    'Scoped to the search alone. Find raises on a range Excel cannot search,
    'and the answer for this class is then "the variable is absent".
    On Error Resume Next
        VariableExists = Not (varRange.Find(What:=searchValue, _
                                            LookAt:=xlWhole, _
                                            MatchCase:=True, _
                                            LookIn:=xlValues, _
                                            SearchOrder:=xlByRows) Is Nothing)
    On Error GoTo 0
End Function

'@label:addcolumn
'@sub-title Append a column at the end
'@details
'Adds a new column header after the current data block. No data is populated,
'allowing callers to later fill values or formulas as required.
'@params
'- colName Column header to append
Public Sub AddColumn(ByVal colName As String)
    Dim targetSheet As Worksheet
    Dim endCol As Long

    If Trim$(colName) = vbNullString Then Exit Sub

    Set targetSheet = Wksh()
    endCol = DictEndColumn()

    targetSheet.Cells(StartRow(), endCol + 1).Value = colName
    ClearPreparedCache
End Sub

'@label:insertcolumn
'@sub-title Insert a column relative to another
'@details
'Inserts a column immediately after the specified anchor, ensuring the header
'is merged with surrounding formatting when available so designer tables retain
'their layout.
'@params
'- colName Column to insert
'- After Existing column used as anchor
Public Sub InsertColumn(ByVal colName As String, ByVal After As String)
    Dim anchorRange As Range
    Dim headerCell As Range
    Dim hostSheet As Worksheet

    If colName = vbNullString Then Exit Sub

    If Not ColumnExists(After) Then
        LogInfo "Unable to insert column " & colName & " because anchor " & After & " does not exist", checkingWarning
        Exit Sub
    End If

    If ColumnExists(colName) Then
        LogInfo "Skipping insert of column " & colName & " because it already exists", checkingNote
        Exit Sub
    End If

    Set anchorRange = DataRange(After, includeHeaders:=True)
    Set hostSheet = Wksh()
    anchorRange.Offset(, 1).EntireColumn.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    anchorRange.Offset(, 1).Cells(1, 1).Value = colName

    'Scoped to the merge alone. A dictionary with no banner row above the
    'headers has no cells to merge, and the insert itself has already worked.
    On Error Resume Next
        Set headerCell = hostSheet.Range(anchorRange.Cells(-1, 1), anchorRange.Cells(0, 2))
        headerCell.Merge
    On Error GoTo 0

    ClearPreparedCache
End Sub

'@label:removecolumn
'@sub-title Remove a column by name
'@details
'Deletes the specified column from the dictionary, logging diagnostics when the
'column cannot be located to aid troubleshooting in designer flows.
'@params
'- colName Column header name
Public Sub RemoveColumn(ByVal colName As String)
    Dim columnRange As Range

    If Not ColumnExists(colName) Then
        LogInfo "Column " & colName & " was not found while attempting removal", checkingWarning
        Exit Sub
    End If

    Set columnRange = DataRange(colName)

    'Scoped to the delete alone. A protected worksheet refuses it, and the
    'caller is told through the checking log rather than by a raise.
    On Error Resume Next
        columnRange.EntireColumn.Delete
    On Error GoTo 0

    ClearPreparedCache
End Sub

'@label:RenameColumn
'@sub-title Rename a dictionary column header
Public Sub RenameColumn(ByVal currentName As String, _
                        ByVal newName As String, _
                        Optional ByVal strictSearch As Boolean = True, _
                        Optional ByVal matchCase As Boolean = False)

    If LenB(newName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "New column name cannot be empty"
    End If

    Data.RenameColumn currentName, newName, strictSearch, matchCase
    ClearPreparedCache
End Sub


'@section Row Operations
'===============================================================================

'@label:managerows
'@sub-title Add or remove rows on the dictionary listobject
'@details
'The dictionary listobject is taken as the first one on the worksheet, because
'this entry point has no selection to start from. `InsertRows` and `DeleteRows`
'work from the caller's selection and reach the table through it.
'Not every dictionary is table-backed; with no ListObject the attempt is logged.
'@param del Optional Boolean. True removes rows, False adds them.
Public Sub ManageRows(Optional ByVal del As Boolean = False)
    Dim hostSheet As Worksheet
    Dim listObj As ListObject
    Dim table As CustomTable

    Set hostSheet = Wksh()

    'Scoped to the lookup alone. A worksheet with no listobject raises here,
    'and the guard below turns that into a logged note.
    On Error Resume Next
        Set listObj = hostSheet.ListObjects(1)
    On Error GoTo 0

    If listObj Is Nothing Then
        LogInfo "No listobject available to " & IIf(del, "remove", "add") & " rows", checkingNote
        Exit Sub
    End If

    Set table = CustomTable.Create(listObj)
    If del Then
        table.RemoveRows
    Else
        table.AddRows
    End If
    If table.HasCheckings Then HarvestCollaboratorCheckings table.CheckingValues

    ClearPreparedCache
End Sub

'@label:InsertRows
'@sub-title Insert rows relative to a target selection
'@details
'Uses the worksheet selection height to determine how many rows to insert in the
'dictionary listobject. When the selection falls outside the table the request
'is logged and ignored to keep the structure stable.
'@param targetCell Range anchor describing the selection to mirror.
'@param insertShift Optional Boolean. When True worksheet rows are inserted to protect stacked tables.
Public Sub InsertRows(ByVal targetCell As Range, _
                      Optional ByVal insertShift As Boolean = False)
    Dim listObj As ListObject
    Dim table As CustomTable

    Set listObj = ResolveRowTarget(targetCell, "insert")
    If listObj Is Nothing Then Exit Sub

    Set table = CustomTable.Create(listObj)
    table.InsertRowsAt targetCell:=targetCell, insertShift:=insertShift
    If table.HasCheckings Then HarvestCollaboratorCheckings table.CheckingValues

    ClearPreparedCache
End Sub

'@label:DeleteRows
'@sub-title Delete rows intersecting the supplied selection
Public Sub DeleteRows(ByVal targetCell As Range, _
                      Optional ByVal includeIds As Boolean = True, _
                      Optional ByVal forceShift As Boolean = False)

    Dim listObj As ListObject
    Dim table As CustomTable

    Set listObj = ResolveRowTarget(targetCell, "delete")
    If listObj Is Nothing Then Exit Sub

    Set table = CustomTable.Create(listObj)
    table.DeleteRowsAt targetCell:=targetCell, _
                        includeIds:=includeIds, _
                        forceShift:=forceShift
    If table.HasCheckings Then HarvestCollaboratorCheckings table.CheckingValues

    ClearPreparedCache
End Sub

'@label:resolve-row-target
'@sub-title Resolve the listobject behind a selection
Private Function ResolveRowTarget(ByVal targetCell As Range, ByVal actionLabel As String) As ListObject
    'Validate that the supplied selection belongs to the dictionary worksheet and lives inside a listobject.
    Dim hostSheet As Worksheet
    Dim selectionSheet As Worksheet
    Dim normalizedAction As String

    normalizedAction = LCase$(actionLabel)
    Set hostSheet = Wksh()

    If targetCell Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the target selection is missing", checkingWarning
        Exit Function
    End If

    'Scoped to the worksheet read alone. It raises when the sheet holding the
    'selection has since been deleted, and the guard below reports that.
    On Error Resume Next
        Set selectionSheet = targetCell.Worksheet
    On Error GoTo 0

    If selectionSheet Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the selection worksheet could not be resolved", checkingWarning
        Exit Function
    End If

    If Not (selectionSheet Is hostSheet) Then
        LogInfo "Unable to " & normalizedAction & " rows because the selection belongs to " & selectionSheet.Name, checkingWarning
        Exit Function
    End If

    'Scoped to the listobject read alone. A selection outside any table raises,
    'and the guard below turns that into a logged warning.
    On Error Resume Next
        Set ResolveRowTarget = targetCell.ListObject
    On Error GoTo 0

    If ResolveRowTarget Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the selection is outside the dictionary listobject", checkingWarning
    End If
End Function


'@section Specialised Views
'===============================================================================

'@label:resolve-special-vars
'@sub-title Filter variables on a tag column
'@details
'Shared filtering routine behind `ChoicesVars`, `GeoVars`, and `TimeVars`. The
'function inspects the control column (or a custom column) and returns the
'variables that match the supplied criteria. Only variables living on
'horizontal sheets (`hlist2D`) are relevant for these helpers.
'
'The three columns are read as three blocks and matched in memory. `EventSetup`
'asks for five different filters back to back, so this used to walk the same
'three columns cell by cell five times.
Private Function ResolveSpecialVars(ByVal firstCondition As String, _
                                    Optional ByVal secondCondition As String = vbNullString, _
                                    Optional ByVal conditionName As String = "Control") As BetterArray
    Dim result As BetterArray
    Dim controlValues As Variant
    Dim variableValues As Variant
    Dim sheetValues As Variant
    Dim lastIndex As Long
    Dim idx As Long
    Dim controlValue As String
    Dim sheetValue As String

    Set result = New BetterArray

    If Not ColumnExists(conditionName) Then
        LogInfo "Unable to compute special vars. Column " & conditionName & " is missing", checkingWarning
        Set ResolveSpecialVars = result.Clone()
        Exit Function
    End If

    If Not ColumnExists("variable name") Or Not ColumnExists("sheet type") Then
        LogInfo "Unable to compute special vars because required metadata columns are missing", checkingWarning
        Set ResolveSpecialVars = result.Clone()
        Exit Function
    End If

    If (firstCondition = vbNullString) And (secondCondition = vbNullString) Then
        LogInfo "Unable to compute special vars because both conditions are missing", checkingWarning
        Set ResolveSpecialVars = result.Clone()
        Exit Function
    End If

    controlValues = ColumnValues(DataRange(conditionName))
    variableValues = ColumnValues(DataRange("variable name"))
    sheetValues = ColumnValues(DataRange("sheet type"))

    If Not (IsArray(controlValues) And IsArray(variableValues) And IsArray(sheetValues)) Then
        Set ResolveSpecialVars = result.Clone()
        Exit Function
    End If

    lastIndex = UBound(controlValues)
    If UBound(sheetValues) < lastIndex Then lastIndex = UBound(sheetValues)
    If UBound(variableValues) < lastIndex Then lastIndex = UBound(variableValues)

    For idx = 1 To lastIndex
        controlValue = TextOf(controlValues(idx))
        sheetValue = TextOf(sheetValues(idx))

        If sheetValue = SHEETTYPEHORIZONTAL Then
            If (controlValue = firstCondition) Or ((controlValue = secondCondition) And (secondCondition <> vbNullString)) Then
                result.Push TextOf(variableValues(idx))
            End If
        End If
    Next

    Set ResolveSpecialVars = result.Clone()
End Function

'@label:specialvars
'@prop-title Filtered variable list by criteria
'@details
'Returns variables that match one or two control values within a chosen column.
'Only considers variables on horizontal sheets to align with designer
'expectations.
'@params
'- firstCondition Primary tag to match
'- secondCondition Optional secondary tag
'- conditionName Optional column header that stores the tags. Defaults to
'  "Control"
Public Property Get SpecialVars(ByVal firstCondition As String, _
                                Optional ByVal secondCondition As String = vbNullString, _
                                Optional ByVal conditionName As String = "Control") As BetterArray
    Set SpecialVars = ResolveSpecialVars(firstCondition, secondCondition, conditionName)
End Property

'@label:choicesvars
'@prop-title Choice variables list
'@details
'Shortcut returning variables controlled by manual or formula-driven choice
'lists.
'@return BetterArray
Public Property Get ChoicesVars() As BetterArray
    Set ChoicesVars = ResolveSpecialVars("choice_manual", "choice_formula")
End Property

'@label:geovars
'@prop-title Geographic variables list
'@details
'Convenience accessor for variables whose controls drive geographic selections
'(geo/hf).
'@return BetterArray
Public Property Get GeoVars() As BetterArray
    Set GeoVars = ResolveSpecialVars("geo", "hf")
End Property

'@label:timevars
'@prop-title Time variables list
'@details
'Returns variables tied to dates by filtering through the `Variable Type`
'column.
'@return BetterArray
Public Property Get TimeVars() As BetterArray
    Set TimeVars = ResolveSpecialVars("date", conditionName:="Variable Type")
End Property


'@section Preparation
'===============================================================================

'@label:clear-prepared
'@sub-title Drop the held preparation answer
'@details
'Called by every routine that adds, removes or renames a column, or moves the
'last data row, because `Prepared` reads six column names, two forbidden ones
'and the font colour of the cell below the data.
Private Sub ClearPreparedCache()
    this.preparedValue = False
End Sub

'@label:prepared
'@prop-title Dictionary preparation flag
'@details
'Indicates whether all helper columns have been generated for designer
'workflows. The heuristic mirrors the legacy implementation by inspecting
'specific markers instead of tracking explicit state.
'
'A True answer is held on the object. VBA evaluates every operand of `And`, so
'one read costs eight column searches plus a font read, and
'`Formulas.ParsedLinelistFormula` reaches it once per token in every formula in
'the workbook through `LLSheets.VariableAddress`. That hot path always runs on a
'prepared dictionary, so holding True is where the whole win is.
'
'A False answer is read again every time. Preparation is what turns False into
'True, and another object over the same worksheet can be the one that runs it,
'so a held False could go stale with nothing to drop it. Every routine here that
'can turn True back into False drops the held answer itself.
'@return Boolean
Public Property Get Prepared() As Boolean
    Dim hostSheet As Worksheet

    If this.preparedValue Then
        Prepared = True
        Exit Property
    End If

    Set hostSheet = Wksh()
    this.preparedValue = (hostSheet.Cells(DictEndRow() + 1, StartColumn()).Font.color = vbBlue) And _
               ColumnExists("table name") And ColumnExists("column index") And _
               ColumnExists("visibility") And ColumnExists("crf index") And _
               ColumnExists("crf choices") And ColumnExists("crf status") And _
               (Not ColumnExists("main section sort")) And (Not ColumnExists("main section index"))

    Prepared = this.preparedValue
End Property

'@label:prepare
'@sub-title Prepare dictionary for designer workflows
'@details
'Extends the dictionary with all helper metadata required by the designer:
'table names, CRF indices, geo duplications, and visibility columns. The method
'guards against repeated execution by checking the `Prepared` flag and logs
'issues discovered during enrichment.
'
'The order of operations is fixed: exported workbooks depend on it. The sheet is
'sorted twice, once by `AppendNumberColumn` to group the rows by sheet and again
'inside `AddMainSectionSort` to order each sheet by main section. The second
'sort numbers its groups in first-seen order, so it needs the rows already
'grouped by sheet. Both sorts stay.
'@params
'- preservedSheetNames BetterArray of sheet names that should keep their
'  original values
'- geoObject LLGeo instance required for geo expansions. Taken as Object and
'  reached late.
'- tablePrefix Optional String prefix applied to generated table names
Public Sub Prepare(Optional ByVal preservedSheetNames As BetterArray, _
                   Optional ByVal geoObject As Object, _
                   Optional ByVal tablePrefix As String = "table")

    'Augment the user dictionary with all helper columns required by the
    'designer (table names, CRF metadata, geo expansions, etc.).
    Dim controlRange As Range

    If Prepared Then Exit Sub

    EnsureRequiredColumns

    If Not ColumnExists("table name") Then AppendNumberColumn "sheet name", tablePrefix, "table name"
    If Not ColumnExists("column index") Then AppendColumnIndex

    Set controlRange = DataRange("control")

    If Not controlRange Is Nothing Then
        If TypeName(geoObject) <> "LLGeo" Then
            LogInfo "Geo object should be of type LLGeo, geolines not append", checkingError
        Else
            If ColumnExists("control") Then
                If Not controlRange.Find(What:="geo", lookAt:=xlWhole, MatchCase:=True) Is Nothing Then
                    AppendGeoLines geoObject
                End If

                'Read the control column again. AppendGeoLines inserts eleven
                'rows per geo variable, and the range taken above stops where
                'the dictionary used to end.
                Set controlRange = DataRange("control")

                If Not controlRange Is Nothing Then
                    If Not controlRange.Find(What:="hf", lookAt:=xlWhole, MatchCase:=True) Is Nothing Then
                        AppendHFLines
                    End If
                End If
            End If
        End If
    End If

    UniqueVarNames
    GoodSheetNames preservedSheetNames

    If Not ColumnExists("visibility") Then AddVisibility
    If Not ColumnExists("crf index") Then AppendCRFLineIndex

    Clean removeAddedColumns:=False

    With Wksh()
        .Cells(DictEndRow() + 1, StartColumn()).Font.color = vbBlue
    End With

    ClearPreparedCache
End Sub

'@label:EnsureRequiredColumns
'@sub-title Ensure required columns are present in the dictionary
Private Sub EnsureRequiredColumns()
    Dim colName As String
    Dim counter As Long
    Dim requiredCols As BetterArray

    Set requiredCols = New BetterArray
    requiredCols.Items = Split(DEFAULTDICTIONARYCOLUMNS, "|")

    For counter = requiredCols.LowerBound To requiredCols.UpperBound
        colName = requiredCols.Item(counter)
        'Check existence for all columns except exports
        If Not ColumnExists(colName) Then
            ThrowError ProjectError.ElementNotFound, "Column -" & colName & "- not found in the dictionary"
        End If
    Next
End Sub

'@label:valid-dictionary-column
'@sub-title Decide in memory whether a header belongs to the dictionary
'@details
'Same rule as `ColumnExists` with `checkValidity` set: the header must sit in
'the dictionary header row, and it must either be in the allowed list or start
'with the word "export". The two log messages match the ones `ColumnExists`
'writes, so a cleaned workbook still reports the same notes.
'@param columnName String. The header, already lower-cased.
'@param knownHeaders String array. Lower-cased headers of the dictionary block.
'@param knownCount Long. How many entries of that array are filled.
'@return Boolean. True when the column is kept.
Private Function ValidDictionaryColumn(ByVal columnName As String, _
                                       ByRef knownHeaders() As String, _
                                       ByVal knownCount As Long) As Boolean
    If ListPosition(knownHeaders, knownCount, columnName) = 0 Then
        LogInfo "Column " & columnName & " is missing from dictionary", checkingWarning
        Exit Function
    End If

    EnsureColumnsInitialized
    ValidDictionaryColumn = this.allowedColumns.Includes(columnName) Or (Split(columnName, " ")(0) = "export")

    If Not ValidDictionaryColumn Then
        LogInfo "Column " & columnName & " exists but is not part of the allowed dictionary columns", checkingWarning
    End If
End Function

'@label:clean
'@sub-title Remove unexpected columns
'@details
'Deletes columns that fall outside the supported dictionary schema. When the
'optional argument is True, helper columns introduced during preparation are
'removed as well so the dictionary reverts to the user-provided structure.
'
'The header row is read once and every decision is taken in memory. Asking the
'worksheet for each header and then searching the worksheet again for its name
'cost two crossings a column.
'
'The walk starts at the dictionary start column. It used to start at column 1,
'so for a dictionary that does not begin in column A the columns to its left
'were read as headers and deleted.
'
'This deletes ENTIRE worksheet columns. Anything else living on the Dictionary
'sheet to the right of the table goes with them.
'@params
'- removeAddedColumns Optional Boolean. Remove preparation columns too.
Public Sub Clean(Optional ByVal removeAddedColumns As Boolean = False)
    Dim hostSheet As Worksheet
    Dim usedCols As Long
    Dim firstCol As Long
    Dim dataEndCol As Long
    Dim headerRow As Long
    Dim headerValues As Variant
    Dim knownHeaders() As String
    Dim knownCount As Long
    Dim colIndex As Long
    Dim counter As Long
    Dim columnName As String
    Dim removeAdded As Boolean

    Set hostSheet = Wksh()
    EnsureColumnsInitialized

    firstCol = StartColumn()
    headerRow = StartRow()
    usedCols = hostSheet.UsedRange.Columns.Count + hostSheet.UsedRange.Column - 1

    If usedCols < firstCol Then Exit Sub

    dataEndCol = DictEndColumn()
    headerValues = RowValues(hostSheet.Range(hostSheet.Cells(headerRow, firstCol), _
                                             hostSheet.Cells(headerRow, usedCols)))

    'Headers inside the dictionary block are what `Data.ColumnExists` can see.
    'A header sitting past the last data column reads as absent, which is how a
    'stray column beyond the table is removed.
    ReDim knownHeaders(1 To UBound(headerValues))
    For counter = 1 To UBound(headerValues)
        If firstCol + counter - 1 <= dataEndCol Then
            columnName = LCase$(TextOf(headerValues(counter)))
            If LenB(columnName) > 0 Then
                knownCount = knownCount + 1
                knownHeaders(knownCount) = columnName
            End If
        End If
    Next

    For colIndex = usedCols To firstCol Step -1
        columnName = LCase$(TextOf(headerValues(colIndex - firstCol + 1)))
        removeAdded = removeAddedColumns And this.addedColumns.Includes(columnName)

        If (Not ValidDictionaryColumn(columnName, knownHeaders, knownCount)) Or removeAdded Then
            hostSheet.Columns(colIndex).EntireColumn.Delete
        End If
    Next

    ClearPreparedCache
End Sub

'@label:good-sheet-names
'@sub-title Suffix sheet names that clash with reserved names
'@details
'The sheet-name column is read once and written back only when a name really
'moved, so a dictionary with nothing to rename keeps its cells untouched.
Private Sub GoodSheetNames(ByVal preservedSheetNames As BetterArray)
    Dim sheetNameIndex As Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim sheetValues As Variant
    Dim counter As Long
    Dim currValue As String
    Dim changed As Boolean

    If preservedSheetNames Is Nothing Then Exit Sub

    firstRow = StartRow() + 1
    lastRow = DictEndRow()
    If lastRow < firstRow Then Exit Sub

    sheetNameIndex = ColumnIndex("sheet name")
    sheetValues = SheetColumnValues(sheetNameIndex, firstRow, lastRow)

    For counter = 1 To UBound(sheetValues)
        currValue = TextOf(sheetValues(counter))
        If preservedSheetNames.Includes(currValue) Then
            sheetValues(counter) = currValue & "_"
            changed = True
            LogInfo "Sheet name " & currValue & " duplicated and suffixed with underscore during preparation", checkingNote
        End If
    Next

    If changed Then WriteSheetColumn sheetNameIndex, firstRow, sheetValues
End Sub

'@label:normalise-variable-name
'@sub-title Clean one variable name in plain VBA
'@details
'Reproduces the three worksheet functions this used to call, with no crossing
'into Excel: U+00A0 NO-BREAK SPACE becomes a space, codes below 32 are dropped (worksheet
'CLEAN), the ends are trimmed and runs of spaces inside the text collapse to one
'(worksheet TRIM does that as well as trimming), then the remaining spaces
'become underscores.
'
'AscW answers a NEGATIVE number above code point 32767, so the control-code test
'guards on zero as well.
'@param rawName String. The value read from the variable name cell.
'@return String. The normalised identifier.
Private Function NormaliseVariableName(ByVal rawName As String) As String
    Dim source As String
    Dim result As String
    Dim counter As Long
    Dim charCode As Long
    Dim oneChar As String
    Dim previousWasSpace As Boolean

    'ChrW$ reads 160 as a Unicode code point, so U+00A0 comes back on every
    'machine. Chr$ would read it as a byte position in the machine's ANSI
    'codepage, which on this Mac holds the dagger.
    source = Replace(rawName, ChrW$(160), " ")

    For counter = 1 To Len(source)
        oneChar = Mid$(source, counter, 1)
        charCode = AscW(oneChar)

        If charCode >= 0 And charCode < 32 Then
            'Dropped, the same as the worksheet CLEAN function.
        ElseIf oneChar = " " Then
            If Not previousWasSpace Then
                result = result & " "
                previousWasSpace = True
            End If
        Else
            result = result & oneChar
            previousWasSpace = False
        End If
    Next

    NormaliseVariableName = Replace(Trim$(result), " ", "_")
End Function

'@label:unique-var-names
'@sub-title Normalise variable names and force uniqueness
'@details
'Trims whitespace, replaces spaces with underscores, and appends an underscore
'when two rows end up with the same identifier.
'
'Every short name is collected and reported in ONE message, and nothing is
'written when one is found. Raising on the first bad row left the rows above
'normalised and the rows below untouched, so the user was told about one row and
'handed a half-rewritten dictionary.
Private Sub UniqueVarNames()
    Dim firstRow As Long
    Dim lastRow As Long
    Dim varIndex As Long
    Dim varValues As Variant
    Dim seen() As String
    Dim seenCount As Long
    Dim counter As Long
    Dim varName As String
    Dim shortRows As String
    Dim shortCount As Long

    EnsureColumnsInitialized

    firstRow = StartRow() + 1
    lastRow = DictEndRow()
    If lastRow < firstRow Then Exit Sub

    varIndex = ColumnIndex("variable name")
    varValues = SheetColumnValues(varIndex, firstRow, lastRow)
    ReDim seen(1 To UBound(varValues))

    For counter = 1 To UBound(varValues)
        varName = NormaliseVariableName(TextOf(varValues(counter)))

        If Len(varName) < 4 Then
            shortCount = shortCount + 1
            If shortCount <= MAXREPORTEDROWS Then
                If LenB(shortRows) > 0 Then shortRows = shortRows & ", "
                shortRows = shortRows & CStr(firstRow + counter - 1)
            End If
        End If

        If ListPosition(seen, seenCount, varName) > 0 Then varName = varName & "_"
        seenCount = seenCount + 1
        seen(seenCount) = varName
        varValues(counter) = varName
    Next

    If shortCount > 0 Then
        If shortCount > MAXREPORTEDROWS Then
            shortRows = shortRows & " and " & CStr(shortCount - MAXREPORTEDROWS) & " more"
        End If
        ThrowError ProjectError.InvalidArgument, _
                   "Variable names should be at least 4 characters. Please review variables at row " & shortRows
    End If

    WriteSheetColumn varIndex, firstRow, varValues
End Sub

'@label:add-visibility
'@sub-title Derive the visibility and CRF status columns
'@details
'The logic mirrors the designer wizard so that prepared dictionaries behave the
'same whether produced manually or through automation. The status and control
'columns are read as two blocks and the two answers are written as two blocks.
'
'The status decides both columns, and then the CONTROL has the last word on the
'CRF one: a formula, a choice_formula, a case_when and a list_auto read
'"always hidden" there whatever their status is.
'
'A row whose status is filled in with something this dictionary does not know
'keeps both cells empty and the value is logged. A row with no status at all
'keeps them empty and is passed over in silence, since there is nothing there
'for anybody to have got wrong.
Private Sub AddVisibility()
    Dim firstRow As Long
    Dim lastRow As Long
    Dim statusIdx As Long
    Dim visibilityIdx As Long
    Dim crfStatusIdx As Long
    Dim controlIdx As Long
    Dim statusValues As Variant
    Dim controlValues As Variant
    Dim visibilityValues As Variant
    Dim crfStatusValues As Variant
    Dim counter As Long
    Dim statusValue As String
    Dim controlValue As String

    AddColumn "visibility"
    AddColumn "crf status"

    firstRow = StartRow() + 1
    lastRow = DictEndRow()
    If lastRow < firstRow Then Exit Sub

    statusIdx = ColumnIndex("status")
    visibilityIdx = ColumnIndex("visibility")
    crfStatusIdx = ColumnIndex("crf status")
    controlIdx = ColumnIndex("control")

    statusValues = SheetColumnValues(statusIdx, firstRow, lastRow)
    controlValues = SheetColumnValues(controlIdx, firstRow, lastRow)

    ReDim visibilityValues(1 To UBound(statusValues))
    ReDim crfStatusValues(1 To UBound(statusValues))

    For counter = 1 To UBound(statusValues)
        statusValue = TextOf(statusValues(counter))
        controlValue = TextOf(controlValues(counter))

        'A variable the designer gave no status at all keeps both cells empty
        'and says nothing about it. An empty cell is not a value anybody wrote
        'wrong, so the note below is kept for a status that IS filled in and is
        'none of the ones this dictionary knows.
        If LenB(Trim$(statusValue)) > 0 Then
            Select Case statusValue
                Case "hidden"
                    visibilityValues(counter) = "always hidden"
                    crfStatusValues(counter) = "always hidden"
                Case "mandatory"
                    visibilityValues(counter) = "always visible"
                    crfStatusValues(counter) = "visible on creation"
                Case "optional, hidden"
                    visibilityValues(counter) = "hidden on creation"
                    crfStatusValues(counter) = "hidden on creation"
                Case "optional, visible", "optional"
                    visibilityValues(counter) = "visible on creation"
                    crfStatusValues(counter) = "visible on creation"
                Case Else
                    LogInfo "Unhandled status value " & statusValue & " at row " & CStr(firstRow + counter - 1), checkingNote
            End Select
        End If

        'Whatever the status says. A calculated column has nothing to give on a
        'form somebody fills in by hand, and an automatic list has nothing but
        'values already typed elsewhere, so neither belongs on a CRF at any
        'status -- mandatory, optional or unwritten. This used to sit inside
        'the mandatory branch alone, which left an optional formula reading
        '"visible on creation" here while the CRF held it hidden anyway.
        'ShowHide.IsLockedStatus names the same four controls and holds them
        'hidden on the printed sheet as well, which has no derived column of
        'its own: the printed layer reads `register book`, written by hand.
        'The two lists have to be changed together.
        Select Case controlValue
            Case "formula", "choice_formula", "case_when", "list_auto"
                crfStatusValues(counter) = "always hidden"
        End Select
    Next

    WriteSheetColumn visibilityIdx, firstRow, visibilityValues
    WriteSheetColumn crfStatusIdx, firstRow, crfStatusValues
End Sub

'@label:append-geo-lines
'@sub-title Expand every geo control variable into twelve rows
'@details
'Each "geo" control variable becomes 12 contiguous rows:
'  4 admin levels   (adm1-4) with geo1-4 controls
'  4 p-codes        (pcode_adm1-4) hidden, formula-driven
'  4 concatenations (concat_adm1-4) hidden, formula-driven
'
'The original geo row is modified in place to become adm1 and 11 new rows are
'inserted right after it. `DictEndRow()` is read on every pass ON PURPOSE, so
'the search range follows the rows this routine keeps inserting. Do not turn
'this into a single block read.
'
'`geoObject.GeoNames(...)` is read per level; `LLGeo` holds those in its own
'cache, so the repeat costs nothing here.
Private Sub AppendGeoLines(ByVal geoObject As Object)
    Const ADMIN_LEVELS As Long = 4
    Const GEO_EXPANSION As Long = 12
    Const INSERTED_ROWS As Long = 11

    Dim hostSheet As Worksheet
    Dim controlIdx As Long
    Dim varIdx As Long
    Dim mainLabelIdx As Long
    Dim subSectionIdx As Long
    Dim registerIdx As Long
    Dim statusIdx As Long
    Dim positionIdx As Long
    Dim controlDetailsIdx As Long
    Dim endCol As Long
    Dim geoRow As Long
    Dim insertRow As Long
    Dim varName As String
    Dim varPosition As Long
    Dim baseValues As Variant
    Dim searchRange As Range
    Dim foundCell As Range
    Dim targetRange As Range
    Dim counter As Long
    Dim fillRow As Long
    Dim formulaValue As String

    Set hostSheet = Wksh()
    controlIdx = ColumnIndex("control")
    varIdx = ColumnIndex("variable name")
    mainLabelIdx = ColumnIndex("main label")
    subSectionIdx = ColumnIndex("sub section")
    registerIdx = ColumnIndex("register book")
    statusIdx = ColumnIndex("status")
    positionIdx = ColumnIndex("column index")
    controlDetailsIdx = ColumnIndex("control details")
    endCol = DictEndColumn()
    geoRow = StartRow() + 1

    Do
        'Find next "geo" control from current position downward
        Set searchRange = hostSheet.Range( _
            hostSheet.Cells(geoRow, controlIdx), _
            hostSheet.Cells(DictEndRow(), controlIdx))
        Set foundCell = searchRange.Find(What:="geo", lookAt:=xlWhole, MatchCase:=True)
        If foundCell Is Nothing Then Exit Do

        geoRow = foundCell.Row
        varName = hostSheet.Cells(geoRow, varIdx).Value
        varPosition = hostSheet.Cells(geoRow, positionIdx).Value

        'Move main label into sub section when sub section is blank
        If IsEmpty(hostSheet.Cells(geoRow, subSectionIdx)) And _
           (Not IsEmpty(hostSheet.Cells(geoRow, mainLabelIdx))) Then
            hostSheet.Cells(geoRow, subSectionIdx).Value = hostSheet.Cells(geoRow, mainLabelIdx).Value
            hostSheet.Cells(geoRow, mainLabelIdx).Value = vbNullString
        End If

        'Snapshot base row values before modification
        baseValues = hostSheet.Range( _
            hostSheet.Cells(geoRow, StartColumn()), _
            hostSheet.Cells(geoRow, endCol)).Value

        'Modify original row to adm1
        hostSheet.Cells(geoRow, varIdx).Value = "adm1_" & varName
        hostSheet.Cells(geoRow, mainLabelIdx).Value = geoObject.GeoNames("adm1_name")
        hostSheet.Cells(geoRow, controlIdx).Value = "geo1"

        'Insert 11 rows right after adm1
        insertRow = geoRow + 1
        hostSheet.Rows(insertRow & ":" & (insertRow + INSERTED_ROWS - 1)).Insert _
            Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        'Populate all inserted rows with base values
        For counter = 0 To INSERTED_ROWS - 1
            Set targetRange = hostSheet.Range( _
                hostSheet.Cells(insertRow + counter, StartColumn()), _
                hostSheet.Cells(insertRow + counter, endCol))
            targetRange.Value = baseValues
        Next

        'Admin levels 2-4
        For counter = 2 To ADMIN_LEVELS
            fillRow = insertRow + counter - 2
            hostSheet.Cells(fillRow, varIdx).Value = "adm" & counter & "_" & varName
            hostSheet.Cells(fillRow, controlIdx).Value = "geo" & counter
            hostSheet.Cells(fillRow, mainLabelIdx).Value = geoObject.GeoNames("adm" & counter & "_name")
            hostSheet.Cells(fillRow, positionIdx).Value = varPosition + counter - 1
        Next

        'P-codes 1-4 (offset past 3 admin rows)
        For counter = 1 To ADMIN_LEVELS
            fillRow = insertRow + (ADMIN_LEVELS - 1) + counter - 1
            hostSheet.Cells(fillRow, statusIdx).Value = "hidden"
            hostSheet.Cells(fillRow, registerIdx).Value = "hidden"
            hostSheet.Cells(fillRow, controlIdx).Value = "formula"
            hostSheet.Cells(fillRow, mainLabelIdx).Value = "geo_pcode_adm" & counter
            hostSheet.Cells(fillRow, varIdx).Value = "pcode_adm" & counter & "_" & varName
            hostSheet.Cells(fillRow, positionIdx).Value = varPosition + ADMIN_LEVELS + counter - 1
            formulaValue = "INDEX(T_ADM" & counter & ", MATCH(concat_adm" & counter & "_" & varName & _
                           ", adm" & counter & "_concat, 0), " & (counter + 2) & ")"
            hostSheet.Cells(fillRow, controlDetailsIdx).Value = _
                "IF(ISNA(" & formulaValue & "), """", " & formulaValue & ")"
        Next

        'Concatenations 1-4 (offset past admin + pcode rows)
        For counter = 1 To ADMIN_LEVELS
            fillRow = insertRow + (ADMIN_LEVELS - 1) + ADMIN_LEVELS + counter - 1
            hostSheet.Cells(fillRow, statusIdx).Value = "hidden"
            hostSheet.Cells(fillRow, registerIdx).Value = "hidden"
            hostSheet.Cells(fillRow, controlIdx).Value = "formula"
            hostSheet.Cells(fillRow, mainLabelIdx).Value = "geo_concat_adm" & counter
            hostSheet.Cells(fillRow, varIdx).Value = "concat_adm" & counter & "_" & varName
            hostSheet.Cells(fillRow, positionIdx).Value = varPosition + (2 * ADMIN_LEVELS) + counter - 1
            hostSheet.Cells(fillRow, controlDetailsIdx).Value = _
                "GEOCONCAT(adm1_" & varName & ", " & counter & ")"
        Next

        'Advance past all expanded rows
        geoRow = geoRow + GEO_EXPANSION
    Loop

    ClearPreparedCache
End Sub

'@label:append-hf-lines
'@sub-title Give every health facility variable its hidden p-code row
'@details
'A health facility variable occupies TWO worksheet columns, and until now it
'only ever wrote one. AppendColumnIndex reserves the pair -- its `Case "hf"`
'advances the next variable by 2 -- and LLSpatial.HF_COLUMN_SPAN agrees the
'block is two wide. Nothing filled the second. LLSheets sizes the data table to
'`max(variable count, largest column index)`, so the table stretched over a
'header cell nobody had written and Excel named that column "Column 1" itself.
'A phantom column in every linelist carrying a health facility.
'
'This is the geo expansion cut to its smallest form: one inserted row rather
'than eleven, hidden the same way, controlled by "formula" the same way. The
'p-code is looked up on the facility column ALONE, matched against `hf_concat`
'and read out of `hf_pcode`, both of which LLGeo publishes off T_HF beside the
'four admin pairs.
'
'The variable name is used exactly as the dictionary holds it. The geo branch
'rewrites its own row to `adm1_<name>` and can build prefixed names from what
'it started with; no such rewrite happens to a health facility row, so the name
'in the cell is the name VarWriter puts in the header, and it is the only
'spelling that will match the column at run time.
Private Sub AppendHFLines()
    Const HF_EXPANSION As Long = 2

    Dim hostSheet As Worksheet
    Dim controlIdx As Long
    Dim varIdx As Long
    Dim mainLabelIdx As Long
    Dim registerIdx As Long
    Dim statusIdx As Long
    Dim positionIdx As Long
    Dim controlDetailsIdx As Long
    Dim endCol As Long
    Dim hfRow As Long
    Dim insertRow As Long
    Dim varName As String
    Dim varPosition As Long
    Dim baseValues As Variant
    Dim searchRange As Range
    Dim foundCell As Range
    Dim targetRange As Range
    Dim formulaValue As String

    Set hostSheet = Wksh()
    controlIdx = ColumnIndex("control")
    varIdx = ColumnIndex("variable name")
    mainLabelIdx = ColumnIndex("main label")
    registerIdx = ColumnIndex("register book")
    statusIdx = ColumnIndex("status")
    positionIdx = ColumnIndex("column index")
    controlDetailsIdx = ColumnIndex("control details")
    endCol = DictEndColumn()
    hfRow = StartRow() + 1

    Do
        'Find the next "hf" control from the current position downward.
        Set searchRange = hostSheet.Range( _
            hostSheet.Cells(hfRow, controlIdx), _
            hostSheet.Cells(DictEndRow(), controlIdx))
        Set foundCell = searchRange.Find(What:="hf", lookAt:=xlWhole, MatchCase:=True)
        If foundCell Is Nothing Then Exit Do

        hfRow = foundCell.Row
        varName = hostSheet.Cells(hfRow, varIdx).Value
        varPosition = hostSheet.Cells(hfRow, positionIdx).Value

        'Snapshot the row before it gains a neighbour, so the inserted row
        'starts from the same sheet name, section and table name.
        baseValues = hostSheet.Range( _
            hostSheet.Cells(hfRow, StartColumn()), _
            hostSheet.Cells(hfRow, endCol)).Value

        insertRow = hfRow + 1
        hostSheet.Rows(insertRow).Insert _
            Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        Set targetRange = hostSheet.Range( _
            hostSheet.Cells(insertRow, StartColumn()), _
            hostSheet.Cells(insertRow, endCol))
        targetRange.Value = baseValues

        hostSheet.Cells(insertRow, statusIdx).Value = "hidden"
        hostSheet.Cells(insertRow, registerIdx).Value = "hidden"
        hostSheet.Cells(insertRow, controlIdx).Value = "formula"
        hostSheet.Cells(insertRow, mainLabelIdx).Value = "geo_pcode_hf"
        hostSheet.Cells(insertRow, varIdx).Value = "pcode_" & varName
        hostSheet.Cells(insertRow, positionIdx).Value = varPosition + 1

        'INDEX over the named p-code column rather than T_HF with a column
        'number. The geo rows count columns into T_ADM<n>; here the name LLGeo
        'published is used directly, so a reordered geobase cannot silently
        'return the wrong field.
        formulaValue = "INDEX(hf_pcode, MATCH(" & varName & ", hf_concat, 0))"
        hostSheet.Cells(insertRow, controlDetailsIdx).Value = _
            "IF(ISNA(" & formulaValue & "), """", " & formulaValue & ")"

        'Past the pair.
        hfRow = hfRow + HF_EXPANSION
    Loop

    ClearPreparedCache
End Sub

'@label:append-number-column
'@sub-title Number each distinct group and sort the dictionary on it
'@details
'Assigns an incrementing number to each distinct sheet (or user-defined group)
'so the dictionary can later produce stable table names, sorts the whole block
'on that number, then writes the visible name column as a formula.
'
'The source column index is resolved ONCE. It used to be resolved inside the row
'loop, and `DataSheet.ColumnIndex` is a header scan plus a search, so an N-row
'pass cost N searches.
'
'The name column stays a formula write: the formula is written, autofilled and
'then flattened on purpose (`sortRange.Value = sortRange.Value`).
Private Sub AppendNumberColumn(Optional ByVal onColName As String = "sheet name", _
                               Optional ByVal tablePrefix As String = "table", _
                               Optional ByVal newColumn As String = "table name")
    Dim hostSheet As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim endCol As Long
    Dim sourceIdx As Long
    Dim tableOrder As Long
    Dim sourceValues As Variant
    Dim numbers As Variant
    Dim seen() As String
    Dim seenOrder() As Long
    Dim seenCount As Long
    Dim counter As Long
    Dim position As Long
    Dim groupName As String
    Dim sortRange As Range

    Set hostSheet = Wksh()
    firstRow = StartRow() + 1
    lastRow = DictEndRow()
    endCol = DictEndColumn()
    sourceIdx = ColumnIndex(onColName)

    hostSheet.Cells(StartRow(), endCol + 1).Value = "number_" & newColumn

    If lastRow >= firstRow Then
        sourceValues = SheetColumnValues(sourceIdx, firstRow, lastRow)
        ReDim numbers(1 To UBound(sourceValues))
        ReDim seen(1 To UBound(sourceValues))
        ReDim seenOrder(1 To UBound(sourceValues))

        For counter = 1 To UBound(sourceValues)
            groupName = TextOf(sourceValues(counter))
            position = ListPosition(seen, seenCount, groupName)

            If position = 0 Then
                tableOrder = tableOrder + 1
                seenCount = seenCount + 1
                seen(seenCount) = groupName
                seenOrder(seenCount) = tableOrder
                numbers(counter) = tableOrder
            Else
                numbers(counter) = seenOrder(position)
            End If
        Next

        WriteSheetColumn endCol + 1, firstRow, numbers
    End If

    Set sortRange = DataRange(includeHeaders:=True)
    sortRange.Sort key1:=DataRange("number_" & newColumn), Order1:=xlAscending, Header:=xlYes

    'The number column moved the last data column, so it is read again here.
    endCol = DictEndColumn()
    hostSheet.Cells(StartRow(), endCol + 1).Value = newColumn
    hostSheet.Cells(StartRow() + 1, endCol + 1).Formula = "= " & Chr$(34) & tablePrefix & Chr$(34) & " " & Chr$(38) & " " & hostSheet.Cells(StartRow() + 1, endCol).Address(RowAbsolute:=False)
    Set sortRange = DataRange(newColumn)
    hostSheet.Cells(StartRow() + 1, endCol + 1).AutoFill sortRange, Type:=xlFillValues
    hostSheet.UsedRange.Calculate
    sortRange.Value = sortRange.Value

    ClearPreparedCache
End Sub

'@label:add-main-section-sort
'@sub-title Add the main section sort helper and order the sheet on it
'@details
'Writes a formula column holding "<sheet name> - <main section>", or the row
'number when the main section is empty, then numbers and sorts on it through
'`AppendNumberColumn`. The formula write and its autofill stay as they are.
Private Sub AddMainSectionSort()
    Dim hostSheet As Worksheet
    Dim sheetNameAddress As String
    Dim mainSecAddress As String
    Dim sortFormula As String
    Dim sortRange As Range
    Dim mainSectionIdx As Long
    Dim sheetNameIdx As Long
    Dim endCol As Long

    Set hostSheet = Wksh()
    mainSectionIdx = ColumnIndex("main section")
    sheetNameIdx = ColumnIndex("sheet name")
    endCol = DictEndColumn()

    hostSheet.Cells(StartRow(), endCol + 1).Value = "main section sort"
    sheetNameAddress = hostSheet.Cells(StartRow() + 1, sheetNameIdx).Address(RowAbsolute:=False) & " "
    mainSecAddress = hostSheet.Cells(StartRow() + 1, mainSectionIdx).Address(RowAbsolute:=False)

    sortFormula = "= IF(" & mainSecAddress & "=" & Chr$(34) & Chr$(34) & _
                  ", ROW(" & mainSecAddress & "), " & sheetNameAddress & _
                  Chr$(38) & Chr$(34) & " - " & Chr$(34) & Chr$(38) & _
                  mainSecAddress & ")"

    hostSheet.Cells(StartRow() + 1, endCol + 1).Formula = sortFormula
    Set sortRange = DataRange("main section sort")
    hostSheet.Cells(StartRow() + 1, endCol + 1).AutoFill sortRange, Type:=xlFillValues
    sortRange.Calculate
    AppendNumberColumn "main section sort", "main section", "main section index"
End Sub

'@label:sort
'@sub-title Sort the dictionary by sheet and main section
'@details
'Reached from the designer Sort button on the setup Dictionary sheet, through
'`EventSetup.SortTables`.
'
'The clean afterwards keeps the generated columns. The three helper columns the
'sort writes ("main section sort", "main section index" and its number_ helper)
'sit outside the allowed list, so the plain clean already removes them. Passing
'True here also deleted table name, column index, visibility, crf index, crf
'choices, crf status and list auto, so a sort on a prepared dictionary undid the
'preparation.
'@export
Public Sub Sort()
    AddMainSectionSort
    Clean removeAddedColumns:=False
End Sub

'@label:append-column-index
'@sub-title Compute the final column order used by the designer
'@details
'Sheet-level grouping is respected and gaps are inserted for geo/hf variables
'that expand into several columns. The three source columns are read as three
'blocks and the answer is written as one block.
Private Sub AppendColumnIndex()
    Dim hostSheet As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim endCol As Long
    Dim sheetNameIdx As Long
    Dim sheetTypeIdx As Long
    Dim controlIdx As Long
    Dim sheetNameValues As Variant
    Dim sheetTypeValues As Variant
    Dim controlValues As Variant
    Dim orders As Variant
    Dim counter As Long
    Dim sheetName As String
    Dim varOrder As Long
    Dim prevControl As String

    AddMainSectionSort

    Set hostSheet = Wksh()
    firstRow = StartRow() + 1
    lastRow = DictEndRow()
    endCol = DictEndColumn()

    sheetNameIdx = ColumnIndex("sheet name")
    sheetTypeIdx = ColumnIndex("sheet type")
    controlIdx = ColumnIndex("control")

    hostSheet.Cells(StartRow(), endCol + 1).Value = "column index"
    ClearPreparedCache

    If lastRow < firstRow Then Exit Sub

    sheetNameValues = SheetColumnValues(sheetNameIdx, firstRow, lastRow)
    sheetTypeValues = SheetColumnValues(sheetTypeIdx, firstRow, lastRow)
    controlValues = SheetColumnValues(controlIdx, firstRow, lastRow)

    ReDim orders(1 To UBound(sheetNameValues))
    sheetName = vbNullString
    varOrder = 1

    For counter = 1 To UBound(sheetNameValues)
        If TextOf(sheetNameValues(counter)) <> sheetName Then
            varOrder = IIf(TextOf(sheetTypeValues(counter)) = SHEETTYPEVERTICAL, 4, 1)
            sheetName = TextOf(sheetNameValues(counter))
        Else
            'The row above decides the gap. On the first data row there is no
            'row above inside the block, so the default step applies.
            If counter > 1 Then
                prevControl = TextOf(controlValues(counter - 1))
            Else
                prevControl = vbNullString
            End If

            Select Case prevControl
                Case "geo"
                    varOrder = varOrder + 12
                Case "hf"
                    varOrder = varOrder + 2
                Case Else
                    varOrder = varOrder + 1
            End Select
        End If

        orders(counter) = varOrder
    Next

    WriteSheetColumn endCol + 1, firstRow, orders
End Sub

'@label:append-crf-line-index
'@sub-title Populate the CRF indices
'@details
'Form designers map dictionary variables to the rows shown in CRF-style layouts.
'The walk goes through sections, subsections and controls to insert spacing
'exactly like the interactive designer.
'
'Five columns are read as five blocks and the two answers are written as two
'blocks. This was the heaviest loop in the class: five range reads and up to two
'writes per row.
'
'A row that needs no choices marker keeps its cell Empty. Writing an empty
'string there would count as content for the last-row scan.
Private Sub AppendCRFLineIndex()
    Dim hostSheet As Worksheet
    Dim endCol As Long
    Dim firstRow As Long
    Dim subSectionValues As Variant
    Dim sectionValues As Variant
    Dim sheetValues As Variant
    Dim controlValues As Variant
    Dim detailsValues As Variant
    Dim crfIndexValues As Variant
    Dim choicesValues As Variant
    Dim crfIndexRange As Range
    Dim lineCounter As Long
    Dim currentSection As String
    Dim currentSubSection As String
    Dim currentSheet As String
    Dim currentControl As String
    Dim currentDetails As String
    Dim previousDetails As String
    Dim rowSection As String
    Dim rowSubSection As String
    Dim rowSheet As String
    Dim counter As Long
    Dim lastIndex As Long

    Set hostSheet = Wksh()
    endCol = DictEndColumn()
    firstRow = StartRow() + 1

    subSectionValues = ColumnValues(DataRange("sub section"))
    sectionValues = ColumnValues(DataRange("main section"))
    sheetValues = ColumnValues(DataRange("sheet name"))
    controlValues = ColumnValues(DataRange("control"))
    detailsValues = ColumnValues(DataRange("control details"))

    If Not (IsArray(subSectionValues) And IsArray(sectionValues) And IsArray(sheetValues) _
            And IsArray(controlValues) And IsArray(detailsValues)) Then
        LogInfo "CRF indices were not computed because a source column could not be read", checkingError
        Exit Sub
    End If

    hostSheet.Cells(StartRow(), endCol + 1).Value = "crf index"
    hostSheet.Cells(StartRow(), endCol + 2).Value = "crf choices"
    ClearPreparedCache

    Set crfIndexRange = DataRange("crf index")
    If crfIndexRange Is Nothing Then Exit Sub

    'All six blocks span the same rows. The smallest one still bounds the walk,
    'so a dictionary whose columns disagree gives a short answer rather than a
    'subscript error in the middle of preparation.
    lastIndex = crfIndexRange.Rows.Count
    If UBound(subSectionValues) < lastIndex Then lastIndex = UBound(subSectionValues)
    If UBound(sectionValues) < lastIndex Then lastIndex = UBound(sectionValues)
    If UBound(sheetValues) < lastIndex Then lastIndex = UBound(sheetValues)
    If UBound(controlValues) < lastIndex Then lastIndex = UBound(controlValues)
    If UBound(detailsValues) < lastIndex Then lastIndex = UBound(detailsValues)
    If lastIndex < 1 Then Exit Sub

    ReDim crfIndexValues(1 To lastIndex)
    ReDim choicesValues(1 To lastIndex)

    lineCounter = 0
    currentSection = vbNullString
    currentSubSection = vbNullString
    currentSheet = vbNullString

    For counter = 1 To UBound(crfIndexValues)
        previousDetails = currentDetails
        currentControl = TextOf(controlValues(counter))
        currentDetails = TextOf(detailsValues(counter))
        rowSheet = TextOf(sheetValues(counter))
        rowSection = TextOf(sectionValues(counter))
        rowSubSection = TextOf(subSectionValues(counter))

        If currentSheet <> rowSheet Then
            lineCounter = 6
            currentSheet = rowSheet
            currentSection = rowSection
            currentSubSection = rowSubSection
        ElseIf (currentSection <> rowSection) And (rowSection <> vbNullString) Then
            lineCounter = lineCounter + 7
            currentSection = rowSection
            currentSubSection = rowSubSection
        ElseIf (currentSubSection <> rowSubSection) And (rowSubSection <> vbNullString) Then
            lineCounter = lineCounter + 6
            currentSubSection = rowSubSection
        Else
            lineCounter = lineCounter + 2
        End If

        If (currentControl = "choice_manual") And (currentDetails <> previousDetails) Then
            lineCounter = lineCounter + 2
            choicesValues(counter) = "yes"
        End If

        crfIndexValues(counter) = lineCounter
    Next

    WriteSheetColumn endCol + 1, firstRow, crfIndexValues
    WriteSheetColumn endCol + 2, firstRow, choicesValues
End Sub


'@section Data Exchange
'===============================================================================

'@label:import
'@sub-title Import dictionary data
'@details
'Replaces the current dictionary content with data copied from another
'worksheet. The incoming data is wrapped in a temporary `DataSheet` so the
'existing cleaning and formatting routines can be reused.
'
'With `clearSheet` True the whole worksheet is cleared. Sheet-level hidden names
'survive a `Cells.Clear`, and the export counter is one of them, which is why
'the source counter is read first.
'@params
'- fromWksh Source worksheet
'- fromStartRow Long. Start row of the source data
'- fromStartcol Long. Start column of the source data
'- clearSheet Optional Boolean. Clear the whole host worksheet first.
Public Sub Import(ByVal fromWksh As Worksheet, _
                  ByVal fromStartRow As Long, _
                  ByVal fromStartcol As Long, _
                  Optional ByVal clearSheet As Boolean = False)

    Dim targetData As DataSheet
    Dim sourceData As DataSheet
    Dim hostSheet As Worksheet
    Dim importedTotal As Long

    Set sourceData = DataSheet.Create(fromWksh, fromStartRow, fromStartcol)
    Set hostSheet = Wksh()
    importedTotal = StoredTotalExportsFromSheet(fromWksh)

    If importedTotal > 0 Then
        ApplyTotalExports importedTotal, persist:=True, suppressInvalidLog:=True
    Else
        ApplyTotalExports TotalNumberOfExports, persist:=True, suppressInvalidLog:=True
    End If

    EnsureColumnsInitialized
    EnsureFormatColumns

    If clearSheet Then
        hostSheet.Cells.Clear
        this.allowedColumns.ToExcelRange hostSheet.Cells(StartRow(), StartColumn()), TransposeValues:=True
    End If

    Set targetData = Data()
    Clean removeAddedColumns:=True
    targetData.Import sourceData, False

    ClearPreparedCache
End Sub

'@label:export
'@sub-title Export dictionary data
'@details
'Copies the dictionary into the destination workbook with optional filtering,
'table creation, and sheet visibility adjustments to mirror the interactive
'export workflow.
'
'The prepared marker is written from the TARGET worksheet and it is written
'whatever `addListObject` says. The copy always lands on the first cell of the
'target, so taking the row from the source put the marker four rows below the
'data for every setup workbook, whose dictionary starts at row 5, and `Prepared`
'then answered False on the copy. `LLExporter.AddDictionary` writes the same
'marker itself, at the same cell, because it could not rely on this one.
'@params
'- toWkb Destination workbook
'- exportType Optional String. Filter by column name. Defaults to "__all__"
'- addListObject Optional Boolean. Create a listobject in the destination.
'  Defaults to True
'- Hide Optional Long. Visibility setting applied after export. Defaults to
'  xlSheetHidden
Public Sub Export(ByVal toWkb As Workbook, _
                  Optional ByVal exportType As String = "__all__", _
                  Optional ByVal addListObject As Boolean = True, _
                  Optional ByVal Hide As Long = xlSheetHidden)

    'Visual tweaks (fonts, colours) match the manual export workflow.
    Dim targetData As DataSheet
    Dim exportRange As Range
    Dim targetSheet As Worksheet
    Dim sheetName As String
    Dim tableName As String
    Dim markerRow As Long

    EnsureFormatColumns
    Set targetData = Data()

    If ColumnExists(exportType) Then
        targetData.Export toWkb, filteredVarName:=exportType, Hide:=Hide
    Else
        targetData.Export toWkb, Hide:=Hide
    End If

    sheetName = Wksh().Name
    Set targetSheet = toWkb.Worksheets(sheetName)

    PersistTotalExportsOn targetSheet, TotalNumberOfExports

    If addListObject Then
        Set exportRange = DataRange(includeHeaders:=True)

        'Export is always on the first cell of the worksheet
        With targetSheet
            Set exportRange = .Range(.Cells(1, 1), .Cells(exportRange.Rows.Count, exportRange.Columns.Count))
        End With

        tableName = "o" & LCase$(sheetName)
        targetSheet.ListObjects.Add(xlSrcRange, exportRange, , xlYes).Name = tableName
        targetSheet.ListObjects(tableName).TableStyle = ""
    End If

    markerRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
    targetSheet.Cells(markerRow, 1).Font.color = vbBlue
End Sub

'@label:translate
'@sub-title Translate dictionary labels
'@details
'Delegates the translation of user-facing columns (labels, notes, sections) to
'the provided translation object while protecting formula columns from direct
'text replacement.
'@params
'- TransObject TranslationObject responsible for translations
Public Sub Translate(ByVal TransObject As TranslationObject)
    'Control details may contain formulas, so the `containsFormula` flag is
    'passed for that column to avoid breaking references.
    Dim transCols As BetterArray
    Dim idx As Long
    Dim targetRange As Range

    Set transCols = New BetterArray
    transCols.Push "main label", "sub label", "dev comments", "sheet name", "note", _
                   "main section", "sub section", "message"

    For idx = transCols.LowerBound To transCols.UpperBound
        If ColumnExists(transCols.Item(idx)) Then
            Set targetRange = DataRange(transCols.Item(idx))
            TransObject.TranslateRange targetRange
        End If
    Next

    If ColumnExists("control details") Then
        Set targetRange = DataRange("control details")
        TransObject.TranslateRange targetRange, containsFormula:=True
    End If

    'Calculate the datarange (some computations can result into errors)
    On Error Resume Next
        DataRange.Calculate
    On Error GoTo 0
End Sub


'@section Checkings
'===============================================================================

'@label:log-info
'@sub-title Record one event in the shared checking object
'@details
'All noteworthy events (missing columns, normalisation steps, etc.) are piped
'into the shared checking object so setup and designer screens can surface
'diagnostics to the user. The host worksheet name is read only when the checking
'object is built; it used to cost two crossings on every logged line, and `Clean`
'can log once per deleted column.
Private Sub LogInfo(ByVal label As String, Optional ByVal scope As Byte = checkingNote)
    If this.internalChecks Is Nothing Then
        Set this.internalChecks = Checking.Create(Wksh().Name)
    End If

    this.internalChecks.Add CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@label:harvest-collaborator-checkings
'@sub-title Fold what a collaborator filed into this dictionary's entries
'@details
'`DataSheet` and `CustomTable` each keep a store of their own, and until
'this line nobody read either of them. An import that aborted, or a column
'name a table could not find, was filed and then dropped.
'@param collaboratorChecks Checking. What the collaborator filed.
Private Sub HarvestCollaboratorCheckings(ByVal collaboratorChecks As Checking)
    If collaboratorChecks Is Nothing Then Exit Sub
    If collaboratorChecks.Length() <= 0 Then Exit Sub

    If this.internalChecks Is Nothing Then
        Set this.internalChecks = Checking.Create(Wksh().Name)
    End If

    this.internalChecks.Append collaboratorChecks
    this.hasChecks = True
End Sub

'@label:merge-collaborator-checkings
'@sub-title Take the data sheet's entries before the report is read
'@details
'The data sheet lives as long as the dictionary does, so its entries are
'folded in when the report is asked for. The flag keeps the fold to once:
'the caller reads `HasCheckings` and then `CheckingValues`, and appending
'twice would double every line.
Private Sub MergeCollaboratorCheckings()
    If this.collaboratorsMerged Then Exit Sub
    this.collaboratorsMerged = True

    If this.dictDataSheet Is Nothing Then Exit Sub
    If Not this.dictDataSheet.HasCheckings Then Exit Sub

    HarvestCollaboratorCheckings this.dictDataSheet.CheckingValues
End Sub

'@label:has-checkings
'@prop-title True once something was logged
'@return Boolean
Public Property Get HasCheckings() As Boolean
    MergeCollaboratorCheckings
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title The logged events
'@return Checking
Public Property Get CheckingValues() As Checking
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = this.internalChecks
End Property


'@section Errors
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError and record it
'@details
'The message is logged BEFORE the raise. Logging afterwards meant the line never
'ran, so every raise from this class was missing from the generation report,
'which is the one place the user sees why a build stopped.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    LogInfo errNumber & " - " & errorMessage, checkingError
    Err.Raise errNumber, CLASSNAME, errorMessage
End Sub
