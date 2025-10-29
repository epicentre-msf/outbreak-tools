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


Implements ILLdictionary

'The linelist dictionary class centralises all logic that manipulates the Excel
'"Dictionary" worksheet used across setup and designer workflows. It merges
'behaviour from the legacy implementations so consumers interact with a single
'surface area and tests target a unified object model.

Private Type TDictionary
    dictDataSheet As IDataSheet
    numberOfExports As Long
    hasChecks As Boolean
    hiddenValues As IHiddenNames
End Type

Private this As TDictionary
Private dictionaryColumns As BetterArray
Private dictionaryAddedColumns As BetterArray

Private Const SHEETTYPEVERTICAL As String = "vlist1D"
Private Const SHEETTYPEHORIZONTAL As String = "hlist2D"
Private Const DEFAULTNUMBEROFEXPORTS As Long = 20
Private Const DEFAULTDICTIONARYCOLUMNS As String = "variable name|main label|dev comments|editable label|sub label|note|sheet name|sheet type|main section|sub section|status|register book|personal identifier|variable type|variable format|control|control details|unique|min|max|alert|message|formatting condition|formatting values|lock cells"
Private Const DEFAULTCREATEDCOLUMNS As String = "table name|column index|visibility|crf index|crf choices|crf status|list auto"
Private Const TOTAL_EXPORTS_NAME As String = "__ll_exports_total__"

'Checking support
Private checkCounter As Long
Private internalChecks As IChecking

'@section Instantiation
'===============================================================================

'@label:create
'@prop-title Create a dictionary object
'@details
'Creates a linelist dictionary wrapper around the supplied worksheet. The
'routine initialises a backing `DataSheet` so all subsequent access goes through
'a consistent abstraction and callers interact with a predeclared instance.
'@params
'- dictWksh Worksheet hosting the dictionary data
'- dictStartRow Long. Header row of the dictionary
'- dictStartColumn Long. First column of the dictionary
'- numberOfExports Optional Long. Expected number of export columns
'@return An ILLdictionary instance
'@depends DataSheet, Checking
'@export
Public Function Create(ByVal dictWksh As Worksheet, ByVal dictStartRow As Long, _
                       ByVal dictStartColumn As Long, _
                       Optional ByVal numberOfExports As Long = DEFAULTNUMBEROFEXPORTS) As ILLdictionary

    Dim customDataSheet As IDataSheet

    Set customDataSheet = DataSheet.Create(dictWksh, dictStartRow, dictStartColumn)

    With New LLdictionary
        Set .Data = customDataSheet
        .InitialiseTotalExports numberOfExports
        Set Create = .Self
    End With
End Function

'@label:init-exports
'@sub-title Initialise total export counter
'@details
'Loads the stored number of export columns from the worksheet when available.
'When no persisted value exists, the routine falls back to the requested input
'and persists it so subsequent sessions pick up the same configuration.
Public Sub InitialiseTotalExports(ByVal requestedTotal As Long)
    Dim storedTotal As Long

    storedTotal = StoredTotalExportsFromSheet(Wksh())

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

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor so consuming code can fluently retrieve the instance
'returned by `Create` without repeatedly typing `Set x = New LLdictionary`.
'@return ILLdictionary
Public Property Get Self() As ILLdictionary
    Set Self = Me
End Property


'@section Data Elements
'===============================================================================

'@label:data
'@prop-title DataSheet backing store
'@details
'Expose the lazily initialised `DataSheet` wrapper used for all worksheet
'interactions. Consumers rely on this to perform advanced queries without
'directly manipulating Excel ranges.
'@return IDataSheet
Public Property Get Data() As IDataSheet
    Set Data = this.dictDataSheet
End Property

Public Property Set Data(ByVal customDataSheet As IDataSheet)
    Set this.dictDataSheet = customDataSheet
    Set this.hiddenValues = Nothing
End Property

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
    Dim previousTotal As Long

    previousTotal = this.numberOfExports
    normalisedTotal = NormaliseTotalExports(requestedTotal, suppressInvalidLog)

    this.numberOfExports = normalisedTotal

    If normalisedTotal <> previousTotal Then ResetColumnCaches

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

'@label:reset-export-cache
Private Sub ResetColumnCaches()
    Set dictionaryColumns = Nothing
    Set dictionaryAddedColumns = Nothing
End Sub

'@label:persist-exports
Private Sub PersistTotalExports(ByVal totalCount As Long)
    Dim store As IHiddenNames

    Set store = HiddenNameStore()
    If store Is Nothing Then Exit Sub

    If totalCount < 0 Then totalCount = 0

    store.EnsureName TOTAL_EXPORTS_NAME, totalCount, HiddenNameTypeLong
    store.SetValue TOTAL_EXPORTS_NAME, totalCount
End Sub

'@label:persist-exports-sheet
Private Sub PersistTotalExportsOnSheet(ByVal sheet As Worksheet, ByVal totalCount As Long)
    Dim store As IHiddenNames

    If sheet Is Nothing Then Exit Sub
    If totalCount < 0 Then totalCount = 0

    Set store = HiddenNameStore(sheet)
    If store Is Nothing Then Exit Sub

    store.EnsureName TOTAL_EXPORTS_NAME, totalCount, HiddenNameTypeLong
    store.SetValue TOTAL_EXPORTS_NAME, totalCount
End Sub

'@label:stored-exports
Private Function StoredTotalExportsFromSheet(ByVal sheet As Worksheet) As Long
    Dim store As IHiddenNames

    If sheet Is Nothing Then Exit Function

    Set store = HiddenNameStore(sheet)
    If store Is Nothing Then Exit Function

    StoredTotalExportsFromSheet = store.ValueAsLong(TOTAL_EXPORTS_NAME, 0)
End Function

'@label:hidden-name-store
Private Function HiddenNameStore(Optional ByVal sheet As Worksheet) As IHiddenNames
    Dim target As Worksheet

    If sheet Is Nothing Then
        If this.dictDataSheet Is Nothing Then Exit Function
        Set target = Wksh()
        If target Is Nothing Then Exit Function

        If this.hiddenValues Is Nothing Then
            Set this.hiddenValues = HiddenNames.Create(target)
        ElseIf Not this.hiddenValues.Sheet Is target Then
            Set this.hiddenValues = HiddenNames.Create(target)
        End If

        Set HiddenNameStore = this.hiddenValues
    Else
        Set HiddenNameStore = HiddenNames.Create(sheet)
    End If
End Function

'@label:try-sheet-name
Private Function TrySheetNameDefinition(ByVal sheet As Worksheet, ByVal nameId As String) As Name
    Dim definition As Name

    If sheet Is Nothing Then Exit Function

    For Each definition In sheet.Names
        If StrComp(ExtractSimpleName(definition.Name), nameId, vbTextCompare) = 0 Then
            Set TrySheetNameDefinition = definition
            Exit Function
        End If
    Next definition
End Function

'@label:extract-name
Private Function ExtractSimpleName(ByVal qualifiedName As String) As String
    Dim exclPos As Long

    exclPos = InStr(qualifiedName, "!")
    If exclPos = 0 Then
        ExtractSimpleName = qualifiedName
    Else
        ExtractSimpleName = Mid$(qualifiedName, exclPos + 1)
    End If
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
'@return Long
Private Property Get DictEndRow() As Long
    DictEndRow = Data().DataEndRow
End Property

'@label:endcol
'@prop-title Last data column
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


'@section Column Configuration
'===============================================================================

Private Sub EnsureColumnsInitialized()
    'Initialise the cached column lists on demand. Using `BetterArray`
    'structures keeps lookups platform-agnostic while providing helper
    'functions such as `Includes`.
    Dim counter As Long

    If Not dictionaryColumns Is Nothing Then Exit Sub

    Set dictionaryColumns = New BetterArray
    Set dictionaryAddedColumns = New BetterArray

    dictionaryColumns.Items = Split(DEFAULTDICTIONARYCOLUMNS, "|")
                           

    For counter = 1 To TotalNumberOfExports()
        dictionaryColumns.Push "export " & CStr(counter)
    Next

    'Columns introduced during preparation (e.g. table name, CRF helpers) are
    'tracked separately so we can optionally strip them during cleaning.
    dictionaryAddedColumns.Items = Split(DEFAULTCREATEDCOLUMNS, "|") 

    dictionaryColumns.Concat dictionaryAddedColumns.Items
End Sub

Private Sub EnsureFormatColumns()
    this.dictDataSheet.AddFormatsColumns False, True, "formatting condition", "formatting values", "variable name", "control", "main label", "lock cells", "dev comments"
End Sub

'@label:columnexists
'@prop-title Test if a column exists
'@details
'Checks whether a column header is present in the dictionary and optionally
'verifies it belongs to the known schema. Recording the attempt helps capture
'incomplete dictionaries during setup.
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
        exists = dictionaryColumns.Includes(colName) Or (Split(colName, " ")(0) = "export")
        If Not exists Then LogInfo "Column " & colName & " exists but is not part of the allowed dictionary columns", checkingWarning
    ElseIf (Not exists) And checkValidity Then
        LogInfo "Column " & colName & " is missing from dictionary", checkingWarning
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
'Collects a distinct list of values for the requested column, ignoring
'case-sensitive duplicates to reflect how dictionary metadata is typically
'interpreted.
'@params
'- colName Column header
'@return BetterArray
Public Function UniqueValues(ByVal colName As String) As BetterArray
    'Collect a distinct list of entries for the supplied column. Downstream
    'components use this information to populate combo boxes or validate user
    'input against known dictionary metadata.

    Dim values As BetterArray
    Dim counter As Long
    Dim lastRow As Long
    Dim nextValue As String
    Dim dictSheet As Worksheet
    Dim colIdx As Long

    Set values = New BetterArray

    If Not ColumnExists(colName) Then
        LogInfo "Unable to compute unique values. Column " & colName & " not found", checkingWarning
        Set UniqueValues = values.Clone()
        Exit Function
    End If

    lastRow = DictEndRow()
    colIdx = ColumnIndex(colName)
    Set dictSheet = Wksh()

    For counter = StartRow() + 1 To lastRow
        nextValue = dictSheet.Cells(counter, colIdx).Value
        If Not values.Includes(nextValue) Then values.Push nextValue
    Next

    Set UniqueValues = values.Clone()
End Function

'@label:variableexists
'@prop-title Variable presence test
'@details
'Tests the dictionary for a specific variable identifier by scanning the
'"variable name" column. Used to guard operations that may attempt to insert
'duplicate entries.
'@params
'- varName Variable identifier
'@return Boolean
Public Function VariableExists(ByVal varName As String) As Boolean
    Dim varRange As Range
    Set varRange = DataRange("variable name")

    If varRange Is Nothing Then
        LogInfo "Unable to locate variable name column while checking for variable " & varName, checkingError
        VariableExists = False
    Else
        VariableExists = Not (varRange.Find(What:=varName, lookAt:=xlWhole, MatchCase:=True) Is Nothing)
    End If
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

    On Error Resume Next
        Set headerCell = hostSheet.Range(anchorRange.Cells(-1, 1), anchorRange.Cells(0, 2))
        headerCell.Merge
    On Error GoTo 0
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
    On Error Resume Next
        columnRange.EntireColumn.Delete
    On Error GoTo 0
End Sub


'@section Row Operations
'===============================================================================

Private Sub ManageRows(Optional ByVal del As Boolean = False)
    'Wrapper around the worksheet listobject to add or remove rows while
    'preserving formatting. Not every dictionary is table-backed; if no
    'ListObject is found we simply log the attempt for traceability.
    Dim hostSheet As Worksheet
    Dim listObj As ListObject
    Dim table As ICustomTable

    Set hostSheet = Wksh()

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
End Sub

'@label:addrows
'@sub-title Append rows via listobject
'@details
'Adds rows to the listobject backing the dictionary while preserving table
'formatting. When no listobject is available the action is logged for
'traceability.
Public Sub AddRows()
    ManageRows
End Sub

'@label:removerows
'@sub-title Remove rows via listobject
'@details
'Removes rows from the dictionary listobject using the shared `ManageRows`
'helper. Logs a note when a listobject is not present.
Public Sub RemoveRows()
    ManageRows del:=True
End Sub


'@section Specialised Views
'===============================================================================

Private Function ResolveSpecialVars(ByVal firstCondition As String, _
                                    Optional ByVal secondCondition As String = vbNullString, _
                                    Optional ByVal conditionName As String = "Control") As BetterArray
    'Shared filtering routine behind `ChoicesVars`, `GeoVars`, and
    '`TimeVars`. The function inspects the control column (or a custom column)
    'and returns the variables that match the supplied criteria. Only variables
    'living on horizontal sheets (`hlist2D`) are relevant for these helpers.
    Dim result As BetterArray
    Dim controlRange As Range
    Dim variableRange As Range
    Dim sheetRange As Range
    Dim lastRow As Long
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

    Set controlRange = DataRange(conditionName)
    Set variableRange = DataRange("variable name")
    Set sheetRange = DataRange("sheet type")
    lastRow = DictEndRow()

    For idx = 1 To lastRow - StartRow()
        controlValue = controlRange.Cells(idx, 1).Value
        sheetValue = sheetRange.Cells(idx, 1).Value

        If sheetValue = SHEETTYPEHORIZONTAL Then
            If (controlValue = firstCondition) Or ((controlValue = secondCondition) And (secondCondition <> vbNullString)) Then
                result.Push variableRange.Cells(idx, 1).Value
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

'@label:prepared
'@prop-title Dictionary preparation flag
'@details
'Indicates whether all helper columns have been generated for designer
'workflows. The heuristic mirrors the legacy implementation by inspecting
'specific markers instead of tracking explicit state.
'@return Boolean
Public Property Get Prepared() As Boolean
    'Quick heuristic used throughout the designer flows to check whether the
    'dictionary already contains all generated helper columns. Preparation is a
    'one-time operation, so we compare against the known set of structural
    'markers instead of storing additional state.
    Dim hostSheet As Worksheet

    Set hostSheet = Wksh()
    Prepared = (hostSheet.Cells(DictEndRow() + 1, StartColumn()).Font.color = vbBlue) And _
               ColumnExists("table name") And ColumnExists("column index") And _
               ColumnExists("visibility") And ColumnExists("crf index") And _
               ColumnExists("crf choices") And ColumnExists("crf status") And _
               (Not ColumnExists("main section sort")) And (Not ColumnExists("main section index"))
End Property

'@label:prepare
'@sub-title Prepare dictionary for designer workflows
'@details
'Extends the dictionary with all helper metadata required by the designer:
'table names, CRF indices, geo duplications, and visibility columns. The method
'guards against repeated execution by checking the `Prepared` flag and logs
'issues discovered during enrichment.
'@params
'- PreservedSheetNames BetterArray of sheet names that should keep their
'  original values
'- geoObject LLGeo instance required for geo expansions
'- msgTrads ITranslationObject used during column creation
'- tablePrefix Optional String prefix applied to generated table names
Public Sub Prepare(Optional ByVal PreservedSheetNames As BetterArray, _
                   Optional ByVal geoObject As Object, _
                   Optional ByVal tablePrefix As String = "table")

    'Augment the user dictionary with all helper columns required by the
    'designer (table names, CRF metadata, geo expansions, etc.). The order of
    'operations mirrors the legacy implementations to keep exported workbooks
    'compatible with existing automation.
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
            If Not controlRange.Find(What:="geo", lookAt:=xlWhole, MatchCase:=True) Is Nothing Then
                If ColumnExists("control") Then AppendGeoLines geoObject
            End If
        End If
    End If

    UniqueVarNames
    GoodSheetNames PreservedSheetNames

    If Not ColumnExists("visibility") Then AddVisibility
    If Not ColumnExists("crf index") Then AppendCRFLineIndex

    Clean removeAddedColumns:=False

    With Wksh()
        .Cells(DictEndRow() + 1, StartColumn()).Font.color = vbBlue
    End With
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
            ThrowError ProjectError.ElementNotFound, "Column -" & colName &  "- not found in the dictionary"
        End If
    Next
End Sub

'@label:clean
'@sub-title Remove unexpected columns
'@details
'Deletes columns that fall outside the supported dictionary schema. When the
'optional argument is True, helper columns introduced during preparation are
'removed as well so the dictionary reverts to the user-provided structure.
'@params
'- removeAddedColumns Optional Boolean. Remove preparation columns too.
Public Sub Clean(Optional ByVal removeAddedColumns As Boolean = False)
    'Strip unexpected columns from the dictionary. When `removeAddedColumns` is
    'True we also delete helper columns that were previously generated during
    'preparation so the dictionary reverts to the original user-supplied shape.
    Dim hostSheet As Worksheet
    Dim usedCols As Long
    Dim colIndex As Long
    Dim columnName As String
    Dim removeAdded As Boolean

    Set hostSheet = Wksh()
    EnsureColumnsInitialized

    usedCols = hostSheet.UsedRange.Columns.Count + hostSheet.UsedRange.Column - 1

    For colIndex = usedCols To 1 Step -1
        columnName = hostSheet.Cells(StartRow(), colIndex).Value
        removeAdded = removeAddedColumns And dictionaryAddedColumns.Includes(LCase$(columnName))
        If (Not ColumnExists(columnName, True)) Or removeAdded Then
            hostSheet.Columns(colIndex).EntireColumn.Delete
        End If
    Next
End Sub

Private Sub GoodSheetNames(ByVal PreservedSheetNames As BetterArray)
    Dim hostSheet As Worksheet
    Dim sheetNameIndex As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim currValue As String

    If PreservedSheetNames Is Nothing Then Exit Sub

    Set hostSheet = Wksh()
    sheetNameIndex = ColumnIndex("sheet name")
    lastRow = DictEndRow()

    For rowIndex = StartRow() + 1 To lastRow
        currValue = hostSheet.Cells(rowIndex, sheetNameIndex).Value
        If PreservedSheetNames.Includes(currValue) Then
            hostSheet.Cells(rowIndex, sheetNameIndex).Value = currValue & "_"
            LogInfo "Sheet name " & currValue & " duplicated and suffixed with underscore during preparation", checkingNote
        End If
    Next
End Sub

Private Sub UniqueVarNames()
    'Normalise variable names so downstream code can rely on predictable
    'identifiers. The rules mimic the designer implementation: trim whitespace,
    'replace spaces with underscores, and force uniqueness by appending an
    'underscore when clashes occur.
    Dim hostSheet As Worksheet
    Dim seen As BetterArray
    Dim lastRow As Long
    Dim varIndex As Long
    Dim rowIndex As Long
    Dim varName As String

    EnsureColumnsInitialized

    Set hostSheet = Wksh()
    Set seen = New BetterArray
    lastRow = DictEndRow()
    varIndex = ColumnIndex("variable name")

    For rowIndex = StartRow() + 1 To lastRow
        varName = hostSheet.Cells(rowIndex, varIndex).Value
        varName = Application.WorksheetFunction.Substitute(varName, Chr$(160), " ")
        varName = Application.WorksheetFunction.Clean(varName)
        varName = Application.WorksheetFunction.Trim(varName)
        varName = Replace(varName, " ", "_")

        If Len(varName) < 4 Then
            ThrowError ProjectError.InvalidArgument, "Variable names should be at least 4 characters. Please review variable at row " & rowIndex
        End If

        If seen.Includes(varName) Then varName = varName & "_"
        seen.Push varName
        hostSheet.Cells(rowIndex, varIndex).Value = varName
    Next
End Sub

Private Sub AddVisibility()
    'Derive the visibility and CRF status columns based on the current status
    'and control definitions. The logic mirrors the designer wizard so that
    'prepared dictionaries behave the same whether produced manually or via
    'automation.
    Dim hostSheet As Worksheet
    Dim lastRow As Long
    Dim statusIdx As Long
    Dim visibilityIdx As Long
    Dim crfStatusIdx As Long
    Dim controlIdx As Long
    Dim rowIndex As Long
    Dim statusValue As String
    Dim controlValue As String

    AddColumn "visibility"
    AddColumn "crf status"

    Set hostSheet = Wksh()
    lastRow = DictEndRow
    statusIdx = ColumnIndex("status")
    visibilityIdx = ColumnIndex("visibility")
    crfStatusIdx = ColumnIndex("crf status")
    controlIdx = ColumnIndex("control")

    For rowIndex = StartRow() + 1 To lastRow
        statusValue = hostSheet.Cells(rowIndex, statusIdx).Value
        Select Case statusValue
            Case "hidden"
                hostSheet.Cells(rowIndex, visibilityIdx).Value = "always hidden"
                hostSheet.Cells(rowIndex, crfStatusIdx).Value = "always hidden"
            Case "mandatory"
                hostSheet.Cells(rowIndex, visibilityIdx).Value = "always visible"
                controlValue = hostSheet.Cells(rowIndex, controlIdx).Value
                If controlValue = "formula" Or controlValue = "choice_formula" Or controlValue = "case_when" Then
                    hostSheet.Cells(rowIndex, crfStatusIdx).Value = "always hidden"
                Else
                    hostSheet.Cells(rowIndex, crfStatusIdx).Value = "visible on creation"
                End If
            Case "optional, hidden"
                hostSheet.Cells(rowIndex, visibilityIdx).Value = "hidden on creation"
                hostSheet.Cells(rowIndex, crfStatusIdx).Value = "hidden on creation"
            Case "optional, visible", "optional"
                hostSheet.Cells(rowIndex, visibilityIdx).Value = "visible on creation"
                hostSheet.Cells(rowIndex, crfStatusIdx).Value = "visible on creation"
            Case Else
                LogInfo "Unhandled status value " & statusValue & " at row " & rowIndex, checkingNote
        End Select
    Next
End Sub

'Note: AppendGeoLines, AppendNumberColumn, AppendColumnIndex, AppendCRFLineIndex
'retain much of the original designer implementation with minor guard rails.

Private Sub AppendGeoLines(ByVal geoObject As Object)
    'Insert the extra hierarchy required for geographic controls. For each
    '"geo" variable we duplicate the base row to create admin levels, pcodes,
    'and concatenated helper fields. The implementation intentionally mirrors
    'the original designer code so exported templates remain backward
    'compatible.
    Dim dictRange As Range
    Dim geoRow As Long
    Dim hostSheet As Worksheet
    Dim endRow As Long
    Dim endCol As Long
    Dim baseRowRange As Range
    Dim pastedRange As Range
    Dim varName As String
    Dim varPosition As Long
    Dim counter As Long
    Dim controlIdx As Long
    Dim varIdx As Long
    Dim mainLabelIdx As Long
    Dim subSectionIdx As Long
    Dim registerIdx As Long
    Dim statusIdx As Long
    Dim positionIdx As Long
    Dim controlDetailsIdx As Long
    Dim formulaValue As String

    Set hostSheet = Wksh()
    Set dictRange = DataRange("control")
    controlIdx = ColumnIndex("control")
    varIdx = ColumnIndex("variable name")
    mainLabelIdx = ColumnIndex("main label")
    subSectionIdx = ColumnIndex("sub section")
    registerIdx = ColumnIndex("register book")
    statusIdx = ColumnIndex("status")
    positionIdx = ColumnIndex("column index")
    controlDetailsIdx = ColumnIndex("control details")
    endRow = DictEndRow()
    endCol = DictEndColumn()
    geoRow = StartRow() + 1

    Do While Not dictRange.Find(What:="geo", lookAt:=xlWhole, MatchCase:=True) Is Nothing
        geoRow = dictRange.Find(What:="geo", lookAt:=xlWhole, MatchCase:=True).Row
        Set baseRowRange = hostSheet.Range(hostSheet.Cells(geoRow, StartColumn()), hostSheet.Cells(geoRow, endCol))
        varName = hostSheet.Cells(geoRow, varIdx).Value
        varPosition = hostSheet.Cells(geoRow, positionIdx).Value

        If IsEmpty(hostSheet.Cells(geoRow, subSectionIdx)) And (Not IsEmpty(hostSheet.Cells(geoRow, mainLabelIdx))) Then
            hostSheet.Cells(geoRow, subSectionIdx).Value = hostSheet.Cells(geoRow, mainLabelIdx).Value
            hostSheet.Cells(geoRow, mainLabelIdx).Value = vbNullString
        End If

        hostSheet.Cells(geoRow, varIdx).Value = "adm1_" & varName
        hostSheet.Cells(geoRow, mainLabelIdx).Value = geoObject.GeoNames("adm1_name")
        hostSheet.Cells(geoRow, controlIdx).Value = "geo1"

        geoRow = geoRow + 1
        For counter = 12 To 2 Step -1
            hostSheet.Rows(geoRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Set pastedRange = hostSheet.Range(hostSheet.Cells(geoRow, StartColumn()), hostSheet.Cells(geoRow, endCol))
            pastedRange.Value = baseRowRange.Value
            hostSheet.Cells(geoRow, varIdx).Value = "adm" & counter & "_" & varName
            hostSheet.Cells(geoRow, controlIdx).Value = "geo" & counter
            hostSheet.Cells(geoRow, mainLabelIdx).Value = geoObject.GeoNames("adm" & counter & "_name")
            hostSheet.Cells(geoRow, positionIdx).Value = varPosition + counter - 1
            geoRow = geoRow + 1
        Next

        For counter = 1 To 4
            hostSheet.Cells(geoRow + counter, statusIdx).Value = "hidden"
            hostSheet.Cells(geoRow + counter, registerIdx).Value = "hidden"
            hostSheet.Cells(geoRow + counter, controlIdx).Value = "formula"
            hostSheet.Cells(geoRow + counter, mainLabelIdx).Value = "geo_pcode_adm" & counter
            hostSheet.Cells(geoRow + counter, varIdx).Value = "pcode_adm" & counter & "_" & varName
            formulaValue = "INDEX(T_ADM" & counter & ", MATCH(concat_adm" & counter & "_" & varName & ", adm" & counter & "_concat, 0), " & (counter + 2) & ")"
            hostSheet.Cells(geoRow + counter, controlDetailsIdx) = "IF(ISNA(" & formulaValue & "), """", " & formulaValue & ")"
        Next

        geoRow = geoRow + 5
        For counter = 1 To 4
            hostSheet.Cells(geoRow + counter, statusIdx).Value = "hidden"
            hostSheet.Cells(geoRow + counter, registerIdx).Value = "hidden"
            hostSheet.Cells(geoRow + counter, controlIdx).Value = "formula"
            hostSheet.Cells(geoRow + counter, mainLabelIdx).Value = "geo_concat_adm" & counter
            hostSheet.Cells(geoRow + counter, varIdx).Value = "concat_adm" & counter & "_" & varName
            hostSheet.Cells(geoRow + counter, controlDetailsIdx) = "GEOCONCAT(adm1_" & varName & ", " & counter & ")"
        Next

        geoRow = geoRow + 5
        Set dictRange = hostSheet.Range(hostSheet.Cells(geoRow, controlIdx), hostSheet.Cells(endRow, controlIdx))
    Loop
End Sub

Private Sub AppendNumberColumn(Optional ByVal onColName As String = "sheet name", _
                               Optional ByVal tablePrefix As String = "table", _
                               Optional ByVal newColumn As String = "table name")
    'Assign an incrementing number to each distinct sheet (or user-defined
    'group) so the dictionary can later produce stable table names.
    Dim lastRow As Long
    Dim endCol As Long
    Dim tableOrder As Long
    Dim sheetNames As BetterArray
    Dim sheetTrack As BetterArray
    Dim hostSheet As Worksheet
    Dim rowIndex As Long
    Dim sheetName As String
    Dim sortRange As Range

    Set hostSheet = Wksh()
    lastRow = DictEndRow()
    endCol = DictEndColumn()
    tableOrder = 1
    Set sheetNames = New BetterArray
    Set sheetTrack = New BetterArray

    hostSheet.Cells(StartRow(), endCol + 1).Value = "number_" & newColumn
    sheetTrack.Push hostSheet.Cells(StartRow() + 1, ColumnIndex(onColName)).Value
    sheetNames.Push tableOrder

    For rowIndex = StartRow() + 1 To lastRow
        sheetName = hostSheet.Cells(rowIndex, ColumnIndex(onColName)).Value
        If sheetTrack.Includes(sheetName) Then
            hostSheet.Cells(rowIndex, endCol + 1).Value = sheetNames.Item(sheetTrack.IndexOf(sheetName))
        Else
            tableOrder = tableOrder + 1
            sheetTrack.Push sheetName
            sheetNames.Push tableOrder
            hostSheet.Cells(rowIndex, endCol + 1).Value = tableOrder
        End If
    Next

    Set sortRange = DataRange
    sortRange.Sort key1:=DataRange("number_" & newColumn), Order1:=xlAscending

    '@Ignore AssignmentNotUsed
    endCol = DictEndColumn()
    hostSheet.Cells(StartRow(), endCol + 1).Value = newColumn
    hostSheet.Cells(StartRow() + 1, endCol + 1).Formula = "= " & Chr$(34) & tablePrefix & Chr$(34) & " " & Chr$(38) & " " & hostSheet.Cells(StartRow() + 1, endCol).Address(Rowabsolute:=False)
    Set sortRange = DataRange(newColumn)
    hostSheet.Cells(StartRow() + 1, endCol + 1).AutoFill sortRange, Type:=xlFillValues
    hostSheet.Calculate
    sortRange.Value = sortRange.Value
End Sub

Private Sub AppendColumnIndex()
    'Compute the final column order used by the designer when producing
    'workbooks. The routine ensures sheet-level grouping is respected while
    'inserting gaps for geo/hf variables that expand into multiple columns.
    Dim hostSheet As Worksheet
    Dim endCol As Long
    Dim lastRow As Long
    Dim sheetTypeIdx As Long
    Dim sheetNameIdx As Long
    Dim mainSectionIdx As Long
    Dim controlIdx As Long
    Dim sheetName As String
    Dim varOrder As Long
    Dim prevControl As String
    Dim rowIndex As Long
    Dim sortRange As Range
    Dim firstSheetCell As String
    Dim firstSectionCell As String

    Set hostSheet = Wksh()
    endCol = DictEndColumn()
    lastRow = DictEndRow()
    sheetTypeIdx = ColumnIndex("sheet type")
    sheetNameIdx = ColumnIndex("sheet name")
    mainSectionIdx = ColumnIndex("main section")
    controlIdx = ColumnIndex("control")

    hostSheet.Cells(StartRow(), endCol + 1).Value = "main section sort"
    firstSheetCell = hostSheet.Cells(StartRow() + 1, sheetNameIdx).Address(Rowabsolute:=False) & " "
    firstSectionCell = hostSheet.Cells(StartRow() + 1, mainSectionIdx).Address(Rowabsolute:=False)
    hostSheet.Cells(StartRow() + 1, endCol + 1).Formula = "= IF(" & firstSectionCell & "=" & Chr$(34) & Chr$(34) & ", ROW(" & firstSectionCell & "), " & firstSheetCell & Chr$(38) & Chr$(34) & " - " & Chr$(34) & Chr$(38) & firstSectionCell & ")"
    Set sortRange = DataRange("main section sort")
    hostSheet.Cells(StartRow() + 1, endCol + 1).AutoFill sortRange, Type:=xlFillValues

    AppendNumberColumn "main section sort", "main section", "main section index"

    endCol = DictEndColumn()
    hostSheet.Cells(StartRow(), endCol + 1).Value = "column index"
    sheetName = vbNullString
    varOrder = 1

    For rowIndex = StartRow() + 1 To lastRow
        If hostSheet.Cells(rowIndex, sheetNameIdx).Value <> sheetName Then
            varOrder = IIf(hostSheet.Cells(rowIndex, sheetTypeIdx).Value = SHEETTYPEVERTICAL, 4, 1)
            sheetName = hostSheet.Cells(rowIndex, sheetNameIdx).Value
        Else
            prevControl = hostSheet.Cells(rowIndex - 1, controlIdx).Value
            Select Case prevControl
                Case "geo"
                    varOrder = varOrder + 12
                Case "hf"
                    varOrder = varOrder + 2
                Case Else
                    varOrder = varOrder + 1
            End Select
        End If
        hostSheet.Cells(rowIndex, endCol + 1).Value = varOrder
    Next
End Sub

Private Sub AppendCRFLineIndex()
    'Populate CRF indices so form designers can map dictionary variables to the
    'rows displayed in CRF-style layouts. The logic walks through sections,
    'subsections, and controls to insert spacing exactly like the interactive
    'designer.
    Dim hostSheet As Worksheet
    Dim endCol As Long
    Dim subSectionRange As Range
    Dim sectionRange As Range
    Dim sheetRange As Range
    Dim controlRange As Range
    Dim detailsRange As Range
    Dim crfIndexRange As Range
    Dim choicesRange As Range
    Dim lineCounter As Long
    Dim currentSection As String
    Dim currentSubSection As String
    Dim currentSheet As String
    Dim currentControl As String
    Dim currentDetails As String
    Dim previousDetails As String
    Dim rowIndex As Long

    Set hostSheet = Wksh()
    endCol = DictEndColumn()
    Set subSectionRange = DataRange("sub section")
    Set sectionRange = DataRange("main section")
    Set sheetRange = DataRange("sheet name")
    Set controlRange = DataRange("control")
    Set detailsRange = DataRange("control details")

    hostSheet.Cells(StartRow(), endCol + 1).Value = "crf index"
    hostSheet.Cells(StartRow(), endCol + 2).Value = "crf choices"

    Set crfIndexRange = DataRange("crf index")
    Set choicesRange = DataRange("crf choices")

    lineCounter = 0
    currentSection = vbNullString
    currentSubSection = vbNullString
    currentSheet = vbNullString

    For rowIndex = 1 To crfIndexRange.Rows.Count
        previousDetails = currentDetails
        currentControl = controlRange.Cells(rowIndex, 1).Value
        currentDetails = detailsRange.Cells(rowIndex, 1).Value

        If currentSheet <> sheetRange.Cells(rowIndex, 1).Value Then
            lineCounter = 6
            currentSheet = sheetRange.Cells(rowIndex, 1).Value
            currentSection = sectionRange.Cells(rowIndex, 1).Value
            currentSubSection = subSectionRange.Cells(rowIndex, 1).Value
        ElseIf (currentSection <> sectionRange.Cells(rowIndex, 1).Value) And (sectionRange.Cells(rowIndex, 1).Value <> vbNullString) Then
            lineCounter = lineCounter + 7
            currentSection = sectionRange.Cells(rowIndex, 1).Value
            currentSubSection = subSectionRange.Cells(rowIndex, 1).Value
        ElseIf (currentSubSection <> subSectionRange.Cells(rowIndex, 1).Value) And (subSectionRange.Cells(rowIndex, 1).Value <> vbNullString) Then
            lineCounter = lineCounter + 6
            currentSubSection = subSectionRange.Cells(rowIndex, 1).Value
        Else
            lineCounter = lineCounter + 2
        End If

        If (currentControl = "choice_manual") And (currentDetails <> previousDetails) Then
            lineCounter = lineCounter + 2
            choicesRange.Cells(rowIndex, 1).Value = "yes"
        End If

        crfIndexRange.Cells(rowIndex, 1).Value = lineCounter
    Next
End Sub


'@section Data Exchange
'===============================================================================

'@label:import
'@sub-title Import dictionary data
'@details
'Replaces the current dictionary content with data copied from another
'worksheet. The incoming data is wrapped in a temporary `DataSheet` so the
'existing cleaning and formatting routines can be reused.
'@params
'- fromWksh Source worksheet
'- fromStartRow Long. Start row of the source data
'- fromStartcol Long. Start column of the source data
Public Sub Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    'Replace the current dictionary content with data imported from another
    'worksheet. Helper columns are initialised before the import so the
    'datasheet can drop unexpected headers and reapply formatting.
    Dim targetData As IDataSheet
    Dim sourceData As IDataSheet
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

    hostSheet.Cells.Clear
    EnsureColumnsInitialized
    dictionaryColumns.ToExcelRange hostSheet.Cells(StartRow(), StartColumn()), TransposeValues:=True

    EnsureFormatColumns
    Set targetData = Data()
    Clean removeAddedColumns:=True
    targetData.Import sourceData, False
End Sub

'@label:export
'@sub-title Export dictionary data
'@details
'Copies the dictionary into the destination workbook with optional filtering,
'table creation, and sheet visibility adjustments to mirror the interactive
'export workflow.
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
    'Copy the dictionary to another workbook, optionally filtering by export
    'type and adding a listobject for backwards compatibility with legacy
    'macros. Visual tweaks (fonts, colours) match the manual export workflow.
    Dim targetData As IDataSheet
    Dim exportRange As Range
    Dim targetSheet As Worksheet
    Dim sheetName As String

    EnsureFormatColumns
    Set targetData = Data()
    

    If ColumnExists(exportType) Then
        targetData.Export toWkb, filteredVarName:=exportType, Hide:=Hide
    Else
        targetData.Export toWkb, Hide:=Hide
    End If

    sheetName = Wksh().Name
    Set targetSheet = toWkb.Worksheets(sheetName)

    PersistTotalExportsOnSheet targetSheet, TotalNumberOfExports

    If addListObject Then
        Set exportRange = DataRange(includeHeaders:=True)

        'Export is always on the first cell of the worksheet
        With targetSheet
            Set exportRange = .Range(.Cells(1, 1), .Cells(exportRange.Rows.Count, exportRange.Columns.Count))
        End With

        targetSheet.ListObjects.Add(xlSrcRange, exportRange, , xlYes).Name = "o" & LCase$(sheetName)
        targetSheet.Cells(DictEndRow() + 1, StartColumn()).Font.color = vbBlue
        targetSheet.ListObjects("o" & LCase$(sheetName)).TableStyle = ""

    End If
End Sub

'@label:translate
'@sub-title Translate dictionary labels
'@details
'Delegates the translation of user-facing columns (labels, notes, sections) to
'the provided translation object while protecting formula columns from direct
'text replacement.
'@params
'- TransObject ITranslationObject responsible for translations
Public Sub Translate(ByVal TransObject As ITranslationObject)
    'Delegate translation of human-readable columns to the provided translation
    'object. Control details may contain formulas, so we pass the
    '`containsFormula` flag when appropriate to avoid breaking references.
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
End Sub


'@section Checkings
'===============================================================================

Private Sub LogInfo(ByVal label As String, Optional ByVal scope As Byte = checkingNote)
    'All noteworthy events (missing columns, normalisation steps, etc.) are
    'piped into the shared checking object so setup and designer screens can
    'surface diagnostics to the user.
    Dim checkName As String

    checkName = Wksh().Name

    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)
    internalChecks.Add CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1
    this.hasChecks = True
End Sub

Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

Private Property Get CheckingValues() As Object
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = internalChecks
End Property


'@section Errors
'===============================================================================

Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    'Standardised error helper used across the class. Centralising the wording
    'keeps raised errors consistent with those emitted by other components.
    Err.Raise errNumber, "LLdictionary",   errorMessage
    LogInfo errNumber & " - " & errorMessage, checkingError
End Sub


'@section Interface
'===============================================================================

'@jump:data
Private Property Get ILLdictionary_Data() As IDataSheet
    Set ILLdictionary_Data = Data
End Property

'@jump:datarange
Private Property Get ILLdictionary_DataRange(Optional ByVal colName As String = "__all__", _
                                             Optional ByVal includeHeaders As Boolean = False) As Range
    Set ILLdictionary_DataRange = DataRange(colName, includeHeaders)
End Property

'@jump:prepared
Private Property Get ILLdictionary_Prepared() As Boolean
    ILLdictionary_Prepared = Prepared
End Property

'@jump:exports
Private Property Get ILLdictionary_TotalNumberOfExports() As Long
    ILLdictionary_TotalNumberOfExports = TotalNumberOfExports
End Property

Private Property Let ILLdictionary_TotalNumberOfExports(ByVal numberOfExports As Long)
    TotalNumberOfExports = numberOfExports
End Property

'@jump:columnexists
Private Function ILLdictionary_ColumnExists(ByVal colName As String, Optional ByVal checkValidity As Boolean) As Boolean
    ILLdictionary_ColumnExists = ColumnExists(colName, checkValidity)
End Function

'@jump:uniquevalues
Private Function ILLdictionary_UniqueValues(ByVal colName As String) As BetterArray
    Set ILLdictionary_UniqueValues = UniqueValues(colName)
End Function

'@jump:variableexists
Private Function ILLdictionary_VariableExists(ByVal varName As String) As Boolean
    ILLdictionary_VariableExists = VariableExists(varName)
End Function

'@jump:addcolumn
Private Sub ILLdictionary_AddColumn(ByVal colName As String)
    AddColumn colName
End Sub

'@jump:insertcolumn
Private Sub ILLdictionary_InsertColumn(ByVal colName As String, ByVal After As String)
    InsertColumn colName, After
End Sub

'@jump:removecolumn
Private Sub ILLdictionary_RemoveColumn(ByVal colName As String)
    RemoveColumn colName
End Sub

'@jump:addrows
Private Sub ILLdictionary_AddRows()
    AddRows
End Sub

'@jump:removerows
Private Sub ILLdictionary_RemoveRows()
    RemoveRows
End Sub

Private Sub ILLdictionary_ManageRows(Optional ByVal del As Boolean = False)
    ManageRows del
End Sub

'@jump:choicesvars
Private Property Get ILLdictionary_ChoicesVars() As BetterArray
    Set ILLdictionary_ChoicesVars = ChoicesVars
End Property

'@jump:geovars
Private Property Get ILLdictionary_GeoVars() As BetterArray
    Set ILLdictionary_GeoVars = GeoVars
End Property

'@jump:timevars
Private Property Get ILLdictionary_TimeVars() As BetterArray
    Set ILLdictionary_TimeVars = TimeVars
End Property

'@jump:specialvars
Private Property Get ILLdictionary_SpecialVars(ByVal firstCondition As String, _
                                               Optional ByVal secondCondition As String = vbNullString, _
                                               Optional ByVal conditionName As String = "Control") As BetterArray
    Set ILLdictionary_SpecialVars = SpecialVars(firstCondition, secondCondition, conditionName)
End Property

'@jump:prepare
Private Sub ILLdictionary_Prepare(Optional ByVal PreservedSheetNames As BetterArray, _
                                  Optional ByVal geoObject As Object, _
                                  Optional ByVal tablePrefix As String = "table")
    Prepare PreservedSheetNames, geoObject, tablePrefix
End Sub

'@jump:clean
Private Sub ILLdictionary_Clean(Optional ByVal removeAddedColumns As Boolean = False)
    Clean removeAddedColumns
End Sub

'@jump:import
Private Sub ILLdictionary_Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    Import fromWksh, fromStartRow, fromStartcol
End Sub

'@jump:export
Private Sub ILLdictionary_Export(ByVal toWkb As Workbook, _
                                 Optional ByVal exportType As String = "__all__", _
                                 Optional ByVal addListObject As Boolean = True, _
                                 Optional ByVal Hide As Long = xlSheetHidden)
    Export toWkb, exportType, addListObject, Hide
End Sub

'@jump:translate
Private Sub ILLdictionary_Translate(ByVal TransObject As ITranslationObject)
    Translate TransObject
End Sub
