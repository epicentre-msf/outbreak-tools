VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecsColumnMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Caches table specification headers for constant-time lookups")
'@depends BetterArray
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableSpecsColumnMap

Private Type TColumnMapState
    header As Range
    tableRow As Range
    columnEntries As BetterArray
    aliasEntries As BetterArray
    isPrimed As Boolean
End Type

Private this As TColumnMapState
Private Const CLASSNAME As String = "TableSpecsColumnMap"

'@section Lifecycle
'===============================================================================
'Ensure internal caches start in a known state as soon as the class is created.
Private Sub Class_Initialize()
    ResetCaches
End Sub

'@section Factory
'===============================================================================
'@label:Create
'Expose a convenience factory that validates inputs and primes the cache once.
Public Function Create(ByVal headerRange As Range, _
                       ByVal tableRange As Range) As ITableSpecsColumnMap

    ValidateRangeParameter headerRange, "headerRange"
    ValidateRangeParameter tableRange, "tableRange"

    With New TableSpecsColumnMap
        Set .HeaderRange = headerRange
        Set .TableRange = tableRange
        .Refresh
        Set Create = .Self
    End With
End Function

'@label:Self
'Return the current instance cast as its interface to encourage loose coupling.
Public Property Get Self() As ITableSpecsColumnMap
    Set Self = Me
End Property

'@section Properties
'===============================================================================
'@label:GetHeaderRange
'Expose the header range bound to the cache.
Public Property Get HeaderRange() As Range
    Set HeaderRange = this.header
End Property

'@label:SetHeaderRange
'Assign a new header range and mark caches as invalid to be rebuilt on demand.
Public Property Set HeaderRange(ByVal headerRange As Range)
    ValidateRangeParameter headerRange, "headerRange"
    Set this.header = headerRange
    InvalidateCaches
End Property

'@label:GetTableRange
'Return the table/specification range that stores values associated to headers.
Public Property Get TableRange() As Range
    Set TableRange = this.tableRow
End Property

'@label:SetTableRange
'Assign a new table/specification range and invalidate cached results.
Public Property Set TableRange(ByVal tableRange As Range)
    ValidateRangeParameter tableRange, "tableRange"
    Set this.tableRow = tableRange
    InvalidateCaches
End Property

'@section Public API
'===============================================================================
'@label:Invalidate
'Mark the cache as dirty so the next lookup rebuilds column metadata.
Public Sub Invalidate()
    InvalidateCaches
End Sub

'@label:Refresh
'Rebuild the cache immediately to avoid incurring work on the next lookup.
Public Sub Refresh()
    PrimeCache
End Sub

'@label:ColumnExists
'Check if a column or fragment exists in the cached header list.
Public Function ColumnExists(ByVal columnName As String) As Boolean
    Dim columnIndexValue As Long
    ColumnExists = TryResolveColumnIndex(columnName, columnIndexValue)
End Function

'@label:ColumnIndex
'Retrieve the one-based position of a column using the cached metadata.
Public Function ColumnIndex(ByVal columnName As String) As Long
    Dim columnIndexValue As Long

    If TryResolveColumnIndex(columnName, columnIndexValue) Then
        ColumnIndex = columnIndexValue
    Else
        ColumnIndex = -1
    End If
End Function

'@label:Value
'Return the cell value located at the cached column index within the table.
Public Function Value(ByVal columnName As String) As String
    Dim columnIndexValue As Long
    Dim targetCell As Range

    Value = vbNullString

    If Not TryResolveColumnIndex(columnName, columnIndexValue) Then Exit Function
    Set targetCell = ResolveValueCell(columnIndexValue)
    If targetCell Is Nothing Then Exit Function

    Value = ReadCellText(targetCell)
End Function

'@section Cache internals
'===============================================================================
'@label:PrimeCache
'Build the header cache once, capturing column positions and normalised keys.
Private Sub PrimeCache()
    EnsureHeaderRangeReady
    ResetCaches

    BuildColumnEntries
    this.isPrimed = True
End Sub

'@label:InvalidateCaches
'Reset cached structures without clearing the configured ranges.
Private Sub InvalidateCaches()
    ResetCaches
    this.isPrimed = False
End Sub

'@label:ResetCaches
'Initialise array-backed caches used to store header metadata and aliases.
Private Sub ResetCaches()
    Set this.columnEntries = New BetterArray
    this.columnEntries.LowerBound = 1

    Set this.aliasEntries = New BetterArray
    this.aliasEntries.LowerBound = 1
End Sub

'@label:EnsurePrimed
'Safeguard lookups by priming the cache lazily the first time it is required.
Private Sub EnsurePrimed()
    If Not this.isPrimed Then PrimeCache
End Sub

'@label:BuildColumnEntries
'Iterate the header range once to record each column name and its position.
Private Sub BuildColumnEntries()
    Dim headerRow As Range
    Dim headerCell As Range
    Dim normalizedKey As String
    Dim headerText As String
    Dim columnIndexValue As Long

    Set headerRow = this.header.Rows(1)
    columnIndexValue = 0

    For Each headerCell In headerRow.Cells
        columnIndexValue = columnIndexValue + 1
        headerText = Trim$(CStr(headerCell.Value))
        normalizedKey = NormalizeKey(headerText)

        If normalizedKey <> vbNullString Then
            RecordColumnEntry headerText, normalizedKey, columnIndexValue
            StoreAlias normalizedKey, columnIndexValue
        End If
    Next headerCell
End Sub

'@label:RecordColumnEntry
'Store a new header entry if the normalised key has not been registered yet.
Private Sub RecordColumnEntry(ByVal headerText As String, _
                              ByVal normalizedKey As String, _
                              ByVal columnIndexValue As Long)
    Dim entry As Variant

    If ColumnEntryExists(normalizedKey) Then Exit Sub

    entry = Array(headerText, normalizedKey, columnIndexValue)
    this.columnEntries.Push entry
End Sub

'@label:ColumnEntryExists
'Determine whether the cache already contains the provided normalised key.
Private Function ColumnEntryExists(ByVal normalizedKey As String) As Boolean
    Dim idx As Long
    Dim entry As Variant

    ColumnEntryExists = False

    If this.columnEntries Is Nothing Then Exit Function
    If this.columnEntries.Length = 0 Then Exit Function

    For idx = this.columnEntries.LowerBound To this.columnEntries.UpperBound
        entry = this.columnEntries.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(1)), normalizedKey, vbBinaryCompare) = 0 Then
                ColumnEntryExists = True
                Exit Function
            End If
        End If
    Next idx
End Function

'@label:TryResolveColumnIndex
'Centralise the logic that resolves a column index using cache aliases and fallbacks.
Private Function TryResolveColumnIndex(ByVal columnName As String, _
                                       ByRef columnIndexValue As Long) As Boolean
    Dim normalizedKey As String

    TryResolveColumnIndex = False
    columnIndexValue = 0

    normalizedKey = NormalizeKey(columnName)
    If normalizedKey = vbNullString Then Exit Function

    EnsurePrimed

    If TryGetAlias(normalizedKey, columnIndexValue) Then
        TryResolveColumnIndex = True
        Exit Function
    End If

    If TryGetIndexForNormalizedKey(normalizedKey, columnIndexValue) Then
        StoreAlias normalizedKey, columnIndexValue
        TryResolveColumnIndex = True
        Exit Function
    End If

    If TryGetIndexByPartialMatch(normalizedKey, columnIndexValue) Then
        StoreAlias normalizedKey, columnIndexValue
        TryResolveColumnIndex = True
    End If
End Function

'@label:TryGetIndexForNormalizedKey
'Resolve an index using an exact match on the normalised header name.
Private Function TryGetIndexForNormalizedKey(ByVal normalizedKey As String, _
                                             ByRef columnIndexValue As Long) As Boolean
    Dim idx As Long
    Dim entry As Variant

    TryGetIndexForNormalizedKey = False
    columnIndexValue = 0

    If this.columnEntries Is Nothing Then Exit Function
    If this.columnEntries.Length = 0 Then Exit Function

    For idx = this.columnEntries.LowerBound To this.columnEntries.UpperBound
        entry = this.columnEntries.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(1)), normalizedKey, vbBinaryCompare) = 0 Then
                columnIndexValue = CLng(entry(2))
                TryGetIndexForNormalizedKey = True
                Exit Function
            End If
        End If
    Next idx
End Function

'@label:TryGetIndexByPartialMatch
'Fallback resolution that mimics Range.Find with partial, case-insensitive matching.
Private Function TryGetIndexByPartialMatch(ByVal normalizedKey As String, _
                                           ByRef columnIndexValue As Long) As Boolean
    Dim idx As Long
    Dim entry As Variant
    Dim headerKey As String

    TryGetIndexByPartialMatch = False
    columnIndexValue = 0

    If this.columnEntries Is Nothing Then Exit Function
    If this.columnEntries.Length = 0 Then Exit Function

    For idx = this.columnEntries.LowerBound To this.columnEntries.UpperBound
        entry = this.columnEntries.Item(idx)
        If IsArray(entry) Then
            headerKey = CStr(entry(1))
            If InStr(1, headerKey, normalizedKey, vbBinaryCompare) > 0 Then
                columnIndexValue = CLng(entry(2))
                TryGetIndexByPartialMatch = True
                Exit Function
            End If
        End If
    Next idx
End Function

'@label:TryGetAlias
'Look up an already resolved alias to avoid scanning the header list again.
Private Function TryGetAlias(ByVal normalizedKey As String, _
                             ByRef columnIndexValue As Long) As Boolean
    Dim idx As Long
    Dim entry As Variant

    TryGetAlias = False
    columnIndexValue = 0

    If this.aliasEntries Is Nothing Then Exit Function
    If this.aliasEntries.Length = 0 Then Exit Function

    For idx = this.aliasEntries.LowerBound To this.aliasEntries.UpperBound
        entry = this.aliasEntries.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(0)), normalizedKey, vbBinaryCompare) = 0 Then
                columnIndexValue = CLng(entry(1))
                TryGetAlias = True
                Exit Function
            End If
        End If
    Next idx
End Function

'@label:StoreAlias
'Persist an alias so future lookups for the same fragment shortcut directly to the index.
Private Sub StoreAlias(ByVal normalizedKey As String, ByVal columnIndexValue As Long)
    Dim entry As Variant

    If normalizedKey = vbNullString Then Exit Sub
    If columnIndexValue <= 0 Then Exit Sub

    EnsureAliasCache
    RemoveAlias normalizedKey

    entry = Array(normalizedKey, columnIndexValue)
    this.aliasEntries.Push entry
End Sub

'@label:EnsureAliasCache
'Create the alias cache when it has not been initialised yet.
Private Sub EnsureAliasCache()
    If this.aliasEntries Is Nothing Then
        Set this.aliasEntries = New BetterArray
        this.aliasEntries.LowerBound = 1
    End If
End Sub

'@label:RemoveAlias
'Drop an existing alias to keep mappings unique per key.
Private Sub RemoveAlias(ByVal normalizedKey As String)
    Dim idx As Long

    If this.aliasEntries Is Nothing Then Exit Sub
    If this.aliasEntries.Length = 0 Then Exit Sub

    For idx = this.aliasEntries.UpperBound To this.aliasEntries.LowerBound Step -1
        If IsArray(this.aliasEntries.Item(idx)) Then
            If StrComp(CStr(this.aliasEntries.Item(idx)(0)), normalizedKey, vbBinaryCompare) = 0 Then
                this.aliasEntries.Remove idx
            End If
        End If
    Next idx
End Sub

'@label:ResolveValueCell
'Fetch the cell within the table range that corresponds to the cached index.
Private Function ResolveValueCell(ByVal columnIndexValue As Long) As Range
    Dim dataRow As Range
    Dim maxColumns As Long

    Set ResolveValueCell = Nothing

    EnsureTableRangeReady

    Set dataRow = this.tableRow.Rows(1)
    maxColumns = dataRow.Columns.Count

    If columnIndexValue < 1 Then Exit Function
    If columnIndexValue > maxColumns Then Exit Function

    On Error Resume Next
        Set ResolveValueCell = dataRow.Cells(1, columnIndexValue)
    On Error GoTo 0
End Function

'@label:ReadCellText
'Extract text from a cell while gracefully handling errors.
Private Function ReadCellText(ByVal targetCell As Range) As String
    On Error Resume Next
        ReadCellText = CStr(targetCell.Value)
    On Error GoTo 0
End Function

'@label:NormalizeKey
'Normalise keys by trimming whitespace and converting to lowercase.
Private Function NormalizeKey(ByVal value As String) As String
    NormalizeKey = LCase$(Trim$(value))
End Function

'@label:ValidateRangeParameter
'Ensure the provided range parameter is a valid Range reference.
Private Sub ValidateRangeParameter(ByVal candidate As Range, ByVal argumentName As String)
    If candidate Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   argumentName & " cannot be Nothing"
        Exit Sub
    End If

    If StrComp(TypeName(candidate), "Range", vbBinaryCompare) <> 0 Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   argumentName & " must be a valid Range reference"
    End If
End Sub

'@label:EnsureHeaderRangeReady
'Raise a descriptive error when the header range has not been configured yet.
Private Sub EnsureHeaderRangeReady()
    If this.header Is Nothing Then
        ThrowError CLng(ProjectError.ObjectNotInitialized), _
                   "Header range must be assigned before using the column map"
    End If
End Sub

'@label:EnsureTableRangeReady
'Raise a descriptive error when the table range is missing while retrieving values.
Private Sub EnsureTableRangeReady()
    If this.tableRow Is Nothing Then
        ThrowError CLng(ProjectError.ObjectNotInitialized), _
                   "Table range must be assigned before reading values"
    End If
End Sub

'@label:ThrowError
'Centralised error helper leveraging ProjectError codes for consistency.
Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function ITableSpecsColumnMap_Create(ByVal headerRange As Range, _
                                             ByVal tableRange As Range) As ITableSpecsColumnMap
    Set ITableSpecsColumnMap_Create = Create(headerRange, tableRange)
End Function

Private Property Get ITableSpecsColumnMap_Self() As ITableSpecsColumnMap
    Set ITableSpecsColumnMap_Self = Self
End Property

Private Property Get ITableSpecsColumnMap_HeaderRange() As Range
    Set ITableSpecsColumnMap_HeaderRange = HeaderRange
End Property

Private Property Set ITableSpecsColumnMap_HeaderRange(ByVal headerRange As Range)
    Set Me.HeaderRange = headerRange
End Property

Private Property Get ITableSpecsColumnMap_TableRange() As Range
    Set ITableSpecsColumnMap_TableRange = TableRange
End Property

Private Property Set ITableSpecsColumnMap_TableRange(ByVal tableRange As Range)
    Set Me.TableRange = tableRange
End Property

Private Sub ITableSpecsColumnMap_Invalidate()
    Invalidate
End Sub

Private Sub ITableSpecsColumnMap_Refresh()
    Refresh
End Sub

Private Function ITableSpecsColumnMap_ColumnExists(ByVal columnName As String) As Boolean
    ITableSpecsColumnMap_ColumnExists = ColumnExists(columnName)
End Function

Private Function ITableSpecsColumnMap_ColumnIndex(ByVal columnName As String) As Long
    ITableSpecsColumnMap_ColumnIndex = ColumnIndex(columnName)
End Function

Private Function ITableSpecsColumnMap_Value(ByVal columnName As String) As String
    ITableSpecsColumnMap_Value = Value(columnName)
End Function
