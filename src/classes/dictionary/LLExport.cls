VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manage linelist and setup exports"

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Manage linelist and setup exports")

'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, ProcedureNotUsed, SuperfluousAnnotationArgument


'@class LLExport
'@description
'Manages the export specification sheet, the table where each row defines one
'export of a linelist. Create binds the worksheet through a DataSheet.
'NumberOfExports counts the rows, IsActive says whether an export is switched
'on, ActiveExportNumbers lists the ones that are, and ColumnValue reads any
'column of one row. ManageRows, InsertRows, DeleteRows, AddRows and RemoveRows
'work on the rows, Sort orders the table, Translate rewrites the two translated
'columns, and ImportSpecs and ExportSpecs move the sheet between workbooks.
'SyncDictionaryExports keeps the export columns of the dictionary in step with
'this table.
'
'THE FILE NAME OF AN EXPORT
'-------------------------------------------------------------------------------
'ExportFileName builds the name one export is saved under. The template comes
'from the "file name" column, and its chunks are replaced from the dictionary.
'A version suffix read from Passwords and a timestamp are added at the end, and
'the whole name is cleaned of the characters a file name cannot carry. An
'export asked for with exportAll is named from the workbook instead. An empty
'template falls back to a default one.
'
'WHAT THE INSTANCE HOLDS
'-------------------------------------------------------------------------------
'The specification block and its header row are read once and answered from
'memory. The class drops that copy whenever it writes to the table, and
'ResetCaches drops it for a caller that wrote to the sheet itself.
'
'WHAT IS REPORTED
'-------------------------------------------------------------------------------
'Entries filed while a member runs leave through HasCheckings and
'CheckingValues.
'@depends BetterArray, Checking, CustomTable, DataSheet, LLdictionary, LLVariables, Passwords, TranslationObject

Option Explicit


Private Type TLLExport
    expData As DataSheet
    collaboratorsMerged As Boolean          'True once the data sheet's entries were taken
    exportSheet As Worksheet
    dictInt As LLdictionary
    variables As LLVariables
    checks As Checking
    hasChecks As Boolean
    checkCounter As Long
    sanitizeList As Variant
    specHeaders As Variant
    specBlock As Variant
    specLoaded As Boolean
    exportTableAdapter As CustomTable       'The last table adapter ExportTable built
    sealed As Boolean
End Type

Private Const COL_STATUS As String = "status"
Private Const COL_FILE_NAME As String = "file name"
Private Const COL_LABEL_BUTTON As String = "label button"
Private Const COL_EXPORT_NUMBER As String = "export number"
Private Const COL_INCLUDE_IDS As String = "include personal identifiers"
Private Const STATUS_ACTIVE As String = "active"
Private Const DEFAULT_INCLUDE_IDS As String = "no"
Private Const TIMESTAMP_FORMAT As String = "yyyymmdd-HhNn"
Private Const CHECKING_CONTEXT As String = "LLExport checks"
Private Const CLASSNAME As String = "LLExport"

Private this As TLLExport

'@section Factory
'===============================================================================
'@label:Create
'@prop-title Create a unified export handler
'@details
'Creates a linelist export handler over the specification sheet provided by the
'caller. The implementation caches the underlying datasheet, worksheet, and
'sanitisation rules so subsequent calls can resolve filenames and table rows
'without repeatedly scanning the workbook.
'@param exportSheet Worksheet hosting the export specification table.
'@param startRow Optional; first row of the export table (defaults to 1).
'@param startColumn Optional; first column of the export table (defaults to 1).
'@return A fully initialised `LLExport` instance.
'@throws ProjectError.ObjectNotInitialized when the worksheet is missing.
Public Function Create(ByVal exportSheet As Worksheet, _
                       Optional ByVal startRow As Long = 1, _
                       Optional ByVal startColumn As Long = 1) As LLExport

    Dim exportData As DataSheet

    If exportSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Worksheet is required to create " & CLASSNAME
    End If

    Set exportData = DataSheet.Create(exportSheet, startRow, startColumn, True, "ExportsDataSheet")

    Dim newExport As LLExport
    Set newExport = New LLExport
    Set newExport.InternalData = exportData
    Set newExport.InternalWorksheet = exportSheet
    newExport.ResetCaches
    newExport.Seal
    Set Create = newExport
End Function

'@label:SetInternalData
'@param exportData Datasheet wrapper over the specification table.
Public Property Set InternalData(ByVal exportData As DataSheet)
    GuardNotSealed "InternalData"
    Set this.expData = exportData
End Property

'@label:SetInternalWorksheet
'@param exportSheet Worksheet that stores the export specification table.
Public Property Set InternalWorksheet(ByVal exportSheet As Worksheet)
    GuardNotSealed "InternalWorksheet"
    Set this.exportSheet = exportSheet
End Property

'@label:Seal
'@prop-title Prevent further changes to creation-only setters
Public Sub Seal()
    this.sealed = True
End Sub

'@label:GuardNotSealed
'@prop-title Guard creation-only setters after sealing
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@section Properties
'===============================================================================
'@label:Data
'@prop-title Underlying export datasheet
'@return Datasheet abstraction wrapping the export specification table.
Public Property Get Data() As DataSheet
    EnsureDataReady
    Set Data = this.expData
End Property

'@label:NumberOfExports
'@prop-title Count the configured export rows
'@details
'Counts rows within the specification table excluding headers. The count is
'read from the table itself, so it always agrees with what the sheet holds.
'@return Total number of export definitions currently present.
Public Property Get NumberOfExports() As Long
    EnsureDataReady
    EnsureWorksheetReady
    NumberOfExports = CountExportRows()
End Property

'@label:HasCheckings
'@prop-title Whether informational checkings were captured
'@label:harvest-collaborator-checkings
'@sub-title Fold what a collaborator filed into these entries
'@details
'`DataSheet` and `CustomTable` each keep a store of their own, and nobody
'read either of them before this line.
'@param collaboratorChecks Checking. What the collaborator filed.
Private Sub HarvestCollaboratorCheckings(ByVal collaboratorChecks As Checking)
    If collaboratorChecks Is Nothing Then Exit Sub
    If collaboratorChecks.Length() <= 0 Then Exit Sub

    If this.checks Is Nothing Then Set this.checks = Checking.Create(CHECKING_CONTEXT)
    this.checks.Append collaboratorChecks
    this.hasChecks = True
End Sub

'@label:merge-collaborator-checkings
'@sub-title Take the collaborators' entries before the report is read
'@details
'The data sheet lives as long as this object does; the table adapter is the
'last one `ExportTable` built. Both are folded in when the report is asked
'for, and the flag keeps the fold to once because the caller reads
'`HasCheckings` and then `CheckingValues`.
Private Sub MergeCollaboratorCheckings()
    If this.collaboratorsMerged Then Exit Sub
    this.collaboratorsMerged = True

    If Not (this.expData Is Nothing) Then
        If this.expData.HasCheckings Then HarvestCollaboratorCheckings this.expData.CheckingValues
    End If

    If Not (this.exportTableAdapter Is Nothing) Then
        If this.exportTableAdapter.HasCheckings Then _
            HarvestCollaboratorCheckings this.exportTableAdapter.CheckingValues
    End If
End Sub

Public Property Get HasCheckings() As Boolean
    MergeCollaboratorCheckings
    HasCheckings = this.hasChecks
End Property

'@label:CheckingValues
'@prop-title Collected informational diagnostics
'@return Checking object containing log entries, if any.
Public Property Get CheckingValues() As Checking
    If this.hasChecks Then
        Set CheckingValues = this.checks
    End If
End Property

'@section Row Management
'===============================================================================

'@label:ManageRows
'@prop-title Align export rows and dictionary columns
'@details
'Handle the addition or removal of export rows while keeping the dictionary
'columns aligned with the actual identifiers present in the worksheet. The
'logic operates on identifiers, so arbitrary deletions leave unrelated export
'columns in place.
'@param del Optional flag switching to row removal mode.
'@param dict Optional dictionary used to add/remove export columns.
'@param rowCount Optional threshold for row removal; rows with non-empty cells at or below the threshold are trimmed (defaults to 2).
Public Sub ManageRows(Optional ByVal del As Boolean = False, _
                      Optional ByVal dict As LLdictionary, _
                      Optional ByVal rowCount As Long = 2)

    Dim tableObj As CustomTable
    Dim previousCount As Long
    Dim exportNumbers As BetterArray

    EnsureDataReady
    EnsureWorksheetReady
    ResolveDictionary dict, required:=False

    Set tableObj = ExportTable()
    previousCount = CountExportRows()

    If del Then
        'Allow rowCount = 0, no problem with that here, nothing is done
        'rowcount default value is 2 because we have at least 2 values on one row (export id, and identifier)
        If rowCount < 0 Then rowCount = 2
        tableObj.RemoveRows totalCount:=rowCount, includeIds:=False
    Else
        tableObj.AddRows nbRows:=1, includeIds:=False
        ApplyIncludeDefaults tableObj, previousCount + 1, CountExportRows()
    End If

    Set exportNumbers = EnsureExportIdentifiers(tableObj)
    DropSpecification

    If DictionaryIsAvailable() Then SyncDictionaryExportsCore exportNumbers, this.dictInt
End Sub

'@label:InsertRows
'@prop-title Insert export rows and keep dictionary columns aligned
'@param targetCell Range anchor representing where to mirror inserted rows.
'@param dict Optional dictionary to shift export columns.
'@param insertShift Optional Boolean to insert worksheet rows when necessary. Defaults to False.
Public Sub InsertRows(ByVal targetCell As Range, _
                      Optional ByVal dict As LLdictionary, _
                      Optional ByVal insertShift As Boolean = False)

    Dim tableObj As CustomTable
    Dim exportNumbers As BetterArray
    Dim insertIndex As Long
    Dim lastDefaultRow As Long
    Dim actualInserted As Long
    Dim previousCount As Long
    Dim newCount As Long

    If targetCell Is Nothing Then Exit Sub

    EnsureDataReady
    EnsureWorksheetReady
    ResolveDictionary dict, required:=False

    Set tableObj = ExportTable()
    previousCount = CountExportRows()

    insertIndex = ResolveSelectionRowIndex(tableObj, targetCell)
    If insertIndex <= 0 Then insertIndex = previousCount + 1
    If insertIndex > previousCount + 1 Then insertIndex = previousCount + 1

    tableObj.InsertRowsAt targetCell:=targetCell, insertShift:=insertShift, includeIds:=False

    newCount = CountExportRows()
    actualInserted = newCount - previousCount

    If actualInserted <= 0 Then
        DropSpecification
        Exit Sub
    End If

    lastDefaultRow = insertIndex + actualInserted - 1
    If lastDefaultRow > newCount Then lastDefaultRow = newCount

    ApplyIncludeDefaults tableObj, insertIndex, lastDefaultRow
    Set exportNumbers = AssignExportNumbers(tableObj, insertIndex, actualInserted)
    DropSpecification

    If DictionaryIsAvailable() Then SyncDictionaryExportsCore exportNumbers, this.dictInt
End Sub

'@label:DeleteRows
'@prop-title Delete export rows intersecting the supplied selection
'@param targetCell Range holding the rows the user selected.
'@param dict Optional dictionary to prune export columns from.
'@param forceShift Optional Boolean asking the table to shift worksheet rows up.
Public Sub DeleteRows(ByVal targetCell As Range, _
                      Optional ByVal dict As LLdictionary, _
                      Optional ByVal forceShift As Boolean = False)

    Dim tableObj As CustomTable
    Dim exportNumbers As BetterArray
    Dim previousCount As Long
    Dim newCount As Long

    If targetCell Is Nothing Then Exit Sub

    EnsureDataReady
    EnsureWorksheetReady
    ResolveDictionary dict, required:=False

    Set tableObj = ExportTable()
    previousCount = CountExportRows()

    tableObj.DeleteRowsAt targetCell:=targetCell, _
                          includeIds:=False, _
                          forceShift:=forceShift

    newCount = CountExportRows()
    If newCount = previousCount Then
        DropSpecification
        Exit Sub
    End If

    Set exportNumbers = EnsureExportIdentifiers(tableObj)
    DropSpecification

    If DictionaryIsAvailable() Then SyncDictionaryExportsCore exportNumbers, this.dictInt
End Sub

'@label:AddRows
'@prop-title Append export definition rows
'@param dict Optional dictionary to align export columns with.
Public Sub AddRows(Optional ByVal dict As LLdictionary)
    ManageRows dict:=dict
End Sub

'@label:RemoveRows
'@prop-title Remove empty export rows
'@param dict Optional dictionary to align export columns with.
'@param rowCount Optional threshold; rows with non-empty cell counts at or below this value are removed (defaults to 2).
Public Sub RemoveRows(Optional ByVal dict As LLdictionary, _
                      Optional ByVal rowCount As Long = 2)
    ManageRows del:=True, dict:=dict, rowCount:=rowCount
End Sub

'@label:SyncDictionaryExports
'@prop-title Synchronise dictionary export columns
'@param dict Optional dictionary; the cached one is used when this is missing.
Public Sub SyncDictionaryExports(Optional ByVal dict As LLdictionary)
    Dim tableObj As CustomTable
    Dim exportNumbers As BetterArray

    EnsureDataReady
    EnsureWorksheetReady

    Set tableObj = ExportTable()
    Set exportNumbers = CollectExportNumbers(tableObj)

    ResolveDictionary dict, required:=False

    SyncDictionaryExportsCore exportNumbers, this.dictInt
End Sub

'@label:Sort
'@sub-title Sort export definitions and normalize identifiers
'@param dict Optional dictionary whose export columns follow the new order.
Public Sub Sort(Optional ByVal dict As LLdictionary)
    Dim tableObj As CustomTable
    Dim sortError As Long
    Dim sortMessage As String

    EnsureDataReady
    EnsureWorksheetReady

    Set tableObj = ExportTable()

    'A table that refuses to sort still gets its identifiers renumbered below,
    'so the failure is recorded and the routine carries on.
    Err.Clear
    On Error Resume Next
        tableObj.Sort colName:=COL_EXPORT_NUMBER
        sortError = Err.Number
        sortMessage = Err.Description
        If sortError <> 0 Then Err.Clear
    On Error GoTo 0

    If sortError <> 0 Then
        LogInfo "Sorting the export table failed: " & sortMessage, checkingWarning
    End If

    ResolveDictionary dict, required:=False

    SequentiallyRenameExports tableObj, this.dictInt
    DropSpecification
End Sub

'@section Row Coordination Helpers
'===============================================================================

'@label:ApplyIncludeDefaults
'@sub-title Write the default identifier flag into the rows that have none
'@param tableObj CustomTable wrapping the export specification table.
'@param startIndex First row of the range to fill, 1-based within the table body.
'@param endIndex Last row of the range to fill.
Private Sub ApplyIncludeDefaults(ByVal tableObj As CustomTable, _
                                 ByVal startIndex As Long, _
                                 ByVal endIndex As Long)
    Dim includeColumn As Range
    Dim cellValues As Variant
    Dim lastIndex As Long
    Dim idx As Long

    If startIndex <= 0 Then Exit Sub
    If endIndex < startIndex Then Exit Sub

    Set includeColumn = tableObj.DataRange(COL_INCLUDE_IDS)
    If includeColumn Is Nothing Then Exit Sub

    lastIndex = endIndex
    If lastIndex > includeColumn.Rows.Count Then lastIndex = includeColumn.Rows.Count
    If startIndex > lastIndex Then Exit Sub

    cellValues = ColumnValues(includeColumn)

    For idx = startIndex To lastIndex
        If LenB(CStr(cellValues(idx))) = 0 Then
            With includeColumn.Cells(idx, 1)
                .Value = DEFAULT_INCLUDE_IDS
                .Locked = False
            End With
        End If
    Next idx
End Sub

'@label:ResolveSelectionRowIndex
'@sub-title Turn a worksheet selection into a row index inside the table body
'@param tableObj CustomTable wrapping the export specification table.
'@param targetCell Range the user selected.
'@return Long. The 1-based row index, clamped to one past the last row.
Private Function ResolveSelectionRowIndex(ByVal tableObj As CustomTable, _
                                          ByVal targetCell As Range) As Long

    Dim exportRange As Range
    Dim relativeRow As Long

    Set exportRange = tableObj.DataRange(COL_EXPORT_NUMBER)
    If exportRange Is Nothing Then Exit Function

    relativeRow = targetCell.Row - exportRange.Row + 1
    If relativeRow < 1 Then relativeRow = 1
    If relativeRow > exportRange.Rows.Count Then relativeRow = exportRange.Rows.Count + 1

    ResolveSelectionRowIndex = relativeRow
End Function

'@label:AssignExportNumbers
'@sub-title Give the freshly inserted rows the next free export numbers
'@details
'The numbers in use are read once and kept in memory, so the routine answers
'every later question about them without going back to the worksheet.
'@param tableObj CustomTable wrapping the export specification table.
'@param startIndex First inserted row, 1-based within the table body.
'@param insertCount How many rows were inserted.
'@return BetterArray. Every export number in use once the rows are numbered.
Private Function AssignExportNumbers(ByVal tableObj As CustomTable, _
                                     ByVal startIndex As Long, _
                                     ByVal insertCount As Long) As BetterArray

    Dim exportRange As Range
    Dim existing As BetterArray
    Dim idx As Long
    Dim nextNumber As Long

    Set existing = CollectExportNumbers(tableObj)
    If existing Is Nothing Then Set existing = NewNumberList()
    Set AssignExportNumbers = existing

    If insertCount <= 0 Then Exit Function

    Set exportRange = tableObj.DataRange(COL_EXPORT_NUMBER)
    If exportRange Is Nothing Then Exit Function

    For idx = 0 To insertCount - 1
        nextNumber = NextAvailableExportNumber(existing)
        exportRange.Cells(startIndex + idx, 1).Value = ExportColumnName(nextNumber)
        existing.Push nextNumber
    Next idx
End Function

'@label:SequentiallyRenameExports
'@sub-title Renumber every export row from 1 and follow up in the dictionary
'@param tableObj CustomTable wrapping the export specification table.
'@param dict Optional dictionary whose export columns follow the new order.
Private Sub SequentiallyRenameExports(ByVal tableObj As CustomTable, _
                                      Optional ByVal dict As LLdictionary)

    Dim exportRange As Range
    Dim cellValues As Variant
    Dim written() As Variant
    Dim previousNames() As String
    Dim total As Long
    Dim idx As Long

    Set exportRange = tableObj.DataRange(COL_EXPORT_NUMBER)
    If exportRange Is Nothing Then Exit Sub

    total = exportRange.Rows.Count
    If total = 0 Then Exit Sub

    cellValues = ColumnValues(exportRange)

    ReDim previousNames(1 To total)
    ReDim written(1 To total, 1 To 1)

    For idx = 1 To total
        previousNames(idx) = CStr(cellValues(idx))
        written(idx, 1) = ExportColumnName(idx)
    Next idx

    exportRange.Value = written

    If Not dict Is Nothing Then
        RenameDictionaryExports dict, previousNames, total
    End If
End Sub

'@label:RenameDictionaryExports
'@sub-title Move the dictionary export columns onto the new sequence
'@details
'The rename runs in two passes through temporary names so an old "Export 3"
'can become "Export 1" while another column still carries that name.
'@param dict Dictionary holding the export columns.
'@param previousNames The export identifiers as they read before the sort.
'@param total How many export rows there are now.
Private Sub RenameDictionaryExports(ByVal dict As LLdictionary, _
                                    ByRef previousNames() As String, _
                                    ByVal total As Long)

    Dim placed As BetterArray
    Dim existing As BetterArray
    Dim idx As Long
    Dim pointer As Long
    Dim tempName As String
    Dim normalized As String

    If dict Is Nothing Then Exit Sub
    If total <= 0 Then Exit Sub

    'First pass: rename old export columns to temporary placeholders.
    For idx = 1 To total
        normalized = NormalizedExportHeader(previousNames(idx))
        If LenB(normalized) > 0 Then
            If dict.ColumnExists(normalized) Then
                tempName = TemporaryExportColumn(idx)
                dict.RenameColumn normalized, tempName, strictSearch:=True, matchCase:=False
            End If
        End If
    Next idx

    'Second pass: rename placeholders to sequential export names or create columns when missing.
    Set placed = NewNumberList()
    For idx = 1 To total
        tempName = TemporaryExportColumn(idx)
        If dict.ColumnExists(tempName) Then
            dict.RenameColumn tempName, ExportColumnName(idx), strictSearch:=True, matchCase:=True
        Else
            EnsureDictionaryExportColumn dict, idx, placed
        End If
        placed.Push idx
    Next idx

    'Remove extra export columns beyond the active count.
    Set existing = EnumerateDictionaryExportNumbers(dict)
    If Not existing Is Nothing Then
        For pointer = existing.UpperBound To existing.LowerBound Step -1
            If CLng(existing.Item(pointer)) > total Then
                dict.RemoveColumn ExportColumnName(CLng(existing.Item(pointer)))
            End If
        Next pointer
    End If

    dict.TotalNumberOfExports = total
End Sub

'@label:NextAvailableExportNumber
'@sub-title Pick the lowest export number that is still free
'@param existing BetterArray of the export numbers already in use.
'@return Long. The next free number.
Private Function NextAvailableExportNumber(ByVal existing As BetterArray) As Long
    Dim candidate As Long

    candidate = HighestNumber(existing)

    Do
        candidate = candidate + 1
    Loop While ContainsNumber(existing, candidate)

    NextAvailableExportNumber = candidate
End Function

'@label:HighestNumber
'@sub-title Largest value in a list of export numbers
'@param numbers BetterArray of export numbers, possibly Nothing.
'@return Long. The largest value, or zero for an empty list.
Private Function HighestNumber(ByVal numbers As BetterArray) As Long
    Dim idx As Long
    Dim current As Long

    If numbers Is Nothing Then Exit Function
    For idx = numbers.LowerBound To numbers.UpperBound
        current = CLng(numbers.Item(idx))
        If current > HighestNumber Then HighestNumber = current
    Next idx
End Function

'@label:NewNumberList
'@sub-title Build an empty 1-based list of export numbers
'@return BetterArray. An empty list ready to be pushed into.
Private Function NewNumberList() As BetterArray
    Dim numbers As BetterArray

    Set numbers = New BetterArray
    numbers.LowerBound = 1
    Set NewNumberList = numbers
End Function

'@label:TemporaryExportColumn
'@sub-title Placeholder header used while export columns swap names
'@param orderIndex Position in the new sequence.
'@return String. The placeholder header.
Private Function TemporaryExportColumn(ByVal orderIndex As Long) As String
    TemporaryExportColumn = "__exp_tmp_" & Format$(orderIndex, "0000")
End Function

'@label:NormalizedExportHeader
'@sub-title Turn an export identifier into the dictionary column header
'@param headerText The identifier as it reads on the export sheet.
'@return String. The matching column header, empty when the text carries no number.
Private Function NormalizedExportHeader(ByVal headerText As String) As String
    Dim parsed As Long

    parsed = ParseExportIdentifier(headerText)
    If parsed <= 0 Then Exit Function

    NormalizedExportHeader = ExportColumnName(parsed)
End Function

'@label:EnsureExportIdentifiers
'@sub-title Give every export row a unique identifier
'@details
'The identifier column is read as one block, settled in memory and written back
'in one go. A row whose number repeats an earlier one, or whose text carries no
'number, takes the next free number.
'@param tableObj CustomTable wrapping the export specification table.
'@return BetterArray. The export numbers in worksheet order.
Private Function EnsureExportIdentifiers(ByVal tableObj As CustomTable) As BetterArray
    Dim exportRange As Range
    Dim cellValues As Variant
    Dim written() As Variant
    Dim assigned() As Long
    Dim usedNumbers As BetterArray
    Dim result As BetterArray
    Dim idx As Long
    Dim parsed As Long
    Dim maxNumber As Long
    Dim total As Long

    Set exportRange = tableObj.DataRange(COL_EXPORT_NUMBER)
    If exportRange Is Nothing Then Exit Function

    cellValues = ColumnValues(exportRange)
    total = UBound(cellValues)
    If total <= 0 Then Exit Function

    ReDim assigned(1 To total)
    Set usedNumbers = NewNumberList()

    'First sighting of a number wins; every other row is left at zero.
    For idx = 1 To total
        parsed = ParseExportIdentifier(cellValues(idx))
        If parsed > 0 Then
            If Not ContainsNumber(usedNumbers, parsed) Then
                usedNumbers.Push parsed
                assigned(idx) = parsed
                If parsed > maxNumber Then maxNumber = parsed
            End If
        End If
    Next idx

    'Every row still at zero takes the next number above the highest in use.
    For idx = 1 To total
        If assigned(idx) = 0 Then
            maxNumber = maxNumber + 1
            assigned(idx) = maxNumber
        End If
    Next idx

    ReDim written(1 To total, 1 To 1)
    Set result = NewNumberList()

    For idx = 1 To total
        written(idx, 1) = ExportColumnName(assigned(idx))
        result.Push assigned(idx)
    Next idx

    exportRange.Value = written

    If result.Length = 0 Then Exit Function
    Set EnsureExportIdentifiers = result
End Function

'@label:CollectExportNumbers
'@sub-title Read the export numbers the specification table holds
'@details
'The identifier column is read as one block and parsed in memory.
'@param tableObj CustomTable wrapping the export specification table.
'@return BetterArray. The export numbers in worksheet order, Nothing when there are none.
Private Function CollectExportNumbers(ByVal tableObj As CustomTable) As BetterArray
    Dim exportRange As Range
    Dim cellValues As Variant
    Dim result As BetterArray
    Dim idx As Long
    Dim parsed As Long

    Set exportRange = tableObj.DataRange(COL_EXPORT_NUMBER)
    If exportRange Is Nothing Then Exit Function

    cellValues = ColumnValues(exportRange)
    Set result = NewNumberList()

    For idx = 1 To UBound(cellValues)
        parsed = ParseExportIdentifier(cellValues(idx))
        If parsed > 0 Then
            If Not ContainsNumber(result, parsed) Then
                result.Push parsed
            End If
        End If
    Next idx

    If result.Length = 0 Then Exit Function
    Set CollectExportNumbers = result
End Function

'@label:EnumerateDictionaryExportNumbers
'@sub-title Read the export numbers the dictionary columns carry
'@details
'The dictionary header row is read as one block, so a wide dictionary costs one
'worksheet crossing instead of one per column.
'@param dict Dictionary to inspect.
'@return BetterArray. The export numbers found, Nothing when there are none.
Private Function EnumerateDictionaryExportNumbers(ByVal dict As LLdictionary) As BetterArray
    Dim headers As Range
    Dim headerValues As Variant
    Dim result As BetterArray
    Dim idx As Long
    Dim parsed As Long

    If dict Is Nothing Then Exit Function

    Set headers = dict.DataRange(includeHeaders:=True)
    If headers Is Nothing Then Exit Function

    headerValues = RowValues(headers.Rows(1))
    Set result = NewNumberList()

    For idx = 1 To UBound(headerValues)
        parsed = ParseExportIdentifier(headerValues(idx))
        If parsed > 0 Then
            If Not ContainsNumber(result, parsed) Then
                result.Push parsed
            End If
        End If
    Next idx

    If result.Length = 0 Then Exit Function
    Set EnumerateDictionaryExportNumbers = result
End Function

'@label:ContainsNumber
'@sub-title Whether a list of export numbers holds a value
'@param numbers BetterArray of export numbers, possibly Nothing.
'@param target The value to look for.
'@return Boolean. True when the value is in the list.
Private Function ContainsNumber(ByVal numbers As BetterArray, ByVal target As Long) As Boolean
    If numbers Is Nothing Then Exit Function
    If numbers.Length = 0 Then Exit Function
    ContainsNumber = numbers.Includes(target)
End Function

'@label:ParseExportIdentifier
'@sub-title Read the number out of an export identifier
'@details
'Anything that starts with "export" and ends in digits gives that number.
'"export" on its own, and text such as "exported", give zero: the conversion
'runs under On Error Resume Next so a bad tail reads as "no number".
'@param rawValue The cell value to read.
'@return Long. The export number, or zero.
Private Function ParseExportIdentifier(ByVal rawValue As Variant) As Long
    Dim text As String
    Dim suffix As String

    If IsEmpty(rawValue) Then Exit Function

    text = LCase$(Trim$(CStr(rawValue)))
    If LenB(text) = 0 Then Exit Function

    If Left$(text, 6) <> "export" Then Exit Function

    suffix = Trim$(Mid$(text, 7))
    If LenB(suffix) = 0 Then Exit Function

    On Error Resume Next
        ParseExportIdentifier = CLng(suffix)
    On Error GoTo 0
End Function

'@label:SyncDictionaryExportsCore
'@sub-title Make the dictionary export columns match the export rows
'@details
'The dictionary header is read once and the list of export columns is kept in
'memory as columns are added, so a sync of n exports costs one header read.
'The dictionary is then told how many export columns it carries.
'@param exportNumbers BetterArray of the export numbers on the specification sheet.
'@param dict Dictionary to align.
Private Sub SyncDictionaryExportsCore(ByVal exportNumbers As BetterArray, ByVal dict As LLdictionary)
    Dim existing As BetterArray
    Dim idx As Long
    Dim number As Long
    Dim columnCount As Long

    If dict Is Nothing Then Exit Sub

    Set existing = EnumerateDictionaryExportNumbers(dict)
    If existing Is Nothing Then Set existing = NewNumberList()

    If Not exportNumbers Is Nothing Then
        columnCount = exportNumbers.Length
        For idx = exportNumbers.LowerBound To exportNumbers.UpperBound
            number = CLng(exportNumbers.Item(idx))
            If Not ContainsNumber(existing, number) Then
                EnsureDictionaryExportColumn dict, number, existing
                existing.Push number
            End If
        Next idx
    End If

    For idx = existing.LowerBound To existing.UpperBound
        number = CLng(existing.Item(idx))
        If Not ContainsNumber(exportNumbers, number) Then
            dict.RemoveColumn ExportColumnName(number)
        End If
    Next idx

    'The dictionary counter holds how many export COLUMNS it carries, which is
    'one per export row once the two sides agree.
    dict.TotalNumberOfExports = columnCount
End Sub

'@label:EnsureDictionaryExportColumn
'@sub-title Add one export column to the dictionary in the right place
'@details
'The new column goes right after the highest numbered export column below it,
'and after "unique" when there is none. The list of numbers already placed is
'passed in, so the dictionary header is read once per sync.
'@param dict Dictionary to add the column to.
'@param exportIndex Export number the column belongs to.
'@param existing BetterArray of the export numbers already present.
Private Sub EnsureDictionaryExportColumn(ByVal dict As LLdictionary, _
                                         ByVal exportIndex As Long, _
                                         ByVal existing As BetterArray)
    Dim columnName As String
    Dim anchorName As String
    Dim previousIdx As Long

    If dict Is Nothing Then Exit Sub

    columnName = ExportColumnName(exportIndex)
    If dict.ColumnExists(columnName) Then Exit Sub

    previousIdx = exportIndex - 1
    Do While previousIdx >= 1
        If ContainsNumber(existing, previousIdx) Then
            anchorName = ExportColumnName(previousIdx)
            Exit Do
        End If
        previousIdx = previousIdx - 1
    Loop

    If LenB(anchorName) = 0 Then
        If dict.ColumnExists("unique") Then
            anchorName = "unique"
        Else
            dict.AddColumn columnName
            Exit Sub
        End If
    End If

    If dict.ColumnExists(anchorName) Then
        dict.InsertColumn columnName, anchorName
    Else
        dict.AddColumn columnName
    End If
End Sub

'@label:ExportColumnName
'@sub-title Dictionary column header for an export number
'@details
'The header is capitalised here while `LLdictionary` lower-cases what it is
'given and searches with matchCase False. Both sides have to keep that true.
'@param exportIndex Export number.
'@return String. The column header.
Private Function ExportColumnName(ByVal exportIndex As Long) As String
    ExportColumnName = "Export " & CStr(exportIndex)
End Function

'@section SpecManagement
'===============================================================================

'@label:ImportSpecs
'@prop-title Import export specifications from another worksheet
'@param fromWksh Source worksheet containing the export table.
'@param fromStartRow First row of the source data.
'@param fromStartColumn First column of the source data.
'@throws ProjectError.ObjectNotInitialized when the source worksheet is missing.
Public Sub ImportSpecs(ByVal fromWksh As Worksheet, _
                       ByVal fromStartRow As Long, _
                       ByVal fromStartColumn As Long)

    Dim sourceData As DataSheet

    EnsureDataReady
    If fromWksh Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Source worksheet is missing"
    Set sourceData = DataSheet.Create(fromWksh, fromStartRow, fromStartColumn)

    this.expData.Import sourceData
    ResetCaches
End Sub

'@label:ExportSpecs
'@prop-title Export specifications into another workbook
'@param toWkb Destination workbook.
'@param Hide Optional visibility flag for the exported sheet.
'@throws ProjectError.ObjectNotInitialized when the target workbook is missing.
Public Sub ExportSpecs(ByVal toWkb As Workbook, _
                       Optional ByVal Hide As Long = xlSheetHidden)

    EnsureDataReady

    If toWkb Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Target workbook is missing"
    this.expData.Export toWkb, Hide:=Hide, includeNames:=True
End Sub


'@section FileName
'===============================================================================
'@label:ExportFileName
'@prop-title Build the filename for a specific export definition
'@param exportNumber Target export definition number starting at 1.
'@param dict Dictionary providing metadata for chunk substitutions.
'@param pass Password provider used to append version/public key information.
'@param exportAll Optional flag forcing the "export all" naming schema.
'@return Sanitised filename incorporating dictionary-driven chunks and metadata.
'@throws ProjectError.InvalidArgument when the export number is invalid.
'@throws ProjectError.ObjectNotInitialized when required collaborators are missing.
Public Function ExportFileName(ByVal exportNumber As Long, _
                               ByVal dict As LLdictionary, _
                               ByVal pass As Passwords, _
                               Optional ByVal exportAll As Boolean = False) As String

    Dim currWb As Workbook
    Dim template As String
    Dim sanitizedPath As String
    Dim suffix As String

    EnsureDataReady
    EnsureWorksheetReady
    ValidatePasswords pass
    ResolveDictionary dict

    Set currWb = this.exportSheet.Parent

    If exportAll Then
        ExportFileName = SanitizeChunk(WorkbookBaseName(currWb)) & "_export_all__" & BuildTimestamp
        Exit Function
    End If

    ValidateExportNumber exportNumber

    If Not IsActive(exportNumber) Then
        LogInfo "Export " & CStr(exportNumber) & " is inactive; generating filename for traceability"
    End If

    template = Trim$(ReadValue(COL_FILE_NAME, exportNumber, required:=True))

    If LenB(template) = 0 Then
        LogInfo "Empty file name template for export " & CStr(exportNumber)
        template = DefaultTemplate(exportNumber)
    End If

    sanitizedPath = ComposeTemplatePath(template)
    suffix = BuildVersionSuffix(pass)

    ExportFileName = sanitizedPath & suffix & "__" & BuildTimestamp
End Function

'@section Export Count
'===============================================================================

'@label:CountExportRows
'@sub-title How many rows the specification table holds
'@return Long. The row count, zero when there is no data store.
Private Function CountExportRows() As Long
    Dim exportRange As Range

    If this.expData Is Nothing Then Exit Function

    Set exportRange = this.expData.DataRange(includeHeaders:=False)
    If Not exportRange Is Nothing Then
        CountExportRows = exportRange.Rows.Count
    End If
End Function


'@section Utilities
'===============================================================================
'@label:IsActive
'@prop-title Determine whether an export is marked active
'@param exportNumber Export definition number to inspect.
'@return True when the status column contains "active" (case-insensitive).
Public Function IsActive(ByVal exportNumber As Long) As Boolean
    Dim statusValue As String
    statusValue = ReadValue(COL_STATUS, exportNumber, required:=False)
    IsActive = (StrComp(statusValue, STATUS_ACTIVE, vbTextCompare) = 0)
End Function

'@label:ColumnValue
'@prop-title Retrieve a column value for a given export row
'@param exportNumber Export definition number to read.
'@param columnName Column header to retrieve.
'@return Column value as string (empty string when not found).
Public Function ColumnValue(ByVal exportNumber As Long, ByVal columnName As String) As String
    ColumnValue = ReadValue(columnName, exportNumber, required:=False)
End Function

'@label:ActiveExportNumbers
'@prop-title Enumerate export numbers marked as active
'@return BetterArray containing active export numbers using 1-based numbering.
Public Function ActiveExportNumbers() As BetterArray
    Dim total As Long
    Dim idx As Long
    Dim result As BetterArray

    EnsureDataReady
    total = NumberOfExports()

    Set result = NewNumberList()

    For idx = 1 To total
        If IsActive(idx) Then result.Push idx
    Next idx

    Set ActiveExportNumbers = result
End Function

'@label:Translate
'@sub-title Translate the current LLExport
'@details
'A setup workbook that lacks one of the two translated columns is logged and
'the other column is still translated.
'@param trads Translation table to read the wording from.
'@throws ProjectError.ObjectNotInitialized when the translation object is missing.
Public Sub Translate(ByVal trads As TranslationObject)

    Dim labelRange As Range
    Dim fileNameRange As Range

    If trads Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Translations are not initialised"
    End If

    EnsureDataReady

    Set labelRange = SpecColumnRange(COL_LABEL_BUTTON)
    If labelRange Is Nothing Then
        LogInfo "Export column '" & COL_LABEL_BUTTON & "' is missing, nothing to translate", checkingWarning
    Else
        trads.TranslateRange labelRange
    End If

    Set fileNameRange = SpecColumnRange(COL_FILE_NAME)
    If fileNameRange Is Nothing Then
        LogInfo "Export column '" & COL_FILE_NAME & "' is missing, nothing to translate", checkingWarning
    Else
        trads.TranslateRange fileNameRange, containsFormula:=True
    End If

End Sub

'@section Private Helpers
'===============================================================================
'@label:EnsureDataReady
'@prop-title Ensure the datasheet state is initialised
Private Sub EnsureDataReady()
    If this.expData Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Export data store not initialised"
    End If
End Sub

'@label:EnsureWorksheetReady
'@prop-title Ensure the worksheet reference is initialised
Private Sub EnsureWorksheetReady()
    If this.exportSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Export worksheet not initialised"
    End If
End Sub

'@label:ResolveDictionary
'@prop-title Keep the dictionary the caller passed, or reuse the cached one
'@details
'A dictionary passed in becomes the cached one. When nothing is passed the
'cached one stands. This is the single place the class picks its dictionary.
'@param dict Optional dictionary supplied by the caller.
'@param required Optional; raise when no dictionary can be resolved. Defaults to True.
'@throws ProjectError.ObjectNotInitialized when required is True and none is found.
Private Sub ResolveDictionary(Optional ByVal dict As LLdictionary, _
                              Optional ByVal required As Boolean = True)
    If Not dict Is Nothing Then Set this.dictInt = dict

    If required And (this.dictInt Is Nothing) Then
        ThrowError ProjectError.ObjectNotInitialized, "LLExport Internal dictionary is required; none found"
    End If
End Sub

'@label:EnsureVariablesReady
'@prop-title Lazily instantiate dictionary-backed variable helper
'@param dict Optional dictionary the variables read from.
Private Sub EnsureVariablesReady(Optional ByVal dict As LLdictionary)
    Dim currentDict As LLdictionary
    Dim existingDict As LLdictionary

    ResolveDictionary dict

    If Not DictionaryIsAvailable() Then Exit Sub

    Set currentDict = InternalDictionary()

    If this.variables Is Nothing Then
        Set this.variables = LLVariables.Create(currentDict)
    Else
        Set existingDict = this.variables.Dictionary
        If (existingDict Is Nothing) Or (Not (existingDict Is currentDict)) Then
            Set this.variables = LLVariables.Create(currentDict)
        End If
    End If
End Sub

'@label:InternalDictionary
'@prop-title Dictionary the class is currently working with
'@return LLdictionary. The cached dictionary.
Private Property Get InternalDictionary() As LLdictionary
    ResolveDictionary
    Set InternalDictionary = this.dictInt
End Property

'@label:DictionaryIsAvailable
'@sub-title Whether a dictionary is cached
'@return Boolean. True when the class holds one.
Private Function DictionaryIsAvailable() As Boolean
    DictionaryIsAvailable = Not (this.dictInt Is Nothing)
End Function

'@label:ValidateExportNumber
'@prop-title Validate export number bounds
'@param exportNumber The number to check.
Private Sub ValidateExportNumber(ByVal exportNumber As Long)
    Dim total As Long
    total = NumberOfExports()

    If total = 0 Then
        ThrowError ProjectError.ElementNotFound, "No export specifications available"
    End If

    If (exportNumber < 1 Or exportNumber > total) Then
        ThrowError ProjectError.InvalidArgument, _
                  "Export number " & CStr(exportNumber) & " is out of range"
    End If
End Sub

'@label:ValidatePasswords
'@prop-title Validate password provider dependency
'@param pass The password object to check.
Private Sub ValidatePasswords(ByVal pass As Passwords)
    If pass Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Passwords object is required"
    End If
End Sub

'@label:ExportTable
'@prop-title Retrieve export table wrapper
'@details
'VBA evaluates both sides of Or, so the two guards below are written as two
'separate tests: reading Count off Nothing would raise on the very case the
'first test is there to catch.
'@return CustomTable. The wrapper over the first table on the export sheet.
Private Function ExportTable() As CustomTable
    Dim sheetTables As ListObjects
    Dim lo As ListObject

    EnsureWorksheetReady

    Set sheetTables = this.exportSheet.ListObjects
    If sheetTables Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "No table found on export worksheet"
    End If

    If sheetTables.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, "No table found on export worksheet"
    End If

    'The adapter of the previous call is done with, so whatever it filed is
    'taken now. Reading it at report time alone would keep the last one only.
    If Not (this.exportTableAdapter Is Nothing) Then
        If this.exportTableAdapter.HasCheckings Then _
            HarvestCollaboratorCheckings this.exportTableAdapter.CheckingValues
    End If

    Set lo = sheetTables(1)
    Set this.exportTableAdapter = CustomTable.Create(lo, COL_EXPORT_NUMBER)
    Set ExportTable = this.exportTableAdapter
End Function

'@section Specification Reads
'===============================================================================

'@label:LoadSpecification
'@sub-title Read the whole specification table once
'@details
'The header row and the data body are each read in one crossing and kept.
'Every later read of a cell is answered from those two arrays, so an export
'that reads eight columns costs two crossings instead of eight searches.
'The copy is dropped by `DropSpecification` whenever the class writes to the
'table, and by `ResetCaches` for a caller that wrote to the sheet itself.
Private Sub LoadSpecification()
    Dim headerRng As Range
    Dim bodyRng As Range

    If this.specLoaded Then Exit Sub

    EnsureDataReady

    Set headerRng = this.expData.HeaderRange
    Set bodyRng = this.expData.DataRange(includeHeaders:=False)

    If headerRng Is Nothing Then Exit Sub
    If bodyRng Is Nothing Then Exit Sub

    this.specHeaders = RowValues(headerRng)
    this.specBlock = BlockValues(bodyRng)
    this.specLoaded = True
End Sub

'@label:DropSpecification
'@sub-title Forget the copy of the specification table
Private Sub DropSpecification()
    this.specHeaders = Empty
    this.specBlock = Empty
    this.specLoaded = False
End Sub

'@label:SpecColumnIndex
'@sub-title Position of a column in the specification table
'@details
'Whole string, case-insensitive, leftmost wins. That is the same rule the
'worksheet search used before, applied to the header row held in memory.
'@param columnName Header to look for.
'@return Long. The 1-based position, zero when the header is missing.
Private Function SpecColumnIndex(ByVal columnName As String) As Long
    Dim idx As Long

    If Not this.specLoaded Then Exit Function

    For idx = 1 To UBound(this.specHeaders)
        If StrComp(Trim$(CStr(this.specHeaders(idx))), columnName, vbTextCompare) = 0 Then
            SpecColumnIndex = idx
            Exit Function
        End If
    Next idx
End Function

'@label:SpecRowCount
'@sub-title How many rows the copy of the specification table holds
'@return Long. The row count, zero when nothing is loaded.
Private Function SpecRowCount() As Long
    If Not this.specLoaded Then Exit Function
    SpecRowCount = UBound(this.specBlock, 1)
End Function

'@label:ReadValue
'@prop-title Read a column value for one export row
'@param columnName Header of the column to read.
'@param exportNumber Row to read, 1-based within the table body.
'@param required Optional; raise when the column or the row is missing.
'@return String. The cell value, empty when it cannot be read and required is False.
'@throws ProjectError.ElementNotFound when a required column is missing.
'@throws ProjectError.InvalidArgument when a required row is out of range.
Private Function ReadValue(ByVal columnName As String, _
                           ByVal exportNumber As Long, _
                           Optional ByVal required As Boolean = False) As String

    Dim columnIdx As Long

    LoadSpecification

    columnIdx = SpecColumnIndex(columnName)
    If columnIdx = 0 Then
        If required Then
            ThrowError ProjectError.ElementNotFound, _
                       "Export specification column '" & columnName & "' not found"
        End If
        Exit Function
    End If

    If exportNumber < 1 Or exportNumber > SpecRowCount() Then
        If required Then
            ThrowError ProjectError.InvalidArgument, _
                      "Export number " & CStr(exportNumber) & " is outside the specification range"
        End If
        Exit Function
    End If

    ReadValue = CStr(this.specBlock(exportNumber, columnIdx))
End Function

'@label:SpecColumnRange
'@sub-title Resolve one specification column as a worksheet Range
'@details
'`Translate` writes back into the cells, so it needs the Range itself rather
'than the values held in memory.
'@param columnName Header of the column.
'@return Range. The column body, Nothing when the header is missing.
Private Function SpecColumnRange(ByVal columnName As String) As Range

    On Error GoTo ColumnMissing
        Set SpecColumnRange = this.expData.DataRange(colName:=columnName, _
                                                     includeHeaders:=False, _
                                                     strictSearch:=True, _
                                                     matchCase:=False)
    On Error GoTo 0
    Exit Function

ColumnMissing:
    On Error GoTo 0
End Function

'@label:RowValues
'@sub-title Read a one-row Range as a 1-based list of values
'@param rowRange Range spanning one row.
'@return Variant. A 1-based one-dimensional array.
Private Function RowValues(ByVal rowRange As Range) As Variant
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long
    Dim nbColumns As Long

    nbColumns = rowRange.Columns.Count
    ReDim flat(1 To nbColumns)

    If nbColumns = 1 Then
        flat(1) = rowRange.Cells(1, 1).Value
    Else
        rawValues = rowRange.Value
        For counter = 1 To nbColumns
            flat(counter) = rawValues(1, counter)
        Next counter
    End If

    RowValues = flat
End Function

'@label:ColumnValues
'@sub-title Read a one-column Range as a 1-based list of values
'@param columnRange Range spanning one column.
'@return Variant. A 1-based one-dimensional array.
Private Function ColumnValues(ByVal columnRange As Range) As Variant
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long
    Dim nbRows As Long

    nbRows = columnRange.Rows.Count
    ReDim flat(1 To nbRows)

    If nbRows = 1 Then
        flat(1) = columnRange.Cells(1, 1).Value
    Else
        rawValues = columnRange.Value
        For counter = 1 To nbRows
            flat(counter) = rawValues(counter, 1)
        Next counter
    End If

    ColumnValues = flat
End Function

'@label:BlockValues
'@sub-title Read a Range as a 1-based two-dimensional array
'@details
'Range.Value hands back a plain value for one cell and a two-dimensional array
'otherwise. This gives callers one shape.
'@param blockRange Range to read.
'@return Variant. A 1-based two-dimensional array.
Private Function BlockValues(ByVal blockRange As Range) As Variant
    Dim block() As Variant

    If blockRange.Rows.Count = 1 And blockRange.Columns.Count = 1 Then
        ReDim block(1 To 1, 1 To 1)
        block(1, 1) = blockRange.Value
        BlockValues = block
    Else
        BlockValues = blockRange.Value
    End If
End Function

'@label:ResetCaches
'@prop-title Reset cached references and diagnostics
'@details
'Callers that write to the export sheet themselves call this so the next read
'goes back to the worksheet.
Public Sub ResetCaches()
    Set this.dictInt = Nothing
    Set this.variables = Nothing
    this.checkCounter = 0
    Set this.checks = Nothing
    this.hasChecks = False
    this.sanitizeList = Empty
    DropSpecification
End Sub


'@label:ComposeTemplatePath
'@prop-title Build the filename path from template chunks
'@param template The file name template read from the specification table.
'@return String. The resolved chunks joined by a double underscore.
Private Function ComposeTemplatePath(ByVal template As String) As String

    Dim pathParts As BetterArray
    Dim chunks As Variant
    Dim idx As Long
    Dim resolved As String
    Dim hostWorkbook As Workbook

    Set hostWorkbook = this.exportSheet.Parent

    chunks = Split(template, "+")
    Set pathParts = New BetterArray
    pathParts.LowerBound = LBound(chunks)

    For idx = LBound(chunks) To UBound(chunks)
        resolved = ResolveTemplateChunk(CStr(chunks(idx)), hostWorkbook)
        pathParts.Push resolved
    Next idx

    ComposeTemplatePath = pathParts.ToString(separator:="__", _
                                            OpeningDelimiter:=vbNullString, _
                                            ClosingDelimiter:=vbNullString, _
                                            QuoteStrings:=False)
End Function

'@label:ResolveTemplateChunk
'@prop-title Resolve a single template chunk into a safe string
'@param rawChunk One piece of the template.
'@param hostWorkbook Workbook holding the named ranges the chunk may point at.
'@return String. The resolved and sanitised value.
Private Function ResolveTemplateChunk(ByVal rawChunk As String, _
                                      ByVal hostWorkbook As Workbook) As String

    Dim chunk As String
    Dim value As String

    chunk = Trim$(rawChunk)
    If LenB(chunk) = 0 Then
        ResolveTemplateChunk = "chunk"
        Exit Function
    End If

    If IsLiteralTemplateChunk(chunk) Then
        ResolveTemplateChunk = SanitizeChunk(StripLiteralChunk(chunk))
        Exit Function
    End If

    value = ResolveVariableChunk(chunk, hostWorkbook)

    If LenB(value) = 0 Then
        LogInfo "Chunk resolved to empty string: " & chunk
        value = SanitizeChunk(chunk)
    End If

    ResolveTemplateChunk = value
End Function

'@label:IsLiteralTemplateChunk
'@prop-title Determine whether a chunk represents a literal string
'@param chunk The chunk text.
'@return Boolean. True when the chunk is wrapped in quotes.
Private Function IsLiteralTemplateChunk(ByVal chunk As String) As Boolean
    Dim firstChar As String
    Dim lastChar As String

    If Len(chunk) < 2 Then Exit Function

    firstChar = Left$(chunk, 1)
    lastChar = Right$(chunk, 1)

    If firstChar = Chr$(34) And lastChar = Chr$(34) Then
        IsLiteralTemplateChunk = True
        Exit Function
    End If

    If firstChar = "'" And lastChar = "'" Then
        IsLiteralTemplateChunk = True
        Exit Function
    End If
End Function

'@label:StripLiteralChunk
'@prop-title Remove wrapping quotes from literal chunk values
'@param chunk The chunk text.
'@return String. The chunk without its wrapping quotes.
Private Function StripLiteralChunk(ByVal chunk As String) As String
    Dim lengthValue As Long

    lengthValue = Len(chunk)
    If lengthValue >= 2 Then
        If (Left$(chunk, 1) = Chr$(34) And Right$(chunk, 1) = Chr$(34)) _
           Or (Left$(chunk, 1) = "'" And Right$(chunk, 1) = "'") Then
            StripLiteralChunk = Mid$(chunk, 2, lengthValue - 2)
            Exit Function
        End If
    End If

    StripLiteralChunk = chunk
End Function

'@label:ResolveVariableChunk
'@prop-title Resolve dictionary-driven chunk value
'@param chunk Name of the variable the chunk points at.
'@param hostWorkbook Workbook holding the named range.
'@return String. The sanitised value, empty when the chunk cannot be resolved.
Private Function ResolveVariableChunk(ByVal chunk As String, _
                                      ByVal hostWorkbook As Workbook) As String

    Dim vars As LLVariables
    Dim sheetName As String
    Dim sheetType As String
    Dim targetSheet As Worksheet
    Dim namedRange As Range
    Dim resolved As String

    EnsureVariablesReady
    Set vars = this.variables

    If vars Is Nothing Then Exit Function
    If Not vars.Contains(chunk) Then Exit Function

    sheetName = vars.SheetName(chunk)
    sheetType = vars.Value("sheet type", chunk)

    If LenB(sheetName) = 0 Then Exit Function
    If StrComp(sheetType, "vlist1D", vbTextCompare) <> 0 Then Exit Function

    Set targetSheet = TryGetWorksheet(hostWorkbook, sheetName)
    If targetSheet Is Nothing Then
        LogInfo "Worksheet '" & sheetName & "' for chunk '" & chunk & "' is missing"
        Exit Function
    End If

    Set namedRange = TryGetNamedRange(hostWorkbook, chunk, targetSheet)

    If namedRange Is Nothing Then
        LogInfo "Named range for chunk '" & chunk & "' is missing", checkingWarning
        Exit Function
    End If

    resolved = CStr(namedRange.Value)
    resolved = SanitizeChunk(resolved)

    ResolveVariableChunk = resolved
End Function

'@label:TryGetWorksheet
'@prop-title Safely retrieve a worksheet by name
'@param hostWorkbook Workbook to look in.
'@param sheetName Name of the worksheet.
'@return Worksheet. Nothing when there is no such sheet.
Private Function TryGetWorksheet(ByVal hostWorkbook As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set TryGetWorksheet = hostWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

'@label:TryGetNamedRange
'@prop-title Safely retrieve a named range for a chunk
'@param hostWorkbook Workbook holding the workbook-level names.
'@param chunk Name to look for.
'@param targetSheet Worksheet to fall back on for a sheet-level name.
'@return Range. Nothing when there is no such name.
Private Function TryGetNamedRange(ByVal hostWorkbook As Workbook, _
                                  ByVal chunk As String, _
                                  ByVal targetSheet As Worksheet) As Range

    On Error Resume Next
        Set TryGetNamedRange = hostWorkbook.Names(chunk).RefersToRange
    On Error GoTo 0

    If TryGetNamedRange Is Nothing Then
        On Error Resume Next
            Set TryGetNamedRange = targetSheet.Range(chunk)
        On Error GoTo 0
    End If
End Function

'@label:SanitizeChunk
'@prop-title Sanitise individual chunk values for filenames
'@param chunk The value to clean.
'@return String. The cleaned value, "chunk" when nothing is left.
Private Function SanitizeChunk(ByVal chunk As String) As String
    Dim sanitized As String
    sanitized = Trim$(chunk)
    sanitized = ReplaceInvalidCharacters(sanitized)
    If LenB(sanitized) = 0 Then sanitized = "chunk"
    SanitizeChunk = sanitized
End Function

'@label:ReplaceInvalidCharacters
'@prop-title Replace characters unsupported in filenames
'@details
'Each pass works on the result of the one before it, so every character in the
'list is replaced. The value goes straight into a SaveAs path.
'@param value The text to clean.
'@return String. The text with every listed character turned into an underscore.
Private Function ReplaceInvalidCharacters(ByVal value As String) As String
    Dim invalidChar As Variant

    EnsureSanitizeList

    For Each invalidChar In this.sanitizeList
        value = Replace(value, CStr(invalidChar), "_")
    Next invalidChar

    ReplaceInvalidCharacters = value
End Function

'@label:EnsureSanitizeList
'@prop-title Initialise invalid filename character list
Private Sub EnsureSanitizeList()
    If IsEmpty(this.sanitizeList) Then
        this.sanitizeList = Array("<", ">", ":", "\", "/", "|", "?", "*", _
                                  Chr$(34), "&", ".", " ")
    End If
End Sub

'@label:BuildVersionSuffix
'@prop-title Compose version/public key suffix for filenames
'@param pass Password provider holding the version and the public key.
'@return String. The suffix to append to the file name.
Private Function BuildVersionSuffix(ByVal pass As Passwords) As String
    Dim versionValue As String
    Dim publicKeyValue As String

    versionValue = SanitizeChunk(pass.Value("version"))
    publicKeyValue = SanitizeChunk(pass.Value("publickey"))

    BuildVersionSuffix = "__v" & versionValue & "-" & publicKeyValue
End Function

'@label:BuildTimestamp
'@prop-title Compose timestamp fragment for filenames
'@return String. The current date and time in the file name format.
Private Function BuildTimestamp() As String
    BuildTimestamp = Format$(Now, TIMESTAMP_FORMAT)
End Function

'@label:DefaultTemplate
'@prop-title Provide fallback filename template
'@param exportNumber The export the template belongs to.
'@return String. The fallback template.
Private Function DefaultTemplate(ByVal exportNumber As Long) As String
    DefaultTemplate = "export" & CStr(exportNumber)
End Function


'@label:WorkbookBaseName
'@prop-title Remove Excel extensions from workbook name
'@param hostWorkbook The workbook to name.
'@return String. The lower-cased name without its extension.
Private Function WorkbookBaseName(ByVal hostWorkbook As Workbook) As String
    Dim cleaned As String

    cleaned = LCase$(hostWorkbook.Name)
    cleaned = Replace(cleaned, ".xlsm", vbNullString)
    cleaned = Replace(cleaned, ".xlsb", vbNullString)
    cleaned = Replace(cleaned, ".xlsx", vbNullString)
    cleaned = Replace(cleaned, ".xls", vbNullString)

    WorkbookBaseName = cleaned
End Function

'@label:LogInfo
'@prop-title Append informational diagnostics to the Checking log
'@param label The message to record.
'@param scope Optional checking scope. Defaults to checkingInfo.
Private Sub LogInfo(ByVal label As String, Optional ByVal scope As Byte = checkingInfo)
    If this.checks Is Nothing Then Set this.checks = Checking.Create(CHECKING_CONTEXT)
    this.checks.Add CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@section Error Handling
'===============================================================================
'@label:ThrowError
'@prop-title Raise a ProjectError coded runtime error
'@param errNumb The ProjectError value to raise.
'@param errorMessage The message to carry with it.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub
