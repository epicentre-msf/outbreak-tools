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
'@Folder("General")
'@ModuleDescription("Manage linelist and setup exports")

'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, ProcedureNotUsed, SuperfluousAnnotationArgument


'@depends Checking, DataSheet, LLVariables, HiddenNames
Option Explicit

Implements ILLExport

Private Type TLLExport
    expData As IDataSheet
    exportSheet As Worksheet
    dictInt As ILLdictionary
    variables As ILLVariables
    checks As IChecking
    hasChecks As Boolean
    checkCounter As Long
    sanitizeList As Variant
    store As IHiddenNames
End Type

Private Const COL_STATUS As String = "status"
Private Const COL_FILE_NAME As String = "file name"
Private Const COL_EXPORT_NUMBER As String = "export number"
Private Const COL_INCLUDE_IDS As String = "include personal identifiers"
Private Const STATUS_ACTIVE As String = "active"
Private Const DEFAULT_INCLUDE_IDS As String = "no"
Private Const TIMESTAMP_FORMAT As String = "yyyymmdd-HhNn"
Private Const CHECKING_CONTEXT As String = "LLExport checks"
Private Const TOTAL_EXPORTS_NAME As String = "__ll_exports_total__"
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
'@return A fully initialised `ILLExport` instance.
'@throws ProjectError.ObjectNotInitialized when the worksheet is missing.
Public Function Create(ByVal exportSheet As Worksheet, _
                       Optional ByVal startRow As Long = 1, _
                       Optional ByVal startColumn As Long = 1) As ILLExport

    Dim exportData As IDataSheet

    If exportSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Worksheet is required to create " & CLASSNAME
    End If

    Set exportData = DataSheet.Create(exportSheet, startRow, startColumn, True, "ExportsDataSheet")

    With New LLExport
        Set .InternalData = exportData
        Set .InternalWorksheet = exportSheet
        .ResetCaches
        Set Create = .Self()
    End With
End Function

'@label:SetInternalData
'@param exportData Datasheet wrapper over the specification table.
Public Property Set InternalData(ByVal exportData As IDataSheet)
    Set this.expData = exportData
End Property

'@label:SetInternalWorksheet
'@param exportSheet Worksheet that stores the export specification table.
Public Property Set InternalWorksheet(ByVal exportSheet As Worksheet)
    Set this.exportSheet = exportSheet
    Set this.store = Nothing
End Property

'@label:Self
'@prop-title Return the current instance as interface
'@return The object cast as `ILLExport`.
Public Property Get Self() As ILLExport
    Set Self = Me
End Property

'@section Properties
'===============================================================================
'@label:Data
'@prop-title Underlying export datasheet
'@return Datasheet abstraction wrapping the export specification table.
Public Property Get Data() As IDataSheet
    EnsureDataReady
    Set Data = this.expData
End Property

'@label:Worksheet
'@prop-title Worksheet hosting the export specification table
'@return Worksheet reference used by the handler.
Public Property Get Worksheet() As Worksheet
    EnsureWorksheetReady
    Set Worksheet = this.exportSheet
End Property

'@label:NumberOfExports
'@prop-title Count the configured export rows
'@details Counts rows within the specification table excluding headers.
'@return Total number of export definitions currently present.
Public Property Get NumberOfExports() As Long
    EnsureDataReady
    EnsureWorksheetReady
    NumberOfExports = StoredTotalExports()
    If NumberOfExports <= 0 Then
        NumberOfExports = CountExportRows()
        PersistTotalExports NumberOfExports
    End If
End Property

'@label:HasCheckings
'@prop-title Whether informational checkings were captured
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:CheckingValues
'@prop-title Collected informational diagnostics
'@return Checking object containing log entries, if any.
Public Property Get CheckingValues() As IChecking
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
'logic operates on identifiers rather than simple counts so arbitrary deletions
'do not remove unrelated export columns.
'@param dict Optional dictionary used to add/remove export columns.
'@param del Optional flag switching to row removal mode.
'@param rowCount Optional threshold for row removal; rows with non-empty cells at or below the threshold are trimmed (defaults to 0, removing empty rows).
Private Sub ManageRows(Optional ByVal del As Boolean = False, _ 
                       Optional ByVal dict As ILLdictionary, _
                       Optional ByVal rowCount As Long = 2)

    Dim tableObj As ICustomTable
    Dim previousCount As Long
    Dim newCount As Long
    Dim exportNumbers As BetterArray
    Dim dictionaryReady As Boolean
    Dim dictRef As ILLdictionary

    
    EnsureDataReady
    EnsureWorksheetReady
    EnsureDictionaryReady dict, Not dict Is Nothing, False
    dictionaryReady = DictionaryIsAvailable()
    If dictionaryReady Then Set dictRef = this.dictInt
   
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

    EnsureExportIdentifiers tableObj

    newCount = CountExportRows()
    PersistTotalExports newCount

    Set exportNumbers = CollectExportNumbers(tableObj)
    If dictionaryReady Then SyncDictionaryExportsCore exportNumbers, dictRef
End Sub

'@label:InsertRows
'@prop-title Insert export rows and keep dictionary columns aligned
'@param targetCell Range anchor representing where to mirror inserted rows.
'@param dict Optional dictionary to shift export columns.
'@param insertShift Optional Boolean to insert worksheet rows when necessary. Defaults to False.
Public Sub InsertRows(ByVal targetCell As Range, _
                      Optional ByVal dict As ILLdictionary, _
                      Optional ByVal insertShift As Boolean = False)

    Dim tableObj As ICustomTable
    Dim exportNumbers As BetterArray
    Dim dictionaryReady As Boolean
    Dim dictRef As ILLdictionary
    Dim insertIndex As Long
    Dim insertCount As Long
    Dim actualInserted As Long
    Dim previousCount As Long
    Dim newCount As Long

    If (targetCell Is Nothing) Then Exit Sub

    EnsureDataReady
    EnsureWorksheetReady

    EnsureDictionaryReady dict, (Not dict Is Nothing), False

    dictionaryReady = DictionaryIsAvailable()
    If dictionaryReady Then Set dictRef = this.dictInt

    Set tableObj = ExportTable()
    insertCount = targetCell.Rows.Count
    previousCount = CountExportRows()

    insertIndex = ResolveSelectionRowIndex(tableObj, targetCell)
    If insertIndex <= 0 Then insertIndex = previousCount + 1
    If insertIndex > previousCount + 1 Then insertIndex = previousCount + 1

    tableObj.InsertRowsAt targetCell:=targetCell, insertShift:=insertShift, includeIds:=False

    newCount = CountExportRows()
    actualInserted = newCount - previousCount

    If actualInserted <= 0 Then
        PersistTotalExports newCount
        Exit Sub
    End If

    ApplyIncludeDefaults tableObj, insertIndex, Application.WorksheetFunction.Min(newCount, insertIndex + actualInserted - 1)
    AssignExportNumbers tableObj, insertIndex, actualInserted

    PersistTotalExports newCount

    Set exportNumbers = CollectExportNumbers(tableObj)
    If dictionaryReady Then SyncDictionaryExportsCore exportNumbers, dictRef
End Sub

'@label:DeleteRows
'@prop-title Delete export rows intersecting the supplied selection
Public Sub DeleteRows(ByVal targetCell As Range, _
                      Optional ByVal dict As ILLdictionary, _
                      Optional ByVal forceShift As Boolean = False)

    Dim tableObj As ICustomTable
    Dim exportNumbers As BetterArray
    Dim dictionaryReady As Boolean
    Dim dictRef As ILLdictionary
    Dim previousCount As Long
    Dim newCount As Long

    If targetCell Is Nothing Then Exit Sub

    EnsureDataReady
    EnsureWorksheetReady

    EnsureDictionaryReady dict, Not dict Is Nothing, False
    dictionaryReady = DictionaryIsAvailable()
    If dictionaryReady Then Set dictRef = this.dictInt

    Set tableObj = ExportTable()
    previousCount = CountExportRows()

    tableObj.DeleteRowsAt targetCell:=targetCell, _
                          includeIds:=False, _
                          forceShift:=forceShift

    newCount = CountExportRows()
    If newCount = previousCount Then
        PersistTotalExports newCount
        Exit Sub
    End If

    EnsureExportIdentifiers tableObj
    PersistTotalExports newCount

    Set exportNumbers = CollectExportNumbers(tableObj)
    If dictionaryReady Then SyncDictionaryExportsCore exportNumbers, dictRef
End Sub

'@label:AddRows
'@prop-title Append export definition rows
'@param dict Optional dictionary to align export columns with.
Public Sub AddRows(Optional ByVal dict As ILLdictionary)
    ManageRows dict:=dict
End Sub

'@label:RemoveRows
'@prop-title Remove empty export rows
'@param rowCount Optional threshold; rows with non-empty cell counts at or below this value are removed (defaults to 0).
Public Sub RemoveRows(Optional ByVal dict As ILLdictionary, _ 
                      Optional ByVal rowCount As Long = 2)
    ManageRows del:=True, dict:=dict, rowCount:=rowCount
End Sub

'@label:SyncDictionaryExports
'@prop-title Synchronise dictionary export columns
Public Sub SyncDictionaryExports(Optional ByVal dict As ILLdictionary)
    Dim tableObj As ICustomTable
    Dim exportNumbers As BetterArray
    Dim targetDict As ILLdictionary

    EnsureDataReady
    EnsureWorksheetReady

    Set tableObj = ExportTable()
    Set exportNumbers = CollectExportNumbers(tableObj)

    EnsureDictionaryReady dict, Not dict Is Nothing, False

    If Not (this.dictInt Is Nothing) Then
        Set targetDict = this.dictInt
    Else
        Set targetDict = dict
    End If

    SyncDictionaryExportsCore exportNumbers, targetDict
End Sub

'@label:Sort
'@sub-title Sort export definitions and normalize identifiers
Public Sub Sort(Optional ByVal dict As ILLdictionary)
    Dim tableObj As ICustomTable
    Dim dictRef As ILLdictionary

    EnsureDataReady
    EnsureWorksheetReady

    Set tableObj = ExportTable()

    On Error Resume Next
        tableObj.Sort colName:=COL_EXPORT_NUMBER
    On Error GoTo 0

    EnsureDictionaryReady dict, Not dict Is Nothing, False
    If Not (this.dictInt Is Nothing) Then
        Set dictRef = this.dictInt
    Else
        Set dictRef = dict
    End If

    SequentiallyRenameExports tableObj, dictRef
End Sub

'@section Row Coordination Helpers
'===============================================================================

Private Sub ApplyIncludeDefaults(ByVal tableObj As ICustomTable, _
                                 ByVal startIndex As Long, _
                                 ByVal endIndex As Long)
    Dim includeColumn As Range
    Dim idx As Long

    If startIndex <= 0 Or endIndex < startIndex Then Exit Sub

    Set includeColumn = tableObj.DataRange(COL_INCLUDE_IDS)
    If includeColumn Is Nothing Then Exit Sub

    For idx = startIndex To endIndex
        If idx > includeColumn.Rows.Count Then Exit For
        If LenB(CStr(includeColumn.Cells(idx, 1).Value)) = 0 Then
            includeColumn.Cells(idx, 1).Value = DEFAULT_INCLUDE_IDS
            includeColumn.Cells(idx, 1).Locked = False
        End If
    Next idx
End Sub

Private Function ResolveSelectionRowIndex(ByVal tableObj As ICustomTable, _
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

Private Sub AssignExportNumbers(ByVal tableObj As ICustomTable, _
                                ByVal startIndex As Long, _
                                ByVal insertCount As Long)

    Dim exportRange As Range
    Dim existing As BetterArray
    Dim idx As Long
    Dim nextNumber As Long

    If insertCount <= 0 Then Exit Sub

    Set exportRange = tableObj.DataRange(COL_EXPORT_NUMBER)
    If exportRange Is Nothing Then Exit Sub

    Set existing = CollectExportNumbers(tableObj)
    For idx = 0 To insertCount - 1
        nextNumber = NextAvailableExportNumber(existing)
        exportRange.Cells(startIndex + idx, 1).Value = ExportColumnName(nextNumber)
        If existing Is Nothing Then
            Set existing = New BetterArray
            existing.LowerBound = 1
        End If
        existing.Push nextNumber
    Next idx
End Sub

Private Sub SequentiallyRenameExports(ByVal tableObj As ICustomTable, _
                                      Optional ByVal dict As ILLdictionary)

    Dim exportRange As Range
    Dim total As Long
    Dim idx As Long
    Dim previousNames() As String

    Set exportRange = tableObj.DataRange(COL_EXPORT_NUMBER)
    If exportRange Is Nothing Then Exit Sub

    total = exportRange.Rows.Count
    If total = 0 Then Exit Sub

    ReDim previousNames(1 To total)
    For idx = 1 To total
        previousNames(idx) = CStr(exportRange.Cells(idx, 1).Value)
        exportRange.Cells(idx, 1).Value = ExportColumnName(idx)
    Next idx

    PersistTotalExports total

    If( Not dict Is Nothing) Then
        RenameDictionaryExports dict, previousNames, total
    End If
End Sub

Private Sub RenameDictionaryExports(ByVal dict As ILLdictionary, _
                                    ByRef previousNames() As String, _
                                    ByVal total As Long)

    Dim idx As Long
    Dim tempName As String
    Dim normalized As String

    If dict Is Nothing Then Exit Sub
    If total <= 0 Then Exit Sub

    'First pass: rename old export columns to temporary placeholders.
    For idx = 1 To total
        normalized = NormalizedExportHeader(previousNames(idx))
        If LenB(normalized) = 0 Then
            GoTo ContinueFor
        End If
        If dict.ColumnExists(normalized) Then
            tempName = TemporaryExportColumn(idx)
            dict.RenameColumn normalized, tempName, strictSearch:=True, matchCase:=False
        End If
    
ContinueFor:
    Next idx

    'Second pass: rename placeholders to sequential export names or create columns when missing.
    For idx = 1 To total
        tempName = TemporaryExportColumn(idx)
        If dict.ColumnExists(tempName) Then
            dict.RenameColumn tempName, ExportColumnName(idx), strictSearch:=True, matchCase:=True
        Else
            EnsureDictionaryExportColumn dict, idx
        End If
    Next idx

    'Remove extra export columns beyond the active count.
    Dim existing As BetterArray
    Dim pointer As Long

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

Private Function NextAvailableExportNumber(ByVal existing As BetterArray) As Long
    Dim candidate As Long

    candidate = Application.WorksheetFunction.Max(StoredTotalExports(), HighestNumber(existing))

    Do
        candidate = candidate + 1
    Loop While ContainsNumber(existing, candidate)

    NextAvailableExportNumber = candidate
End Function

Private Function HighestNumber(ByVal numbers As BetterArray) As Long
    Dim idx As Long
    Dim current As Long

    If numbers Is Nothing Then Exit Function
    For idx = numbers.LowerBound To numbers.UpperBound
        current = CLng(numbers.Item(idx))
        If current > HighestNumber Then HighestNumber = current
    Next idx
End Function

Private Function TemporaryExportColumn(ByVal orderIndex As Long) As String
    TemporaryExportColumn = "__exp_tmp_" & Format$(orderIndex, "0000")
End Function

Private Function NormalizedExportHeader(ByVal headerText As String) As String
    Dim parsed As Long

    parsed = ParseExportIdentifier(headerText)
    If parsed <= 0 Then Exit Function

    NormalizedExportHeader = ExportColumnName(parsed)
End Function

Private Sub EnsureExportIdentifiers(ByVal tableObj As ICustomTable)
    Dim exportRange As Range
    Dim usedNumbers As BetterArray
    Dim pendingRows As BetterArray
    Dim idx As Long
    Dim parsed As Long
    Dim maxNumber As Long
    Dim rowIndex As Long

    Set exportRange = tableObj.DataRange(COL_EXPORT_NUMBER)
    If exportRange Is Nothing Then Exit Sub

    Set usedNumbers = New BetterArray
    usedNumbers.LowerBound = 1
    Set pendingRows = New BetterArray
    pendingRows.LowerBound = 1
    maxNumber = 0

    For idx = 1 To exportRange.Rows.Count
        parsed = ParseExportIdentifier(exportRange.Cells(idx, 1).Value)
        If parsed > 0 Then
            If ContainsNumber(usedNumbers, parsed) Then
                exportRange.Cells(idx, 1).Value = vbNullString
                pendingRows.Push idx
            Else
                usedNumbers.Push parsed
                exportRange.Cells(idx, 1).Value = ExportColumnName(parsed)
                If parsed > maxNumber Then maxNumber = parsed
            End If
        Else
            exportRange.Cells(idx, 1).Value = vbNullString
            pendingRows.Push idx
        End If
    Next idx

    If pendingRows.Length > 0 Then
        For idx = pendingRows.LowerBound To pendingRows.UpperBound
            rowIndex = CLng(pendingRows.Item(idx))
            maxNumber = maxNumber + 1
            exportRange.Cells(rowIndex, 1).Value = ExportColumnName(maxNumber)
        Next idx
    End If
End Sub

Private Function CollectExportNumbers(ByVal tableObj As ICustomTable) As BetterArray
    Dim exportRange As Range
    Dim result As BetterArray
    Dim idx As Long
    Dim parsed As Long

    Set exportRange = tableObj.DataRange(COL_EXPORT_NUMBER)
    If exportRange Is Nothing Then Exit Function

    Set result = New BetterArray
    result.LowerBound = 1

    For idx = 1 To exportRange.Rows.Count
        parsed = ParseExportIdentifier(exportRange.Cells(idx, 1).Value)
        If parsed > 0 Then
            If Not ContainsNumber(result, parsed) Then
                result.Push parsed
            End If
        End If
    Next idx

    If result.Length = 0 Then
        Set CollectExportNumbers = Nothing
    Else
        Set CollectExportNumbers = result
    End If
End Function

Private Function EnumerateDictionaryExportNumbers(ByVal dict As ILLdictionary) As BetterArray
    Dim headers As Range
    Dim result As BetterArray
    Dim idx As Long
    Dim parsed As Long
    If dict Is Nothing Then Exit Function

    Set headers = dict.DataRange(includeHeaders:=True)
    If headers Is Nothing Then Exit Function

    Set result = New BetterArray
    result.LowerBound = 1

    For idx = 1 To headers.Columns.Count
        parsed = ParseExportIdentifier(headers.Cells(1, idx).Value)
        If parsed > 0 Then
            If Not ContainsNumber(result, parsed) Then
                result.Push parsed
            End If
        End If
    Next idx

    If result.Length = 0 Then
        Set EnumerateDictionaryExportNumbers = Nothing
    Else
        Set EnumerateDictionaryExportNumbers = result
    End If
End Function

Private Function ContainsNumber(ByVal numbers As BetterArray, ByVal target As Long) As Boolean
    If numbers Is Nothing Then Exit Function
    If numbers.Length = 0 Then Exit Function
    ContainsNumber = numbers.Includes(target)
End Function

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

Private Sub SyncDictionaryExportsCore(ByVal exportNumbers As BetterArray, ByVal dict As ILLdictionary)
    Dim existing As BetterArray
    Dim idx As Long
    Dim number As Long
    Dim highest As Long

    If dict Is Nothing Then Exit Sub

    Set existing = EnumerateDictionaryExportNumbers(dict)

    If Not exportNumbers Is Nothing Then
        For idx = exportNumbers.LowerBound To exportNumbers.UpperBound
            number = CLng(exportNumbers.Item(idx))
            EnsureDictionaryExportColumn dict, number
            If number > highest Then highest = number
        Next idx
    End If

    If Not existing Is Nothing Then
        For idx = existing.LowerBound To existing.UpperBound
            number = CLng(existing.Item(idx))
            If Not ContainsNumber(exportNumbers, number) Then
                dict.RemoveColumn ExportColumnName(number)
            ElseIf number > highest Then
                highest = number
            End If
        Next idx
    End If
    dict.TotalNumberOfExports = highest
End Sub

Private Sub EnsureDictionaryExportColumn(ByVal dict As ILLdictionary, ByVal exportIndex As Long)
    Dim columnName As String
    Dim anchorName As String
    Dim previousIdx As Long
    Dim existing As BetterArray

    If dict Is Nothing Then Exit Sub

    columnName = ExportColumnName(exportIndex)
    If dict.ColumnExists(columnName) Then Exit Sub

    Set existing = EnumerateDictionaryExportNumbers(dict)

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
    
    Dim sourceData As IDataSheet
    
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

    Dim sourceData As IDataSheet
    
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
                               ByVal dict As ILLdictionary, _
                               ByVal pass As IPasswords, _
                               Optional ByVal exportAll As Boolean = False) As String

    Dim currWb As Workbook
    Dim template As String
    Dim sanitizedPath As String
    Dim suffix As String

    EnsureDataReady
    EnsureWorksheetReady
    ValidatePasswords pass
    EnsureDictionaryReady dict, True

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

'@section Export Count Tracking
'===============================================================================

'@sub-title Resolve worksheet name manager
Private Function HiddenNameStore() As IHiddenNames
    If this.exportSheet Is Nothing Then Exit Function

    If this.store Is Nothing Then
        Set this.store = HiddenNames.Create(this.exportSheet)
    End If

    Set HiddenNameStore = this.store
End Function

Private Function CountExportRows() As Long
    Dim exportRange As Range

    If this.expData Is Nothing Then Exit Function

    Set exportRange = this.expData.DataRange(includeHeaders:=False)
    If Not exportRange Is Nothing Then
        CountExportRows = exportRange.Rows.Count
    End If
End Function

Private Function StoredTotalExports() As Long
    Dim store As IHiddenNames

    Set store = HiddenNameStore()
    If store Is Nothing Then Exit Function

    StoredTotalExports = store.ValueAsLong(TOTAL_EXPORTS_NAME, 0)
End Function

Private Sub PersistTotalExports(ByVal totalCount As Long)
    Dim store As IHiddenNames

    If totalCount < 0 Then totalCount = 0

    Set store = HiddenNameStore()
    If store Is Nothing Then Exit Sub

    store.EnsureName TOTAL_EXPORTS_NAME, totalCount, HiddenNameTypeLong
    store.SetValue TOTAL_EXPORTS_NAME, totalCount
End Sub

Private Sub EnsureStoredTotalExports()
    Dim store As IHiddenNames
    Dim total As Long

    If this.exportSheet Is Nothing Then Exit Sub
    If this.expData Is Nothing Then Exit Sub

    total = CountExportRows()

    Set store = HiddenNameStore()
    If store Is Nothing Then Exit Sub

    store.EnsureName TOTAL_EXPORTS_NAME, total, HiddenNameTypeLong
    store.SetValue TOTAL_EXPORTS_NAME, total
End Sub


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

    Set result = New BetterArray
    result.LowerBound = 1

    For idx = 1 To total
        If IsActive(idx) Then result.Push idx
    Next idx

    Set ActiveExportNumbers = result
End Function

'@label:Translate
'@sub-title Translate the current LLExport
Private Sub Translate(ByVal trads As ITranslationObject)
    
    If trads Is Nothing Then 
        ThrowError ProjectError.ObjectNotInitialized, "Translations are not initialised"
    End If

    trads.TranslateRange Data.DataRange("label button")
    trads.TranslateRange Data.DataRange("file name"), containsFormula:=True

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

'@label:EnsureDictionaryReady
'@prop-title Validate dictionary dependency
Private Sub EnsureDictionaryReady(Optional ByVal dict As ILLdictionary, Optional ByVal resetInt As Boolean = False, Optional ByVal required As Boolean = True)
    If resetInt Then Set this.dictInt = Nothing

    If Not dict Is Nothing Then
        Set this.dictInt = dict
    End If

    If required And (this.dictInt Is Nothing) Then
        ThrowError ProjectError.ObjectNotInitialized, "LLExport Internal dictionary is required; none found"
    End If
End Sub

'@label:EnsureVariablesReady
'@prop-title Lazily instantiate dictionary-backed variable helper
Private Sub EnsureVariablesReady(Optional ByVal dict As ILLdictionary)
    EnsureDictionaryReady dict

    If Not DictionaryIsAvailable() Then Exit Sub

    Dim currentDict As ILLdictionary
    Set currentDict = InternalDictionary()

    If this.variables Is Nothing Then
        Set this.variables = LLVariables.Create(currentDict)
    Else
        Dim existingDict As ILLdictionary
        Set existingDict = this.variables.Dictionary
        If (existingDict Is Nothing) Or (Not (existingDict Is currentDict)) Then
            Set this.variables = LLVariables.Create(currentDict)
        End If
    End If
End Sub

'@label:InternalDictionary
Private Property Get InternalDictionary() As ILLdictionary
    EnsureDictionaryReady
    Set InternalDictionary = this.dictInt
End Property

'@label:InternalDictionary
Private Property Set InternalDictionary(ByVal dict As ILLdictionary)
    EnsureDictionaryReady dict
End Property

Private Function DictionaryIsAvailable() As Boolean
    DictionaryIsAvailable = Not (this.dictInt Is Nothing)
End Function

'@label:ValidateExportNumber
'@prop-title Validate export number bounds
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
Private Sub ValidatePasswords(ByVal pass As IPasswords)
    If pass Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Passwords object is required"
    End If
End Sub

'@label:ExportTable
'@prop-title Retrieve export table wrapper
Private Function ExportTable() As ICustomTable
    Dim listObjects As ListObjects
    Dim lo As ListObject

    EnsureWorksheetReady

    Set listObjects = this.exportSheet.ListObjects
    If listObjects Is Nothing Or listObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, "No table found on export worksheet"
    End If

    Set lo = listObjects(1)
    Set ExportTable = CustomTable.Create(lo, COL_EXPORT_NUMBER)
End Function

'@label:ReadValue
'@prop-title Read a column value from the cached datasheet
Private Function ReadValue(ByVal columnName As String, _
                           ByVal exportNumber As Long, _
                           Optional ByVal required As Boolean = False) As String

    Dim columnRange As Range

    Set columnRange = ResolveColumnRange(columnName, required)

    If columnRange Is Nothing Then Exit Function

    If exportNumber < 1 Or exportNumber > columnRange.Rows.Count Then
        If required Then
            ThrowError ProjectError.InvalidArgument, _
                      "Export number " & CStr(exportNumber) & " is outside the specification range"
        Else
            Exit Function
        End If
    End If

    ReadValue = CStr(columnRange.Cells(exportNumber, 1).Value)
End Function

'@label:ResolveColumnRange
'@prop-title Resolve and cache a specification column
Private Function ResolveColumnRange(ByVal columnName As String, _
                                    ByVal required As Boolean) As Range

    Dim resolved As Range

    On Error GoTo ColumnMissing
        Set resolved = this.expData.DataRange(colName:=columnName, _
                                             includeHeaders:=False, _
                                             strictSearch:=True, _
                                             matchCase:=False)
    On Error GoTo 0

    Set ResolveColumnRange = resolved
    Exit Function

ColumnMissing:
    On Error GoTo 0
    If required Then
        ThrowError ProjectError.ElementNotFound, _
                   "Export specification column '" & columnName & "' not found"
    End If
End Function

'@label:ResetCaches
'@prop-title Reset cached references and diagnostics
Public Sub ResetCaches()
    Set this.dictInt = Nothing
    Set this.variables = Nothing
    this.checkCounter = 0
    Set this.checks = Nothing
    this.sanitizeList = Empty
    Set this.store = Nothing
    EnsureStoredTotalExports
End Sub


'@label:ComposeTemplatePath
'@prop-title Build the filename path from template chunks
Private Function ComposeTemplatePath(ByVal template As String) As String

    Dim pathParts As BetterArray
    Dim chunks As Variant
    Dim idx As Long
    Dim resolved As String
    Dim workbook As Workbook
    
    Set workbook = this.exportSheet.Parent

    chunks = Split(template, "+")
    Set pathParts = New BetterArray
    pathParts.LowerBound = LBound(chunks)

    For idx = LBound(chunks) To UBound(chunks)
        resolved = ResolveTemplateChunk(CStr(chunks(idx)), workbook)
        pathParts.Push resolved
    Next idx

    ComposeTemplatePath = pathParts.ToString(separator:="__", _
                                            OpeningDelimiter:=vbNullString, _
                                            ClosingDelimiter:=vbNullString, _
                                            QuoteStrings:=False)
End Function

'@label:ResolveTemplateChunk
'@prop-title Resolve a single template chunk into a safe string
Private Function ResolveTemplateChunk(ByVal rawChunk As String, _
                                      ByVal workbook As Workbook) As String

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

    value = ResolveVariableChunk(chunk, workbook)

    If LenB(value) = 0 Then
        LogInfo "Chunk resolved to empty string: " & chunk
        value = SanitizeChunk(chunk)
    End If

    ResolveTemplateChunk = value
End Function

'@label:IsLiteralTemplateChunk
'@prop-title Determine whether a chunk represents a literal string
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
Private Function ResolveVariableChunk(ByVal chunk As String, _
                                       ByVal workbook As Workbook) As String

    Dim vars As ILLVariables
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

    Set targetSheet = TryGetWorksheet(workbook, sheetName)
    If targetSheet Is Nothing Then
        LogInfo "Worksheet '" & sheetName & "' for chunk '" & chunk & "' is missing"
        Exit Function
    End If

    Set namedRange = TryGetNamedRange(workbook, chunk, targetSheet)

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
Private Function TryGetWorksheet(ByVal workbook As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set TryGetWorksheet = workbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

'@label:TryGetNamedRange
'@prop-title Safely retrieve a named range for a chunk
Private Function TryGetNamedRange(ByVal workbook As Workbook, _
                                  ByVal chunk As String, _
                                  ByVal targetSheet As Worksheet) As Range

    On Error Resume Next
        Set TryGetNamedRange = workbook.Names(chunk).RefersToRange
    On Error GoTo 0

    If TryGetNamedRange Is Nothing Then
        On Error Resume Next
            Set TryGetNamedRange = targetSheet.Range(chunk)
        On Error GoTo 0
    End If
End Function

'@label:SanitizeChunk
'@prop-title Sanitise individual chunk values for filenames
Private Function SanitizeChunk(ByVal chunk As String) As String
    Dim sanitized As String
    sanitized = Trim$(chunk)
    sanitized = ReplaceInvalidCharacters(sanitized)
    If LenB(sanitized) = 0 Then sanitized = "chunk"
    SanitizeChunk = sanitized
End Function

'@label:ReplaceInvalidCharacters
'@prop-title Replace characters unsupported in filenames
Private Function ReplaceInvalidCharacters(ByVal value As String) As String
    Dim invalidChar As Variant
    Dim newValue As String

    EnsureSanitizeList

    For Each invalidChar In this.sanitizeList
        newValue = Replace(value, CStr(invalidChar), "_")
    Next invalidChar

    ReplaceInvalidCharacters = newValue
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
Private Function BuildVersionSuffix(ByVal pass As IPasswords) As String
    Dim versionValue As String
    Dim publicKeyValue As String

    versionValue = SanitizeChunk(pass.Value("version"))
    publicKeyValue = SanitizeChunk(pass.Value("publickey"))

    BuildVersionSuffix = "__v" & versionValue & "-" & publicKeyValue
End Function

'@label:BuildTimestamp
'@prop-title Compose timestamp fragment for filenames
Private Function BuildTimestamp() As String
    BuildTimestamp = Format$(Now, TIMESTAMP_FORMAT)
End Function

'@label:DefaultTemplate
'@prop-title Provide fallback filename template
Private Function DefaultTemplate(ByVal exportNumber As Long) As String
    DefaultTemplate = "export" & CStr(exportNumber)
End Function


'@label:WorkbookBaseName
'@prop-title Remove Excel extensions from workbook name
Private Function WorkbookBaseName(ByVal workbook As Workbook) As String
    Dim cleaned As String

    cleaned = LCase$(workbook.Name)
    cleaned = Replace(cleaned, ".xlsm", vbNullString)
    cleaned = Replace(cleaned, ".xlsb", vbNullString)
    cleaned = Replace(cleaned, ".xlsx", vbNullString)
    cleaned = Replace(cleaned, ".xls", vbNullString)

    WorkbookBaseName = cleaned
End Function

'@label:LogInfo
'@prop-title Append informational diagnostics to the Checking log
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
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub

'@section Interface Implementation
'===============================================================================
Private Property Get ILLExport_Data() As IDataSheet
    Set ILLExport_Data = Data
End Property

Private Property Get ILLExport_Worksheet() As Worksheet
    Set ILLExport_Worksheet = Worksheet
End Property

Private Property Get ILLExport_NumberOfExports() As Long
    ILLExport_NumberOfExports = NumberOfExports
End Property

Private Property Get ILLExport_HasCheckings() As Boolean
    ILLExport_HasCheckings = HasCheckings
End Property

Private Property Get ILLExport_CheckingValues() As IChecking
    Set ILLExport_CheckingValues = CheckingValues
End Property

Private Sub ILLExport_AddRows(Optional ByVal dict As ILLdictionary)
    AddRows dict
End Sub

Private Sub ILLExport_InsertRows(ByVal targetCell As Range, _
                                 Optional ByVal dict As ILLdictionary, _
                                 Optional ByVal insertShift As Boolean = False)
    InsertRows targetCell, dict, insertShift
End Sub

Private Sub ILLExport_DeleteRows(ByVal targetCell As Range, _
                                 Optional ByVal dict As ILLdictionary, _
                                 Optional ByVal forceShift As Boolean = False)
    DeleteRows targetCell, dict, forceShift
End Sub

Private Sub ILLExport_RemoveRows(Optional ByVal dict As ILLdictionary, Optional ByVal rowCount As Long = 2)
    RemoveRows dict, rowCount
End Sub

Private Sub ILLExport_ManageRows(Optional ByVal del As Boolean = False, _ 
                                 Optional ByVal dict As ILLdictionary, _ 
                                 Optional ByVal rowCount As Long = 2)
    ManageRows del:=del, dict:=dict, rowCount:=rowCount
End Sub

Private Sub ILLExport_SyncDictionaryExports(Optional ByVal dict As ILLdictionary)
    SyncDictionaryExports dict
End Sub

Private Sub ILLExport_Sort(Optional ByVal dict As ILLdictionary)
    Sort dict
End Sub

Private Sub ILLExport_ImportSpecs(ByVal fromWksh As Worksheet, _
                                  ByVal fromStartRow As Long, _
                                  ByVal fromStartColumn As Long)
    ImportSpecs fromWksh, fromStartRow, fromStartColumn
End Sub

Private Sub ILLExport_ExportSpecs(ByVal toWkb As Workbook, _
                                  Optional ByVal Hide As Long = xlSheetHidden)
    ExportSpecs toWkb, Hide
End Sub

Private Sub ILLExport_InvalidateCaches()
    ResetCaches
End Sub

Private Function ILLExport_ExportFileName(ByVal exportNumber As Long, _
                                          ByVal dict As ILLdictionary, _
                                          ByVal pass As IPasswords, _
                                          Optional ByVal exportAll As Boolean = False) As String
    ILLExport_ExportFileName = ExportFileName(exportNumber, dict, pass, exportAll)
End Function

Private Function ILLExport_IsActive(ByVal exportNumber As Long) As Boolean
    ILLExport_IsActive = IsActive(exportNumber)
End Function

Private Function ILLExport_ColumnValue(ByVal exportNumber As Long, ByVal columnName As String) As String
    ILLExport_ColumnValue = ColumnValue(exportNumber, columnName)
End Function

Private Function ILLExport_ActiveExportNumbers() As BetterArray
    Set ILLExport_ActiveExportNumbers = ActiveExportNumbers
End Function


Private Sub ILLExport_Translate(ByVal trads As ITranslationObject)
    Translate trads
End Sub


Private Sub ILLExport_ResetCaches()
    ResetCaches
End Sub