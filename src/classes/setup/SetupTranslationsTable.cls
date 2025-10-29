VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetupTranslationsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Improved translations table manager leveraging caching and guarded updates"

'@PredeclaredId
'@Folder("Setup")
'@ModuleDescription("Improved translations table manager leveraging caching and guarded updates")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@depends ProjectError, ApplicationState, IApplicationState, HiddenNames, IHiddenNames

Option Explicit

Implements ISetupTranslationsTable

'@section State and constants
'===============================================================================
Private Type TSetupTranslations
    table As ListObject               'Managed translations table
    tagHeader As Range                'Header cell immediately to the left storing translation tags
    labelCache As BetterArray         'Cached label lookups for label to row mappings
    processedRanges As BetterArray    'Ranges processed during the current update cycle
    displayPrompts As Boolean         'Toggle to avoid message prompts during automated tests
    stateScope As IApplicationState   'Optional application guard for busy sections
    tagsIntegrated As Boolean         'Tracks whether the helper tag column is currently merged in the table
End Type

Private this As TSetupTranslations

Private Const CLASS_NAME As String = "SetupTranslationsTable"
Private Const COUNTER_NAME As String = "_SetupTranslationsCounter"
Private Const STATUS_UPDATED As String = "yes"
Private Const MODE_TEXT As String = "translate as text"
Private Const MODE_FORMULA As String = "translate as formula"
Private Const TAG_HEADER_TITLE As String = "__TagInternal__"
Private Const TAG_SEPARATOR As String = "__"


'@section Factory helpers
'===============================================================================
'@sub-title Create a configured translations table manager instance.
'@param translationTable ListObject representing the translations table.
'@param stateScope Optional scope guarding Excel state while updates run.
'@return ISetupTranslationsTable ready to process registry updates.
Public Function Create(ByVal translationTable As ListObject, Optional ByVal stateScope As IApplicationState = Nothing) As ISetupTranslationsTable
    Dim instance As SetupTranslationsTable

    ValidateTranslationTable translationTable

    Set instance = New SetupTranslationsTable
    instance.InitialiseCore translationTable, stateScope

    Set Create = instance.Self
End Function

'@sub-title Store the supplied table and optional scope inside the current instance.
Public Sub InitialiseCore(ByVal translationTable As ListObject, Optional ByVal stateScope As IApplicationState = Nothing)
    With this
        Set .table = translationTable
        Set .stateScope = stateScope
        .displayPrompts = True
        .tagsIntegrated = False
    End With

    ConfigureTagColumn
    ResetCache
    ResetProcessedRanges
End Sub


'@section Public configuration
'===============================================================================
'@sub-title Worksheet hosting the translations table.
Public Property Get HostWorksheet() As Worksheet
    ValidateInitialised
    Set HostWorksheet = this.table.Parent
End Property

'@sub-title Name of the managed translations listobject.
Public Property Get TableName() As String
    ValidateInitialised
    TableName = this.table.Name
End Property

'@sub-title Self reference used by factories returning interface instances.
Public Property Get Self() As ISetupTranslationsTable
    Set Self = Me
End Property

'@sub-title Allow callers to toggle display prompts used during updates.
Public Sub SetDisplayPrompts(ByVal state As Boolean)
    this.displayPrompts = state
End Sub

'@sub-title Guarded access to the application scope.
Private Function ResolveScope() As IApplicationState
    If this.stateScope Is Nothing Then
        Set this.stateScope = ApplicationState.Create(Application)
    End If

    Set ResolveScope = this.stateScope
End Function


'@section Language management
'===============================================================================
'@sub-title Ensure all provided languages exist as table headers.
Public Sub EnsureLanguages(Optional ByVal languages As String = vbNullString)
    ValidateInitialised

    Dim languageList() As String
    Dim counter As Long
    Dim languageName As String

    If LenB(Trim$(languages)) = 0 Then Exit Sub

    languageList = Split(languages, ";")
    For counter = LBound(languageList) To UBound(languageList)
        languageName = Trim$(languageList(counter))
        If LenB(languageName) > 0 Then
            EnsureLanguageColumn languageName
        End If
    Next counter
End Sub

'@sub-title Create the language column when it does not exist.
Private Sub EnsureLanguageColumn(ByVal languageName As String)
    Dim column As ListColumn

    For Each column In this.table.ListColumns
        If StrComp(column.Name, languageName, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next column

    With this.table.ListColumns.Add
        .Name = languageName
        On Error Resume Next
            .Range.EntireColumn.ColumnWidth = 40
        On Error GoTo 0
    End With
End Sub

'@sub-title Promote the requested language column to become the primary language.
Public Sub SwitchDefaultLanguage(ByVal languageName As String)
    Dim targetIndex As Long
    Dim scope As IApplicationState

    ValidateInitialised

    languageName = Trim$(languageName)
    If LenB(languageName) = 0 Then Exit Sub

    targetIndex = LanguageColumnIndex(languageName)
    If targetIndex = 0 Then
        ThrowProjectError ProjectError.InvalidArgument, "Language '" & languageName & "' does not exist in the translations table."
    End If

    If targetIndex = LabelColumnIndex() Then Exit Sub

    Set scope = ResolveScope()
    scope.ApplyBusyState True

    On Error GoTo CleanFail

    SwapLanguageColumns 1, targetIndex
    ResetCache
    EnsureCache
    SortTranslations
    ApplyFormatting

    scope.Restore
    Exit Sub

CleanFail:
    If Not scope Is Nothing Then scope.Restore
    ThrowProjectError ProjectError.SomethingWentWrong, "Failed to switch default language: " & Err.Description
End Sub

'@sub-title Retrieve the available languages excluding the default column.
Public Function SwitchableLanguages() As BetterArray
    Dim languages As BetterArray
    Dim columnIndex As Long
    Dim startIndex As Long

    ValidateInitialised

    Set languages = CreateBuffer()

    startIndex = LabelColumnIndex() + 1
    If startIndex > this.table.ListColumns.Count Then
        Set SwitchableLanguages = languages
        Exit Function
    End If

    For columnIndex = startIndex To this.table.ListColumns.Count
        languages.Push this.table.ListColumns(columnIndex).Name
    Next columnIndex

    Set SwitchableLanguages = languages
End Function

'@section Language utilities
'===============================================================================
'@sub-title Resolve the column index for a provided language header.
Private Function LanguageColumnIndex(ByVal languageName As String) As Long
    Dim trimmedName As String
    Dim columnIndex As Long
    Dim startIndex As Long

    trimmedName = Trim$(languageName)
    If LenB(trimmedName) = 0 Then Exit Function

    startIndex = LabelColumnIndex()
    For columnIndex = startIndex To this.table.ListColumns.Count
        If StrComp(this.table.ListColumns(columnIndex).Name, trimmedName, vbTextCompare) = 0 Then
            LanguageColumnIndex = columnIndex
            Exit Function
        End If
    Next columnIndex
End Function

'@sub-title Swap the headers, data, and totals between two language columns.
Private Sub SwapLanguageColumns(ByVal firstIndex As Long, ByVal secondIndex As Long)
    Dim firstColumn As ListColumn
    Dim secondColumn As ListColumn
    Dim firstHeader As String
    Dim secondHeader As String
    Dim temporaryHeader As String
    Dim firstRange As Range
    Dim secondRange As Range
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim hasTotalsRow As Boolean
    Dim firstTotalsCalc As XlTotalsCalculation
    Dim secondTotalsCalc As XlTotalsCalculation
    Dim firstTotalFormula As Variant
    Dim secondTotalFormula As Variant

    Set firstColumn = this.table.ListColumns(firstIndex)
    Set secondColumn = this.table.ListColumns(secondIndex)

    firstHeader = firstColumn.Name
    secondHeader = secondColumn.Name

    On Error Resume Next
        Set firstRange = firstColumn.DataBodyRange
        Set secondRange = secondColumn.DataBodyRange
    On Error GoTo 0

    hasTotalsRow = this.table.ShowTotals
    If hasTotalsRow Then
        firstTotalsCalc = firstColumn.TotalsCalculation
        secondTotalsCalc = secondColumn.TotalsCalculation
        firstTotalFormula = firstColumn.Total
        secondTotalFormula = secondColumn.Total
    End If

    temporaryHeader = UniqueTemporaryHeader()
    secondColumn.Name = temporaryHeader
    firstColumn.Name = secondHeader
    secondColumn.Name = firstHeader

    If hasTotalsRow Then
        firstColumn.TotalsCalculation = secondTotalsCalc
        secondColumn.TotalsCalculation = firstTotalsCalc
        secondColumn.Total = firstTotalFormula
        firstColumn.Total = secondTotalFormula
    End If

    If firstRange Is Nothing Or secondRange Is Nothing Then Exit Sub

    rowCount = firstRange.Rows.Count
    If secondRange.Rows.Count < rowCount Then
        rowCount = secondRange.Rows.Count
    End If
    If rowCount = 0 Then Exit Sub

    For rowIndex = 1 To rowCount
        SwapCellContents firstRange.Cells(rowIndex, 1), secondRange.Cells(rowIndex, 1)
    Next rowIndex
End Sub

'@sub-title Generate a temporary header that does not collide with existing columns.
Private Function UniqueTemporaryHeader() As String
    Dim candidate As String
    Dim attempt As Long

    attempt = 0
    Do
        candidate = "__switch_tmp_" & CStr(attempt)
        attempt = attempt + 1
    Loop While LanguageColumnIndex(candidate) <> 0

    UniqueTemporaryHeader = candidate
End Function

'@sub-title Swap the formula and number format between two cells.
Private Sub SwapCellContents(ByVal leftCell As Range, ByVal rightCell As Range)
    Dim tempFormula As Variant
    Dim tempNumberFormat As String

    tempFormula = leftCell.Formula
    leftCell.Formula = rightCell.Formula
    rightCell.Formula = tempFormula

    tempNumberFormat = leftCell.NumberFormat
    leftCell.NumberFormat = rightCell.NumberFormat
    rightCell.NumberFormat = tempNumberFormat
End Sub


'@section Registry processing
'===============================================================================
'@sub-title Reset the hidden update sequence counter to zero.
Public Sub ResetSequence(ByVal registrySheet As Worksheet)
    ValidateInitialised
    ValidateRegistrySheet registrySheet
    PersistCounterValue registrySheet, 0
End Sub

'@sub-title Update the translations table using the supplied registry sheet.
Public Sub UpdateFromRegistry(ByVal registrySheet As Worksheet, Optional ByVal languages As String = vbNullString)
    Dim scope As IApplicationState
    Dim updateSequence As Long

    ValidateInitialised
    ValidateRegistrySheet registrySheet

    EnsureLanguages languages

    EnsureTagColumnReady

    On Error GoTo CleanFail

    BeginTagIntegration
    Set scope = ResolveScope()
    scope.ApplyBusyState True

    updateSequence = NextUpdateSequence(registrySheet)
    ResetCache
    ResetProcessedRanges
    EnsureCache
    ProcessRegistry registrySheet, updateSequence
    RemoveObsoleteLabels updateSequence
    ResetCache
    SortTranslations
    ApplyFormatting
    NumberOfMissing

    EndTagIntegration
    scope.Restore
    Exit Sub

CleanFail:
    EndTagIntegration
    If Not scope Is Nothing Then scope.Restore
    ThrowProjectError ProjectError.SomethingWentWrong, "Failed to update translations: " & Err.Description
End Sub

'@sub-title Iterate all registry tables and process each watcher row.
Private Sub ProcessRegistry(ByVal registrySheet As Worksheet, ByVal updateSequence As Long)
    Dim registryTable As ListObject

    For Each registryTable In registrySheet.ListObjects
        ProcessRegistryTable registryTable, updateSequence
    Next registryTable
End Sub

'@sub-title Process a single registry listobject.
Private Sub ProcessRegistryTable(ByVal registryTable As ListObject, ByVal updateSequence As Long)
    Dim row As ListRow
    Dim rngName As String
    Dim status As String
    Dim mode As String
    Dim rngTag As String
    Dim sourceRange As Range

    For Each row In registryTable.ListRows
        rngName = Trim$(CStr(row.Range.Cells(1, 1).Value))
        status = Trim$(CStr(row.Range.Cells(1, 2).Value))
        mode = Trim$(CStr(row.Range.Cells(1, 3).Value))

        If ShouldProcessRange(status, updateSequence) And LenB(rngName) > 0 Then
            TrackProcessedRange rngName
            Set sourceRange = ResolveNamedRange(registryTable.Parent.Parent, rngName)
            If Not sourceRange Is Nothing Then
                rngTag = rngName & TAG_SEPARATOR & CStr(updateSequence)
                ProcessSourceRange sourceRange, rngTag, mode
            End If
        End If
    Next row
End Sub

'@sub-title Determine whether the named range should be processed this cycle.
Private Function ShouldProcessRange(ByVal status As String, ByVal updateSequence As Long) As Boolean
    If updateSequence = 1 Then
        ShouldProcessRange = True
    Else
        ShouldProcessRange = (StrComp(status, STATUS_UPDATED, vbTextCompare) = 0)
    End If
End Function

'@sub-title Dispatch processing based on the requested translation mode.
Private Sub ProcessSourceRange(ByVal sourceRange As Range, ByVal rngTag As String, ByVal mode As String)
    Select Case LCase$(Trim$(mode))
        Case MODE_FORMULA
            ProcessFormulaRange sourceRange, rngTag
        Case MODE_TEXT, vbNullString
            ProcessTextRange sourceRange, rngTag
        Case Else
            ThrowProjectError ProjectError.InvalidArgument, "Unknown translation mode '" & mode & "' for range '" & rngTag & "'."
    End Select
End Sub

'@sub-title Add literal cell values to the translations table.
Private Sub ProcessTextRange(ByVal sourceRange As Range, ByVal rngTag As String)
    Dim cell As Range

    For Each cell In sourceRange.Cells
        If LenB(Trim$(CStr(cell.Value))) > 0 Then
            AddChunk CStr(cell.Value), rngTag
        End If
    Next cell
End Sub

'@sub-title Extract quoted portions of formulas and add them as chunks.
Private Sub ProcessFormulaRange(ByVal sourceRange As Range, ByVal rngTag As String)
    Dim cell As Range
    Dim tokens As BetterArray
    Dim counter As Long
    Dim chunk As String

    For Each cell In sourceRange.Cells
        Set tokens = ExtractFormulaChunks(CStr(cell.Formula))
        If Not tokens Is Nothing Then
            For counter = tokens.LowerBound To tokens.UpperBound
                chunk = Trim$(CStr(tokens.Item(counter)))
                If LenB(chunk) > 0 Then
                    AddChunk chunk, rngTag
                End If
            Next counter
        End If
    Next cell
End Sub

'@sub-title Extract quoted text segments from a formula while handling escaped quotes.
Private Function ExtractFormulaChunks(ByVal formulaText As String) As BetterArray
    Dim cleaned As String
    Dim position As Long
    Dim chunkStart As Long
    Dim buffer As BetterArray
    Dim token As String

    cleaned = Replace(formulaText, Chr$(34) & Chr$(34), vbNullString)

    For position = 1 To Len(cleaned)
        If Mid$(cleaned, position, 1) = Chr$(34) Then
            If chunkStart = 0 Then
                chunkStart = position + 1
            ElseIf position > chunkStart Then
                token = Mid$(cleaned, chunkStart, position - chunkStart)
                If buffer Is Nothing Then Set buffer = CreateBuffer()
                buffer.Push token
                chunkStart = 0
            End If
        End If
    Next position

    Set ExtractFormulaChunks = buffer
End Function


'@section Chunk management
'===============================================================================
'@sub-title Add or update a translation chunk with caching support.
Private Sub AddChunk(ByVal label As String, ByVal rngTag As String)
    Dim rowIndex As Long
    Dim trimmedLabel As String

    trimmedLabel = Trim$(label)
    If LenB(trimmedLabel) = 0 Then Exit Sub

    rowIndex = CachedRowIndex(trimmedLabel)
    If rowIndex > 0 Then
        SetTagForRow rowIndex, rngTag
    Else
        rowIndex = AppendRow(trimmedLabel, rngTag)
        AddToCache trimmedLabel, rowIndex
    End If
End Sub

'@sub-title Retrieve cached row index when the label exists.
Private Function CachedRowIndex(ByVal label As String) As Long
    Dim cache As BetterArray
    Dim idx As Long
    Dim entry As Variant

    EnsureCache
    Set cache = this.labelCache
    If cache Is Nothing Then Exit Function
    If cache.Length = 0 Then Exit Function

    For idx = cache.LowerBound To cache.UpperBound
        entry = cache.Item(idx)
        If MatchLabel(CStr(entry(0)), label) Then
            CachedRowIndex = CLng(entry(1))
            Exit For
        End If
    Next idx
End Function

'@sub-title Persist the label inside the cache for future lookups.
Private Sub AddToCache(ByVal label As String, ByVal rowIndex As Long)
    Dim entry(0 To 1) As Variant
    Dim existing As Variant
    Dim idx As Long

    If this.labelCache Is Nothing Then
        Set this.labelCache = CreateBuffer()
    End If

    If this.labelCache.Length > 0 Then
        For idx = this.labelCache.LowerBound To this.labelCache.UpperBound
            existing = this.labelCache.Item(idx)
            If MatchLabel(CStr(existing(0)), label) Then Exit Sub
        Next idx
    End If

    entry(0) = label
    entry(1) = rowIndex
    this.labelCache.Push entry
End Sub

'@sub-title Append a new list row and populate the mandatory cells.
Private Function AppendRow(ByVal label As String, ByVal rngTag As String) As Long
    Dim newRow As ListRow

    Set newRow = this.table.ListRows.Add
    newRow.Range.Cells(1, LabelColumnIndex).Value = label
    SetTagForRow newRow.Index, rngTag

    AppendRow = newRow.Index
End Function

'@sub-title Update the helper column with the latest tag for the given row.
Private Sub SetTagForRow(ByVal rowIndex As Long, ByVal rngTag As String)
    Dim cell As Range

    Set cell = TagDataCell(rowIndex)
    If cell Is Nothing Then Exit Sub

    cell.Value = rngTag
End Sub

'@sub-title Build or refresh the label cache from the current table state.
Private Sub EnsureCache()
    Dim data As Range
    Dim values As BetterArray
    Dim rowIndex As Long
    Dim label As String

    If Not this.labelCache Is Nothing Then Exit Sub

    Set this.labelCache = CreateBuffer()
    Set data = LabelColumnData()

    If data Is Nothing Then Exit Sub
    Set values = New BetterArray
    values.FromExcelRange data, DetectLastRow:=False, DetectLastColumn:=False
    Debug.Print "Retrieved values"
    Debug.Print "Values lowerbound"; values.LowerBound
    Debug.Print "Values upperbound"; values.UpperBound
    Debug.Print "First Item"; values.Item(values.LowerBound)

    For rowIndex = values.LowerBound To values.UpperBound
        Debug.Print "Row index: "; rowIndex
        label = Trim$(CStr(values.Item(rowIndex)))
        Debug.Print "Label: "; label
        If LenB(label) > 0 Then
            AddToCache label, rowIndex
        End If

    Next rowIndex
End Sub

'@sub-title Clear the cached labels forcing the next EnsureCache to rebuild it.
Private Sub ResetCache()
    Set this.labelCache = Nothing
End Sub

'@sub-title Helper to compare cached labels using trimmed case-insensitive comparison.
Private Function MatchLabel(ByVal cachedLabel As String, ByVal candidate As String) As Boolean
    MatchLabel = (StrComp(Trim$(cachedLabel), Trim$(candidate), vbTextCompare) = 0)
End Function

'@section Process tracking
'===============================================================================
'@sub-title Track the ranges processed during the current registry cycle.
Private Sub TrackProcessedRange(ByVal rngName As String)
    Dim trimmed As String
    Dim idx As Long

    trimmed = Trim$(rngName)
    If LenB(trimmed) = 0 Then Exit Sub

    If this.processedRanges Is Nothing Then
        Set this.processedRanges = CreateBuffer()
    End If

    If this.processedRanges.Length = 0 Then
        this.processedRanges.Push trimmed
        Exit Sub
    End If

    For idx = this.processedRanges.LowerBound To this.processedRanges.UpperBound
        If StrComp(CStr(this.processedRanges.Item(idx)), trimmed, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next idx

    this.processedRanges.Push trimmed
End Sub

'@sub-title Determine whether a registry range was processed during the active cycle.
Private Function WasRangeProcessed(ByVal rngName As String) As Boolean
    Dim trimmed As String
    Dim idx As Long

    If this.processedRanges Is Nothing Then Exit Function
    If this.processedRanges.Length = 0 Then Exit Function

    trimmed = Trim$(rngName)
    For idx = this.processedRanges.LowerBound To this.processedRanges.UpperBound
        If StrComp(CStr(this.processedRanges.Item(idx)), trimmed, vbTextCompare) = 0 Then
            WasRangeProcessed = True
            Exit Function
        End If
    Next idx
End Function

'@sub-title Reset processed ranges tracking.
Private Sub ResetProcessedRanges()
    Set this.processedRanges = Nothing
End Sub

'@sub-title Remove labels that no longer exist in processed ranges.
Private Sub RemoveObsoleteLabels(ByVal updateSequence As Long)
    Dim rowIndex As Long
    Dim tagText As String
    Dim rangeName As String
    Dim sequenceValue As Long
    Dim shouldDelete As Boolean

    If this.table.ListRows.Count = 0 Then Exit Sub

    For rowIndex = this.table.ListRows.Count To 1 Step -1
        tagText = TagValueAt(rowIndex)
        shouldDelete = False

        If LenB(tagText) = 0 Then
            shouldDelete = True
        ElseIf ParseTag(tagText, rangeName, sequenceValue) Then
            If WasRangeProcessed(rangeName) And sequenceValue < updateSequence Then
                shouldDelete = True
            End If
        End If

        If shouldDelete Then
            DeleteRowAt rowIndex
        End If
    Next rowIndex
End Sub

'@sub-title Delete a row within the translations table when the index is valid.
Private Sub DeleteRowAt(ByVal rowIndex As Long)
    Dim helperCell As Range

    If rowIndex < 1 Then Exit Sub
    If rowIndex > this.table.ListRows.Count Then Exit Sub

    If Not this.tagsIntegrated Then
        Set helperCell = TagDataCell(rowIndex)
        If Not helperCell Is Nothing Then
            On Error Resume Next
                helperCell.Delete Shift:=xlShiftUp
            On Error GoTo 0
        End If
    End If

    this.table.ListRows(rowIndex).Delete
End Sub

'@sub-title Parse helper tag text into its components.
Private Function ParseTag(ByVal tagText As String, ByRef rangeName As String, ByRef sequenceValue As Long) As Boolean
    Dim parts() As String
    Dim numberText As String

    parts = Split(tagText, TAG_SEPARATOR)
    If UBound(parts) < 1 Then Exit Function

    rangeName = Trim$(parts(0))
    numberText = Trim$(parts(1))

    If LenB(rangeName) = 0 Then Exit Function
    If LenB(numberText) = 0 Then Exit Function
    If Not IsNumeric(numberText) Then Exit Function

    sequenceValue = CLng(numberText)
    ParseTag = True
End Function

'@sub-title Create a BetterArray buffer with deterministic bounds.
Private Function CreateBuffer() As BetterArray
    Dim buffer As BetterArray
    Set buffer = New BetterArray
    buffer.LowerBound = 0
    Set CreateBuffer = buffer
End Function

'@section Reporting
'===============================================================================
'@sub-title Summarise missing translations per language and optionally prompt the user.
Public Function NumberOfMissing() As String
    Dim columnIndex As Long
    Dim dataRange As Range
    Dim missingCount As Long
    Dim messageLines As BetterArray
    Dim summary As String
    Dim firstLanguageIndex As Long

    ValidateInitialised

    Set messageLines = CreateBuffer()
    messageLines.Push "Translation Updated!"

    firstLanguageIndex = LabelColumnIndex() + 1
    If firstLanguageIndex > this.table.ListColumns.Count Then GoTo BuildSummary

    For columnIndex = firstLanguageIndex To this.table.ListColumns.Count
        On Error Resume Next
            Set dataRange = this.table.ListColumns(columnIndex).DataBodyRange
        On Error GoTo 0

        If Not dataRange Is Nothing Then
            missingCount = Application.WorksheetFunction.CountBlank(dataRange)
            If missingCount > 0 Then
                messageLines.Push missingCount & " labels are missing for column " & this.table.ListColumns(columnIndex).Name & "."
            End If
        End If
        Set dataRange = Nothing
    Next columnIndex

BuildSummary:
    summary = JoinLines(messageLines)
    NumberOfMissing = summary

    If this.displayPrompts Then
        MsgBox summary, vbInformation + vbOKOnly, "Done!"
    End If
End Function

'@sub-title Join BetterArray text lines using environment specific new lines.
Private Function JoinLines(ByVal lines As BetterArray) As String
    Dim idx As Long
    Dim builder As String

    If lines Is Nothing Then Exit Function
    If lines.Length = 0 Then Exit Function

    For idx = lines.LowerBound To lines.UpperBound
        If Len(builder) = 0 Then
            builder = CStr(lines.Item(idx))
        Else
            builder = builder & vbLf & CStr(lines.Item(idx))
        End If
    Next idx

    JoinLines = builder
End Function

'@sub-title Access the label column data range.
Private Function LabelColumnData() As Range
    If this.table.ListRows.Count = 0 Then Exit Function
    On Error Resume Next
        Set LabelColumnData = this.table.ListColumns(LabelColumnIndex).DataBodyRange
    On Error GoTo 0
End Function

'@sub-title Retrieve the helper tag column data range.
Private Function TagColumnDataRange() As Range
    Dim helperHeader As Range

    If this.table.ListRows.Count = 0 Then Exit Function

    If this.tagsIntegrated Then
        On Error Resume Next
            Set TagColumnDataRange = this.table.ListColumns(1).DataBodyRange
        On Error GoTo 0
    Else
        Set helperHeader = RequireTagHeader()
        Set TagColumnDataRange = helperHeader.Offset(1, 0).Resize(this.table.ListRows.Count, 1)
    End If
End Function

'@sub-title Ensure the helper column header is located and configured.
Private Sub ConfigureTagColumn()
    Dim header As Range

    ValidateInitialised

    Set header = this.table.HeaderRowRange.Cells(1, 1)
    If header.Column = 1 Then
        ThrowProjectError ProjectError.ErrorUnexpectedState, "Translations table requires a helper column immediately to the left."
    End If

    Set this.tagHeader = header.Offset(0, -1)
    this.tagHeader.Value = TAG_HEADER_TITLE
End Sub

Private Sub EnsureTagColumnReady()
    Dim unused As Range
    Set unused = RequireTagHeader()
End Sub

Private Function RequireTagHeader() As Range
    If Not IsRangeValid(this.tagHeader) Then
        AttachTagHeader
    End If

    If Not IsRangeValid(this.tagHeader) Then
        ThrowProjectError ProjectError.ErrorUnexpectedState, _
                          "Translations table requires a helper column titled '" & TAG_HEADER_TITLE & "' immediately to the left of the listobject."
    End If

    Set RequireTagHeader = this.tagHeader
End Function

Private Sub AttachTagHeader()
    On Error GoTo Failed
        ConfigureTagColumn
        Exit Sub
Failed:
    Err.Clear
    Set this.tagHeader = Nothing
End Sub

Private Function IsRangeValid(ByVal candidate As Range) As Boolean
    Dim probe As Variant

    If candidate Is Nothing Then Exit Function

    On Error GoTo Invalid
        probe = candidate.Value
        IsRangeValid = True
    On Error GoTo 0
    Exit Function

Invalid:
    Err.Clear
    IsRangeValid = False
End Function

'@section Tag integration helpers
'===============================================================================
'@sub-title Temporarily include the helper tag column inside the managed table.
Private Sub BeginTagIntegration()
    Dim helperHeader As Range
    Dim expandedRange As Range

    If this.tagsIntegrated Then Exit Sub

    ValidateInitialised
    Set helperHeader = RequireTagHeader()

    Set expandedRange = helperHeader.Resize(this.table.Range.Rows.Count, this.table.Range.Columns.Count + 1)
    this.table.Resize expandedRange
    this.tagsIntegrated = True
End Sub

'@sub-title Restore the translations table to exclude the helper tag column.
Private Sub EndTagIntegration()
    Dim restoredStart As Range
    Dim restoredRange As Range
    Dim totalColumns As Long
    Dim totalRows As Long

    If Not this.tagsIntegrated Then Exit Sub

    ValidateInitialised

    totalColumns = this.table.Range.Columns.Count
    totalRows = this.table.Range.Rows.Count

    If totalColumns > 1 Then
        Set restoredStart = this.table.HeaderRowRange.Cells(1, 2)
        Set restoredRange = restoredStart.Resize(totalRows, totalColumns - 1)
        this.table.Resize restoredRange
    End If

    this.tagsIntegrated = False
End Sub

'@sub-title Resolve the current column index storing translation labels.
Private Function LabelColumnIndex() As Long
    If this.tagsIntegrated Then
        LabelColumnIndex = 2
    Else
        LabelColumnIndex = 1
    End If
End Function

'@sub-title Resolve the current tag data cell for a provided row.
Private Function TagDataCell(ByVal rowIndex As Long) As Range
    Dim tagRange As Range
    Dim tagColumn As ListColumn

    If rowIndex < 1 Then Exit Function
    If rowIndex > this.table.ListRows.Count Then Exit Function

    If this.tagsIntegrated Then
        Set tagColumn = this.table.ListColumns(1)
        If tagColumn.DataBodyRange Is Nothing Then Exit Function
        On Error Resume Next
            Set TagDataCell = tagColumn.DataBodyRange.Cells(rowIndex, 1)
        On Error GoTo 0
    Else
        Set tagRange = TagColumnDataRange()
        If tagRange Is Nothing Then Exit Function
        If rowIndex > tagRange.Rows.Count Then Exit Function
        Set TagDataCell = tagRange.Cells(rowIndex, 1)
    End If
End Function

'@sub-title Retrieve the stored tag text for the provided row index.
Private Function TagValueAt(ByVal rowIndex As Long) As String
    Dim cell As Range

    Set cell = TagDataCell(rowIndex)
    If cell Is Nothing Then Exit Function

    TagValueAt = Trim$(CStr(cell.Value))
End Function


'@section Sorting and formatting
'===============================================================================
'@sub-title Sort the translations by their primary language column.
Private Sub SortTranslations()
    On Error Resume Next
        With this.table.Sort
            .SortFields.Clear
            .SortFields.Add Key:=this.table.ListColumns(LabelColumnIndex).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    On Error GoTo 0
End Sub

'@sub-title Apply light formatting and duplicate highlighting once per update.
Private Sub ApplyFormatting()
    Dim headerRange As Range
    Dim dataRange As Range
    Dim tableRange As Range
    Dim columnRange As Range
    Dim unlockRange As Range
    Dim unlockOffset As Long
    Dim unlockWidth As Long
    Dim tagHeader As Range
    Dim tagData As Range

    Set headerRange = this.table.HeaderRowRange
    Set dataRange = this.table.DataBodyRange
    Set tableRange = this.table.Range

    On Error Resume Next
        this.table.TableStyle = vbNullString
        tableRange.Interior.Color = vbWhite
        tableRange.Font.Color = vbBlack
        tableRange.EntireRow.RowHeight = 30
        tableRange.EntireColumn.ColumnWidth = 60
        headerRange.Font.Color = RGB(34, 96, 162)
        headerRange.Interior.Color = RGB(221, 235, 247)
        headerRange.Locked = False
    On Error GoTo 0

    Set tagHeader = RequireTagHeader()
    tagHeader.Font.Color = vbWhite

    Set tagData = TagColumnDataRange()
    If Not tagData Is Nothing Then tagData.Font.Color = vbWhite

    On Error Resume Next
        With tableRange
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .Color = RGB(34, 96, 162)
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .Color = RGB(34, 96, 162)
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .Color = RGB(34, 96, 162)
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .Color = RGB(34, 96, 162)
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .Color = RGB(34, 96, 162)
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .Color = RGB(34, 96, 162)
            End With
        End With
    On Error GoTo 0

    If dataRange Is Nothing Then Exit Sub

    On Error Resume Next
        dataRange.FormatConditions.Delete
    On Error GoTo 0

    unlockWidth = this.table.ListColumns.Count - LabelColumnIndex()
    If unlockWidth > 0 Then
        unlockOffset = LabelColumnIndex()
        On Error Resume Next
            Set unlockRange = dataRange.Offset(0, unlockOffset).Resize(dataRange.Rows.Count, unlockWidth)
            If Not unlockRange Is Nothing Then unlockRange.Locked = False
        On Error GoTo 0
    End If

    On Error Resume Next
        Set columnRange = dataRange.Columns(LabelColumnIndex())
    On Error GoTo 0

    If columnRange Is Nothing Then Exit Sub

    Do While Not columnRange Is Nothing And Not Intersect(columnRange, dataRange) Is Nothing
        columnRange.FormatConditions.AddUniqueValues
        With columnRange.FormatConditions(columnRange.FormatConditions.Count)
            .DupeUnique = xlDuplicate
            .Interior.Color = vbRed
        End With
        Set columnRange = columnRange.Offset(0, 1)
        If columnRange.Column > dataRange.Columns(dataRange.Columns.Count).Column Then Exit Do
    Loop
End Sub


'@section Registry bookkeeping
'===============================================================================
'@sub-title Retrieve and increment the registry update sequence.
Private Function NextUpdateSequence(ByVal registrySheet As Worksheet) As Long
    Dim currentValue As Long

    currentValue = CounterNumericValue(registrySheet)
    If currentValue < 0 Then currentValue = 0

    NextUpdateSequence = currentValue + 1
    PersistCounterValue registrySheet, NextUpdateSequence
End Function

'@sub-title Persist the counter value inside a hidden worksheet-level name.
Private Sub PersistCounterValue(ByVal registrySheet As Worksheet, ByVal counterValue As Long)
    Dim store As IHiddenNames

    If counterValue < 0 Then counterValue = 0

    Set store = HiddenNameStore(registrySheet)
    If store Is Nothing Then Exit Sub

    store.EnsureName COUNTER_NAME, counterValue, HiddenNameTypeLong
    store.SetValue COUNTER_NAME, counterValue
End Sub

'@sub-title Retrieve the current counter value stored in worksheet names.
Private Function CounterNumericValue(ByVal registrySheet As Worksheet) As Long
    Dim store As IHiddenNames

    Set store = HiddenNameStore(registrySheet)
    If store Is Nothing Then Exit Function

    If Not store.HasName(COUNTER_NAME) Then
        store.EnsureName COUNTER_NAME, 0, HiddenNameTypeLong
        store.SetValue COUNTER_NAME, 0
    End If

    CounterNumericValue = store.ValueAsLong(COUNTER_NAME, 0)
End Function

Private Function HiddenNameStore(ByVal registrySheet As Worksheet) As IHiddenNames
    If registrySheet Is Nothing Then Exit Function
    On Error Resume Next
        Set HiddenNameStore = HiddenNames.Create(registrySheet)
    On Error GoTo 0
End Function


'@section Validation and utilities
'===============================================================================
'@sub-title Ensure the manager has been initialised before use.
Private Sub ValidateInitialised()
    If this.table Is Nothing Then
        ThrowProjectError ProjectError.ObjectNotInitialized, "Translations table was not initialised."
    End If
End Sub

'@sub-title Validate the supplied translations table.
Private Sub ValidateTranslationTable(ByVal translationTable As ListObject)
    If translationTable Is Nothing Then
        ThrowProjectError ProjectError.InvalidArgument, "Please supply a valid translations listobject."
    End If
End Sub

'@sub-title Validate the registry sheet before processing.
Private Sub ValidateRegistrySheet(ByVal registrySheet As Worksheet)
    If registrySheet Is Nothing Then
        ThrowProjectError ProjectError.InvalidArgument, "Registry worksheet cannot be Nothing."
    End If
End Sub

'@sub-title Resolve a named range from the workbook, returning Nothing when missing.
Private Function ResolveNamedRange(ByVal workbook As Workbook, ByVal rangeName As String) As Range
    Dim definedName As Name
    On Error Resume Next
        Set definedName = workbook.Names(rangeName)
    On Error GoTo 0

    If definedName Is Nothing Then Exit Function

    On Error Resume Next
        Set ResolveNamedRange = definedName.RefersToRange
    On Error GoTo 0
End Function

'@sub-title Raise a project error with a contextual message.
Private Sub ThrowProjectError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Sub ISetupTranslationsTable_EnsureLanguages(Optional ByVal languages As String = vbNullString)
    EnsureLanguages languages
End Sub

Private Sub ISetupTranslationsTable_SwitchDefaultLanguage(ByVal languageName As String)
    SwitchDefaultLanguage languageName
End Sub

Private Function ISetupTranslationsTable_SwitchableLanguages() As BetterArray
    Set ISetupTranslationsTable_SwitchableLanguages = SwitchableLanguages
End Function

Private Property Get ISetupTranslationsTable_HostWorksheet() As Worksheet
    Set ISetupTranslationsTable_HostWorksheet = HostWorksheet
End Property

Private Sub ISetupTranslationsTable_Initialise(ByVal translationTable As ListObject, Optional ByVal stateScope As IApplicationState)
    ValidateTranslationTable translationTable
    InitialiseCore translationTable, stateScope
End Sub

Private Property Get ISetupTranslationsTable_Self() As ISetupTranslationsTable
    Set ISetupTranslationsTable_Self = Self
End Property

Private Property Get ISetupTranslationsTable_TableName() As String
    ISetupTranslationsTable_TableName = TableName
End Property

Private Sub ISetupTranslationsTable_UpdateFromRegistry(ByVal registrySheet As Worksheet, Optional ByVal languages As String = vbNullString)
    UpdateFromRegistry registrySheet, languages
End Sub

Private Sub ISetupTranslationsTable_ResetSequence(ByVal registrySheet As Worksheet)
    ResetSequence registrySheet
End Sub

Private Function ISetupTranslationsTable_NumberOfMissing() As String
    ISetupTranslationsTable_NumberOfMissing = NumberOfMissing
End Function

Private Sub ISetupTranslationsTable_SetDisplayPrompts(ByVal state As Boolean)
    SetDisplayPrompts state
End Sub
