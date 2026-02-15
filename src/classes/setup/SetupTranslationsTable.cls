VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetupTranslationsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Translations table manager with helper tag column"

'@PredeclaredId
'@Folder("Setup")
'@ModuleDescription("Translations table manager with helper tag column")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@depends ProjectError, ApplicationState, IApplicationState, HiddenNames, IHiddenNames, CustomTable, ICustomTable

Option Explicit

Implements ISetupTranslationsTable

'@class SetupTranslationsTable
'@description
'Manages a translations ListObject backed by a helper tag column.
'Synchronises translations from a watcher registry, handles language
'column creation and promotion, removes obsolete and duplicate labels,
'reports missing and duplicate entries, and exports the table to
'external workbooks. Uses HiddenNames for sequence counters.

'@section State and constants
'===============================================================================
Private Type TSetupTranslations
    table As ListObject               'Managed translations table
    tagHeader As Range                'Header cell immediately to the left storing translation tags
    labelValues() As Variant          'In-memory label column for lookups during ProcessRegistry
    labelCount As Long                'Number of entries in the label buffer
    processedRanges As BetterArray    'Ranges processed during the current update cycle
    displayPrompts As Boolean         'Toggle to avoid message prompts during automated tests
    stateScope As IApplicationState   'Optional application guard for busy sections
    tagsIntegrated As Boolean         'Tracks whether the helper tag column is currently merged in the table
    missingSummary As BetterArray     'Keep missing summary on non default languages
    duplicatesSummary As BetterArray  'keep duplicate summary on all the languages
    tagBuffer() As Variant            'In-memory tag column buffer for batch writes during ProcessRegistry
    tagBufferSize As Long             'Current number of entries in the tag buffer
    dirty As Boolean                  'Tracks whether rows were added or removed during the current update
End Type

Private this As TSetupTranslations

Private Const CLASS_NAME As String = "SetupTranslationsTable"
Private Const COUNTER_NAME As String = "_SetupTranslationsCounter"
Private Const STATUS_UPDATED As String = "yes"
Private Const MODE_TEXT As String = "translate as text"
Private Const MODE_FORMULA As String = "translate as formula"
Private Const MODE_WATCH As String = "watch for update"
Private Const TAG_HEADER_TITLE As String = "__TagInternal__"
Private Const TAG_SEPARATOR As String = "__"
Private Const LANGUAGES_NAME_ID As String = "__SetupTranslationsLanguages__"
Private Const TRANSLATIONS_NAME_PREFIX As String = "__SetupTranslations"

Private Type TDuplicateSummary
    HasDuplicates As Boolean
    Message As String
End Type


'@section Factory helpers
'===============================================================================
'@label:create
'@sub-title Create a configured translations table manager instance.
'@param translationTable ListObject. The managed translations table.
'@param stateScope Optional IApplicationState. Guard for Excel settings during updates.
'@return ISetupTranslationsTable. Ready to process registry updates.
'@export
Public Function Create(ByVal translationTable As ListObject, Optional ByVal stateScope As IApplicationState = Nothing) As ISetupTranslationsTable
    Dim instance As SetupTranslationsTable

    ValidateTranslationTable translationTable

    Set instance = New SetupTranslationsTable
    instance.InitialiseCore translationTable, stateScope

    Set Create = instance.Self
End Function

'@label:initialise-core
'@sub-title Store the supplied table and optional scope inside the current instance.
Public Sub InitialiseCore(ByVal translationTable As ListObject, Optional ByVal stateScope As IApplicationState = Nothing)
    With this
        Set .table = translationTable
        Set .stateScope = stateScope
        .displayPrompts = True
        .tagsIntegrated = False
    End With

    ConfigureTagColumn
    ResetProcessedRanges
    PersistLanguageList
End Sub


'@section Public configuration
'===============================================================================
'@label:host-worksheet
'@prop-title Worksheet hosting the translations table.
Public Property Get HostWorksheet() As Worksheet
    ValidateInitialised
    Set HostWorksheet = this.table.Parent
End Property

'@label:table-name
'@prop-title Name of the managed translations ListObject.
Public Property Get TableName() As String
    ValidateInitialised
    TableName = this.table.Name
End Property

'@label:self
'@prop-title Self reference used by factories returning interface instances.
Public Property Get Self() As ISetupTranslationsTable
    Set Self = Me
End Property

'@label:set-display-prompts
'@sub-title Allow callers to toggle display prompts used during updates.
Public Sub SetDisplayPrompts(ByVal state As Boolean)
    this.displayPrompts = state
End Sub

'@label:resolve-scope
'@sub-title Guarded access to the application scope.
Private Function ResolveScope() As IApplicationState
    If this.stateScope Is Nothing Then
        Set this.stateScope = ApplicationState.Create(Application)
    End If

    Set ResolveScope = this.stateScope
End Function


'@section Language management
'===============================================================================
'@label:ensure-languages
'@sub-title Ensure all provided languages exist as table headers.
Private Sub EnsureLanguages(Optional ByVal lang As String = vbNullString)
    ValidateInitialised

    Dim languageList() As String
    Dim counter As Long
    Dim languageName As String
    Dim languages As String

    If Not TryGetWrittenLanguages(languages, lang) Then Exit Sub

    languageList = Split(languages, ";")
    For counter = LBound(languageList) To UBound(languageList)
        languageName = Trim$(languageList(counter))
        If LenB(languageName) > 0 Then
            EnsureLanguageColumn languageName
        End If
    Next counter

    PersistLanguageList
End Sub

'@label:try-get-written-languages
'@sub-title Collect language names written to the right of the table.
Private Function TryGetWrittenLanguages(ByRef languages As String, ByVal newLang As String) As Boolean

    Dim cellRng As Range

    Set cellRng = this.table.ListColumns(this.table.ListColumns.Count).Range.Offset(, 1).Cells(1, 1)

    If (LenB(newLang) <> 0) Then languages = newLang

    Do While (LenB(cellRng.Value) <> 0)

        If (LenB(languages) <> 0) Then
            languages = languages & ";" & CStr(cellRng.Value)
        Else
            languages = CStr(cellRng.Value)
        End If

        cellRng.ClearContents
        Set cellRng = cellRng.Offset(, 1)
    Loop

    TryGetWrittenLanguages = (LenB(languages) <> 0)
End Function

'@label:ensure-language-column
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

'@label:persist-language-list
'@sub-title Persist the available languages into the worksheet hidden names.
Private Sub PersistLanguageList()
    Dim store As IHiddenNames
    Dim buffer As BetterArray
    Dim switchable As BetterArray
    Dim idx As Long
    Dim header As String

    Set store = MetadataStore()
    If store Is Nothing Then Exit Sub

    Set buffer = CreateBuffer()

    header = DefaultLanguageHeader()
    If LenB(header) > 0 Then buffer.Push header

    Set switchable = SwitchableLanguages
    If Not switchable Is Nothing Then
        buffer.Concat switchable.Items
    End If

    store.EnsureName LANGUAGES_NAME_ID, vbNullString, HiddenNameTypeString
    store.SetValue LANGUAGES_NAME_ID, buffer.ToString(Separator:=";", OpeningDelimiter:=vbNullString, ClosingDelimiter:=vbNullString)
End Sub

'@label:switch-default-language
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
    scope.ApplyBusyState suppressEvents:=True

    On Error GoTo CleanFail

    SwapLanguageColumns 1, targetIndex
    SortTranslations
    ApplyFormatting
    PersistLanguageList

    scope.Restore
    Exit Sub

CleanFail:
    If Not scope Is Nothing Then scope.Restore
    ThrowProjectError ProjectError.SomethingWentWrong, "Failed to switch default language: " & Err.Description
End Sub

'@label:switchable-languages
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

'@label:languages
'@prop-title Languages available for translation excluding the default column unless requested.
Public Property Get Languages(Optional ByVal includeDefault As Boolean = False) As BetterArray
    Dim allLanguages As BetterArray
    Dim switchable As BetterArray
    Dim defaultHeader As String
    Dim idx As Long

    Set switchable = SwitchableLanguages()

    If Not includeDefault Then
        Set Languages = switchable
        Exit Property
    End If

    Set allLanguages = CreateBuffer()
    defaultHeader = DefaultLanguageHeader()
    If LenB(defaultHeader) > 0 Then allLanguages.Push defaultHeader

    If Not (switchable Is Nothing) Then
        allLanguages.Concat switchable.Items
    End If

    Set Languages = allLanguages
End Property

'@section Language utilities
'===============================================================================
'@label:default-language-header
'@sub-title Resolve the current default language header.
Private Function DefaultLanguageHeader() As String
    Dim labelIndex As Long

    ValidateInitialised

    labelIndex = LabelColumnIndex()
    If labelIndex < 1 Then Exit Function
    If labelIndex > this.table.ListColumns.Count Then Exit Function

    On Error Resume Next
        DefaultLanguageHeader = CStr(this.table.ListColumns(labelIndex).Name)
    On Error GoTo 0
End Function

'@label:language-column-index
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

'@label:resolve-language-column
'@sub-title Resolve the data body range for a given language column.
Private Function ResolveLanguageColumn(ByVal languageName As String) As Range
    Dim columnIndex As Long
    Dim columnRange As Range

    columnIndex = LanguageColumnIndex(languageName)
    If columnIndex = 0 Then Exit Function

    On Error Resume Next
        Set columnRange = this.table.ListColumns(columnIndex).DataBodyRange
    On Error GoTo 0

    Set ResolveLanguageColumn = columnRange
End Function

'@label:swap-language-columns
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

'@label:unique-temporary-header
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

'@label:swap-cell-contents
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
'@label:reset-sequence
'@sub-title Reset the hidden update sequence counter to zero.
Public Sub ResetSequence(ByVal registrySheet As Worksheet)
    ValidateInitialised
    ValidateRegistrySheet registrySheet
    PersistCounterValue registrySheet, 0
End Sub

'@label:update-from-registry
'@sub-title Update the translations table using the supplied registry sheet.
Public Sub UpdateFromRegistry(ByVal registrySheet As Worksheet, _
                             Optional ByVal languages As String = vbNullString)
    Dim scope As IApplicationState
    Dim updateSequence As Long
    Dim errNumb As Long
    Dim errDesc As String

    ValidateInitialised
    ValidateRegistrySheet registrySheet

    EnsureLanguages languages

    EnsureTagColumnReady

    On Error GoTo CleanFail

    BeginTagIntegration
    Set scope = ResolveScope()
    scope.ApplyBusyState True

    updateSequence = NextUpdateSequence(registrySheet)
    ResetProcessedRanges
    this.dirty = False

    InitTagBuffer
    InitLabelBuffer
    ProcessRegistry registrySheet, updateSequence
    FlushTagBuffer

    RemoveObsoleteLabels updateSequence
    DeduplicateLabels
    SortTranslations
    If this.dirty Then ApplyFormatting
    NumberOfMissing
    UpdateDuplicatesSummary

    EndTagIntegration
    scope.Restore
    Exit Sub

CleanFail:
    errNumb = Err.Number
    errDesc = Err.Description
    EndTagIntegration
    If Not scope Is Nothing Then scope.Restore
    ThrowProjectError errNumb, "Failed to update translations: " & errDesc
End Sub

'@label:process-registry
'@sub-title Iterate all registry tables and process each watcher row.
Private Sub ProcessRegistry(ByVal registrySheet As Worksheet, ByVal updateSequence As Long)
    Dim registryTable As ListObject

    For Each registryTable In registrySheet.ListObjects
        ProcessRegistryTable registryTable, updateSequence
    Next registryTable
End Sub

'@label:process-registry-table
'@sub-title Process a single registry ListObject.
Private Sub ProcessRegistryTable(ByVal registryTable As ListObject, ByVal updateSequence As Long)
    Dim rngName As String
    Dim status As String
    Dim mode As String
    Dim rngTag As String
    Dim sourceRange As Range
    Dim dataRange As Range
    Dim values As Variant
    Dim rowIndex As Long
    Dim columnCount As Long
    Dim columnLower As Long

    Set dataRange = registryTable.DataBodyRange
    If dataRange Is Nothing Then Exit Sub

    values = dataRange.Value2
    If Not IsArray(values) Then
        rngName = Trim$(CStr(values))
        status = vbNullString
        mode = vbNullString
        If ShouldProcessRange(status, updateSequence) And LenB(rngName) > 0 Then
            TrackProcessedRange rngName
            Set sourceRange = ResolveNamedRange(registryTable.Parent.Parent, rngName)
            If Not sourceRange Is Nothing Then
                rngTag = rngName & TAG_SEPARATOR & CStr(updateSequence)
                ProcessSourceRange sourceRange, rngTag, mode
            End If
        End If
        Exit Sub
    End If

    columnCount = UBound(values, 2)
    columnLower = LBound(values, 2)

    For rowIndex = LBound(values, 1) To UBound(values, 1)
        rngName = Trim$(CStr(values(rowIndex, columnLower + 1)))
        If columnCount >= 3 Then status = Trim$(CStr(values(rowIndex, columnLower + 2))) Else status = vbNullString
        If columnCount >= 4 Then mode = Trim$(CStr(values(rowIndex, columnLower + 3))) Else mode = vbNullString

        If ShouldProcessRange(status, updateSequence) And LenB(rngName) > 0 Then
            TrackProcessedRange rngName
            Set sourceRange = ResolveNamedRange(registryTable.Parent.Parent, rngName)
            If Not sourceRange Is Nothing Then
                rngTag = rngName & TAG_SEPARATOR & CStr(updateSequence)
                ProcessSourceRange sourceRange, rngTag, mode
            End If
        End If
    Next rowIndex
End Sub

'@label:should-process-range
'@sub-title Determine whether the named range should be processed this cycle.
Private Function ShouldProcessRange(ByVal status As String, ByVal updateSequence As Long) As Boolean
    If updateSequence = 1 Then
        ShouldProcessRange = True
    Else
        ShouldProcessRange = (StrComp(status, STATUS_UPDATED, vbTextCompare) = 0)
    End If
End Function

'@label:process-source-range
'@sub-title Dispatch processing based on the requested translation mode.
Private Sub ProcessSourceRange(ByVal sourceRange As Range, ByVal rngTag As String, ByVal mode As String)
    Select Case LCase$(Trim$(mode))
        Case MODE_FORMULA
            ProcessFormulaRange sourceRange, rngTag
        Case MODE_TEXT
            ProcessTextRange sourceRange, rngTag
        Case Else
            If (mode <> MODE_WATCH) Then
                ThrowProjectError ProjectError.InvalidArgument, "Unknown translation mode '" & mode & "' for range '" & rngTag & "'."
            End If
    End Select
End Sub

'@label:process-text-range
'@sub-title Add literal cell values to the translations table.
Private Sub ProcessTextRange(ByVal sourceRange As Range, ByVal rngTag As String)
    Dim values As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim chunk As String

    values = sourceRange.Value2
    If IsArray(values) Then
        firstRow = LBound(values, 1)
        lastRow = UBound(values, 1)
        firstCol = LBound(values, 2)
        lastCol = UBound(values, 2)
        For rowIndex = firstRow To lastRow
            For colIndex = firstCol To lastCol
                chunk = Trim$(CStr(values(rowIndex, colIndex)))
                If LenB(chunk) > 0 Then
                    AddChunk chunk, rngTag
                End If
            Next colIndex
        Next rowIndex
    Else
        chunk = Trim$(CStr(values))
        If LenB(chunk) > 0 Then
            AddChunk chunk, rngTag
        End If
    End If
End Sub

'@label:process-formula-range
'@sub-title Extract quoted portions of formulas and add them as chunks.
Private Sub ProcessFormulaRange(ByVal sourceRange As Range, ByVal rngTag As String)
    Dim formulas As Variant
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim tokens As BetterArray
    Dim counter As Long
    Dim chunk As String

    formulas = sourceRange.Formula
    If IsArray(formulas) Then
        firstRow = LBound(formulas, 1)
        lastRow = UBound(formulas, 1)
        firstCol = LBound(formulas, 2)
        lastCol = UBound(formulas, 2)
        For rowIndex = firstRow To lastRow
            For colIndex = firstCol To lastCol
                Set tokens = ExtractFormulaChunks(CStr(formulas(rowIndex, colIndex)))
                If Not tokens Is Nothing Then
                    For counter = tokens.LowerBound To tokens.UpperBound
                        chunk = Trim$(CStr(tokens.Item(counter)))
                        If LenB(chunk) > 0 Then
                            AddChunk chunk, rngTag
                        End If
                    Next counter
                End If
            Next colIndex
        Next rowIndex
    Else
        Set tokens = ExtractFormulaChunks(CStr(formulas))
        If Not tokens Is Nothing Then
            For counter = tokens.LowerBound To tokens.UpperBound
                chunk = Trim$(CStr(tokens.Item(counter)))
                If LenB(chunk) > 0 Then
                    AddChunk chunk, rngTag
                End If
            Next counter
        End If
    End If
End Sub

'@label:extract-formula-chunks
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
'@label:add-chunk
'@sub-title Add or update a translation chunk in the table.
Private Sub AddChunk(ByVal label As String, ByVal rngTag As String)
    Dim rowIndex As Long
    Dim trimmedLabel As String

    trimmedLabel = Trim$(label)
    If LenB(trimmedLabel) = 0 Then Exit Sub

    rowIndex = FindLabelRow(trimmedLabel)

    If rowIndex > 0 Then
        If rowIndex <= this.tagBufferSize Then
            this.tagBuffer(rowIndex) = rngTag
        End If
    Else
        AppendRow trimmedLabel, rngTag
    End If
End Sub

'@label:append-row
'@sub-title Append a new list row and populate the mandatory cells.
Private Function AppendRow(ByVal label As String, ByVal rngTag As String) As Long
    Dim newRow As ListRow

    Set newRow = this.table.ListRows.Add
    newRow.Range.Cells(1, LabelColumnIndex).Value = label

    this.tagBufferSize = this.tagBufferSize + 1
    ReDim Preserve this.tagBuffer(1 To this.tagBufferSize)
    this.tagBuffer(this.tagBufferSize) = rngTag

    this.labelCount = this.labelCount + 1
    ReDim Preserve this.labelValues(1 To this.labelCount)
    this.labelValues(this.labelCount) = label

    this.dirty = True
    AppendRow = newRow.Index
End Function


'@section Label buffer
'===============================================================================
'@label:init-label-buffer
'@sub-title Read the label column into a 1D in-memory buffer for lookups.
Private Sub InitLabelBuffer()
    Dim dataRange As Range
    Dim rawValues As Variant
    Dim idx As Long

    Set dataRange = LabelColumnData()

    If dataRange Is Nothing Then
        this.labelCount = 0
        ReDim this.labelValues(1 To 1)
        Exit Sub
    End If

    rawValues = dataRange.Value2

    If IsArray(rawValues) Then
        this.labelCount = UBound(rawValues, 1) - LBound(rawValues, 1) + 1
        ReDim this.labelValues(1 To this.labelCount)
        For idx = 1 To this.labelCount
            this.labelValues(idx) = Trim$(CStr(rawValues(LBound(rawValues, 1) - 1 + idx, LBound(rawValues, 2))))
        Next idx
    Else
        this.labelCount = 1
        ReDim this.labelValues(1 To 1)
        this.labelValues(1) = Trim$(CStr(rawValues))
    End If
End Sub

'@label:find-label-row
'@sub-title Find the row index for a label using a simple linear scan of the label buffer.
Private Function FindLabelRow(ByVal label As String) As Long
    Dim idx As Long

    For idx = 1 To this.labelCount
        If StrComp(CStr(this.labelValues(idx)), label, vbTextCompare) = 0 Then
            FindLabelRow = idx
            Exit Function
        End If
    Next idx
End Function


'@section Tag buffer
'===============================================================================
'@label:init-tag-buffer
'@sub-title Read the tag column into a 1D in-memory buffer for batch writes.
Private Sub InitTagBuffer()
    Dim tagRange As Range
    Dim rawTags As Variant
    Dim idx As Long

    Set tagRange = TagColumnDataRange()

    If tagRange Is Nothing Then
        this.tagBufferSize = 0
        ReDim this.tagBuffer(1 To 1)
        Exit Sub
    End If

    rawTags = tagRange.Value2

    If IsArray(rawTags) Then
        this.tagBufferSize = UBound(rawTags, 1) - LBound(rawTags, 1) + 1
        ReDim this.tagBuffer(1 To this.tagBufferSize)
        For idx = 1 To this.tagBufferSize
            this.tagBuffer(idx) = rawTags(LBound(rawTags, 1) - 1 + idx, LBound(rawTags, 2))
        Next idx
    Else
        this.tagBufferSize = 1
        ReDim this.tagBuffer(1 To 1)
        this.tagBuffer(1) = rawTags
    End If
End Sub

'@label:flush-tag-buffer
'@sub-title Write the accumulated tag buffer back to the tag column in one operation.
Private Sub FlushTagBuffer()
    Dim tagRange As Range
    Dim outArr() As Variant
    Dim idx As Long

    If this.tagBufferSize = 0 Then Exit Sub

    Set tagRange = TagColumnDataRange()
    If tagRange Is Nothing Then Exit Sub

    ReDim outArr(1 To this.tagBufferSize, 1 To 1)
    For idx = 1 To this.tagBufferSize
        outArr(idx, 1) = this.tagBuffer(idx)
    Next idx

    tagRange.Resize(this.tagBufferSize, 1).Value = outArr
End Sub

'@section Process tracking
'===============================================================================
'@label:track-processed-range
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

'@label:was-range-processed
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

'@label:reset-processed-ranges
'@sub-title Reset processed ranges tracking.
Private Sub ResetProcessedRanges()
    Set this.processedRanges = Nothing
End Sub

'@label:remove-obsolete-labels
'@sub-title Remove labels whose source ranges were processed but whose tags are stale.
Private Sub RemoveObsoleteLabels(ByVal updateSequence As Long)
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim tagText As String
    Dim rangeName As String
    Dim sequenceValue As Long
    Dim tagRange As Range
    Dim tagValues As Variant
    Dim hasArray As Boolean
    Dim firstTagRow As Long
    Dim tagColumnIndex As Long

    rowCount = this.table.ListRows.Count
    If rowCount = 0 Then Exit Sub

    'Batch-read all tag values in one operation
    Set tagRange = TagColumnDataRange()
    If tagRange Is Nothing Then Exit Sub

    tagValues = tagRange.Value2
    hasArray = IsArray(tagValues)
    If hasArray Then
        firstTagRow = LBound(tagValues, 1)
        tagColumnIndex = LBound(tagValues, 2)
    End If

    'Walk rows from bottom to top so that deletions do not shift indices
    For rowIndex = rowCount To 1 Step -1
        If hasArray Then
            tagText = Trim$(CStr(tagValues(firstTagRow - 1 + rowIndex, tagColumnIndex)))
        Else
            tagText = Trim$(CStr(tagValues))
        End If

        'Delete when: (a) tag is empty, or (b) range was processed and sequence is old
        If LenB(tagText) = 0 Then
            this.table.ListRows(rowIndex).Delete
            this.dirty = True
        ElseIf ParseTag(tagText, rangeName, sequenceValue) Then
            If WasRangeProcessed(rangeName) And sequenceValue < updateSequence Then
                this.table.ListRows(rowIndex).Delete
                this.dirty = True
            End If
        End If
    Next rowIndex
End Sub

'@label:deduplicate-labels
'@sub-title Remove duplicate labels keeping the row with the oldest (lowest) tag sequence.
Private Sub DeduplicateLabels()
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim firstLangIndex As Long
    Dim labelData As Variant
    Dim tagData As Variant
    Dim tagRange As Range
    Dim labelValue As String
    Dim tagText As String
    Dim rangeName As String
    Dim seqValue As Long

    'Parallel arrays: seenLabels(i), bestRows(i), bestSeqs(i) form a single record
    Dim seenLabels As BetterArray
    Dim bestRows As BetterArray
    Dim bestSeqs As BetterArray
    Dim foundIdx As Long
    Dim searchIdx As Long

    'Collect row indices to delete, then delete from bottom to top
    Dim indicesToDelete() As Long
    Dim deleteCount As Long
    Dim prevRow As Long
    Dim idx As Long

    firstLangIndex = LabelColumnIndex()
    rowCount = this.table.ListRows.Count
    If rowCount < 2 Then Exit Sub
    If firstLangIndex > this.table.ListColumns.Count Then Exit Sub

    'Batch-read label and tag data
    On Error Resume Next
        labelData = this.table.ListColumns(firstLangIndex).DataBodyRange.Value2
        Set tagRange = TagColumnDataRange()
        If Not tagRange Is Nothing Then tagData = tagRange.Value2
    On Error GoTo 0

    If IsEmpty(labelData) Then Exit Sub

    Set seenLabels = New BetterArray
    seenLabels.LowerBound = 0
    Set bestRows = New BetterArray
    bestRows.LowerBound = 0
    Set bestSeqs = New BetterArray
    bestSeqs.LowerBound = 0

    deleteCount = 0
    ReDim indicesToDelete(1 To rowCount)

    For rowIndex = 1 To rowCount
        'Read label
        If IsArray(labelData) Then
            labelValue = Trim$(CStr(labelData(rowIndex, 1)))
        Else
            labelValue = Trim$(CStr(labelData))
        End If
        If LenB(labelValue) = 0 Then GoTo NextDedupRow

        'Read tag sequence number
        seqValue = 0
        tagText = vbNullString
        If IsArray(tagData) Then
            tagText = Trim$(CStr(tagData(rowIndex, 1)))
        ElseIf Not IsEmpty(tagData) Then
            tagText = Trim$(CStr(tagData))
        End If
        If LenB(tagText) > 0 Then ParseTag tagText, rangeName, seqValue

        'Search for this label in the already-seen list (case-sensitive)
        foundIdx = -1
        For searchIdx = seenLabels.LowerBound To seenLabels.UpperBound
            If StrComp(CStr(seenLabels.Item(searchIdx)), labelValue, vbBinaryCompare) = 0 Then
                foundIdx = searchIdx
                Exit For
            End If
        Next searchIdx

        If foundIdx >= 0 Then
            'Duplicate: keep the row with the older (lower) sequence number
            If seqValue < CLng(bestSeqs.Item(foundIdx)) Then
                'Current row is older — mark previous best row for deletion
                prevRow = CLng(bestRows.Item(foundIdx))
                deleteCount = deleteCount + 1
                indicesToDelete(deleteCount) = prevRow
                bestRows.Item(foundIdx) = rowIndex
                bestSeqs.Item(foundIdx) = seqValue
            Else
                'Stored row is older or equal — mark current row for deletion
                deleteCount = deleteCount + 1
                indicesToDelete(deleteCount) = rowIndex
            End If
        Else
            'First time seeing this label
            seenLabels.Push labelValue
            bestRows.Push rowIndex
            bestSeqs.Push seqValue
        End If

NextDedupRow:
    Next rowIndex

    'Sort indices descending so deletions do not shift remaining indices
    If deleteCount > 0 Then
        SortIndicesDescending indicesToDelete, deleteCount
        For idx = 1 To deleteCount
            this.table.ListRows(indicesToDelete(idx)).Delete
        Next idx
        this.dirty = True
    End If
End Sub

'@label:sort-indices-descending
'@sub-title Sort a 1-based Long array in descending order (simple insertion sort).
Private Sub SortIndicesDescending(ByRef arr() As Long, ByVal count As Long)
    Dim idx As Long
    Dim jdx As Long
    Dim temp As Long

    For idx = 2 To count
        temp = arr(idx)
        jdx = idx - 1
        Do While jdx >= 1
            If arr(jdx) >= temp Then Exit Do
            arr(jdx + 1) = arr(jdx)
            jdx = jdx - 1
        Loop
        arr(jdx + 1) = temp
    Next idx
End Sub

'@label:parse-tag
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

'@label:create-buffer
'@sub-title Create a BetterArray buffer with deterministic bounds.
Private Function CreateBuffer() As BetterArray
    Dim buffer As BetterArray
    Set buffer = New BetterArray
    buffer.LowerBound = 0
    Set CreateBuffer = buffer
End Function

'@label:build-duplicate-line
'@sub-title Format a single duplicate-count line for reporting.
Private Function BuildDuplicateLine(ByVal labelValue As String, ByVal occurrenceCount As Long) As String
    BuildDuplicateLine = """" & labelValue & """ has " & CStr(occurrenceCount) & " duplicates"
End Function


'@section Reporting
'===============================================================================
'@label:missing-labels
'@sub-title Count missing translation labels for a specific language.
Private Function MissingLabels(ByVal lang As String) As Long
    Dim trimmedLanguage As String
    Dim columnIndex As Long
    Dim dataRange As Range

    ValidateInitialised

    trimmedLanguage = Trim$(lang)
    If LenB(trimmedLanguage) = 0 Then
        ThrowProjectError ProjectError.InvalidArgument, "Please supply a language name."
    End If

    columnIndex = LanguageColumnIndex(trimmedLanguage)
    If columnIndex = 0 Then
        ThrowProjectError ProjectError.InvalidArgument, "Language '" & trimmedLanguage & "' does not exist in the translations table."
    End If

    On Error Resume Next
        Set dataRange = this.table.ListColumns(columnIndex).DataBodyRange
    On Error GoTo 0

    If dataRange Is Nothing Then Exit Function

    MissingLabels = Application.WorksheetFunction.CountBlank(dataRange)
End Function

'@label:number-of-missing
'@sub-title Summarise missing translations per language and optionally prompt the user.
Private Function NumberOfMissing() As String
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
    summary = messageLines.ToString(Separator:=vbLf, OpeningDelimiter:=vbNullString, ClosingDelimiter:=vbNullString)
    UpdateMissingSummary messageLines
    NumberOfMissing = summary

    If this.displayPrompts Then
        MsgBox summary, vbInformation + vbOKOnly, "Done!"
    End If
End Function

'@label:update-missing-summary
'@sub-title Update internal missing summary.
Private Sub UpdateMissingSummary(ByVal summary As BetterArray)

    If this.missingSummary Is Nothing Then
        Set this.missingSummary = CreateBuffer()
    End If

    Set this.missingSummary = summary.Clone()
End Sub

'@label:internal-missing-summary
'@prop-title Cached missing-labels summary across languages.
Private Property Get InternalMissingSummary() As BetterArray
    If this.missingSummary Is Nothing Then
        Set this.missingSummary = CreateBuffer()
    End If
    Set InternalMissingSummary = this.missingSummary.Clone()
End Property

'@label:update-duplicates-summary
'@sub-title Build duplicate summary for every language column.
Private Sub UpdateDuplicatesSummary()
    Dim langs As BetterArray
    Dim langIdx As Long
    Dim targetLang As String
    Dim summary As BetterArray
    Dim duplicateReport As TDuplicateSummary

    If this.duplicatesSummary Is Nothing Then
        Set this.duplicatesSummary = CreateBuffer()
    End If

    Set summary = CreateBuffer()
    Set langs = Languages(True)

    If Not (langs Is Nothing) Then
        For langIdx = langs.LowerBound To langs.UpperBound
            targetLang = Trim$(CStr(langs.Item(langIdx)))
            If LenB(targetLang) > 0 Then
                duplicateReport = ComputeDuplicateSummary(targetLang)
                If duplicateReport.HasDuplicates Then
                    summary.Push duplicateReport.Message
                End If
            End If
        Next langIdx
    End If

    Set this.duplicatesSummary = summary.Clone()
End Sub


'@label:internal-duplicates-summary
'@prop-title Cached duplicates summary across languages.
Private Property Get InternalDuplicatesSummary() As BetterArray
    If this.duplicatesSummary Is Nothing Then
        Set this.duplicatesSummary = CreateBuffer()
    End If

    Set InternalDuplicatesSummary = this.duplicatesSummary.Clone()
End Property


'@label:duplicate-labels
'@sub-title Report duplicate translation labels when they exist.
Private Function DuplicateLabels(ByRef duplicateMessage As String, Optional ByVal languageName As String = vbNullString) As Boolean
    Dim summary As TDuplicateSummary

    summary = ComputeDuplicateSummary(languageName)
    duplicateMessage = summary.Message
    DuplicateLabels = summary.HasDuplicates
End Function

'@label:compute-duplicate-summary
'@sub-title Compute duplicates for one language or all languages using simple linear counting.
Private Function ComputeDuplicateSummary(ByVal languageName As String) As TDuplicateSummary
    Dim labelRange As Range
    Dim targetRange As Range
    Dim rowIndex As Long
    Dim currentLabel As String
    Dim summary As TDuplicateSummary
    Dim columnTitle As String
    Dim messageLines As BetterArray
    Dim duplicatesFound As Boolean
    Dim trimmedLanguage As String
    Dim values As Variant
    Dim rowCount As Long
    Dim seenLabels As BetterArray
    Dim seenCounts As BetterArray
    Dim foundIdx As Long
    Dim searchIdx As Long
    Dim bucketIndex As Long

    ValidateInitialised
    Set labelRange = LabelColumnData()
    If labelRange Is Nothing Then GoTo Done

    trimmedLanguage = Trim$(languageName)
    If LenB(trimmedLanguage) = 0 Then
        Set targetRange = labelRange
        columnTitle = this.table.ListColumns(LabelColumnIndex).Name
    Else
        Set targetRange = ResolveLanguageColumn(trimmedLanguage)
        If targetRange Is Nothing Then GoTo Done
        columnTitle = trimmedLanguage
    End If

    'Batch read entire column
    values = targetRange.Value2
    If IsArray(values) Then
        rowCount = UBound(values, 1) - LBound(values, 1) + 1
    Else
        rowCount = 1
    End If

    Set seenLabels = New BetterArray
    seenLabels.LowerBound = 0
    Set seenCounts = New BetterArray
    seenCounts.LowerBound = 0
    Set messageLines = CreateBuffer()

    'Count occurrences using simple linear search
    For rowIndex = 1 To rowCount
        If IsArray(values) Then
            currentLabel = Trim$(CStr(values(LBound(values, 1) - 1 + rowIndex, LBound(values, 2))))
        Else
            currentLabel = Trim$(CStr(values))
        End If
        If LenB(currentLabel) = 0 Then GoTo ContinueLoop

        'Linear search for existing entry
        foundIdx = -1
        For searchIdx = seenLabels.LowerBound To seenLabels.UpperBound
            If StrComp(CStr(seenLabels.Item(searchIdx)), currentLabel, vbBinaryCompare) = 0 Then
                foundIdx = searchIdx
                Exit For
            End If
        Next searchIdx

        If foundIdx >= 0 Then
            'Existing label: increment count
            seenCounts.Item(foundIdx) = CLng(seenCounts.Item(foundIdx)) + 1
        Else
            'New label: add to lists
            seenLabels.Push currentLabel
            seenCounts.Push CLng(1)
        End If

ContinueLoop:
    Next rowIndex

    For bucketIndex = seenLabels.LowerBound To seenLabels.UpperBound
        If seenCounts.Item(bucketIndex) > 1 Then
            If Not duplicatesFound Then
                messageLines.Push "Duplicate labels detected in column " & columnTitle & "!"
                duplicatesFound = True
            End If
            messageLines.Push BuildDuplicateLine(CStr(seenLabels.Item(bucketIndex)), CLng(seenCounts.Item(bucketIndex)))
        End If
    Next bucketIndex

    If duplicatesFound Then
        summary.HasDuplicates = True
        summary.Message = messageLines.ToString(Separator:=vbLf, OpeningDelimiter:=vbNullString, ClosingDelimiter:=vbNullString)
    Else
        summary.HasDuplicates = False
        summary.Message = vbNullString
    End If

Done:
    ComputeDuplicateSummary = summary
End Function

'@label:label-column-data
'@sub-title Access the label column data range.
Private Function LabelColumnData() As Range
    If this.table.ListRows.Count = 0 Then Exit Function
    On Error Resume Next
        Set LabelColumnData = this.table.ListColumns(LabelColumnIndex).DataBodyRange
    On Error GoTo 0
End Function


'@section Tag column infrastructure
'===============================================================================
'@label:tag-column-data-range
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

'@label:apply-tag-column-formatting
'@sub-title Ensure the helper column keeps its stealth formatting regardless of state.
Private Sub ApplyTagColumnFormatting()
    Dim header As Range
    Dim tagData As Range

    On Error Resume Next
        Set header = RequireTagHeader()
    On Error GoTo 0

    If Not header Is Nothing Then
        On Error Resume Next
            header.Value = TAG_HEADER_TITLE
            header.Font.Color = vbWhite
            header.Interior.Color = vbWhite
        On Error GoTo 0
    End If

    Set tagData = TagColumnDataRange()
    If Not tagData Is Nothing Then
        On Error Resume Next
            tagData.Font.Color = vbWhite
            With tagData
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .Color = vbWhite
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .Color = vbWhite
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .Color = vbWhite
                End With
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .Color = vbWhite
                End With
                With .Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .Color = vbWhite
                End With
                With .Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .Color = vbWhite
                End With
        End With
        On Error GoTo 0
    End If
End Sub

'@label:apply-lang-column-formatting
'@sub-title Apply font, sizing, and duplicate highlighting to a language column.
Private Sub ApplyLangColumnFormatting(ByVal langRng As Range)

    If langRng Is Nothing Then Exit Sub

    With langRng

        .Cells(0, 1).Locked = False 'Unlock the first cell for the language
        .Cells.Font.Name = "Consolas"
        .Cells.Font.Size = 16
        .Cells(0, 1).Font.Size = 18
        .Cells.WrapText = True
        .Cells.EntireRow.AutoFit
        .Cells.HorizontalAlignment = xlHAlignLeft
        .Cells.VerticalAlignment = xlVAlignCenter

    End With

    'Formatting conditions
    langRng.FormatConditions.AddUniqueValues

    With langRng.FormatConditions(langRng.FormatConditions.Count)
        .DupeUnique = xlDuplicate
        .Interior.Color = vbRed
    End With

End Sub

'@label:configure-tag-column
'@sub-title Ensure the helper column header is located and configured.
Private Sub ConfigureTagColumn()
    Dim header As Range
    Dim helperHeader As Range
    Dim currentTitle As String

    ValidateInitialised

    Set header = this.table.HeaderRowRange.Cells(1, 1)
    If header.Column = 1 Then
        ThrowProjectError ProjectError.ErrorUnexpectedState, "Translations table requires a helper column immediately to the left."
    End If

    Set helperHeader = header.Offset(0, -1)
    Set this.tagHeader = helperHeader

    On Error Resume Next
        currentTitle = CStr(helperHeader.Value)
    On Error GoTo 0

    If StrComp(currentTitle, TAG_HEADER_TITLE, vbBinaryCompare) <> 0 Then
        On Error Resume Next
        helperHeader.Value = TAG_HEADER_TITLE
        On Error GoTo 0
    End If

    ApplyTagColumnFormatting
End Sub

'@label:ensure-tag-column-ready
'@sub-title Verify the helper tag column is accessible.
Private Sub EnsureTagColumnReady()
    Dim unused As Range
    Set unused = RequireTagHeader()
End Sub

'@label:require-tag-header
'@sub-title Return the tag header cell, raising an error when unavailable.
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

'@label:attach-tag-header
'@sub-title Attempt to reconfigure and attach the tag header cell.
Private Sub AttachTagHeader()
    On Error GoTo Failed
        ConfigureTagColumn
        Exit Sub
Failed:
    Err.Clear
    Set this.tagHeader = Nothing
End Sub

'@label:is-range-valid
'@sub-title Test whether a Range reference is still alive and readable.
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
'@label:begin-tag-integration
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

    On Error Resume Next
        this.table.ListColumns(1).Name = TAG_HEADER_TITLE
    On Error GoTo 0

    ApplyTagColumnFormatting
End Sub

'@label:end-tag-integration
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
    ApplyTagColumnFormatting
End Sub

'@label:label-column-index
'@sub-title Resolve the current column index storing translation labels.
Private Function LabelColumnIndex() As Long
    If this.tagsIntegrated Then
        LabelColumnIndex = 2
    Else
        LabelColumnIndex = 1
    End If
End Function


'@section Sorting and formatting
'===============================================================================
'@label:sort-translations
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

'@label:apply-formatting
'@sub-title Apply light formatting and duplicate highlighting once per update.
Private Sub ApplyFormatting()
    Dim dataRange As Range
    Dim tableRange As Range
    Dim columnRange As Range
    Dim unlockRange As Range
    Dim unlockOffset As Long
    Dim unlockWidth As Long


    Set dataRange = this.table.DataBodyRange
    Set tableRange = this.table.Range

    On Error Resume Next
        this.table.TableStyle = "TableStyleLight11"
        tableRange.EntireColumn.ColumnWidth = 60
        tableRange.Rows(1).Font.Color = vbWhite
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
        ApplyLangColumnFormatting columnRange
        Set columnRange = columnRange.Offset(0, 1)
        If columnRange.Column > dataRange.Columns(dataRange.Columns.Count).Column Then Exit Do
    Loop

    ApplyTagColumnFormatting
End Sub


'@section Data exchange
'===============================================================================
'@label:export
'@sub-title Export translations into a destination workbook.
Public Sub Export(ByVal exportWorkbook As Workbook, Optional ByVal Hide As Long = xlSheetVisible)
    Dim exportSheet As Worksheet
    Dim tableAdapter As ICustomTable
    Dim hostSheet As Worksheet
    Dim store As IHiddenNames

    ValidateInitialised
    If exportWorkbook Is Nothing Then
        ThrowProjectError ProjectError.ObjectNotInitialized, "Target workbook is missing; cannot export translations."
    End If

    Set hostSheet = HostWorksheet

    'Test if the export workbook has the translation sheet name
    On Error Resume Next
        Set exportSheet = exportWorkbook.Worksheets(hostSheet.Name)
    On Error GoTo 0

    If (exportSheet Is Nothing) Then
        Set exportSheet = exportWorkbook.Worksheets.Add(After:=exportWorkbook.Worksheets(exportWorkbook.Worksheets.Count))
        exportSheet.Name = hostSheet.Name
    Else
       exportSheet.Cells.Clear
    End If

    Set tableAdapter = CustomTable.Create(this.table)

    'Start exporting at column two to preserve the helper tag column offset.
    tableAdapter.Export exportSheet, startColumn:=2, startLine:=1, addListObject:=True

    Set store = MetadataStore()
    'Persist translation metadata (languages, formatting markers) alongside the exported table.
    If Not store Is Nothing Then store.ExportNames exportSheet

    On Error Resume Next
        exportSheet.Rows(1).Font.Bold = True
        exportSheet.Cells.Font.Name = "Consolas"
        exportSheet.Cells.HorizontalAlignment = xlHAlignLeft
        exportSheet.Cells.VerticalAlignment = xlVAlignCenter
        exportSheet.Cells.EntireRow.AutoFit
        exportSheet.Cells.EntireColumn.AutoFit
        exportSheet.Visible = Hide
    On Error GoTo 0
End Sub


'@section Registry bookkeeping
'===============================================================================
'@label:next-update-sequence
'@sub-title Retrieve and increment the registry update sequence.
Private Function NextUpdateSequence(ByVal registrySheet As Worksheet) As Long
    Dim currentValue As Long

    currentValue = CounterNumericValue(registrySheet)
    If currentValue < 0 Then currentValue = 0

    NextUpdateSequence = currentValue + 1
    PersistCounterValue registrySheet, NextUpdateSequence
End Function

'@label:persist-counter-value
'@sub-title Persist the counter value inside a hidden worksheet-level name.
Private Sub PersistCounterValue(ByVal registrySheet As Worksheet, ByVal counterValue As Long)
    Dim store As IHiddenNames

    If counterValue < 0 Then counterValue = 0

    Set store = HiddenNameStore(registrySheet)
    If store Is Nothing Then Exit Sub

    store.EnsureName COUNTER_NAME, counterValue, HiddenNameTypeLong
    store.SetValue COUNTER_NAME, counterValue
End Sub

'@label:counter-numeric-value
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

'@label:hidden-name-store
'@sub-title Create a HiddenNames wrapper for the supplied registry sheet.
Private Function HiddenNameStore(ByVal registrySheet As Worksheet) As IHiddenNames
    If registrySheet Is Nothing Then Exit Function
    On Error Resume Next
        Set HiddenNameStore = HiddenNames.Create(registrySheet)
    On Error GoTo 0
End Function

'@label:metadata-store
'@sub-title Resolve the hidden name store hosting metadata.
Private Function MetadataStore() As IHiddenNames
    Dim hostSh As Worksheet

    If this.table Is Nothing Then Exit Function
    Set hostSh = this.table.Parent
    If hostSh Is Nothing Then Exit Function

    On Error Resume Next
        Set MetadataStore = HiddenNames.Create(hostSh)
    On Error GoTo 0
End Function


'@section Validation and utilities
'===============================================================================
'@label:validate-initialised
'@sub-title Ensure the manager has been initialised before use.
Private Sub ValidateInitialised()
    If this.table Is Nothing Then
        ThrowProjectError ProjectError.ObjectNotInitialized, "Translations table was not initialised."
    End If
End Sub

'@label:validate-translation-table
'@sub-title Validate the supplied translations table.
Private Sub ValidateTranslationTable(ByVal translationTable As ListObject)
    If translationTable Is Nothing Then
        ThrowProjectError ProjectError.InvalidArgument, "Please supply a valid translations listobject."
    End If
End Sub

'@label:validate-registry-sheet
'@sub-title Validate the registry sheet before processing.
Private Sub ValidateRegistrySheet(ByVal registrySheet As Worksheet)
    If registrySheet Is Nothing Then
        ThrowProjectError ProjectError.InvalidArgument, "Registry worksheet cannot be Nothing."
    End If
End Sub

'@label:resolve-named-range
'@sub-title Resolve a named range from the workbook, returning Nothing when missing.
Private Function ResolveNamedRange(ByVal hostwb As Workbook, ByVal rangeName As String) As Range
    Dim definedName As Name
    On Error Resume Next
        Set definedName = hostwb.Names(rangeName)
    On Error GoTo 0

    If definedName Is Nothing Then Exit Function

    On Error Resume Next
        Set ResolveNamedRange = definedName.RefersToRange
    On Error GoTo 0
End Function

'@label:throw-project-error
'@sub-title Raise a project error with a contextual message.
Private Sub ThrowProjectError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub


'@section Interface implementation
'===============================================================================
'@description
'Thin delegation stubs wiring ISetupTranslationsTable members to
'the corresponding public or private methods on this class.

Private Sub ISetupTranslationsTable_EnsureLanguages(Optional ByVal languages As String = vbNullString)
    EnsureLanguages languages
End Sub

Private Sub ISetupTranslationsTable_SwitchDefaultLanguage(ByVal languageName As String)
    SwitchDefaultLanguage languageName
End Sub

Private Function ISetupTranslationsTable_SwitchableLanguages() As BetterArray
    Set ISetupTranslationsTable_SwitchableLanguages = SwitchableLanguages
End Function

Private Property Get ISetupTranslationsTable_Languages(Optional ByVal includeDefault As Boolean = False) As BetterArray
    Set ISetupTranslationsTable_Languages = Languages(includeDefault)
End Property

Private Property Get ISetupTranslationsTable_HostWorksheet() As Worksheet
    Set ISetupTranslationsTable_HostWorksheet = HostWorksheet
End Property

'@label:initialise
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

Private Sub ISetupTranslationsTable_Export(ByVal exportWorkbook As Workbook, Optional ByVal Hide As Long = xlSheetVisible)
    Export exportWorkbook, Hide
End Sub

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

Private Function ISetupTranslationsTable_MissingLabels(ByVal lang As String) As Long
    ISetupTranslationsTable_MissingLabels = MissingLabels(lang)
End Function

Private Function ISetupTranslationsTable_DuplicateLabels(ByRef duplicateMessage As String, Optional ByVal languageName As String = vbNullString) As Boolean
    ISetupTranslationsTable_DuplicateLabels = DuplicateLabels(duplicateMessage, languageName)
End Function

Private Property Get ISetupTranslationsTable_InternalMissingSummary() As BetterArray
    Set ISetupTranslationsTable_InternalMissingSummary = InternalMissingSummary
End Property

Private Property Get ISetupTranslationsTable_InternalDuplicatesSummary() As BetterArray
    Set ISetupTranslationsTable_InternalDuplicatesSummary = InternalDuplicatesSummary
End Property
