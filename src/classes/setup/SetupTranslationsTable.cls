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
'@depends ProjectError, ApplicationState, HiddenNames, CustomTable, BetterArray

Option Explicit

'@class SetupTranslationsTable
'@description
'Manages the translations ListObject of a setup workbook, together with the
'helper tag column that sits beside it. UpdateFromRegistry reads the watched
'ranges of the update registry and writes one row per label. EnsureLanguages
'creates the language columns, SwitchDefaultLanguage promotes one of them,
'SwitchableLanguages and Languages list what is there, and ResetSequence puts
'the row counter back. MissingLabels, NumberOfMissing, DuplicateLabels and
'UnresolvedRanges report on the table, and Export writes it into another
'workbook.
'
'LABEL MATCHING POLICY
'-------------------------------------------------------------------------------
'A label carries its case. "Hello" and "hello" are two labels, they take two
'rows, and each one stands for itself. FindLabelRow, DeduplicateLabels and
'ComputeDuplicateSummary all answer from one keyed Collection built by
'LabelKey, and that key spells the capitalisation out so two spellings never
'share it. A Collection matches its own keys without regard to case, so the
'case travels inside the key. TranslationObject.LookupRow reads this table at
'run time and compares with vbBinaryCompare.
'
'Column titles, registry status words, mode keywords and defined-name lookups
'are matched without regard to case. Excel matches defined names that way
'itself.
'
'DUPLICATES ARE COLOURED BY GROUP
'-------------------------------------------------------------------------------
'ApplyFormatting gives every duplicate group in a column a fill of its own,
'walked off the hue wheel by DuplicateGroupColor. A column painted in one colour
'says which cells are duplicated and nothing about which of them go together,
'and a translator staring at forty red cells has to read every one to find the
'pairs. One rule per group answers both questions at a glance.
'
'CollectDuplicateGroups finds the groups, and the summary message reads its
'groups from the same function, so the sheet and the message never disagree. The
'rules are EXACT tests, which is what carries the capitalisation through to the
'sheet: "Hello" and "hello" are two labels here and neither takes the colour of
'the other.
'
'The red rule that used to paint every duplicate is still there, added last so
'it holds the lowest priority. It now marks a duplicate this class could not
'group: one past the end of the palette, one whose label is too long for a rule
'formula, or one whose formula Excel refused.
'
'THE HELPER TAG COLUMN
'-------------------------------------------------------------------------------
'The translations ListObject sits with one spare worksheet column to its left
'holding a tag per row, shaped <rangename>__<sequence>. That column lives
'outside the ListObject most of the time. During an update the table is resized
'one column left so each tag travels with its row through the sort, then
'resized back. LabelColumnIndex() answers 2 while the tag column is inside the
'table and 1 otherwise, and every column index in this class comes from that
'one function.
'
'CREATE WRITES NOTHING
'-------------------------------------------------------------------------------
'Create locates and checks the helper column. The title, the formatting and the
'hidden language list are written by the update path alone, through
'EnsureTagColumnReady and EnsureLanguages, so a manager can be built over a
'protected sheet and read from.
'
'SOURCE RANGES ARE CONTIGUOUS
'-------------------------------------------------------------------------------
'A registry range refers to one block of cells on one worksheet.
'SetupPreparation and UpdatedValues both build their names over a single table
'column. ProcessTextRange and ProcessFormulaRange read each source in one
'crossing, and Range.Value2 answers for the first area of a range, so a name
'built by hand over scattered cells brings in its first block alone.
'
'ERROR NUMBERS
'-------------------------------------------------------------------------------
'ProjectError.ErrorUnexpectedState is 1004, which is also Excel's generic error
'number. An Excel 1004 that escapes an update is raised again as 1004, so a
'caller reading the number alone sees one value for both.

'@section State and constants
'===============================================================================
Private Type TDuplicateSummary
    HasDuplicates As Boolean
    Message As String
End Type

Private Type TSetupTranslations
    table As ListObject               'Managed translations table
    tagHeader As Range                'Header cell immediately to the left storing translation tags
    labelValues() As Variant          'In-memory label column for lookups during ProcessRegistry
    labelCount As Long                'Number of entries in the label buffer
    labelCapacity As Long             'Allocated size of the label buffer
    labelIndex As Collection          'Case-carrying label key -> row index, rebuilt each cycle
    pendingCount As Long              'Labels appended in memory and not yet written to the sheet
    processedRanges As BetterArray    'Ranges processed during the current update cycle
    unresolvedRanges As BetterArray   'Registry ranges that failed to resolve during the cycle
    displayPrompts As Boolean         'Toggle to avoid message prompts during automated tests
    stateScope As ApplicationState    'Optional application guard for busy sections
    tagsIntegrated As Boolean         'Tracks whether the helper tag column is currently merged in the table
    missingSummary As BetterArray     'Keep missing summary on non default languages
    duplicatesSummary As BetterArray  'keep duplicate summary on all the languages
    tagBuffer() As Variant            'In-memory tag column buffer for batch writes during ProcessRegistry
    tagBufferSize As Long             'Current number of entries in the tag buffer
    tagCapacity As Long               'Allocated size of the tag buffer
    metadataStore As HiddenNames      'Cached hidden name store on the host worksheet
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
Private Const KEY_CASE_SEPARATOR As String = "|"
Private Const LANGUAGES_NAME_ID As String = "__SetupTranslationsLanguages__"

'Registry column headers, as UpdatedValues writes them. Each column is found
'by header name first and falls back to its published position, so an older
'registry with different header text still loads.
Private Const REGISTRY_HEADER_RANGE As String = "rngname"
Private Const REGISTRY_HEADER_STATUS As String = "updated"
Private Const REGISTRY_HEADER_MODE As String = "headername"
Private Const REGISTRY_COLUMN_RANGE As Long = 2
Private Const REGISTRY_COLUMN_STATUS As Long = 3
Private Const REGISTRY_COLUMN_MODE As Long = 4
Private Const REGISTRY_COLUMN_COUNT As Long = 4

'Tag given to a row that carries a label and no tag. Such a row was never
'claimed by a watched range, so no watched range can declare it obsolete.
'Sequence 0 is the lowest, so a hand-written row wins over a regenerated one
'when the two carry the same label.
Private Const MANUAL_TAG_NAME As String = "__manual__"

Private Const BUFFER_GROWTH_MARGIN As Long = 16
Private Const MAX_LANGUAGE_SCAN As Long = 64
Private Const MAX_DELETE_BATCH As Long = 500

'The duplicate palette. One hue per group, walked with a stride that is prime to
'the hue count so the whole wheel is still used and two groups that follow each
'other in the column land half a wheel apart. A lap of the wheel is one shade
'band, and each further lap drops the lightness, so the palette holds
'DUPLICATE_HUE_COUNT * DUPLICATE_SHADE_COUNT colours before it repeats.
Private Const DUPLICATE_HUE_COUNT As Long = 12
Private Const DUPLICATE_HUE_STRIDE As Long = 5
Private Const DUPLICATE_SHADE_COUNT As Long = 2
Private Const DUPLICATE_HUE_START As Double = 20
Private Const DUPLICATE_SATURATION As Double = 0.72
Private Const DUPLICATE_LIGHTNESS_TOP As Double = 0.86
Private Const DUPLICATE_LIGHTNESS_STEP As Double = 0.22
'DUPLICATE_HUE_COUNT hues times DUPLICATE_SHADE_COUNT shade bands.
Private Const MAX_DUPLICATE_GROUPS As Long = 24

'Excel refuses a format condition formula past 255 characters. A label that long
'gets no group rule and is left to the red catch-all.
Private Const MAX_RULE_FORMULA_LENGTH As Long = 255


'@section Factory helpers
'===============================================================================
'@label:create
'@sub-title Create a configured translations table manager instance.
'@param translationTable ListObject. The managed translations table.
'@param stateScope Optional ApplicationState. Guard for Excel settings during updates.
'@return SetupTranslationsTable. Ready to process registry updates.
'@export
Public Function Create(ByVal translationTable As ListObject, Optional ByVal stateScope As ApplicationState = Nothing) As SetupTranslationsTable
    Dim instance As SetupTranslationsTable

    ValidateTranslationTable translationTable

    Set instance = New SetupTranslationsTable
    instance.InitialiseCore translationTable, stateScope

    Set Create = instance
End Function

'@label:initialise-core
'@sub-title Store the supplied table and optional scope inside the current instance.
Public Sub InitialiseCore(ByVal translationTable As ListObject, Optional ByVal stateScope As ApplicationState = Nothing)
    With this
        Set .table = translationTable
        Set .stateScope = stateScope
        .displayPrompts = True
        .tagsIntegrated = False
    End With

    ConfigureTagColumn
    ResetProcessedRanges
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

'@label:set-display-prompts
'@sub-title Allow callers to toggle display prompts used during updates.
Public Sub SetDisplayPrompts(ByVal state As Boolean)
    this.displayPrompts = state
End Sub

'@label:languages-name-id
'@prop-title The HiddenName key of the persisted language list.
'@details
'Name of the worksheet-level HiddenName this class writes the semicolon
'separated language list under on its Translations sheet. The designer reads
'the list back through this key when a setup file is loaded, so the key lives
'here, on the class that writes it. The answer is a constant, so the property
'works on the predeclared instance.
'@return String. The HiddenName identifier.
'@export
Public Property Get LanguagesNameId() As String
    LanguagesNameId = LANGUAGES_NAME_ID
End Property

'@label:resolve-scope
'@sub-title Guarded access to the application scope.
Private Function ResolveScope() As ApplicationState
    If this.stateScope Is Nothing Then
        Set this.stateScope = ApplicationState.Create(Application)
    End If

    Set ResolveScope = this.stateScope
End Function


'@section Language management
'===============================================================================
'@label:ensure-languages
'@sub-title Ensure all provided languages exist as table headers.
Public Sub EnsureLanguages(Optional ByVal lang As String = vbNullString)
    ValidateInitialised

    Dim languageList() As String
    Dim counter As Long
    Dim languageName As String
    Dim languages As String
    Dim typedCells As Range

    If Not TryGetWrittenLanguages(languages, typedCells, lang) Then
        PersistLanguageList
        Exit Sub
    End If

    languageList = Split(languages, ";")
    For counter = LBound(languageList) To UBound(languageList)
        languageName = Trim$(languageList(counter))
        If LenB(languageName) > 0 Then
            EnsureLanguageColumn languageName
        End If
    Next counter

    'The typed cells are cleared once every column exists. When a column add
    'fails the text the user typed is still on the sheet.
    If Not typedCells Is Nothing Then typedCells.ClearContents

    PersistLanguageList
End Sub

'@label:try-get-written-languages
'@sub-title Collect language names written to the right of the table.
'@param languages ByRef String. Receives the semicolon separated language list.
'@param typedCells ByRef Range. Receives the cells holding the typed names, to clear later.
'@param newLang String. Extra language requested by the caller.
Private Function TryGetWrittenLanguages(ByRef languages As String, ByRef typedCells As Range, ByVal newLang As String) As Boolean

    Dim cellRng As Range
    Dim typedNames As BetterArray
    Dim scanned As Long
    Dim idx As Long
    Dim lastColumn As Long

    Set typedNames = CreateBuffer()
    Set typedCells = Nothing

    Set cellRng = this.table.ListColumns(this.table.ListColumns.Count).Range.Offset(, 1).Cells(1, 1)
    lastColumn = cellRng.Worksheet.Columns.Count

    Do While scanned < MAX_LANGUAGE_SCAN
        If LenB(cellRng.Value) = 0 Then Exit Do

        typedNames.Push SafeText(cellRng.Value)
        If typedCells Is Nothing Then
            Set typedCells = cellRng
        Else
            Set typedCells = Application.Union(typedCells, cellRng)
        End If

        scanned = scanned + 1
        If cellRng.Column >= lastColumn Then Exit Do
        Set cellRng = cellRng.Offset(, 1)
    Loop

    If (LenB(newLang) <> 0) Then languages = newLang

    If typedNames.Length > 0 Then
        For idx = typedNames.LowerBound To typedNames.UpperBound
            If (LenB(languages) <> 0) Then
                languages = languages & ";" & CStr(typedNames.Item(idx))
            Else
                languages = CStr(typedNames.Item(idx))
            End If
        Next idx
    End If

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

    this.table.ListColumns.Add.Name = languageName
End Sub

'@label:persist-language-list
'@sub-title Persist the available languages into the worksheet hidden names.
Private Sub PersistLanguageList()
    Dim store As HiddenNames
    Dim buffer As BetterArray
    Dim switchable As BetterArray
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
    Dim scope As ApplicationState

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

    SwapLanguageColumns LabelColumnIndex(), targetIndex
    SortTranslations
    ApplyFormatting
    PersistLanguageList

    scope.Restore
    Exit Sub

CleanFail:
    'The handler must not raise. A throw here would replace the real error and
    'leave Excel with screen updating off.
    On Error Resume Next
        If Not scope Is Nothing Then scope.Restore silent:=True
    On Error GoTo 0
    ThrowProjectError ProjectError.SomethingWentWrong, "Failed to switch default language: " & Err.Description
End Sub

'@label:remove-language
'@sub-title Delete one language column from the translations table.
'@details
'Only the columns after the default language are languages, so the structural
'columns can never be deleted through here. The default language column
'stays: promote another language first with SwitchDefaultLanguage. The stored
'language list is refreshed once the column is gone.
'@param languageName String. Header of the language column to delete.
Public Sub RemoveLanguage(ByVal languageName As String)
    Dim targetIndex As Long

    ValidateInitialised

    languageName = Trim$(languageName)
    If LenB(languageName) = 0 Then Exit Sub

    targetIndex = LanguageColumnIndex(languageName)
    If targetIndex = 0 Then
        ThrowProjectError ProjectError.InvalidArgument, "Language '" & languageName & "' does not exist in the translations table."
    End If

    If targetIndex = LabelColumnIndex() Then
        ThrowProjectError ProjectError.InvalidArgument, "Language '" & languageName & "' is the default language. Promote another language first."
    End If

    this.table.ListColumns(targetIndex).Delete
    PersistLanguageList
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
'@sub-title Swap the headers and the data between two language columns.
'@details The data body of each column is read and written as one block, so a
'         swap costs four Excel crossings whatever the row count. Number
'         formats stay where they are: translation columns hold text. The
'         translations table carries no totals row, so nothing here handles one.
Private Sub SwapLanguageColumns(ByVal firstIndex As Long, ByVal secondIndex As Long)
    Dim firstColumn As ListColumn
    Dim secondColumn As ListColumn
    Dim firstHeader As String
    Dim secondHeader As String
    Dim temporaryHeader As String
    Dim firstRange As Range
    Dim secondRange As Range
    Dim firstBlock As Variant
    Dim secondBlock As Variant
    Dim rowCount As Long

    Set firstColumn = this.table.ListColumns(firstIndex)
    Set secondColumn = this.table.ListColumns(secondIndex)

    firstHeader = firstColumn.Name
    secondHeader = secondColumn.Name

    On Error Resume Next
        Set firstRange = firstColumn.DataBodyRange
        Set secondRange = secondColumn.DataBodyRange
    On Error GoTo 0

    temporaryHeader = UniqueTemporaryHeader()
    secondColumn.Name = temporaryHeader
    firstColumn.Name = secondHeader
    secondColumn.Name = firstHeader

    If firstRange Is Nothing Or secondRange Is Nothing Then Exit Sub

    rowCount = firstRange.Rows.Count
    If secondRange.Rows.Count < rowCount Then
        rowCount = secondRange.Rows.Count
    End If
    If rowCount = 0 Then Exit Sub

    Set firstRange = firstRange.Resize(rowCount, 1)
    Set secondRange = secondRange.Resize(rowCount, 1)

    firstBlock = firstRange.Formula
    secondBlock = secondRange.Formula
    firstRange.Formula = secondBlock
    secondRange.Formula = firstBlock
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
'@details The sequence counter is written after the rows settle. A run that
'         fails halfway leaves the counter where it was, so the retry works at
'         the same sequence and the rows the failed run tagged survive.
Public Sub UpdateFromRegistry(ByVal registrySheet As Worksheet, _
                             Optional ByVal languages As String = vbNullString)
    Dim scope As ApplicationState
    Dim updateSequence As Long
    Dim errNumb As Long
    Dim errDesc As String

    ValidateInitialised
    ValidateRegistrySheet registrySheet

    On Error GoTo CleanFail

    EnsureLanguages languages
    EnsureTagColumnReady

    BeginTagIntegration
    Set scope = ResolveScope()
    scope.ApplyBusyState True

    updateSequence = PlannedUpdateSequence(registrySheet)
    ResetProcessedRanges
    this.dirty = False

    InitLabelBuffer
    InitTagBuffer
    ProcessRegistry registrySheet, updateSequence
    FlushPendingRows
    FlushTagBuffer

    RemoveObsoleteLabels updateSequence
    DeduplicateLabels
    PersistCounterValue registrySheet, updateSequence

    SortTranslations
    If this.dirty Then ApplyFormatting

    EndTagIntegration

    'The prompts run once the table is back in its resting shape, so the user
    'reads them over a table with no tag column showing.
    NumberOfMissing
    UpdateDuplicatesSummary
    ReleaseUpdateBuffers

    scope.Restore
    Exit Sub

CleanFail:
    errNumb = Err.Number
    errDesc = Err.Description
    'Both calls below must stay quiet. A throw inside the handler would replace
    'the real error and leave Excel with screen updating off and events
    'suppressed.
    On Error Resume Next
        EndTagIntegration
        If Not scope Is Nothing Then scope.Restore silent:=True
        ReleaseUpdateBuffers
    On Error GoTo 0
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
    Dim nameColumn As Long
    Dim statusColumn As Long
    Dim modeColumn As Long
    Dim hostWorkbook As Workbook

    Set dataRange = registryTable.DataBodyRange
    If dataRange Is Nothing Then Exit Sub

    nameColumn = RegistryColumnByHeader(registryTable, REGISTRY_HEADER_RANGE)
    columnCount = registryTable.ListColumns.Count

    'The sheet that holds the registries also holds the name index table
    'UpdatedValues keeps beside them. That one carries three columns and no
    'rngname header, so it is skipped. Reading it as a registry would take a
    'table name for a range name.
    If nameColumn = 0 And columnCount < REGISTRY_COLUMN_COUNT Then Exit Sub

    'Anything that names a range column, or is wide enough to be read by
    'position, has to carry the whole published shape.
    If columnCount < REGISTRY_COLUMN_COUNT Then
        ThrowRegistryShape registryTable, columnCount
    End If

    If nameColumn = 0 Then nameColumn = REGISTRY_COLUMN_RANGE
    statusColumn = RegistryColumn(registryTable, REGISTRY_HEADER_STATUS, REGISTRY_COLUMN_STATUS)
    modeColumn = RegistryColumn(registryTable, REGISTRY_HEADER_MODE, REGISTRY_COLUMN_MODE)

    values = dataRange.Value2
    If Not IsArray(values) Then Exit Sub

    columnLower = LBound(values, 2)
    Set hostWorkbook = registryTable.Parent.Parent

    For rowIndex = LBound(values, 1) To UBound(values, 1)
        rngName = SafeText(values(rowIndex, columnLower + nameColumn - 1))
        status = SafeText(values(rowIndex, columnLower + statusColumn - 1))
        mode = SafeText(values(rowIndex, columnLower + modeColumn - 1))

        If ShouldProcessRange(status, updateSequence) And LenB(rngName) > 0 Then
            Set sourceRange = ResolveNamedRange(hostWorkbook, rngName)
            If sourceRange Is Nothing Then
                'A range that does not resolve was not processed. Marking it
                'processed would let RemoveObsoleteLabels delete every row
                'tagged to it.
                TrackUnresolvedRange rngName
            Else
                TrackProcessedRange rngName
                rngTag = rngName & TAG_SEPARATOR & CStr(updateSequence)
                ProcessSourceRange sourceRange, rngTag, mode
            End If
        End If
    Next rowIndex
End Sub

'@label:registry-column
'@sub-title Resolve a registry column by header name, falling back to its published position.
Private Function RegistryColumn(ByVal registryTable As ListObject, _
                                ByVal headerName As String, _
                                ByVal fallbackColumn As Long) As Long
    RegistryColumn = RegistryColumnByHeader(registryTable, headerName)
    If RegistryColumn = 0 Then RegistryColumn = fallbackColumn
End Function

'@label:registry-column-by-header
'@sub-title Find a registry column by header name, answering 0 when the header is absent.
Private Function RegistryColumnByHeader(ByVal registryTable As ListObject, _
                                        ByVal headerName As String) As Long
    Dim column As ListColumn

    For Each column In registryTable.ListColumns
        If StrComp(Trim$(column.Name), headerName, vbTextCompare) = 0 Then
            RegistryColumnByHeader = column.Index
            Exit Function
        End If
    Next column
End Function

'@label:throw-registry-shape
'@sub-title Report a registry table that carries too few columns.
Private Sub ThrowRegistryShape(ByVal registryTable As ListObject, ByVal columnCount As Long)
    ThrowProjectError ProjectError.ErrorUnexpectedState, _
                      "Registry table '" & registryTable.Name & "' has " & CStr(columnCount) & _
                      " columns and " & CStr(REGISTRY_COLUMN_COUNT) & " are required."
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
'@details The mode is lowered and trimmed once, and every branch reads that one
'         value. A watched range and a blank mode both carry no labels, so both
'         are skipped without a word.
Private Sub ProcessSourceRange(ByVal sourceRange As Range, ByVal rngTag As String, ByVal mode As String)
    Dim requestedMode As String

    requestedMode = LCase$(Trim$(mode))

    Select Case requestedMode
        Case MODE_FORMULA
            ProcessFormulaRange sourceRange, rngTag
        Case MODE_TEXT
            ProcessTextRange sourceRange, rngTag
        Case MODE_WATCH, vbNullString
            'A watched range holds no labels of its own.
        Case Else
            ThrowProjectError ProjectError.InvalidArgument, "Unknown translation mode '" & mode & "' for range '" & rngTag & "'."
    End Select
End Sub

'@label:process-text-range
'@sub-title Add literal cell values to the translations table.
Private Sub ProcessTextRange(ByVal sourceRange As Range, ByVal rngTag As String)
    Dim values As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim readRange As Range
    Dim chunk As String

    Set readRange = EffectiveSourceRange(sourceRange)
    If readRange Is Nothing Then Exit Sub

    values = readRange.Value2
    If IsArray(values) Then
        For rowIndex = LBound(values, 1) To UBound(values, 1)
            For colIndex = LBound(values, 2) To UBound(values, 2)
                chunk = SafeText(values(rowIndex, colIndex))
                If LenB(chunk) > 0 Then
                    AddChunk chunk, rngTag
                End If
            Next colIndex
        Next rowIndex
    Else
        chunk = SafeText(values)
        If LenB(chunk) > 0 Then
            AddChunk chunk, rngTag
        End If
    End If
End Sub

'@label:process-formula-range
'@sub-title Extract quoted portions of formulas and add them as chunks.
Private Sub ProcessFormulaRange(ByVal sourceRange As Range, ByVal rngTag As String)
    Dim formulaValues As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim readRange As Range
    Dim tokens As BetterArray

    Set readRange = EffectiveSourceRange(sourceRange)
    If readRange Is Nothing Then Exit Sub

    formulaValues = readRange.Formula
    If IsArray(formulaValues) Then
        For rowIndex = LBound(formulaValues, 1) To UBound(formulaValues, 1)
            For colIndex = LBound(formulaValues, 2) To UBound(formulaValues, 2)
                Set tokens = ExtractFormulaChunks(SafeText(formulaValues(rowIndex, colIndex)))
                AddTokens tokens, rngTag
            Next colIndex
        Next rowIndex
    Else
        Set tokens = ExtractFormulaChunks(SafeText(formulaValues))
        AddTokens tokens, rngTag
    End If
End Sub

'@label:add-tokens
'@sub-title Add every extracted formula token as a chunk.
Private Sub AddTokens(ByVal tokens As BetterArray, ByVal rngTag As String)
    Dim counter As Long
    Dim chunk As String

    If tokens Is Nothing Then Exit Sub

    For counter = tokens.LowerBound To tokens.UpperBound
        chunk = SafeText(tokens.Item(counter))
        If LenB(chunk) > 0 Then
            AddChunk chunk, rngTag
        End If
    Next counter
End Sub

'@label:effective-source-range
'@sub-title Narrow a source range to the part of the sheet that holds data.
'@details A defined name over a whole column answers with a 1,048,576 row
'         array. The intersection with the used range keeps the read small.
Private Function EffectiveSourceRange(ByVal sourceRange As Range) As Range
    Dim parentSheet As Worksheet

    If sourceRange Is Nothing Then Exit Function

    On Error Resume Next
        Set parentSheet = sourceRange.Worksheet
    On Error GoTo 0

    If parentSheet Is Nothing Then
        Set EffectiveSourceRange = sourceRange
        Exit Function
    End If

    On Error Resume Next
        Set EffectiveSourceRange = Application.Intersect(sourceRange, parentSheet.UsedRange)
    On Error GoTo 0
End Function

'@label:extract-formula-chunks
'@sub-title Extract quoted text segments from a formula while handling escaped quotes.
'@details The walk jumps from quote to quote with InStr, so the cost is one
'         call per quote. Escaped double quotes are stripped first, which is
'         what makes IF(A1="", ...) skip its empty comparison string.
Private Function ExtractFormulaChunks(ByVal formulaText As String) As BetterArray
    Dim cleaned As String
    Dim openQuote As Long
    Dim closeQuote As Long
    Dim buffer As BetterArray
    Dim token As String

    cleaned = Replace(formulaText, Chr$(34) & Chr$(34), vbNullString)

    openQuote = InStr(1, cleaned, Chr$(34))
    Do While openQuote > 0
        closeQuote = InStr(openQuote + 1, cleaned, Chr$(34))
        If closeQuote = 0 Then Exit Do

        If closeQuote > openQuote + 1 Then
            token = Mid$(cleaned, openQuote + 1, closeQuote - openQuote - 1)
            If buffer Is Nothing Then Set buffer = CreateBuffer()
            buffer.Push token
        End If

        openQuote = InStr(closeQuote + 1, cleaned, Chr$(34))
    Loop

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

    If rowIndex = 0 Then
        AppendRow trimmedLabel, rngTag
        Exit Sub
    End If

    'The label buffer and the tag buffer grow together from the same row count,
    'so they can only disagree through a fault. Report it: a dropped tag turns
    'into a row with no tag, and the next run has to guess what that row is.
    If rowIndex > this.tagBufferSize Then
        ThrowProjectError ProjectError.ErrorUnexpectedState, _
                          "Label '" & trimmedLabel & "' maps to row " & CStr(rowIndex) & _
                          " and the tag buffer holds " & CStr(this.tagBufferSize) & " rows."
    End If

    this.tagBuffer(rowIndex) = rngTag
End Sub

'@label:append-row
'@sub-title Record a new label in memory, to be written to the sheet in one batch.
Private Sub AppendRow(ByVal label As String, ByVal rngTag As String)
    this.labelCount = this.labelCount + 1
    EnsureLabelCapacity this.labelCount
    this.labelValues(this.labelCount) = label
    RegisterLabelKey label, this.labelCount

    this.tagBufferSize = this.tagBufferSize + 1
    EnsureTagCapacity this.tagBufferSize
    this.tagBuffer(this.tagBufferSize) = rngTag

    this.pendingCount = this.pendingCount + 1
    this.dirty = True
End Sub

'@label:flush-pending-rows
'@sub-title Grow the table once and write every new label in one block.
Private Sub FlushPendingRows()
    Dim existingRows As Long
    Dim idx As Long
    Dim labelBlock() As Variant
    Dim labelColumn As Range

    If this.pendingCount = 0 Then Exit Sub

    existingRows = this.labelCount - this.pendingCount

    ReDim labelBlock(1 To this.pendingCount, 1 To 1)
    For idx = 1 To this.pendingCount
        labelBlock(idx, 1) = this.labelValues(existingRows + idx)
    Next idx

    GrowTableRows this.pendingCount

    Set labelColumn = this.table.ListColumns(LabelColumnIndex()).DataBodyRange
    labelColumn.Cells(existingRows + 1, 1).Resize(this.pendingCount, 1).Value = labelBlock

    this.pendingCount = 0
End Sub

'@label:grow-table-rows
'@sub-title Extend the managed table by the requested number of rows in one call.
Private Sub GrowTableRows(ByVal extraRows As Long)
    Dim tableRange As Range
    Dim below As Range

    If extraRows <= 0 Then Exit Sub

    Set tableRange = this.table.Range
    Set below = tableRange.Rows(tableRange.Rows.Count).Offset(1, 0).Resize(extraRows)

    'Anything sitting under the table is pushed down first. A plain Resize
    'grows the table over those cells and the label write that follows
    'overwrites whatever they held. Adding one ListRow at a time used to shift
    'them down, and this keeps that behaviour at two Excel calls.
    If Application.WorksheetFunction.CountA(below) > 0 Then
        below.Insert Shift:=xlDown
    End If

    this.table.Resize tableRange.Resize(tableRange.Rows.Count + extraRows, tableRange.Columns.Count)
End Sub

'@label:ensure-label-capacity
'@sub-title Grow the label buffer geometrically so appends stay linear.
Private Sub EnsureLabelCapacity(ByVal wanted As Long)
    Dim wasEmpty As Boolean

    If wanted <= this.labelCapacity Then Exit Sub

    wasEmpty = (this.labelCapacity = 0)
    If wasEmpty Then this.labelCapacity = BUFFER_GROWTH_MARGIN

    Do While this.labelCapacity < wanted
        this.labelCapacity = this.labelCapacity * 2
    Loop

    If wasEmpty Then
        ReDim this.labelValues(1 To this.labelCapacity)
    Else
        ReDim Preserve this.labelValues(1 To this.labelCapacity)
    End If
End Sub

'@label:ensure-tag-capacity
'@sub-title Grow the tag buffer geometrically so appends stay linear.
Private Sub EnsureTagCapacity(ByVal wanted As Long)
    Dim wasEmpty As Boolean

    If wanted <= this.tagCapacity Then Exit Sub

    wasEmpty = (this.tagCapacity = 0)
    If wasEmpty Then this.tagCapacity = BUFFER_GROWTH_MARGIN

    Do While this.tagCapacity < wanted
        this.tagCapacity = this.tagCapacity * 2
    Loop

    If wasEmpty Then
        ReDim this.tagBuffer(1 To this.tagCapacity)
    Else
        ReDim Preserve this.tagBuffer(1 To this.tagCapacity)
    End If
End Sub


'@section Label buffer
'===============================================================================
'@label:init-label-buffer
'@sub-title Read the label column into a buffer and build the lookup for the cycle.
Private Sub InitLabelBuffer()
    Dim dataRange As Range
    Dim rawValues As Variant
    Dim idx As Long
    Dim label As String

    Set this.labelIndex = New Collection
    this.pendingCount = 0

    Set dataRange = LabelColumnData()

    If dataRange Is Nothing Then
        this.labelCount = 0
        this.labelCapacity = 0
        EnsureLabelCapacity BUFFER_GROWTH_MARGIN
        Exit Sub
    End If

    rawValues = dataRange.Value2

    If IsArray(rawValues) Then
        this.labelCount = UBound(rawValues, 1) - LBound(rawValues, 1) + 1
    Else
        this.labelCount = 1
    End If

    this.labelCapacity = 0
    EnsureLabelCapacity this.labelCount + BUFFER_GROWTH_MARGIN

    For idx = 1 To this.labelCount
        label = BlockText(rawValues, idx)
        this.labelValues(idx) = label
        RegisterLabelKey label, idx
    Next idx
End Sub

'@label:label-key
'@sub-title Build the lookup key for a label, one key per spelling.
'@details The leading character keeps a numeric label legal as a Collection key.
'         The trailing signature is what makes two spellings different keys: a
'         Collection matches its own keys without regard to case, so dropping
'         the label in on its own would fold "Hello" onto "hello" no matter how
'         it was cased here.
Private Function LabelKey(ByVal label As String) As String
    LabelKey = "k" & label & KEY_CASE_SEPARATOR & CaseSignature(label)
End Function

'@label:case-signature
'@sub-title Spell a label's capitalisation out in digits.
'@details Digits carry no case of their own, so a Collection holds this part of
'         the key exactly as written. A character reads 1 when it differs from
'         its own lower case, which is the whole of what a case-insensitive
'         comparison throws away.
Private Function CaseSignature(ByVal label As String) As String
    Dim idx As Long
    Dim labelLength As Long
    Dim character As String
    Dim digits() As String

    labelLength = Len(label)
    If labelLength = 0 Then Exit Function

    ReDim digits(1 To labelLength)

    For idx = 1 To labelLength
        character = Mid$(label, idx, 1)
        If character = LCase$(character) Then
            digits(idx) = "0"
        Else
            digits(idx) = "1"
        End If
    Next idx

    CaseSignature = Join(digits, vbNullString)
End Function

'@label:register-label-key
'@sub-title Record the row a label sits on, keeping the first row when the same spelling repeats.
Private Sub RegisterLabelKey(ByVal label As String, ByVal rowIndex As Long)
    If LenB(label) = 0 Then Exit Sub

    On Error Resume Next
        this.labelIndex.Add rowIndex, LabelKey(label)
    On Error GoTo 0
End Sub

'@label:find-label-row
'@sub-title Find the row index for a label through the keyed lookup.
Private Function FindLabelRow(ByVal label As String) As Long
    If this.labelIndex Is Nothing Then Exit Function

    On Error Resume Next
        FindLabelRow = this.labelIndex(LabelKey(label))
    On Error GoTo 0
End Function

'@label:label-at
'@sub-title Read a label out of the in-memory buffer.
Private Function LabelAt(ByVal rowIndex As Long) As String
    If rowIndex < 1 Then Exit Function
    If rowIndex > this.labelCount Then Exit Function

    LabelAt = CStr(this.labelValues(rowIndex))
End Function


'@section Tag buffer
'===============================================================================
'@label:init-tag-buffer
'@sub-title Read the tag column into a buffer for batch writes.
'@details A row that carries a label and no tag is given the manual tag. It was
'         never claimed by a watched range, so it must survive the update.
Private Sub InitTagBuffer()
    Dim tagRange As Range
    Dim rawTags As Variant
    Dim idx As Long
    Dim tagText As String

    Set tagRange = TagColumnDataRange()

    If tagRange Is Nothing Then
        this.tagBufferSize = 0
        this.tagCapacity = 0
        EnsureTagCapacity BUFFER_GROWTH_MARGIN
        Exit Sub
    End If

    rawTags = tagRange.Value2

    If IsArray(rawTags) Then
        this.tagBufferSize = UBound(rawTags, 1) - LBound(rawTags, 1) + 1
    Else
        this.tagBufferSize = 1
    End If

    this.tagCapacity = 0
    EnsureTagCapacity this.tagBufferSize + BUFFER_GROWTH_MARGIN

    For idx = 1 To this.tagBufferSize
        tagText = BlockText(rawTags, idx)
        If LenB(tagText) = 0 Then
            If LenB(LabelAt(idx)) > 0 Then
                tagText = MANUAL_TAG_NAME & TAG_SEPARATOR & "0"
            End If
        End If
        this.tagBuffer(idx) = tagText
    Next idx
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

'@label:release-update-buffers
'@sub-title Drop the per-cycle buffers so a large table leaves no memory behind.
Private Sub ReleaseUpdateBuffers()
    Erase this.labelValues
    Erase this.tagBuffer

    this.labelCount = 0
    this.labelCapacity = 0
    this.tagBufferSize = 0
    this.tagCapacity = 0
    this.pendingCount = 0

    Set this.labelIndex = Nothing
    Set this.processedRanges = Nothing
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

    For idx = this.processedRanges.LowerBound To this.processedRanges.UpperBound
        If StrComp(CStr(this.processedRanges.Item(idx)), trimmed, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next idx

    this.processedRanges.Push trimmed
End Sub

'@label:track-unresolved-range
'@sub-title Record a registry range whose defined name could not be resolved.
Private Sub TrackUnresolvedRange(ByVal rngName As String)
    Dim trimmed As String
    Dim idx As Long

    trimmed = Trim$(rngName)
    If LenB(trimmed) = 0 Then Exit Sub

    If this.unresolvedRanges Is Nothing Then
        Set this.unresolvedRanges = CreateBuffer()
    End If

    For idx = this.unresolvedRanges.LowerBound To this.unresolvedRanges.UpperBound
        If StrComp(CStr(this.unresolvedRanges.Item(idx)), trimmed, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next idx

    this.unresolvedRanges.Push trimmed
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
    Set this.unresolvedRanges = Nothing
End Sub

'@label:unresolved-ranges
'@prop-title Registry ranges that failed to resolve during the last update.
Public Property Get UnresolvedRanges() As BetterArray
    If this.unresolvedRanges Is Nothing Then
        Set this.unresolvedRanges = CreateBuffer()
    End If

    Set UnresolvedRanges = this.unresolvedRanges.Clone()
End Property

'@label:remove-obsolete-labels
'@sub-title Remove labels whose source ranges were processed but whose tags are stale.
'@details A row that carries a label and no tag stays. Only a row with neither
'         a label nor a tag is dropped, and that is the blank row a ListObject
'         keeps when it holds no data.
Private Sub RemoveObsoleteLabels(ByVal updateSequence As Long)
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim tagText As String
    Dim labelText As String
    Dim rangeName As String
    Dim sequenceValue As Long
    Dim tagValues As Variant
    Dim labelBlock As Variant
    Dim tagRange As Range
    Dim staleRows() As Long
    Dim staleCount As Long

    rowCount = this.table.ListRows.Count
    If rowCount = 0 Then Exit Sub

    Set tagRange = TagColumnDataRange()
    If tagRange Is Nothing Then Exit Sub

    tagValues = ColumnBlock(tagRange)
    labelBlock = ColumnBlock(LabelColumnData())

    ReDim staleRows(1 To rowCount)

    For rowIndex = 1 To rowCount
        tagText = BlockText(tagValues, rowIndex)
        labelText = BlockText(labelBlock, rowIndex)

        If LenB(tagText) = 0 Then
            If LenB(labelText) = 0 Then
                staleCount = staleCount + 1
                staleRows(staleCount) = rowIndex
            End If
        ElseIf ParseTag(tagText, rangeName, sequenceValue) Then
            If WasRangeProcessed(rangeName) And sequenceValue < updateSequence Then
                staleCount = staleCount + 1
                staleRows(staleCount) = rowIndex
            End If
        End If
    Next rowIndex

    If staleCount > 0 Then
        DeleteRowsByIndex staleRows, staleCount
        this.dirty = True
    End If
End Sub

'@label:deduplicate-labels
'@sub-title Remove duplicate labels keeping the row with the oldest (lowest) tag sequence.
Private Sub DeduplicateLabels()
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim labelBlock As Variant
    Dim tagBlock As Variant
    Dim labelValue As String
    Dim tagText As String
    Dim rangeName As String
    Dim seqValue As Long
    Dim seenSlots As Collection
    Dim bestRow() As Long
    Dim bestSeq() As Long
    Dim seenCount As Long
    Dim slot As Long
    Dim indicesToDelete() As Long
    Dim deleteCount As Long

    rowCount = this.table.ListRows.Count
    If rowCount < 2 Then Exit Sub
    If LabelColumnIndex() > this.table.ListColumns.Count Then Exit Sub

    labelBlock = ColumnBlock(LabelColumnData())
    tagBlock = ColumnBlock(TagColumnDataRange())
    If IsEmpty(labelBlock) Then Exit Sub

    Set seenSlots = New Collection
    ReDim bestRow(1 To rowCount)
    ReDim bestSeq(1 To rowCount)
    ReDim indicesToDelete(1 To rowCount)

    For rowIndex = 1 To rowCount
        labelValue = BlockText(labelBlock, rowIndex)

        If LenB(labelValue) > 0 Then
            seqValue = 0
            tagText = BlockText(tagBlock, rowIndex)
            If LenB(tagText) > 0 Then ParseTag tagText, rangeName, seqValue

            slot = 0
            On Error Resume Next
                slot = seenSlots(LabelKey(labelValue))
            On Error GoTo 0

            If slot = 0 Then
                seenCount = seenCount + 1
                bestRow(seenCount) = rowIndex
                bestSeq(seenCount) = seqValue
                seenSlots.Add seenCount, LabelKey(labelValue)
            ElseIf seqValue < bestSeq(slot) Then
                'The current row is older, so the row held until now goes.
                deleteCount = deleteCount + 1
                indicesToDelete(deleteCount) = bestRow(slot)
                bestRow(slot) = rowIndex
                bestSeq(slot) = seqValue
            Else
                deleteCount = deleteCount + 1
                indicesToDelete(deleteCount) = rowIndex
            End If
        End If
    Next rowIndex

    If deleteCount > 0 Then
        DeleteRowsByIndex indicesToDelete, deleteCount
        this.dirty = True
    End If
End Sub

'@label:delete-rows-by-index
'@sub-title Delete the listed table rows, batching them into one Excel call each.
'@details The indices are sorted downward so a flushed batch leaves the
'         remaining indices valid. Union slows down past a few thousand areas,
'         so the batch is capped.
Private Sub DeleteRowsByIndex(ByRef indices() As Long, ByVal count As Long)
    Dim idx As Long
    Dim batch As Range
    Dim areaCount As Long

    If count <= 0 Then Exit Sub

    SortIndicesDescending indices, count

    For idx = 1 To count
        If batch Is Nothing Then
            Set batch = this.table.ListRows(indices(idx)).Range
        Else
            Set batch = Application.Union(batch, this.table.ListRows(indices(idx)).Range)
        End If

        areaCount = areaCount + 1
        If areaCount >= MAX_DELETE_BATCH Then
            batch.Delete Shift:=xlShiftUp
            Set batch = Nothing
            areaCount = 0
        End If
    Next idx

    If Not batch Is Nothing Then batch.Delete Shift:=xlShiftUp
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
'@details The split happens at the LAST separator, so a range name that itself
'         carries the separator still yields its sequence number.
Private Function ParseTag(ByVal tagText As String, ByRef rangeName As String, ByRef sequenceValue As Long) As Boolean
    Dim cut As Long
    Dim numberText As String

    cut = InStrRev(tagText, TAG_SEPARATOR)
    If cut < 2 Then Exit Function

    rangeName = Trim$(Left$(tagText, cut - 1))
    numberText = Trim$(Mid$(tagText, cut + Len(TAG_SEPARATOR)))

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
Public Function MissingLabels(ByVal lang As String) As Long
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
    AppendUnresolvedLine messageLines
    summary = messageLines.ToString(Separator:=vbLf, OpeningDelimiter:=vbNullString, ClosingDelimiter:=vbNullString)
    UpdateMissingSummary messageLines
    NumberOfMissing = summary

    If this.displayPrompts Then
        MsgBox summary, vbInformation + vbOKOnly, "Done!"
    End If
End Function

'@label:append-unresolved-line
'@sub-title Name the registry ranges that could not be resolved during the update.
Private Sub AppendUnresolvedLine(ByVal messageLines As BetterArray)
    If this.unresolvedRanges Is Nothing Then Exit Sub
    If this.unresolvedRanges.Length = 0 Then Exit Sub

    messageLines.Push "These ranges could not be found: " & _
                      this.unresolvedRanges.ToString(Separator:=", ", OpeningDelimiter:=vbNullString, ClosingDelimiter:=vbNullString) & "."
End Sub

'@label:update-missing-summary
'@sub-title Update internal missing summary.
Private Sub UpdateMissingSummary(ByVal summary As BetterArray)
    Set this.missingSummary = summary.Clone()
End Sub

'@label:internal-missing-summary
'@prop-title Cached missing-labels summary across languages.
Public Property Get InternalMissingSummary() As BetterArray
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
Public Property Get InternalDuplicatesSummary() As BetterArray
    If this.duplicatesSummary Is Nothing Then
        Set this.duplicatesSummary = CreateBuffer()
    End If

    Set InternalDuplicatesSummary = this.duplicatesSummary.Clone()
End Property


'@label:duplicate-labels
'@sub-title Report duplicate translation labels when they exist.
Public Function DuplicateLabels(ByRef duplicateMessage As String, Optional ByVal languageName As String = vbNullString) As Boolean
    Dim summary As TDuplicateSummary

    summary = ComputeDuplicateSummary(languageName)
    duplicateMessage = summary.Message
    DuplicateLabels = summary.HasDuplicates
End Function

'@label:compute-duplicate-summary
'@sub-title Count duplicates for one language or for the label column.
'@details Every label is counted through one keyed Collection, so the cost is
'         linear in the row count. Two spellings of the same word are two
'         labels here, so neither is reported as a duplicate of the other.
Private Function ComputeDuplicateSummary(ByVal languageName As String) As TDuplicateSummary
    Dim labelRange As Range
    Dim targetRange As Range
    Dim summary As TDuplicateSummary
    Dim columnTitle As String
    Dim messageLines As BetterArray
    Dim trimmedLanguage As String
    Dim groupLabels() As String
    Dim groupCounts() As Long
    Dim groupCount As Long
    Dim groupIndex As Long

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

    groupCount = CollectDuplicateGroups(targetRange, groupLabels, groupCounts)
    If groupCount = 0 Then GoTo Done

    Set messageLines = CreateBuffer()
    messageLines.Push "Duplicate labels detected in column " & columnTitle & "!"

    For groupIndex = 1 To groupCount
        messageLines.Push BuildDuplicateLine(groupLabels(groupIndex), groupCounts(groupIndex))
    Next groupIndex

    summary.HasDuplicates = True
    summary.Message = messageLines.ToString(Separator:=vbLf, OpeningDelimiter:=vbNullString, ClosingDelimiter:=vbNullString)

Done:
    ComputeDuplicateSummary = summary
End Function

'@label:collect-duplicate-groups
'@sub-title Find every value a column holds more than once, in first-appearance order.
'@details One keyed Collection carries the whole count, so the cost is linear in
'         the row count. The key comes from LabelKey and spells the
'         capitalisation out, so "Hello" and "hello" are counted apart and
'         neither is a duplicate of the other. Both the summary message and the
'         sheet colouring read their groups from here, which is what keeps the
'         two of them saying the same thing.
'@param targetRange Range. The one column to count.
'@param groupLabels ByRef String(). Receives the duplicated values, 1 based.
'@param groupCounts ByRef Long(). Receives how many times each one appears, 1 based.
'@return Long. Number of duplicate groups found.
Private Function CollectDuplicateGroups(ByVal targetRange As Range, _
                                        ByRef groupLabels() As String, _
                                        ByRef groupCounts() As Long) As Long
    Dim values As Variant
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim currentLabel As String
    Dim seenSlots As Collection
    Dim seenLabel() As String
    Dim seenCount() As Long
    Dim uniqueCount As Long
    Dim slot As Long
    Dim bucketIndex As Long
    Dim groupCount As Long

    If targetRange Is Nothing Then Exit Function

    values = targetRange.Value2
    If IsArray(values) Then
        rowCount = UBound(values, 1) - LBound(values, 1) + 1
    Else
        rowCount = 1
    End If
    If rowCount = 0 Then Exit Function

    Set seenSlots = New Collection
    ReDim seenLabel(1 To rowCount)
    ReDim seenCount(1 To rowCount)

    For rowIndex = 1 To rowCount
        currentLabel = BlockText(values, rowIndex)

        If LenB(currentLabel) > 0 Then
            slot = 0
            On Error Resume Next
                slot = seenSlots(LabelKey(currentLabel))
            On Error GoTo 0

            If slot = 0 Then
                uniqueCount = uniqueCount + 1
                seenLabel(uniqueCount) = currentLabel
                seenCount(uniqueCount) = 1
                seenSlots.Add uniqueCount, LabelKey(currentLabel)
            Else
                seenCount(slot) = seenCount(slot) + 1
            End If
        End If
    Next rowIndex

    ReDim groupLabels(1 To rowCount)
    ReDim groupCounts(1 To rowCount)

    For bucketIndex = 1 To uniqueCount
        If seenCount(bucketIndex) > 1 Then
            groupCount = groupCount + 1
            groupLabels(groupCount) = seenLabel(bucketIndex)
            groupCounts(groupCount) = seenCount(bucketIndex)
        End If
    Next bucketIndex

    CollectDuplicateGroups = groupCount
End Function

'@label:label-column-data
'@sub-title Access the label column data range.
Private Function LabelColumnData() As Range
    If this.table.ListRows.Count = 0 Then Exit Function
    On Error Resume Next
        Set LabelColumnData = this.table.ListColumns(LabelColumnIndex).DataBodyRange
    On Error GoTo 0
End Function

'@label:column-block
'@sub-title Read a one column range in a single crossing.
Private Function ColumnBlock(ByVal source As Range) As Variant
    If source Is Nothing Then Exit Function
    ColumnBlock = source.Value2
End Function

'@label:block-text
'@sub-title Read one row out of a column block as text.
Private Function BlockText(ByVal block As Variant, ByVal rowIndex As Long) As String
    If IsArray(block) Then
        BlockText = SafeText(block(LBound(block, 1) - 1 + rowIndex, LBound(block, 2)))
    ElseIf rowIndex = 1 Then
        BlockText = SafeText(block)
    End If
End Function

'@label:safe-text
'@sub-title Turn a sheet value into text, answering empty for an error or a Null.
'@details A cell holding #N/A or #REF! answers with an Error variant, and CStr
'         on one raises 13. That killed a whole update and reported a type
'         mismatch that pointed nowhere.
Private Function SafeText(ByVal candidate As Variant) As String
    If IsError(candidate) Then Exit Function
    If IsNull(candidate) Then Exit Function

    SafeText = Trim$(CStr(candidate))
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

    'TestTagColumnStaysInvisibleAfterUpdate reads all three back, so a failure
    'here shows as a red test.
    If Not header Is Nothing Then
        header.Value = TAG_HEADER_TITLE
        header.Font.Color = vbWhite
        header.Interior.Color = vbWhite
    End If

    Set tagData = TagColumnDataRange()
    If tagData Is Nothing Then Exit Sub

    tagData.Font.Color = vbWhite
    PaintTagBorders tagData
End Sub

'@label:paint-tag-borders
'@sub-title Paint every edge of the tag range white so the column reads as empty.
Private Sub PaintTagBorders(ByVal target As Range)
    PaintHairlineBorder target, xlEdgeTop
    PaintHairlineBorder target, xlEdgeBottom
    PaintHairlineBorder target, xlEdgeLeft
    PaintHairlineBorder target, xlEdgeRight

    'An inside edge exists only where there are two rows or two columns for it
    'to sit between. Asking for one that cannot exist raises 1004, and the tag
    'range is one column wide.
    If target.Rows.Count > 1 Then PaintHairlineBorder target, xlInsideHorizontal
    If target.Columns.Count > 1 Then PaintHairlineBorder target, xlInsideVertical
End Sub

'@label:paint-hairline-border
'@sub-title Paint one edge of a range as a white hairline.
Private Sub PaintHairlineBorder(ByVal target As Range, ByVal edge As Long)
    With target.Borders(edge)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = vbWhite
    End With
End Sub

'@label:apply-lang-column-formatting
'@sub-title Apply font, sizing, and duplicate highlighting to a language column.
'@details The group rules go on first and the red catch-all goes on last.
'         FormatConditions.Add appends, so the rule added last holds the lowest
'         priority and shows through only where no group rule claimed the cell.
'@param langRng Range. The language column to format.
'@param painted ByRef Long. How many groups the pass has coloured so far, carried
'       across every column so no two groups share a colour. See
'       HighlightDuplicateGroups.
Private Sub ApplyLangColumnFormatting(ByVal langRng As Range, ByRef painted As Long)

    If langRng Is Nothing Then Exit Sub

    With langRng

        .Cells(0, 1).Locked = False 'Unlock the first cell for the language
        .Cells.Font.Name = "Consolas"
        .Cells.Font.Size = 16
        .Cells(0, 1).Font.Size = 18
        .Cells.WrapText = True
        .Cells.HorizontalAlignment = xlHAlignLeft
        .Cells.VerticalAlignment = xlVAlignCenter

    End With

    HighlightDuplicateGroups langRng, painted

    'Formatting conditions
    langRng.FormatConditions.AddUniqueValues

    With langRng.FormatConditions(langRng.FormatConditions.Count)
        .DupeUnique = xlDuplicate
        .Interior.Color = vbRed
    End With

End Sub

'@label:highlight-duplicate-groups
'@sub-title Give every duplicate group in a column a fill of its own.
'@details A column painted in one colour says which cells are duplicated and
'         nothing about which of them belong together. One rule per group says
'         both, and the eye pairs them without reading a word.
'
'         Three things send a group back to the red catch-all: a palette that
'         has run out (MAX_DUPLICATE_GROUPS), a label whose rule formula would
'         pass MAX_RULE_FORMULA_LENGTH, and a formula Excel refuses. So red no
'         longer reads as "duplicate". It reads as "duplicate this could not
'         group", and a sheet with red on it is still telling the truth.
'
'         THE COUNTER IS THE CALLER'S, NOT THIS ROUTINE'S
'         It used to start at zero on every column, so the first group of every
'         language drew the same colour. An untranslated row carries the source
'         spelling in both columns, so those same-coloured cells landed on the
'         same rows and the sheet read as though whole rows were highlighted.
'         Duplicates are counted within one column and always were; what crossed
'         the columns was the colour. Carrying the count across the pass gives
'         every group on the sheet a colour of its own, and the palette is then
'         spent sheet-wide rather than per column.
'@param langRng Range. The language column to paint.
'@param painted ByRef Long. Groups coloured so far in this pass, advanced here.
Private Sub HighlightDuplicateGroups(ByVal langRng As Range, ByRef painted As Long)
    Dim groupLabels() As String
    Dim groupCounts() As Long
    Dim groupCount As Long
    Dim groupIndex As Long
    Dim anchor As String

    groupCount = CollectDuplicateGroups(langRng, groupLabels, groupCounts)
    If groupCount = 0 Then Exit Sub

    'A format condition formula is read from the top left cell of the range it
    'is given and applied down from there, so the anchor is that cell with its
    'row left relative.
    anchor = langRng.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    For groupIndex = 1 To groupCount
        If painted >= MAX_DUPLICATE_GROUPS Then Exit For
        If AddGroupCondition(langRng, anchor, groupLabels(groupIndex), DuplicateGroupColor(painted)) Then
            painted = painted + 1
        End If
    Next groupIndex
End Sub

'@label:add-group-condition
'@sub-title Give one duplicate group its own format condition.
'@details EXACT is what carries the case, so the sheet groups labels the way
'         CollectDuplicateGroups counts them and "Hello" never shares a colour
'         with "hello".
'
'         Range.Formula reads US syntax whatever the host locale is and Formula1
'         follows it, so the formula is written in US form with a comma and no
'         localised form is attempted. FormatConditions.Add appends without
'         complaint, so the count is what says whether it was accepted.
'@return Boolean. True when the group came away with a rule.
Private Function AddGroupCondition(ByVal langRng As Range, _
                                   ByVal anchor As String, _
                                   ByVal label As String, _
                                   ByVal fillColor As Long) As Boolean
    Dim conditionsBefore As Long
    Dim ruleFormula As String

    ruleFormula = GroupConditionFormula(anchor, label)
    If LenB(ruleFormula) = 0 Then Exit Function

    conditionsBefore = langRng.FormatConditions.Count

    On Error Resume Next
        langRng.FormatConditions.Add Type:=xlExpression, Formula1:=ruleFormula
    On Error GoTo 0

    If langRng.FormatConditions.Count = conditionsBefore Then Exit Function

    With langRng.FormatConditions(langRng.FormatConditions.Count)
        .Interior.Color = fillColor
        .StopIfTrue = True
    End With

    AddGroupCondition = True
End Function

'@label:group-condition-formula
'@sub-title Build the EXACT test for one duplicate group.
'@details A label carrying a double quote has to have it doubled or the formula
'         closes early and Excel refuses the whole rule.
'@return String. The formula, or empty when it would pass the length cap.
Private Function GroupConditionFormula(ByVal anchor As String, _
                                       ByVal label As String) As String
    Dim quoted As String
    Dim candidate As String

    quoted = Replace(label, Chr$(34), Chr$(34) & Chr$(34))
    candidate = "=EXACT(" & anchor & "," & Chr$(34) & quoted & Chr$(34) & ")"

    If Len(candidate) > MAX_RULE_FORMULA_LENGTH Then Exit Function

    GroupConditionFormula = candidate
End Function

'@label:duplicate-color-count
'@prop-title How many duplicate groups the palette colours before it repeats.
Public Property Get DuplicateColorCount() As Long
    DuplicateColorCount = MAX_DUPLICATE_GROUPS
End Property

'@label:duplicate-group-color
'@sub-title Colour for the nth duplicate group.
'@details The hue walks the wheel in strides of DUPLICATE_HUE_STRIDE. The stride
'         and the hue count share no factor, so every hue is still reached, and
'         two groups that follow each other land most of a wheel apart rather
'         than one step apart. Each further lap of the wheel drops the lightness
'         by one band. Past the palette the index wraps, so two groups do share
'         a colour there; HighlightDuplicateGroups stops before that happens and
'         leaves the rest to the red catch-all.
'@param groupIndex Long. Zero based position of the group.
'@return Long. An RGB value ready for Interior.Color.
'@export
Public Function DuplicateGroupColor(ByVal groupIndex As Long) As Long
    Dim wrapped As Long
    Dim hueStep As Long
    Dim lap As Long
    Dim hue As Double
    Dim lightness As Double

    If groupIndex < 0 Then groupIndex = 0

    wrapped = groupIndex Mod MAX_DUPLICATE_GROUPS
    hueStep = (wrapped * DUPLICATE_HUE_STRIDE) Mod DUPLICATE_HUE_COUNT
    lap = wrapped \ DUPLICATE_HUE_COUNT

    hue = DUPLICATE_HUE_START + hueStep * (360# / DUPLICATE_HUE_COUNT)
    Do While hue >= 360#
        hue = hue - 360#
    Loop

    lightness = DUPLICATE_LIGHTNESS_TOP - lap * DUPLICATE_LIGHTNESS_STEP

    DuplicateGroupColor = HslToRgb(hue, DUPLICATE_SATURATION, lightness)
End Function

'@label:hsl-to-rgb
'@sub-title Convert one HSL triplet into the Long an Interior.Color takes.
'@details Hue picks the sector of the wheel, chroma is how far the colour sits
'         from grey at that lightness, and the match value lifts all three
'         channels so the result reads back at the lightness asked for. Stepping
'         hue alone then keeps every group equally pale, which is what makes the
'         black Consolas text on top stay readable whatever colour a group drew.
'@param hue Double. Degrees, 0 to 360.
'@param saturation Double. 0 to 1.
'@param lightness Double. 0 to 1.
'@return Long. RGB colour.
Private Function HslToRgb(ByVal hue As Double, _
                          ByVal saturation As Double, _
                          ByVal lightness As Double) As Long
    Dim chroma As Double
    Dim sextant As Double
    Dim midPoint As Double
    Dim match As Double
    Dim sector As Long
    Dim redPart As Double
    Dim greenPart As Double
    Dim bluePart As Double

    chroma = (1# - Abs(2# * lightness - 1#)) * saturation
    sextant = hue / 60#
    sector = Int(sextant) Mod 6
    midPoint = chroma * (1# - Abs((sextant - 2# * Int(sextant / 2#)) - 1#))
    match = lightness - chroma / 2#

    Select Case sector
        Case 0
            redPart = chroma
            greenPart = midPoint
        Case 1
            redPart = midPoint
            greenPart = chroma
        Case 2
            greenPart = chroma
            bluePart = midPoint
        Case 3
            greenPart = midPoint
            bluePart = chroma
        Case 4
            redPart = midPoint
            bluePart = chroma
        Case Else
            redPart = chroma
            bluePart = midPoint
    End Select

    HslToRgb = RGB(ColorByte(redPart + match), ColorByte(greenPart + match), ColorByte(bluePart + match))
End Function

'@label:color-byte
'@sub-title Turn a 0 to 1 channel into the 0 to 255 byte RGB expects.
Private Function ColorByte(ByVal channel As Double) As Long
    Dim scaled As Long

    scaled = CLng(channel * 255#)
    If scaled < 0 Then scaled = 0
    If scaled > 255 Then scaled = 255

    ColorByte = scaled
End Function

'@label:configure-tag-column
'@sub-title Locate and validate the helper column. No cell is written here.
Private Sub ConfigureTagColumn()
    Dim header As Range

    ValidateInitialised

    Set header = this.table.HeaderRowRange.Cells(1, 1)
    If header.Column = 1 Then
        ThrowProjectError ProjectError.ErrorUnexpectedState, "Translations table requires a helper column immediately to the left."
    End If

    Set this.tagHeader = header.Offset(0, -1)
End Sub

'@label:ensure-tag-column-ready
'@sub-title Write the helper column title and its stealth formatting.
Private Sub EnsureTagColumnReady()
    Dim header As Range
    Dim currentTitle As String

    Set header = RequireTagHeader()

    currentTitle = SafeText(header.Value)
    If StrComp(currentTitle, TAG_HEADER_TITLE, vbBinaryCompare) <> 0 Then
        header.Value = TAG_HEADER_TITLE
    End If

    ApplyTagColumnFormatting
End Sub

'@label:require-tag-header
'@sub-title Return the tag header cell, raising an error when unavailable.
Private Function RequireTagHeader() As Range
    If Not IsTagHeaderAttached() Then
        AttachTagHeader
    End If

    If Not IsTagHeaderAttached() Then
        ThrowProjectError ProjectError.ErrorUnexpectedState, _
                          "Translations table requires a helper column titled '" & TAG_HEADER_TITLE & "' immediately to the left of the listobject."
    End If

    Set RequireTagHeader = this.tagHeader
End Function

'@label:is-tag-header-attached
'@sub-title Test that the held header cell is alive and still sits beside the table.
'@details Deleting the helper column invalidates the held Range. Inserting a
'         column between the helper and the table leaves it alive and no longer
'         adjacent, and resizing over the wrong column follows from that.
Private Function IsTagHeaderAttached() As Boolean
    Dim tableHeader As Range

    If Not IsRangeValid(this.tagHeader) Then Exit Function

    'While the tag column sits inside the table it IS the first table column,
    'so there is nothing to the left of it to compare against.
    If this.tagsIntegrated Then
        IsTagHeaderAttached = True
        Exit Function
    End If

    On Error GoTo Detached
        Set tableHeader = this.table.HeaderRowRange.Cells(1, 1)
        If tableHeader.Column = 1 Then Exit Function
        IsTagHeaderAttached = (tableHeader.Offset(0, -1).Address(External:=True) = this.tagHeader.Address(External:=True))
    On Error GoTo 0
    Exit Function

Detached:
    Err.Clear
    IsTagHeaderAttached = False
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

    'A rename that fails leaves the tag column inside the table under whatever
    'name it carried. Every index here comes from LabelColumnIndex(), which
    'counts positions, so the update would still run and the sheet would be
    'left holding a column nobody can name.
    this.table.ListColumns(1).Name = TAG_HEADER_TITLE
End Sub

'@label:end-tag-integration
'@sub-title Restore the translations table to exclude the helper tag column.
'@details The tag formatting is applied here alone, once per cycle.
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
    'One row or none is already in order, and that is the case where Apply has
    'nothing to work on. Past that a failed sort is reported: the caller has
    'been told the table is ordered and acts on it.
    If this.table.ListRows.Count < 2 Then Exit Sub

    With this.table.Sort
        .SortFields.Clear
        .SortFields.Add Key:=this.table.ListColumns(LabelColumnIndex).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With
End Sub

'@label:apply-formatting
'@sub-title Apply light formatting and duplicate highlighting once per update.
'@details The row autofit runs once at the end. EntireRow.AutoFit measures
'         every used column on the sheet, so one call per language column was
'         the same measurement done again and again.
Private Sub ApplyFormatting()
    Dim dataRange As Range
    Dim tableRange As Range
    Dim columnRange As Range
    Dim unlockRange As Range
    Dim unlockOffset As Long
    Dim unlockWidth As Long
    Dim paintedGroups As Long


    Set dataRange = this.table.DataBodyRange
    Set tableRange = this.table.Range

    'TestTableHeaderRowIsWhite and TestTagColumnStaysInvisibleAfterUpdate both
    'read back what these three writes produce, so a failure shows as a red test.
    this.table.TableStyle = "TableStyleLight11"
    tableRange.EntireColumn.ColumnWidth = 60
    tableRange.Rows(1).Font.Color = vbWhite

    If dataRange Is Nothing Then Exit Sub

    'A delete that fails leaves the old rules in place and a new one is added
    'below on every update, until the sheet is slow to scroll.
    'TestFormatConditionsDoNotAccumulate watches the count.
    dataRange.FormatConditions.Delete

    unlockWidth = this.table.ListColumns.Count - LabelColumnIndex()
    If unlockWidth > 0 Then
        unlockOffset = LabelColumnIndex()
        'An offset past the last column of the sheet answers Nothing, and that
        'is a legitimate "there is nothing to unlock".
        On Error Resume Next
            Set unlockRange = dataRange.Offset(0, unlockOffset).Resize(dataRange.Rows.Count, unlockWidth)
        On Error GoTo 0

        'A language column left locked stops the user typing a translation into
        'it once the sheet is protected, with nothing saying why.
        If Not unlockRange Is Nothing Then unlockRange.Locked = False
    End If

    On Error Resume Next
        Set columnRange = dataRange.Columns(LabelColumnIndex())
    On Error GoTo 0

    If columnRange Is Nothing Then Exit Sub

    'One counter for the whole pass, so the first duplicate group of the second
    'language does not draw the colour the first language already used.
    Do While Not columnRange Is Nothing And Not Intersect(columnRange, dataRange) Is Nothing
        ApplyLangColumnFormatting columnRange, paintedGroups
        Set columnRange = columnRange.Offset(0, 1)
        If columnRange.Column > dataRange.Columns(dataRange.Columns.Count).Column Then Exit Do
    Loop

    dataRange.EntireRow.AutoFit
End Sub


'@section Data exchange
'===============================================================================
'@label:export
'@sub-title Export translations into a destination workbook.
Public Sub Export(ByVal exportWorkbook As Workbook, Optional ByVal Hide As Long = xlSheetVisible)
    Dim exportSheet As Worksheet
    Dim tableAdapter As CustomTable
    Dim hostSheet As Worksheet
    Dim store As HiddenNames

    ValidateInitialised
    If exportWorkbook Is Nothing Then
        ThrowProjectError ProjectError.ObjectNotInitialized, "Target workbook is missing; cannot export translations."
    End If

    Set hostSheet = HostWorksheet

    'The exported workbook reads its language list from the hidden name copied
    'below, so the name is brought up to date first. Create leaves it alone,
    'and a caller that only exports may never have run an update.
    PersistLanguageList

    'Test if the export workbook has the translation sheet name
    On Error Resume Next
        Set exportSheet = exportWorkbook.Worksheets(hostSheet.Name)
    On Error GoTo 0

    If (exportSheet Is Nothing) Then
        Set exportSheet = exportWorkbook.Worksheets.Add(After:=exportWorkbook.Worksheets(exportWorkbook.Worksheets.Count))
        exportSheet.Name = hostSheet.Name
    End If

    Set tableAdapter = CustomTable.Create(this.table)

    'Start exporting at column two to preserve the helper tag column offset.
    'CustomTable.Export clears the sheet itself, so one clear covers both paths.
    tableAdapter.Export exportSheet, startColumn:=2, startLine:=1, addListObject:=True

    Set store = MetadataStore()
    'Persist translation metadata (languages, formatting markers) alongside the exported table.
    If Not store Is Nothing Then store.ExportNames exportSheet

    exportSheet.Rows(1).Font.Bold = True
    exportSheet.Cells.Font.Name = "Consolas"
    exportSheet.Cells.HorizontalAlignment = xlHAlignLeft
    exportSheet.Cells.VerticalAlignment = xlVAlignCenter
    exportSheet.Cells.EntireRow.AutoFit
    exportSheet.Cells.EntireColumn.AutoFit

    'A visibility write that fails leaves an internal sheet on show in the
    'workbook the user receives.
    exportSheet.Visible = Hide
End Sub


'@section Registry bookkeeping
'===============================================================================
'@label:planned-update-sequence
'@sub-title Work out the sequence this cycle runs at, without writing it.
Private Function PlannedUpdateSequence(ByVal registrySheet As Worksheet) As Long
    Dim currentValue As Long

    currentValue = CounterNumericValue(registrySheet)
    If currentValue < 0 Then currentValue = 0

    PlannedUpdateSequence = currentValue + 1
End Function

'@label:persist-counter-value
'@sub-title Persist the counter value inside a hidden worksheet-level name.
Private Sub PersistCounterValue(ByVal registrySheet As Worksheet, ByVal counterValue As Long)
    Dim store As HiddenNames

    If counterValue < 0 Then counterValue = 0

    Set store = HiddenNameStore(registrySheet)
    If store Is Nothing Then Exit Sub

    store.EnsureName COUNTER_NAME, counterValue, HiddenNameTypeLong
    store.SetValue COUNTER_NAME, counterValue
End Sub

'@label:counter-numeric-value
'@sub-title Retrieve the current counter value stored in worksheet names.
Private Function CounterNumericValue(ByVal registrySheet As Worksheet) As Long
    Dim store As HiddenNames

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
'@details A store that cannot be built means the sequence counter is never
'         written. The next run then reads zero, treats the cycle as sequence 1
'         and reprocesses every row whatever its status, so this failure has to
'         be heard.
Private Function HiddenNameStore(ByVal registrySheet As Worksheet) As HiddenNames
    If registrySheet Is Nothing Then Exit Function

    Set HiddenNameStore = HiddenNames.Create(registrySheet)
End Function

'@label:metadata-store
'@sub-title Resolve the hidden name store hosting metadata.
Private Function MetadataStore() As HiddenNames
    Dim hostSh As Worksheet

    If Not this.metadataStore Is Nothing Then
        Set MetadataStore = this.metadataStore
        Exit Function
    End If

    If this.table Is Nothing Then Exit Function
    Set hostSh = this.table.Parent
    If hostSh Is Nothing Then Exit Function

    'A store that cannot be built means the language list is never persisted,
    'and the exported workbook carries no languages.
    Set this.metadataStore = HiddenNames.Create(hostSh)

    Set MetadataStore = this.metadataStore
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
'@param errNumber Long. The error number to raise. Callers also pass a raw Err.Number.
Private Sub ThrowProjectError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise errNumber, CLASS_NAME, message
End Sub
