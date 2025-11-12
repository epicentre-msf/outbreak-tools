VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MasterSetupVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Lightweight manager for the Master Setup Variables table"

'@PredeclaredId
'@Folder("MasterSetup")
'@ModuleDescription("Lightweight manager for the Master Setup Variables table wrapping CustomTable interactions")
'@depends CustomTable, ICustomTable, HiddenNames, IHiddenNames, BetterArray, ProjectError, IDropdownLists, ILLChoices, ITranslationObject, DataSheet, IDataSheet
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

Implements IMasterSetupVariables


'@section State and constants
'===============================================================================
Private Type TMasterSetupVariables
    table As ListObject              'Underlying variables table
    wrapper As ICustomTable          'Cached CustomTable wrapper for row management
    hiddenNames As IHiddenNames      'Worksheet-level hidden name manager
    colOrder As Long                 'Cached column index for "Variable Order"
    colSection As Long               'Cached column index for "Variable Section"
    colName As Long                  'Cached column index for "Variable Name"
    colLabel As Long                 'Cached column index for "Label"
    colDefaultChoice As Long         'Cached column index for "Default Choice"
    colChoices As Long               'Cached column index for "Choices Values"
    colDefaultStatus As Long         'Cached column index for "Default Status"
    colComments As Long              'Cached column index for "Comments"
    initialised As Boolean           'Flag that tracks whether initialization logic executed
End Type

Private this As TMasterSetupVariables

Private Const CLASS_NAME As String = "MasterSetupVariables"
Private Const NAME_PREFIX As String = "__MSV__"
Private Const NAME_INITIALISED As String = NAME_PREFIX & "Initialised"
Private Const NAME_COL_ORDER As String = NAME_PREFIX & "ColOrder"
Private Const NAME_COL_SECTION As String = NAME_PREFIX & "ColSection"
Private Const NAME_COL_NAME As String = NAME_PREFIX & "ColName"
Private Const NAME_COL_LABEL As String = NAME_PREFIX & "ColLabel"
Private Const NAME_COL_DEFAULT_CHOICE As String = NAME_PREFIX & "ColDefaultChoice"
Private Const NAME_COL_CHOICES As String = NAME_PREFIX & "ColChoices"
Private Const NAME_COL_DEFAULT_STATUS As String = NAME_PREFIX & "ColDefaultStatus"
Private Const NAME_COL_COMMENTS As String = NAME_PREFIX & "ColComments"
Private Const NAME_SHEET_TAG As String = "sheetTag"

Private Const HEADER_ORDER As String = "Variable Order"
Private Const HEADER_SECTION As String = "Variable Section"
Private Const HEADER_NAME As String = "Variable Name"
Private Const HEADER_LABEL As String = "Variable Label"
Private Const HEADER_DEFAULT_CHOICE As String = "Default Choice"
Private Const HEADER_CHOICES As String = "Choices Values"
Private Const HEADER_DEFAULT_STATUS As String = "Default Status"
Private Const HEADER_COMMENTS As String = "Comments"

Private Const STATUS_DROPDOWN_NAME As String = "__var_status"
Private Const STATUS_MANDATORY As String = "mandatory"
Private Const STATUS_VISIBLE As String = "optional, visible"
Private Const STATUS_HIDDEN As String = "optional, hidden"

Private Const DUPLICATE_FILL_COLOR As Long = 12830708       'RGB(244, 199, 195)
Private Const DUPLICATE_FONT_COLOR As Long = 393372         'RGB(156, 0, 6)
Private Const VARIABLE_NAME_RANGE As String = "__Col__Variables"


'@section Factory helpers
'===============================================================================
'@sub-title Create a configured variables table manager instance.
'@param listObject ListObject representing the Master Setup Variables worksheet table.
'@return IMasterSetupVariables ready to manage the associated ListObject.
Public Function Create(ByVal listObject As ListObject) As IMasterSetupVariables

    ValidateListObject listObject
    EnsureRequiredColumns listObject

    With New MasterSetupVariables
        .AssignListObject listObject
        Set Create = .Self
    End With
End Function

'@sub-title Self reference used by factories returning interface instances.
Public Property Get Self() As IMasterSetupVariables
    Set Self = Me
End Property


'@section Public API
'===============================================================================
'@sub-title Initialise validation, hidden names, and duplicate detection.
'@param dropdowns IDropdownLists provider used to configure the Default Status validation.
Public Sub Initialise(ByVal dropdowns As IDropdownLists)
    ValidateReady
    ApplyDropdownValidation dropdowns
    PersistColumnMetadata
    PersistInitialisedFlag True
    ApplyDuplicateFormatting
End Sub

'@sub-title Refresh the concatenated choices values for the supplied variable.
'@param variableName String identifier of the variable row to update.
'@param choices ILLChoices providing category listings.
Public Sub RefreshChoices(ByVal variableName As String, ByVal choices As ILLChoices)
    Dim rowIndex As Long
    Dim concatenated As String

    ValidateReady
    variableName = Trim$(variableName)
    If LenB(variableName) = 0 Then Exit Sub

    If choices Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "RefreshChoices requires an ILLChoices instance."
    End If

    rowIndex = LocateVariableRow(variableName)
    If rowIndex = 0 Then
        ThrowError ProjectError.ElementNotFound, "Variable '" & variableName & "' does not exist in the variables table."
    End If

    concatenated = vbNullString
    If choices.ChoiceExists(variableName) Then
        concatenated = choices.ConcatenateCategories(variableName, " | ")
    End If

    WriteValue rowIndex, HEADER_CHOICES, concatenated
End Sub

'@sub-title Retrieve the label for the supplied variable.
Public Function LabelFor(ByVal variableName As String) As String
    ValidateReady
    LabelFor = ValueFor(variableName, HEADER_LABEL)
End Function

'@sub-title Retrieve the default choice for the supplied variable.
Public Function DefaultChoiceFor(ByVal variableName As String) As String
    ValidateReady
    DefaultChoiceFor = ValueFor(variableName, HEADER_DEFAULT_CHOICE)
End Function

'@sub-title Retrieve the concatenated choices values for the supplied variable.
Public Function ChoicesFor(ByVal variableName As String) As String
    ValidateReady
    ChoicesFor = ValueFor(variableName, HEADER_CHOICES)
End Function

'@sub-title Retrieve the default status for the supplied variable.
Public Function DefaultStatusFor(ByVal variableName As String) As String
    ValidateReady
    DefaultStatusFor = ValueFor(variableName, HEADER_DEFAULT_STATUS)
End Function

'@sub-title Retrieve the section label for the supplied variable.
Public Function SectionFor(ByVal variableName As String) As String
    ValidateReady
    SectionFor = ValueFor(variableName, HEADER_SECTION)
End Function

'@sub-title Append rows to the underlying table.
Public Sub AddRows(Optional ByVal rowCount As Long = 20)
    ValidateReady
    TableWrapper.AddRows nbRows:=rowCount
End Sub

'@sub-title Remove rows from the underlying table.
Public Sub RemoveRows(Optional ByVal rowCount As Long = 1)
    Dim idx As Long
    ValidateReady
    TableWrapper.RemoveRows rowCount
End Sub

'@sub-title Translate the header row with the supplied translations object.
Public Sub TranslateHeaders(ByVal trads As ITranslationObject)
    Dim headerRange As Range

    ValidateReady
    If trads Is Nothing Then Exit Sub

    Set headerRange = TableRange(includeHeaders:=True).Rows(1)
    If headerRange Is Nothing Then Exit Sub

    trads.TranslateRange headerRange
End Sub

'@sub-title Translate the key content columns without affecting headers.
Public Sub TranslateContent(ByVal trads As ITranslationObject)
    ValidateReady
    If trads Is Nothing Then Exit Sub
    TranslateColumn trads, HEADER_SECTION
    TranslateColumn trads, HEADER_LABEL
    TranslateColumn trads, HEADER_COMMENTS
End Sub

'@sub-title Translate both headers and content using the supplied translation object.
Public Sub TranslateAll(ByVal contentsTrads As ITranslationObject, Optional ByVal ribbonTrads As ITranslationObject)
    TranslateContent contentsTrads
    If (Not ribbonTrad Is Nothing) Then
        TranslateHeaders trads
    End If
End Sub

'@sub-title Apply workbook-scoped validation for variable names onto the supplied range.
Public Sub SetVariableValidation(ByVal targetRange As Range)
    Dim workbookStore As IHiddenNames

    ValidateReady
    If targetRange Is Nothing Then Exit Sub

    Set workbookStore = WorkbookHiddenNameStore()
    If workbookStore Is Nothing Then Exit Sub
    If Not workbookStore.HasName(VARIABLE_NAME_RANGE) Then Exit Sub

    ApplyVariableDropdownValidation targetRange, VARIABLE_NAME_RANGE
End Sub

'@sub-title Clone the managed table into another workbook.
'@param targetWorkbook Workbook receiving the cloned table.
'@param sheetName Optional sheet name; defaults to the source worksheet name.
'@param startCell Optional starting cell; defaults to the source table anchor.
'@return IMasterSetupVariables bound to the cloned ListObject.
Public Function CloneToWorkbook(ByVal targetWorkbook As Workbook, _
                                Optional ByVal sheetName As String = vbNullString, _
                                Optional ByVal startCell As Range = Nothing) As IMasterSetupVariables

    Dim targetSheet As Worksheet
    Dim targetTable As ListObject

    ValidateReady
    If targetWorkbook Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "CloneToWorkbook requires a target workbook."
    End If

    If LenB(sheetName) = 0 Then
        sheetName = HostSheet.Name
    End If

    Set targetSheet = ResolveOrCreateSheet(targetWorkbook, sheetName)
    Set targetTable = CopyIntoSheet(targetSheet)

    ExportMetadata targetSheet

    Set CloneToWorkbook = MasterSetupVariables.Create(targetTable)
End Function

'@sub-title Restore current table contents from a peer instance.
'@param source IMasterSetupVariables providing data and metadata.
Public Sub RestoreFrom(ByVal source As IMasterSetupVariables)
    ValidateReady
    If source Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "RestoreFrom requires a source instance."
    End If

    source.CopyToListObject this.table
    source.ExportMetadata HostSheet
    ResetColumnCache
    LoadColumnMetadata
End Sub

'@sub-title Export the managed table to a plain worksheet range.
'@param targetSheet Worksheet receiving the exported values.
Public Sub ExportToWorksheet(ByVal targetSheet As Worksheet)
    ValidateReady
    If targetSheet Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "ExportToWorksheet requires a target sheet."
    End If

    TableWrapper.Export targetSheet
    ExportMetadata targetSheet
End Sub

'@sub-title Import table data from a plain worksheet range.
'@param sourceSheet Worksheet containing headers and rows matching the variables schema.
Public Sub ImportFromWorksheet(ByVal sourceSheet As Worksheet)
    Dim usedRange As Range
    Dim headerRow As Long
    Dim headerColumn As Long
    Dim sourceData As IDataSheet

    ValidateReady
    If sourceSheet Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "ImportFromWorksheet requires a source sheet."
    End If

    On Error Resume Next
        Set usedRange = sourceSheet.UsedRange
    On Error GoTo 0

    If usedRange Is Nothing Then Exit Sub

    headerRow = usedRange.Row
    headerColumn = usedRange.Column
    If headerRow <= 0 Or headerColumn <= 0 Then Exit Sub

    Set sourceData = DataSheet.Create(sourceSheet, headerRow, headerColumn)

    TableWrapper.Import sourceData

    EnsureRequiredColumns this.table
    ResetColumnCache
    PersistColumnMetadata
End Sub

'@sub-title Copy the managed table content into a supplied ListObject.
'@param targetTable ListObject receiving headers, values, and formatting.
Public Sub CopyToListObject(ByVal targetTable As ListObject)
    Dim targetWrapper As ICustomTable

    ValidateReady
    ValidateListObject targetTable

    Set targetWrapper = CustomTable.Create(targetTable)
    targetWrapper.Import TableWrapper, keepSourceHeaders:=True

    On Error Resume Next
        targetWrapper.Table.Name = this.table.Name
        targetWrapper.Table.TableStyle = this.table.TableStyle
    On Error GoTo 0

    EnsureRequiredColumns targetWrapper.Table
End Sub

'@sub-title Export the hidden metadata required to resolve translated columns.
'@param targetSheet Worksheet receiving the hidden name definitions.
Public Sub ExportMetadata(ByVal targetSheet As Worksheet)
    Dim store As IHiddenNames

    ValidateReady
    If targetSheet Is Nothing Then Exit Sub

    Set store = HiddenNameStore()
    If store Is Nothing Then Exit Sub

    store.ExportNames targetSheet, NAME_PREFIX
End Sub

'@sub-title Expose the managed ListObject.
Public Property Get Table() As ListObject
    ValidateReady
    Set Table = this.table
End Property

'@sub-title Indicates whether initialisation logic already executed.
Public Property Get Initialised() As Boolean
    Initialised = this.initialised
End Property

'@sub-title Return the range for the supplied column header, optionally including headers.
Public Function DataRange(ByVal headerName As String, _
                          Optional ByVal includeHeaders As Boolean = False) As Range
    ValidateReady
    Set DataRange = ColumnRange(headerName, includeHeaders)
End Function


'@section Internal helpers
'===============================================================================
Public Sub AssignListObject(ByVal listObject As ListObject)
    Set this.table = listObject
    Set this.wrapper = Nothing
    Set this.hiddenNames = Nothing
    this.initialised = False
    ResetColumnCache
    LoadColumnMetadata
End Sub

Private Sub LoadColumnMetadata()
    Dim store As IHiddenNames

    Set store = HiddenNameStore()
    If store Is Nothing Then Exit Sub

    this.initialised = store.ValueAsBoolean(NAME_INITIALISED, False)
    CacheColumnIndex HEADER_ORDER, store.ValueAsLong(NAME_COL_ORDER, 0)
    CacheColumnIndex HEADER_SECTION, store.ValueAsLong(NAME_COL_SECTION, 0)
    CacheColumnIndex HEADER_NAME, store.ValueAsLong(NAME_COL_NAME, 0)
    CacheColumnIndex HEADER_LABEL, store.ValueAsLong(NAME_COL_LABEL, 0)
    CacheColumnIndex HEADER_DEFAULT_CHOICE, store.ValueAsLong(NAME_COL_DEFAULT_CHOICE, 0)
    CacheColumnIndex HEADER_CHOICES, store.ValueAsLong(NAME_COL_CHOICES, 0)
    CacheColumnIndex HEADER_DEFAULT_STATUS, store.ValueAsLong(NAME_COL_DEFAULT_STATUS, 0)
    CacheColumnIndex HEADER_COMMENTS, store.ValueAsLong(NAME_COL_COMMENTS, 0)
End Sub

Private Sub PersistColumnMetadata()
    Dim store As IHiddenNames

    Set store = HiddenNameStore()
    If store Is Nothing Then Exit Sub

    PersistColumnIndex store, NAME_COL_ORDER, HEADER_ORDER
    PersistColumnIndex store, NAME_COL_SECTION, HEADER_SECTION
    PersistColumnIndex store, NAME_COL_NAME, HEADER_NAME
    PersistColumnIndex store, NAME_COL_LABEL, HEADER_LABEL
    PersistColumnIndex store, NAME_COL_DEFAULT_CHOICE, HEADER_DEFAULT_CHOICE
    PersistColumnIndex store, NAME_COL_CHOICES, HEADER_CHOICES
    PersistColumnIndex store, NAME_COL_DEFAULT_STATUS, HEADER_DEFAULT_STATUS
    PersistColumnIndex store, NAME_COL_COMMENTS, HEADER_COMMENTS

    store.EnsureName NAME_INITIALISED, True, HiddenNameTypeBoolean
    store.SetValue NAME_INITIALISED, True
    this.initialised = True
End Sub

Private Sub PersistColumnIndex(ByVal store As IHiddenNames, _
                               ByVal nameId As String, _
                               ByVal headerName As String)
    Dim columnIndex As Long

    columnIndex = ColumnIndex(headerName)
    If columnIndex <= 0 Then Exit Sub

    store.EnsureName nameId, columnIndex, HiddenNameTypeLong
    store.SetValue nameId, columnIndex
    CacheColumnIndex headerName, columnIndex
End Sub

Private Function CopyIntoSheet(ByVal targetSheet As Worksheet) As ListObject
    Dim tableName As String
    Dim exportedTable As ListObject
    Dim exportRange As Range

    tableName = this.table.Name

    targetSheet.Cells.Clear
    TableWrapper.Export targetSheet

    On Error Resume Next
        Set exportRange = targetSheet.UsedRange
    On Error GoTo 0
    If exportRange Is Nothing Then Exit Function

    Set exportedTable = targetSheet.ListObjects.Add(xlSrcRange:=exportRange, XlListObjectHasHeaders:=xlYes)

    On Error Resume Next
        exportedTable.Name = tableName
        exportedTable.TableStyle = this.table.TableStyle
    On Error GoTo 0

    EnsureRequiredColumns exportedTable

    Set CopyIntoSheet = exportedTable
End Function

Private Function ResolveOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set ResolveOrCreateSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ResolveOrCreateSheet Is Nothing Then
        Set ResolveOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ResolveOrCreateSheet.Name = sheetName
    End If
End Function

Private Sub ApplyDropdownValidation(ByVal dropdowns As IDropdownLists)
    Dim statusRange As Range
    Dim values As BetterArray

    Set statusRange = ColumnRange(HEADER_DEFAULT_STATUS)
    If statusRange Is Nothing Then Exit Sub

    If dropdowns Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "Initialise requires an IDropdownLists dependency."
    End If

    If Not dropdowns.Exists(STATUS_DROPDOWN_NAME) Then
        Set values = New BetterArray
        values.LowerBound = 1
        values.Push STATUS_MANDATORY, STATUS_VISIBLE, STATUS_HIDDEN
        dropdowns.Add values, STATUS_DROPDOWN_NAME
    End If

    dropdowns.SetValidation statusRange, STATUS_DROPDOWN_NAME, alertType:="info"
End Sub

Private Sub ApplyDuplicateFormatting()
    Dim nameRange As Range
    Dim formatCondition As FormatCondition

    Set nameRange = ColumnRange(HEADER_NAME)
    If nameRange Is Nothing Then Exit Sub

    On Error Resume Next
        nameRange.FormatConditions.Delete
    On Error GoTo 0

    On Error Resume Next
        Set formatCondition = nameRange.FormatConditions.AddUniqueValues
    On Error GoTo 0

    If formatCondition Is Nothing Then Exit Sub

    formatCondition.DupeUnique = xlDuplicate
    formatCondition.Interior.Color = DUPLICATE_FILL_COLOR
    formatCondition.Font.Color = DUPLICATE_FONT_COLOR
End Sub

Private Function ColumnRange(ByVal headerName As String, Optional ByVal includeHeaders As Boolean = False) As Range
    Dim columnIndex As Long
    Dim listColumn As ListColumn

    columnIndex = ColumnIndex(headerName)
    If columnIndex <= 0 Then Exit Function

    Set listColumn = this.table.ListColumns(columnIndex)
    If listColumn Is Nothing Then Exit Function

    If includeHeaders Then
        Set ColumnRange = listColumn.Range
    Else
        Set ColumnRange = listColumn.DataBodyRange
    End If
End Function

Private Sub ApplyVariableDropdownValidation(ByVal targetRange As Range, ByVal listName As String)
    Dim formulaText As String

    If targetRange Is Nothing Then Exit Sub
    listName = Trim$(listName)
    If LenB(listName) = 0 Then Exit Sub

    formulaText = "=" & listName

    On Error Resume Next
        targetRange.Validation.Delete
        targetRange.Validation.Add Type:=xlValidateList, _
                                   AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, _
                                   Formula1:=formulaText
        targetRange.Validation.IgnoreBlank = True
        targetRange.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub

Private Function ColumnIndex(ByVal headerName As String) As Long
    Dim cached As Long
    Dim store As IHiddenNames
    Dim nameId As String
    Dim column As ListColumn

    cached = CachedColumnIndex(headerName)
    If cached > 0 Then
        ColumnIndex = cached
        Exit Function
    End If

    Set store = HiddenNameStore()
    If Not store Is Nothing Then
        nameId = ColumnHiddenName(headerName)
        If LenB(nameId) > 0 Then
            cached = store.ValueAsLong(nameId, 0)
            If cached > 0 And cached <= this.table.ListColumns.Count Then
                CacheColumnIndex headerName, cached
                ColumnIndex = cached
                Exit Function
            End If
        End If
    End If

    For Each column In this.table.ListColumns
        If StrComp(Trim$(column.Name), headerName, vbTextCompare) = 0 Then
            cached = column.Index
            Exit For
        End If
    Next column

    If cached > 0 Then
        CacheColumnIndex headerName, cached
    End If

    ColumnIndex = cached
End Function

Private Sub CacheColumnIndex(ByVal headerName As String, ByVal columnIndex As Long)
    If columnIndex <= 0 Then Exit Sub

    Select Case headerName
        Case HEADER_ORDER
            this.colOrder = columnIndex
        Case HEADER_SECTION
            this.colSection = columnIndex
        Case HEADER_NAME
            this.colName = columnIndex
        Case HEADER_LABEL
            this.colLabel = columnIndex
        Case HEADER_DEFAULT_CHOICE
            this.colDefaultChoice = columnIndex
        Case HEADER_CHOICES
            this.colChoices = columnIndex
        Case HEADER_DEFAULT_STATUS
            this.colDefaultStatus = columnIndex
        Case HEADER_COMMENTS
            this.colComments = columnIndex
    End Select
End Sub

Private Function CachedColumnIndex(ByVal headerName As String) As Long
    Select Case headerName
        Case HEADER_ORDER
            CachedColumnIndex = this.colOrder
        Case HEADER_SECTION
            CachedColumnIndex = this.colSection
        Case HEADER_NAME
            CachedColumnIndex = this.colName
        Case HEADER_LABEL
            CachedColumnIndex = this.colLabel
        Case HEADER_DEFAULT_CHOICE
            CachedColumnIndex = this.colDefaultChoice
        Case HEADER_CHOICES
            CachedColumnIndex = this.colChoices
        Case HEADER_DEFAULT_STATUS
            CachedColumnIndex = this.colDefaultStatus
        Case HEADER_COMMENTS
            CachedColumnIndex = this.colComments
        Case Else
            CachedColumnIndex = 0
    End Select
End Function

Private Sub ResetColumnCache()
    this.colOrder = 0
    this.colSection = 0
    this.colName = 0
    this.colLabel = 0
    this.colDefaultChoice = 0
    this.colChoices = 0
    this.colDefaultStatus = 0
    this.colComments = 0
End Sub

Private Function ColumnHiddenName(ByVal headerName As String) As String
    Select Case headerName
        Case HEADER_ORDER
            ColumnHiddenName = NAME_COL_ORDER
        Case HEADER_SECTION
            ColumnHiddenName = NAME_COL_SECTION
        Case HEADER_NAME
            ColumnHiddenName = NAME_COL_NAME
        Case HEADER_LABEL
            ColumnHiddenName = NAME_COL_LABEL
        Case HEADER_DEFAULT_CHOICE
            ColumnHiddenName = NAME_COL_DEFAULT_CHOICE
        Case HEADER_CHOICES
            ColumnHiddenName = NAME_COL_CHOICES
        Case HEADER_DEFAULT_STATUS
            ColumnHiddenName = NAME_COL_DEFAULT_STATUS
        Case HEADER_COMMENTS
            ColumnHiddenName = NAME_COL_COMMENTS
        Case Else
            ColumnHiddenName = vbNullString
    End Select
End Function

Private Function HiddenNameStore() As IHiddenNames
    If this.table Is Nothing Then Exit Function

    If this.hiddenNames Is Nothing Then
        Set this.hiddenNames = HiddenNames.Create(HostSheet)
    End If

    Set HiddenNameStore = this.hiddenNames
End Function

Private Function WorkbookHiddenNameStore() As IHiddenNames
    Dim wb As Workbook

    Set wb = HostWorkbook
    If wb Is Nothing Then Exit Function

    On Error Resume Next
        Set WorkbookHiddenNameStore = HiddenNames.Create(wb)
    On Error GoTo 0
End Function

Private Function TableWrapper() As ICustomTable
    If this.wrapper Is Nothing Then
        Set this.wrapper = CustomTable.Create(this.table)
    End If

    Set TableWrapper = this.wrapper
End Function

Private Function HostSheet() As Worksheet
    Set HostSheet = this.table.Parent
End Function

Private Function HostWorkbook() As Workbook
    Dim sh As Worksheet

    Set sh = HostSheet
    If sh Is Nothing Then Exit Function

    Set HostWorkbook = sh.Parent
End Function

Private Function TableRange(Optional ByVal includeHeaders As Boolean = False) As Range
    Dim rng As Range

    Set rng = this.table.DataBodyRange
    If includeHeaders Then
        Set rng = this.table.Range
    End If

    Set TableRange = rng
End Function

Private Function LocateVariableRow(ByVal variableName As String) As Long
    Dim nameRange As Range
    Dim rowCounter As Long
    Dim totalRows As Long

    Set nameRange = ColumnRange(HEADER_NAME)
    If nameRange Is Nothing Then Exit Function

    totalRows = nameRange.Rows.Count
    For rowCounter = 1 To totalRows
        If StrComp(CStr(nameRange.Cells(rowCounter, 1).Value), variableName, vbTextCompare) = 0 Then
            LocateVariableRow = rowCounter
            Exit Function
        End If
    Next rowCounter
End Function

Private Function ValueFor(ByVal variableName As String, ByVal headerName As String) As String
    Dim rowIndex As Long
    Dim targetRange As Range

    variableName = Trim$(variableName)
    If LenB(variableName) = 0 Then Exit Function

    rowIndex = LocateVariableRow(variableName)
    If rowIndex = 0 Then Exit Function

    Set targetRange = ColumnRange(headerName)
    If targetRange Is Nothing Then Exit Function

    ValueFor = CStr(targetRange.Cells(rowIndex, 1).Value)
End Function

Private Sub WriteValue(ByVal rowIndex As Long, ByVal headerName As String, ByVal newValue As String)
    Dim targetRange As Range

    Set targetRange = ColumnRange(headerName)
    If targetRange Is Nothing Then Exit Sub
    If rowIndex <= 0 Or rowIndex > targetRange.Rows.Count Then Exit Sub

    targetRange.Cells(rowIndex, 1).Value = newValue
End Sub

Private Sub TranslateColumn(ByVal trads As ITranslationObject, ByVal headerName As String)
    Dim targetRange As Range

    Set targetRange = ColumnRange(headerName)
    If targetRange Is Nothing Then Exit Sub

    trads.TranslateRange targetRange
End Sub

Private Sub EnsureRequiredColumns(ByVal listObject As ListObject)
    Dim required As Variant
    Dim header As Variant
    Dim store As IHiddenNames

    required = Array(HEADER_ORDER, HEADER_SECTION, HEADER_NAME, HEADER_LABEL, _
                     HEADER_DEFAULT_CHOICE, HEADER_CHOICES, HEADER_DEFAULT_STATUS, HEADER_COMMENTS)

    If listObject Is Nothing Then Exit Sub

    If this.initialised Then
        Set store = HiddenNameStore()
    End If

    For Each header In required
        If this.initialised And Not store Is Nothing Then
            EnsureColumnAtStoredIndex listObject, CStr(header), store
        Else
            EnsureColumnExistsSimple listObject, CStr(header)
        End If
    Next header
End Sub

Private Sub EnsureColumnExistsSimple(ByVal listObject As ListObject, ByVal headerName As String)
    Dim column As ListColumn
    Dim found As Boolean

    For Each column In listObject.ListColumns
        If StrComp(Trim$(column.Name), headerName, vbTextCompare) = 0 Then
            If StrComp(column.Name, headerName, vbBinaryCompare) <> 0 Then
                column.Name = headerName
            End If
            found = True
            Exit For
        End If
    Next column

    If Not found Then
        With listObject.ListColumns.Add
            .Name = headerName
        End With
    End If
End Sub

Private Sub EnsureColumnAtStoredIndex(ByVal listObject As ListObject, _
                                      ByVal headerName As String, _
                                      ByVal store As IHiddenNames)
    Dim targetIndex As Long
    Dim hiddenName As String
    Dim current As ListColumn
    Dim column As ListColumn

    hiddenName = ColumnHiddenName(headerName)
    targetIndex = store.ValueAsLong(hiddenName, 0)

    If targetIndex <= 0 Then
        EnsureColumnExistsSimple listObject, headerName
        Exit Sub
    End If

    If targetIndex > listObject.ListColumns.Count Then
        Set current = listObject.ListColumns.Add(Position:=targetIndex)
        current.Name = headerName
        Exit Sub
    End If

    For Each column In listObject.ListColumns
        If column.Index = targetIndex Then
            Set current = column
            Exit For
        End If
    Next column

    If current Is Nothing Then
        Set current = listObject.ListColumns.Add(Position:=targetIndex)
    End If

    current.Name = headerName
End Sub

Private Sub PersistInitialisedFlag(ByVal state As Boolean)
    Dim store As IHiddenNames

    Set store = HiddenNameStore()
    If store Is Nothing Then Exit Sub

    store.EnsureName NAME_INITIALISED, state, HiddenNameTypeBoolean
    store.SetValue NAME_INITIALISED, state
    store.EnsureName NAME_SHEET_TAG, "variables", HiddenNameTypeString
    store.SetValue NAME_SHEET_TAG, "variables"
    this.initialised = state
End Sub

Private Sub ValidateReady()
    If this.table Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "MasterSetupVariables requires initialisation through Create."
    End If
End Sub

Private Sub ValidateListObject(ByVal listObject As ListObject)
    If listObject Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "A valid ListObject reference is required."
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise errNumber, CLASS_NAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Sub IMasterSetupVariables_Initialise(ByVal dropdowns As IDropdownLists)
    Initialise dropdowns
End Sub

Private Sub IMasterSetupVariables_RefreshChoices(ByVal variableName As String, ByVal choices As ILLChoices)
    RefreshChoices variableName, choices
End Sub

Private Function IMasterSetupVariables_LabelFor(ByVal variableName As String) As String
    IMasterSetupVariables_LabelFor = LabelFor(variableName)
End Function

Private Function IMasterSetupVariables_DefaultChoiceFor(ByVal variableName As String) As String
    IMasterSetupVariables_DefaultChoiceFor = DefaultChoiceFor(variableName)
End Function

Private Function IMasterSetupVariables_ChoicesFor(ByVal variableName As String) As String
    IMasterSetupVariables_ChoicesFor = ChoicesFor(variableName)
End Function

Private Function IMasterSetupVariables_DefaultStatusFor(ByVal variableName As String) As String
    IMasterSetupVariables_DefaultStatusFor = DefaultStatusFor(variableName)
End Function

Private Function IMasterSetupVariables_SectionFor(ByVal variableName As String) As String
    IMasterSetupVariables_SectionFor = SectionFor(variableName)
End Function

Private Sub IMasterSetupVariables_AddRows(Optional ByVal count As Long = 5)
    AddRows count
End Sub

Private Sub IMasterSetupVariables_RemoveRows(Optional ByVal count As Long = 1)
    RemoveRows count
End Sub

Private Sub IMasterSetupVariables_TranslateHeaders(ByVal trads As ITranslationObject)
    TranslateHeaders trads
End Sub

Private Sub IMasterSetupVariables_TranslateContent(ByVal trads As ITranslationObject)
    TranslateContent trads
End Sub

Private Sub IMasterSetupVariables_TranslateAll(ByVal trads As ITranslationObject)
    TranslateAll trads
End Sub

Private Function IMasterSetupVariables_CloneToWorkbook(ByVal targetWorkbook As Workbook, _
                                                       Optional ByVal sheetName As String = vbNullString, _
                                                       Optional ByVal startCell As Range = Nothing) As IMasterSetupVariables
    Set IMasterSetupVariables_CloneToWorkbook = CloneToWorkbook(targetWorkbook, sheetName, startCell)
End Function

Private Sub IMasterSetupVariables_RestoreFrom(ByVal source As IMasterSetupVariables)
    RestoreFrom source
End Sub

Private Sub IMasterSetupVariables_ExportToWorksheet(ByVal targetSheet As Worksheet)
    ExportToWorksheet targetSheet
End Sub

Private Sub IMasterSetupVariables_ImportFromWorksheet(ByVal sourceSheet As Worksheet)
    ImportFromWorksheet sourceSheet
End Sub

Private Sub IMasterSetupVariables_CopyToListObject(ByVal targetTable As ListObject)
    CopyToListObject targetTable
End Sub

Private Sub IMasterSetupVariables_ExportMetadata(ByVal targetSheet As Worksheet)
    ExportMetadata targetSheet
End Sub

Private Property Get IMasterSetupVariables_Table() As ListObject
    Set IMasterSetupVariables_Table = Table
End Property

Private Property Get IMasterSetupVariables_Initialised() As Boolean
    IMasterSetupVariables_Initialised = Initialised
End Property
