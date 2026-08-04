VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLChoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist choice Class"


'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Linelist choice Class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class LLChoices
'@description
'Manages the choice lists a workbook offers to dropdowns and cascading
'selectors. A choice is named by its list name and carries a long label, a
'short label and an order number, and the four columns sit on one worksheet
'reached through a DataSheet. AddChoice appends a list, RemoveChoice deletes
'one, ChoiceExists answers whether a list is there, Sort orders the sheet and
'Translate rewrites the labels. AllChoices, Categories and
'ConcatenateCategories give the lists back. ManageRows, AddRows, InsertRows,
'DeleteRows and RemoveRows work on the rows of the sheet, and Import and Export
'move the sheet between workbooks.
'
'WHAT THE INSTANCE HOLDS
'-------------------------------------------------------------------------------
'A member that needs the four choice columns reads the header row once as a
'block and matches the names in memory for the rest of that call. Nothing is
'held between calls: AddChoice appends rows, RemoveChoice deletes whole rows
'and Import clears the sheet, so the header row and the end row are read again
'on every call.
'
'EXCEL STATE BELONGS TO THE CALLER
'-------------------------------------------------------------------------------
'This class leaves Application.ScreenUpdating, Calculation and EnableEvents
'where it finds them, and builds no ApplicationState.
'@depends DataSheet, CustomTable, Checking, BetterArray, TranslationObject
'@version 1.0 (2026-02-09)


'@section State
'===============================================================================
'Backing fields and constants for the choice object.

Private Type TChoice
    choiceDataSheet As DataSheet
    hasChecks As Boolean
    choiceName As String
    sealed As Boolean
    checkCounter As Long                   'Unique key for each logged message
    internalChecks As Checking             'Where the logged messages are kept
End Type

Private this As TChoice

Private Const LISTNAMECODE As String = "list name"
Private Const TRANSLABELCODE As String = "translated label"
Private Const LABELCODE   As String = "label"
Private Const LABELSHORTCODE As String = "short label"
Private Const ORDERINLISTCODE   As String = "ordering list"
Private Const CLASSNAME   As String = "LLChoices"

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new LLChoices instance
'@details
'Wraps the given worksheet in a DataSheet and returns a configured
'LLChoices instance. The four columns this class works with are
'"list name", "ordering list", "label" and "short label", and AddHeaders
'writes them to the sheet in that order.
'@param choiceWorksheet Worksheet. The worksheet hosting the choices.
'@param choiceStartRow Long. First row of choice data.
'@param choiceStartColumn Long. First column of choice data.
'@param choiceName Optional String. Name for the choice object. Defaults to vbNullString.
'@return LLChoices. The new instance.
'@export
Public Function Create(ByVal choiceWorksheet As Worksheet, _
                       ByVal choiceStartRow As Long, _
                       ByVal choiceStartColumn As Long, _
                       Optional ByVal choiceName As String = vbNullString) As _
                       LLChoices

    Dim customDataSheet As DataSheet
    Dim instance As LLChoices

    Set customDataSheet = DataSheet.Create(choiceWorksheet, choiceStartRow, choiceStartColumn)
    Set instance = New LLChoices

    Set instance.ChoiceData = customDataSheet
    instance.Name = choiceName
    instance.Seal
    Set Create = instance
End Function

'@section Elements
'===============================================================================

'@label:choice-data
'@prop-title Associated DataSheet object
'@return DataSheet. The backing DataSheet.
'@export
Public Property Get ChoiceData() As DataSheet
    Set ChoiceData = this.choiceDataSheet
End Property

'@label:choice-data-set
'@prop-title Assign the backing DataSheet
'@param customDataSheet DataSheet. The DataSheet to assign.
Public Property Set ChoiceData(ByVal customDataSheet As DataSheet)
    GuardNotSealed "ChoiceData"
    Set this.choiceDataSheet = customDataSheet
End Property

'@label:wksh
'@prop-title Worksheet hosting the choice lists
'@return Worksheet. The host worksheet.
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = ChoiceData.Wksh
End Property

'@label:start-row
'@prop-title First row of choice data
'@return Long. The start row.
Public Property Get StartRow() As Long
    StartRow = ChoiceData.DataStartRow
End Property

'@label:start-column
'@prop-title First column of choice data
'@return Long. The start column.
Public Property Get StartColumn() As Long
    StartColumn = ChoiceData.DataStartColumn
End Property

'@section Header lookup
'===============================================================================
'One read of the header row answers every column question this class asks.
'The three private wrappers that used to sit here are gone with the searches
'they paid for: DataRange (which asked DataSheet to find a column by its
'header text), ChoiceLabel and HasValidHeaders. Every member now reads the
'header row once through HeaderValues and takes its column positions, its
'label column and its "are the four columns there" answer from that one
'read. AddChoice used to work out where to append from the row count of a
'DataRange that carried its header row; it now says DataEndRow + 1, which is
'the same number.

'@label:header-values
'@sub-title Read the header row once as a 1-based list of values
'@details
'Range.Value hands back a plain value when the Range is one cell and a
'two-dimensional array otherwise, so both shapes are flattened here and the
'callers do not repeat the test.
'The block is deliberately NOT kept on the instance. Callers write to these
'sheets without going through this object, so it is read again on the next
'call.
'@param firstColumn Long. Out. The worksheet column the header row starts at.
'@return Variant. A 1-based one-dimensional array of header values.
Private Function HeaderValues(ByRef firstColumn As Long) As Variant
    Dim headerRng As Range
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long
    Dim nbColumns As Long

    Set headerRng = ChoiceData.HeaderRange
    firstColumn = headerRng.Column
    nbColumns = headerRng.Columns.Count
    ReDim flat(1 To nbColumns)

    If nbColumns = 1 Then
        flat(1) = headerRng.Cells(1, 1).Value
    Else
        rawValues = headerRng.Value
        For counter = 1 To nbColumns
            flat(counter) = rawValues(1, counter)
        Next
    End If

    HeaderValues = flat
End Function

'@label:column-at
'@sub-title Find one column in a header row already read
'@details
'Whole-string and case-insensitive, which is the rule this class has always
'used: every lookup it made went through DataRange, and DataRange forces
'matchCase:=False. The leftmost match wins, the same way DataSheet answers.
'@param headers Variant. The header values from HeaderValues.
'@param colName String. Column header to look for.
'@param firstColumn Long. The worksheet column the header row starts at.
'@return Long. The worksheet column of the match, or 0 when there is none.
Private Function ColumnAt(ByRef headers As Variant, ByVal colName As String, _
                          ByVal firstColumn As Long) As Long
    Dim counter As Long

    For counter = 1 To UBound(headers)
        If StrComp(CellText(headers(counter)), colName, vbTextCompare) = 0 Then
            ColumnAt = firstColumn + counter - 1
            Exit Function
        End If
    Next
End Function

'@label:cell-text
'@sub-title Read one cell value as text without raising
'@details
'A cell holding an error value (#N/A and the like) makes CStr raise a type
'mismatch. Such a cell can never match a column name or a choice name, so
'it reads as empty text instead of stopping the caller.
'@param cellValue Variant. The raw cell value.
'@return String. The value as text, or empty text.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@label:choice-label-in
'@sub-title Resolve the label column name from a header row already read
'@details
'The translated label column wins when the sheet carries one, otherwise the
'default label column is used.
'@param headers Variant. The header values from HeaderValues.
'@param firstColumn Long. The worksheet column the header row starts at.
'@return String. The resolved column name.
Private Function ChoiceLabelIn(ByRef headers As Variant, ByVal firstColumn As Long) As String
    ChoiceLabelIn = LABELCODE
    If ColumnAt(headers, TRANSLABELCODE, firstColumn) > 0 Then ChoiceLabelIn = TRANSLABELCODE
End Function

'@label:body-block
'@sub-title Build a Range over the body rows of one or more columns
'@details
'This is what DataSheet.DataRange(name, includeHeaders:=False) builds, but
'from a column position that is already known instead of from a second
'search for the header text.
'@param dataObject DataSheet. The backing DataSheet.
'@param firstColumn Long. Left worksheet column.
'@param lastColumn Long. Right worksheet column.
'@return Range. The body block.
Private Function BodyBlock(ByVal dataObject As DataSheet, _
                           ByVal firstColumn As Long, _
                           ByVal lastColumn As Long) As Range
    Dim sh As Worksheet

    Set sh = dataObject.Wksh()
    Set BodyBlock = sh.Range(sh.Cells(dataObject.DataStartRow() + 1, firstColumn), _
                             sh.Cells(dataObject.DataEndRow(), lastColumn))
End Function

'@label:column-values
'@sub-title Read one worksheet column as a 1-based list of values
'@param blockRng Range. A single-column block.
'@return Variant. A 1-based one-dimensional array of cell values.
Private Function ColumnValues(ByVal blockRng As Range) As Variant
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long
    Dim nbRows As Long

    nbRows = blockRng.Rows.Count
    ReDim flat(1 To nbRows)

    If nbRows = 1 Then
        flat(1) = blockRng.Cells(1, 1).Value
    Else
        rawValues = blockRng.Value
        For counter = 1 To nbRows
            flat(counter) = rawValues(counter, 1)
        Next
    End If

    ColumnValues = flat
End Function

'@section Validation
'===============================================================================

'@label:has-all-choice-columns
'@sub-title Check a header row already read for the four choice columns
'@details
'The four ColumnExists calls this replaces cost about 40 worksheet round
'trips, and VBA does not short-circuit And, so all four always ran.
'@param headers Variant. The header values from HeaderValues.
'@param firstColumn Long. The worksheet column the header row starts at.
'@return Boolean. True when all four required columns are found.
Private Function HasAllChoiceColumns(ByRef headers As Variant, ByVal firstColumn As Long) As Boolean
    If ColumnAt(headers, LISTNAMECODE, firstColumn) = 0 Then Exit Function
    If ColumnAt(headers, LABELCODE, firstColumn) = 0 Then Exit Function
    If ColumnAt(headers, LABELSHORTCODE, firstColumn) = 0 Then Exit Function
    If ColumnAt(headers, ORDERINLISTCODE, firstColumn) = 0 Then Exit Function
    HasAllChoiceColumns = True
End Function

'@section Choices
'===============================================================================

'@label:name
'@prop-title Name of the choice object
'@details
'The lazy default is written straight into the backing field, so a Get can
'change the object. It has to be that way: Seal closes the Let, and the
'default can only be worked out once the worksheet is known.
'@return String. The choice object name.
'@export
Public Property Get Name() As String

    If LenB(this.choiceName) = 0 Then
        On Error Resume Next
            this.choiceName = "LLCHOICE"
            this.choiceName = Wksh.Name
        On Error GoTo 0
    End If

    Name = this.choiceName
End Property

'@label:name-set
'@prop-title Assign the choice object name
'@param choiceName String. The name to assign.
Public Property Let Name(ByVal choiceName As String)
    GuardNotSealed "Name"
    this.choiceName = choiceName
End Property

'@label:list-name-range
'@sub-title Body range of the list name column, or Nothing when it is missing
'@details
'ChoiceExists used to resolve this column twice: once through ColumnExists,
'only to be able to answer with a quiet False when the column is missing,
'and once more inside DataRange, which raises instead. The header row is
'read once here and the range is built from the position it gives.
'@return Range. The list name column body, or Nothing.
Private Function ListNameRange() As Range
    Dim headers As Variant
    Dim firstColumn As Long
    Dim listColumn As Long

    headers = HeaderValues(firstColumn)
    listColumn = ColumnAt(headers, LISTNAMECODE, firstColumn)
    If listColumn = 0 Then Exit Function

    Set ListNameRange = BodyBlock(ChoiceData, listColumn, listColumn)
End Function

'@label:find-choice-cell
'@sub-title Find the first cell of the list name column holding a choice name
'@details
'The four search arguments are kept exactly as they have always been. Three
'things follow from them and they are known, not accidents: Range.Find
'reuses whatever LookIn, SearchOrder and SearchDirection the last search
'anywhere in Excel used, including the user's own Ctrl+F; it always reads
'*, ? and ~ as wildcards even under xlWhole; and it skips hidden rows.
'@param choiceName String. Choice name to search for.
'@param listRng Range. The list name column body.
'@return Range. The matching cell, or Nothing.
Private Function FindChoiceCell(ByVal choiceName As String, ByVal listRng As Range) As Range
    Set FindChoiceCell = listRng.Find( _
        What:=choiceName, _
        lookAt:=xlWhole, _
        MatchCase:=True, _
        After:=listRng.Cells(listRng.Rows.Count, 1) _
    )
End Function

'@label:choice-exists
'@sub-title Check whether a choice exists
'@details
'Answers a quiet False when the list name column is missing or the name is
'empty. Matching is case-SENSITIVE here, which is not the rule Categories
'uses -- AutoFilter criteria are case-insensitive and cannot be made
'otherwise.
'@param choiceName String. Choice name to search for.
'@return Boolean. True when the choice is found.
'@export
Public Property Get ChoiceExists(ByVal choiceName As String) As Boolean
    Dim listRng As Range

    'Empty choicename should return a False
    If LenB(choiceName) = 0 Then Exit Property

    'If the column listName() is not found, return false
    Set listRng = ListNameRange()
    If listRng Is Nothing Then Exit Property

    ChoiceExists = Not (FindChoiceCell(choiceName, listRng) Is Nothing)
End Property

'@label:sort
'@sub-title Sort choices in the worksheet
'@details
'Sorts on the list name first, then on the ordering column. The header row
'is read once and the three ranges are built from the positions it gives;
'the old shape asked the worksheet for the same header row four times.
'@export
Public Sub Sort()

    Dim headers As Variant
    Dim firstColumn As Long
    Dim listColumn As Long
    Dim orderColumn As Long
    Dim dataObject As DataSheet
    Dim sh As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim lastColumn As Long

    headers = HeaderValues(firstColumn)
    If Not HasAllChoiceColumns(headers, firstColumn) Then
        LogMissingHeaders "sorting"
        Exit Sub
    End If

    listColumn = ColumnAt(headers, LISTNAMECODE, firstColumn)
    orderColumn = ColumnAt(headers, ORDERINLISTCODE, firstColumn)

    Set dataObject = ChoiceData
    Set sh = dataObject.Wksh
    firstRow = dataObject.DataStartRow()
    lastRow = dataObject.DataEndRow()
    lastColumn = firstColumn + UBound(headers) - 1

    With sh
        .Range(.Cells(firstRow, firstColumn), .Cells(lastRow, lastColumn)).Sort _
            key1:=.Range(.Cells(firstRow + 1, listColumn), .Cells(lastRow, listColumn)), _
            Order1:=xlAscending, _
            key2:=.Range(.Cells(firstRow + 1, orderColumn), .Cells(lastRow, orderColumn)), _
            order2:=xlAscending, _
            Header:=xlYes
    End With
End Sub

'@label:add-choice
'@sub-title Add a new choice with its categories
'@details
'Appends a new choice entry with long labels and optional short labels.
'Silently exits when the choice already exists or the data is empty.
'The four columns are resolved from one header read, and the list name and
'the ordering number are written as one block each. They used to be written
'cell by cell, with both column numbers read back from the worksheet on
'every pass of the loop.
'Known and unchanged: on a sheet holding only headers DataEndRow answers
'one row below the header, so the first append lands one row low and leaves
'a blank row behind.
'@param choiceName String. Name of the choice to add.
'@param longLabelData BetterArray. Long label values for the categories.
'@param shortLabelData Optional Object. BetterArray of short label values. Defaults to Nothing.
'@export
Public Sub AddChoice(ByVal choiceName As String, _
                      ByVal longLabelData As BetterArray, _
                      Optional ByVal shortLabelData As Object = Nothing)

    Dim headers As Variant
    Dim firstColumn As Long
    Dim listColumn As Long
    Dim labelColumn As Long
    Dim orderColumn As Long
    Dim shortLabelColumn As Long
    Dim dataObject As DataSheet
    Dim choicewksh As Worksheet
    Dim appendRow As Long
    Dim counter As Long
    Dim nbCategories As Long
    Dim hasShortLabel As Boolean
    Dim listBlock() As Variant
    Dim orderBlock() As Variant

    hasShortLabel = False

    'Test if the choice already exists
    If ChoiceExists(choiceName) Then
         LogInfo "The choice " & choiceName & "already exists and is not added"
         Exit Sub
    End If

    'Test if categories table is empty
    If longLabelData.Length = 0 Then
        LogInfo "The provided choice " & choiceName & " does not have any categories and is not added"
        Exit Sub
    End If

    If TypeName(shortLabelData) = "BetterArray" Then
        If (longLabelData.Length <> shortLabelData.Length) Or _
           (longLabelData.LowerBound <> shortLabelData.LowerBound) Then
            LogInfo "long and short label tables should have the same length,  no categories added for choice " & choiceName
            Exit Sub
        Else
            hasShortLabel = True
        End If
    End If

    'Test if the headers (converted or not exists)
    headers = HeaderValues(firstColumn)
    If Not HasAllChoiceColumns(headers, firstColumn) Then
        LogMissingHeaders "adding the choice " & choiceName
        Exit Sub
    End If

    listColumn = ColumnAt(headers, LISTNAMECODE, firstColumn)
    labelColumn = ColumnAt(headers, ChoiceLabelIn(headers, firstColumn), firstColumn)
    orderColumn = ColumnAt(headers, ORDERINLISTCODE, firstColumn)
    If hasShortLabel Then shortLabelColumn = ColumnAt(headers, LABELSHORTCODE, firstColumn)

    Set dataObject = ChoiceData
    Set choicewksh = dataObject.Wksh

    'One row past the last row that holds data. The old shape reached the
    'same number the long way round: StartRow plus the row count of a range
    'that carried its header row, which is DataEndRow + 1.
    appendRow = dataObject.DataEndRow() + 1

    nbCategories = longLabelData.Length

    longLabelData.ToExcelRange choicewksh.Cells(appendRow, labelColumn)
    If hasShortLabel Then shortLabelData.ToExcelRange choicewksh.Cells(appendRow, shortLabelColumn)

    ReDim listBlock(1 To nbCategories, 1 To 1)
    ReDim orderBlock(1 To nbCategories, 1 To 1)
    For counter = 1 To nbCategories
        listBlock(counter, 1) = choiceName
        orderBlock(counter, 1) = counter
    Next

    choicewksh.Cells(appendRow, listColumn).Resize(nbCategories, 1).Value = listBlock
    choicewksh.Cells(appendRow, orderColumn).Resize(nbCategories, 1).Value = orderBlock
End Sub

'@label:remove-one-choice
'@sub-title Remove every row of a single choice
'@details
'The caller has already sorted the sheet, so every row of one choice sits
'in a single unbroken run. The run is found by reading the list name column
'once and matching in memory, and then taken in ONE worksheet delete.
'It used to delete a row at a time, which is one of the most expensive
'things Excel does, and the loop was only correct because Offset(1) and the
'upward shift of the rows below cancelled each other exactly.
'EntireRow.Delete takes the WHOLE worksheet row, on a sheet this class does
'not own on its own. That is unchanged.
'@param choiceName String. The choice to remove.
'@param listColumn Long. Worksheet column holding the list names.
Private Sub RemoveOneChoice(ByVal choiceName As String, ByVal listColumn As Long)

    Dim dataObject As DataSheet
    Dim sh As Worksheet
    Dim listValues As Variant
    Dim firstRow As Long
    Dim counter As Long
    Dim firstMatch As Long
    Dim lastMatch As Long

    Set dataObject = ChoiceData
    Set sh = dataObject.Wksh
    firstRow = dataObject.DataStartRow() + 1
    listValues = ColumnValues(BodyBlock(dataObject, listColumn, listColumn))

    For counter = 1 To UBound(listValues)
        If StrComp(CellText(listValues(counter)), choiceName, vbBinaryCompare) = 0 Then
            If firstMatch = 0 Then firstMatch = counter
            lastMatch = counter
        ElseIf firstMatch > 0 Then
            Exit For
        End If
    Next

    If firstMatch = 0 Then
        LogAndTrace "Choice " & choiceName & " can not be deleted, it does not exists"
        Exit Sub
    End If

    sh.Range(sh.Cells(firstRow + firstMatch - 1, listColumn), _
             sh.Cells(firstRow + lastMatch - 1, listColumn)).EntireRow.Delete
End Sub

'@label:remove-choice
'@sub-title Remove one or multiple choices
'@details
'The sheet is sorted ONCE here, for the whole call. The sort used to run
'inside the per-name routine, so removing five names paid five full
'worksheet sorts.
'@param choiceNames String. Choice names separated by sep.
'@param sep Optional String. Delimiter. Defaults to ";".
'@export
Public Sub RemoveChoice(ByVal choiceNames As String, Optional ByVal sep As String = ";")
    Dim namesData As BetterArray
    Dim counter As Long
    Dim headers As Variant
    Dim firstColumn As Long
    Dim listColumn As Long

    headers = HeaderValues(firstColumn)
    If Not HasAllChoiceColumns(headers, firstColumn) Then
        LogMissingHeaders "removing the choices " & choiceNames
        Exit Sub
    End If

    listColumn = ColumnAt(headers, LISTNAMECODE, firstColumn)

    'Sorting groups every list name into one unbroken run, which is what
    'lets each name go in a single row delete. Deleting rows never moves a
    'column, so the position found above stays right for the whole loop.
    Sort

    Set namesData = New BetterArray
    namesData.Items = Split(choiceNames, sep)

    For counter = namesData.LowerBound To namesData.UpperBound
        RemoveOneChoice CStr(namesData.Item(counter)), listColumn
    Next
End Sub

'@label:translate
'@sub-title Translate the choice worksheet labels
'@details
'The BODY of the two label columns only. TranslateRange writes back every
'cell it is handed, and this used to hand it ranges that started on the
'header cell, because DataRange defaults includeHeaders to True. A
'translation table holding a tag called "label" or "short label" therefore
'replaced the HEADER with its translation, and from that moment on every
'member of this class that needs the four headers left in silence and
'Categories came back empty.
'@param trads TranslationObject. Translation provider.
'@export
Public Sub Translate(ByVal trads As TranslationObject)

    Dim headers As Variant
    Dim firstColumn As Long
    Dim labelColumn As Long
    Dim shortLabelColumn As Long
    Dim dataObject As DataSheet

    headers = HeaderValues(firstColumn)
    If Not HasAllChoiceColumns(headers, firstColumn) Then
        LogMissingHeaders "translating the labels"
        Exit Sub
    End If

    labelColumn = ColumnAt(headers, ChoiceLabelIn(headers, firstColumn), firstColumn)
    shortLabelColumn = ColumnAt(headers, LABELSHORTCODE, firstColumn)
    Set dataObject = ChoiceData

    'Translate the label
    trads.TranslateRange BodyBlock(dataObject, labelColumn, labelColumn)

    'Translate the short label
    trads.TranslateRange BodyBlock(dataObject, shortLabelColumn, shortLabelColumn)
End Sub

'@label:manage-rows
'@sub-title Add or remove rows via a single toggle
'@details
'Works on the FIRST ListObject of the host sheet, whatever it is, and
'leaves in silence when there is none. The two branches are not symmetric:
'adding uses CustomTable's own default of 5 rows, removing takes 1.
'@param del Optional Boolean. When True, removes rows. Defaults to False.
'@export
Public Sub ManageRows(Optional ByVal del As Boolean = False)
    Dim sh As Worksheet
    Dim tableList As ListObject
    Dim tabl As CustomTable

    Set sh = Wksh()

    On Error Resume Next
    Set tableList = sh.ListObjects(1)
    On Error GoTo 0
    If tableList Is Nothing Then Exit Sub

    Set tabl = CustomTable.Create(tableList)
    If del Then
        tabl.RemoveRows totalCount:=1
    Else
        tabl.AddRows
    End If
End Sub

'@label:add-rows
'@sub-title Add rows to the choice table
'@export
Public Sub AddRows()
    ManageRows del:=False
End Sub

'@label:insert-rows
'@sub-title Insert rows at the selected position
'@details
'Mirrors the selection height inside the choices ListObject. When the
'selection is invalid the attempt is logged and ignored.
'@param targetCell Range. Selection anchoring the insertion point.
'@param insertShift Optional Boolean. When True, inserts worksheet rows. Defaults to False.
'@export
Public Sub InsertRows(ByVal targetCell As Range, _
                       Optional ByVal insertShift As Boolean = False)
    Dim tableList As ListObject
    Dim tabl As CustomTable

    Set tableList = ResolveRowTarget(targetCell, "insert")
    If tableList Is Nothing Then Exit Sub

    Set tabl = CustomTable.Create(tableList)
    tabl.InsertRowsAt targetCell:=targetCell, insertShift:=insertShift
End Sub

'@label:delete-rows
'@sub-title Delete rows intersecting the selection
'@param targetCell Range. Selection identifying rows to delete.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion. Defaults to False.
'@export
Public Sub DeleteRows(ByVal targetCell As Range, _
                       Optional ByVal includeIds As Boolean = True, _
                       Optional ByVal forceShift As Boolean = False)
    Dim tableList As ListObject
    Dim tabl As CustomTable

    Set tableList = ResolveRowTarget(targetCell, "delete")
    If tableList Is Nothing Then Exit Sub

    Set tabl = CustomTable.Create(tableList)
    tabl.DeleteRowsAt targetCell:=targetCell, _
                      includeIds:=includeIds, _
                      forceShift:=forceShift
End Sub

'@label:remove-rows
'@sub-title Remove empty rows from the choice table
'@export
Public Sub RemoveRows()
    ManageRows del:=True
End Sub

'@label:resolve-row-target
'@sub-title Validate a selection and resolve its ListObject
'@details
'The "the worksheet could not be resolved" branch looks unreachable and is
'not: targetCell can point at a sheet that has since been deleted, and then
'reading its Worksheet raises. The trap above it is what keeps
'selectionSheet Nothing, and the next test would read .Name off Nothing, so
'both guards stay.
'This resolves the table from the SELECTION, while ManageRows takes the
'first table of the sheet. On a choices sheet carrying more than one table
'those are two different answers.
'@param targetCell Range. The user-selected cell to validate.
'@param actionLabel String. Human-readable action name used in warning messages.
'@return ListObject. The ListObject, or Nothing when invalid.
Private Function ResolveRowTarget(ByVal targetCell As Range, ByVal actionLabel As String) As ListObject
    'Validate that the provided selection belongs to the choices worksheet and intersects a listobject.
    Dim hostSheet As Worksheet
    Dim selectionSheet As Worksheet
    Dim normalizedAction As String

    normalizedAction = LCase$(actionLabel)
    Set hostSheet = Wksh()

    If targetCell Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the target selection is missing", checkingWarning
        Exit Function
    End If

    On Error Resume Next
        Set selectionSheet = targetCell.Worksheet
    On Error GoTo 0

    If selectionSheet Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the worksheet could not be resolved", checkingWarning
        Exit Function
    End If

    If Not (selectionSheet Is hostSheet) Then
        LogInfo "Unable to " & normalizedAction & " rows because the selection is on " & selectionSheet.Name, checkingWarning
        Exit Function
    End If

    On Error Resume Next
        Set ResolveRowTarget = targetCell.ListObject
    On Error GoTo 0

    If ResolveRowTarget Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the selection is outside the choices table", checkingWarning
    End If
End Function

'@label:all-choices
'@prop-title All distinct choice names
'@details
'Worksheet order is kept: the first time a name is seen is where it lands.
'The column is read as ONE block and the names are de-duplicated in memory.
'It used to read every cell TWICE from the worksheet, once to ask whether
'the value was already held and once to push it, and the "already held"
'test copied the whole internal array on every call. On a 2,000-row choices
'sheet that was about 8,000 worksheet crossings.
'@return BetterArray. The distinct choice names. Never Nothing.
'@export
Public Property Get AllChoices() As BetterArray
    Dim lstChoi As BetterArray
    Dim headers As Variant
    Dim firstColumn As Long
    Dim listColumn As Long
    Dim listValues As Variant
    Dim seen() As String
    Dim currentKey As String
    Dim alreadySeen As Boolean
    Dim nbSeen As Long
    Dim nbRows As Long
    Dim counter As Long
    Dim index As Long

    'Assigned before any exit. This property used to leave without ever
    'writing the property name, so the caller got Nothing and the next line
    'raised a bare "Object variable not set".
    Set lstChoi = New BetterArray
    Set AllChoices = lstChoi

    headers = HeaderValues(firstColumn)
    If Not HasAllChoiceColumns(headers, firstColumn) Then
        LogMissingHeaders "listing the choices"
        Exit Property
    End If

    listColumn = ColumnAt(headers, LISTNAMECODE, firstColumn)
    listValues = ColumnValues(BodyBlock(ChoiceData, listColumn, listColumn))
    nbRows = UBound(listValues)
    ReDim seen(1 To nbRows)

    For counter = 1 To nbRows
        currentKey = CellText(listValues(counter))
        alreadySeen = False

        'A choice's rows sit together on the sheet, so the name just kept is
        'almost always the same one. Testing it first turns the scan below
        'into a single compare for nearly every row.
        If nbSeen > 0 Then alreadySeen = (StrComp(seen(nbSeen), currentKey, vbBinaryCompare) = 0)

        If Not alreadySeen Then
            For index = 1 To nbSeen
                If StrComp(seen(index), currentKey, vbBinaryCompare) = 0 Then
                    alreadySeen = True
                    Exit For
                End If
            Next
        End If

        If Not alreadySeen Then
            nbSeen = nbSeen + 1
            seen(nbSeen) = currentKey
            lstChoi.Push listValues(counter)
        End If
    Next
End Property


'@section Categories
'===============================================================================

'@label:categories
'@prop-title Categories of a choice
'@details
'Returns the long label categories for the specified choice. When
'useShortlabels is True, short labels replace long labels where
'available. Falls back to long labels for missing short entries.
'Matching on the choice name is case-INSENSITIVE, because AutoFilter
'criteria are and cannot be made otherwise. There used to be a full
'Range.Find in front of the filter, only to decide whether to bother
'filtering; the filter already answers "no rows matched" with a
'zero-length array. Dropping it means a sheet holding both list_a and
'List_A now answers as if they were one name. That is a data fault on the
'sheet, and it is also one of the ways the filter itself used to go wrong.
'A missing header is logged and answered with an empty array, the way every
'other member of this class answers. It used to raise out of the class.
'@param choiceName String. The choice to retrieve categories from.
'@param useShortlabels Optional Boolean. When True, uses short labels. Defaults to False.
'@return BetterArray. The category values. Never Nothing. Safe to change in place.
'@export
Public Property Get Categories(ByVal choiceName As String, _
                               Optional ByVal useShortlabels As Boolean = False) _
                               As BetterArray

    Dim dataObject As DataSheet
    Dim lngcat As BetterArray                    'Long and short betterarrays returned
    Dim shrtcat As BetterArray
    Dim headers As Variant
    Dim firstColumn As Long
    Dim counter As Long

    Set lngcat = New BetterArray
    lngcat.LowerBound = 1
    Set Categories = lngcat

    'An empty name never named a choice. It is answered here, because an
    'empty AutoFilter criteria means "blank cells" to Excel, not "nothing".
    If LenB(choiceName) = 0 Then Exit Property

    headers = HeaderValues(firstColumn)
    If ColumnAt(headers, LISTNAMECODE, firstColumn) = 0 Then
        LogMissingHeaders "reading the categories of " & choiceName
        Exit Property
    End If
    If ColumnAt(headers, LABELCODE, firstColumn) = 0 Then
        LogMissingHeaders "reading the categories of " & choiceName
        Exit Property
    End If
    If useShortlabels Then
        If ColumnAt(headers, LABELSHORTCODE, firstColumn) = 0 Then
            LogMissingHeaders "reading the short labels of " & choiceName
            Exit Property
        End If
    End If

    Set dataObject = ChoiceData

    'FilterData already hands back a Clone that nothing else holds, so the
    'result goes straight out and is safe for the caller to change in place.
    Set lngcat = dataObject.FilterData( _
                            varName:=LISTNAMECODE, _
                            criteriaName:=choiceName, _
                            returnedColumnName:=LABELCODE, _
                            matchCase:=False, _
                            includeHeaders:=False)

    'Only look for shortlabels in case there is a need
    If useShortlabels Then
        'Short Categories values
        Set shrtcat = dataObject.FilterData( _
            varName:=LISTNAMECODE, _
            criteriaName:=choiceName, _
            returnedColumnName:=LABELSHORTCODE, _
            matchCase:=False, _
            includeHeaders:=False _
        )

        'In case the long and short categories don't have the same
        'length, return the long one, because of a probable error
        If lngcat.Length = shrtcat.Length Then
            'short categories and long categories have the same length, replace some
            'of the long values by the short one (non empty short ones)
            For counter = lngcat.LowerBound To lngcat.UpperBound
                If shrtcat.Item(counter) <> vbNullString Then
                    lngcat.Item(counter) = shrtcat.Item(counter)
                End If
            Next
        End If
    End If

    Set Categories = lngcat
End Property


'@label:concatenate-categories
'@prop-title Concatenate all categories of a choice
'@details
'Joins the categories with a separator. Optionally translates each
'category via the supplied TranslationObject.
'@param choiceName String. The choice whose categories to concatenate.
'@param sep Optional String. Separator. Defaults to " | ".
'@param trads Optional Object. A TranslationObject. Defaults to Nothing.
'@return String. The concatenated categories.
'@export
Public Property Get ConcatenateCategories( _
        ByVal choiceName As String, _
        Optional ByVal sep As String = " | ", _
        Optional ByVal trads As Object = Nothing) As String

    Dim lst As BetterArray
    Dim counter As Long

    Set lst = Categories(choiceName)

    If lst.Length = 0 Then Exit Property

    If Not trads Is Nothing Then
        If StrComp(TypeName(trads), "TranslationObject", vbTextCompare) = 0 Then
            For counter = lst.LowerBound To lst.UpperBound
                lst.Item(counter) = trads.TranslatedValue(lst.Item(counter))
            Next
        End If
    End If

    ConcatenateCategories = lst.ToString( _
                     separator:=sep, _
                     OpeningDelimiter:=vbNullString, _
                     ClosingDelimiter:=vbNullString, _
                     QuoteStrings:=False)
End Property

'@section DataExchange
'===============================================================================
'Persist the choices sheet or hydrate from external sources.

'@label:add-headers
'@sub-title Write the default headers to the choice worksheet
'@details
'Writes the four standard column headers starting at the configured origin
'cell, in this order: list name, ordering list, label, short label.
'The label header is always the DEFAULT one. A sheet that carried a
'translated label column loses it here.
'@export
Public Sub AddHeaders()

    Dim sh As Worksheet
    Dim cellRng As Range

    Set sh = Wksh()
    Set cellRng = sh.Cells(StartRow(), StartColumn())

    cellRng.Value = LISTNAMECODE
    cellRng.Offset(, 1).Value = ORDERINLISTCODE
    cellRng.Offset(, 2).Value = LABELCODE
    cellRng.Offset(, 3).Value = LABELSHORTCODE
End Sub

'@label:export
'@sub-title Export the current choice object to a workbook
'@details
'Delegates to DataSheet.Export, including hidden names.
'@param Wkb Workbook. Destination workbook.
'@param Hide Long. Worksheet visibility constant. Defaults to xlSheetHidden.
'@export
Public Sub Export(ByVal Wkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    ChoiceData.Export Wkb, Hide:=Hide, includeNames:=True
End Sub

'@label:import
'@sub-title Import choices from a worksheet
'@details
'Optionally clears the sheet and rewrites headers before importing.
'Creates a temporary DataSheet from the source range and delegates to
'DataSheet.Import.
'Two things to know before passing clearSheet:=True. Cells.Clear takes the
'ListObject of the sheet with it, and ManageRows, AddRows, InsertRows and
'DeleteRows all need that table, so they do nothing until something builds
'it again. And the import itself matches source columns to target columns
'case-SENSITIVELY, unlike every lookup this class makes on its own.
'@param fromWksh Worksheet. Source worksheet containing choices.
'@param fromStartRow Long. First data row in the source.
'@param fromStartcol Long. First data column in the source.
'@param clearSheet Boolean. When True, clears the sheet before import. Defaults to False.
'@export
Public Sub Import(ByVal fromWksh As Worksheet, _
                   ByVal fromStartRow As Long, _
                   ByVal fromStartcol As Long, _
                   Optional ByVal clearSheet As Boolean = False)

    Dim impChoiData As DataSheet
    Dim sh As Worksheet

    'clear the choices sheet before imports
    If clearSheet Then
        Set sh = Wksh()
        sh.Cells.Clear
        AddHeaders
    End If

    'Import the data
    Set impChoiData = DataSheet.Create(fromWksh, fromStartRow, fromStartcol)

    ChoiceData.Import impChoiData
End Sub

'@section Checkings
'===============================================================================
'Internal logging helpers for the checking subsystem.

'@label:log-info
'@sub-title Append a diagnostic entry to the internal checking object
'@details
'Lazily creates the Checking instance on first call. Increments the
'internal counter used as the entry key.
'@param label String. The message to record.
'@param scope Byte. Severity level. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "Choices " & Name() & " object informations"

    'First, instanciante the internalCheks Objects
    If this.internalChecks Is Nothing Then Set this.internalChecks = Checking.Create(checkName)

    'Add new informations to the checking object
    this.internalChecks.Add CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:log-and-trace
'@sub-title Log a message and print the debug trace
'@details
'The "log it, show the trace, then leave" tail was written out four times,
'on the two failure paths of the choice removal and on the header guards of
'AllChoices and the removal itself.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level. Defaults to checkingNote.
Private Sub LogAndTrace(ByVal label As String, _
                        Optional ByVal scope As Byte = checkingNote)
    LogInfo label, scope
    ShowDebug
End Sub

'@label:log-missing-headers
'@sub-title Log that a member left because a required header is missing
'@details
'Six members need the choice columns and they used to write this tail three
'different ways: Sort logged nothing at all, Translate logged at
'checkingNote, and the rest at checkingWarning. There is one shape now, at
'checkingWarning, because a missing header always means the caller's work
'was dropped. Sort is the one that changes: it used to leave without a
'word, which gave the user a button that did nothing and said nothing.
'@param actionLabel String. What was not done, in plain words.
Private Sub LogMissingHeaders(ByVal actionLabel As String)
    LogAndTrace "Some headers of the choice " & Name() & " object are missing, " & _
                actionLabel & " was not done", checkingWarning
End Sub

'@label:show-debug
'@sub-title Print diagnostic summary to the Immediate window
Private Sub ShowDebug()
    Debug.Print "Check the informations of choice object " & Name()
End Sub

'@label:has-checkings
'@prop-title Whether the choice object has logged diagnostic messages
'@return Boolean. True when diagnostic entries exist.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@details
'Returns the internal Checking object containing all logged messages.
'Returns Nothing when no entries have been recorded.
'@return Object. A Checking instance, or Nothing.
'@export
Public Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = this.internalChecks
End Property

'@section Instantiation Guard
'===============================================================================
'Set-once sealing for the creation-time setters.

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumb Long. ProjectError code.
'@param errorMessage String. Descriptive message.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, "Error : " & Chr(34) & errorMessage
End Sub

'@label:seal
'@sub-title Seal the instance so setup setters can no longer be written
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against writes after sealing
'@param propName String. Name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub
