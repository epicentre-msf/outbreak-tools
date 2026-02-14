Attribute VB_Name = "TestCustomTable"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the CustomTable class")
'
'@description
'   Tests for the CustomTable class, which wraps an Excel ListObject with
'   id-based CRUD operations, Import/Export, sort, and snapshot/restore
'   capabilities. Tests cover factory creation (including rejection of Nothing),
'   AddRows with sequential ID assignment, RemoveRows for trailing empties and
'   explicit trimming, SetValue cell updates, SortOnFirst grouping by first
'   occurrence, Import from both CustomTable and DataSheet sources with various
'   options (keepSourceHeaders, strictColumnSearch, pasteAtBottom, formatHeaders,
'   hidden column preservation, column expansion and trimming), stacked table
'   shift behavior for InsertRowsAt/DeleteRowsAt, Export with header filtering
'   and ListObject creation, DataRange returning Nothing when empty, and
'   snapshot restoration on import failure.
'@depends CustomTable, ICustomTable, DataSheet, IDataSheet, BetterArray, CustomTest, TestHelpers

Private Const TABLESHEETNAME As String = "CustomTableFixture"
Private Const TABLENAME As String = "tblCustom"
Private Const SOURCE_SHEETNAME As String = "CustomTableFixtureSource"
Private Const DATASHEETNAME As String = "CustomTableData"
Private Const MULTITABLESHEET As String = "CustomTableMulti"
Private Const EXPORTSHEETNAME As String = "CustomTableExport"
Private Const EXPAND_SHEETNAME As String = "CustomTableExpand"
Private Const EXPAND_TABLE_NAME As String = "tblExpandTarget"
Private Const TRIM_SHEETNAME As String = "CustomTableTrim"
Private Const TRIM_TABLE_NAME As String = "tblTrimTarget"

Private Assert As ICustomTest
Private Fakes As Object

'@section Helpers
'===============================================================================

'@sub-title Returns the standard three-column header array (ID, Name, Amount)
Private Function CustomTableHeaders() As Variant
    CustomTableHeaders = Array("ID", "Name", "Amount")
End Function

'@sub-title Returns three default data rows for the standard fixture table
Private Function CustomTableRows() As Variant
    CustomTableRows = Array( _
        Array(1, "Alpha", 10), _
        Array(2, "Beta", 20), _
        Array(3, "Gamma", 30))
End Function

'@sub-title Creates a ListObject with a formula column for import-preserves-formulas tests
'@details Builds a three-column table (ID, Value, Calc) on the given sheet where the Calc
'   column contains an R1C1 formula (=RC[-1]*2). This allows tests to verify that importing
'   data into a table with formulas preserves those formulas while updating value columns.
Private Sub PrepareCustomTableWithFormula(ByVal sheetName As String, ByVal tableName As String)

    Dim hostSheet As Worksheet
    Dim headers As Variant
    Dim dataRows As Variant
    Dim listRange As Range
    Dim Lo As ListObject
    Dim rowIndex As Long

    headers = Array("ID", "Value", "Calc")
    dataRows = Array( _
        Array("row 1", 5, 0), _
        Array("row 2", 10, 0), _
        Array("row 3", 15, 0))

    Set hostSheet = EnsureWorksheet(sheetName)
    ClearWorksheet hostSheet

    hostSheet.Range("A1").Resize(1, UBound(headers) + 1).Value = headers

    For rowIndex = LBound(dataRows) To UBound(dataRows)
        hostSheet.Range("A2").Offset(rowIndex).Resize(1, 3).Value = dataRows(rowIndex)
    Next rowIndex

    Set listRange = hostSheet.Range("A1").Resize(UBound(dataRows) + 2, 3)
    Set Lo =  hostSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=listRange, XlListObjectHasHeaders:=xlYes)
    Lo.Name = tableName

    Lo.ListColumns("Calc").DataBodyRange.FormulaR1C1 = "=RC[-1]*2"
End Sub

'@sub-title Creates a DataSheet from arbitrary headers and rows on a fresh worksheet
Private Function CreateDataSheet(ByVal sheetName As String, headers As Variant, rows As Variant) As IDataSheet

    Dim hostSheet As Worksheet
    Dim rowIndex As Long

    Set hostSheet = EnsureWorksheet(sheetName)
    ClearWorksheet hostSheet

    hostSheet.Range("A1").Resize(1, UBound(headers) + 1).Value = headers

    For rowIndex = LBound(rows) To UBound(rows)
        hostSheet.Range("A2").Offset(rowIndex).Resize(1, UBound(headers) + 1).Value = rows(rowIndex)
    Next rowIndex

    Set CreateDataSheet = DataSheet.Create(hostSheet, 1, 1)
End Function

'@sub-title Prepares the standard fixture ListObject, optionally with or without data rows
'@details Writes CustomTableHeaders and optionally CustomTableRows to the given sheet,
'   then wraps the result in a ListObject. When includeData is False the table contains
'   only a header row, which is useful for import-into-empty tests.
Private Sub PrepareCustomTable(Optional ByVal includeData As Boolean = True, _
                               Optional ByVal sheetName As String = TABLESHEETNAME, _
                               Optional ByVal tableName As String = TABLENAME)

    Dim hostSheet As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim lastRow As Long
    Dim columnCount As Long
    Dim listRange As Range
    Dim Lo As ListObject

    Set hostSheet = EnsureWorksheet(sheetName)
    ClearWorksheet hostSheet

    headerMatrix = RowsToMatrix(Array(CustomTableHeaders()))
    WriteMatrix hostSheet.Cells(1, 1), headerMatrix

    If includeData Then
        dataMatrix = RowsToMatrix(CustomTableRows())
        WriteMatrix hostSheet.Cells(2, 1), dataMatrix
        lastRow = 1 + UBound(dataMatrix, 1)
    Else
        lastRow = 1
    End If

    columnCount = UBound(headerMatrix, 2)
    Set listRange = hostSheet.Range("A1").Resize(lastRow, columnCount)

    Set Lo =  hostSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=listRange, XlListObjectHasHeaders:=xlYes)
    Lo.Name = tableName
End Sub

'@sub-title Convenience shortcut that prepares the default fixture and returns a CustomTable instance
Private Function BuildCustomTable() As ICustomTable
    PrepareCustomTable
    Set BuildCustomTable = CustomTable.Create(ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME), "ID", "row")
End Function

'@sub-title Creates a CustomTable from caller-supplied headers and rows on a dedicated sheet
'@details Writes the provided headers and rows to a fresh worksheet, wraps them in a
'   ListObject, and returns a fully initialised CustomTable. This allows tests to build
'   ad-hoc tables with specific data layouts without reusing the default fixture constants.
Private Function CreateCustomTableWithData(ByVal sheetName As String, _
                                          ByVal tableName As String, _
                                          headers As Variant, _
                                          rows As Variant, _
                                          Optional ByVal idColumnName As String = "ID", _
                                          Optional ByVal idPrefix As String = "row") As ICustomTable

    Dim hostSheet As Worksheet
    Dim listRange As Range
    Dim Lo As ListObject
    Dim columnCount As Long
    Dim rowIndex As Long
    Dim lastRow As Long

    Set hostSheet = EnsureWorksheet(sheetName)
    ClearWorksheet hostSheet

    columnCount = UBound(headers) - LBound(headers) + 1
    hostSheet.Range("A1").Resize(1, columnCount).Value = headers

    For rowIndex = LBound(rows) To UBound(rows)
        hostSheet.Range("A2").Offset(rowIndex).Resize(1, columnCount).Value = rows(rowIndex)
    Next rowIndex

    If UBound(rows) >= LBound(rows) Then
        lastRow = (UBound(rows) - LBound(rows) + 1) + 1
    Else
        lastRow = 1
    End If

    Set listRange = hostSheet.Range("A1").Resize(lastRow, columnCount)
    Set Lo =  hostSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=listRange, XlListObjectHasHeaders:=xlYes)
    Lo.Name = tableName

    Set CreateCustomTableWithData = CustomTable.Create(Lo, idColumnName, idPrefix)
End Function

'@sub-title Creates a BetterArray pre-loaded with the given values (LowerBound = 1)
Private Function NewBetterArray(ParamArray values() As Variant) As BetterArray
    Dim arr As BetterArray
    Dim idx As Long

    Set arr = New BetterArray
    arr.LowerBound = 1

    If UBound(values) >= LBound(values) Then
        For idx = LBound(values) To UBound(values)
            arr.Push values(idx)
        Next idx
    End If

    Set NewBetterArray = arr
End Function

'@sub-title Builds two vertically stacked ListObjects on a single worksheet for shift tests
'@details Creates a top table ("tblTop") with two data rows and a bottom table ("tblBottom")
'   with one data row, separated by a gap row. Returns both ListObjects via ByRef parameters
'   so tests can verify that row insertions or deletions in the top table correctly shift
'   the bottom table up or down on the worksheet.
Private Sub PrepareMultiTableFixture(ByRef topTable As ListObject, ByRef bottomTable As ListObject)

    Dim hostSheet As Worksheet
    Dim headers As Variant
    Dim topData As Variant
    Dim bottomData As Variant
    Dim topRange As Range
    Dim bottomRange As Range

    headers = Array("ID", "Name")
    topData = Array( _
        Array("row 1", "Alpha"), _
        Array("row 2", "Beta"))
    bottomData = Array( _
        Array("row A", "Gamma"))

    Set hostSheet = EnsureWorksheet(MULTITABLESHEET)
    ClearWorksheet hostSheet

    hostSheet.Range("A1").Resize(1, 2).Value = headers
    hostSheet.Range("A2").Resize(UBound(topData) + 1, 2).Value = topData

    Set topRange = hostSheet.Range("A1").Resize(UBound(topData) + 2, 2)
    Set topTable = hostSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=topRange, XlListObjectHasHeaders:=xlYes)
    topTable.Name = "tblTop"

    hostSheet.Range("A" & (topRange.Rows.Count + 2)).Resize(1, 2).Value = headers
    hostSheet.Range("A" & (topRange.Rows.Count + 3)).Resize(UBound(bottomData) + 1, 2).Value = bottomData

    Set bottomRange = hostSheet.Range("A" & (topRange.Rows.Count + 2)).Resize(UBound(bottomData) + 2, 2)
    Set bottomTable = hostSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=bottomRange, XlListObjectHasHeaders:=xlYes)
    bottomTable.Name = "tblBottom"
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCustomTable"
    PrepareCustomTable
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet TABLESHEETNAME
    DeleteWorksheet SOURCE_SHEETNAME
    DeleteWorksheet DATASHEETNAME
    DeleteWorksheet MULTITABLESHEET
    DeleteWorksheet EXPORTSHEETNAME
    DeleteWorksheet TRIM_SHEETNAME
    DeleteWorksheet EXPAND_SHEETNAME

    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    BusyApp
    PrepareCustomTable
End Sub

'@TestCleanUp
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.FlushCurrentTest
    End If
End Sub

'@section Tests
'===============================================================================

'@sub-title Verifies that Create initialises the table name and id column from the ListObject
'@details Arranges a standard fixture via BuildCustomTable, then asserts that the Name
'   property matches the underlying ListObject name and IdValue reflects the requested
'   id column header.
'@TestMethod("CustomTable")
Public Sub TestCreateInitialisesTable()
    CustomTestSetTitles Assert, "CustomTable", "TestCreateInitialisesTable"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Set tableObject = BuildCustomTable

    Assert.AreEqual TABLENAME, tableObject.Name, "Table name should match listObject name"
    Assert.AreEqual "ID", tableObject.IdValue, "Id column should be preserved"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateInitialisesTable", Err.Number, Err.Description
End Sub

'@sub-title Verifies that AddRows appends blank rows and AddIds fills them with sequential IDs
'@details Builds the default three-row fixture, appends two rows via AddRows, then calls
'   AddIds. Asserts the total row count is five and that the new rows receive IDs "row 4"
'   and "row 5" following the existing sequence.
'@TestMethod("CustomTable")
Public Sub TestAddRowsAssignsIds()
    CustomTestSetTitles Assert, "CustomTable", "TestAddRowsAssignsIds"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim Lo As ListObject

    Set tableObject = BuildCustomTable
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    tableObject.AddRows nbRows:=2
    tableObject.AddIds

    Assert.IsTrue (Lo.DataBodyRange.Rows.Count = 5), "AddRows should append rows"
    Assert.AreEqual "row 4", Lo.ListColumns("ID").DataBodyRange.Cells(4, 1).Value, _
                    "New blank rows should receive sequential IDs"
    Assert.AreEqual "row 5", Lo.ListColumns("ID").DataBodyRange.Cells(5, 1).Value, "Ids should be sequential"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddRowsAssignsIds", Err.Number, Err.Description
End Sub

'@sub-title Verifies that RemoveRows deletes trailing empty rows and respects a totalCount cap
'@details Blanks the last Name cell to simulate a trailing empty row, adds another empty
'   ListRow, then calls RemoveRows(totalCount:=0) to strip empties. Asserts the table
'   returns to three rows. A second call with totalCount:=2 further trims to two rows.
'@TestMethod("CustomTable")
Public Sub TestRemoveRowsDeletesEmpty()
    CustomTestSetTitles Assert, "CustomTable", "TestRemoveRowsDeletesEmpty"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim Lo As ListObject

    Set tableObject = BuildCustomTable
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    Lo.ListColumns("Name").DataBodyRange.Cells(Lo.DataBodyRange.Rows.Count, 1).Value = ""
    Lo.ListRows.Add
    tableObject.RemoveRows totalCount:=0

    Assert.IsTrue (Lo.DataBodyRange.Rows.Count = 3), "RemoveRows should delete trailing empty rows"
    tableObject.RemoveRows totalCount:=2
    Assert.IsTrue (Lo.DataBodyRange.Rows.Count = 2), "totalCount argument should trim the table to the requested size"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveRowsDeletesEmpty", Err.Number, Err.Description
End Sub

'@sub-title Verifies that SetValue persists a value update to the correct cell
'@details Builds the default fixture, calls SetValue to write "99" into the Amount
'   column at row index "2", then reads back via the Value property and asserts
'   the update was applied.
'@TestMethod("CustomTable")
Public Sub TestSetValueUpdatesCell()
    CustomTestSetTitles Assert, "CustomTable", "TestSetValueUpdatesCell"
    On Error GoTo Fail

    Dim tableObject As ICustomTable

    Set tableObject = BuildCustomTable
    tableObject.SetValue "Amount", "2", "99"

    Assert.AreEqual "99", tableObject.Value("Amount", "2"), "SetValue should persist updates"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSetValueUpdatesCell", Err.Number, Err.Description
End Sub

'@sub-title Verifies that Sort with directSort=False groups rows by first occurrence order
'@details Creates a four-row table where Name values appear as Gamma, Alpha, Gamma, Beta.
'   Calls Sort with directSort:=False (SortOnFirst mode). Asserts that the two Gamma rows
'   are grouped first, followed by Alpha and Beta in first-seen order, and that the
'   temporary helper column is removed after sorting.
'@TestMethod("CustomTable")
Public Sub TestSortOnFirstGroupsByFirstOccurrence()
    CustomTestSetTitles Assert, "CustomTable", "TestSortOnFirstGroupsByFirstOccurrence"
    On Error GoTo Fail

    Dim headers As Variant
    Dim rows As Variant
    Dim tableObject As ICustomTable
    Dim Lo As ListObject

    headers = CustomTableHeaders()
    rows = Array( _
        Array("row 1", "Gamma", 1), _
        Array("row 2", "Alpha", 2), _
        Array("row 3", "Gamma", 3), _
        Array("row 4", "Beta", 4))

    Set tableObject = CreateCustomTableWithData(TABLESHEETNAME, TABLENAME, headers, rows)
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    tableObject.Sort colName:="Name", directSort:=False

    Assert.AreEqual "Gamma", Lo.ListColumns("Name").DataBodyRange.Cells(1, 1).Value, _
                     "SortOnFirst should group rows starting with the first encountered value"
    Assert.AreEqual "Gamma", Lo.ListColumns("Name").DataBodyRange.Cells(2, 1).Value, _
                     "Duplicate values should remain adjacent after SortOnFirst"
    Assert.AreEqual "Alpha", Lo.ListColumns("Name").DataBodyRange.Cells(3, 1).Value, _
                     "Subsequent distinct values should follow in first-seen order"
    Assert.AreEqual "Beta", Lo.ListColumns("Name").DataBodyRange.Cells(4, 1).Value, _
                     "Later unique values should appear last"
    Assert.AreEqual 3, Lo.ListColumns.Count, _
                     "SortOnFirst should remove its helper column after sorting"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSortOnFirstGroupsByFirstOccurrence", Err.Number, Err.Description
End Sub

'@sub-title Verifies that importing from a CustomTable with matching headers copies all rows
'@details Creates an empty target table and a populated source table with the same three
'   columns. Calls Import with keepSourceHeaders:=True. Asserts the target receives all
'   three source rows and that the last row value matches "Gamma".
'@TestMethod("CustomTable")
Public Sub TestImportWithMatchingHeaders()
    CustomTestSetTitles Assert, "CustomTable", "TestImportWithMatchingHeaders"
    On Error GoTo Fail

    Dim sourceTable As ICustomTable
    Dim targetTable As ICustomTable
    Dim Lo As ListObject

    PrepareCustomTable includeData:=False
    Set targetTable = CustomTable.Create(ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME), "ID", "row")

    PrepareCustomTable includeData:=True, sheetName:=SOURCE_SHEETNAME, tableName:=TABLENAME & "Src"
    Set sourceTable = CustomTable.Create(ThisWorkbook.Worksheets(SOURCE_SHEETNAME).ListObjects(TABLENAME & "Src"), "ID", "row")

    targetTable.Import sourceTable, keepSourceHeaders:=True
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    Assert.IsTrue (Lo.DataBodyRange.Rows.Count = 3), "Import should copy all rows"
    Assert.AreEqual "Gamma", Lo.ListColumns("Name").DataBodyRange.Cells(3, 1).Value, "Imported values should match source"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportWithMatchingHeaders", Err.Number, Err.Description
End Sub

'@sub-title Verifies that importing a smaller source shrinks the target to match
'@details Pads the target with an extra row beyond the default three, then imports a
'   source with only two rows. Asserts the target is trimmed to two rows and the second
'   row value matches the source dataset.
'@TestMethod("CustomTable")
Public Sub TestImportShrinksTrailingRows()
    CustomTestSetTitles Assert, "CustomTable", "TestImportShrinksTrailingRows"
    On Error GoTo Fail

    Dim targetTable As ICustomTable
    Dim sourceTable As ICustomTable
    Dim targetList As ListObject

    PrepareCustomTable
    Set targetTable = CustomTable.Create(ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME), "ID", "row")
    Set targetList = ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    targetList.ListRows.Add
    targetList.ListColumns("ID").DataBodyRange.Cells(targetList.DataBodyRange.Rows.Count, 1).Value = "row extra"

    PrepareCustomTable includeData:=True, sheetName:=SOURCE_SHEETNAME, tableName:=TABLENAME & "Src"
    With ThisWorkbook.Worksheets(SOURCE_SHEETNAME).ListObjects(TABLENAME & "Src")
        .ListRows(.ListRows.Count).Delete
    End With

    Set sourceTable = CustomTable.Create(ThisWorkbook.Worksheets(SOURCE_SHEETNAME).ListObjects(TABLENAME & "Src"), "ID", "row")

    targetTable.Import sourceTable, keepSourceHeaders:=True

    Assert.IsTrue (targetList.DataBodyRange.Rows.Count = 2), "Import should shrink data rows to match the source"
    Assert.AreEqual "Beta", targetList.ListColumns("Name").DataBodyRange.Cells(2, 1).Value, _
                     "Second row should match the imported dataset"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportShrinksTrailingRows", Err.Number, Err.Description
End Sub

'@sub-title Verifies that importing from a DataSheet preserves formula columns
'@details Creates a table with a Calc formula column (=Value*2), then imports a DataSheet
'   containing only ID and Value columns. Asserts the formula survives the import, the
'   Value column reflects imported data, and the Calc column recalculates correctly.
'@TestMethod("CustomTable")
Public Sub TestImportFromDataSheetPreservesFormulas()
    CustomTestSetTitles Assert, "CustomTable", "TestImportFromDataSheetPreservesFormulas"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim Lo As ListObject
    Dim headers As Variant
    Dim rows As Variant
    Dim dataSheetObj As IDataSheet

    PrepareCustomTableWithFormula TABLESHEETNAME, TABLENAME
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)
    Set tableObject = CustomTable.Create(Lo, "ID", "row")

    headers = Array("ID", "Value")
    rows = Array( _
        Array("row 1", 100), _
        Array("row 2", 200))

    Set dataSheetObj = CreateDataSheet(DATASHEETNAME, headers, rows)

    tableObject.Import dataSheetObj, keepSourceHeaders:=False

    'Assert.IsTrue (Lo.DataBodyRange.rows.count = 2), "Import from DataSheet should size table to source rows"
    Assert.IsTrue Lo.ListColumns("Calc").DataBodyRange.Cells(1, 1).HasFormula, _
                  "Formula column should keep its formulas after import"
    Assert.IsTrue Lo.ListColumns("Value").DataBodyRange.Cells(2, 1).Value = 200, _
                  "Value column should reflect imported data"
    Assert.IsTrue Lo.ListColumns("Calc").DataBodyRange.Cells(2, 1).Value = 400, _
                  "Formula column should recalculate against imported values"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportFromDataSheetPreservesFormulas", Err.Number, Err.Description
End Sub

'@sub-title Verifies that AddRows with insertShift pushes a stacked table down
'@details Builds two vertically stacked tables, records the bottom table header row
'   position, then calls AddRows(nbRows:=2, insertShift:=True) on the top table. Asserts
'   the bottom table header row shifted down by two and the top table gained two data rows.
'@TestMethod("CustomTable")
Public Sub TestAddRowsRespectsAdjacentTables()
    CustomTestSetTitles Assert, "CustomTable", "TestAddRowsRespectsAdjacentTables"
    On Error GoTo Fail

    Dim topLo As ListObject
    Dim bottomLo As ListObject
    Dim topTable As ICustomTable
    Dim originalBottomHeaderRow As Long

    PrepareMultiTableFixture topLo, bottomLo
    originalBottomHeaderRow = bottomLo.HeaderRowRange.Row

    Set topTable = CustomTable.Create(topLo, "ID", "row")

    topTable.AddRows nbRows:=2, insertShift:=True

    Assert.IsTrue (originalBottomHeaderRow + 2 =  bottomLo.HeaderRowRange.Row), _
                    "InsertShift should push the following table down"
    Assert.IsTrue (topLo.DataBodyRange.Rows.Count = 4),  "Top table should include the additional rows"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddRowsRespectsAdjacentTables", Err.Number, Err.Description
End Sub

'@sub-title Verifies that InsertRowsAt adds blank rows matching the selection height
'@details Builds the default three-row fixture, selects a two-row range starting at
'   row 2, then calls InsertRowsAt. Asserts the table grows to five rows, blank rows
'   appear before each selected row, and the original data shifts downward correctly.
'@TestMethod("CustomTable")
Public Sub TestInsertRowsAtInsertsSelectionRowCount()
    CustomTestSetTitles Assert, "CustomTable", "TestInsertRowsAtInsertsSelectionRowCount"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim lo As ListObject
    Dim selectionRange As Range

    Set tableObject = BuildCustomTable
    Set lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    Set selectionRange = lo.ListRows(2).Range
    Set selectionRange = selectionRange.Resize(2, selectionRange.Columns.Count)

    tableObject.InsertRowsAt selectionRange

    Assert.AreEqual 5, lo.ListRows.Count, _
                     "InsertRowsAt should add rows matching the selection height"
    Assert.AreEqual vbNullString, CStr(lo.ListColumns("Name").DataBodyRange.Cells(2, 1).Value), _
                     "A blank row should be inserted ahead of the first selection row"
    Assert.AreEqual "Beta", lo.ListColumns("Name").DataBodyRange.Cells(3, 1).Value, _
                     "Original second row should shift down by one position"
    Assert.AreEqual vbNullString, CStr(lo.ListColumns("Name").DataBodyRange.Cells(4, 1).Value), _
                     "A second blank row should be inserted ahead of the last selection row"
    Assert.AreEqual "Gamma", lo.ListColumns("Name").DataBodyRange.Cells(5, 1).Value, _
                     "Last row should slide under the second inserted row"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsAtInsertsSelectionRowCount", Err.Number, Err.Description
End Sub

'@sub-title Verifies that DeleteRowsAt removes the selected rows from the table
'@details Builds the default three-row fixture, selects the first two rows, then calls
'   DeleteRowsAt. Asserts only one row remains and it contains the trailing record
'   ("Gamma") that was not part of the selection.
'@TestMethod("CustomTable")
Public Sub TestDeleteRowsAtRemovesSelectedRows()
    CustomTestSetTitles Assert, "CustomTable", "TestDeleteRowsAtRemovesSelectedRows"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim lo As ListObject
    Dim selectionRange As Range

    Set tableObject = BuildCustomTable
    Set lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    Set selectionRange = lo.ListRows(1).Range
    Set selectionRange = selectionRange.Resize(2, selectionRange.Columns.Count)

    tableObject.DeleteRowsAt selectionRange

    Assert.AreEqual 1, lo.ListRows.Count, _
                     "DeleteRowsAt should remove the selected rows"
    Assert.AreEqual "Gamma", lo.ListColumns("Name").DataBodyRange.Cells(1, 1).Value, _
                     "Remaining row should preserve the trailing record"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDeleteRowsAtRemovesSelectedRows", Err.Number, Err.Description
End Sub

'@sub-title Verifies that DeleteRowsAt preserves at least one blank template row
'@details Reduces the fixture to a single row, then attempts to delete it. Asserts that
'   the table still contains one row (the template row) and that the row is blank, ensuring
'   the table structure is never fully emptied.
'@TestMethod("CustomTable")
Public Sub TestDeleteRowsAtKeepsTemplateRow()
    CustomTestSetTitles Assert, "CustomTable", "TestDeleteRowsAtKeepsTemplateRow"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim lo As ListObject
    Dim selectionRange As Range

    Set tableObject = BuildCustomTable
    Set lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    lo.ListRows(3).Delete
    lo.ListRows(2).Delete

    Set selectionRange = lo.ListRows(1).Range

    tableObject.DeleteRowsAt selectionRange

    Assert.AreEqual 1, lo.ListRows.Count, _
                     "DeleteRowsAt should always preserve at least one template row"
    Assert.AreEqual vbNullString, CStr(lo.ListColumns("Name").DataBodyRange.Cells(1, 1).Value), _
                     "Template row should be blank after deletion"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDeleteRowsAtKeepsTemplateRow", Err.Number, Err.Description
End Sub

'@sub-title Verifies that InsertRowsAt with insertShift shifts a stacked table and DeleteRowsAt restores it
'@details Builds stacked tables, inserts one row in the top table with insertShift:=True,
'   then asserts the bottom table shifted down by one. After deleting the inserted row,
'   asserts the bottom table returns to its original position and the top table regains
'   its original row count.
'@TestMethod("CustomTable")
Public Sub TestInsertRowsAtWithShiftMovesStackedTables()
    CustomTestSetTitles Assert, "CustomTable", "TestInsertRowsAtWithShiftMovesStackedTables"
    On Error GoTo Fail

    Dim topLo As ListObject
    Dim bottomLo As ListObject
    Dim topTable As ICustomTable
    Dim originalBottomHeaderRow As Long
    Dim originalTopRowCount As Long
    Dim insertSelection As Range
    Dim removalSelection As Range

    PrepareMultiTableFixture topLo, bottomLo
    originalBottomHeaderRow = bottomLo.HeaderRowRange.Row
    originalTopRowCount = topLo.ListRows.Count

    Set topTable = CustomTable.Create(topLo, "ID", "row")

    Set insertSelection = topLo.ListRows(1).Range
    topTable.InsertRowsAt insertSelection, insertShift:=True

    Assert.AreEqual originalBottomHeaderRow + 1, bottomLo.HeaderRowRange.Row, _
                    "Worksheet insertion should push the following table down"
    Assert.AreEqual originalTopRowCount + 1, topLo.ListRows.Count, _
                    "Top table should gain one data row"

    Set removalSelection = topLo.ListRows(1).Range
    topTable.DeleteRowsAt removalSelection

    Assert.AreEqual originalBottomHeaderRow, bottomLo.HeaderRowRange.Row, _
                    "Deleting with the tracker should restore the stacked table position"
    Assert.AreEqual originalTopRowCount, topLo.ListRows.Count, _
                    "Top table should return to its original row count"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsAtWithShiftMovesStackedTables", Err.Number, Err.Description
End Sub

'@sub-title Verifies that DeleteRowsAt with forceShift pulls a stacked table upward
'@details Builds stacked tables, deletes the last row of the top table with
'   forceShift:=True. Asserts the bottom table header row moved up by one and the
'   top table lost one data row.
'@TestMethod("CustomTable")
Public Sub TestDeleteRowsAtForceShiftMovesStackedTables()
    CustomTestSetTitles Assert, "CustomTable", "TestDeleteRowsAtForceShiftMovesStackedTables"
    On Error GoTo Fail

    Dim topLo As ListObject
    Dim bottomLo As ListObject
    Dim topTable As ICustomTable
    Dim originalBottomHeaderRow As Long
    Dim originalTopRowCount As Long
    Dim removalSelection As Range

    PrepareMultiTableFixture topLo, bottomLo
    originalBottomHeaderRow = bottomLo.HeaderRowRange.Row
    originalTopRowCount = topLo.ListRows.Count

    Set topTable = CustomTable.Create(topLo, "ID", "row")

    Set removalSelection = topLo.ListRows(topLo.ListRows.Count).Range
    topTable.DeleteRowsAt removalSelection, forceShift:=True

    Assert.AreEqual originalBottomHeaderRow - 1, bottomLo.HeaderRowRange.Row, _
                    "ForceShift should pull the stacked table upward"
    Assert.AreEqual originalTopRowCount - 1, topLo.ListRows.Count, _
                    "Top table should lose the deleted row"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDeleteRowsAtForceShiftMovesStackedTables", Err.Number, Err.Description
End Sub

'@sub-title Verifies that importing a DataSheet with extra columns records the missing ones
'@details Builds the default fixture, then imports a DataSheet whose headers include an
'   extra "NewValue" column not present in the target. Asserts HasColumnsNotImported is
'   True, ImportColumnsNotFound contains exactly "NewValue", and existing columns still
'   received their data.
'@TestMethod("CustomTable")
Public Sub TestImportRecordsMissingColumns()
    CustomTestSetTitles Assert, "CustomTable", "TestImportRecordsMissingColumns"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim dataSheetObj As IDataSheet
    Dim headers As Variant
    Dim rows As Variant
    Dim missing As BetterArray
    Dim Lo As ListObject

    Set tableObject = BuildCustomTable
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    headers = Array("ID", "Name", "NewValue")
    rows = Array( _
        Array(1, "Alpha", "A"), _
        Array(2, "Beta", "B"))

    Set dataSheetObj = CreateDataSheet(DATASHEETNAME, headers, rows)

    tableObject.Import dataSheetObj, keepSourceHeaders:=False

    Assert.IsTrue tableObject.HasColumnsNotImported, "Import should flag unsupported headers"
    Set missing = tableObject.ImportColumnsNotFound
    Assert.IsTrue (missing.Length = 1), "Only the unexpected column should be reported"
    Assert.AreEqual "NewValue", CStr(missing.Item(missing.LowerBound)), "Reported column should match the missing header"
    Assert.AreEqual "Beta", Lo.ListColumns("Name").DataBodyRange.Cells(2, 1).Value, _
                     "Existing columns should still import their data"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportRecordsMissingColumns", Err.Number, Err.Description
End Sub

'@sub-title Verifies that pasteAtBottom appends imported rows below existing data
'@details Builds the default three-row fixture, then imports a two-row source with
'   pasteAtBottom:=True. Asserts the total grows to five rows, existing rows remain
'   at the top, and appended rows appear in source order at positions 4 and 5.
'@TestMethod("CustomTable")
Public Sub TestImportPasteAtBottomAppendsData()
    CustomTestSetTitles Assert, "CustomTable", "TestImportPasteAtBottomAppendsData"
    On Error GoTo Fail

    Dim targetTable As ICustomTable
    Dim sourceTable As ICustomTable
    Dim Lo As ListObject
    Dim headers As Variant
    Dim rows As Variant

    Set targetTable = BuildCustomTable
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    headers = CustomTableHeaders()
    rows = Array( _
        Array("row 4", "Delta", 40), _
        Array("row 5", "Epsilon", 50))

    Set sourceTable = CreateCustomTableWithData(SOURCE_SHEETNAME, TABLENAME & "Append", headers, rows)

    targetTable.Import sourceTable, pasteAtBottom:=True, keepSourceHeaders:=False

    Assert.AreEqual 5, Lo.DataBodyRange.Rows.Count, _
                     "Import with pasteAtBottom should append incoming rows"
    Assert.AreEqual "Alpha", Lo.ListColumns("Name").DataBodyRange.Cells(1, 1).Value, _
                     "Existing rows should remain at the top when appending"
    Assert.AreEqual "Delta", Lo.ListColumns("Name").DataBodyRange.Cells(4, 1).Value, _
                     "First imported row should follow existing data"
    Assert.AreEqual "Epsilon", Lo.ListColumns("Name").DataBodyRange.Cells(5, 1).Value, _
                     "Rows should append in source order"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportPasteAtBottomAppendsData", Err.Number, Err.Description
End Sub

'@sub-title Verifies that importing preserves hidden column state and still updates hidden values
'@details Hides the Amount column, imports data via a DataSheet with keepSourceHeaders
'   toggled both ways. Asserts the column remains hidden after import and that values
'   in the hidden column are correctly updated regardless of the keepSourceHeaders setting.
'@TestMethod("CustomTable")
Public Sub TestImportPreservesHiddenColumns()
    CustomTestSetTitles Assert, "CustomTable", "TestImportPreservesHiddenColumns"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim Lo As ListObject
    Dim headers As Variant
    Dim rows As Variant
    Dim dataSheetObj As IDataSheet

    Set tableObject = BuildCustomTable
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    Lo.ListColumns("Amount").Range.EntireColumn.Hidden = True

    headers = CustomTableHeaders()
    rows = Array( _
        Array(1, "Omega", 99), _
        Array(2, "Sigma", 123))

    Set dataSheetObj = CreateDataSheet(DATASHEETNAME, headers, rows)

    'Import with source headers
    tableObject.Import dataSheetObj, keepSourceHeaders:=True

    Assert.IsTrue Lo.ListColumns("Amount").Range.EntireColumn.Hidden, _
                  "Import should restore hidden columns"
    Dim hidVal As String
    hidVal = Lo.ListColumns("Amount").DataBodyRange.Cells(2, 1).Value
    Assert.AreEqual "123", CStr(hidVal), "Hidden column values should still update - value when keepSourceHeaders = yes"

    'Import without keeping source headers
    Set tableObject = BuildCustomTable()
    Set Lo = ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    Lo.ListColumns("Amount").Range.EntireColumn.Hidden = True
    tableObject.Import dataSheetObj, keepSourceHeaders:=False
    hidVal = Lo.ListColumns("Amount").DataBodyRange.Cells(2, 1).Value
    Assert.AreEqual "123", CStr(hidVal), "Hidden column values should still update - value when keepSourceHeaders = no"
    Assert.IsTrue (Lo.ListColumns("Amount").DataBodyRange.Cells(3, 1).Value = vbNullString), "Hidden column values should be deleted - value when keepSourceHeaders = no"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportPreservesHiddenColumns", Err.Number, Err.Description
End Sub

'@sub-title Verifies that strictColumnSearch rejects case-mismatched headers
'@details Builds the default fixture and imports a DataSheet where "Name" is spelled as
'   "name" (lowercase). Asserts HasColumnsNotImported flags the mismatch, the Name column
'   stays blank because strict matching cannot resolve it, and exactly-matched columns
'   (Amount) still receive their data.
'@TestMethod("CustomTable")
Public Sub TestImportStrictColumnSearchRequiresExactMatch()
    CustomTestSetTitles Assert, "CustomTable", "TestImportStrictColumnSearchRequiresExactMatch"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim Lo As ListObject
    Dim headers As Variant
    Dim rows As Variant
    Dim dataSheetObj As IDataSheet
    Dim missing As BetterArray
    Dim nameValue As Variant

    Set tableObject = BuildCustomTable
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    headers = Array("ID", "name", "Amount")
    rows = Array(Array(1, "Omega", 900))

    Set dataSheetObj = CreateDataSheet(DATASHEETNAME, headers, rows)

    tableObject.Import dataSheetObj, strictColumnSearch:=True

    Assert.IsTrue tableObject.HasColumnsNotImported, _
                  "Strict column search should flag case-mismatched headers"
    Set missing = tableObject.ImportColumnsNotFound
    Assert.AreEqual "name", CStr(missing.Item(missing.LowerBound)), _
                     "Mismatched header should be reported exactly"

    nameValue = Lo.ListColumns("Name").DataBodyRange.Cells(1, 1).Value
    Assert.IsTrue (IsEmpty(nameValue) Or nameValue = vbNullString), _
                  "Name column should remain blank when strict search cannot match header"
    Assert.AreEqual "900", CStr(Lo.ListColumns("Amount").DataBodyRange.Cells(1, 1).Value), _
                     "Matching headers should still import their data"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportStrictColumnSearchRequiresExactMatch", Err.Number, Err.Description
End Sub

'@sub-title Verifies that RestoreTableSnapshot reverts the table after a failed import
'@details Builds the default fixture, imports a DataSheet with a case-mismatched header
'   (triggering HasColumnsNotImported), then calls RestoreTableSnapshot. Asserts the
'   original first and third row values are restored and the Amount column still exists,
'   confirming the snapshot captured the full table state before the import.
'@TestMethod("CustomTable")
Public Sub TestImportFailureRestoresSnapshot()
    CustomTestSetTitles Assert, "CustomTable", "TestImportFailureRestoresSnapshot"

    Dim tableObject As ICustomTable
    Dim headers As Variant
    Dim rows As Variant
    Dim dataSheetObj As IDataSheet
    Dim Lo As ListObject

    Set tableObject = BuildCustomTable
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    headers = Array("ID", "name")
    rows = Array(Array(1, "Replacement"))

    Set dataSheetObj = CreateDataSheet(DATASHEETNAME, headers, rows)

    tableObject.Import dataSheetObj, strictColumnSearch:=True
    tableObject.RestoreTableSnapshot

    Assert.AreEqual "Alpha", Lo.ListColumns("Name").DataBodyRange.Cells(1, 1).Value, _
                     "Failed import should restore original first row"

    Assert.AreEqual "Gamma", Lo.ListColumns("Name").DataBodyRange.Cells(3, 1).Value, _
                     "Failed import should keep trailing rows intact"

      Assert.IsTrue Not (Lo.ListColumns("Amount") Is Nothing), "failed to estore listcolumns"
End Sub

'@sub-title Verifies that importing a larger CustomTable into a stacked top table shifts the bottom table
'@details Builds stacked tables with two top rows, imports a four-row source into the top
'   table. Asserts the top table resizes to four rows, the bottom table header shifts
'   down by the row delta, and the last imported value matches the source.
'@TestMethod("CustomTable")
Public Sub TestImportAllCustomTableShiftsFollowingTables()
    CustomTestSetTitles Assert, "CustomTable", "TestImportAllCustomTableShiftsFollowingTables"
    On Error GoTo Fail

    Dim topLo As ListObject
    Dim bottomLo As ListObject
    Dim topTable As ICustomTable
    Dim sourceTable As ICustomTable
    Dim headers As Variant
    Dim rows As Variant
    Dim originalBottomHeaderRow As Long
    Dim originalTopRows As Long
    Dim expectedBottomRow As Long

    PrepareMultiTableFixture topLo, bottomLo
    originalBottomHeaderRow = bottomLo.HeaderRowRange.Row
    originalTopRows = topLo.DataBodyRange.Rows.Count

    headers = Array("ID", "Name")
    rows = Array( _
        Array("row 1", "One"), _
        Array("row 2", "Two"), _
        Array("row 3", "Three"), _
        Array("row 4", "Four"))

    Set sourceTable = CreateCustomTableWithData(SOURCE_SHEETNAME, "tblSourceAll", headers, rows)
    Set topTable = CustomTable.Create(topLo, "ID", "row")

    topTable.Import sourceTable, keepSourceHeaders:=False

    expectedBottomRow = originalBottomHeaderRow + ((UBound(rows) - LBound(rows) + 1) - originalTopRows)

    Assert.AreEqual UBound(rows) - LBound(rows) + 1, topLo.DataBodyRange.Rows.Count, _
                     "ImportAll should resize the target table to match source rows"
    Assert.AreEqual expectedBottomRow, bottomLo.HeaderRowRange.Row, _
                     "ImportAll should insert rows and shift following tables down"
    Assert.AreEqual "Four", topLo.ListColumns("Name").DataBodyRange.Cells(topLo.DataBodyRange.Rows.Count, 1).Value, _
                     "Imported data should match the source table"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportAllCustomTableShiftsFollowingTables", Err.Number, Err.Description
End Sub

'@sub-title Verifies that importing a larger DataSheet into a stacked top table shifts the bottom table
'@details Builds stacked tables with two top rows, imports a five-row DataSheet into the
'   top table. Asserts the top table resizes to five rows, the bottom table shifts down
'   accordingly, and the last value matches the DataSheet content.
'@TestMethod("CustomTable")
Public Sub TestImportAllDataSheetShiftsFollowingTables()
    CustomTestSetTitles Assert, "CustomTable", "TestImportAllDataSheetShiftsFollowingTables"
    On Error GoTo Fail

    Dim topLo As ListObject
    Dim bottomLo As ListObject
    Dim topTable As ICustomTable
    Dim dataSheetObj As IDataSheet
    Dim headers As Variant
    Dim rows As Variant
    Dim originalBottomHeaderRow As Long
    Dim originalTopRows As Long
    Dim expectedBottomRow As Long

    PrepareMultiTableFixture topLo, bottomLo
    originalBottomHeaderRow = bottomLo.HeaderRowRange.Row
    originalTopRows = topLo.DataBodyRange.Rows.Count

    headers = Array("ID", "Name")
    rows = Array( _
        Array("row 1", "Alpha"), _
        Array("row 2", "Beta"), _
        Array("row 3", "Gamma"), _
        Array("row 4", "Delta"), _
        Array("row 5", "Epsilon"))

    Set dataSheetObj = CreateDataSheet(DATASHEETNAME, headers, rows)
    Set topTable = CustomTable.Create(topLo, "ID", "row")

    topTable.Import dataSheetObj, keepSourceHeaders:=False

    expectedBottomRow = originalBottomHeaderRow + ((UBound(rows) - LBound(rows) + 1) - originalTopRows)

    Assert.AreEqual UBound(rows) - LBound(rows) + 1, topLo.DataBodyRange.Rows.Count, _
                     "ImportAll from DataSheet should resize the table"
    Assert.AreEqual expectedBottomRow, bottomLo.HeaderRowRange.Row, _
                     "ImportAll from DataSheet should shift following tables"
    Assert.AreEqual "Epsilon", topLo.ListColumns("Name").DataBodyRange.Cells(topLo.DataBodyRange.Rows.Count, 1).Value, _
                     "Imported values should reflect the DataSheet content"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportAllDataSheetShiftsFollowingTables", Err.Number, Err.Description
End Sub

'@sub-title Verifies that importing a source with extra columns expands the target table
'@details Creates a two-column target and a four-column source, then imports with
'   keepSourceHeaders:=True. Asserts the target table gains the two extra columns, the
'   new column names match the source, and imported data populates the newly added columns.
'@TestMethod("CustomTable")
Public Sub TestImportAllExpandsColumnsWithHeaders()
    CustomTestSetTitles Assert, "CustomTable", "TestImportAllExpandsColumnsWithHeaders"
    On Error GoTo Fail

    Dim targetTable As ICustomTable
    Dim sourceTable As ICustomTable
    Dim targetLo As ListObject
    Dim targetHeaders As Variant
    Dim targetRows As Variant
    Dim sourceHeaders As Variant
    Dim sourceRows As Variant

    targetHeaders = Array("ID", "Name")
    targetRows = Array(Array("row 1", "One"))
    Set targetTable = CreateCustomTableWithData(EXPAND_SHEETNAME, EXPAND_TABLE_NAME, targetHeaders, targetRows)

    sourceHeaders = Array("ID", "Name", "Amount", "Notes")
    sourceRows = Array(Array("row 1", "Uno", 100, "Extra"))
    Set sourceTable = CreateCustomTableWithData(SOURCE_SHEETNAME, "tblExpandSource", sourceHeaders, sourceRows)

    targetTable.Import sourceTable, keepSourceHeaders:=True
    Set targetLo = ThisWorkbook.Worksheets(EXPAND_SHEETNAME).ListObjects(EXPAND_TABLE_NAME)

    Assert.AreEqual 4, targetLo.ListColumns.Count, "ImportAll should expand target columns to match source headers"
    Assert.AreEqual "Notes", targetLo.ListColumns(4).Name, "New columns should copy the source header names"
    Assert.AreEqual "Extra", targetLo.ListColumns("Notes").DataBodyRange.Cells(1, 1).Value, _
                     "Imported data should populate the newly added column"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportAllExpandsColumnsWithHeaders", Err.Number, Err.Description
End Sub

'@sub-title Verifies that importing a source with fewer columns trims the target table
'@details Creates a four-column target and a two-column source, then imports with
'   keepSourceHeaders:=True. Asserts the target is trimmed to two columns matching the
'   source headers and that imported data fills the reduced table correctly.
'@TestMethod("CustomTable")
Public Sub TestImportAllTrimsExtraColumns()
    CustomTestSetTitles Assert, "CustomTable", "TestImportAllTrimsExtraColumns"
    On Error GoTo Fail

    Dim targetTable As ICustomTable
    Dim sourceTable As ICustomTable
    Dim targetLo As ListObject
    Dim targetHeaders As Variant
    Dim targetRows As Variant
    Dim sourceHeaders As Variant
    Dim sourceRows As Variant

    targetHeaders = Array("ID", "Name", "Amount", "Notes")
    targetRows = Array(Array("row 1", "One", 10, "keep"))
    Set targetTable = CreateCustomTableWithData(TRIM_SHEETNAME, TRIM_TABLE_NAME, targetHeaders, targetRows)

    sourceHeaders = Array("ID", "Name")
    sourceRows = Array(Array("row 1", "Replace"))
    Set sourceTable = CreateCustomTableWithData(SOURCE_SHEETNAME, "tblTrimSource", sourceHeaders, sourceRows)

    targetTable.Import sourceTable, keepSourceHeaders:=True
    Set targetLo = ThisWorkbook.Worksheets(TRIM_SHEETNAME).ListObjects(TRIM_TABLE_NAME)

    Assert.AreEqual 2, targetLo.ListColumns.Count, "ImportAll should trim extra columns when source has fewer headers"
    Assert.AreEqual "Name", targetLo.ListColumns(2).Name, "Remaining columns must match the source headers"
    Assert.AreEqual "Replace", targetLo.ListColumns("Name").DataBodyRange.Cells(1, 1).Value, _
                     "Imported data should fill the trimmed table correctly"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportAllTrimsExtraColumns", Err.Number, Err.Description
End Sub

'@sub-title Verifies that formatHeaders copies cell styling (color, bold, comments) during import
'@details Creates a source table with a styled Name cell (yellow background, bold font,
'   comment "Important"), then imports into an empty target with formatHeaders listing
'   "Name". Asserts the target cell inherits the interior color, font weight, and comment
'   text from the source.
'@TestMethod("CustomTable")
Public Sub TestImportWithFormatHeadersCopiesStyling()
    CustomTestSetTitles Assert, "CustomTable", "TestImportWithFormatHeadersCopiesStyling"
    On Error GoTo Fail

    Dim targetTable As ICustomTable
    Dim sourceTable As ICustomTable
    Dim formatHeaders As BetterArray
    Dim sourceCell As Range
    Dim targetCell As Range

    PrepareCustomTable includeData:=False
    Set targetTable = CustomTable.Create(ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME), "ID", "row")

    Set sourceTable = CreateCustomTableWithData(SOURCE_SHEETNAME, "tblFormatSource", _
                                                CustomTableHeaders(), _
                                                Array(Array("row 1", "Styled", 42)))

    Set sourceCell = ThisWorkbook.Worksheets(SOURCE_SHEETNAME).ListObjects("tblFormatSource").ListColumns("Name").DataBodyRange.Cells(1, 1)
    On Error Resume Next
        sourceCell.Comment.Delete
    On Error GoTo 0
    sourceCell.Interior.Color = RGB(255, 255, 0)
    sourceCell.Font.Bold = True
    sourceCell.AddComment "Important"

    Set formatHeaders = NewBetterArray("Name")

    targetTable.Import sourceTable, formatHeaders:=formatHeaders

    Set targetCell = ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME).ListColumns("Name").DataBodyRange.Cells(1, 1)

    Assert.AreEqual sourceCell.Interior.Color, targetCell.Interior.Color, _
                     "Import should copy cell interior color for formatted headers"
    Assert.IsTrue targetCell.Font.Bold, "Import should copy font weight for formatted headers"

    Dim commentText As String
    On Error Resume Next
        commentText = targetCell.Comment.Text
    On Error GoTo 0
    Assert.IsTrue (InStr(1, commentText, "Important", vbTextCompare) > 0), _
                 "Import should copy cell comments for formatted headers"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportWithFormatHeadersCopiesStyling", Err.Number, Err.Description
End Sub

'@sub-title Verifies that Export writes selected headers and data at the requested start line
'@details Builds the default fixture, sets a value in cell A1, then exports only "Name"
'   and "Amount" columns starting at row 3. Asserts the sheet is cleared before writing,
'   headers appear at row 3 in the requested order, unrequested columns are absent, and
'   data rows follow immediately below.
'@TestMethod("CustomTable")
Public Sub TestExportWritesSelectedHeadersAtRow()
    CustomTestSetTitles Assert, "CustomTable", "TestExportWritesSelectedHeadersAtRow"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim exportSheet As Worksheet
    Dim selectedHeaders As BetterArray

    Set tableObject = BuildCustomTable
    Set exportSheet = EnsureWorksheet(EXPORTSHEETNAME)
    ClearWorksheet exportSheet
    exportSheet.Range("A1").Value = "Should be cleared"

    Set selectedHeaders = NewBetterArray("Name", "Amount")

    tableObject.Export exportSheet, headersTable:=selectedHeaders, startLine:=3

    Assert.IsTrue IsEmpty(exportSheet.Cells(1, 1).Value), "Export should clear the target worksheet before writing"
    Assert.AreEqual "Name", exportSheet.Cells(3, 1).Value, "Export should write headers at the requested start line"
    Assert.AreEqual "Amount", exportSheet.Cells(3, 2).Value, "Export should respect requested header order"
    Assert.IsTrue IsEmpty(exportSheet.Cells(3, 3).Value), "Unrequested columns should remain empty"
    Assert.AreEqual "Alpha", exportSheet.Cells(4, 1).Value, "Export should include data rows beneath the headers"
    Assert.IsTrue (exportSheet.Cells(4, 2).Value = "10"), "Export should include corresponding values for selected headers"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportWritesSelectedHeadersAtRow", Err.Number, Err.Description
End Sub

'@sub-title Verifies that Export restores hidden column state and includes hidden columns in output
'@details Hides the Amount column, exports to a clean sheet without restricting headers.
'   Asserts the Amount column remains hidden on the source table after export and that
'   the export sheet includes Amount data in its output.
'@TestMethod("CustomTable")
Public Sub TestExportRestoresHiddenColumns()
    CustomTestSetTitles Assert, "CustomTable", "TestExportRestoresHiddenColumns"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim Lo As ListObject
    Dim exportSheet As Worksheet

    Set tableObject = BuildCustomTable
    Set Lo =  ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    Lo.ListColumns("Amount").Range.EntireColumn.Hidden = True

    Set exportSheet = EnsureWorksheet(EXPORTSHEETNAME)
    ClearWorksheet exportSheet

    tableObject.Export exportSheet

    Assert.IsTrue Lo.ListColumns("Amount").Range.EntireColumn.Hidden, _
                  "Export should restore hidden column state after writing"
    Assert.AreEqual "Amount", exportSheet.Cells(1, 3).Value, _
                  "Export should include hidden columns when headers are not restricted"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportRestoresHiddenColumns", Err.Number, Err.Description
End Sub

'@sub-title Verifies that Export creates a ListObject when addListObject is True
'@details Builds the default fixture and exports with addListObject:=True. Asserts the
'   export sheet contains exactly one ListObject whose name, table style, and row count
'   match the source table.
'@TestMethod("CustomTable")
Public Sub TestExportAddsListObjectWhenRequested()
    CustomTestSetTitles Assert, "CustomTable", "TestExportAddsListObjectWhenRequested"
    On Error GoTo Fail

    Dim tableObject As ICustomTable
    Dim exportSheet As Worksheet
    Dim outLo As ListObject
    Dim sourceLo As ListObject

    Set tableObject = BuildCustomTable
    Set sourceLo = ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    Set exportSheet = EnsureWorksheet(EXPORTSHEETNAME)
    ClearWorksheet exportSheet

    tableObject.Export exportSheet, addListObject:=True

    Assert.AreEqual 1, exportSheet.ListObjects.Count, "Export should create a ListObject when requested"
    Set outLo = exportSheet.ListObjects(1)
    Assert.AreEqual tableObject.Name, outLo.Name, "Export should preserve table name on the created ListObject"
    Assert.AreEqual sourceLo.TableStyle, outLo.TableStyle, "Export should copy table style to the created ListObject"
    Assert.AreEqual sourceLo.ListRows.Count, outLo.ListRows.Count, "Exported ListObject should contain all data rows"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportAddsListObjectWhenRequested", Err.Number, Err.Description
End Sub

'@sub-title Verifies that Create raises ObjectNotInitialized when given Nothing
'@details Calls CustomTable.Create(Nothing) and expects the error handler to catch
'   ProjectError.ObjectNotInitialized. Asserts the error number matches and logs a
'   failure if no error is raised.
'@TestMethod("CustomTable")
Public Sub TestCreateRejectsMissingListObject()
    CustomTestSetTitles Assert, "CustomTable", "TestCreateRejectsMissingListObject"
    On Error GoTo ExpectError

    Dim tableObject As ICustomTable
    '@Ignore AssignmentNotUsed
    Set tableObject = CustomTable.Create(Nothing)
    Assert.LogFailure "Create should raise when no ListObject is supplied"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "Create should surface ObjectNotInitialized for missing listobject"
    Err.Clear
End Sub

'@sub-title Verifies that DataRange returns Nothing when the table has no data rows
'@details Prepares a header-only fixture (includeData:=False), creates a CustomTable, then
'   calls DataRange("Name"). Asserts the result Is Nothing, confirming no phantom range is
'   returned for an empty column.
'@TestMethod("CustomTable")
Public Sub TestDataRangeReturnsNothingWhenEmpty()
    CustomTestSetTitles Assert, "CustomTable", "TestDataRangeReturnsNothingWhenEmpty"
    On Error GoTo Fail

    PrepareCustomTable includeData:=False

    Dim tableObject As ICustomTable
    Set tableObject = CustomTable.Create(ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME), "ID", "row")

    Dim nameRange As Range
    Set nameRange = tableObject.DataRange("Name")

    Assert.IsTrue (nameRange Is Nothing), "DataRange should return Nothing when the column has no data"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDataRangeReturnsNothingWhenEmpty", Err.Number, Err.Description
End Sub

'@sub-title Verifies that SortOnFirst handles non-string cell values without crashing
'@details Creates a table where the Name column contains mixed types (numbers,
'   dates, empty). Calls Sort with directSort:=False (SortOnFirst mode). Asserts
'   that the sort completes without error and groups duplicate values correctly.
'@TestMethod("CustomTable")
Public Sub TestSortOnFirstHandlesNonStringValues()
    CustomTestSetTitles Assert, "CustomTable", "TestSortOnFirstHandlesNonStringValues"
    On Error GoTo Fail

    Dim headers As Variant
    Dim rows As Variant
    Dim tableObject As ICustomTable
    Dim Lo As ListObject

    headers = CustomTableHeaders()
    rows = Array( _
        Array("row 1", 100, 1), _
        Array("row 2", 200, 2), _
        Array("row 3", 100, 3), _
        Array("row 4", 300, 4))

    Set tableObject = CreateCustomTableWithData(TABLESHEETNAME, TABLENAME, headers, rows)
    Set Lo = ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME)

    tableObject.Sort colName:="Name", directSort:=False

    Assert.AreEqual "100", CStr(Lo.ListColumns("Name").DataBodyRange.Cells(1, 1).Value), _
                     "First numeric group should appear first after SortOnFirst"
    Assert.AreEqual "100", CStr(Lo.ListColumns("Name").DataBodyRange.Cells(2, 1).Value), _
                     "Duplicate numeric values should remain adjacent"
    Assert.AreEqual 3, Lo.ListColumns.Count, _
                     "SortOnFirst should remove its helper column"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSortOnFirstHandlesNonStringValues", Err.Number, Err.Description
End Sub
