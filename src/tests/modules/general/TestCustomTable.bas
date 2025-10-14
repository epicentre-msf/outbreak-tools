Attribute VB_Name = "TestCustomTable"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")

Private Const TABLESHEETNAME As String = "CustomTableFixture"
Private Const TABLENAME As String = "tblCustom"
Private Const SOURCE_SHEETNAME As String = "CustomTableFixtureSource"
Private Const DATASHEETNAME As String = "CustomTableData"
Private Const MULTITABLESHEET As String = "CustomTableMulti"
Private Const EXPORTSHEETNAME As String = "CustomTableExport"

Private Assert As ICustomTest
Private Fakes As Object

'@section Helpers
'===============================================================================

Private Function CustomTableHeaders() As Variant
    CustomTableHeaders = Array("ID", "Name", "Amount")
End Function

Private Function CustomTableRows() As Variant
    CustomTableRows = Array( _
        Array(1, "Alpha", 10), _
        Array(2, "Beta", 20), _
        Array(3, "Gamma", 30))
End Function

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

Private Function BuildCustomTable() As ICustomTable
    PrepareCustomTable
    Set BuildCustomTable = CustomTable.Create(ThisWorkbook.Worksheets(TABLESHEETNAME).ListObjects(TABLENAME), "ID", "row")
End Function

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

    Assert.IsTrue (Lo.DataBodyRange.rows.count = 2), "Import from DataSheet should size table to source rows"
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
