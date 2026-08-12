Attribute VB_Name = "TestDesignerMulti"
Attribute VB_Description = "Unit tests for Multi group table operations"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates Multi group table operations: add rows, remove rows, duplicate, import, export, the write-once row IDs on T_Multi, and the driver helpers that read a row into the Main entries.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private MultiSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TABLE_MULTI As String = "T_Multi"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDesignerMulti"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    BusyApp

    Set FixtureWorkbook = NewWorkbook
    Set MultiSheet = EnsureWorksheet("GenerateMultiple", FixtureWorkbook)
    CreateMultiTable MultiSheet
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set MultiSheet = Nothing
    Set FixtureWorkbook = Nothing

    RestoreApp
End Sub


'@section AddRows Tests
'===============================================================================
'@TestMethod("DesignerMulti.AddRows")
Public Sub TestAddRowsIncreasesRowCount()
    CustomTestSetTitles Assert, "DesignerMulti", "TestAddRowsIncreasesRowCount"
    On Error GoTo Fail

    'Arrange
    Dim lo As ListObject
    Set lo = MultiSheet.ListObjects(TABLE_MULTI)
    Dim initialRowCount As Long
    initialRowCount = lo.ListRows.Count

    Dim table As CustomTable
    Set table = CustomTable.Create(lo)

    'Act
    table.AddRows nbRows:=5, insertShift:=False, includeIds:=False

    'Assert
    Assert.AreEqual initialRowCount + 5, lo.ListRows.Count, _
                    "AddRows should increase row count by 5."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestAddRowsIncreasesRowCount", Err.Number, Err.Description
End Sub


'@section RemoveRows Tests
'===============================================================================
'@TestMethod("DesignerMulti.RemoveRows")
Public Sub TestRemoveRowsClearsEmptyRows()
    CustomTestSetTitles Assert, "DesignerMulti", "TestRemoveRowsClearsEmptyRows"
    On Error GoTo Fail

    'Arrange: add 5 empty rows, record the row count with data
    Dim lo As ListObject
    Set lo = MultiSheet.ListObjects(TABLE_MULTI)

    'Fill the first row so it is not empty
    lo.ListRows(1).Range.Cells(1, 2).Value = "test_value"
    Dim filledRowCount As Long
    filledRowCount = 1

    Dim table As CustomTable
    Set table = CustomTable.Create(lo)
    table.AddRows nbRows:=5, insertShift:=False, includeIds:=False

    'Act
    table.RemoveRows totalCount:=0, includeIds:=False, forceShift:=False

    'Assert: only the filled row should remain
    Assert.AreEqual filledRowCount, lo.ListRows.Count, _
                    "RemoveRows should leave only non-empty rows."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestRemoveRowsClearsEmptyRows", Err.Number, Err.Description
End Sub


'@section DuplicateRow Tests
'===============================================================================
'@TestMethod("DesignerMulti.DuplicateRow")
Public Sub TestDuplicateRowCopiesValues()
    CustomTestSetTitles Assert, "DesignerMulti", "TestDuplicateRowCopiesValues"
    On Error GoTo Fail

    'Arrange: fill row 1 with known values
    Dim lo As ListObject
    Set lo = MultiSheet.ListObjects(TABLE_MULTI)

    lo.ListRows(1).Range.Cells(1, 2).Value = "setup_path.xlsb"
    lo.ListRows(1).Range.Cells(1, 3).Value = "geo_path.xlsx"
    lo.ListRows(1).Range.Cells(1, 4).Value = "C:\output"

    Dim originalCount As Long
    originalCount = lo.ListRows.Count

    'Act: insert a duplicate after row 1
    lo.ListRows.Add Position:=2
    lo.ListRows(2).Range.Value = lo.ListRows(1).Range.Value

    'Assert
    Assert.AreEqual originalCount + 1, lo.ListRows.Count, _
                    "Duplicate should add one row."
    Assert.AreEqual "setup_path.xlsb", CStr(lo.ListRows(2).Range.Cells(1, 2).Value), _
                    "Duplicated row should have same setups value."
    Assert.AreEqual "geo_path.xlsx", CStr(lo.ListRows(2).Range.Cells(1, 3).Value), _
                    "Duplicated row should have same geobases value."
    Assert.AreEqual "C:\output", CStr(lo.ListRows(2).Range.Cells(1, 4).Value), _
                    "Duplicated row should have same output folders value."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestDuplicateRowCopiesValues", Err.Number, Err.Description
End Sub


'@section Import Tests
'===============================================================================
'@TestMethod("DesignerMulti.Import")
Public Sub TestImportReplacesTableData()
    CustomTestSetTitles Assert, "DesignerMulti", "TestImportReplacesTableData"
    On Error GoTo Fail

    'Arrange: create a source T_Multi on a separate worksheet
    Dim sourceSheet As Worksheet
    Set sourceSheet = EnsureWorksheet("SourceMulti", FixtureWorkbook)
    CreateMultiTable sourceSheet
    sourceSheet.ListObjects(TABLE_MULTI).Name = "T_Multi_Source"

    Dim sourceLo As ListObject
    Set sourceLo = sourceSheet.ListObjects("T_Multi_Source")
    sourceLo.ListRows(1).Range.Cells(1, 2).Value = "imported_setup.xlsb"
    sourceLo.ListRows(1).Range.Cells(1, 3).Value = "imported_geo.xlsx"

    Dim sourceTable As CustomTable
    Set sourceTable = CustomTable.Create(sourceLo)

    'Target table
    Dim targetLo As ListObject
    Set targetLo = MultiSheet.ListObjects(TABLE_MULTI)
    targetLo.ListRows(1).Range.Cells(1, 2).Value = "old_setup.xlsb"

    Dim targetTable As CustomTable
    Set targetTable = CustomTable.Create(targetLo)

    'Act
    targetTable.Import sourceTable

    'Assert
    Assert.AreEqual "imported_setup.xlsb", _
                    CStr(targetLo.ListRows(1).Range.Cells(1, 2).Value), _
                    "Import should replace setups value with source data."
    Assert.AreEqual "imported_geo.xlsx", _
                    CStr(targetLo.ListRows(1).Range.Cells(1, 3).Value), _
                    "Import should replace geobases value with source data."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestImportReplacesTableData", Err.Number, Err.Description
End Sub


'@section Export Tests
'===============================================================================
'@TestMethod("DesignerMulti.Export")
Public Sub TestExportWritesToWorksheet()
    CustomTestSetTitles Assert, "DesignerMulti", "TestExportWritesToWorksheet"
    On Error GoTo Fail

    'Arrange: fill T_Multi with data
    Dim lo As ListObject
    Set lo = MultiSheet.ListObjects(TABLE_MULTI)
    lo.ListRows(1).Range.Cells(1, 2).Value = "export_setup.xlsb"
    lo.ListRows(1).Range.Cells(1, 3).Value = "export_geo.xlsx"

    Dim table As CustomTable
    Set table = CustomTable.Create(lo)

    Dim exportSheet As Worksheet
    Set exportSheet = EnsureWorksheet("ExportTarget", FixtureWorkbook)

    'Act
    table.Export sh:=exportSheet, startLine:=1, startColumn:=1, addListObject:=True

    'Assert: the export sheet should have a ListObject with the same headers
    Assert.IsTrue exportSheet.ListObjects.Count > 0, _
                  "Export should create a ListObject on the target sheet."

    Dim exportLo As ListObject
    Set exportLo = exportSheet.ListObjects(1)
    Assert.AreEqual lo.ListColumns.Count, exportLo.ListColumns.Count, _
                    "Exported table should have the same number of columns."
    Assert.AreEqual "ID", exportLo.ListColumns(1).Name, _
                    "First column header should be 'ID'."
    Assert.AreEqual "setups", exportLo.ListColumns(2).Name, _
                    "Second column header should be 'setups'."
    Assert.AreEqual "export_setup.xlsb", _
                    CStr(exportLo.ListRows(1).Range.Cells(1, 2).Value), _
                    "Exported data should match source data."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestExportWritesToWorksheet", Err.Number, Err.Description
End Sub


'@section EnsureRowIds Tests
'===============================================================================
'@TestMethod("DesignerMulti.EnsureRowIds")
Public Sub TestEnsureRowIdsFillsOnlyBlankIds()
    CustomTestSetTitles Assert, "DesignerMulti", "TestEnsureRowIdsFillsOnlyBlankIds"
    On Error GoTo Fail

    'Arrange: three rows, two with IDs written and a blank one between them
    Dim lo As ListObject
    Set lo = MultiSheet.ListObjects(TABLE_MULTI)
    lo.ListRows.Add
    lo.ListRows.Add

    lo.ListRows(1).Range.Cells(1, 1).Value = "Operation- 1"
    lo.ListRows(2).Range.Cells(1, 1).Value = vbNullString
    lo.ListRows(3).Range.Cells(1, 1).Value = "Operation- 7"

    'Act
    Dim hasIdColumn As Boolean
    hasIdColumn = EventsDesignerMulti.EnsureRowIds(lo)

    'Assert: an ID is written once, so the two written IDs stay and the
    'blank cell gets the number after the largest one
    Assert.IsTrue hasIdColumn, "EnsureRowIds should find the ID column."
    Assert.AreEqual "Operation- 1", CStr(lo.ListRows(1).Range.Cells(1, 1).Value), _
                    "A written ID should keep its value."
    Assert.AreEqual "Operation- 8", CStr(lo.ListRows(2).Range.Cells(1, 1).Value), _
                    "A blank ID should get the next free number."
    Assert.AreEqual "Operation- 7", CStr(lo.ListRows(3).Range.Cells(1, 1).Value), _
                    "A written ID should keep its value after the fill."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestEnsureRowIdsFillsOnlyBlankIds", Err.Number, Err.Description
End Sub


'@section Multi driver Tests
'===============================================================================
'@TestMethod("DesignerMulti.Driver")
Public Sub TestOutputNameFromCellKeepsABareName()
    CustomTestSetTitles Assert, "DesignerMulti", "TestOutputNameFromCellKeepsABareName"
    On Error GoTo Fail

    Assert.AreEqual "my_linelist", _
                    EventsDesignerMulti.OutputNameFromCell("my_linelist"), _
                    "A bare name should come back unchanged."
    Assert.AreEqual "my_linelist", _
                    EventsDesignerMulti.OutputNameFromCell("  my_linelist  "), _
                    "The name should come back trimmed."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestOutputNameFromCellKeepsABareName", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.Driver")
Public Sub TestOutputNameFromCellStripsPathAndExtension()
    CustomTestSetTitles Assert, "DesignerMulti", "TestOutputNameFromCellStripsPathAndExtension"
    On Error GoTo Fail

    'A row that built holds the full written path, and a re-run reads it
    'back. The folder and the extension go so the name stays stable.
    Assert.AreEqual "my_linelist", _
                    EventsDesignerMulti.OutputNameFromCell("/out/folder/my_linelist.xlsb"), _
                    "A slash path with the .xlsb extension should reduce to the name."
    Assert.AreEqual "my_linelist", _
                    EventsDesignerMulti.OutputNameFromCell("C:\out\my_linelist.xlsb"), _
                    "A backslash path should reduce to the name too."
    Assert.AreEqual "my_linelist", _
                    EventsDesignerMulti.OutputNameFromCell("my_linelist.xlsb"), _
                    "A bare name with the extension should lose the extension."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestOutputNameFromCellStripsPathAndExtension", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.Driver")
Public Sub TestCountBuildRowsCountsFilledSetupCells()
    CustomTestSetTitles Assert, "DesignerMulti", "TestCountBuildRowsCountsFilledSetupCells"
    On Error GoTo Fail

    'Arrange: three rows, two with a setup path
    Dim lo As ListObject
    Set lo = MultiSheet.ListObjects(TABLE_MULTI)
    lo.ListRows.Add
    lo.ListRows.Add

    lo.ListRows(1).Range.Cells(1, 2).Value = "first_setup.xlsb"
    lo.ListRows(3).Range.Cells(1, 2).Value = "third_setup.xlsb"

    'Act and assert
    Assert.AreEqual CLng(2), EventsDesignerMulti.CountBuildRows(lo), _
                    "The driver should count the two rows whose setups cell is filled."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestCountBuildRowsCountsFilledSetupCells", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.Driver")
Public Sub TestWriteRowEntriesLandsRowValuesOnMain()
    CustomTestSetTitles Assert, "DesignerMulti", "TestWriteRowEntriesLandsRowValuesOnMain"
    On Error GoTo Fail

    'Arrange: a Main-shaped sheet carrying the entry ranges the build reads
    Dim mainSheet As Worksheet
    Set mainSheet = EnsureWorksheet("Main", FixtureWorkbook)
    FixtureWorkbook.Names.Add Name:="RNG_PathDico", RefersTo:=mainSheet.Range("A1")
    FixtureWorkbook.Names.Add Name:="RNG_PathGeo", RefersTo:=mainSheet.Range("A2")
    FixtureWorkbook.Names.Add Name:="RNG_LLDir", RefersTo:=mainSheet.Range("A3")
    FixtureWorkbook.Names.Add Name:="RNG_LLName", RefersTo:=mainSheet.Range("A4")
    FixtureWorkbook.Names.Add Name:="RNG_LLPwdOpen", RefersTo:=mainSheet.Range("A5")
    FixtureWorkbook.Names.Add Name:="RNG_LLPassword", RefersTo:=mainSheet.Range("A6")
    FixtureWorkbook.Names.Add Name:="RNG_LangSetup", RefersTo:=mainSheet.Range("A7")
    FixtureWorkbook.Names.Add Name:="RNG_LLForm", RefersTo:=mainSheet.Range("A8")
    FixtureWorkbook.Names.Add Name:="RNG_DefaultEpiWeek", RefersTo:=mainSheet.Range("A9")
    FixtureWorkbook.Names.Add Name:="RNG_DesignLL", RefersTo:=mainSheet.Range("A10")

    'A filled row, with a full path in the output files cell as a re-run reads it
    Dim lo As ListObject
    Set lo = MultiSheet.ListObjects(TABLE_MULTI)
    lo.ListRows(1).Range.Cells(1, 2).Value = "/setups/measles.xlsb"
    lo.ListRows(1).Range.Cells(1, 3).Value = "/geo/geobase.xlsx"
    lo.ListRows(1).Range.Cells(1, 4).Value = "/out"
    lo.ListRows(1).Range.Cells(1, 5).Value = "/out/measles_ll.xlsb"
    lo.ListRows(1).Range.Cells(1, 6).Value = "open-secret"
    lo.ListRows(1).Range.Cells(1, 7).Value = "debug-secret"
    lo.ListRows(1).Range.Cells(1, 8).Value = "ENG"
    lo.ListRows(1).Range.Cells(1, 9).Value = "FRA - Francais"
    lo.ListRows(1).Range.Cells(1, 10).Value = "2"
    lo.ListRows(1).Range.Cells(1, 11).Value = "Standard"

    Dim entry As DesignerEntry
    Set entry = DesignerEntry.Create(mainSheet)

    'Act
    EventsDesignerMulti.WriteRowEntries lo, 1, entry

    'Assert: the row's values read back through the entry keys
    Assert.AreEqual "/setups/measles.xlsb", entry.ValueOf("setuppath"), _
                    "The setups cell should land on the setup path entry."
    Assert.AreEqual "/geo/geobase.xlsx", entry.ValueOf("geopath"), _
                    "The geobases cell should land on the geo path entry."
    Assert.AreEqual "/out", entry.ValueOf("lldir"), _
                    "The output folders cell should land on the output folder entry."
    Assert.AreEqual "measles_ll", entry.ValueOf("llname"), _
                    "The output files cell should land as the bare file name."
    Assert.AreEqual "open-secret", entry.ValueOf("llpassword"), _
                    "The password cell should land on the open password entry."
    Assert.AreEqual "debug-secret", entry.ValueOf("debugpassword"), _
                    "The debugging password cell should land on the debug password entry."
    Assert.AreEqual "ENG", entry.ValueOf("setuplang"), _
                    "The dictionary language cell should land on the setup language entry."
    Assert.AreEqual "FRA - Francais", entry.ValueOf("lllang"), _
                    "The interface language cell should land on the linelist language entry."
    Assert.AreEqual "2", entry.ValueOf("epiweekstart"), _
                    "The epiweek cell should land on the epiweek entry."
    Assert.AreEqual "Standard", entry.ValueOf("design"), _
                    "The design cell should land on the design entry."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestWriteRowEntriesLandsRowValuesOnMain", Err.Number, Err.Description
End Sub


'@section Test helpers
'===============================================================================

'@sub-title Create a T_Multi ListObject with the settled headers
'@details
'Writes the T_Multi header row (the ID column first, then the eleven
'columns) and one empty data row on the supplied worksheet, then
'converts the range to a ListObject named T_Multi.
'@param sh Worksheet. The worksheet to create the table on.
Private Sub CreateMultiTable(ByVal sh As Worksheet)
    Dim headers As Variant
    headers = Array("ID", "setups", "geobases", "output folders", "output files", _
                    "output file password", "output file debugging password", _
                    "language of the dictionary", "language of the interface", _
                    "epiweek start", "design", "result")

    Dim idx As Long
    For idx = LBound(headers) To UBound(headers)
        sh.Cells(1, idx - LBound(headers) + 1).Value = headers(idx)
    Next idx

    'Add one empty data row so DataBodyRange exists
    sh.Cells(2, 1).Value = vbNullString

    Dim dataRange As Range
    Set dataRange = sh.Range(sh.Cells(1, 1), sh.Cells(2, UBound(headers) - LBound(headers) + 1))

    Dim lo As ListObject
    Set lo = sh.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=dataRange, _
        XlListObjectHasHeaders:=xlYes)
    lo.Name = TABLE_MULTI
End Sub
