Attribute VB_Name = "TestDesignerMulti"
Attribute VB_Description = "Unit tests for Multi group table operations"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates Multi group table operations: add rows, remove rows, duplicate, import, and export on T_Multi.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private MultiSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TABLE_MULTI As String = "T_Multi"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    TestHelpers.BusyApp
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
    TestHelpers.RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    TestHelpers.BusyApp

    Set FixtureWorkbook = TestHelpers.NewWorkbook
    Set MultiSheet = TestHelpers.EnsureWorksheet("GenerateMultiple", FixtureWorkbook)
    CreateMultiTable MultiSheet
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        TestHelpers.DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set MultiSheet = Nothing
    Set FixtureWorkbook = Nothing

    TestHelpers.RestoreApp
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

    Dim table As ICustomTable
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
    lo.ListRows(1).Range.Cells(1, 1).Value = "test_value"
    Dim filledRowCount As Long
    filledRowCount = 1

    Dim table As ICustomTable
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

    lo.ListRows(1).Range.Cells(1, 1).Value = "setup_path.xlsb"
    lo.ListRows(1).Range.Cells(1, 2).Value = "geo_path.xlsx"
    lo.ListRows(1).Range.Cells(1, 3).Value = "C:\output"

    Dim originalCount As Long
    originalCount = lo.ListRows.Count

    'Act: insert a duplicate after row 1
    lo.ListRows.Add Position:=2
    lo.ListRows(2).Range.Value = lo.ListRows(1).Range.Value

    'Assert
    Assert.AreEqual originalCount + 1, lo.ListRows.Count, _
                    "Duplicate should add one row."
    Assert.AreEqual "setup_path.xlsb", CStr(lo.ListRows(2).Range.Cells(1, 1).Value), _
                    "Duplicated row should have same setups value."
    Assert.AreEqual "geo_path.xlsx", CStr(lo.ListRows(2).Range.Cells(1, 2).Value), _
                    "Duplicated row should have same geobases value."
    Assert.AreEqual "C:\output", CStr(lo.ListRows(2).Range.Cells(1, 3).Value), _
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
    Set sourceSheet = TestHelpers.EnsureWorksheet("SourceMulti", FixtureWorkbook)
    CreateMultiTable sourceSheet
    sourceSheet.ListObjects(TABLE_MULTI).Name = "T_Multi_Source"

    Dim sourceLo As ListObject
    Set sourceLo = sourceSheet.ListObjects("T_Multi_Source")
    sourceLo.ListRows(1).Range.Cells(1, 1).Value = "imported_setup.xlsb"
    sourceLo.ListRows(1).Range.Cells(1, 2).Value = "imported_geo.xlsx"

    Dim sourceTable As ICustomTable
    Set sourceTable = CustomTable.Create(sourceLo)

    'Target table
    Dim targetLo As ListObject
    Set targetLo = MultiSheet.ListObjects(TABLE_MULTI)
    targetLo.ListRows(1).Range.Cells(1, 1).Value = "old_setup.xlsb"

    Dim targetTable As ICustomTable
    Set targetTable = CustomTable.Create(targetLo)

    'Act
    targetTable.Import sourceTable

    'Assert
    Assert.AreEqual "imported_setup.xlsb", _
                    CStr(targetLo.ListRows(1).Range.Cells(1, 1).Value), _
                    "Import should replace setups value with source data."
    Assert.AreEqual "imported_geo.xlsx", _
                    CStr(targetLo.ListRows(1).Range.Cells(1, 2).Value), _
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
    lo.ListRows(1).Range.Cells(1, 1).Value = "export_setup.xlsb"
    lo.ListRows(1).Range.Cells(1, 2).Value = "export_geo.xlsx"

    Dim table As ICustomTable
    Set table = CustomTable.Create(lo)

    Dim exportSheet As Worksheet
    Set exportSheet = TestHelpers.EnsureWorksheet("ExportTarget", FixtureWorkbook)

    'Act
    table.Export sh:=exportSheet, startLine:=1, startColumn:=1, addListObject:=True

    'Assert: the export sheet should have a ListObject with the same headers
    Assert.IsTrue exportSheet.ListObjects.Count > 0, _
                  "Export should create a ListObject on the target sheet."

    Dim exportLo As ListObject
    Set exportLo = exportSheet.ListObjects(1)
    Assert.AreEqual lo.ListColumns.Count, exportLo.ListColumns.Count, _
                    "Exported table should have the same number of columns."
    Assert.AreEqual "setups", exportLo.ListColumns(1).Name, _
                    "First column header should be 'setups'."
    Assert.AreEqual "export_setup.xlsb", _
                    CStr(exportLo.ListRows(1).Range.Cells(1, 1).Value), _
                    "Exported data should match source data."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestExportWritesToWorksheet", Err.Number, Err.Description
End Sub


'@section Test helpers
'===============================================================================

'@sub-title Create a T_Multi ListObject with the expected headers
'@details
'Writes the T_Multi header row and one empty data row on the supplied
'worksheet, then converts the range to a ListObject named T_Multi.
'@param sh Worksheet. The worksheet to create the table on.
Private Sub CreateMultiTable(ByVal sh As Worksheet)
    Dim headers As Variant
    headers = Array("setups", "geobases", "output folders", "output files", _
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
