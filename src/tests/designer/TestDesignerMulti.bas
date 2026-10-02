Attribute VB_Name = "TestDesignerMulti"
Attribute VB_Description = "Unit tests for Multi group table operations"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates Multi group table operations: add rows, remove rows, duplicate, import, export, the write-once row IDs on T_Multi, the driver helpers that read a row into the Main entries, the ribbon file the run gives every row, the report section title of a row, the shared setup language extraction, the generation drivers: GenerateOne raising a step outcome in place and through a host, the file-name pre-flight, and on Windows the single build and the multi loop through a hidden instance, and the summary a script reads a generation run back from.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'THE HOSTED BUILDS RUN OVER A COPY OF THIS DRIVER
'-------------------------------------------------------------------------------
'The instance tests copy this driver into a hidden Excel, the way a
'designer press copies the designer, because the driver is the one
'workbook of the run whose project carries BuildSteps and compiles. A Main
'sheet with the entry ranges is stood up on the driver for the test and
'taken off in the cleanup. The first step fails in the instance, since
'the copy is no designer, and that failure is the answer under test: it
'crossed the processes as an outcome string, the entries crossed the other
'way with the step, and the instance is still there to release.
'
'THE PROCESS COUNT
'-------------------------------------------------------------------------------
'The instance tests count the Excel processes of the machine before and
'after, through WMI, so a leaked instance fails the test. Another Excel
'starting or stopping on the machine during one of these tests moves the
'count and fails it too; run them again when that happens.
'@depends EventsDesignerMulti, EventsDesignerAdvanced, BuildSteps, GenerationHost, DesignerEntry, ProgressBar, GenerationLog, CustomTest, TestHelpersLite

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private MultiSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TABLE_MULTI As String = "T_Multi"
Private Const SHEET_MAIN As String = "Main"

'The host of the running instance test, released by TestCleanup whatever the exit
Private heldHost As GenerationHost

'True while a test holds a Main sheet of its own on this driver, so the
'cleanup takes off what the test made and nothing else
Private driverMainMade As Boolean

'The folders beside this driver: the designer copy, and the output folder
'the builds point at
Private Const COPY_FOLDER_NAME As String = "multi_copy"
Private Const OUTPUT_FOLDER_NAME As String = "multi_out"

'The eleven entry ranges DesignerEntry writes, in the order they take the
'rows of a Main sheet stood up for a test
Private Const ENTRY_RANGE_NAMES As String = "RNG_PathDico,RNG_PathGeo,RNG_LLDir,RNG_LLName," & _
                                            "RNG_LLPwdOpen,RNG_LLPassword,RNG_LangSetup,RNG_LLForm," & _
                                            "RNG_DefaultEpiWeek,RNG_DesignLL,RNG_LLTemp"

'How long a process count is waited for, in seconds
Private Const PROCESS_WAIT_SECONDS As Long = 30

'What closes the keys of DesignerLastSummary and opens its free text
Private Const REPORT_MARKER As String = "--report--"


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
        ReleaseHeldHost
        RemoveDriverMain
        ClearOutputFolder
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
        ReleaseHeldHost
        RemoveDriverMain
        ClearOutputFolder
        DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set MultiSheet = Nothing
    Set FixtureWorkbook = Nothing

    'A hosted build can hand the screen to another workbook
    ThisWorkbook.Activate

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

'@TestMethod("DesignerMulti.Driver")
Public Sub TestWriteRowEntriesLandsTheRunRibbonFile()
    CustomTestSetTitles Assert, "DesignerMulti", "TestWriteRowEntriesLandsTheRunRibbonFile"
    On Error GoTo Fail

    'Arrange: a Main-shaped sheet carrying the template range the build reads
    Dim mainSheet As Worksheet
    Set mainSheet = EnsureWorksheet("Main", FixtureWorkbook)
    FixtureWorkbook.Names.Add Name:="RNG_LLTemp", RefersTo:=mainSheet.Range("A11")
    mainSheet.Range("A11").Value = "/templates/from_the_main_sheet.xlsb"

    Dim lo As ListObject
    Set lo = MultiSheet.ListObjects(TABLE_MULTI)
    lo.ListRows(1).Range.Cells(1, 2).Value = "/setups/measles.xlsb"

    Dim entry As DesignerEntry
    Set entry = DesignerEntry.Create(mainSheet)

    'Act: the run's own ribbon file rides with the row
    EventsDesignerMulti.WriteRowEntries lo, 1, entry, "/ribbons/run_template.xlsb"

    'Assert: the row builds with the run's file, and the Main cell is overwritten
    Assert.AreEqual "/ribbons/run_template.xlsb", entry.ValueOf("temppath"), _
                    "The ribbon file of the run should land on the template entry."

    'Act: a run with no ribbon file clears the entry, so the buttons build
    EventsDesignerMulti.WriteRowEntries lo, 1, entry

    'Assert
    Assert.AreEqual vbNullString, entry.ValueOf("temppath"), _
                    "A run with no ribbon file should leave the template entry empty."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestWriteRowEntriesLandsTheRunRibbonFile", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.Driver")
Public Sub TestRowSectionTitleNamesTheRowAndItsSetup()
    CustomTestSetTitles Assert, "DesignerMulti", "TestRowSectionTitleNamesTheRowAndItsSetup"
    On Error GoTo Fail

    'Arrange
    Dim lo As ListObject
    Set lo = MultiSheet.ListObjects(TABLE_MULTI)
    lo.ListRows(1).Range.Cells(1, 1).Value = "Operation- 3"

    'Act and assert: the ID and the file name of the setup
    Assert.AreEqual "Operation- 3 - measles.xlsb", _
                    EventsDesignerMulti.RowSectionTitle(lo, 1, "/setups/measles.xlsb"), _
                    "The section title should name the row ID and the setup file."

    'Act and assert: a row with no ID is named by its position
    lo.ListRows(1).Range.Cells(1, 1).Value = vbNullString
    Assert.AreEqual "row 1 - measles.xlsb", _
                    EventsDesignerMulti.RowSectionTitle(lo, 1, "/setups/measles.xlsb"), _
                    "A row with no ID should be named by its position in the table."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestRowSectionTitleNamesTheRowAndItsSetup", Err.Number, Err.Description
End Sub


'@section SetupLanguages Tests
'===============================================================================
'@TestMethod("DesignerMulti.SetupLanguages")
Public Sub TestSetupLanguagesReadsHiddenNamesList()
    CustomTestSetTitles Assert, "DesignerMulti", "TestSetupLanguagesReadsHiddenNamesList"
    On Error GoTo Fail

    'Arrange: the persisted language list of a setup Translations sheet.
    'The list wins over anything a table on the sheet carries.
    Dim tradSheet As Worksheet
    Set tradSheet = EnsureWorksheet("Translations", FixtureWorkbook)

    Dim store As HiddenNames
    Set store = HiddenNames.Create(tradSheet)
    store.EnsureName SetupTranslationsTable.LanguagesNameId, "ENG;FRA", HiddenNameTypeString
    store.SetValue SetupTranslationsTable.LanguagesNameId, "ENG;FRA"

    'Act
    Dim langValues As BetterArray
    Set langValues = EventsDesignerAdvanced.SetupLanguages(tradSheet)

    'Assert
    Assert.AreEqual CLng(2), langValues.Length, _
                    "The persisted list should answer its two languages."
    Assert.AreEqual "ENG", CStr(langValues.Item(1)), _
                    "The first language should be the first list entry."
    Assert.AreEqual "FRA", CStr(langValues.Item(2)), _
                    "The second language should be the second list entry."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestSetupLanguagesReadsHiddenNamesList", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.SetupLanguages")
Public Sub TestSetupLanguagesFallbackDropsTagColumns()
    CustomTestSetTitles Assert, "DesignerMulti", "TestSetupLanguagesFallbackDropsTagColumns"
    On Error GoTo Fail

    'Arrange: no persisted list, so the fallback reads the header row of
    'the first ListObject. The internal tag column is table machinery and
    'used to land in the language dropdown.
    Dim tradSheet As Worksheet
    Set tradSheet = EnsureWorksheet("Translations", FixtureWorkbook)

    tradSheet.Cells(1, 1).Value = "ENG"
    tradSheet.Cells(1, 2).Value = "FRA"
    tradSheet.Cells(1, 3).Value = "__TagInternal__"
    tradSheet.Cells(2, 1).Value = "hello"
    tradSheet.Cells(2, 2).Value = "bonjour"
    tradSheet.Cells(2, 3).Value = "k1"

    Dim lo As ListObject
    Set lo = tradSheet.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=tradSheet.Range(tradSheet.Cells(1, 1), tradSheet.Cells(2, 3)), _
        XlListObjectHasHeaders:=xlYes)
    lo.Name = "Tab_Translations"

    'Act
    Dim langValues As BetterArray
    Set langValues = EventsDesignerAdvanced.SetupLanguages(tradSheet)

    'Assert
    Assert.AreEqual CLng(2), langValues.Length, _
                    "The fallback should answer the two language headers."
    Assert.AreEqual "ENG", CStr(langValues.Item(1)), _
                    "The first header should be kept."
    Assert.AreEqual "FRA", CStr(langValues.Item(2)), _
                    "The second header should be kept."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestSetupLanguagesFallbackDropsTagColumns", Err.Number, Err.Description
End Sub


'@section Generation driver Tests
'===============================================================================
'@TestMethod("DesignerMulti.Generation")
Public Sub TestGenerateOneRaisesTheStepOutcome()
    CustomTestSetTitles Assert, "DesignerMulti", "TestGenerateOneRaisesTheStepOutcome"
    On Error GoTo Fail

    'Arrange: a Main over the fixture, its setup entry naming a missing file,
    'and a run log over the fixture so the checkings have somewhere to land
    Dim entry As DesignerEntry
    Set entry = MainEntry(FixtureWorkbook)
    FillEntries entry, MissingSetupPath()
    EventsDesignerAdvanced.StartRunLog entry.ValueOf("setuppath"), entry.ValueOf("llname"), FixtureWorkbook

    'Act: the first step fails inside its own handler and GenerateOne raises
    'what it answered
    Dim errNumber As Long
    Dim errDesc As String
    On Error Resume Next
    EventsDesignerAdvanced.GenerateOne entry
    errNumber = Err.Number
    errDesc = Err.Description
    On Error GoTo Fail

    EventsDesignerAdvanced.FinishRunLog "test run"

    'Assert
    Assert.IsTrue errNumber <> 0, "A step that failed should raise out of GenerateOne."
    Assert.IsTrue LenB(errDesc) > 0, "The raise should carry the description of the step's fault."
    Assert.IsFalse Left$(errDesc, 6) = "ERROR ", _
                   "The description should be the fault alone, the lead stripped: " & errDesc
    AssertKeptPathIsOnDisk

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestGenerateOneRaisesTheStepOutcome", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.Generation")
Public Sub TestGenerateOneWithAnInPlaceHostRunsTheStepsHere()
    CustomTestSetTitles Assert, "DesignerMulti", "TestGenerateOneWithAnInPlaceHostRunsTheStepsHere"
    On Error GoTo Fail

    'Arrange: a host on the in-place path, never acquired; the steps run in
    'this project the way they do with no host at all
    Dim entry As DesignerEntry
    Set entry = MainEntry(FixtureWorkbook)
    FillEntries entry, MissingSetupPath()
    EventsDesignerAdvanced.StartRunLog entry.ValueOf("setuppath"), entry.ValueOf("llname"), FixtureWorkbook

    Dim host As GenerationHost
    Set host = GenerationHost.Create(FixtureWorkbook, HostPathInPlace)

    Dim bar As ProgressBar
    Set bar = ProgressBar.Create(MultiSheet.Range("N1:R1"), 3)

    'Act
    Dim errNumber As Long
    Dim errDesc As String
    On Error Resume Next
    EventsDesignerAdvanced.GenerateOne entry, Nothing, host, bar
    errNumber = Err.Number
    errDesc = Err.Description
    On Error GoTo Fail

    EventsDesignerAdvanced.FinishRunLog "test run"

    'Assert: the fault came back, the host was never touched, the bar sized and unmoved
    Assert.IsTrue errNumber <> 0, "A step that failed should raise out of GenerateOne."
    Assert.IsTrue LenB(errDesc) > 0, "The raise should carry the description of the step's fault."
    Assert.IsFalse host.IsAcquired, "An in-place host needs no Acquire for the steps to run."
    Assert.AreEqual CLng(5), bar.Maximum, "The bar should be sized to the fixed steps before the count is known."
    Assert.AreEqual CLng(0), bar.Value, "No step finished, so the bar should still stand at zero."
    AssertKeptPathIsOnDisk

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestGenerateOneWithAnInPlaceHostRunsTheStepsHere", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.Generation")
Public Sub TestCheckBuildFileNamesRefusesASharedNameAndFilesIt()
    CustomTestSetTitles Assert, "DesignerMulti", "TestCheckBuildFileNamesRefusesASharedNameAndFilesIt"
    On Error GoTo Fail

    'Arrange: a setup and a template sharing a file name in two folders, and
    'a run log to take the refusal
    Dim entry As DesignerEntry
    Set entry = MainEntry(FixtureWorkbook)
    FillEntries entry, JoinPath(OutputFolder(), "same_name.xlsb")
    entry.AddInfo JoinPath(OutputFolder(), "elsewhere", "same_name.xlsb"), "temppath"

    Dim runLog As GenerationLog
    Set runLog = EventsDesignerAdvanced.StartRunLog(vbNullString, vbNullString, FixtureWorkbook)

    Dim linesBefore As Long
    linesBefore = runLog.RecordLength

    Dim host As GenerationHost
    Set host = GenerationHost.Create(FixtureWorkbook, HostPathInPlace)

    'Act
    Dim clashText As String
    clashText = EventsDesignerAdvanced.CheckBuildFileNames(host, entry, _
                                                           JoinPath(OutputFolder(), "OBTApp_", "__temp.xlsb"))
    EventsDesignerAdvanced.FinishRunLog "test run"

    'Assert: both paths are named, and the refusal is in the report
    Assert.IsTrue InStr(1, clashText, "share the name") > 0, _
                  "A setup and a template of one name should be refused: " & clashText
    Assert.IsTrue InStr(1, clashText, entry.ValueOf("setuppath")) > 0, _
                  "The refusal should name the setup path."
    Assert.IsTrue InStr(1, clashText, entry.ValueOf("temppath")) > 0, _
                  "The refusal should name the template path."
    Assert.IsTrue runLog.RecordLength > linesBefore, _
                  "The refusal should land in the run log."

    'Act and assert: distinct names pass
    entry.AddInfo vbNullString, "temppath"
    Assert.AreEqual vbNullString, _
                    EventsDesignerAdvanced.CheckBuildFileNames(host, entry, _
                                                               JoinPath(OutputFolder(), "OBTApp_", "__temp.xlsb")), _
                    "Distinct names should answer no refusal."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestCheckBuildFileNamesRefusesASharedNameAndFilesIt", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.Generation")
Public Sub TestGenerateOneThroughTheInstanceCrossesTheEntries()
    CustomTestSetTitles Assert, "DesignerMulti", "TestGenerateOneThroughTheInstanceCrossesTheEntries"
    On Error GoTo Fail

    If Not InstancePathAvailable() Then
        Assert.IsTrue True, "The instance path exists on Windows alone; nothing to check here."
        Exit Sub
    End If

    'Arrange: a Main on this driver with its entries EMPTY, copied into the
    'instance; the entries are written after the copy, so the only way they
    'reach the copy's Main is across Run with the first step
    Dim before As Long
    before = ExcelProcessCount()

    Dim entry As DesignerEntry
    Set entry = MainEntry(ThisWorkbook)

    Set heldHost = GenerationHost.Create(ThisWorkbook, HostPathInstance)
    heldHost.Acquire
    heldHost.OpenDesignerCopy CopyFolder()

    FillEntries entry, MissingSetupPath()
    EventsDesignerAdvanced.StartRunLog entry.ValueOf("setuppath"), entry.ValueOf("llname"), FixtureWorkbook

    Dim bar As ProgressBar
    Set bar = ProgressBar.Create(MultiSheet.Range("N1:R1"), 3)

    'Act
    Dim errNumber As Long
    Dim errDesc As String
    On Error Resume Next
    EventsDesignerAdvanced.GenerateOne entry, Nothing, heldHost, bar
    errNumber = Err.Number
    errDesc = Err.Description
    On Error GoTo Fail

    'Assert: the step failed in the instance and answered, the entries reached
    'the copy, the instance is still alive, the bar stands at zero
    Assert.IsTrue errNumber <> 0, "The step that failed in the instance should raise out of GenerateOne."
    Assert.IsTrue LenB(errDesc) > 0, "The raise should carry the description the step answered."
    Assert.IsFalse InStr(1, errDesc, "stopped answering") > 0, _
                   "A step that answered should leave the instance alive: " & errDesc
    Assert.IsFalse heldHost.InstanceStopped, "The instance should be marked alive after an answered step."
    Assert.AreEqual MissingSetupPath(), CopyMainValue("RNG_PathDico"), _
                    "The setup entry should have reached the copy's Main across Run."
    Assert.AreEqual OutputFolder(), CopyMainValue("RNG_LLDir"), _
                    "The output folder entry should have reached the copy's Main across Run."
    Assert.AreEqual CLng(0), bar.Value, "No step finished, so the bar should still stand at zero."
    AssertKeptPathIsOnDisk

    'Act: the instance goes away behind the host's back, and the next build
    'answers the stopped outcome without a call
    heldHost.DesignerCopy.Close SaveChanges:=False
    heldHost.HostApplication.Quit

    On Error Resume Next
    EventsDesignerAdvanced.GenerateOne entry, Nothing, heldHost, bar
    errNumber = Err.Number
    errDesc = Err.Description
    On Error GoTo Fail

    EventsDesignerAdvanced.FinishRunLog "test run"

    'Assert: the stopped outcome names the first step, the abort keeps nothing
    Assert.IsTrue errNumber <> 0, "A build over an instance gone should raise."
    Assert.IsTrue InStr(1, errDesc, "stopped answering after BuildSteps.BuildBeginEntries") > 0, _
                  "The raise should say the instance stopped answering after the first step: " & errDesc
    Assert.IsTrue heldHost.InstanceStopped, "The instance should be marked stopped."
    Assert.AreEqual vbNullString, EventsDesignerAdvanced.AbortBuild(heldHost), _
                    "An abort over an instance gone should keep nothing."

    'Act and assert: the release names the handle and the process leaves
    Dim outcome As String
    outcome = heldHost.ReleaseInstance()
    Assert.IsTrue Left$(outcome, 6) = "ERROR ", _
                  "ReleaseInstance after the instance is gone should answer an error outcome: " & outcome
    Set heldHost = Nothing
    Assert.AreEqual before, WaitForProcessCount(before), _
                    "No Excel process should be left behind."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestGenerateOneThroughTheInstanceCrossesTheEntries", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.Generation")
Public Sub TestGenerateMultipleRowsThroughTheInstanceKeepsGoing()
    CustomTestSetTitles Assert, "DesignerMulti", "TestGenerateMultipleRowsThroughTheInstanceKeepsGoing"
    On Error GoTo Fail

    If Not InstancePathAvailable() Then
        Assert.IsTrue True, "The instance path exists on Windows alone; nothing to check here."
        Exit Sub
    End If

    'Arrange: three rows over one host and one copy. The first names a
    'setup that is on disk and builds in the instance until the copy, which
    'is no designer, refuses; the second names a setup carrying the
    'template's file name; the third names a setup that is not on disk.
    'The setups are scratch workbooks, so the test leans on no asset file.
    Dim before As Long
    before = ExcelProcessCount()

    Dim rowOneSetup As String
    Dim sharedNameSetup As String
    Dim templatePath As String
    rowOneSetup = JoinPath(OutputFolder(), "row_one_setup.xlsb")
    sharedNameSetup = JoinPath(OutputFolder(), "shared", "shared_name.xlsb")
    templatePath = JoinPath(OutputFolder(), "shared_name.xlsb")
    MakeScratchWorkbook rowOneSetup
    MakeScratchWorkbook sharedNameSetup
    MakeScratchWorkbook templatePath

    Dim entry As DesignerEntry
    Set entry = MainEntry(ThisWorkbook)

    Dim lo As ListObject
    Set lo = MultiSheet.ListObjects(TABLE_MULTI)
    lo.ListRows.Add
    lo.ListRows.Add
    FillMultiRow lo, 1, rowOneSetup
    FillMultiRow lo, 2, sharedNameSetup
    FillMultiRow lo, 3, MissingSetupPath()

    EventsDesignerAdvanced.StartRunLog vbNullString, vbNullString, FixtureWorkbook

    Set heldHost = GenerationHost.Create(ThisWorkbook, HostPathInstance)
    heldHost.Acquire
    heldHost.OpenDesignerCopy CopyFolder()

    'Act
    Dim builtCount As Long
    Dim failedCount As Long
    EventsDesignerMulti.GenerateMultipleRows lo, entry, Nothing, builtCount, failedCount, _
                                             templatePath, heldHost

    EventsDesignerAdvanced.FinishRunLog "test run"

    'Assert: every row answered in its result cell and the loop kept going
    Assert.AreEqual CLng(0), builtCount, "No row can build over a copy that is no designer."
    Assert.AreEqual CLng(3), failedCount, "The three rows should each count as failed or refused."
    Assert.IsTrue Left$(ResultOf(lo, 1), 7) = "Failed:", _
                  "The first row should fail inside the instance: " & ResultOf(lo, 1)
    Assert.IsFalse InStr(1, ResultOf(lo, 1), "stopped answering") > 0, _
                   "The first row's fault should come from the step, with the instance alive: " & ResultOf(lo, 1)
    Assert.IsTrue Left$(ResultOf(lo, 2), 8) = "Refused:", _
                  "The second row should be refused before its build: " & ResultOf(lo, 2)
    Assert.IsTrue InStr(1, ResultOf(lo, 2), sharedNameSetup) > 0, _
                  "The refusal should name the setup path: " & ResultOf(lo, 2)
    Assert.IsTrue InStr(1, ResultOf(lo, 2), templatePath) > 0, _
                  "The refusal should name the template path: " & ResultOf(lo, 2)
    Assert.IsTrue Left$(ResultOf(lo, 3), 8) = "Refused:", _
                  "The third row should be refused by the entry checks: " & ResultOf(lo, 3)
    Assert.IsFalse heldHost.InstanceStopped, "The instance should be alive after the three rows."
    Assert.AreEqual rowOneSetup, CopyMainValue("RNG_PathDico"), _
                    "The copy's Main should carry the entries of the last row that reached its build."

    'Act and assert: the release quits the instance and leaves no process
    Assert.AreEqual "OK", heldHost.ReleaseInstance(), "ReleaseInstance should answer OK."
    Set heldHost = Nothing
    Assert.AreEqual before, WaitForProcessCount(before), _
                    "No Excel process should be left behind."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestGenerateMultipleRowsThroughTheInstanceKeepsGoing", Err.Number, Err.Description
End Sub


'@section Run summary Tests
'===============================================================================
'DesignerLastSummary is what a script reads a generation run back from, since
'the answer of Application.Run carries nothing about what was built. The keys
'are read by position on the R side, so their order is under test here.

'@TestMethod("DesignerMulti.Generation")
Public Sub TestLastSummaryAnswersItsKeysInOrder()
    CustomTestSetTitles Assert, "DesignerMulti", "TestLastSummaryAnswersItsKeysInOrder"
    On Error GoTo Fail

    'Arrange: the run log a run opens before anything is built
    Dim entry As DesignerEntry
    Set entry = MainEntry(FixtureWorkbook)
    FillEntries entry, MissingSetupPath()
    EventsDesignerAdvanced.StartRunLog entry.ValueOf("setuppath"), entry.ValueOf("llname"), FixtureWorkbook

    'Act
    Dim summaryLines() As String
    summaryLines = Split(EventsDesignerAdvanced.DesignerLastSummary(), vbLf)

    'Assert: seven keys in the settled order, then the marker
    Assert.IsTrue UBound(summaryLines) >= 7, _
                  "The summary should carry seven keys and the marker."
    Assert.AreEqual "outcome=", summaryLines(0), _
                    "outcome= leads the block, empty until a run answers."
    Assert.AreEqual "linelist=", summaryLines(1), "linelist= comes second."
    Assert.AreEqual "log=", summaryLines(2), "log= comes third."
    Assert.AreEqual "sheets=0", summaryLines(3), "sheets= counts the data entry sheets built."
    Assert.AreEqual "variables=0", summaryLines(4), "variables= counts the variables written."
    Assert.AreEqual "built=0", summaryLines(5), "built= counts the linelists of the run."
    Assert.AreEqual "failed=0", summaryLines(6), "failed= counts the builds that did not finish."
    Assert.AreEqual REPORT_MARKER, summaryLines(7), "The marker closes the keys."

    EventsDesignerAdvanced.FinishRunLog "test run"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestLastSummaryAnswersItsKeysInOrder", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.Generation")
Public Sub TestLastSummaryCarriesTheRunLogUnderTheMarker()
    CustomTestSetTitles Assert, "DesignerMulti", "TestLastSummaryCarriesTheRunLogUnderTheMarker"
    On Error GoTo Fail

    'Arrange: a run that opened its log, failed at its first step and closed
    Dim entry As DesignerEntry
    Set entry = MainEntry(FixtureWorkbook)
    FillEntries entry, MissingSetupPath()
    EventsDesignerAdvanced.StartRunLog entry.ValueOf("setuppath"), entry.ValueOf("llname"), FixtureWorkbook

    On Error Resume Next
    EventsDesignerAdvanced.GenerateOne entry
    On Error GoTo Fail

    EventsDesignerAdvanced.FinishRunLog "the run under test"

    'Act
    Dim freeText As String
    freeText = FreeTextOf(EventsDesignerAdvanced.DesignerLastSummary())

    'Assert: the record of the run is the free text, one line per entry
    Assert.IsTrue LenB(freeText) > 0, "The free text should carry the record of the run."
    Assert.IsTrue InStr(1, freeText, "the run under test") > 0, _
                  "The closing outcome should be in the free text: " & freeText
    Assert.IsTrue InStr(1, freeText, vbLf) > 0, _
                  "The record should read one line per entry."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestLastSummaryCarriesTheRunLogUnderTheMarker", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerMulti.Generation")
Public Sub TestStartRunLogEmptiesTheSummaryOfTheRunBefore()
    CustomTestSetTitles Assert, "DesignerMulti", "TestStartRunLogEmptiesTheSummaryOfTheRunBefore"
    On Error GoTo Fail

    'Arrange: one run, closed with an outcome of its own
    Dim entry As DesignerEntry
    Set entry = MainEntry(FixtureWorkbook)
    FillEntries entry, MissingSetupPath()
    EventsDesignerAdvanced.StartRunLog entry.ValueOf("setuppath"), entry.ValueOf("llname"), FixtureWorkbook
    EventsDesignerAdvanced.FinishRunLog "the run before"

    Assert.IsTrue InStr(1, FreeTextOf(EventsDesignerAdvanced.DesignerLastSummary()), _
                        "the run before") > 0, _
                  "The first run should be in its own summary."

    'Act: the next run opens its log
    EventsDesignerAdvanced.StartRunLog entry.ValueOf("setuppath"), entry.ValueOf("llname"), FixtureWorkbook

    'Assert: nothing of the run before is left to read
    Dim summaryText As String
    summaryText = EventsDesignerAdvanced.DesignerLastSummary()

    Assert.IsFalse InStr(1, FreeTextOf(summaryText), "the run before") > 0, _
                   "A new run should not answer the record of the run before it."
    Assert.IsTrue InStr(1, summaryText, "outcome=" & vbLf) > 0, _
                  "A new run should answer an empty outcome."
    Assert.IsTrue InStr(1, summaryText, "built=0") > 0, _
                  "A new run should count no linelist built."

    EventsDesignerAdvanced.FinishRunLog "test run"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestStartRunLogEmptiesTheSummaryOfTheRunBefore", Err.Number, Err.Description
End Sub


'@section Generation driver helpers
'===============================================================================

'@sub-title The free text of a summary: everything past the report marker
'@param summaryText String. What DesignerLastSummary answered.
'@return String. The text after the marker, empty when there is none.
Private Function FreeTextOf(ByVal summaryText As String) As String
    Dim markerAt As Long

    markerAt = InStr(1, summaryText, REPORT_MARKER)
    If markerAt = 0 Then Exit Function

    FreeTextOf = Mid$(summaryText, markerAt + Len(REPORT_MARKER) + 1)
End Function

'@sub-title Whether this Excel can start a second instance
'@return Boolean. True on Windows.
Private Function InstancePathAvailable() As Boolean
    #If Mac Then
        InstancePathAvailable = False
    #Else
        InstancePathAvailable = True
    #End If
End Function

'@sub-title The number of Excel processes on the machine
'@return Long. The count, or -1 when it cannot be read.
Private Function ExcelProcessCount() As Long
    #If Mac Then
        ExcelProcessCount = -1
    #Else
        Dim management As Object
        Dim processes As Object

        On Error GoTo Unreadable
        Set management = GetObject("winmgmts:\\.\root\cimv2")
        Set processes = management.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
        ExcelProcessCount = processes.Count
        On Error GoTo 0
        Exit Function

Unreadable:
        ExcelProcessCount = -1
    #End If
End Function

'@sub-title Wait for the Excel process count to reach a value
'@details
'A quit instance takes a moment to leave the process table, so the count
'is read again until it answers the expected value or the wait runs out.
'@param expected Long. The count waited for.
'@return Long. The last count read.
Private Function WaitForProcessCount(ByVal expected As Long) As Long
    Dim startedAt As Single
    Dim observed As Long

    startedAt = Timer
    Do
        observed = ExcelProcessCount()
        If observed = expected Then Exit Do
        DoEvents
    Loop While Timer - startedAt < PROCESS_WAIT_SECONDS And Timer >= startedAt

    WaitForProcessCount = observed
End Function

'@sub-title The folder the designer copy is written in
'@return String. An existing folder beside this driver.
Private Function CopyFolder() As String
    CopyFolder = BuildTempFolder(ThisWorkbook, COPY_FOLDER_NAME)
End Function

'@sub-title The output folder the builds of this suite point at
'@return String. An existing folder beside this driver.
Private Function OutputFolder() As String
    OutputFolder = BuildTempFolder(ThisWorkbook, OUTPUT_FOLDER_NAME)
End Function

'@sub-title A setup path no file is at
'@return String.
Private Function MissingSetupPath() As String
    MissingSetupPath = JoinPath(OutputFolder(), "missing_setup.xlsb")
End Function

'@sub-title Write an empty workbook at a path, as the setup or template of a row
'@details
'The entry checks want the file on disk and nothing more; the build in the
'instance then refuses it, which is the answer under test.
'@param filePath String. Where the workbook is saved. Its folder is made first.
Private Sub MakeScratchWorkbook(ByVal filePath As String)
    Dim scratch As Workbook

    EnsureFolder ParentFolder(filePath)

    Set scratch = Application.Workbooks.Add
    scratch.SaveAs fileName:=filePath, fileFormat:=xlExcel12
    scratch.Close SaveChanges:=False
End Sub

'@sub-title A Main worksheet carrying the entry ranges the build reads, and an entry over it
'@details
'The eleven named ranges DesignerEntry writes, one cell each, on a Main
'sheet of the workbook. The names are workbook-scoped, the way the
'designer carries them.
'@param targetBook Workbook. The workbook to stand the Main up on.
'@return DesignerEntry. The entry manager over that Main.
Private Function MainEntry(ByVal targetBook As Workbook) As DesignerEntry
    Dim mainSheet As Worksheet
    Dim names() As String
    Dim index As Long

    If targetBook Is ThisWorkbook Then
        driverMainMade = Not WorksheetExists(SHEET_MAIN, ThisWorkbook)
    End If

    Set mainSheet = EnsureWorksheet(SHEET_MAIN, targetBook)

    names = Split(ENTRY_RANGE_NAMES, ",")
    For index = LBound(names) To UBound(names)
        targetBook.Names.Add Name:=names(index), RefersTo:=mainSheet.Cells(index + 1, 1)
    Next index

    Set MainEntry = DesignerEntry.Create(mainSheet)
End Function

'@sub-title Write the entries of one build: a setup, the output folder, a name, the languages, a design
'@param entry DesignerEntry. The entry manager over a Main.
'@param setupPath String. The setup entry.
Private Sub FillEntries(ByVal entry As DesignerEntry, ByVal setupPath As String)
    entry.AddInfo setupPath, "setuppath"
    entry.AddInfo vbNullString, "geopath"
    entry.AddInfo OutputFolder(), "lldir"
    entry.AddInfo "hosted_probe", "llname"
    entry.AddInfo vbNullString, "llpassword"
    entry.AddInfo vbNullString, "debugpassword"
    entry.AddInfo "ENG", "setuplang"
    entry.AddInfo "ENG", "lllang"
    entry.AddInfo "1", "epiweekstart"
    entry.AddInfo "Standard", "design"
    entry.AddInfo vbNullString, "temppath"
End Sub

'@sub-title Fill one T_Multi row with a setup and the values the entry checks want
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@param setupPath String. The setups cell.
Private Sub FillMultiRow(ByVal lo As ListObject, ByVal rowIdx As Long, ByVal setupPath As String)
    With lo.ListRows(rowIdx).Range
        .Cells(1, 2).Value = setupPath
        .Cells(1, 4).Value = OutputFolder()
        .Cells(1, 5).Value = "hosted_row_" & CStr(rowIdx)
        .Cells(1, 8).Value = "ENG"
        .Cells(1, 9).Value = "ENG"
        .Cells(1, 10).Value = "1"
        .Cells(1, 11).Value = "Standard"
    End With
End Sub

'@sub-title The result cell of one T_Multi row
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@return String.
Private Function ResultOf(ByVal lo As ListObject, ByVal rowIdx As Long) As String
    ResultOf = CStr(lo.ListRows(rowIdx).Range.Cells(1, 12).Value)
End Function

'@sub-title One entry read off the Main of the copy in the instance
'@param rangeName String. The named range.
'@return String. The cell text.
Private Function CopyMainValue(ByVal rangeName As String) As String
    Dim copyBook As Workbook

    Set copyBook = heldHost.DesignerCopy
    CopyMainValue = CStr(copyBook.Worksheets(SHEET_MAIN).Range(rangeName).Value)
End Function

'@sub-title The kept file of the last build, when there is one, is on disk
Private Sub AssertKeptPathIsOnDisk()
    Dim keptPath As String

    keptPath = EventsDesignerAdvanced.LastKeptPath()
    If LenB(keptPath) = 0 Then Exit Sub

    Assert.IsTrue LenB(Dir$(keptPath)) > 0, _
                  "A kept path answered by the build should point at a file: " & keptPath
End Sub

'@sub-title Take the Main sheet and its names off this driver
Private Sub RemoveDriverMain()
    Dim names() As String
    Dim index As Long
    Dim previousAlerts As Boolean

    If Not driverMainMade Then Exit Sub

    On Error Resume Next
    names = Split(ENTRY_RANGE_NAMES, ",")
    For index = LBound(names) To UBound(names)
        ThisWorkbook.Names(names(index)).Delete
    Next index

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(SHEET_MAIN).Delete
    Application.DisplayAlerts = previousAlerts
    On Error GoTo 0

    driverMainMade = False
End Sub

'@sub-title ReleaseInstance the host of the running test
Private Sub ReleaseHeldHost()
    On Error Resume Next
    If Not heldHost Is Nothing Then heldHost.ReleaseInstance
    Set heldHost = Nothing
    On Error GoTo 0
End Sub

'@sub-title Empty the output folder of this suite, the scratch folder under it included
Private Sub ClearOutputFolder()
    Dim scratchFolder As String

    On Error Resume Next
    scratchFolder = JoinPath(OutputFolder(), "OBTApp_")
    Kill JoinPath(scratchFolder, "*")
    RmDir scratchFolder
    Kill JoinPath(OutputFolder(), "shared", "*")
    RmDir JoinPath(OutputFolder(), "shared")
    Kill JoinPath(OutputFolder(), "*")
    On Error GoTo 0
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
