Attribute VB_Name = "TestLLExport"

Option Explicit

'@Folder("CustomTests")
'@ModuleDescription("Tests for LLExport class covering creation, row operations, dictionary synchronisation, sorting, filename generation, and export spec replication.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'Validates the LLExport class which manages the export specification table in
'the linelist builder. Tests cover instantiation, row lifecycle (add, insert,
'delete, remove), dictionary column synchronisation, sequential sort renaming,
'export-spec replication into destination workbooks, filename template
'resolution with literal and variable chunks, active/inactive status filtering,
'and threshold-based row removal.
'@depends LLExport, LLdictionary, Passwords, CustomTest, BetterArray, TestHelpersLite, DictionaryTestFixture, PasswordsTestFixture


Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const EXPORT_SHEET As String = "LLExportSpec"
Private Const SOURCE_SHEET As String = "LLExportSource"
Private Const DICT_SHEET As String = "LLExportDict"
Private Const VLIST_SHEET As String = "vlist1D-sheet1"
Private Const PASSWORD_SHEET As String = "LLExportPasswords"
Private Const DICT_LO_NAME As String = "Tab_Dictionary"

Private Assert As CustomTest
Private DictionarySheet As Worksheet
Private ExportSheet As Worksheet
Private VListSheet As Worksheet
Private Manager As LLExport
Private PasswordSheet As Worksheet
Private PasswordsSubject As Passwords

'@section Lifecycle
'===============================================================================

'@sub-title Initialise the test module and prepare shared fixtures
'@details
'Public, so the headless runner reaches it: the runner calls every lifecycle
'hook through Application.Run, which cannot see a Private Sub. The handler
'matters as much. An error escaping a lifecycle hook aborts the WHOLE module,
'so a fixture problem here would drop all 23 tests and the run would report no
'failures, because nothing ran.
'@ModuleInitialize
Public Sub ModuleInitialize()
    On Error GoTo Fail
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLExport"
    PrepareTestSheets
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "ModuleInitialize", Err.Number, Err.Description
End Sub

'@sub-title Tear down worksheets and print test results
'@details
'Every step before RestoreApp is wrapped, because RestoreApp has to run
'whatever happened above: the hooks here call BusyApp, which puts Excel in
'manual calculation, and the next module in the run would inherit that.
'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        ThisWorkbook.Names("choi_v1").Delete
        DeleteWorksheet EXPORT_SHEET
        DeleteWorksheet DICT_SHEET
        DeleteWorksheet VLIST_SHEET
        DeleteWorksheet PASSWORD_SHEET
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Reset test sheets and create fresh Manager and Passwords instances before each test
'@TestInitialize
Public Sub TestInitialize()
    On Error GoTo Fail
    BusyApp
    PrepareTestSheets
    Set Manager = LLExport.Create(ExportSheet)
    Set PasswordsSubject = Passwords.Create(PasswordSheet)
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInitialize", Err.Number, Err.Description
End Sub

'@sub-title Flush current test results and release object references after each test
'@TestCleanup
Public Sub TestCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.FlushCurrentTest
        End If
    On Error GoTo 0

    Set Manager = Nothing
    Set PasswordsSubject = Nothing
End Sub

'@section Creation
'===============================================================================

'@sub-title Verify that Create initialises the Data property and reports the row count
'@details
'Arranges by creating a Manager via LLExport.Create with a single-row export
'table. Asserts that Data is not Nothing and that NumberOfExports agrees with
'the rows the table holds.
'@TestMethod("LLExport")
Public Sub TestLLExportCreateInitialisesData()
    CustomTestSetTitles Assert, "LLExport", "TestLLExportCreateInitialisesData"
    On Error GoTo Fail

    Assert.IsTrue (Not Manager.Data Is Nothing), "Expected Data to be initialised"
    Assert.AreEqual 1, Manager.NumberOfExports, "Should report single export row"
    Assert.AreEqual TableRowCount(), Manager.NumberOfExports, _
                    "NumberOfExports should agree with the rows the table holds"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLLExportCreateInitialisesData", Err.Number, Err.Description
End Sub

'@section ExportSpecs
'===============================================================================

'@sub-title Verify that ExportSpecs replicates the specification sheet into the destination workbook
'@details
'Arranges by reading the current number of exports, then creates a temporary
'workbook and calls ExportSpecs to copy the export sheet into it. Asserts that
'the destination carries a sheet of the same name, with the same number of data
'rows and the same file name template. Closes the temporary workbook without
'saving.
'@TestMethod("LLExport")
Public Sub TestExportSpecsCopiesSpecificationSheet()
    CustomTestSetTitles Assert, "LLExport", "TestExportSpecsCopiesSpecificationSheet"
    Dim exportBook As Workbook
    Dim exportedManager As LLExport
    Dim expectedTotal As Long
    Dim expectedTemplate As String

    On Error GoTo Fail

    expectedTotal = Manager.NumberOfExports
    expectedTemplate = Manager.ColumnValue(1, "file name")

    Set exportBook = NewWorkbook
    Manager.ExportSpecs exportBook, Hide:=xlSheetVisible

    Set exportedManager = LLExport.Create(exportBook.Worksheets(EXPORT_SHEET))

    Assert.AreEqual expectedTotal, exportedManager.NumberOfExports, _
                    "ExportSpecs should replicate every specification row into the destination workbook."
    Assert.AreEqual expectedTemplate, exportedManager.ColumnValue(1, "file name"), _
                    "ExportSpecs should replicate the file name template."

    exportBook.Close SaveChanges:=False
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportSpecsCopiesSpecificationSheet", Err.Number, Err.Description
    On Error Resume Next
        If Not exportBook Is Nothing Then exportBook.Close SaveChanges:=False
    On Error GoTo 0
End Sub

'@sub-title Verify that ImportSpecs writes "Export 1" over an imported bare "1"
'@details
'A setup written before this class filled the export number column holds the
'number alone, and ParseExportIdentifier answers 0 for it. The source carries
'two rows against the host's one, so the second row lands below the ListObject
'and has to be converted through the data block rather than through the table.
'@TestMethod("LLExport")
Public Sub TestImportSpecsTagsBareExportNumbers()
    CustomTestSetTitles Assert, "LLExport", "TestImportSpecsTagsBareExportNumbers"
    On Error GoTo Fail

    Dim sourceSheet As Worksheet
    Dim exportIdx As Long

    exportIdx = ColumnIndexOf("export number")
    Set sourceSheet = PrepareExportSource(1, 2)

    Manager.ImportSpecs sourceSheet, 1, 1

    Assert.AreEqual "Export 1", CStr(ExportSheet.Cells(2, exportIdx).Value), _
                    "An imported bare 1 should be written as Export 1"
    Assert.AreEqual "Export 2", CStr(ExportSheet.Cells(3, exportIdx).Value), _
                    "A row that landed below the table should be tagged as well"

Cleanup:
    DeleteWorksheet SOURCE_SHEET
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportSpecsTagsBareExportNumbers", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@sub-title Verify that ImportSpecs keeps the number an imported row carries
'@details
'The import tags a number that has no word in front of it. It gives no row a
'different number: renumbering belongs to EnsureExportIdentifiers, which runs
'when a row is added or deleted.
'@TestMethod("LLExport")
Public Sub TestImportSpecsKeepsTaggedExportNumbers()
    CustomTestSetTitles Assert, "LLExport", "TestImportSpecsKeepsTaggedExportNumbers"
    On Error GoTo Fail

    Dim sourceSheet As Worksheet
    Dim exportIdx As Long

    exportIdx = ColumnIndexOf("export number")
    Set sourceSheet = PrepareExportSource("Export 3", 7)

    Manager.ImportSpecs sourceSheet, 1, 1

    Assert.AreEqual "Export 3", CStr(ExportSheet.Cells(2, exportIdx).Value), _
                    "A row that already carries its number should keep it"
    Assert.AreEqual "Export 7", CStr(ExportSheet.Cells(3, exportIdx).Value), _
                    "A bare 7 should become Export 7 rather than the next free number"

Cleanup:
    DeleteWorksheet SOURCE_SHEET
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportSpecsKeepsTaggedExportNumbers", Err.Number, Err.Description
    Resume Cleanup
End Sub

'@section AddRows
'===============================================================================

'@sub-title Verify that AddRows appends a row with correct defaults and sequential identifiers
'@details
'Calls AddRows once on a single-row export table. Asserts that NumberOfExports
'increases to two, the new row defaults "include personal identifiers" to "no",
'the existing row is normalised to "export 1", and the new row receives
'"export 2" as its sequential identifier.
'@TestMethod("LLExport")
Public Sub TestAddRowsAppliesDefaults()
    CustomTestSetTitles Assert, "LLExport", "TestAddRowsAppliesDefaults"
    On Error GoTo Fail

    Dim exportIdx As Long

    Manager.AddRows
    Assert.AreEqual 2, Manager.NumberOfExports, "Row count should grow by one"
    Assert.AreEqual "no", Manager.ColumnValue(2, "include personal identifiers"), _
                     "Include personal identifiers should default to 'no'"

    exportIdx = ColumnIndexOf("export number")
    Assert.AreEqual "export 1", ExportSheet.ListObjects(1).DataBodyRange.Cells(1, exportIdx).Value, _
                     "Existing export should be normalised to prefixed identifier"
    Assert.AreEqual "export 2", ExportSheet.ListObjects(1).DataBodyRange.Cells(2, exportIdx).Value, _
                     "Newly added export should receive the next identifier"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddRowsAppliesDefaults", Err.Number, Err.Description
End Sub

'@sub-title Verify that AddRows synchronises Export columns in the dictionary
'@details
'Arranges by creating a dictionary and removing surplus Export columns so only
'Export 1 remains. Calls AddRows with the dictionary parameter. Asserts that
'NumberOfExports increases, the table row count follows, the dictionary gains
'an "Export 2" column, and the dictionary export counter matches the number of
'export rows.
'@TestMethod("LLExport")
Public Sub TestAddRowsSynchronisesDictionaryColumns()
    CustomTestSetTitles Assert, "LLExport", "TestAddRowsSynchronisesDictionaryColumns"
    On Error GoTo Fail

    Dim dict As LLdictionary
    Dim idx As Long
    Dim startTotal As Long

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)

    For idx = 5 To 2 Step -1
        dict.RemoveColumn "Export " & CStr(idx)
    Next idx

    startTotal = Manager.NumberOfExports
    Assert.AreEqual startTotal, TableRowCount(), "NumberOfExports should match the table rows before adding"

    Manager.AddRows dict:=dict

    Assert.AreEqual startTotal + 1, TableRowCount(), "The table should carry one more row after adding"
    Assert.AreEqual startTotal + 1, Manager.NumberOfExports, "NumberOfExports should report the updated total after adding rows"
    Assert.IsTrue dict.ColumnExists("Export 2"), "Dictionary should expose Export 2 column after row addition"
    Assert.AreEqual Manager.NumberOfExports, CLng(dict.TotalNumberOfExports), _
                    "Dictionary export counter should hold one column per export row"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddRowsSynchronisesDictionaryColumns", Err.Number, Err.Description
End Sub

'@sub-title Verify that AddRows works after ResetCaches when no dictionary is supplied
'@details
'Calls ResetCaches to clear any cached state, then calls AddRows without a
'dictionary parameter. Asserts that NumberOfExports increases to two and stays
'aligned with the table row count.
'@TestMethod("LLExport")
Public Sub TestAddRowsWithoutDictionaryAfterReset()
    CustomTestSetTitles Assert, "LLExport", "TestAddRowsWithoutDictionaryAfterReset"
    On Error GoTo Fail

    Manager.ResetCaches
    Manager.AddRows

    Assert.AreEqual 2, Manager.NumberOfExports, "Row count should grow when dictionary is not supplied"
    Assert.AreEqual 2, TableRowCount(), "The table should carry two rows even without a dictionary"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddRowsWithoutDictionaryAfterReset", Err.Number, Err.Description
End Sub

'@section InsertRows
'===============================================================================

'@sub-title Verify that InsertRows applies defaults and synchronises dictionary columns
'@details
'Arranges by creating a dictionary, seeding Export 1 and Export 2 data, then
'calling AddRows to reach two rows. Inserts a row at row 2 using a selection
'range. Asserts that the export count increases to three, inserted rows default
'"include personal identifiers" to "no", existing dictionary data is preserved,
'the new Export 3 column starts blank, and the table row count mirrors the total.
'@TestMethod("LLExport")
Public Sub TestInsertRowsAppliesDefaultsAndSyncsDictionary()
    CustomTestSetTitles Assert, "LLExport", "TestInsertRowsAppliesDefaultsAndSyncsDictionary"
    On Error GoTo Fail

    Dim dict As LLdictionary
    Dim selectionRange As Range
    Dim dictSheet As Worksheet
    Dim dictLo As ListObject

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)
    Set dictSheet = DictionarySheet
    Set dictLo = dictSheet.ListObjects(DICT_LO_NAME)

    EnsureExportColumn dictLo, "Export 1"
    EnsureExportColumn dictLo, "Export 2"

    dictLo.ListColumns("Export 1").DataBodyRange.Cells(1, 1).Value = "Alpha"
    dictLo.ListColumns("Export 2").DataBodyRange.Cells(1, 1).Value = "Beta"

    Manager.AddRows dict:=dict

    Set selectionRange = ExportSheet.ListObjects(1).ListRows(2).Range

    Manager.InsertRows selectionRange, dict:=dict

    Assert.AreEqual 3, Manager.NumberOfExports, "InsertRows should increase export count"
    Assert.AreEqual "no", Manager.ColumnValue(2, "include personal identifiers"), _
                     "Inserted export rows should default include personal identifiers to 'no'"
    Assert.AreEqual "Alpha", CStr(dictLo.ListColumns("Export 1").DataBodyRange.Cells(1, 1).Value), _
                     "Existing export data should remain untouched"
    Assert.AreEqual "Beta", CStr(dictLo.ListColumns("Export 2").DataBodyRange.Cells(1, 1).Value), _
                     "Existing export columns should not be shifted"
    Assert.IsTrue dictLo.ListColumns.Count >= 3, "Dictionary should expose the new export column"
    Assert.AreEqual vbNullString, CStr(dictLo.ListColumns("Export 3").DataBodyRange.Cells(1, 1).Value), _
                     "New export column should start blank"
    Assert.AreEqual Manager.NumberOfExports, TableRowCount(), _
                     "NumberOfExports should mirror the table row count"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsAppliesDefaultsAndSyncsDictionary", Err.Number, Err.Description
End Sub

'@sub-title Verify that InsertRows silently ignores a selection outside the exports table
'@details
'Arranges by recording the initial export count and selecting a cell far outside
'the table body (Z100). Calls InsertRows with the invalid selection. Asserts
'that NumberOfExports and the table row count remain unchanged.
'@TestMethod("LLExport")
Public Sub TestInsertRowsIgnoresSelectionOutsideTable()
    CustomTestSetTitles Assert, "LLExport", "InsertRows ignores selection outside the exports table"
    On Error GoTo Fail

    Dim initialCount As Long
    Dim invalidSelection As Range

    initialCount = Manager.NumberOfExports
    Set invalidSelection = ExportSheet.Range("Z100")

    Manager.InsertRows invalidSelection

    Assert.AreEqual initialCount, Manager.NumberOfExports, _
                     "InsertRows should leave the export count unchanged when the selection is invalid"
    Assert.AreEqual initialCount, TableRowCount(), _
                     "The table row count should remain aligned when no insertion occurs"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsIgnoresSelectionOutsideTable", Err.Number, Err.Description
End Sub

'@section DeleteRows
'===============================================================================

'@sub-title Verify that DeleteRows removes selected rows and prunes dictionary columns
'@details
'Arranges by adding a row via AddRows with a dictionary so there are two rows.
'Selects the second row and calls DeleteRows. Asserts that NumberOfExports
'returns one, the table row count follows, and the dictionary no longer contains
'the "Export 2" column.
'@TestMethod("LLExport")
Public Sub TestDeleteRowsShrinksExportsAndDictionary()
    CustomTestSetTitles Assert, "LLExport", "TestDeleteRowsShrinksExportsAndDictionary"
    On Error GoTo Fail

    Dim dict As LLdictionary
    Dim selectionRange As Range

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)
    Manager.AddRows dict:=dict

    Set selectionRange = ExportSheet.ListObjects(1).ListRows(2).Range
    Manager.DeleteRows selectionRange, dict:=dict

    Assert.AreEqual 1, Manager.NumberOfExports, "DeleteRows should remove export entries"
    Assert.AreEqual 1, TableRowCount(), "The table should carry the remaining row only"
    Assert.IsFalse dict.ColumnExists("Export 2"), "Dictionary should drop columns for deleted exports"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDeleteRowsShrinksExportsAndDictionary", Err.Number, Err.Description
End Sub

'@section Sort
'===============================================================================

'@sub-title Verify that Sort renumbers export identifiers sequentially and reorders dictionary data
'@details
'Arranges by adding two rows to reach three exports, seeding dictionary Export
'columns with known values, and scrambling the "export number" column to 3-1-2
'order. Calls Sort with the dictionary. Asserts that after sorting the export
'number column reads export 1, export 2, export 3 in sequential order.
'@TestMethod("LLExport")
Public Sub TestSortRenamesExportsSequentially()
    CustomTestSetTitles Assert, "LLExport", "TestSortRenamesExportsSequentially"
    On Error GoTo Fail

    Dim dict As LLdictionary
    Dim dictLo As ListObject
    Dim exportNumberIndex As Long
    Dim lo As ListObject

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)
    Set dictLo = DictionarySheet.ListObjects(DICT_LO_NAME)

    EnsureExportColumn dictLo, "Export 1"
    EnsureExportColumn dictLo, "Export 2"
    EnsureExportColumn dictLo, "Export 3"

    Manager.AddRows dict:=dict
    Manager.AddRows dict:=dict

    Set lo = ExportSheet.ListObjects(1)
    exportNumberIndex = lo.ListColumns("export number").Index

    dictLo.ListColumns("Export 1").DataBodyRange.Cells(1, 1).Value = "One"
    dictLo.ListColumns("Export 2").DataBodyRange.Cells(1, 1).Value = "Two"
    dictLo.ListColumns("Export 3").DataBodyRange.Cells(1, 1).Value = "Three"

    lo.DataBodyRange.Cells(1, exportNumberIndex).Value = "export 3"
    lo.DataBodyRange.Cells(2, exportNumberIndex).Value = "export 1"
    lo.DataBodyRange.Cells(3, exportNumberIndex).Value = "export 2"

    Manager.Sort dict

    Assert.AreEqual "export 1", CStr(lo.DataBodyRange.Cells(1, exportNumberIndex).Value), _
                     "First row should be renamed sequentially after sort"
    Assert.AreEqual "export 2", CStr(lo.DataBodyRange.Cells(2, exportNumberIndex).Value), _
                     "Second row should be renamed sequentially after sort"
    Assert.AreEqual "export 3", CStr(lo.DataBodyRange.Cells(3, exportNumberIndex).Value), _
                     "Third row should be renamed sequentially after sort"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSortRenamesExportsSequentially", Err.Number, Err.Description
End Sub

'@section SyncDictionary
'===============================================================================

'@sub-title Verify that SyncDictionaryExports ignores non-export-prefixed dictionary columns
'@details
'Arranges by creating a dictionary, removing "Export 3", and adding a custom
'column "mainlab_3_backup". Calls AddRows with the dictionary. Asserts that the
'custom column is preserved, the table row count reflects actual export rows,
'and the dictionary counter ignores non-export-prefixed columns.
'@TestMethod("LLExport")
Public Sub TestSyncDictionaryIgnoresNonExportPrefixedColumns()
    CustomTestSetTitles Assert, "LLExport", "TestSyncDictionaryIgnoresNonExportPrefixedColumns"
    On Error GoTo Fail

    Dim dict As LLdictionary

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)

    dict.RemoveColumn "Export 3"
    dict.AddColumn "mainlab_3_backup"

    Manager.AddRows dict:=dict

    Assert.IsTrue dict.ColumnExists("mainlab_3_backup"), "Custom column should remain untouched"
    Assert.AreEqual 2, TableRowCount(), "The table should carry the actual export rows"
    Assert.AreEqual 2, CLng(dict.TotalNumberOfExports), "Dictionary counter should ignore non-export-prefixed columns"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSyncDictionaryIgnoresNonExportPrefixedColumns", Err.Number, Err.Description
End Sub

'@sub-title Verify that SyncDictionaryExports works with an explicit dictionary and with the cached one
'@details
'Arranges by adding a row with a dictionary, then removing "Export 2" and
'calling SyncDictionaryExports with the dictionary explicitly. Asserts Export 2
'is restored. Removes Export 2 again and calls SyncDictionaryExports without a
'parameter, verifying it reuses the cached dictionary reference.
'@TestMethod("LLExport")
Public Sub TestPublicSyncDictionaryExportsWithAndWithoutParameter()
    CustomTestSetTitles Assert, "LLExport", "TestPublicSyncDictionaryExportsWithAndWithoutParameter"
    On Error GoTo Fail

    Dim dict As LLdictionary

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)

    Manager.AddRows dict:=dict

    dict.RemoveColumn "Export 2"
    Manager.SyncDictionaryExports dict
    Assert.IsTrue dict.ColumnExists("Export 2"), "SyncDictionaryExports should restore export columns when dictionary provided"

    dict.RemoveColumn "Export 2"
    Manager.SyncDictionaryExports
    Assert.IsTrue dict.ColumnExists("Export 2"), "SyncDictionaryExports should reuse cached dictionary when none supplied"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestPublicSyncDictionaryExportsWithAndWithoutParameter", Err.Number, Err.Description
End Sub

'@section RemoveRows
'===============================================================================

'@sub-title Verify that RemoveRows deletes empty export rows
'@details
'Arranges by adding a row and then clearing the content of the second row.
'Calls RemoveRows. Asserts that only one export remains and its identifier is
'preserved as "export 1".
'@TestMethod("LLExport")
Public Sub TestLLExportRemoveRowsDeletesEmpty()
    CustomTestSetTitles Assert, "LLExport", "TestLLExportRemoveRowsDeletesEmpty"
    On Error GoTo Fail

    Manager.AddRows
    ExportSheet.ListObjects(1).DataBodyRange.Rows(2).ClearContents
    Manager.RemoveRows

    Assert.AreEqual 1, Manager.NumberOfExports, "Removing rows should trim empty rows"
    Assert.AreEqual "export 1", ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("export number")).Value, _
                     "Remaining export should keep its prefixed identifier"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLLExportRemoveRowsDeletesEmpty", Err.Number, Err.Description
End Sub

'@sub-title Verify that RemoveRows prunes dictionary columns for removed exports
'@details
'Arranges by creating a dictionary with reduced Export columns, adding two
'rows to reach three exports, then calling RemoveRows with the dictionary.
'Asserts that the table row count, NumberOfExports, and dictionary columns all
'reflect one remaining export, and the dictionary counter matches.
'@TestMethod("LLExport")
Public Sub TestRemoveRowsPrunesDictionaryColumns()
    CustomTestSetTitles Assert, "LLExport", "TestRemoveRowsPrunesDictionaryColumns"
    On Error GoTo Fail

    Dim dict As LLdictionary
    Dim idx As Long

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)

    For idx = 5 To 4 Step -1
        dict.RemoveColumn "Export " & CStr(idx)
    Next idx

    Manager.AddRows dict:=dict
    Manager.AddRows dict:=dict
    Assert.AreEqual 3, TableRowCount(), "Expected the table to carry the added rows"

    Manager.RemoveRows dict:=dict

    Assert.AreEqual 1, TableRowCount(), "The table should carry the trimmed rows"
    Assert.AreEqual 1, Manager.NumberOfExports, "NumberOfExports should reflect the trimmed export rows"
    Assert.IsFalse dict.ColumnExists("Export 2"), "Export 2 column should be removed when only one export remains"
    Assert.IsFalse dict.ColumnExists("Export 3"), "Export 3 column should be removed when not present"
    Assert.AreEqual 1, CLng(dict.TotalNumberOfExports), "Dictionary export count should match remaining exports"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveRowsPrunesDictionaryColumns", Err.Number, Err.Description
End Sub

'@sub-title Verify that RemoveRows retains dictionary columns for identifiers still present
'@details
'Arranges by adding three rows so four exports exist, manually deleting
'row 3 from the ListObject, then calling RemoveRows. Asserts that the row count
'drops, the table row count aligns, Export 4 is gone because its identifier was
'removed, and Export 3 is also gone.
'@TestMethod("LLExport")
Public Sub TestRemoveRowsKeepsColumnsForRemainingIdentifiers()
    CustomTestSetTitles Assert, "LLExport", "TestRemoveRowsKeepsColumnsForRemainingIdentifiers"
    On Error GoTo Fail

    Dim dict As LLdictionary
    Dim exportIdx As Long

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)

    dict.RemoveColumn "Export 5"

    Manager.AddRows dict:=dict
    Manager.AddRows dict:=dict
    Manager.AddRows dict:=dict

    exportIdx = ColumnIndexOf("export number")
    Assert.AreEqual "export 4", ExportSheet.ListObjects(1).DataBodyRange.Cells(4, exportIdx).Value, _
                     "Fourth export should receive identifier export 4"

    ExportSheet.ListObjects(1).ListRows(3).Delete

    Manager.RemoveRows

    Assert.AreEqual 1, Manager.NumberOfExports, "Row count should drop after deleting one export"
    Assert.AreEqual 1, TableRowCount(), "The table row count should align with the current export count"
    Assert.IsFalse dict.ColumnExists("Export 4"), "Export column 4 should persist when its identifier still exists"
    Assert.IsFalse dict.ColumnExists("Export 3"), "Export column 3 should be removed when identifier is missing"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveRowsKeepsColumnsForRemainingIdentifiers", Err.Number, Err.Description
End Sub

'@sub-title Verify that RemoveRows honours the rowCount threshold parameter
'@details
'Arranges by adding a row to reach two exports, then calls RemoveRows with
'threshold zero which should keep all rows with data. Asserts two rows remain.
'Calls RemoveRows again with threshold two which trims populated rows at or
'below the threshold. Asserts one row remains.
'@TestMethod("LLExport")
Public Sub TestRemoveRowsHonoursThreshold()
    CustomTestSetTitles Assert, "LLExport", "TestRemoveRowsHonoursThreshold"
    On Error GoTo Fail

    Manager.AddRows
    Manager.RemoveRows rowCount:=0
    Assert.AreEqual 2, Manager.NumberOfExports, "Rows with data should remain when threshold is zero"

    Manager.RemoveRows rowCount:=2
    Assert.AreEqual 1, Manager.NumberOfExports, "Rows at or below the threshold should be trimmed"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveRowsHonoursThreshold", Err.Number, Err.Description
End Sub

'@section ExportFileName
'===============================================================================

'@sub-title Verify that ExportFileName resolves variable and literal chunks from the template
'@details
'Calls ExportFileName on row 1 whose template contains a named-range variable
'(choi_v1) and a quoted literal. Asserts the result includes the resolved value
'concatenated with the literal suffix, includes the version marker, and that no
'checking entries are produced for literal-only chunks.
'@TestMethod("LLExport")
Public Sub TestExportFileNameBuildsFromTemplate()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameBuildsFromTemplate"
    On Error GoTo Fail

    Dim fileName As String

    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)
    Assert.IsTrue InStr(1, fileName, "custom_value__literal_suffix", vbTextCompare) > 0, _
                  "Filename should include resolved variable and literal chunks"
    Assert.IsTrue fileName Like "*__vd0099-1234__*", "Version suffix should be appended - Actual finename: " & fileName
    Assert.IsFalse Manager.HasCheckings, "Default template should not trigger checkings for literal chunks"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportFileNameBuildsFromTemplate", Err.Number, Err.Description
End Sub

'@sub-title Verify that an empty variable puts its own name in the file name
'@details
'Writes the default template, empties the cell the choi_v1 named range points
'at, and calls ExportFileName. Asserts the name carries "choi_v1", carries no
'"chunk", and that the run filed a checking entry for the blank value.
'
'"chunk" is the placeholder SanitizeChunk answers for a piece that comes to
'nothing. It used to be applied to the variable value as well, so a blank
'entry answered "chunk", the caller read it as a resolved value, and the word
'landed in the delivered file name once per blank entry. Restores the cell.
'@TestMethod("LLExport")
Public Sub TestExportFileNameNamesAnEmptyVariable()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameNamesAnEmptyVariable"

    Dim fileName As String

    On Error GoTo Fail

    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("file name")).Value = "choi_v1 + ""literal suffix"""
    VListSheet.Range("A1").ClearContents

    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)

    Assert.IsTrue InStr(1, fileName, "choi_v1", vbTextCompare) > 0, _
                  "An empty variable should put its own name in the file name - Actual: " & fileName
    Assert.IsTrue InStr(1, fileName, "chunk", vbTextCompare) = 0, _
                  "An empty variable should leave the chunk placeholder out of the file name - Actual: " & fileName
    Assert.IsTrue Manager.HasCheckings, _
                  "An empty variable should file a checking entry"

    VListSheet.Range("A1").Value = "custom value"
    Exit Sub

Fail:
    VListSheet.Range("A1").Value = "custom value"
    CustomTestLogFailure Assert, "TestExportFileNameNamesAnEmptyVariable", Err.Number, Err.Description
End Sub

'@sub-title Verify that ExportFileName preserves single-quoted literal chunks
'@details
'Overwrites the file name template to use single-quoted literals. Calls
'ExportFileName and asserts the sanitised single-quoted literal appears in the
'result and no checking entries are logged.
'@TestMethod("LLExport")
Public Sub TestExportFileNameHandlesSingleQuotedLiteral()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameHandlesSingleQuotedLiteral"
    On Error GoTo Fail

    Dim fileName As String

    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("file name")).Value = "choi_v1 + 'single literal'"

    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)

    Assert.IsTrue InStr(1, fileName, "custom_value__single_literal", vbTextCompare) > 0, _
                  "Single-quoted literal chunks should be preserved"
    Assert.IsFalse Manager.HasCheckings, "Single-quoted literal chunks should not trigger checkings"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportFileNameHandlesSingleQuotedLiteral", Err.Number, Err.Description
End Sub

'@sub-title Verify that ExportFileName concatenates all-literal templates correctly
'@details
'Overwrites the file name template so it contains only double-quoted literal
'chunks with no variable references. Calls ExportFileName and asserts that the
'sanitised literals are concatenated with underscores and no checking entries
'are produced.
'@TestMethod("LLExport")
Public Sub TestExportFileNameWithOnlyLiteralChunks()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameWithOnlyLiteralChunks"
    On Error GoTo Fail

    Dim fileName As String

    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("file name")).Value = """static chunk"" + ""second part"""

    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)

    Assert.IsTrue InStr(1, fileName, "static_chunk__second_part", vbTextCompare) > 0, _
                  "All-literal templates should concatenate sanitized literals"
    Assert.IsFalse Manager.HasCheckings, "All-literal templates should not trigger checkings"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportFileNameWithOnlyLiteralChunks", Err.Number, Err.Description
End Sub

'@sub-title Verify that every character in the sanitize list is replaced
'@details
'Writes a literal template holding the whole set of characters the class calls
'invalid. Asserts that none of them survives into the file name. Against the
'earlier sanitiser only the last entry in the list, a space, was replaced, so
'a template chunk carrying a slash or a colon reached the SaveAs path.
'@TestMethod("LLExport")
Public Sub TestExportFileNameReplacesEveryInvalidCharacter()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameReplacesEveryInvalidCharacter"
    On Error GoTo Fail

    Dim fileName As String
    Dim invalidChars As Variant
    Dim idx As Long
    Dim badChar As String

    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("file name")).Value = _
        Chr$(34) & "a<b>c:d\e/f|g?h*i&j.k l" & Chr$(34)

    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)

    invalidChars = Array("<", ">", ":", "\", "/", "|", "?", "*", "&", " ")

    For idx = LBound(invalidChars) To UBound(invalidChars)
        badChar = CStr(invalidChars(idx))
        Assert.AreEqual 0, InStr(1, fileName, badChar, vbBinaryCompare), _
                        "Character " & badChar & " should be replaced - Actual filename: " & fileName
    Next idx

    Assert.IsTrue InStr(1, fileName, "a_b_c_d_e_f_g_h_i_j_k_l", vbTextCompare) > 0, _
                  "Every invalid character should become an underscore - Actual filename: " & fileName
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportFileNameReplacesEveryInvalidCharacter", Err.Number, Err.Description
End Sub

'@sub-title Verify that ExportFileName logs a checking entry for inactive exports
'@details
'Arranges by adding a row and setting its status to "inactive". Calls
'ExportFileName for that row. Asserts that HasCheckings is True and a non-empty
'filename is still returned despite the inactive status.
'@TestMethod("LLExport")
Public Sub TestExportFileNameLogsWhenInactive()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameLogsWhenInactive"
    On Error GoTo Fail

    Dim nameValue As String

    Manager.AddRows
    ExportSheet.ListObjects(1).DataBodyRange.Cells(2, ColumnIndexOf("status")).Value = "inactive"

    nameValue = Manager.ExportFileName(2, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)
    Assert.IsTrue Manager.HasCheckings, "Inactive export should log information"
    Assert.IsTrue LenB(nameValue) > 0, "Should still return a filename"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportFileNameLogsWhenInactive", Err.Number, Err.Description
End Sub

'@sub-title Verify that exportAll flag overrides the scope portion of the filename
'@details
'Calls ExportFileName with exportAll set to True. Asserts that the resulting
'filename contains "export_all" indicating the scope override was applied.
'@TestMethod("LLExport")
Public Sub TestExportAllOverridesScope()
    CustomTestSetTitles Assert, "LLExport", "TestExportAllOverridesScope"
    On Error GoTo Fail

    Dim nameValue As String

    nameValue = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject, exportAll:=True)
    Assert.IsTrue InStr(1, nameValue, "export_all", vbTextCompare) > 0, "ExportAll should override scope - fileName : " & nameValue
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportAllOverridesScope", Err.Number, Err.Description
End Sub

'@sub-title Verify that ExportFileName logs a checking entry for unresolvable chunks
'@details
'Overwrites the file name template with a token that does not correspond to any
'named range or known variable. Calls ExportFileName and asserts that
'HasCheckings is True and the fallback filename still includes the sanitised
'unresolved chunk.
'@TestMethod("LLExport")
Public Sub TestExportFileNameLogsMissingChunk()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameLogsMissingChunk"
    On Error GoTo Fail

    Dim fileName As String

    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("file name")).Value = "unknown_chunk"

    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)

    Assert.IsTrue Manager.HasCheckings, "Missing chunk should produce checking entries"
    Assert.IsTrue InStr(1, fileName, "unknown_chunk", vbTextCompare) > 0, _
                  "Fallback filename should include sanitized chunk"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportFileNameLogsMissingChunk", Err.Number, Err.Description
End Sub

'@section Status
'===============================================================================

'@sub-title Verify that IsActive correctly reflects the row status column
'@details
'Asserts that the default row (status "active") returns True from IsActive.
'Overwrites the status cell to "inactive" and asserts IsActive returns False.
'The class keeps a copy of the specification table, so a test that writes to
'the sheet itself calls ResetCaches before reading again.
'@TestMethod("LLExport")
Public Sub TestIsActiveReflectsStatus()
    CustomTestSetTitles Assert, "LLExport", "TestIsActiveReflectsStatus"
    On Error GoTo Fail

    Assert.IsTrue Manager.IsActive(1), "Row with active status should be active"

    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("status")).Value = "inactive"
    Manager.ResetCaches

    Assert.IsFalse Manager.IsActive(1), "Row with inactive status should report false"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestIsActiveReflectsStatus", Err.Number, Err.Description
End Sub

'@sub-title Verify that ActiveExportNumbers returns only active row indices
'@details
'Arranges by adding two rows to reach three exports, then sets the second row
'to "inactive" and the third to "active". Calls ActiveExportNumbers and asserts
'the returned BetterArray contains exactly two entries corresponding to rows 1
'and 3.
'@TestMethod("LLExport")
Public Sub TestActiveExportNumbersReturnsActiveRows()
    CustomTestSetTitles Assert, "LLExport", "TestActiveExportNumbersReturnsActiveRows"
    On Error GoTo Fail

    Dim statusCol As Long
    Dim active As BetterArray
    Dim startIndex As Long

    Manager.AddRows
    Manager.AddRows

    statusCol = ColumnIndexOf("status")

    ExportSheet.ListObjects(1).DataBodyRange.Cells(2, statusCol).Value = "inactive"
    ExportSheet.ListObjects(1).DataBodyRange.Cells(3, statusCol).Value = "active"

    Set active = Manager.ActiveExportNumbers

    Assert.AreEqual 2, active.Length, "Expected two active exports"
    startIndex = active.LowerBound
    Assert.AreEqual 1, CLng(active.Item(startIndex)), "First active export should be row 1"
    Assert.AreEqual 3, CLng(active.Item(startIndex + 1)), "Second active export should be row 3"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestActiveExportNumbersReturnsActiveRows", Err.Number, Err.Description
End Sub

'@section Helpers
'===============================================================================

'@sub-title Build all fixture worksheets needed by the test suite
'@details
'Creates or resets the dictionary, export, variable list, and password sheets.
'Populates the dictionary with a ListObject, seeds the export table via
'PrepareExportTable, writes a named range for choi_v1 resolution, and delegates
'password fixture creation to PasswordsTestFixture.
Private Sub PrepareTestSheets()

    Set DictionarySheet = EnsureWorksheet(DICT_SHEET)
    PrepareDictionaryFixture DICT_SHEET
    With DictionarySheet
        .ListObjects.Add(xlSrcRange, .Range("A1:AD78"), , xlYes).Name = DICT_LO_NAME
    End With

    Set ExportSheet = EnsureWorksheet(EXPORT_SHEET)
    PrepareExportTable ExportSheet

    Set VListSheet = EnsureWorksheet(VLIST_SHEET)

    VListSheet.Range("A1").Value = "custom value"

    On Error Resume Next
    ThisWorkbook.Names("choi_v1").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:="choi_v1", RefersTo:=VListSheet.Range("A1")

    PasswordsTestFixture.PreparePasswordsFixture PASSWORD_SHEET
    Set PasswordSheet = ThisWorkbook.Worksheets(PASSWORD_SHEET)

End Sub

'@sub-title Ensure that a named Export column exists in the dictionary ListObject
Private Sub EnsureExportColumn(ByVal dictLo As ListObject, ByVal columnName As String)
    On Error Resume Next
        dictLo.ListColumns(columnName).Name = columnName
    On Error GoTo 0
    If ColumnExists(dictLo, columnName) Then Exit Sub

    dictLo.ListColumns.Add.Name = columnName
End Sub

'@sub-title Check whether a ListObject contains a column with the given name
Private Function ColumnExists(ByVal dictLo As ListObject, ByVal columnName As String) As Boolean
    On Error Resume Next
        ColumnExists = Not dictLo.ListColumns(columnName) Is Nothing
    On Error GoTo 0
End Function

'@sub-title Populate the export sheet with a header row, data row, and ListObject
Private Sub PrepareExportTable(ByVal targetSheet As Worksheet)
    Dim headers As Variant
    Dim dataRow As Variant
    headers = Array("export number", "status", "label button", _
                    "file format", "file name", "password", _
                    "include personal identifiers", "include p-codes", _
                    "header format", "export metadata sheets", _
                    "export analyses sheets")

    dataRow = Array(1, "active", "Label", "xlsx", "choi_v1 + ""literal suffix""", "pwd", _
                    "", "yes", "default", "no", "no")

    targetSheet.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    targetSheet.Range("A2").Resize(1, UBound(dataRow) + 1).Value = dataRow
    targetSheet.ListObjects.Add SourceType:=xlSrcRange, _
            Source:=targetSheet.Range("A1").Resize(2, UBound(headers) + 1), XlListObjectHasHeaders:=xlYes
End Sub

'@sub-title Build a source export sheet ImportSpecs can be pointed at
'@details
'Two rows, so one of them lands below the single-row ListObject of the host.
'The two export numbers are passed in, which is how one test writes them bare
'and another writes one of them already carrying its word.
Private Function PrepareExportSource(ByVal firstNumber As Variant, _
                                     ByVal secondNumber As Variant) As Worksheet
    Dim sh As Worksheet
    Dim headers As Variant

    headers = Array("export number", "status", "label button", _
                    "file format", "file name", "password", _
                    "include personal identifiers", "include p-codes", _
                    "header format", "export metadata sheets", _
                    "export analyses sheets")

    Set sh = EnsureWorksheet(SOURCE_SHEET)
    sh.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    sh.Range("A2").Resize(1, UBound(headers) + 1).Value = _
        Array(firstNumber, "active", "First", "xlsx", """one""", "pwd", _
              "no", "yes", "default", "no", "no")
    sh.Range("A3").Resize(1, UBound(headers) + 1).Value = _
        Array(secondNumber, "active", "Second", "xlsx", """two""", "pwd", _
              "no", "yes", "default", "no", "no")

    Set PrepareExportSource = sh
End Function

'@sub-title Return the 1-based column index for a header in the export ListObject
Private Function ColumnIndexOf(ByVal headerName As String) As Long
    ColumnIndexOf = ExportSheet.ListObjects(1).ListColumns(headerName).Index
End Function

'@sub-title Count the data rows the export table holds on the worksheet
'@details
'Read straight off the ListObject, so it answers what the sheet holds rather
'than what the class believes.
Private Function TableRowCount() As Long
    Dim lo As ListObject

    Set lo = ExportSheet.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function

    TableRowCount = lo.DataBodyRange.Rows.Count
End Function
