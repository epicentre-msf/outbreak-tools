Attribute VB_Name = "TestLLExport"

Option Explicit

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const EXPORT_SHEET As String = "LLExportSpec"
Private Const DICT_SHEET As String = "LLExportDict"
Private Const VLIST_SHEET As String = "vlist1D-sheet1"
Private Const PASSWORD_SHEET As String = "LLExportPasswords"
Private Const EXPORT_TOTAL_NAME As String = "__ll_exports_total__"

Private Assert As ICustomTest
Private DictionarySheet As Worksheet
Private ExportSheet As Worksheet
Private VListSheet As Worksheet
Private Manager As ILLExport
Private PasswordSheet As Worksheet
Private PasswordsSubject As IPasswords

'@ModuleInitialize
Public Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLExport"
    PrepareTestSheets
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
    ThisWorkbook.Names("choi_v1").Delete
    DeleteWorksheet EXPORT_SHEET
    DeleteWorksheet DICT_SHEET
    DeleteWorksheet VLIST_SHEET
    DeleteWorksheet PASSWORD_SHEET
    On Error GoTo 0
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    PrepareTestSheets
    Set Manager = LLExport.Create(ExportSheet)
    Set PasswordsSubject = Passwords.Create(PasswordSheet)
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.FlushCurrentTest
    End If
    Set Manager = Nothing
    Set PasswordsSubject = Nothing
End Sub

'@TestMethod("LLExport")
Public Sub TestCreateInitialisesData()
    CustomTestSetTitles Assert, "LLExport", "TestCreateInitialisesData"
    Assert.IsTrue (Not Manager.Data Is Nothing), "Expected Data to be initialised"
    Assert.AreEqual 1, Manager.NumberOfExports, "Should report single export row"
    Assert.AreEqual 1, StoredExportTotal(), "Hidden export counter should match initial export count"
End Sub

'@TestMethod("LLExport")
Public Sub TestAddRowsAppliesDefaults()
    CustomTestSetTitles Assert, "LLExport", "TestAddRowsAppliesDefaults"
    Manager.AddRows
    Assert.AreEqual 2, Manager.NumberOfExports, "Row count should grow by one"
    Assert.AreEqual "no", Manager.ColumnValue(2, "include personal identifiers"), _
                     "Include personal identifiers should default to 'no'"
    Dim exportIdx As Long
    exportIdx = ColumnIndexOf("export number")
    Assert.AreEqual "export 1", ExportSheet.ListObjects(1).DataBodyRange.Cells(1, exportIdx).Value, _
                     "Existing export should be normalised to prefixed identifier"
    Assert.AreEqual "export 2", ExportSheet.ListObjects(1).DataBodyRange.Cells(2, exportIdx).Value, _
                     "Newly added export should receive the next identifier"
End Sub

'@TestMethod("LLExport")
Public Sub TestAddRowsSynchronisesDictionaryColumns()
    CustomTestSetTitles Assert, "LLExport", "TestAddRowsSynchronisesDictionaryColumns"

    Dim dict As ILLdictionary
    Dim idx As Long
    Dim startTotal As Long

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)

    For idx = 5 To 2 Step -1
        dict.RemoveColumn "Export " & CStr(idx)
    Next idx

    startTotal = Manager.NumberOfExports
    Assert.AreEqual startTotal, StoredExportTotal(), "Stored export counter should match current number before adding rows"

    Manager.AddRows dict:=dict

    Assert.AreEqual startTotal + 1, StoredExportTotal(), "Stored export counter should increment after adding rows"
    Assert.AreEqual startTotal + 1, Manager.NumberOfExports, "NumberOfExports should report the updated total after adding rows"
    Assert.IsTrue dict.ColumnExists("Export 2"), "Dictionary should expose Export 2 column after row addition"
    Assert.AreEqual StoredExportTotal(), CLng(dict.TotalNumberOfExports), "Dictionary export total should mirror stored counter"
End Sub

'@TestMethod("LLExport")
Public Sub TestAddRowsWithoutDictionaryAfterReset()
    CustomTestSetTitles Assert, "LLExport", "TestAddRowsWithoutDictionaryAfterReset"

    Manager.ResetCaches
    Manager.AddRows

    Assert.AreEqual 2, Manager.NumberOfExports, "Row count should grow when dictionary is not supplied"
    Assert.AreEqual 2, StoredExportTotal(), "Hidden export counter should align with table rows even without dictionary"
End Sub

'@TestMethod("LLExport")
Public Sub TestInsertRowsAppliesDefaultsAndSyncsDictionary()
    CustomTestSetTitles Assert, "LLExport", "TestInsertRowsAppliesDefaultsAndSyncsDictionary"
    On Error GoTo Fail

    Dim dict As ILLdictionary
    Dim selectionRange As Range
    Dim dictSheet As Worksheet
    Dim dictLo As ListObject

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)
    Set dictSheet = DictionarySheet
    Set dictLo = dictSheet.ListObjects("Tab_Dictionary")

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
    Assert.AreEqual Manager.NumberOfExports, StoredExportTotal(), _
                     "Hidden export counter should mirror the table row count"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsAppliesDefaultsAndSyncsDictionary", Err.Number, Err.Description
End Sub

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
    Assert.AreEqual initialCount, StoredExportTotal(), _
                     "Stored export counter should remain aligned when no insertion occurs"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsIgnoresSelectionOutsideTable", Err.Number, Err.Description
End Sub

'@TestMethod("LLExport")
Public Sub TestSortRenamesExportsSequentially()
    CustomTestSetTitles Assert, "LLExport", "TestSortRenamesExportsSequentially"
    On Error GoTo Fail

    Dim dict As ILLdictionary
    Dim dictLo As ListObject
    Dim exportNumberIndex As Long
    Dim lo As ListObject

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)
    Set dictLo = DictionarySheet.ListObjects("Tab_Dictionary")

    EnsureExportColumn dictLo, "Export 1"
    EnsureExportColumn dictLo, "Export 2"
    EnsureExportColumn dictLo, "Export 3"

    dictLo.ListColumns("Export 1").DataBodyRange.Cells(1, 1).Value = "One"
    dictLo.ListColumns("Export 2").DataBodyRange.Cells(1, 1).Value = "Two"
    dictLo.ListColumns("Export 3").DataBodyRange.Cells(1, 1).Value = "Three"

    Manager.AddRows dict:=dict
    Manager.AddRows dict:=dict

    Set lo = ExportSheet.ListObjects(1)
    exportNumberIndex = lo.ListColumns("export number").Index

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

    Assert.AreEqual "Three", CStr(dictLo.ListColumns("Export 1").DataBodyRange.Cells(1, 1).Value), _
                     "Dictionary column originally tied to Export 3 should now be Export 1"
    Assert.AreEqual "One", CStr(dictLo.ListColumns("Export 2").DataBodyRange.Cells(1, 1).Value), _
                     "Dictionary values should follow the new ordering"
    Assert.AreEqual "Two", CStr(dictLo.ListColumns("Export 3").DataBodyRange.Cells(1, 1).Value), _
                     "Dictionary values should follow the new ordering"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSortRenamesExportsSequentially", Err.Number, Err.Description
End Sub

'@TestMethod("LLExport")
Public Sub TestSyncDictionaryIgnoresNonExportPrefixedColumns()
    CustomTestSetTitles Assert, "LLExport", "TestSyncDictionaryIgnoresNonExportPrefixedColumns"

    Dim dict As ILLdictionary

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)

    dict.RemoveColumn "Export 3"
    dict.AddColumn "mainlab_3_backup"

    Manager.AddRows dict:=dict

    Assert.IsTrue dict.ColumnExists("mainlab_3_backup"), "Custom column should remain untouched"
    Assert.AreEqual 2, StoredExportTotal(), "Stored export counter should reflect actual export rows"
    Assert.AreEqual 2, CLng(dict.TotalNumberOfExports), "Dictionary total should ignore non-export-prefixed columns"
End Sub

'@TestMethod("LLExport")
Public Sub TestPublicSyncDictionaryExportsWithAndWithoutParameter()
    CustomTestSetTitles Assert, "LLExport", "TestPublicSyncDictionaryExportsWithAndWithoutParameter"

    Dim dict As ILLdictionary

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)

    Manager.AddRows dict:=dict

    dict.RemoveColumn "Export 2"
    Manager.SyncDictionaryExports dict
    Assert.IsTrue dict.ColumnExists("Export 2"), "SyncDictionaryExports should restore export columns when dictionary provided"

    dict.RemoveColumn "Export 2"
    Manager.SyncDictionaryExports
    Assert.IsTrue dict.ColumnExists("Export 2"), "SyncDictionaryExports should reuse cached dictionary when none supplied"
End Sub

'@TestMethod("LLExport")
Public Sub TestRemoveRowsDeletesEmpty()
    CustomTestSetTitles Assert, "LLExport", "TestRemoveRowsDeletesEmpty"
    Manager.AddRows
    ExportSheet.ListObjects(1).DataBodyRange.Rows(2).ClearContents
    Manager.RemoveRows
    Assert.AreEqual 1, Manager.NumberOfExports, "Removing rows should trim empty rows"
    Assert.AreEqual "export 1", ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("export number")).Value, _
                     "Remaining export should keep its prefixed identifier"
End Sub

'@TestMethod("LLExport")
Public Sub TestRemoveRowsPrunesDictionaryColumns()
    CustomTestSetTitles Assert, "LLExport", "TestRemoveRowsPrunesDictionaryColumns"

    Dim dict As ILLdictionary
    Dim idx As Long

    Set dict = LLdictionary.Create(DictionarySheet, 1, 1)

    For idx = 5 To 4 Step -1
        dict.RemoveColumn "Export " & CStr(idx)
    Next idx

    Manager.AddRows dict:=dict
    Manager.AddRows dict:=dict
    Assert.AreEqual 3, StoredExportTotal(), "Expected stored export counter to include added rows"

    Manager.RemoveRows dict:=dict

    Assert.AreEqual 1, StoredExportTotal(), "Stored export counter should reflect the trimmed rows"
    Assert.AreEqual 1, Manager.NumberOfExports, "NumberOfExports should reflect the trimmed export rows"
    Assert.IsFalse dict.ColumnExists("Export 2"), "Export 2 column should be removed when only one export remains"
    Assert.IsFalse dict.ColumnExists("Export 3"), "Export 3 column should be removed when not present"
    Assert.AreEqual 1, CLng(dict.TotalNumberOfExports), "Dictionary export count should match remaining exports"
End Sub

'@TestMethod("LLExport")
Public Sub TestRemoveRowsKeepsColumnsForRemainingIdentifiers()
    CustomTestSetTitles Assert, "LLExport", "TestRemoveRowsKeepsColumnsForRemainingIdentifiers"

    Dim dict As ILLdictionary
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
    Assert.AreEqual 1, StoredExportTotal(), "Stored export counter should align with current export count"
    Assert.IsFalse dict.ColumnExists("Export 4"), "Export column 4 should persist when its identifier still exists"
    Assert.IsFalse dict.ColumnExists("Export 3"), "Export column 3 should be removed when identifier is missing"
End Sub

'@TestMethod("LLExport")
Public Sub TestExportFileNameBuildsFromTemplate()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameBuildsFromTemplate"
    Dim fileName As String
    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)
    Assert.IsTrue InStr(1, fileName, "custom_value__literal_suffix", vbTextCompare) > 0, _
                  "Filename should include resolved variable and literal chunks"
    Assert.IsTrue fileName Like "*__vd0099-1234__*", "Version suffix should be appended - Actual finename: " & fileName
    Assert.IsFalse Manager.HasCheckings, "Default template should not trigger checkings for literal chunks"
End Sub

'@TestMethod("LLExport")
Public Sub TestExportFileNameHandlesSingleQuotedLiteral()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameHandlesSingleQuotedLiteral"

    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("file name")).Value = "choi_v1 + 'single literal'"

    Dim fileName As String
    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)

    Assert.IsTrue InStr(1, fileName, "custom_value__single_literal", vbTextCompare) > 0, _
                  "Single-quoted literal chunks should be preserved"
    Assert.IsFalse Manager.HasCheckings, "Single-quoted literal chunks should not trigger checkings"
End Sub

'@TestMethod("LLExport")
Public Sub TestExportFileNameWithOnlyLiteralChunks()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameWithOnlyLiteralChunks"

    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("file name")).Value = """static chunk"" + ""second part"""

    Dim fileName As String
    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)

    Assert.IsTrue InStr(1, fileName, "static_chunk__second_part", vbTextCompare) > 0, _
                  "All-literal templates should concatenate sanitized literals"
    Assert.IsFalse Manager.HasCheckings, "All-literal templates should not trigger checkings"
End Sub

'@TestMethod("LLExport")
Public Sub TestExportFileNameLogsWhenInactive()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameLogsWhenInactive"
    Manager.AddRows
    ExportSheet.ListObjects(1).DataBodyRange.Cells(2, ColumnIndexOf("status")).Value = "inactive"
    Dim name As String
    name = Manager.ExportFileName(2, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)
    Assert.IsTrue Manager.HasCheckings, "Inactive export should log information"
    Assert.IsTrue LenB(name) > 0, "Should still return a filename"
End Sub

'@TestMethod("LLExport")
Public Sub TestExportAllOverridesScope()
    CustomTestSetTitles Assert, "LLExport", "TestExportAllOverridesScope"
    Dim name As String
    name = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject, exportAll:=True)
    Assert.IsTrue InStr(1, name, "export_all", vbTextCompare) > 0, "ExportAll should override scope - fileName : " & name
End Sub

'@TestMethod("LLExport")
Public Sub TestIsActiveReflectsStatus()
    CustomTestSetTitles Assert, "LLExport", "TestIsActiveReflectsStatus"
    Assert.IsTrue Manager.IsActive(1), "Row with active status should be active"
    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("status")).Value = "inactive"
    Assert.IsFalse Manager.IsActive(1), "Row with inactive status should report false"
End Sub

'@TestMethod("LLExport")
Public Sub TestRemoveRowsHonoursThreshold()
    CustomTestSetTitles Assert, "LLExport", "TestRemoveRowsHonoursThreshold"

    Manager.AddRows
    Manager.RemoveRows rowCount:=0
    Assert.AreEqual 2, Manager.NumberOfExports, "Rows with data should remain when threshold is zero"

    Manager.RemoveRows rowCount:=2
    Assert.AreEqual 1, Manager.NumberOfExports, "Rows at or below the threshold should be trimmed"
End Sub

'@TestMethod("LLExport")
Public Sub TestActiveExportNumbersReturnsActiveRows()
    CustomTestSetTitles Assert, "LLExport", "TestActiveExportNumbersReturnsActiveRows"
    Manager.AddRows
    Manager.AddRows

    Dim statusCol As Long
    statusCol = ColumnIndexOf("status")

    ExportSheet.ListObjects(1).DataBodyRange.Cells(2, statusCol).Value = "inactive"
    ExportSheet.ListObjects(1).DataBodyRange.Cells(3, statusCol).Value = "active"

    Dim active As BetterArray
    Set active = Manager.ActiveExportNumbers

    Assert.AreEqual 2, active.Length, "Expected two active exports"
    Dim startIndex As Long
    startIndex = active.LowerBound
    Assert.AreEqual 1, CLng(active.Item(startIndex)), "First active export should be row 1"
    Assert.AreEqual 3, CLng(active.Item(startIndex + 1)), "Second active export should be row 3"
End Sub

'@TestMethod("LLExport")
Public Sub TestExportFileNameLogsMissingChunk()
    CustomTestSetTitles Assert, "LLExport", "TestExportFileNameLogsMissingChunk"
    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("file name")).Value = "unknown_chunk"

    Dim fileName As String
    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)

    Assert.IsTrue Manager.HasCheckings, "Missing chunk should produce checking entries"
    Assert.IsTrue InStr(1, fileName, "unknown_chunk", vbTextCompare) > 0, _
                  "Fallback filename should include sanitized chunk"
End Sub

'@section Helpers
'===============================================================================
Private Sub PrepareTestSheets()

    Set DictionarySheet = EnsureWorksheet(DICT_SHEET)
    PrepareDictionaryFixture DICT_SHEET
    
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

Private Sub EnsureExportColumn(ByVal dictLo As ListObject, ByVal columnName As String)
    On Error Resume Next
        dictLo.ListColumns(columnName).Name = columnName
    On Error GoTo 0
    If ColumnExists(dictLo, columnName) Then Exit Sub

    dictLo.ListColumns.Add.Name = columnName
End Sub

Private Function ColumnExists(ByVal dictLo As ListObject, ByVal columnName As String) As Boolean
    On Error Resume Next
        ColumnExists = Not dictLo.ListColumns(columnName) Is Nothing
    On Error GoTo 0
End Function

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

Private Function ColumnIndexOf(ByVal headerName As String) As Long
    ColumnIndexOf = ExportSheet.ListObjects(1).ListColumns(headerName).Index
End Function

Private Function StoredExportTotal() As Long
    Dim definition As Name
    Dim evaluated As String

    On Error Resume Next
        Set definition = ExportSheet.Names(EXPORT_TOTAL_NAME)
    On Error GoTo 0

    If definition Is Nothing Then Exit Function

    On Error Resume Next
        evaluated = Trim$(Replace$(definition.Value, "=", vbNullString))
    On Error GoTo 0

    If LenB(evaluated) <> 0 Then
        StoredExportTotal = CLng(evaluated)
    End If
End Function
