Attribute VB_Name = "TestLLExport"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const EXPORT_SHEET As String = "LLExportSpec"
Private Const DICT_SHEET As String = "LLExportDict"
Private Const VLIST_SHEET As String = "vlist1D-sheet1"
Private Const PASSWORD_SHEET As String = "LLExportPasswords"

Private Assert As Object
Private DictionarySheet As Worksheet
Private ExportSheet As Worksheet
Private VListSheet As Worksheet
Private Manager As ILLExport
Private PasswordSheet As Worksheet
Private PasswordsSubject As IPasswords

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    PrepareTestSheets
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
    ThisWorkbook.Names("choi_v1").Delete
    DeleteWorksheet EXPORT_SHEET
    DeleteWorksheet DICT_SHEET
    DeleteWorksheet VLIST_SHEET
    DeleteWorksheet PASSWORD_SHEET
    On Error GoTo 0
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    PrepareTestSheets
    Set Manager = LLExport.Create(ExportSheet)
    Set PasswordsSubject = Passwords.Create(PasswordSheet)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Manager = Nothing
    Set PasswordsSubject = Nothing
End Sub

'@TestMethod("LLExport")
Private Sub TestCreateInitialisesData()
    Assert.IsTrue (Not Manager.Data Is Nothing), "Expected Data to be initialised"
    Assert.AreEqual 1, Manager.NumberOfExports, "Should report single export row"
End Sub

'@TestMethod("LLExport")
Private Sub TestAddRowsAppliesDefaults()
    Manager.AddRows 1
    Assert.AreEqual 2, Manager.NumberOfExports, "Row count should grow by one"
    Assert.AreEqual "no", Manager.ColumnValue(2, "include personal identifiers"), _
                     "Include personal identifiers should default to 'no'"
End Sub

'@TestMethod("LLExport")
Private Sub TestRemoveRowsDeletesEmpty()
    Manager.AddRows 1
    ExportSheet.ListObjects(1).DataBodyRange.Rows(2).ClearContents
    Manager.RemoveRows totalCount:=0
    Assert.AreEqual 1, Manager.NumberOfExports, "Removing rows should trim empty rows"
End Sub

'@TestMethod("LLExport")
Private Sub TestExportFileNameBuildsFromTemplate()
    Dim fileName As String
    fileName = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)
    Assert.IsTrue InStr(1, fileName, "custom_value", vbTextCompare) > 0, _
                  "Filename should include resolved variable value"
    Assert.IsTrue fileName Like "*__v001-PK__*", "Version suffix should be appended"
End Sub

'@TestMethod("LLExport")
Private Sub TestExportFileNameLogsWhenInactive()
    Manager.AddRows 1
    ExportSheet.ListObjects(1).DataBodyRange.Cells(2, ColumnIndexOf("status")).Value = "inactive"
    Dim name As String
    name = Manager.ExportFileName(2, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject)
    Assert.IsTrue Manager.HasCheckings, "Inactive export should log information"
    Assert.IsTrue LenB(name) > 0, "Should still return a filename"
End Sub

'@TestMethod("LLExport")
Private Sub TestExportAllOverridesScope()
    Dim name As String
    name = Manager.ExportFileName(1, LLdictionary.Create(DictionarySheet, 1, 1), PasswordsSubject, exportAll:=True)
    Assert.IsTrue InStr(1, name, "export_all", vbTextCompare) > 0, "ExportAll should override scope"
End Sub

'@TestMethod("LLExport")
Private Sub TestIsActiveReflectsStatus()
    Assert.IsTrue Manager.IsActive(1), "Row with active status should be active"
    ExportSheet.ListObjects(1).DataBodyRange.Cells(1, ColumnIndexOf("status")).Value = "inactive"
    Assert.IsFalse Manager.IsActive(1), "Row with inactive status should report false"
End Sub

'@TestMethod("LLExport")
Private Sub TestAddRowsRejectsInvalidCount()
    On Error GoTo ExpectError

    Manager.AddRows 0
    Assert.Fail "AddRows should reject counts smaller than one"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Invalid row counts should raise InvalidArgument"
    Err.Clear
End Sub

'@TestMethod("LLExport")
Private Sub TestActiveExportNumbersReturnsActiveRows()
    Manager.AddRows 2

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
Private Sub TestExportFileNameLogsMissingChunk()
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
    PrepareDictionaryFixture DICT_SHEET
    Set DictionarySheet = EnsureWorksheet(DICT_SHEET)

    Set ExportSheet = EnsureWorksheet(EXPORT_SHEET)
    ClearWorksheet ExportSheet
    PrepareExportTable ExportSheet

    Set VListSheet = EnsureWorksheet(VLIST_SHEET)
    ClearWorksheet VListSheet
    VListSheet.Range("A1").Value = "custom value"
    On Error Resume Next
    ThisWorkbook.Names("choi_v1").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="choi_v1", RefersTo:=VListSheet.Range("A1")

    PasswordsTestFixture.PreparePasswordsFixture PASSWORD_SHEET
    Set PasswordSheet = ThisWorkbook.Worksheets(PASSWORD_SHEET)
End Sub

Private Sub PrepareExportTable(ByVal targetSheet As Worksheet)
    Dim headers As Variant
    Dim dataRow As Variant
    headers = Array("export number", "status", "label button", _
                    "file format", "file name", "password", _
                    "include personal identifiers", "include p-codes", _
                    "header format", "export metadata sheets", _
                    "export analyses sheets")

    dataRow = Array(1, "active", "Label", "xlsx", "choi_v1 + custom_value", "pwd", _
                    "", "yes", "default", "no", "no")

    targetSheet.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    targetSheet.Range("A2").Resize(1, UBound(dataRow) + 1).Value = dataRow
    targetSheet.ListObjects.Add SourceType:=xlSrcRange, _
            Source:=targetSheet.Range("A1").Resize(2, UBound(headers) + 1), XlListObjectHasHeaders:=xlYes
End Sub

Private Function ColumnIndexOf(ByVal headerName As String) As Long
    ColumnIndexOf = ExportSheet.ListObjects(1).ListColumns(headerName).Index
End Function
