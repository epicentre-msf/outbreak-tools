Attribute VB_Name = "TestTableSpecsColumnMap"
Option Explicit
Option Private Module

'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Unit tests for the TableSpecsColumnMap helper")

Private Const MAP_SHEET As String = "TableSpecsColumnMap"

Private Assert As Object
Private Map As ITableSpecsColumnMap
Private MapSheet As Worksheet

'@section Module lifecycle
'===============================================================================

'Prepare shared helpers for the entire suite.
'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'Restore global state after all tests ran.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Map = Nothing
    Set MapSheet = Nothing
    Set Assert = Nothing
    DeleteWorksheet MAP_SHEET
End Sub

'@section Test lifecycle
'===============================================================================

'Reset the worksheet and prime the column map before each test.
'@TestInitialize
Private Sub TestInitialize()
    Set MapSheet = EnsureWorksheet(MAP_SHEET)
    SeedWorksheet
    Set Map = TableSpecsColumnMap.Create(MapSheet.Range("A1:D1"), _
                                         MapSheet.Range("A2:D2"))
End Sub

'Clear references captured during the test run.
'@TestCleanup
Private Sub TestCleanup()
    Set Map = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("TableSpecsColumnMap")
Private Sub TestColumnIndexResolvesExactHeaders()
    On Error GoTo Fail

    Assert.AreEqual 1, Map.ColumnIndex("row variable"), "Row variable should resolve to the first column"
    Assert.AreEqual 3, Map.ColumnIndex("percentage"), "Percentage column should resolve to the third position"
    Assert.IsTrue Map.ColumnExists("column variable"), "ColumnExists should confirm known headers"
    Assert.IsFalse Map.ColumnExists("missing header"), "ColumnExists should reject unknown headers"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestColumnIndexResolvesExactHeaders"
End Sub

'@TestMethod("TableSpecsColumnMap")
Private Sub TestColumnIndexSupportsPartialMatches()
    On Error GoTo Fail

    Assert.IsTrue Map.ColumnExists("percent"), "Partial matches should be recognised"
    Assert.AreEqual 2, Map.ColumnIndex("column"), "Partial column lookups should return the matching index"

    'Repeated lookups should pull from the alias cache without raising errors.
    Assert.AreEqual 2, Map.ColumnIndex("column"), "Cached partial lookup should remain stable"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestColumnIndexSupportsPartialMatches"
End Sub

'@TestMethod("TableSpecsColumnMap")
Private Sub TestValueReturnsMatchingCell()
    On Error GoTo Fail

    Assert.AreEqual "row_value", Map.Value("row variable"), "Value should return the row cell content"
    Assert.AreEqual "percentage_value", Map.Value("percentage"), "Value should retrieve the expected column entry"
    Assert.AreEqual vbNullString, Map.Value("unknown"), "Unknown columns should return an empty string"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestValueReturnsMatchingCell"
End Sub

'@TestMethod("TableSpecsColumnMap")
Private Sub TestRefreshRebuildsCacheAfterInvalidation()
    On Error GoTo Fail

    MapSheet.Range("C1").Value = "rate"
    MapSheet.Range("C2").Value = "rate_value"

    Map.Invalidate
    Map.Refresh

    Assert.AreEqual 3, Map.ColumnIndex("rate"), "Refresh should capture updated header labels"
    Assert.AreEqual -1, Map.ColumnIndex("percentage"), "Stale headers should no longer resolve"
    Assert.AreEqual "rate_value", Map.Value("rate"), "Value should align with the refreshed header"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestRefreshRebuildsCacheAfterInvalidation"
End Sub

'@section Worksheet helpers
'===============================================================================

'Populate the dedicated worksheet with predictable headers and values.
Private Sub SeedWorksheet()
    MapSheet.Range("A1").Value = "row variable"
    MapSheet.Range("B1").Value = "column variable"
    MapSheet.Range("C1").Value = "percentage"
    MapSheet.Range("D1").Value = "total flag"

    MapSheet.Range("A2").Value = "row_value"
    MapSheet.Range("B2").Value = "column_value"
    MapSheet.Range("C2").Value = "percentage_value"
    MapSheet.Range("D2").Value = "total_value"
End Sub
