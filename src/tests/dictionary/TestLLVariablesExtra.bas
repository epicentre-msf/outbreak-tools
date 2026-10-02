Attribute VB_Name = "TestLLVariablesExtra"
Attribute VB_Description = "Additional tests for the LLVariables class"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@ModuleDescription("Additional tests for the LLVariables class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'Provides supplementary tests for the LLVariables class that cover edge
'cases and secondary behaviour not addressed in the core test module.
'Tests include empty-name handling, case-insensitive column lookup,
'CellRange for valid and invalid variables, error paths for unknown
'variables, the answer SetValue gives back when it skips a write, and
'column resolution after a header is renamed at the worksheet level.
'@depends LLVariables, LLdictionary, CustomTest, DictionaryTestFixture, TestHelpersLite

Private Const DICT_SHEET As String = "LLVarExtraDict"

Private Assert As CustomTest
Private Dictionary As LLdictionary
Private Variables As LLVariables

'@section Module Lifecycle
'===============================================================================

'@sub-title Initialise the test module and prepare shared fixtures
'@details
'Public, so the headless runner reaches it: the runner calls every lifecycle
'hook through Application.Run, which cannot see a Private Sub. The handler
'matters as much. An error escaping a lifecycle hook aborts the WHOLE module,
'so a fixture problem here would drop every test and the run would report no
'failures, because nothing ran.
'@ModuleInitialize
Public Sub ModuleInitialize()
    On Error GoTo Fail
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLVariablesExtra"
    PrepareDictionaryFixture DICT_SHEET
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "ModuleInitialize", Err.Number, Err.Description
End Sub

'@sub-title Tear down the module by printing results and releasing objects
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
        DeleteWorksheet DICT_SHEET
    On Error GoTo 0

    RestoreApp

    Set Variables = Nothing
    Set Dictionary = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Rebuild the dictionary fixture and create fresh LLVariables before each test
'@TestInitialize
Public Sub TestInitialize()
    On Error GoTo Fail
    BusyApp
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    Set Variables = LLVariables.Create(Dictionary)
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInitialize", Err.Number, Err.Description
End Sub

'@sub-title Flush assertion output and release per-test objects
'@TestCleanup
Public Sub TestCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.Flush
        End If
    On Error GoTo 0

    Set Variables = Nothing
    Set Dictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@sub-title Verify that Contains returns False for an empty variable name
'@details
'Acts by calling Variables.Contains with vbNullString as the variable
'name. Asserts that the method returns False without raising an error,
'confirming that empty-string inputs are handled gracefully rather than
'causing a lookup failure or match against blank cells.
'@TestMethod("LLVariablesExtra")
Public Sub TestContainsReturnsFalseForEmptyName()
    CustomTestSetTitles Assert, "LLVariables", "TestContainsReturnsFalseForEmptyName"
    On Error GoTo Fail

    Assert.IsFalse Variables.Contains(vbNullString), "Contains should return False for empty names"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestContainsReturnsFalseForEmptyName", Err.Number, Err.Description
End Sub

'@sub-title Verify that Value resolves column headers case-insensitively
'@details
'Arranges using the standard fixture which has a "Main Label" column.
'Acts by calling Variables.Value with "main label" in lowercase for the
'known variable "choi_v1". Asserts that the returned value matches the
'expected fixture data, confirming that column header lookup is
'case-insensitive.
'@TestMethod("LLVariablesExtra")
Public Sub TestValueCaseInsensitiveColumnLookup()
    CustomTestSetTitles Assert, "LLVariables", "TestValueCaseInsensitiveColumnLookup"
    On Error GoTo Fail

    Dim val As String
    val = Variables.Value("main label", "choi_v1")
    Assert.AreEqual "Choices on vlist1D", val, _
                     "Value should resolve headers ignoring case differences"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValueCaseInsensitiveColumnLookup", Err.Number, Err.Description
End Sub

'@sub-title Verify that Value returns an empty string for a cell holding an error
'@details
'Arranges by writing an #N/A error value straight into the Dev Comments cell
'of choi_v1. The value is written straight into the cell, because the lifecycle
'hooks put Excel in manual calculation and a formula would sit there
'uncalculated. Acts by reading that cell through Variables.Value. Asserts that
'an empty string comes back.
'Reading the cell straight into a String used to raise a type mismatch, and
'Formulas reads the table name column this way once per token of every
'formula in the workbook.
'@TestMethod("LLVariablesExtra")
Public Sub TestValueReturnsEmptyForErrorCell()
    CustomTestSetTitles Assert, "LLVariables", "TestValueReturnsEmptyForErrorCell"
    On Error GoTo Fail

    Dim targetCell As Range

    Set targetCell = Variables.CellRange("Dev Comments", "choi_v1")
    targetCell.Value = CVErr(xlErrNA)

    Assert.AreEqual vbNullString, Variables.Value("Dev Comments", "choi_v1"), _
                     "A cell holding an error value should read as an empty string"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValueReturnsEmptyForErrorCell", Err.Number, Err.Description
End Sub

'@sub-title Verify that CellRange returns a Range for valid variables and Nothing for unknown ones
'@details
'Arranges using the standard fixture with a "Dev Comments" column. Acts
'by calling CellRange for the known variable "choi_v1" and then for the
'nonexistent variable "__unknown__". Asserts that the first call returns
'a non-Nothing Range object, and the second call returns Nothing,
'confirming correct behaviour for both valid and invalid variable names.
'@TestMethod("LLVariablesExtra")
Public Sub TestCellRangeValidAndInvalid()
    CustomTestSetTitles Assert, "LLVariables", "TestCellRangeValidAndInvalid"
    On Error GoTo Fail

    Dim rng As Range
    Set rng = Variables.CellRange("Dev Comments", "choi_v1")
    Assert.IsTrue (Not rng Is Nothing), "CellRange should return a usable Range for existing values"

    Set rng = Variables.CellRange("Dev Comments", "__unknown__")
    Assert.IsTrue (rng Is Nothing), "CellRange should return Nothing for unknown variables"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCellRangeValidAndInvalid", Err.Number, Err.Description
End Sub

'@sub-title Verify that SetValue raises for a nonexistent variable name
'@details
'Acts by calling SetValue with the variable name "__missing__" which
'does not exist in the dictionary fixture. Asserts that a
'ProjectError.ElementNotFound error is raised, confirming that the
'method validates the variable name before attempting to write and
'surfaces a clear error for unknown variables.
'@TestMethod("LLVariablesExtra")
Public Sub TestSetValueRaisesForUnknownVariable()
    CustomTestSetTitles Assert, "LLVariables", "TestSetValueRaisesForUnknownVariable"
    On Error GoTo ExpectError

    Variables.SetValue "__missing__", "Dev Comments", "value"
    Assert.LogFailure "SetValue should raise when variable is absent"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing variable should raise ElementNotFound when setting values"
    Err.Clear
End Sub

'@sub-title Verify that SetValue says whether it wrote
'@details
'Arranges by writing "existing" into the Dev Comments cell for choi_v1.
'Acts by calling SetValue with onEmpty True, which skips the write, then
'clears the cell and calls it again. Asserts that the first call answers
'False and the second True. The skipped write used to be recorded in a
'checking object that nothing in the tree ever read.
'@TestMethod("LLVariablesExtra")
Public Sub TestSetValueReportsSkippedWrite()
    CustomTestSetTitles Assert, "LLVariables", "TestSetValueReportsSkippedWrite"
    On Error GoTo Fail

    Dim devComments As Range
    Dim written As Boolean

    Set devComments = Dictionary.DataRange("Dev Comments")
    devComments.Cells(2, 1).Value = "existing"

    written = Variables.SetValue("choi_v1", "Dev Comments", "new text", onEmpty:=True)
    Assert.IsFalse written, "SetValue should answer False when it leaves a populated cell alone"

    devComments.Cells(2, 1).ClearContents
    written = Variables.SetValue("choi_v1", "Dev Comments", "new text", onEmpty:=True)
    Assert.IsTrue written, "SetValue should answer True when it writes"
    Assert.AreEqual "new text", devComments.Cells(2, 1).Value, _
                     "SetValue should write into the empty cell"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSetValueReportsSkippedWrite", Err.Number, Err.Description
End Sub

'@sub-title Verify that a renamed header can never answer from the held snapshot
'@details
'Arranges by reading "Dev Comments" once, which loads the header row into
'memory, then renaming that header on the worksheet to "Dev Comments 2".
'Acts by requesting the original header again. Asserts that an empty string
'comes back. The held column is checked with one cell read before it is used,
'so a header renamed behind this object can never make it read the wrong
'column's value.
'@TestMethod("LLVariablesExtra")
Public Sub TestResolveColumnIndexCacheInvalidation()
    CustomTestSetTitles Assert, "LLVariables", "TestResolveColumnIndexCacheInvalidation"
    On Error GoTo Fail

    Dim first As String
    Dim colIdx As Long
    Dim sh As Worksheet

    'Warm the header snapshot for Dev Comments
    '@Ignore AssignmentNotUsed
    first = Variables.Value("Dev Comments", "choi_v1")

    'Rename the header to make the held column stale
    colIdx = Dictionary.Data.ColumnIndex("Dev Comments", shouldExist:=True, matchCase:=False)
    Set sh = Dictionary.Data.Wksh
    sh.Cells(1, colIdx).Value = "Dev Comments 2"

    'Request the old header again; should not error and should return empty
    Assert.AreEqual vbNullString, Variables.Value("Dev Comments", "choi_v1"), _
                     "A held column index should be checked against the current header"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestResolveColumnIndexCacheInvalidation", Err.Number, Err.Description
End Sub
