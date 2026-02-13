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
'variables, checking state after skipped writes, and column-index cache
'invalidation when headers are renamed at the worksheet level.
'@depends LLVariables, LLdictionary, CustomTest, DictionaryTestFixture, TestHelpers

Private Const DICT_SHEET As String = "LLVarExtraDict"

Private Assert As ICustomTest
Private Dictionary As ILLdictionary
Private Variables As ILLVariables

'@section Module Lifecycle
'===============================================================================

'@sub-title Initialise the test module and prepare shared fixtures
'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLVariablesExtra"
    PrepareDictionaryFixture DICT_SHEET
End Sub

'@sub-title Tear down the module by printing results and releasing objects
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet DICT_SHEET
    Set Variables = Nothing
    Set Dictionary = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Rebuild the dictionary fixture and create fresh LLVariables before each test
'@TestInitialize
Private Sub TestInitialize()
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    Set Variables = LLVariables.Create(Dictionary)
End Sub

'@sub-title Flush assertion output and release per-test objects
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
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

'@sub-title Verify that HasCheckings is True after a skipped onEmpty SetValue
'@details
'Arranges by writing "existing" into the Dev Comments cell for choi_v1.
'Acts by calling SetValue with onEmpty True, which should skip the write
'and log a warning. Asserts that HasCheckings returns True and
'CheckingValues is not Nothing, confirming that the skip was recorded
'in the internal checking object for diagnostic purposes.
'@TestMethod("LLVariablesExtra")
Public Sub TestHasCheckingsAfterSkippedSetValue()
    CustomTestSetTitles Assert, "LLVariables", "TestHasCheckingsAfterSkippedSetValue"
    On Error GoTo Fail

    Dim devComments As Range
    Set devComments = Dictionary.DataRange("Dev Comments")
    devComments.Cells(2, 1).Value = "existing"

    Variables.SetValue "choi_v1", "Dev Comments", "new text", onEmpty:=True
    Assert.IsTrue Variables.HasCheckings, "Skipping SetValue should log a warning and create checkings"
    Assert.IsTrue (Not Variables.CheckingValues Is Nothing), "CheckingValues should expose the internal checking object"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestHasCheckingsAfterSkippedSetValue", Err.Number, Err.Description
End Sub

'@sub-title Verify that cached column indexes are re-validated against current headers
'@details
'Arranges by warming the column-index cache via a Value call for
'"Dev Comments", then renaming the header on the worksheet to
'"Dev Comments 2" to simulate a structural change. Acts by requesting
'the original "Dev Comments" header again. Asserts that the method
'returns an empty string rather than stale cached data, confirming
'that column-index caches are re-validated against the actual headers
'and gracefully handle renamed columns.
'@TestMethod("LLVariablesExtra")
Public Sub TestResolveColumnIndexCacheInvalidation()
    CustomTestSetTitles Assert, "LLVariables", "TestResolveColumnIndexCacheInvalidation"
    On Error GoTo Fail

    Dim first As String
    Dim colIdx As Long
    Dim sh As Worksheet

    'Warm the cache for Dev Comments
    '@Ignore AssignmentNotUsed
    first = Variables.Value("Dev Comments", "choi_v1")

    'Rename the header to invalidate the cached index
    colIdx = Dictionary.Data.ColumnIndex("Dev Comments", shouldExist:=True, matchCase:=False)
    Set sh = Dictionary.Data.Wksh
    sh.Cells(1, colIdx).Value = "Dev Comments 2"

    'Request the old header again; should not error and should return empty
    Assert.AreEqual vbNullString, Variables.Value("Dev Comments", "choi_v1"), _
                     "Cached column indexes should be validated against current headers"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestResolveColumnIndexCacheInvalidation", Err.Number, Err.Description
End Sub
