Attribute VB_Name = "TestLLVariables"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@Folder("CustomTests")

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@ModuleDescription("Core tests for the LLVariables class")

'@description
'Validates the core behaviour of the LLVariables class including variable
'lookup, value mutation, cache management, and metadata accessors. Each
'test creates a fresh dictionary fixture so that worksheet state does not
'leak between runs. Error-path tests verify that missing columns and
'invalid state raise the expected ProjectError codes.
'@depends LLVariables, LLdictionary, CustomTest, DictionaryTestFixture, TestHelpers, BetterArray

Private Const DICT_SHEET As String = "LLVarDict"

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
    Assert.SetModuleName "TestLLVariables"
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

'@sub-title Verify that Create raises when the variable name column is absent
'@details
'Arranges by deleting the first column of the dictionary sheet so the
'variable name header is missing. Acts by calling LLVariables.Create with
'the mutilated dictionary. Asserts that a ProjectError.ElementNotFound
'error is raised, confirming the class validates its required column
'during construction.
'@TestMethod("LLVariables")
Public Sub TestCreateFailsWhenNameColumnMissing()
    CustomTestSetTitles Assert, "LLVariables", "TestCreateFailsWhenNameColumnMissing"
    Dim dictSheet As Worksheet

    Set dictSheet = ThisWorkbook.Worksheets(DICT_SHEET)
    dictSheet.Columns(1).Delete

    On Error GoTo ExpectError
        Set Dictionary = LLdictionary.Create(dictSheet, 1, 1)
        Set Variables = LLVariables.Create(Dictionary)
        Assert.LogFailure "Create should raise when variable name column is missing"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing variable-name column should raise ElementNotFound"
    Err.Clear
End Sub

'@sub-title Verify that Contains matches literal wildcard characters
'@details
'Arranges by writing the string "star*value?" into the first variable
'name cell so the name itself contains wildcard characters. Acts by
'calling Contains with the exact string and a case-insensitive variant.
'Asserts that both lookups succeed, confirming that Contains treats
'wildcard characters literally rather than as pattern metacharacters.
'@TestMethod("LLVariables")
Public Sub TestContainsHandlesWildcards()
    CustomTestSetTitles Assert, "LLVariables", "TestContainsHandlesWildcards"
    Dim varRange As Range

    Set varRange = Dictionary.DataRange("Variable Name")
    varRange.Cells(1, 1).Value = "star*value?"

    Set Variables = LLVariables.Create(Dictionary)
    Assert.IsTrue Variables.Contains("star*value?"), "Contains should match literal wildcard characters"
    Assert.IsTrue Variables.Contains("STAR*VALUE?", matchCase:=False), _
                  "Contains should support case-insensitive comparisons when requested"
End Sub

'@sub-title Verify that SetValue respects the onEmpty flag
'@details
'Arranges by writing "existing" into the Dev Comments cell for choi_v1.
'Acts by calling SetValue with onEmpty True, then verifies the cell is
'unchanged. Clears the cell and calls SetValue again with onEmpty True.
'Asserts that the empty cell receives the new value, confirming the
'conditional-write behaviour of the onEmpty parameter.
'@TestMethod("LLVariables")
Public Sub TestSetValueHonoursOnEmpty()
    CustomTestSetTitles Assert, "LLVariables", "TestSetValueHonoursOnEmpty"
    Dim devComments As Range

    Set devComments = Dictionary.DataRange("Dev Comments")
    devComments.Cells(2, 1).Value = "existing"

    Variables.SetValue "choi_v1", "Dev Comments", "new text", onEmpty:=True
    Assert.AreEqual "existing", devComments.Cells(2, 1).Value, _
                     "SetValue should leave populated cells untouched when onEmpty is True"

    devComments.Cells(2, 1).ClearContents
    Variables.SetValue "choi_v1", "Dev Comments", "new text", onEmpty:=True
    Assert.AreEqual "new text", devComments.Cells(2, 1).Value, _
                     "SetValue should update empty cells when onEmpty is True"
End Sub

'@sub-title Verify that Index raises when the column index column is missing
'@details
'Arranges by removing the "Column Index" column from the dictionary.
'Acts by calling Variables.Index for a known variable. Asserts that a
'ProjectError.ElementNotFound error is raised, confirming that Index
'validates the presence of the column-index column before returning
'a result.
'@TestMethod("LLVariables")
Public Sub TestIndexRaisesWhenColumnMissing()
    CustomTestSetTitles Assert, "LLVariables", "TestIndexRaisesWhenColumnMissing"
    Dictionary.RemoveColumn "Column Index"

    On Error GoTo ExpectError
        Dim idx As Long
        '@Ignore VariableNotUsed, AssignmentNotUsed
        idx = Variables.Index("choi_v1")
        Assert.LogFailure "Index should raise when column index column is missing"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing column index should raise ElementNotFound"
    Err.Clear
End Sub

'@sub-title Verify that VariableNames returns a populated BetterArray
'@details
'Acts by calling Variables.VariableNames with no prior arrangement
'beyond the standard fixture. Asserts that the returned BetterArray
'has a positive length and includes the known variable "choi_v1",
'confirming that the method correctly reads variable names from
'the dictionary.
'@TestMethod("LLVariables")
Public Sub TestVariableNamesReturnsBetterArray()
    CustomTestSetTitles Assert, "LLVariables", "TestVariableNamesReturnsBetterArray"
    Dim names As BetterArray

    Set names = Variables.VariableNames
    Assert.IsTrue (names.Length > 0), "VariableNames should return non-empty list"
    Assert.IsTrue names.Includes("choi_v1"), "Expected known variable to appear in VariableNames list"
End Sub

'@sub-title Verify that SetValue raises when the target column is removed after caching
'@details
'Arranges by removing the "Dev Comments" column from the dictionary
'after the Variables object has already been created and may have cached
'column positions. Acts by calling SetValue targeting the removed column.
'Asserts that a ProjectError.ElementNotFound error is raised, confirming
'that stale cache entries do not mask missing columns.
'@TestMethod("LLVariables")
Public Sub TestSetValueRaisesWhenColumnMissingAfterCache()
    CustomTestSetTitles Assert, "LLVariables", "TestSetValueRaisesWhenColumnMissingAfterCache"
    On Error GoTo ExpectError

    Dictionary.RemoveColumn "Dev Comments"
    Variables.SetValue "choi_v1", "Dev Comments", "should fail"
    Assert.LogFailure "SetValue should raise when target column is missing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "SetValue should raise ElementNotFound when column removed after caching"
    Err.Clear
End Sub

'@sub-title Verify that VariableNames reflects new entries after cache invalidation
'@details
'Arranges by warming the VariableNames cache with an initial call, then
'appending a new variable name "cache_test_var" directly to the dictionary
'sheet. Acts by calling InvalidateCaches followed by VariableNames again.
'Asserts that the newly added variable appears in the refreshed list,
'confirming that cache invalidation forces a re-read of the underlying data.
'@TestMethod("LLVariables")
Public Sub TestVariableNamesCacheInvalidation()
    CustomTestSetTitles Assert, "LLVariables", "TestVariableNamesCacheInvalidation"

    Dim newRow As Range
    Dim names As BetterArray

    Variables.VariableNames 'Warm cache
    Set newRow = Dictionary.DataRange("Variable Name")
    newRow.Cells(newRow.Rows.Count + 1, 1).Value = "cache_test_var"

    Variables.InvalidateCaches
    Set names = Variables.VariableNames

    Assert.IsTrue names.Includes("cache_test_var"), _
                  "VariableNames should include new variables after invalidating caches"
End Sub

'@sub-title Verify that metadata helpers return expected dictionary values
'@details
'Acts by calling SheetName, ControlType, and TableName for the known
'variable "choi_v1" against the standard dictionary fixture. Asserts
'that SheetName and ControlType return the expected fixture values, and
'that TableName returns an empty string because the dictionary has not
'been prepared yet, confirming correct delegation to the underlying
'dictionary columns.
'@TestMethod("LLVariables")
Public Sub TestMetadataHelpers()
    CustomTestSetTitles Assert, "LLVariables", "TestMetadataHelpers"
    Dim sheetName As String
    Dim controlType As String
    Dim tableName As String

    sheetName = Variables.SheetName("choi_v1")
    controlType = Variables.ControlType("choi_v1")
    tableName = Variables.TableName("choi_v1")

    Assert.AreEqual "vlist1D-sheet1", sheetName, "SheetName helper should return dictionary sheet name"
    Assert.AreEqual "choice_manual", controlType, "ControlType helper should return control value"
    Assert.IsTrue LenB(tableName) = 0, "TableName helper should empty dictionary table if dictionary is not prepared"
End Sub
