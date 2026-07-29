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
'@depends LLVariables, LLdictionary, CustomTest, DictionaryTestFixture, TestHelpersLite, BetterArray

Private Const DICT_SHEET As String = "LLVarDict"

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
    Assert.SetModuleName "TestLLVariables"
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

    'The arrange has its own handler. An error escaping it would reach the VBE
    'as a modal dialog, and a dialog stops the whole headless run.
    On Error GoTo Fail
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
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateFailsWhenNameColumnMissing", Err.Number, Err.Description
End Sub

'@sub-title Verify that Contains matches literal wildcard characters
'@details
'Arranges by writing the string "star*value?" into the first variable
'name cell so the name itself contains wildcard characters. Acts by
'calling Contains with the exact string and a case-insensitive variant.
'Asserts that both lookups succeed. The match is a string compare in
'memory, so wildcard characters carry no special meaning.
'@TestMethod("LLVariables")
Public Sub TestContainsHandlesWildcards()
    CustomTestSetTitles Assert, "LLVariables", "TestContainsHandlesWildcards"
    On Error GoTo Fail

    Dim varRange As Range

    Set varRange = Dictionary.DataRange("Variable Name")
    varRange.Cells(1, 1).Value = "star*value?"

    Set Variables = LLVariables.Create(Dictionary)
    Assert.IsTrue Variables.Contains("star*value?"), "Contains should match literal wildcard characters"
    Assert.IsTrue Variables.Contains("STAR*VALUE?", matchCase:=False), _
                  "Contains should support case-insensitive comparisons when requested"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestContainsHandlesWildcards", Err.Number, Err.Description
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
    On Error GoTo Fail

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
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSetValueHonoursOnEmpty", Err.Number, Err.Description
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

    'The fixture carries no "Column Index" column: that one is written by
    'LLdictionary.Prepare. RemoveColumn logs a warning and returns for a column
    'it cannot find, so this states the starting point rather than changing it.
    On Error GoTo Fail
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
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestIndexRaisesWhenColumnMissing", Err.Number, Err.Description
End Sub

'@sub-title Verify that Index raises InvalidArgument when the stored index is text
'@details
'Arranges by adding the "Column Index" column, which the fixture leaves to
'LLdictionary.Prepare, and writing the text "abc" into the cell of choi_v1.
'Acts by calling Variables.Index for that variable. Asserts that
'ProjectError.InvalidArgument is raised. The class used to hand back a bare
'type mismatch from CLng, with no variable name anywhere in the message.
'@TestMethod("LLVariables")
Public Sub TestIndexRaisesInvalidArgumentOnTextIndex()
    CustomTestSetTitles Assert, "LLVariables", "TestIndexRaisesInvalidArgumentOnTextIndex"

    Dim indexCell As Range

    'The arrange has its own handler, and the cell is tested before it is
    'written. Reading .Value off Nothing raises error 91, and an error escaping
    'a test proc reaches the VBE as a modal dialog, which stops the whole run.
    On Error GoTo Fail
    Dictionary.AddColumn "Column Index"
    Set indexCell = Variables.CellRange("Column Index", "choi_v1")

    If indexCell Is Nothing Then
        Assert.LogFailure "The Column Index cell of choi_v1 could not be resolved"
        Exit Sub
    End If

    indexCell.Value = "abc"

    On Error GoTo ExpectError
        Dim idx As Long
        '@Ignore VariableNotUsed, AssignmentNotUsed
        idx = Variables.Index("choi_v1")
        Assert.LogFailure "Index should raise when the stored column index is text"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "A column index that is not a number should raise InvalidArgument"
    Err.Clear
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestIndexRaisesInvalidArgumentOnTextIndex", Err.Number, Err.Description
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
    On Error GoTo Fail

    Dim names As BetterArray

    Set names = Variables.VariableNames
    Assert.IsTrue (names.Length > 0), "VariableNames should return non-empty list"
    Assert.IsTrue names.Includes("choi_v1"), "Expected known variable to appear in VariableNames list"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestVariableNamesReturnsBetterArray", Err.Number, Err.Description
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

'@sub-title Verify that a column added after creation still resolves
'@details
'Arranges by reading a value through the Variables object, which loads the
'header row into memory, then adding a brand new column to the dictionary.
'Acts by writing to that new column through SetValue. Asserts that the write
'lands. This is the shape LinelistSpecs.AddListAuto uses: it adds "list auto"
'to the dictionary and then writes to it through a variables object it built
'earlier, so a header snapshot that never refreshed would break the build.
'@TestMethod("LLVariables")
Public Sub TestColumnAddedAfterCreationResolves()
    CustomTestSetTitles Assert, "LLVariables", "TestColumnAddedAfterCreationResolves"
    On Error GoTo Fail

    Dim newColumn As Range

    'Warm the header snapshot before the column exists.
    Assert.AreEqual "choice_manual", Variables.ControlType("choi_v1"), _
                     "The fixture control type should be readable before the new column"

    Dictionary.AddColumn "late column"
    Variables.SetValue "choi_v1", "late column", "written"

    Set newColumn = Variables.CellRange("late column", "choi_v1")
    Assert.IsTrue (Not newColumn Is Nothing), "A column added after creation should resolve"
    Assert.AreEqual "written", CStr(newColumn.Value), _
                     "SetValue should write into a column added after creation"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestColumnAddedAfterCreationResolves", Err.Number, Err.Description
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
    On Error GoTo Fail

    Dim newRow As Range
    Dim names As BetterArray

    Variables.VariableNames 'Warm cache
    Set newRow = Dictionary.DataRange("Variable Name")
    newRow.Cells(newRow.Rows.Count + 1, 1).Value = "cache_test_var"

    Variables.InvalidateCaches
    Set names = Variables.VariableNames

    Assert.IsTrue names.Includes("cache_test_var"), _
                  "VariableNames should include new variables after invalidating caches"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestVariableNamesCacheInvalidation", Err.Number, Err.Description
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
    On Error GoTo Fail

    Dim sheetName As String
    Dim controlType As String
    Dim tableName As String

    sheetName = Variables.SheetName("choi_v1")
    controlType = Variables.ControlType("choi_v1")
    tableName = Variables.TableName("choi_v1")

    Assert.AreEqual "vlist1D-sheet1", sheetName, "SheetName helper should return dictionary sheet name"
    Assert.AreEqual "choice_manual", controlType, "ControlType helper should return control value"
    Assert.IsTrue LenB(tableName) = 0, "TableName helper should empty dictionary table if dictionary is not prepared"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMetadataHelpers", Err.Number, Err.Description
End Sub
