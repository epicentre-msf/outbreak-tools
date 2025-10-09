Attribute VB_Name = "TestLLVariables"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@Folder("CustomTests")
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const DICT_SHEET As String = "LLVarDict"

Private Assert As ICustomTest
Private Dictionary As ILLdictionary
Private Variables As ILLVariables

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLVariables"
    PrepareDictionaryFixture DICT_SHEET
End Sub

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

'@TestInitialize
Private Sub TestInitialize()
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    Set Variables = LLVariables.Create(Dictionary)
End Sub

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

'@TestMethod("LLVariables")
Public Sub TestVariableNamesReturnsBetterArray()
    CustomTestSetTitles Assert, "LLVariables", "TestVariableNamesReturnsBetterArray"
    Dim names As BetterArray

    Set names = Variables.VariableNames
    Assert.IsTrue (names.Length > 0), "VariableNames should return non-empty list"
    Assert.IsTrue names.Includes("choi_v1"), "Expected known variable to appear in VariableNames list"
End Sub

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
