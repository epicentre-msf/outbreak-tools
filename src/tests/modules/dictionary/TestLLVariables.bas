Attribute VB_Name = "TestLLVariables"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const DICT_SHEET As String = "LLVarDict"

Private Assert As Object
Private Dictionary As ILLdictionary
Private Variables As ILLVariables

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    PrepareDictionaryFixture DICT_SHEET
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
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
    Set Variables = Nothing
    Set Dictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("LLVariables")
Private Sub TestCreateFailsWhenNameColumnMissing()
    Dim dictSheet As Worksheet

    Set dictSheet = ThisWorkbook.Worksheets(DICT_SHEET)
    dictSheet.ListObjects(1).ListColumns("Variable Name").Delete

    On Error GoTo ExpectError
        Set Dictionary = LLdictionary.Create(dictSheet, 1, 1)
        Set Variables = LLVariables.Create(Dictionary)
        Assert.Fail "Create should raise when variable name column is missing"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing variable-name column should raise ElementNotFound"
    Err.Clear
End Sub

'@TestMethod("LLVariables")
Private Sub TestContainsHandlesWildcards()
    Dim varRange As Range

    Set varRange = Dictionary.DataRange("Variable Name")
    varRange.Cells(1, 1).Value = "star*value?"

    Set Variables = LLVariables.Create(Dictionary)
    Assert.IsTrue Variables.Contains("star*value?"), "Contains should match literal wildcard characters"
    Assert.IsTrue Variables.Contains("STAR*VALUE?", matchCase:=False), _
                  "Contains should support case-insensitive comparisons when requested"
End Sub

'@TestMethod("LLVariables")
Private Sub TestSetValueHonoursOnEmpty()
    Dim devComments As Range

    Set devComments = Dictionary.DataRange("Dev Comments")
    devComments.Cells(1, 1).Value = "existing"

    Variables.SetValue "choi_v1", "Dev Comments", "new text", onEmpty:=True
    Assert.AreEqual "existing", devComments.Cells(1, 1).Value, _
                     "SetValue should leave populated cells untouched when onEmpty is True"

    devComments.Cells(1, 1).ClearContents
    Variables.SetValue "choi_v1", "Dev Comments", "new text", onEmpty:=True
    Assert.AreEqual "new text", devComments.Cells(1, 1).Value, _
                     "SetValue should update empty cells when onEmpty is True"
End Sub

'@TestMethod("LLVariables")
Private Sub TestIndexRaisesWhenColumnMissing()
    Dictionary.RemoveColumn "Column Index"

    On Error GoTo ExpectError
        Dim idx As Long
        '@Ignore VariableNotUsed, AssignmentNotUsed
        idx = Variables.Index("choi_v1")
        Assert.Fail "Index should raise when column index column is missing"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing column index should raise ElementNotFound"
    Err.Clear
End Sub

'@TestMethod("LLVariables")
Private Sub TestVariableNamesReturnsBetterArray()
    Dim names As BetterArray

    Set names = Variables.VariableNames
    Assert.IsTrue (names.Length > 0), "VariableNames should return non-empty list"
    Assert.IsTrue names.Includes("choi_v1"), "Expected known variable to appear in VariableNames list"
End Sub

'@TestMethod("LLVariables")
Private Sub TestSetValueRaisesWhenColumnMissingAfterCache()
    On Error GoTo ExpectError

    Dictionary.RemoveColumn "Dev Comments"
    Variables.SetValue "choi_v1", "Dev Comments", "should fail"
    Assert.Fail "SetValue should raise when target column is missing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "SetValue should raise ElementNotFound when column removed after caching"
    Err.Clear
End Sub

'@TestMethod("LLVariables")
Private Sub TestVariableNamesCacheInvalidation()
    Dim lo As ListObject
    Dim newRow As ListRow
    Dim names As BetterArray

    Variables.VariableNames 'Warm cache

    Set lo = Dictionary.Data.Wksh.ListObjects(1)
    Set newRow = lo.ListRows.Add
    newRow.Range.Value = lo.ListRows(1).Range.Value
    newRow.Range.Cells(1, 1).Value = "cache_test_var"

    Variables.InvalidateCaches
    Set names = Variables.VariableNames

    Assert.IsTrue names.Includes("cache_test_var"), _
                  "VariableNames should include new variables after invalidating caches"
End Sub

'@TestMethod("LLVariables")
Private Sub TestMetadataHelpers()
    Dim sheetName As String
    Dim controlType As String
    Dim tableName As String
    Dim expectedTable As String

    sheetName = Variables.SheetName("choi_v1")
    controlType = Variables.ControlType("choi_v1")
    tableName = Variables.TableName("choi_v1")
    expectedTable = DictionaryFixtureValue(0, "Table Name")

    Assert.AreEqual "vlist1D-sheet1", sheetName, "SheetName helper should return dictionary sheet name"
    Assert.AreEqual "choice_manual", controlType, "ControlType helper should return control value"
    Assert.AreEqual expectedTable, tableName, "TableName helper should return dictionary table name"
End Sub
