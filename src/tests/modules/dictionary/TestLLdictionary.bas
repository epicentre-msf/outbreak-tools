Attribute VB_Name = "TestLLdictionary"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'Dictionary-focused tests rely on the shared fixture defined in
'`DictionaryTestFixture`. The helper module mirrors the contents of
''src/classes/implements/draft.csv` so every consumer (DataSheet, LLdictionary,
'etc.) exercises the same dataset without touching the filesystem.

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@Folder("Tests")

Private Const DICT_SHEET As String = "LLDictTest"

Private Assert As ICustomTest
Private Dictionary As ILLdictionary

'@section Fixture lifecycle
'===============================================================================

Private Sub ResetDictionarySheet()
    PrepareDictionaryFixture DICT_SHEET
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLdictionary"
    ResetDictionarySheet
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    'Release references captured during `ModuleInitialize`. Keeping things tidy
    'helps when the suite is executed repeatedly within the same Excel session.
    Set Assert = Nothing
    Set Dictionary = Nothing
    DeleteWorksheet DICT_SHEET

End Sub

'@TestInitialize
Private Sub TestInitialize()
    'Refresh the fixture worksheet ahead of each test to guarantee isolation.
    ResetDictionarySheet
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    'Drop references to ensure subsequent tests cannot accidentally reuse
    'stateful resources from previous runs.
    Set Dictionary = Nothing
End Sub

'@TestMethod("LLdictionary")
Public Sub TestCreateInitialisesData()
    CustomTestSetTitles Assert, "LLdictionary", "TestCreateInitialisesData"
    On Error GoTo Fail
    Assert.IsTrue (TypeOf Dictionary Is ILLdictionary), "Expected Create to yield an interface implementation"
    Assert.IsTrue (Dictionary.Data.DataStartRow = 1), "Start row should remain at 1"
    Assert.IsTrue (Dictionary.Data.DataStartColumn = 1), "Start column should remain at 1"
    Assert.IsTrue (Dictionary.Data.Wksh.Name = DICT_SHEET), "Dictionary should target the configured sheet"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateInitialisesData", Err.Number, Err.Description
End Sub


'@TestMethod("LLdictionary")
Public Sub TestColumnExistsAndValidity()
    CustomTestSetTitles Assert, "LLdictionary", "TestColumnExistsAndValidity"
    On Error GoTo Fail
    Assert.IsTrue Dictionary.ColumnExists("variable name"), "variable name column should exist"
    Assert.IsTrue (Not Dictionary.ColumnExists("random column for testing")), "Unexpected column reported as existing"
    Assert.IsTrue (Dictionary.ColumnExists("control", checkValidity:=True)), "Control column should be recognised when validating"
    Assert.IsTrue (Not Dictionary.ColumnExists("column indexes", checkValidity:=True)), "Validation should fail for unsupported header"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestColumnExistsAndValidity", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestVariableAndUniqueValues()
    CustomTestSetTitles Assert, "LLdictionary", "TestVariableAndUniqueValues"
    Dim sheetsList As BetterArray
    Dim expectedSheets As BetterArray
    Dim idx As Long
    Dim firstVar As String

    On Error GoTo Fail

    Set sheetsList = Dictionary.UniqueValues("sheet name")
    Set expectedSheets = DictionaryDistinctValues("Sheet Name")

    Assert.IsTrue (sheetsList.Length = expectedSheets.Length), "UniqueValues should capture all sheet names"

    For idx = expectedSheets.LowerBound To expectedSheets.UpperBound
        Assert.IsTrue sheetsList.Includes(CStr(expectedSheets.Item(idx))), "Expected sheet " & CStr(expectedSheets.Item(idx)) & " to be listed"
    Next idx

    firstVar = DictionaryFixtureValue(0, "Variable Name")
    Assert.IsTrue Dictionary.VariableExists(firstVar), "First fixture variable should exist"
    Assert.IsTrue (Not Dictionary.VariableExists("missing_var")), "Unexpected variable reported as present"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestVariableAndUniqueValues", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestSpecialVariableSelectors()
    CustomTestSetTitles Assert, "LLdictionary", "TestSpecialVariableSelectors"
    Dim choices As BetterArray
    Dim geos As BetterArray
    Dim times As BetterArray
    Dim expectedChoices As BetterArray
    Dim expectedGeos As BetterArray
    Dim expectedTimes As BetterArray
    Dim idx As Long

    On Error GoTo Fail

    Set choices = Dictionary.ChoicesVars
    Set geos = Dictionary.GeoVars
    Set times = Dictionary.TimeVars

    Set expectedChoices = DictionaryControlMatches(Array("choice_manual", "choice_formula"))
    Set expectedGeos = DictionaryControlMatches(Array("geo", "hf"))
    Set expectedTimes = DictionaryFieldEquals("Variable Type", "date")

    Assert.IsTrue (choices.Length = expectedChoices.Length), "ChoicesVars count mismatch"
    For idx = expectedChoices.LowerBound To expectedChoices.UpperBound
        Assert.IsTrue choices.Includes(CStr(expectedChoices.Item(idx))), "ChoicesVars missing " & CStr(expectedChoices.Item(idx))
    Next idx

    Assert.IsTrue (geos.Length = expectedGeos.Length), "GeoVars count mismatch"
    For idx = expectedGeos.LowerBound To expectedGeos.UpperBound
        Assert.IsTrue geos.Includes(CStr(expectedGeos.Item(idx))), "GeoVars missing " & CStr(expectedGeos.Item(idx))
    Next idx

    Assert.IsTrue (times.Length = expectedTimes.Length), "TimeVars count mismatch"
    For idx = expectedTimes.LowerBound To expectedTimes.UpperBound
        Assert.IsTrue times.Includes(CStr(expectedTimes.Item(idx))), "TimeVars missing " & CStr(expectedTimes.Item(idx))
    Next idx

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSpecialVariableSelectors", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestInsertAndRemoveColumn()
    CustomTestSetTitles Assert, "LLdictionary", "TestInsertAndRemoveColumn"

    On Error GoTo Fail

    Dictionary.InsertColumn "custom export", "sheet type"
    Assert.IsTrue (Dictionary.ColumnExists("custom export")), "InsertColumn should add headers"

    Dictionary.RemoveColumn "custom export"
    Assert.IsTrue (Not Dictionary.ColumnExists("custom export")), "RemoveColumn should delete headers"

    Dictionary.AddColumn "after range"
    Assert.IsTrue Dictionary.ColumnExists("after range"), "AddColumn should append headers at the end"
    Dictionary.RemoveColumn "after range"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertAndRemoveColumn", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestCleanRemovesUnknownColumns()
    CustomTestSetTitles Assert, "LLdictionary", "TestCleanRemovesUnknownColumns"
    Dim sh As Worksheet
    Dim endCol As Long

    On Error GoTo Fail

    Set sh = Dictionary.Data.Wksh
    endCol = sh.Cells(1, sh.Columns.Count).End(xlToLeft).Column + 1
    sh.Cells(1, endCol).Value = "temp column"

    Dictionary.Clean removeAddedColumns:=True
    Assert.IsTrue (Not Dictionary.ColumnExists("temp column")), "Clean should remove unknown columns"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCleanRemovesUnknownColumns", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestExportCreatesWorkbook()
    CustomTestSetTitles Assert, "LLdictionary", "TestExportCreatesWorkbook"

    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim expectedRow As Long
    Dim expectedCol As Long

    On Error GoTo Fail

    Set exportBook = NewWorkbook()

    Dictionary.Export exportBook

    Set exportedSheet = exportBook.Worksheets(DICT_SHEET)

    Assert.IsTrue (exportedSheet.ListObjects.Count = 1), "Export should add a table to the destination workbook"

    expectedRow = Dictionary.Data.DataEndRow + 1
    expectedCol = Dictionary.Data.DataStartColumn

    Assert.IsTrue (exportedSheet.Cells(expectedRow, expectedCol).Font.Color = vbBlue), "Export should mark the sheet as prepared"
    Assert.IsTrue (DICTIONARY_FIXTURE_LAST_COLOR = exportedSheet.Cells(Dictionary.Data.DataEndRow, Dictionary.Data.DataEndColumn).Interior.Color), "Fixture formatting should persist after export"

    If Not exportBook Is Nothing Then DeleteWorkbook exportBook
    Exit Sub

Fail:
    If Not exportBook Is Nothing Then DeleteWorkbook exportBook
    CustomTestLogFailure Assert, "TestExportCreatesWorkbook", Err.Number, Err.Description
End Sub
