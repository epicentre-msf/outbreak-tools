Attribute VB_Name = "TestLLdictionary"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'Dictionary-focused tests rely on the shared fixture defined in
'`DictionaryTestFixture`. The helper module mirrors the contents of
''src/classes/implements/draft.csv` so every consumer (DataSheet, LLdictionary,
'etc.) exercises the same dataset without touching the filesystem.

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")


Private Const DICT_SHEET As String = "LLDictTest"
Private Const EXPORT_TOTAL_NAME As String = "__ll_exports_total__"

Private Assert As ICustomTest
Private Dictionary As ILLdictionary

'@section Fixture lifecycle
'===============================================================================

Private Sub ResetDictionarySheet()
    PrepareDictionaryFixture DICT_SHEET
    RemoveDictionaryExportName ThisWorkbook.Worksheets(DICT_SHEET)
End Sub

Private Function EnsureDictionaryListObject() As ListObject
    Dim dictSheet As Worksheet
    Dim dataRange As Range
    Dim listObj As ListObject

    Set dictSheet = ThisWorkbook.Worksheets(DICT_SHEET)

    On Error Resume Next
        dictSheet.ListObjects(1).Delete
    On Error GoTo 0

    Set dataRange = dictSheet.Range("A1").CurrentRegion
    Set listObj = dictSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                            Source:=dataRange, _
                                            XlListObjectHasHeaders:=xlYes)
    listObj.Name = "tblLLDictionary"

    Set EnsureDictionaryListObject = listObj
End Function

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
    RemoveDictionaryExportName ThisWorkbook.Worksheets(DICT_SHEET)
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
    Set expectedTimes = DictionaryControlMatches(Array("date"), "Variable Type")

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
Public Sub TestRenameColumnUpdatesHeader()
    CustomTestSetTitles Assert, "LLdictionary", "TestRenameColumnUpdatesHeader"
    On Error GoTo Fail

    Dictionary.RenameColumn "main label", "main label renamed"

    Assert.IsTrue Dictionary.ColumnExists("main label renamed"), "RenameColumn should update the header"
    Assert.IsFalse Dictionary.ColumnExists("main label"), "Old header should no longer be present"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRenameColumnUpdatesHeader", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestDeleteRowsRemovesSelection()
    CustomTestSetTitles Assert, "LLdictionary", "TestDeleteRowsRemovesSelection"
    On Error GoTo Fail

    Dim lo As ListObject
    Dim baseline As Long

    Set lo = EnsureDictionaryListObject()
    baseline = lo.ListRows.Count

    Dictionary.DeleteRows lo.ListRows(2).Range

    Assert.AreEqual baseline - 1, lo.ListRows.Count, "DeleteRows should remove the targeted row"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDeleteRowsRemovesSelection", Err.Number, Err.Description
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
Public Sub TestInsertRowsMirrorsSelectionHeight()
    CustomTestSetTitles Assert, "LLdictionary", "TestInsertRowsMirrorsSelectionHeight"
    On Error GoTo Fail

    Dim lo As ListObject
    Dim selectionRange As Range
    Dim initialRows As Long
    Dim preservedValue As String

    Set lo = EnsureDictionaryListObject()
    preservedValue = CStr(lo.DataBodyRange.Cells(2, 1).Value)
    initialRows = lo.ListRows.Count

    Set selectionRange = lo.ListRows(2).Range
    Set selectionRange = selectionRange.Resize(2, lo.ListColumns.Count)

    Dictionary.InsertRows selectionRange

    Assert.AreEqual initialRows + 2, lo.ListRows.Count, _
        "InsertRows should add as many entries as rows selected"
    Assert.AreEqual vbNullString, CStr(lo.ListRows(2).Range.Cells(1, 1).Value), _
        "First inserted row should be blank"
    Assert.AreEqual preservedValue, CStr(lo.ListRows(3).Range.Cells(1, 1).Value), _
        "Existing data should shift below inserted rows"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsMirrorsSelectionHeight", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestPrepareAddsHelperColumns()
    CustomTestSetTitles Assert, "LLdictionary", "TestPrepareAddsHelperColumns"

    On Error GoTo Fail

    Dictionary.Prepare

    Assert.IsTrue Dictionary.ColumnExists("column index"), "Prepare should append column index helper column"
    Assert.IsTrue Dictionary.ColumnExists("visibility"), "Prepare should append visibility helper column"
    Assert.IsTrue Dictionary.ColumnExists("crf index"), "Prepare should append CRF index helper column"
    Assert.IsTrue Dictionary.ColumnExists("crf choices"), "Prepare should append CRF choices helper column"
    Assert.IsTrue Dictionary.ColumnExists("crf status"), "Prepare should append CRF status helper column"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestPrepareAddsHelperColumns", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestPrepareRenamesPreservedSheetNames()
    CustomTestSetTitles Assert, "LLdictionary", "TestPrepareRenamesPreservedSheetNames"

    Dim preserved As BetterArray
    Dim sheetNames As BetterArray

    On Error GoTo Fail

    Set preserved = BetterArrayFromList("vlist1D-sheet1", "hlist2D-sheet1")

    Dictionary.Prepare PreservedSheetNames:=preserved

    Set sheetNames = Dictionary.UniqueValues("sheet name")

    Assert.IsTrue sheetNames.Includes("vlist1D-sheet1_"), "Prepare should suffix preserved vertical sheet with underscore"
    Assert.IsTrue sheetNames.Includes("hlist2D-sheet1_"), "Prepare should suffix preserved horizontal sheet with underscore"
    Assert.IsTrue (Not sheetNames.Includes("vlist1D-sheet1")), "Original vertical sheet name should be replaced after prepare"
    Assert.IsTrue (Not sheetNames.Includes("hlist2D-sheet1")), "Original horizontal sheet name should be replaced after prepare"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestPrepareRenamesPreservedSheetNames", Err.Number, Err.Description
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

'@TestMethod("LLdictionary")
Public Sub TestSetTotalNumberOfExportsPersistsName()
    CustomTestSetTitles Assert, "LLdictionary", "TestSetTotalNumberOfExportsPersistsName"

    Dim definition As Name
    Dim sheet As Worksheet

    On Error GoTo Fail

    Set sheet = ThisWorkbook.Worksheets(DICT_SHEET)
    RemoveDictionaryExportName sheet

    Dictionary.TotalNumberOfExports = 37

    Set definition = SheetNameDefinition(sheet, EXPORT_TOTAL_NAME)
    Assert.IsTrue (Not definition Is Nothing), "Expected hidden name to be created on the dictionary sheet"
    Assert.AreEqual 37, NameNumericValue(definition), "Hidden export counter should match configured total"
    Assert.IsFalse definition.Visible, "Hidden export counter must remain invisible"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSetTotalNumberOfExportsPersistsName", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestCreateOverridesStoredExportCounterWhenRequested()
    CustomTestSetTitles Assert, "LLdictionary", "TestCreateOverridesStoredExportCounterWhenRequested"

    Dim sheet As Worksheet
    Dim definition As Name
    Dim created As ILLdictionary

    On Error GoTo Fail

    Set sheet = ThisWorkbook.Worksheets(DICT_SHEET)
    RemoveDictionaryExportName sheet
    sheet.Names.Add Name:=EXPORT_TOTAL_NAME, RefersToR1C1:="=42", Visible:=False

    Set created = LLdictionary.Create(sheet, 1, 1, 35)

    Set definition = SheetNameDefinition(sheet, EXPORT_TOTAL_NAME)
    Assert.IsTrue Not definition Is Nothing, "Hidden counter should remain defined after creation"
    Assert.AreEqual 35, NameNumericValue(definition), "Create should persist the requested total number of exports"
    Assert.AreEqual 35, CLng(created.TotalNumberOfExports), "Returned dictionary should expose the requested export total"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateOverridesStoredExportCounterWhenRequested", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestExportWritesExportCounter()
    CustomTestSetTitles Assert, "LLdictionary", "TestExportWritesExportCounter"

    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim definition As Name

    On Error GoTo Fail

    Dictionary.TotalNumberOfExports = 29
    Set exportBook = NewWorkbook()

    Dictionary.Export exportBook

    Set exportedSheet = exportBook.Worksheets(DICT_SHEET)
    Set definition = SheetNameDefinition(exportedSheet, EXPORT_TOTAL_NAME)

    Assert.IsTrue Not definition Is Nothing, "Export should create hidden name in destination sheet"
    Assert.AreEqual 29, NameNumericValue(definition), "Destination hidden counter should mirror dictionary total"
    Assert.IsFalse definition.Visible, "Exported counter should remain hidden"

    DeleteWorkbook exportBook
    Exit Sub

Fail:
    If Not exportBook Is Nothing Then DeleteWorkbook exportBook
    CustomTestLogFailure Assert, "TestExportWritesExportCounter", Err.Number, Err.Description
End Sub

'@TestMethod("LLdictionary")
Public Sub TestImportRestoresExportCounter()
    CustomTestSetTitles Assert, "LLdictionary", "TestImportRestoresExportCounter"

    Dim exportBook As Workbook
    Dim importedSheet As Worksheet
    Dim definition As Name

    On Error GoTo Fail

    Dictionary.TotalNumberOfExports = 23
    Set exportBook = NewWorkbook()

    Dictionary.Export exportBook

    Set importedSheet = exportBook.Worksheets(DICT_SHEET)
    RemoveDictionaryExportName importedSheet
    importedSheet.Names.Add Name:=EXPORT_TOTAL_NAME, RefersToR1C1:="=11", Visible:=False

    Dictionary.TotalNumberOfExports = 3
    Dictionary.Import importedSheet, 1, 1

    Assert.AreEqual 11, CLng(Dictionary.TotalNumberOfExports), "Import should adopt stored export totals"

    Set definition = SheetNameDefinition(ThisWorkbook.Worksheets(DICT_SHEET), EXPORT_TOTAL_NAME)
    Assert.IsTrue Not (definition Is Nothing), "Dictionary sheet should expose hidden counter after import"
    Assert.AreEqual 11, NameNumericValue(definition), "Dictionary hidden counter should match imported value"

    DeleteWorkbook exportBook
    Exit Sub

Fail:
    If Not exportBook Is Nothing Then DeleteWorkbook exportBook
    CustomTestLogFailure Assert, "TestImportRestoresExportCounter", Err.Number, Err.Description
End Sub

'@section Helpers
'===============================================================================

Private Function SheetNameDefinition(ByVal sheet As Worksheet, ByVal nameId As String) As Name
    Dim definition As Name

    If sheet Is Nothing Then Exit Function

    For Each definition In sheet.Names
        If StrComp(LocalName(definition.Name), nameId, vbTextCompare) = 0 Then
            Set SheetNameDefinition = definition
            Exit Function
        End If
    Next definition
End Function

Private Sub RemoveDictionaryExportName(ByVal sheet As Worksheet)
    Dim definition As Name

    If sheet Is Nothing Then Exit Sub

    Set definition = SheetNameDefinition(sheet, EXPORT_TOTAL_NAME)
    If Not definition Is Nothing Then definition.Delete
End Sub

Private Function LocalName(ByVal qualifiedName As String) As String
    Dim exclPos As Long

    exclPos = InStr(qualifiedName, "!")
    If exclPos = 0 Then
        LocalName = qualifiedName
    Else
        LocalName = Mid$(qualifiedName, exclPos + 1)
    End If
End Function

Private Function NameNumericValue(ByVal definition As Name) As Long
    Dim evaluated As String
    Dim hostWorkbook As Workbook

    If definition Is Nothing Then Exit Function

    On Error Resume Next
        evaluated = Trim$(Replace$(definition.Value, "=", vbNullString))
    On Error GoTo 0

    If LenB(evaluated) <> 0 Then NameNumericValue = CLng(evaluated)
End Function
