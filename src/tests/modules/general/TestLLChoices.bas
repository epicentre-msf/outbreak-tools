Attribute VB_Name = "TestLLChoices"

Option Explicit



'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@TestModule
'@Folder("CustomTests")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const CHOICESSHEET As String = "LLChoicesTest"
Private Const CHOICESTRANSLATIONSHEET As String = "LLChoicesTranslation"
Private Const CHOICESTRANSLATIONTABLE As String = "tblLLChoicesTranslation"
Private Const CHOICESTRANSLATIONLANGUAGE As String = "Translated"
Private Const CHOICESIMPORTSHEET As String = "LLChoicesImportSource"

Private Assert As ICustomTest
Private Choices As ILLChoices

'@section Helpers
'===============================================================================

Private Sub ResetChoices()
    Dim previousEventState As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim errHelpFile As String
    Dim errHelpContext As Long

    previousEventState = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanFail

    PrepareChoicesFixture CHOICESSHEET
    Set Choices = LLChoices.Create(ThisWorkbook.Worksheets(CHOICESSHEET), 1, 1)
    Choices.Wksh.Visible = xlSheetVeryHidden 'Keep fixture sheet out of view during UI-bound tests

CleanExit:
    Application.EnableEvents = previousEventState
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    errHelpFile = Err.HelpFile
    errHelpContext = Err.HelpContext
    Application.EnableEvents = previousEventState
    Err.Raise errNumber, errSource, errDescription, errHelpFile, errHelpContext
End Sub

Private Function TranslatorDataRows() As Variant
    TranslatorDataRows = Array( _
        Array("A", "A", "Alpha"), _
        Array("B", "B", "Bravo"), _
        Array("C", "C", "Charlie"), _
        Array("A short", "A short", "A court"), _
        Array("B short", "B short", "B court"), _
        Array("C short", "C short", "C court"), _
        Array("choice 1", "choice 1", "choix 1"), _
        Array("choice 2", "choice 2", "choix 2"), _
        Array("choice 3", "choice 3", "choix 3"), _
        Array("choice 4", "choice 4", "choix 4"), _
        Array("c1", "c1", "c1 fr"), _
        Array("c2", "c2", "c2 fr"), _
        Array("c4", "c4", "c4 fr"))
End Function

Private Function ImportDataRows() As Variant
    ImportDataRows = Array( _
        Array("imported_list", 1, "Label One", "Short One"), _
        Array("imported_list", 2, "Label Two", "Short Two"), _
        Array("imported_second", 1, "Other Label", "Other Short"))
End Function

Private Function CreateChoicesTranslator() As ITranslationObject
    Dim translationSheet As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim translationTable As ListObject

    Set translationSheet = EnsureWorksheet(CHOICESTRANSLATIONSHEET, visibility:= xlSheetVeryhidden)

    headerMatrix = RowsToMatrix(Array(Array("tag", "English", CHOICESTRANSLATIONLANGUAGE)))
    WriteMatrix translationSheet.Cells(1, 1), headerMatrix

    dataMatrix = RowsToMatrix(TranslatorDataRows())
    WriteMatrix translationSheet.Cells(2, 1), dataMatrix

    Set translationTable = translationSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                           Source:=translationSheet.Range("A1").CurrentRegion, _
                                                           XlListObjectHasHeaders:=xlYes)
    translationTable.Name = CHOICESTRANSLATIONTABLE

    Set CreateChoicesTranslator = TranslationObject.Create(translationTable, CHOICESTRANSLATIONLANGUAGE)
End Function

Private Function CreateChoicesImportSheet() As Worksheet
    Dim importSheet As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    Set importSheet = EnsureWorksheet(CHOICESIMPORTSHEET, visibility:=xlSheetVeryhidden)

    headerMatrix = RowsToMatrix(Array(ChoicesFixtureHeaders()))
    WriteMatrix importSheet.Cells(1, 1), headerMatrix

    dataMatrix = RowsToMatrix(ImportDataRows())
    WriteMatrix importSheet.Cells(2, 1), dataMatrix

    Set CreateChoicesImportSheet = importSheet
End Function

Private Sub CleanupChoicesTranslation()
    BusyApp
    DeleteWorksheet CHOICESTRANSLATIONSHEET
End Sub

Private Sub CleanupChoicesImportSource()
    BusyApp
    DeleteWorksheet CHOICESIMPORTSHEET
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLChoices"
    ResetChoices
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    CleanupChoicesTranslation
    CleanupChoicesImportSource
    DeleteWorksheet CHOICESSHEET
    RestoreApp

    Set Choices = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    ResetChoices
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    CleanupChoicesTranslation
    CleanupChoicesImportSource
    Set Choices = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("LLChoices")
Public Sub TestCreateInitialisesChoice()
    CustomTestSetTitles Assert, "LLChoices", "TestCreateInitialisesChoice"
    Assert.IsTrue (TypeName(Choices) = "LLChoices"), "Expected Create to return ILLChoices implementation"
    Assert.AreEqual CHOICESSHEET, Choices.Wksh.Name, "Choice object should target the configured sheet"
    Assert.IsFalse Choices.HasTranslation, "Fixture does not configure translations by default"
End Sub

'@TestMethod("LLChoices")
Public Sub TestAllChoicesReturnsDistinctNames()
    CustomTestSetTitles Assert, "LLChoices", "TestAllChoicesReturnsDistinctNames"
    On Error GoTo Fail

    Dim listValues As BetterArray
    Dim expected As Variant
    Dim index As Long

    expected = ChoicesFixtureDistinctLists()
    Set listValues = Choices.AllChoices

    Assert.AreEqual UBound(expected) - LBound(expected) + 1, listValues.Length, "Unexpected number of lists returned"

    For index = LBound(expected) To UBound(expected)
        Assert.IsTrue listValues.Includes(CStr(expected(index))), "Missing expected list: " & CStr(expected(index))
    Next index
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAllChoicesReturnsDistinctNames", Err.Number, Err.Description
End Sub

'@TestMethod("LLChoices")
Public Sub TestCategoriesHonoursShortLabels()
    CustomTestSetTitles Assert, "LLChoices", "TestCategoriesHonoursShortLabels"
    On Error GoTo Fail

    Dim shortCategories As BetterArray
    Dim expected As Variant
    Dim index As Long
    Dim current As Long

    expected = Array("c1", "c2", "choice 3", "c4")
    Set shortCategories = Choices.Categories("list_multiple", useShortlabels:=True)

    Assert.AreEqual UBound(expected) - LBound(expected) + 1, shortCategories.Length, "Short labels count mismatch"

    current = shortCategories.LowerBound
    For index = LBound(expected) To UBound(expected)
        Assert.AreEqual CStr(expected(index)), CStr(shortCategories.Item(current)), "Unexpected short label at position " & CStr(index - LBound(expected) + 1)
        current = current + 1
    Next index
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCategoriesHonoursShortLabels", Err.Number, Err.Description
End Sub

'@TestMethod("LLChoices")
Public Sub TestSortReordersChoicesByOrdering()
    CustomTestSetTitles Assert, "LLChoices", "TestSortReordersChoicesByOrdering"
    On Error GoTo Fail

    Dim beforeSort As BetterArray
    Dim afterSort As BetterArray
    Dim expected As Variant
    Dim index As Long
    Dim current As Long

    expected = Array("B", "C", "A")

    Set beforeSort = Choices.Categories("list_uncorrect_order")
    Assert.AreEqual "A", CStr(beforeSort.Item(beforeSort.LowerBound)), "Fixture should start unsorted for target list"

    Choices.Sort
    Set afterSort = Choices.Categories("list_uncorrect_order")

    Assert.AreEqual UBound(expected) - LBound(expected) + 1, afterSort.Length, "Sorted list should contain same number of entries"

    current = afterSort.LowerBound
    For index = LBound(expected) To UBound(expected)
        Assert.AreEqual CStr(expected(index)), CStr(afterSort.Item(current)), "Unexpected ordering at index " & CStr(index - LBound(expected) + 1)
        current = current + 1
    Next index
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSortReordersChoicesByOrdering", Err.Number, Err.Description
End Sub

'@TestMethod("LLChoices")
Public Sub TestConcatenateCategoriesUsesDefaultSeparator()
    CustomTestSetTitles Assert, "LLChoices", "TestConcatenateCategoriesUsesDefaultSeparator"
    On Error GoTo Fail

    Dim resultText As String

    resultText = Choices.ConcatenateCategories("list_multiple")
    Assert.AreEqual "choice 1 | choice 2 | choice 3 | choice 4", resultText, "ConcatenateCategories should respect default separator: ' | '"

    resultText = Choices.ConcatenateCategories("list_multiple", sep:=" -- ")
    Assert.AreEqual "choice 1 -- choice 2 -- choice 3 -- choice 4", resultText, "ConcatenateCategories should respect provided separator"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestConcatenateCategoriesUsesDefaultSeparator", Err.Number, Err.Description
End Sub

'@TestMethod("LLChoices")
Public Sub TestExportCreatesHiddenCopy()
    CustomTestSetTitles Assert, "LLChoices", "TestExportCreatesHiddenCopy"
    On Error GoTo Fail

    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim headers As Variant
    Dim index As Long
    Dim expectedLastRow As Long

    Set exportBook = NewWorkbook()
    Choices.Export exportBook

    Set exportedSheet = exportBook.Worksheets(CHOICESSHEET)
    headers = ChoicesFixtureHeaders()

    For index = LBound(headers) To UBound(headers)
        Assert.AreEqual CStr(headers(index)), CStr(exportedSheet.Cells(1, index - LBound(headers) + 1).Value), "Export should preserve header order"
    Next index

    Assert.AreEqual "list_correct_order", CStr(exportedSheet.Cells(2, 1).Value), "Export should carry first data row"
    Assert.AreEqual "A", CStr(exportedSheet.Cells(2, 3).Value), "Export should include label column"
    Assert.AreEqual xlSheetHidden, exportedSheet.Visible, "Export should hide destination sheet by default"

    expectedLastRow = ChoicesFixtureRowCount() + 1
    Assert.AreEqual expectedLastRow, exportedSheet.Cells(exportedSheet.Rows.Count, 1).End(xlUp).Row, "Export should only include fixture rows"

    DeleteWorkbook exportBook
    Exit Sub

Fail:
    If Not exportBook Is Nothing Then DeleteWorkbook exportBook
    CustomTestLogFailure Assert, "TestExportCreatesHiddenCopy", Err.Number, Err.Description
End Sub

'@TestMethod("LLChoices")
Public Sub TestAddChoiceAddsNewEntries()
    CustomTestSetTitles Assert, "LLChoices", "TestAddChoiceAddsNewEntries"
    On Error GoTo Fail

    Dim longLabels As BetterArray
    Dim shortLabels As BetterArray
    Dim categories As BetterArray

    Set longLabels = BetterArrayFromList("North", "South")
    Set shortLabels = BetterArrayFromList("N", "S")

    Choices.AddChoice "geo_region", longLabels, shortLabels

    Assert.IsTrue Choices.ChoiceExists("geo_region"), "New choice should exist after AddChoice"

    Set categories = Choices.Categories("geo_region", useShortlabels:=True)
    Assert.AreEqual longLabels.Length, categories.Length, "Short categories should match number of provided labels"
    Assert.IsTrue categories.Includes("N"), "Expected short labels to be stored"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddChoiceAddsNewEntries", Err.Number, Err.Description
End Sub

'@TestMethod("LLChoices")
Public Sub TestRemoveChoiceDeletesRequestedList()
    CustomTestSetTitles Assert, "LLChoices", "TestRemoveChoiceDeletesRequestedList"
    On Error GoTo Fail

    Assert.IsTrue Choices.ChoiceExists("list_multiple"), "Precondition failed: fixture should contain list_multiple"
    Choices.RemoveChoice "list_multiple"
    Assert.IsFalse Choices.ChoiceExists("list_multiple"), "RemoveChoice should delete all occurrences of the list"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveChoiceDeletesRequestedList", Err.Number, Err.Description
End Sub

'@TestMethod("LLChoices")
Public Sub TestImportReplacesChoices()
    CustomTestSetTitles Assert, "LLChoices", "TestImportReplacesChoices"
    On Error GoTo Fail

    Dim importSheet As Worksheet
    Dim lists As BetterArray

    Set importSheet = CreateChoicesImportSheet()

    Choices.Import importSheet, 1, 1

    Assert.AreEqual "imported_list", CStr(Choices.Wksh.Cells(2, 1).Value), "Import should replace first list name"
    Assert.AreEqual "Label One", CStr(Choices.Wksh.Cells(2, 3).Value), "Import should copy label column"

    Set lists = Choices.AllChoices
    Assert.IsTrue (lists.Length = 2), "Import should reset unique list names"
    Assert.IsTrue lists.Includes("imported_list"), "Imported list should be present"
    Assert.IsTrue lists.Includes("imported_second"), "Second imported list should be present"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportReplacesChoices", Err.Number, Err.Description
End Sub

'@TestMethod("LLChoices")
Public Sub TestTranslateUpdatesLabels()
    CustomTestSetTitles Assert, "LLChoices", "TestTranslateUpdatesLabels"
    On Error GoTo Fail

    Dim translator As ITranslationObject
    Dim hostSheet As Worksheet

    Set translator = CreateChoicesTranslator()
    Choices.Translate translator

    Set hostSheet = Choices.Wksh
    Assert.AreEqual "Alpha", CStr(hostSheet.Cells(2, 3).Value), "Label should be translated"
    Assert.AreEqual "A court", CStr(hostSheet.Cells(2, 4).Value), "Short label should follow translation"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTranslateUpdatesLabels", Err.Number, Err.Description
End Sub
