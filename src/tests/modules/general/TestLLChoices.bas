Attribute VB_Name = "TestLLChoices"
Attribute VB_Description = "Tests for LLChoices class"
Option Explicit

'===============================================================================
' @ModuleDescription Tests for the LLChoices class, which manages linelist
'   choice lists (dropdown options) stored on a dedicated worksheet. Each choice
'   list is identified by a name and contains ordered rows of labels and short
'   labels that drive data-validation dropdowns in the linelist.
'
' @description This module exercises the full public surface of the ILLChoices
'   interface: factory construction via LLChoices.Create, enumeration of distinct
'   list names (AllChoices), retrieval of categories with optional short-label
'   substitution, sort-by-ordering-column, concatenation with configurable
'   separators, export to an external workbook as a hidden sheet, CRUD
'   operations (AddChoice, RemoveChoice, ChoiceExists), ListObject row
'   manipulation (InsertRows, DeleteRows), bulk import from an external
'   worksheet, and label translation through an ITranslationObject. The module
'   also validates the edge case of requesting categories for a nonexistent
'   choice list.
'
' @depends LLChoices, ILLChoices, TranslationObject, ITranslationObject,
'   BetterArray, CustomTest, TestHelpers, ChoicesTestFixture
'
' Test fixture data is supplied by ChoicesTestFixture.bas, which provides three
' named lists: "list_correct_order" (A/B/C in order 1-2-3),
' "list_uncorrect_order" (A/B/C in order 3-1-2), and "list_multiple" (choice
' 1-4 with some missing short labels). Translation and import fixtures are
' built on-the-fly by private helpers within this module.
'===============================================================================

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLChoices class")

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

' @sub-title Reset the module-level Choices object to a clean fixture state.
' @details Disables Application.EnableEvents while rebuilding the fixture sheet
'   via PrepareChoicesFixture, then creates a fresh LLChoices instance targeted
'   at that sheet. The sheet is hidden to prevent visual flicker during tests.
'   On error the original EnableEvents state is restored before re-raising.
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
    Choices.Wksh.Visible = xlSheetHidden 'Keep fixture sheet out of view during UI-bound tests

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

' @sub-title Create (or recreate) a ListObject on the choices sheet from the
'   current data region.
' @details Deletes any existing ListObject on the sheet, converts the
'   CurrentRegion starting at A1 into a new ListObject named "tblLLChoices",
'   and returns it. This helper is used by tests that exercise row-level
'   ListObject operations (InsertRows, DeleteRows).
Private Function EnsureChoicesListObject() As ListObject
    Dim choiceSheet As Worksheet
    Dim dataRange As Range
    Dim lo As ListObject

    Set choiceSheet = Choices.Wksh

    On Error Resume Next
        choiceSheet.ListObjects(1).Delete
    On Error GoTo 0

    Set dataRange = choiceSheet.Range("A1").CurrentRegion
    Set lo = choiceSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                         Source:=dataRange, _
                                         XlListObjectHasHeaders:=xlYes)
    lo.Name = "tblLLChoices"

    Set EnsureChoicesListObject = lo
End Function

' @sub-title Provide translation mapping rows for the choices fixture.
' @details Returns a Variant array of arrays, each containing (tag, English,
'   Translated) triples. Covers both long and short labels for the three fixture
'   lists, enabling TestTranslateUpdatesLabels to verify that Choices.Translate
'   replaces labels with their translated counterparts.
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

' @sub-title Provide import source rows for the import fixture.
' @details Returns a Variant array of arrays matching the choices header layout
'   (list name, ordering, label, short label). Contains two distinct list names
'   ("imported_list" with one row, "imported_second" with one row) used by
'   TestImportReplacesChoices to verify that Choices.Import completely replaces
'   existing data.
Private Function ImportDataRows() As Variant
    ImportDataRows = Array( _
        Array("imported_list", 1, "Label One", "Short One"), _
        Array("imported_second", 1, "Other Label", "Other Short"))
End Function

' @sub-title Build a translation table and return an ITranslationObject for it.
' @details Creates (or resets) a hidden worksheet named CHOICESTRANSLATIONSHEET,
'   writes header and data rows from TranslatorDataRows, wraps the region in a
'   ListObject named CHOICESTRANSLATIONTABLE, and returns a TranslationObject
'   targeting the CHOICESTRANSLATIONLANGUAGE column. The sheet is cleaned up by
'   CleanupChoicesTranslation in TestCleanup.
Private Function CreateChoicesTranslator() As ITranslationObject
    Dim translationSheet As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim translationTable As ListObject

    Set translationSheet = EnsureWorksheet(CHOICESTRANSLATIONSHEET, visibility:= xlSheetHidden)

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

' @sub-title Build a worksheet that acts as an import source for Choices.Import.
' @details Creates a hidden worksheet named CHOICESIMPORTSHEET, writes the
'   standard choices headers (from ChoicesFixtureHeaders) plus import data rows
'   (from ImportDataRows), and returns the sheet. Cleaned up by
'   CleanupChoicesImportSource in TestCleanup.
Private Function CreateChoicesImportSheet() As Worksheet
    Dim importSheet As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    Set importSheet = EnsureWorksheet(CHOICESIMPORTSHEET, visibility:=xlSheetHidden)

    headerMatrix = RowsToMatrix(Array(ChoicesFixtureHeaders()))
    WriteMatrix importSheet.Cells(1, 1), headerMatrix

    dataMatrix = RowsToMatrix(ImportDataRows())
    WriteMatrix importSheet.Cells(2, 1), dataMatrix

    Set CreateChoicesImportSheet = importSheet
End Function

' @sub-title Delete the translation fixture worksheet if it exists.
Private Sub CleanupChoicesTranslation()
    BusyApp
    DeleteWorksheet CHOICESTRANSLATIONSHEET
End Sub

' @sub-title Delete the import source fixture worksheet if it exists.
Private Sub CleanupChoicesImportSource()
    BusyApp
    DeleteWorksheet CHOICESIMPORTSHEET
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
' @sub-title One-time setup before any test in this module runs.
' @details Ensures the test output sheet exists, creates the CustomTest assert
'   object, registers the module name for reporting, and performs an initial
'   ResetChoices to prepare the fixture sheet.
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLChoices"
    ResetChoices
End Sub

'@ModuleCleanup
' @sub-title One-time teardown after all tests in this module have run.
' @details Prints accumulated test results, removes all fixture worksheets
'   (translation, import, choices), restores application state, and releases
'   object references.
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
' @sub-title Per-test setup: rebuild the fixture from scratch.
' @details Calls BusyApp to suppress screen updates, then ResetChoices to
'   ensure each test starts with an identical, unmodified fixture.
Private Sub TestInitialize()
    BusyApp
    ResetChoices
End Sub

'@TestCleanup
' @sub-title Per-test teardown: flush assertions and remove ephemeral sheets.
' @details Flushes any pending assertion output, removes translation and import
'   sheets that may have been created during the test, and releases the Choices
'   reference.
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
' @sub-title Verify that LLChoices.Create returns a properly initialised object.
' @details Asserts that the module-level Choices variable, created during
'   TestInitialize via LLChoices.Create, resolves to the concrete LLChoices
'   type and that its Wksh property points to the expected fixture sheet. This
'   is a smoke test that validates the factory pattern wiring before any
'   behavioural tests run.
Public Sub TestCreateInitialisesChoice()
    CustomTestSetTitles Assert, "LLChoices", "TestCreateInitialisesChoice"
    Assert.IsTrue (TypeName(Choices) = "LLChoices"), "Expected Create to return ILLChoices implementation"
    Assert.AreEqual CHOICESSHEET, Choices.Wksh.Name, "Choice object should target the configured sheet"
End Sub

'@TestMethod("LLChoices")
' @sub-title Verify that AllChoices returns every distinct list name from the
'   fixture exactly once.
' @details Arranges the expected distinct names from ChoicesFixtureDistinctLists
'   (list_correct_order, list_uncorrect_order, list_multiple). Acts by calling
'   Choices.AllChoices which returns a BetterArray. Asserts that the count
'   matches and that every expected name is included, confirming correct
'   deduplication across multiple fixture rows sharing the same list name.
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
' @sub-title Verify that Categories with useShortlabels:=True returns short
'   labels, falling back to long labels when no short label is present.
' @details Targets the "list_multiple" fixture list whose third entry has an
'   empty short label. The expected result is ("c1", "c2", "choice 3", "c4"),
'   where "choice 3" is the long-label fallback for the missing short label.
'   Asserts both the count and the positional content of every element to
'   confirm short-label substitution logic.
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
' @sub-title Verify that Sort reorders entries within each list by the ordering
'   column.
' @details Uses the "list_uncorrect_order" fixture list where rows are stored in
'   order (A=3, B=1, C=2). First confirms the unsorted state starts with "A",
'   then calls Choices.Sort and retrieves categories again. The expected post-
'   sort order is ("B", "C", "A") corresponding to ordering values 1, 2, 3.
'   Asserts both the count and exact positional sequence.
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
' @sub-title Verify that ConcatenateCategories joins labels with the default
'   separator and accepts a custom separator.
' @details Calls ConcatenateCategories on "list_multiple" without a separator
'   argument and asserts the result equals "choice 1 | choice 2 | choice 3 |
'   choice 4", confirming the default separator is " | ". Then calls it again
'   with sep:=" -- " and asserts the custom separator is honoured. This covers
'   both the default and explicit-separator code paths.
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
' @sub-title Verify that Export copies the choices sheet into an external
'   workbook as a hidden sheet with intact headers and data.
' @details Creates a new blank workbook, calls Choices.Export into it, and then
'   asserts that the exported sheet exists with the expected name, that all
'   fixture headers appear in order, that representative data cells match, that
'   the sheet visibility is xlSheetHidden, and that the total row count matches
'   the fixture. The temporary workbook is deleted after assertions. The Fail
'   handler also deletes the workbook to avoid leaked temp files.
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
' @sub-title Verify that AddChoice inserts a new named list with both long and
'   short labels.
' @details Arranges two BetterArrays of labels ("North"/"South") and short
'   labels ("N"/"S"), then calls Choices.AddChoice with the name "geo_region".
'   Asserts that ChoiceExists returns True for the new name, that the short
'   categories count matches, and that the short labels include "N". This
'   confirms end-to-end row insertion and in-memory cache invalidation.
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
' @sub-title Verify that RemoveChoice deletes all rows belonging to a named
'   list.
' @details First asserts that the fixture contains "list_multiple" as a
'   precondition, then calls Choices.RemoveChoice. Asserts that ChoiceExists
'   returns False afterward, confirming that every row for that list name was
'   removed from the underlying worksheet.
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
' @sub-title Verify that InsertRows adds blank rows matching the height of the
'   selection range while shifting existing data downward.
' @details Wraps the fixture in a ListObject via EnsureChoicesListObject,
'   captures the initial row count and the value in the second data row, then
'   builds a two-row selection range starting at row 2. After calling
'   Choices.InsertRows, asserts that the row count increased by 2, the first
'   inserted row is blank, and the previously second row shifted down to row 3
'   with its value preserved.
Public Sub TestInsertRowsMirrorsSelectionHeight()
    CustomTestSetTitles Assert, "LLChoices", "TestInsertRowsMirrorsSelectionHeight"
    On Error GoTo Fail

    Dim lo As ListObject
    Dim selectionRange As Range
    Dim initialRows As Long
    Dim preservedValue As String

    Set lo = EnsureChoicesListObject()
    preservedValue = CStr(lo.DataBodyRange.Cells(2, 1).Value)
    initialRows = lo.ListRows.Count

    Set selectionRange = lo.ListRows(2).Range
    Set selectionRange = selectionRange.Resize(2, lo.ListColumns.Count)

    Choices.InsertRows selectionRange

    Assert.AreEqual initialRows + 2, lo.ListRows.Count, _
        "InsertRows should add rows matching the selection height"
    Assert.AreEqual vbNullString, CStr(lo.ListRows(2).Range.Cells(1, 1).Value), _
        "First inserted row should be blank"
    Assert.AreEqual preservedValue, CStr(lo.ListRows(3).Range.Cells(1, 1).Value), _
        "Existing data should shift below inserted rows"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsMirrorsSelectionHeight", Err.Number, Err.Description
End Sub

'@TestMethod("LLChoices")
' @sub-title Verify that DeleteRows removes the selected ListObject rows.
' @details Wraps the fixture in a ListObject, records the baseline row count,
'   targets the second data row for deletion, and calls Choices.DeleteRows.
'   Asserts that the row count decreased by exactly one, confirming that only
'   the targeted row was removed without side effects on adjacent rows.
Public Sub TestDeleteRowsRemovesSelection()
    CustomTestSetTitles Assert, "LLChoices", "TestDeleteRowsRemovesSelection"
    On Error GoTo Fail

    Dim lo As ListObject
    Dim selectionRange As Range
    Dim baseline As Long

    Set lo = EnsureChoicesListObject()
    baseline = lo.ListRows.Count

    Set selectionRange = lo.ListRows(2).Range
    Choices.DeleteRows selectionRange

    Assert.AreEqual baseline - 1, lo.ListRows.Count, _
                     "DeleteRows should remove the targeted choice rows"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDeleteRowsRemovesSelection", Err.Number, Err.Description
End Sub

'@TestMethod("LLChoices")
' @sub-title Verify that Categories returns an empty BetterArray when the
'   requested choice name does not exist.
' @details Calls Choices.Categories with a name ("missing_choice") that is not
'   present in the fixture data. Asserts that the returned BetterArray has
'   Length 0, confirming graceful handling of nonexistent list names rather than
'   raising an error.
Public Sub TestCategoriesReturnEmptyForMissingChoice()
    CustomTestSetTitles Assert, "LLChoices", "TestCategoriesReturnEmptyForMissingChoice"
    On Error GoTo Fail

    Dim categories As BetterArray
    Set categories = Choices.Categories("missing_choice")

    Assert.AreEqual 0, categories.Length, "Missing choice should return an empty BetterArray"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCategoriesReturnEmptyForMissingChoice", Err.Number, Err.Description
End Sub


'@TestMethod("LLChoices")
' @sub-title Verify that Import fully replaces existing choices with data from
'   an external worksheet.
' @details Arranges a hidden import worksheet via CreateChoicesImportSheet
'   containing two list names ("imported_list", "imported_second"). Calls
'   Choices.Import and asserts that the underlying sheet now starts with the
'   imported data, that AllChoices returns exactly two entries, and that both
'   imported list names are present. This confirms that Import is destructive:
'   all prior fixture data is replaced.
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
' @sub-title Verify that Translate updates both labels and short labels using
'   an ITranslationObject.
' @details Arranges a translation table mapping "A" to "Alpha" and "A short" to
'   "A court" (among others) via CreateChoicesTranslator. Calls
'   Choices.Translate and reads the underlying sheet cells directly. Asserts
'   that the label column now contains "Alpha" and the short-label column
'   contains "A court" for the first data row, confirming that translation
'   applies to both label types in-place on the worksheet.
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
