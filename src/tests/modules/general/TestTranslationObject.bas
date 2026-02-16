Attribute VB_Name = "TestTranslationObject"
Option Explicit

'@Folder("CustomTests")
'@Folder("Tests")
'@ModuleDescription("Unit tests covering the TranslationObject class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

'@description
'Validates the TranslationObject class, which provides tag-based translation
'backed by a ListObject table. The translation table has a first column of
'tags and subsequent columns for each language. Tests cover the core
'TranslatedValue lookup (matching tag, unknown tag passthrough), formula-aware
'translation where only double-quoted chunks are replaced, TranslateRange for
'cell-by-cell bulk translation, TranslateForm for UserForm controls
'(CommandButton, Label, MultiPage), the Create factory guard against Nothing
'ListObject, ValueExists behaviour when the target language column is missing,
'and LanguagesList header enumeration with and without language columns.
'A fresh three-row translation table (greeting/farewell/status_ok in ENG and
'FRA) is rebuilt in TestInitialize so every test starts from a clean baseline.
'Uses the CustomTest harness (ICustomTest) with CustomTestSetTitles and
'CustomTestLogFailure.
'@depends TranslationObject, ITranslationObject, BetterArray, CustomTest, TestHelpers

Private Assert As ICustomTest
Private TranslationSheet As Worksheet
Private TranslationTable As ListObject
Private Translator As ITranslationObject

Private Const TRANSLATION_SHEET As String = "TST_Translations"
Private Const TRANSLATION_TABLE As String = "TST_TranslationsTable"
Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
'@sub-title Create the test harness and register the module name
Public Sub ModuleInitialize()
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestTranslationObject"
End Sub

'@ModuleCleanup
'@sub-title Print accumulated results and release the harness
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
'@sub-title Prepare a fresh translation worksheet, table, and Translator instance
'@details
'Deletes any leftover TST_Translations sheet, creates a new one, populates a
'three-row ListObject (Tag/ENG/FRA) via PrepareTranslationTable, and builds an
'ENG-targeted Translator from it.
Public Sub TestInitialize()
    TestHelpers.DeleteWorksheets TRANSLATION_SHEET
    Set TranslationSheet = TestHelpers.EnsureWorksheet(TRANSLATION_SHEET)
    PrepareTranslationTable
    Set Translator = TranslationObject.Create(TranslationTable, "ENG")
End Sub

'@TestCleanup
'@sub-title Flush harness output, release objects, and delete the fixture sheet
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set Translator = Nothing
    Set TranslationTable = Nothing
    Set TranslationSheet = Nothing
    TestHelpers.DeleteWorksheets TRANSLATION_SHEET
End Sub


'@section Tests
'===============================================================================
'@TestMethod("TranslationObject")
'@sub-title Verify basic tag lookup, ValueExists, and unknown-tag passthrough
'@details
'Arranges the default ENG translator. Acts by calling TranslatedValue for the
'"greeting" tag. Asserts the returned value equals "Hello". Also verifies that
'ValueExists returns True for the "farewell" tag and that an unknown tag
'("unknown_tag") is returned unchanged, confirming the passthrough behaviour
'for tags not found in the translation table.
Public Sub TestTranslatedValueReturnsMatchingEntry()
    Dim actual As String

    CustomTestSetTitles Assert, "TranslationObject", "TestTranslatedValueReturnsMatchingEntry"

    actual = Translator.TranslatedValue("greeting")
    Assert.AreEqual "Hello", actual, "Expected greeting to translate using the ENG column."
    Assert.IsTrue Translator.ValueExists("farewell"), "ValueExists should report the presence of farewell tag."
    Assert.AreEqual "unknown_tag", Translator.TranslatedValue("unknown_tag"), _
                     "Unknown tags should be returned unchanged."
End Sub

'@TestMethod("TranslationObject")
'@sub-title Verify formula-aware translation replaces only double-quoted chunks
'@details
'Arranges a formula string containing two quoted tags ("greeting" and
'"status_ok") separated by an ampersand concatenation operator. Acts by calling
'TranslatedValue with containsFormula:=True. Asserts that the quoted tags are
'replaced with their ENG translations ("Hello" and "OK") while the unquoted
'formula structure (equals sign, ampersand, spaces) passes through unchanged.
'This confirms the TranslateFormulaText internal parser correctly identifies
'and translates only the double-quoted segments.
Public Sub TestTranslatedValueTranslatesFormulaChunks()
    Dim formulaText As String
    Dim result As String

    CustomTestSetTitles Assert, "TranslationObject", "TestTranslatedValueTranslatesFormulaChunks"

    formulaText = "=" & Chr$(34) & "greeting" & Chr$(34) & " & " & Chr$(34) & "status_ok" & Chr$(34)

    result = Translator.TranslatedValue(formulaText, containsFormula:=True)

    Assert.AreEqual "=" & Chr$(34) & "Hello" & Chr$(34) & " & " & Chr$(34) & "OK" & Chr$(34), _
                     result, _
                     "Only quoted chunks should be translated within formulas."
End Sub

'@TestMethod("TranslationObject")
'@sub-title Verify TranslateRange translates cell values in place
'@details
'Arranges a two-cell column range (E2:E3) on the translation fixture sheet and
'populates it with the tags "greeting" and "farewell". Acts by calling
'TranslateRange on the range. Asserts that cell E2 now contains "Hello" and
'cell E3 contains "Good bye", confirming that TranslateRange iterates through
'each cell and replaces its value with the translated text.
Public Sub TestTranslateRangeTranslatesValues()
    Dim targetRange As Range

    CustomTestSetTitles Assert, "TranslationObject", "TestTranslateRangeTranslatesValues"

    Set targetRange = TranslationSheet.Range("E2:E3")
    targetRange.Cells(1, 1).Value = "greeting"
    targetRange.Cells(2, 1).Value = "farewell"

    Translator.TranslateRange targetRange

    Assert.AreEqual "Hello", CStr(targetRange.Cells(1, 1).Value), "TranslateRange should translate the first cell."
    Assert.AreEqual "Good bye", CStr(targetRange.Cells(2, 1).Value), "TranslateRange should translate the second cell."
End Sub

'@TestMethod("TranslationObject")
'@sub-title Verify TranslateForm translates captions on CommandButton, Label, and MultiPage controls
'@details
'Arranges a draft UserForm via CreateDraftForm and adds three controls: a
'CommandButton named "greeting", a Label named "farewell", and a MultiPage
'with a page named "status_ok". Each control is given an arbitrary old caption.
'Acts by calling TranslateForm on the form. Asserts that the CommandButton
'caption is "Hello", the Label caption is "Good bye", and the MultiPage page
'caption is "OK". This validates that TranslateForm dispatches correctly to
'each supported control type and uses the control Name as the translation tag.
'The form is unloaded after assertions to prevent resource leaks.
Public Sub TestTranslateFormTranslatesSupportedControls()
    Dim testForm As Object
    Dim button As MSForms.CommandButton
    Dim formLabel As MSForms.Label
    Dim multiPage As MSForms.MultiPage

    CustomTestSetTitles Assert, "TranslationObject", "TestTranslateFormTranslatesSupportedControls"

    Set testForm = CreateDraftForm()
    Set button = testForm.Controls.Add("Forms.CommandButton.1", "greeting")
    button.Caption = "old_caption"

    Set formLabel = testForm.Controls.Add("Forms.Label.1", "farewell")
    formLabel.Caption = "old_label"

    Set multiPage = testForm.Controls.Add("Forms.MultiPage.1", "MultiPage1")
    multiPage.Pages.Add
    multiPage.Pages(0).Name = "status_ok"
    multiPage.Pages(0).Caption = "old_page"

    Translator.TranslateForm testForm

    Assert.AreEqual "Hello", button.Caption, "Command button caption should be translated."
    Assert.AreEqual "Good bye", formLabel.Caption, "Label caption should be translated."
    Assert.AreEqual "OK", multiPage.Pages(0).Caption, "MultiPage page captions should be translated."

    Unload testForm
    Set testForm = Nothing
End Sub

'@TestMethod("TranslationObject")
'@sub-title Verify Create factory raises ObjectNotInitialized when ListObject is Nothing
'@details
'Arranges an On Error GoTo handler targeting the ExpectError label. Acts by
'calling TranslationObject.Create with Nothing as the ListObject argument.
'Asserts that execution reaches the error handler and that the raised error
'number matches ProjectError.ObjectNotInitialized. If Create does not raise,
'the test fails explicitly. This confirms the factory guard clause rejects
'uninitialised table arguments.
Public Sub TestCreateRequiresListObject()
    On Error GoTo ExpectError

    CustomTestSetTitles Assert, "TranslationObject", "TestCreateRequiresListObject"

    TranslationObject.Create Nothing, "ENG"
    Assert.Fail "Create should raise when the listobject is missing."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), Err.Number, "Object not initialized error should be raised"
    Err.Clear
End Sub

'@TestMethod("TranslationObject")
'@sub-title Verify ValueExists returns False and TranslatedValue falls back when the language column is absent
'@details
'Arranges a TranslationObject targeting the "DEU" language, which does not
'exist as a column in the fixture table. Acts by calling ValueExists for
'"greeting" and then TranslatedValue for the same tag. Asserts that
'ValueExists returns False and TranslatedValue returns the original tag
'"greeting" unchanged. This confirms the graceful fallback behaviour when
'a requested language column is missing from the translation table.
Public Sub TestValueExistsReturnsFalseWhenLanguageMissing()
    Dim missingLanguage As ITranslationObject
    Dim result As String

    CustomTestSetTitles Assert, "TranslationObject", "TestValueExistsReturnsFalseWhenLanguageMissing"

    Set missingLanguage = TranslationObject.Create(TranslationTable, "DEU")

    Assert.IsFalse missingLanguage.ValueExists("greeting"), _
                   "ValueExists should fail when the language column does not exist."
    result = missingLanguage.TranslatedValue("greeting")
    Assert.AreEqual "greeting", result, _
                     "Missing language should cause TranslatedValue to fall back to the original text."
    Set missingLanguage = Nothing
End Sub

'@TestMethod("TranslationObject")
'@sub-title Verify LanguagesList returns all language column headers in order
'@details
'Arranges the default ENG translator backed by a table with Tag, ENG, and FRA
'columns. Acts by retrieving the LanguagesList BetterArray. Asserts that the
'array length is 2 (excluding the Tag column), that the first item is "ENG",
'and the second item is "FRA". This confirms that LanguagesList correctly
'enumerates language headers while excluding the first helper/tag column.
Public Sub TestLanguagesListReturnsLanguageHeaders()
    Dim languages As BetterArray

    CustomTestSetTitles Assert, "TranslationObject", "TestLanguagesListReturnsLanguageHeaders"

    Set languages = Translator.LanguagesList

    Assert.AreEqual 2&, languages.Length, "LanguagesList should only include translation columns."
    Assert.AreEqual "ENG", CStr(languages.Item(languages.LowerBound)), _
                     "LanguagesList should preserve the header order."
    Assert.AreEqual "FRA", CStr(languages.Item(languages.LowerBound + 1)), _
                     "LanguagesList should capture subsequent language headers."
End Sub

'@TestMethod("TranslationObject")
'@sub-title Verify LanguagesList returns an empty array when no language columns exist
'@details
'Arranges the fixture table by deleting both the FRA and ENG columns, leaving
'only the Tag column. Acts by retrieving the LanguagesList BetterArray from
'the Translator. Asserts that the returned array has a length of 0. This
'validates the edge case where the translation table contains no language
'columns at all, ensuring LanguagesList degrades gracefully to empty.
Public Sub TestLanguagesListReturnsEmptyWhenNoAdditionalColumns()
    Dim languages As BetterArray

    CustomTestSetTitles Assert, "TranslationObject", "TestLanguagesListReturnsEmptyWhenNoAdditionalColumns"

    TranslationTable.ListColumns("FRA").Delete
    TranslationTable.ListColumns("ENG").Delete

    Set languages = Translator.LanguagesList

    Assert.AreEqual 0&, languages.Length, "LanguagesList should be empty when only the helper column remains."
End Sub


'@section Helpers
'===============================================================================

'@sub-title Build the fixture translation ListObject with Tag, ENG, and FRA columns
'@details
'Clears the translation sheet and writes a three-column header row (Tag, ENG,
'FRA) followed by three data rows (greeting/farewell/status_ok with their
'English and French translations). Converts the populated range into a
'ListObject named TST_TranslationsTable and stores it in the module-level
'TranslationTable variable for use by TestInitialize and individual tests.
Private Sub PrepareTranslationTable()
    Dim targetRange As Range

    TranslationSheet.Cells.Clear

    With TranslationSheet
        .Range("A1").Value = "Tag"
        .Range("B1").Value = "ENG"
        .Range("C1").Value = "FRA"
        .Range("A2").Resize(3, 1).Value = Application.Transpose(Array("greeting", "farewell", "status_ok"))
        .Range("B2").Resize(3, 1).Value = Application.Transpose(Array("Hello", "Good bye", "OK"))
        .Range("C2").Resize(3, 1).Value = Application.Transpose(Array("Bonjour", "Au revoir", "D'accord"))
        Set targetRange = .Range("A1").Resize(4, 3)
    End With

    Set TranslationTable = TranslationSheet.ListObjects.Add(xlSrcRange, targetRange, , xlYes)
    TranslationTable.Name = TRANSLATION_TABLE
End Sub

'@sub-title Return an empty draft UserForm for control-translation tests
Private Function CreateDraftForm() As Object
    Set CreateDraftForm = [DraftForm]
End Function
