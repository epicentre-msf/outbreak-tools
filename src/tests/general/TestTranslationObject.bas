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
'Uses the CustomTest harness (CustomTest) with CustomTestSetTitles and
'CustomTestLogFailure.
'@depends TranslationObject, TranslationObject, BetterArray, CustomTest, TestHelpersLite

Private Assert As CustomTest
Private TranslationSheet As Worksheet
Private TranslationTable As ListObject
Private Translator As TranslationObject

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
    DeleteWorksheets TRANSLATION_SHEET
    Set TranslationSheet = EnsureWorksheet(TRANSLATION_SHEET)
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
    DeleteWorksheets TRANSLATION_SHEET
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
    Dim missingLanguage As TranslationObject
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


'@TestMethod("TranslationObject")
'@sub-title Verify comparison operators survive a formula translation into French
'@details
'Arranges the fixture table with the six age-group labels of a real
'CHOICE_FORMULA and their French translations, then builds a FRA translator.
'Acts by calling TranslatedValue with containsFormula:=True on that formula.
'Asserts the five "<" characters and the one ">" character are all still there
'afterwards, and that the quoted labels came back in French. The formula text
'is the one reported from the field, where the age bands read "< 6", "< 9",
'"< 12", "< 5" and "< 15" and the upper band reads ">=15".
Public Sub TestFormulaTranslationKeepsComparisonOperators()
    Dim formulaText As String
    Dim result As String
    Dim frenchTranslator As TranslationObject

    CustomTestSetTitles Assert, "TranslationObject", "TestFormulaTranslationKeepsComparisonOperators"

    PrepareAgeGroupRows
    Set frenchTranslator = TranslationObject.Create(TranslationTable, "FRA")

    formulaText = "CHOICE_FORMULA(list_age_group,age_years=" & Chr$(34) & Chr$(34) & "," & Chr$(34) & Chr$(34) & _
                  ", age_months < 6, " & Chr$(34) & "0 - 5 months" & Chr$(34) & _
                  ", age_months < 9, " & Chr$(34) & "6 - 8 months" & Chr$(34) & _
                  ", age_months < 12, " & Chr$(34) & "9 - 11 months" & Chr$(34) & _
                  ", age_years < 5, " & Chr$(34) & "1 - 4 years" & Chr$(34) & _
                  ", age_years < 15, " & Chr$(34) & "5 - 14 years" & Chr$(34) & _
                  ", AND(age_years>=15, age_years 120) , " & Chr$(34) & "15+ years" & Chr$(34) & ")"

    result = frenchTranslator.TranslatedValue(formulaText, containsFormula:=True)

    Assert.AreEqual 5&, OccurrenceCount(result, "<"), _
                     "Every '<' comparison operator should survive the translation."
    Assert.AreEqual 1&, OccurrenceCount(result, ">"), _
                     "The '>' comparison operator should survive the translation."
    Assert.IsTrue InStr(1, result, Chr$(34) & "0 - 5 mois" & Chr$(34), vbBinaryCompare) > 0, _
                   "The first age band label should come back in French."
    Assert.IsTrue InStr(1, result, Chr$(34) & "15+ ans" & Chr$(34), vbBinaryCompare) > 0, _
                   "The last age band label should come back in French."

    Set frenchTranslator = Nothing
End Sub


'@TestMethod("TranslationObject")
'@sub-title Verify text with no entry in the table comes back character for character
'@details
'Arranges the default ENG translator and three tags that are not in the fixture
'table, each carrying a character the cleaner used to remove: a "<", a ">" and
'both together. Acts by calling TranslatedValue on each. Asserts every one is
'returned exactly as it went in. The miss path ends in the same cleaner the
'formula path uses, so an untranslated cell was altered too, which is the wider
'half of the same defect.
'Each comparison is made on CodePoints rather than on the text, for the reason
'given on that function.
Public Sub TestUntranslatedTextIsReturnedUnchanged()
    CustomTestSetTitles Assert, "TranslationObject", "TestUntranslatedTextIsReturnedUnchanged"

    AssertSameCharacters "age_months < 6", Translator.TranslatedValue("age_months < 6"), _
                         "An untranslated tag holding '<' should come back unchanged."
    AssertSameCharacters "age_years >= 15", Translator.TranslatedValue("age_years >= 15"), _
                         "An untranslated tag holding '>' should come back unchanged."
    AssertSameCharacters "5 < x > 2", Translator.TranslatedValue("5 < x > 2"), _
                         "An untranslated tag holding both operators should come back unchanged."
End Sub

'@TestMethod("TranslationObject")
'@sub-title Verify the cleaner still turns a non-breaking space into an ordinary one
'@details
'Arranges a tag whose words are joined by ChrW$(160), U+00A0, the character the
'cleaner is there to remove. Acts by calling TranslatedValue on it. Asserts the
'result reads "no such tag" joined by code 32, and that no code 160 is left.
'
'The expected value is never written as a string holding ordinary spaces and
'compared against a result that may hold non-breaking ones. That comparison is
'the one the harness cannot make: Assert.AreEqual reported success on exactly
'this pair while two code 160 characters were still in the result. Every
'assertion here is on CodePoints instead, so a 160 can never read as a 32.
'ChrW$ is used rather than Chr$ because Chr$ reads 160 as a byte in the
'machine's ANSI codepage, which on this Mac is the dagger.
Public Sub TestNonBreakingSpaceBecomesOrdinarySpace()
    Dim result As String

    CustomTestSetTitles Assert, "TranslationObject", "TestNonBreakingSpaceBecomesOrdinarySpace"

    result = Translator.TranslatedValue("no" & ChrW$(160) & "such" & ChrW$(160) & "tag")

    AssertSameCharacters "no such tag", result, _
                         "A non-breaking space should be replaced by an ordinary space."
    Assert.AreEqual 0&, OccurrenceCount(result, ChrW$(160)), _
                     "No non-breaking space should survive the cleaner."
    Assert.AreEqual 2&, OccurrenceCount(result, ChrW$(32)), _
                     "Both gaps should be an ordinary space."
End Sub


'@TestMethod("TranslationObject")
'@sub-title Verify a table with no data row translates instead of raising
'@details
'Arranges a second table on the fixture sheet and deletes its one data row,
'so it carries no body range at all. Acts by translating a tag through it and
'by asking whether that tag exists. Asserts the tag comes back unchanged and
'is reported absent, and that neither call raises.
'
'The signature of the copy in memory reads the address of the body range.
'Written as an IIf, that address was read even when the body range was
'Nothing and raised 91, "object variable not set", out of every public call
'of this class. A worksheet function calling through here then answered
'#VALUE! on the sheet.
'
'The row goes through ListRows.Delete, not through Resize onto the header row
'alone: Resize is refused on a sheet carrying another table and takes the run
'into the debugger, where a hidden Excel sits until it is killed.
'
'The handler is what the rest of this module does without. A raise here is a
'failure row; unhandled, it breaks the whole run on a headless Excel.
Public Sub TestEmptyTableTranslatesWithoutRaising()
    Dim emptyTable As ListObject
    Dim emptyTranslator As TranslationObject

    CustomTestSetTitles Assert, "TranslationObject", "TestEmptyTableTranslatesWithoutRaising"

    On Error GoTo Fail

    TranslationSheet.Range("E1:F1").Value = Array("Tag", "ENG")
    TranslationSheet.Range("E2:F2").Value = Array("greeting", "Hello")

    Set emptyTable = TranslationSheet.ListObjects.Add(xlSrcRange, TranslationSheet.Range("E1:F2"), , xlYes)
    'A table left with its header row alone carries no body range, which is
    'the state a translations sheet is in before its first tag is written.
    emptyTable.ListRows(1).Delete

    Assert.IsTrue emptyTable.DataBodyRange Is Nothing, "The fixture table should carry no body range."

    Set emptyTranslator = TranslationObject.Create(emptyTable, "ENG")

    Assert.AreEqual "greeting", emptyTranslator.TranslatedValue("greeting"), _
                     "A table with no data row should answer the tag unchanged."
    Assert.IsFalse emptyTranslator.ValueExists("greeting"), _
                     "A table with no data row holds no tag."
    Assert.AreEqual 1&, emptyTranslator.LanguagesList.Length, _
                     "The one language column should still be listed."

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEmptyTableTranslatesWithoutRaising", Err.Number, Err.Description
End Sub

'@TestMethod("TranslationObject")
'@sub-title Verify a table carrying only its tag column lists no language
'@details
'Arranges a one column table with a value typed in the cell just right of it.
'Acts by reading LanguagesList. Asserts the list is empty: the header block
'used to be read as a plain Offset of the header row, which reached that one
'cell past the table and answered whatever the sheet held there.
Public Sub TestLanguagesListIgnoresTheCellPastTheTable()
    Dim tagOnlyTable As ListObject
    Dim tagOnlyTranslator As TranslationObject

    CustomTestSetTitles Assert, "TranslationObject", "TestLanguagesListIgnoresTheCellPastTheTable"

    On Error GoTo Fail

    TranslationSheet.Range("H1").Value = "Tag"
    TranslationSheet.Range("H2").Value = "greeting"
    TranslationSheet.Range("I1").Value = "NotALanguage"

    Set tagOnlyTable = TranslationSheet.ListObjects.Add(xlSrcRange, TranslationSheet.Range("H1:H2"), , xlYes)
    Set tagOnlyTranslator = TranslationObject.Create(tagOnlyTable, "ENG")

    Assert.AreEqual 0&, tagOnlyTranslator.LanguagesList.Length, _
                     "A table carrying only its tag column should list no language."

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLanguagesListIgnoresTheCellPastTheTable", Err.Number, Err.Description
End Sub


'@section Helpers
'===============================================================================

'@sub-title Add the age-group labels of the probe formula to the fixture table
'@details
'Writes six extra tag rows below the three the fixture already holds, each with
'its English label in the ENG column and its French label in the FRA column.
'The table grows, which changes its signature, so the next translation call
'reads it again on its own.
Private Sub PrepareAgeGroupRows()
    Dim tags As Variant
    Dim english As Variant
    Dim french As Variant
    Dim firstRow As Long

    tags = Array("0 - 5 months", "6 - 8 months", "9 - 11 months", _
                 "1 - 4 years", "5 - 14 years", "15+ years")
    english = tags
    french = Array("0 - 5 mois", "6 - 8 mois", "9 - 11 mois", _
                   "1 - 4 ans", "5 - 14 ans", "15+ ans")

    firstRow = TranslationTable.Range.Row + TranslationTable.Range.Rows.Count

    With TranslationSheet
        .Cells(firstRow, 1).Resize(6, 1).Value = Application.Transpose(tags)
        .Cells(firstRow, 2).Resize(6, 1).Value = Application.Transpose(english)
        .Cells(firstRow, 3).Resize(6, 1).Value = Application.Transpose(french)
    End With

    TranslationTable.Resize TranslationSheet.Cells(TranslationTable.Range.Row, 1).Resize(10, 3)
End Sub

'@sub-title Assert two strings hold the same characters, code point by code point
'@details
'Compares CodePoints of each side rather than the strings themselves, and puts
'the readable text in the message so a failure is still legible.
'@param expected String. The text that should have come back.
'@param actual String. The text that did.
'@param message String. What the assertion is pinning.
Private Sub AssertSameCharacters(ByVal expected As String, _
                                 ByVal actual As String, _
                                 ByVal message As String)

    Assert.AreEqual CodePoints(expected), CodePoints(actual), _
                     message & " (expected text: '" & expected & "', actual text: '" & actual & "')"
End Sub

'@sub-title Spell a string out as its Unicode code points, comma separated
'@details
'A whitespace character cannot be compared through Assert.AreEqual: it reported
'success on "no such tag" against the same words joined by ChrW$(160), because
'the comparison behind it does not separate a non-breaking space from an
'ordinary one. Turning each character into its number first removes the question
'-- 160 and 32 are different numbers, and the assertion is then made between two
'strings of digits and commas, which hold no whitespace to be lenient about.
'AscW is used rather than Asc so the answer is the code point and not a byte in
'the machine's ANSI codepage.
'@param text String. The text to spell out.
'@return String. One number per character, comma separated.
Private Function CodePoints(ByVal text As String) As String
    Dim codes() As String
    Dim idx As Long
    Dim textLength As Long

    textLength = Len(text)
    If textLength = 0 Then Exit Function

    ReDim codes(1 To textLength)

    For idx = 1 To textLength
        codes(idx) = CStr(AscW(Mid$(text, idx, 1)))
    Next idx

    CodePoints = Join(codes, ",")
End Function

'@sub-title Count how many times one character appears in a string
'@param haystack String. The text to scan.
'@param needle String. The character to count.
'@return Long. The number of occurrences.
Private Function OccurrenceCount(ByVal haystack As String, ByVal needle As String) As Long
    Dim position As Long

    If LenB(haystack) = 0 Or LenB(needle) = 0 Then Exit Function

    position = InStr(1, haystack, needle, vbBinaryCompare)
    Do While position > 0
        OccurrenceCount = OccurrenceCount + 1
        position = InStr(position + Len(needle), haystack, needle, vbBinaryCompare)
    Loop
End Function

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
