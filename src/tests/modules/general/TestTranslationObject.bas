Attribute VB_Name = "TestTranslationObject"
Option Explicit

'@Folder("CustomTests")
'@Folder("Tests")
'@ModuleDescription("Unit tests covering the TranslationObject class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

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
Public Sub ModuleInitialize()
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestTranslationObject"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    TestHelpers.DeleteWorksheets TRANSLATION_SHEET
    Set TranslationSheet = TestHelpers.EnsureWorksheet(TRANSLATION_SHEET)
    PrepareTranslationTable
    Set Translator = TranslationObject.Create(TranslationTable, "ENG")
End Sub

'@TestCleanup
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


'@section Helpers
'===============================================================================
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

Private Function CreateDraftForm() As Object
    Set CreateDraftForm = [DraftForm]
End Function
