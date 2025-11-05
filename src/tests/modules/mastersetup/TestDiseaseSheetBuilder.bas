Attribute VB_Name = "TestDiseaseSheetBuilder"
Attribute VB_Description = "Tests ensuring DiseaseSheetBuilder creates worksheets with headers, validations, and tables"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests ensuring DiseaseSheetBuilder creates worksheets with headers, validations, and tables")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const ANCHOR_SHEET As String = "Variables"
Private Const DROPDOWN_SHEET As String = "DropdownStubSheet"
Private Const TRANSLATION_SHEET As String = "SheetBuilderTranslations"

Private Assert As ICustomTest
Private Builder As IDiseaseSheetBuilder
Private Dropdowns As IDropdownLists
Private Translations As ITranslationObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseSheetBuilder"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        CleanupEnvironment
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
    Set Builder = Nothing
    Set Dropdowns = Nothing
    Set Translations = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    CleanupEnvironment
    PrepareEnvironment
End Sub

'@TestCleanup
Private Sub TestCleanup()
    CleanupEnvironment
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseSheetBuilder")
Public Sub TestBuildCreatesWorksheet()
    CustomTestSetTitles Assert, "DiseaseSheetBuilder", "TestBuildCreatesWorksheet"

    Dim diseaseSheet As Worksheet
    Dim languagePrompt As String
    Dim table As ListObject

    On Error GoTo Fail

    languagePrompt = Translations.TranslatedValue("infoSelectLang")

    Set diseaseSheet = Builder.Build("Zeta", 1)

    Assert.AreEqual languagePrompt, diseaseSheet.Cells(2, 2).Value, "Language prompt should be translated"
    Assert.AreEqual 1, diseaseSheet.Cells(2, 3).Value, "Disease index should be recorded"
    Assert.IsTrue NameExists("disLang_1"), "Language name should be created"

    Set table = diseaseSheet.ListObjects("disTab_1")
    Assert.AreEqual "disTab_1", table.Name, "Table should use configured prefix"
    Assert.AreEqual Translate("varLabel", "Main Label"), table.HeaderRowRange.Cells(1, 4).Value, "Headers should be translated"

    Assert.AreEqual "= PARAMVARNAME", diseaseSheet.Cells(5, 2).Validation.Formula1, "Variable list validation should be applied"
    Assert.AreEqual "= PARAMCHOICESLIST", diseaseSheet.Cells(5, 4).Validation.Formula1, "Choice validation should be applied"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBuildCreatesWorksheet", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseSheetBuilder")
Public Sub TestBuildRespectsProvidedLanguage()
    CustomTestSetTitles Assert, "DiseaseSheetBuilder", "TestBuildRespectsProvidedLanguage"

    Dim diseaseSheet As Worksheet

    On Error GoTo Fail

    Set diseaseSheet = Builder.Build("Eta", 2, "FR")

    Assert.AreEqual "FR", diseaseSheet.Cells(2, 2).Value, "Provided language should be preserved"
    Assert.IsTrue NameExists("disLang_2"), "Custom index should create named range"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBuildRespectsProvidedLanguage", Err.Number, Err.Description
End Sub

'@section Helpers
'===============================================================================

Private Sub PrepareEnvironment()
    Dim dropdownSheet As Worksheet
    Dim translationSheet As Worksheet
    Dim data As Variant

    EnsureWorksheet ANCHOR_SHEET

    Set dropdownSheet = EnsureWorksheet(DROPDOWN_SHEET)
    ClearWorksheet dropdownSheet

    dropdownSheet.Range("A1").Value = "Languages"
    dropdownSheet.Range("A2:A3").Value = Application.WorksheetFunction.Transpose(Array("ENG", "FRA"))
    dropdownSheet.Range("B1").Value = "Status"
    dropdownSheet.Range("B2:B3").Value = Application.WorksheetFunction.Transpose(Array("core", "optional"))
    dropdownSheet.Range("C1").Value = "VarNames"
    dropdownSheet.Range("C2:C4").Value = Application.WorksheetFunction.Transpose(Array("var_age", "var_fever", "var_symptoms"))
    dropdownSheet.Range("D1").Value = "Choices"
    dropdownSheet.Range("D2:D4").Value = Application.WorksheetFunction.Transpose(Array("choice_age", "choice_fever", "choice_other"))

    AddName "__languages", dropdownSheet.Range("A2:A3")
    AddName "__var_status", dropdownSheet.Range("B2:B3")
    AddName "PARAMVARNAME", dropdownSheet.Range("C2:C4")
    AddName "PARAMCHOICESLIST", dropdownSheet.Range("D2:D4")

    Set Dropdowns = New TestDropdownStub

    Set translationSheet = EnsureWorksheet(TRANSLATION_SHEET)
    ClearWorksheet translationSheet

    data = Array( _
        Array("tag", "ENG"), _
        Array("selectValue", "Select a value"), _
        Array("infoSelectLang", "Select language"), _
        Array("varOrder", "Variable Order"), _
        Array("varSection", "Variable Section"), _
        Array("varName", "Variable Name"), _
        Array("varLabel", "Main Label"), _
        Array("varChoice", "Control"), _
        Array("choiceVal", "Choices"), _
        Array("varStatus", "Status"), _
        Array("errLang", "Please select a language") _
    )

    translationSheet.Range("A1").Resize(UBound(data) + 1, 2).Value = data
    translationSheet.ListObjects.Add SourceType:=xlSrcRange, Source:=translationSheet.Range("A1").Resize(UBound(data) + 1, 2), _
                                     XlListObjectHasHeaders:=xlYes

    Set Translations = TranslationObject.Create(translationSheet.ListObjects(1), "ENG")
    Set Builder = DiseaseSheetBuilder.Create(ThisWorkbook, Dropdowns, Translations)
End Sub

Private Sub CleanupEnvironment()
    DeleteWorksheetSafe "Zeta"
    DeleteWorksheetSafe "Eta"
    DeleteWorksheetSafe DROPDOWN_SHEET
    DeleteWorksheetSafe TRANSLATION_SHEET
    DeleteNameSafe "disLang_1"
    DeleteNameSafe "disLang_2"
    DeleteNameSafe "__languages"
    DeleteNameSafe "__var_status"
    DeleteNameSafe "PARAMVARNAME"
    DeleteNameSafe "PARAMCHOICESLIST"
End Sub

Private Sub DeleteWorksheetSafe(ByVal sheetName As String)
    On Error Resume Next
        ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
End Sub

Private Sub DeleteNameSafe(ByVal nameValue As String)
    On Error Resume Next
        ThisWorkbook.Names(nameValue).Delete
    On Error GoTo 0
End Sub

Private Function NameExists(ByVal nameValue As String) As Boolean
    On Error Resume Next
        ThisWorkbook.Names(nameValue)
        NameExists = (Err.Number = 0)
        Err.Clear
    On Error GoTo 0
End Function

Private Sub AddName(ByVal nameValue As String, ByVal refersToRange As Range)
    DeleteNameSafe nameValue
    ThisWorkbook.Names.Add Name:=nameValue, RefersTo:=refersToRange
End Sub

Private Function Translate(ByVal key As String, ByVal fallback As String) As String
    Translate = Translations.TranslatedValue(key)
    If LenB(Translate) = 0 Then Translate = fallback
End Function
