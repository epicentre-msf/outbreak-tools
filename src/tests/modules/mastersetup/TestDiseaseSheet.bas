Attribute VB_Name = "TestDiseaseSheet"
Attribute VB_Description = "Tests ensuring DiseaseSheet creates worksheets with headers, validations, and tables"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests ensuring DiseaseSheetBuilder creates worksheets with headers, validations, and tables")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const ANCHOR_SHEET As String = "Variables"
Private Const DROPDOWN_SHEET As String = "DropdownStubSheet"
Private Const TRANSLATION_SHEET As String = "SheetBuilderTranslations"
Private Const LANGUAGES_LIST As String = "__data_languages"
Private Const STATUS_LIST As String = "__var_status"
Private Const CHOICES_LIST As String = "__lst_choices"
Private Const PROHIBITED_LIST As String = "__prohibited_diseases_list"
Private Const DISEASES_LIST As String = "__diseases_list"
Private Const VARIABLE_NAME_RANGE As String = "__Col__Variables"
Private Const MARKER_NAME_PREFIX As String = "DISSHEET"
Private Const SHEET_TAG_NAME As String = "sheetTag"
Private Const NAME_DISNAME As String = "__Var_DISNAME"
Private Const NAME_DISLANG As String = "__Var_DISLANG"
Private Const NAME_INDEX As String = "__Var_DISINDEX"

Private Assert As ICustomTest
Private Builder As IDiseaseSheet
Private Dropdowns As IDropdownLists
Private Translations As ITranslationObject
Private VariablesManager As IMasterSetupVariables

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
    Set VariablesManager = Nothing
    DeleteWorksheet ANCHOR_SHEET
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    CleanupEnvironment
    PrepareEnvironment
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Assert.Flush
    CleanupEnvironment
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseSheetBuilder")
Public Sub TestBuildCreatesWorksheet()
    CustomTestSetTitles Assert, "DiseaseSheetBuilder", "TestBuildCreatesWorksheet"

    Dim diseaseSheet As Worksheet
    Dim table As ListObject
    Dim sheetStore As IHiddenNames
    Dim workbookStore As IHiddenNames
    Dim diseases As BetterArray
    Dim validationFormula As String
    Dim labelHeader As String
    Dim choiceHeader As String
    Dim statusHeader As String
    Dim choicesValueHeader As String

    On Error GoTo Fail

    labelHeader = Translate("varLabel", "Main Label")
    choiceHeader = Translate("varChoice", "Choice")
    statusHeader = Translate("varStatus", "Status")
    choicesValueHeader = Translate("choiceVal", "Choice Values")

    Set diseaseSheet = Builder.Build("Zeta")

    Assert.AreEqual "ENG", diseaseSheet.Cells(2, 2).Value, "Language cell should default to the first dropdown entry."
    Assert.AreEqual MARKER_NAME_PREFIX, diseaseSheet.Cells(2, 4).Value, "Marker cell should identify disease worksheets."
    Assert.IsTrue InStr(1, diseaseSheet.Cells(2, 2).Validation.Formula1, LANGUAGES_LIST, vbTextCompare) > 0, _
                 "Language cell should use the languages dropdown."
    
   
    Set table = diseaseSheet.ListObjects("disTab_001")

  
    
    Assert.AreEqual labelHeader, table.HeaderRowRange.Cells(1, 4).Value, "Headers should be translated"

    validationFormula = table.ListColumns("Variable Name").DataBodyRange.Validation.Formula1
    Assert.IsTrue InStr(1, validationFormula, VARIABLE_NAME_RANGE, vbTextCompare) > 0, _
                 "Variable column should reference the workbook variable list."

    

    validationFormula = table.ListColumns(choiceHeader).DataBodyRange.Validation.Formula1
    Assert.IsTrue InStr(1, validationFormula, CHOICES_LIST, vbTextCompare) > 0, _
                 "Choice column should be validated against the choices dropdown."

    Debug.Print "anchor"
   
    
    validationFormula = table.ListColumns(statusHeader).DataBodyRange.Validation.Formula1
    Assert.IsTrue InStr(1, validationFormula, STATUS_LIST, vbTextCompare) > 0, _
                 "Status column should use the status dropdown."

    
    Assert.IsTrue table.ListColumns(choicesValueHeader).DataBodyRange.Locked, "Choice values column should be locked."
    Assert.IsTrue table.ListColumns(labelHeader).DataBodyRange.Locked, "Translated label column should be locked."
    
    

    Set sheetStore = HiddenNames.Create(diseaseSheet)
    Assert.AreEqual "disease", sheetStore.ValueAsString(SHEET_TAG_NAME), "Sheet tag metadata should be stored."
    Assert.AreEqual "Zeta", sheetStore.ValueAsString(NAME_DISNAME), "Disease name metadata should match the worksheet name."
    Assert.AreEqual "ENG", sheetStore.ValueAsString(NAME_DISLANG), "Language metadata should match the selected language."
    Assert.AreEqual 1&, sheetStore.ValueAsLong(NAME_INDEX, 0), "Disease index should be persisted through hidden names."



    Set workbookStore = HiddenNames.Create(ThisWorkbook)
    Assert.AreEqual "Zeta", workbookStore.ValueAsString(MARKER_NAME_PREFIX & "001"), _
                 "Workbook metadata should map marker names to worksheet names."

    Set diseases = Dropdowns.Values(DISEASES_LIST)
    Assert.IsTrue diseases.Includes("Zeta"), "Diseases dropdown should be updated with the new sheet name."

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBuildCreatesWorksheet", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseSheetBuilder")
Public Sub TestBuildRespectsProvidedLanguage()
    CustomTestSetTitles Assert, "DiseaseSheetBuilder", "TestBuildRespectsProvidedLanguage"

    Dim diseaseSheet As Worksheet
    Dim firstSheet As Worksheet
    Dim workbookStore As IHiddenNames
    Dim diseases As BetterArray

    On Error GoTo Fail

    Set firstSheet = Builder.Build("Alpha")
    Set diseaseSheet = Builder.Build("Eta", "FRA")

    Assert.AreEqual "FRA", diseaseSheet.Cells(2, 2).Value, "Provided language should be preserved"
    Assert.AreEqual "disTab_002", diseaseSheet.ListObjects(1).Name, "Sequential builds should increment the table suffix."

    Set workbookStore = HiddenNames.Create(ThisWorkbook)
    Assert.AreEqual "Eta", workbookStore.ValueAsString(MARKER_NAME_PREFIX & "002"), _
                 "Workbook marker name should reference the latest worksheet."

    Set diseases = Dropdowns.Values(DISEASES_LIST)
    Assert.IsTrue diseases.Includes("Alpha"), "Existing disease names should remain in the dropdown."
    Assert.IsTrue diseases.Includes("Eta"), "New disease names should be appended to the dropdown."

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBuildRespectsProvidedLanguage", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseSheetBuilder")
Public Sub TestBuildRejectsInvalidInputs()
    CustomTestSetTitles Assert, "DiseaseSheetBuilder", "TestBuildRejectsInvalidInputs"

    Dim diseaseSheet As Worksheet

    On Error GoTo Fail

    Assert.AreEqual ProjectError.InvalidArgument, BuildExpectingError(vbNullString), _
                 "Empty disease names should raise invalid argument errors."

    Assert.AreEqual ProjectError.InvalidArgument, BuildExpectingError("Variables"), _
                 "Reserved disease names should be rejected."

    Set diseaseSheet = Builder.Build("Beta")
    Assert.IsTrue Not diseaseSheet Is Nothing, "Control build should succeed for unique names."

    Assert.AreEqual ProjectError.InvalidArgument, BuildExpectingError("Beta"), _
                 "Duplicate disease names should not be allowed."

    Assert.AreEqual ProjectError.InvalidArgument, BuildExpectingError("Gamma", "DEU"), _
                 "Providing an unknown language should be rejected."

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBuildRejectsInvalidInputs", Err.Number, Err.Description
End Sub

'@section Helpers
'===============================================================================

Private Sub PrepareEnvironment()
    Dim dropdownSheet As Worksheet
    Dim translationSheet As Worksheet
    Dim data As Variant
    Dim variablesSheet As Worksheet
    Dim variableTable As ListObject

    Set variablesSheet = EnsureWorksheet(ANCHOR_SHEET)
    ClearWorksheet variablesSheet

    variablesSheet.Range("A1").Value = "Variable Order"
    variablesSheet.Range("B1").Value = "Variable Section"
    variablesSheet.Range("C1").Value = "Variable Name"
    variablesSheet.Range("C2").Value = "var_age"
    variablesSheet.Range("B2").Value = "Age"
    variablesSheet.Range("C3").Value = "var_fever"
    variablesSheet.Range("B3").Value = "Fever"

    Set variableTable = variablesSheet.ListObjects.Add(xlSrcRange, variablesSheet.Range("A1:C3"), _
                                                       XlListObjectHasHeaders:=xlYes)
    variableTable.Name = "TST_MasterVariables"

    Set VariablesManager = MasterSetupVariables.Create(variableTable)
    RegisterVariableName variableTable

    Set dropdownSheet = EnsureWorksheet(DROPDOWN_SHEET)
    ClearWorksheet dropdownSheet

    Set Dropdowns = DropdownLists.Create(dropdownSheet)
    AddDropdownList Dropdowns, LANGUAGES_LIST, Array("ENG", "FRA")
    AddDropdownList Dropdowns, STATUS_LIST, Array("core", "optional")
    AddDropdownList Dropdowns, CHOICES_LIST, Array("choice_age", "choice_fever", "choice_other")
    AddDropdownList Dropdowns, PROHIBITED_LIST, Array("Variables", "Translations")
    AddDropdownList Dropdowns, DISEASES_LIST, Array("", "")

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
        Array("varChoice", "Choice"), _
        Array("choiceVal", "Choice Values"), _
        Array("varStatus", "Status"), _
        Array("errLang", "Please select a language") _
    )

    translationSheet.Range("A1").Resize(UBound(data) + 1, 2).Value = data
    translationSheet.ListObjects.Add SourceType:=xlSrcRange, Source:=translationSheet.Range("A1").Resize(UBound(data) + 1, 2), _
                                     XlListObjectHasHeaders:=xlYes

    Set Translations = TranslationObject.Create(translationSheet.ListObjects(1), "ENG")
    Set Builder = DiseaseSheet.Create(ThisWorkbook, Dropdowns, Translations, VariablesManager)
End Sub

Private Sub CleanupEnvironment()
    DeleteWorksheetSafe "Zeta"
    DeleteWorksheetSafe "Eta"
    DeleteWorksheetSafe "Alpha"
    DeleteWorksheetSafe "Beta"
    DeleteWorksheetSafe "Gamma"
    DeleteWorksheetSafe DROPDOWN_SHEET
    DeleteWorksheetSafe TRANSLATION_SHEET
    ClearWorksheetSafe ANCHOR_SHEET

    DeleteNameSafe VARIABLE_NAME_RANGE
    DeleteNameSafe MARKER_NAME_PREFIX & "001"
    DeleteNameSafe MARKER_NAME_PREFIX & "002"
    DeleteNameSafe MARKER_NAME_PREFIX & "003"

    Set Builder = Nothing
    Set Dropdowns = Nothing
    Set Translations = Nothing
    Set VariablesManager = Nothing
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

Private Function Translate(ByVal key As String, ByVal fallback As String) As String
    Translate = Translations.TranslatedValue(key)
    If LenB(Translate) = 0 Or (Translate = key) Then Translate = fallback
End Function

Private Sub ClearWorksheetSafe(ByVal sheetName As String)
    Dim sh As Worksheet

    On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If Not sh Is Nothing Then
        ClearWorksheet sh
    End If
End Sub

Private Sub RegisterVariableName(ByVal lo As ListObject)
    Dim store As IHiddenNames

    Set store = HiddenNames.Create(ThisWorkbook)
    store.SetListObjectHeader VARIABLE_NAME_RANGE, lo, "Variable Name"
End Sub

Private Sub AddDropdownList(ByVal target As IDropdownLists, ByVal listName As String, ByVal values As Variant)
    Dim listValues As BetterArray

    Set listValues = BuildBetterArray(values)
    If listValues Is Nothing Then Exit Sub

    target.Add listValues, listName
End Sub

Private Function BuildBetterArray(ByVal values As Variant) As BetterArray
    Dim arr As BetterArray
    Dim idx As Long

    If Not IsArray(values) Then Exit Function

    Set arr = New BetterArray
    arr.LowerBound = 1
    For idx = LBound(values) To UBound(values)
        arr.Push CStr(values(idx))
    Next idx

    Set BuildBetterArray = arr
End Function

Private Function BuildExpectingError(ByVal diseaseName As String, Optional ByVal languageTag As String = vbNullString) As Long
    Dim unused As Worksheet

    On Error Resume Next
        Set unused = Builder.Build(diseaseName, languageTag)
        BuildExpectingError = Err.Number
        Err.Clear
    On Error GoTo 0

    If BuildExpectingError = 0 And Not unused Is Nothing Then
        DeleteWorksheetSafe unused.Name
    End If
End Function
