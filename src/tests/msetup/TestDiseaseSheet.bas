Attribute VB_Name = "TestDiseaseSheet"
Attribute VB_Description = "Tests ensuring DiseaseSheet creates worksheets with headers, validations, and tables"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests ensuring DiseaseSheet creates worksheets with headers, validations, and tables")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const ANCHOR_SHEET As String = "Variables"
Private Const DROPDOWN_SHEET As String = "DropdownStubSheet"
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
Private Const NAME_DISCODE As String = "__Var_DISCODE"

Private Assert As CustomTest
Private Builder As DiseaseSheet
Private Dropdowns As DropdownLists
Private VariablesManager As MasterSetupVariables

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseSheet"
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

'@TestMethod("DiseaseSheet")
Public Sub TestBuildCreatesWorksheet()
    CustomTestSetTitles Assert, "DiseaseSheet", "TestBuildCreatesWorksheet"

    Dim diseaseWksh As Worksheet
    Dim table As ListObject
    Dim sheetStore As HiddenNames
    Dim workbookStore As HiddenNames
    Dim diseases As BetterArray
    Dim validationFormula As String
    Dim labelHeader As String
    Dim choiceHeader As String
    Dim statusHeader As String
    Dim choicesValueHeader As String

    On Error GoTo Fail

    labelHeader = "Main Label"
    choiceHeader = "Choice"
    statusHeader = "Status"
    choicesValueHeader = "Choice Values"

    Set diseaseWksh = Builder.Build("Zeta")

    Assert.AreEqual "ENG", diseaseWksh.Cells(2, 2).Value, "Language cell should default to the first dropdown entry."
    Assert.IsTrue InStr(1, diseaseWksh.Cells(2, 2).Validation.Formula1, LANGUAGES_LIST, vbTextCompare) > 0, _
                 "Language cell should use the languages dropdown."
    
   
    Set table = diseaseWksh.ListObjects("disTab_001")

  
    
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

    'The sheet ships protected: what the user picks in has to be unlocked.
    Assert.IsFalse diseaseWksh.Cells(2, 2).Locked, "The language cell stays open on a protected sheet."
    Assert.IsFalse table.ListColumns("Variable Name").DataBodyRange.Locked, "The name column stays open on a protected sheet."
    Assert.IsFalse table.ListColumns(choiceHeader).DataBodyRange.Locked, "The choice column stays open on a protected sheet."
    Assert.IsFalse table.ListColumns(statusHeader).DataBodyRange.Locked, "The status column stays open on a protected sheet."

    'White sheet; the language cell and the headers carry the house blue.
    Assert.AreEqual RGB(0, 82, 155), CLng(diseaseWksh.Cells(2, 2).Interior.Color), "The language cell carries the house blue."
    Assert.AreEqual vbWhite, CLng(diseaseWksh.Cells(2, 2).Font.Color), "The language cell writes in white."
    Assert.AreEqual RGB(0, 82, 155), CLng(table.HeaderRowRange.Cells(1, 1).Interior.Color), "The table headers carry the house blue."
    Assert.AreEqual vbWhite, CLng(table.HeaderRowRange.Cells(1, 1).Font.Color), "The table headers write in white."
    Assert.AreEqual vbWhite, CLng(table.DataBodyRange.Cells(1, 1).Interior.Color), "The table body stays white."
    Assert.AreEqual vbWhite, CLng(diseaseWksh.Cells(1, 1).Interior.Color), "The sheet around the table stays white."

    'The gridlines are off, so the table carries its own dotted gray frame.
    Assert.AreEqual CLng(xlDot), CLng(table.DataBodyRange.Cells(1, 1).Borders(xlEdgeBottom).LineStyle), _
                    "The first body cell carries a dotted bottom edge."
    Assert.AreEqual RGB(166, 166, 166), CLng(table.DataBodyRange.Cells(1, 1).Borders(xlEdgeBottom).Color), _
                    "The frame is gray."



    Set sheetStore = HiddenNames.Create(diseaseWksh)
    Assert.AreEqual "disease", sheetStore.ValueAsString(SHEET_TAG_NAME), "Sheet tag metadata should be stored."
    Assert.AreEqual "Zeta", sheetStore.ValueAsString(NAME_DISNAME), "Disease name metadata should match the worksheet name."
    Assert.AreEqual "ENG", sheetStore.ValueAsString(NAME_DISLANG), "Language metadata should match the selected language."
    Assert.AreEqual 1&, sheetStore.ValueAsLong(NAME_INDEX, 0), "Disease index should be persisted through hidden names."
    Assert.AreEqual MARKER_NAME_PREFIX & "001", sheetStore.ValueAsString(NAME_DISCODE), "The disease code should live in the hidden names."



    Set workbookStore = HiddenNames.Create(ThisWorkbook)
    Assert.AreEqual "Zeta", workbookStore.ValueAsString(MARKER_NAME_PREFIX & "001"), _
                 "Workbook metadata should map marker names to worksheet names."

    Set diseases = Dropdowns.Values(DISEASES_LIST)
    Assert.IsTrue diseases.Includes("Zeta"), "Diseases dropdown should be updated with the new sheet name."

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBuildCreatesWorksheet", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseSheet")
Public Sub TestBuildRespectsProvidedLanguage()
    CustomTestSetTitles Assert, "DiseaseSheet", "TestBuildRespectsProvidedLanguage"

    Dim diseaseWksh As Worksheet
    Dim firstSheet As Worksheet
    Dim workbookStore As HiddenNames
    Dim diseases As BetterArray

    On Error GoTo Fail

    Set firstSheet = Builder.Build("Alpha")
    Set diseaseWksh = Builder.Build("Eta", "FRA")

    Assert.AreEqual "FRA", diseaseWksh.Cells(2, 2).Value, "Provided language should be preserved"
    Assert.AreEqual "disTab_002", diseaseWksh.ListObjects(1).Name, "Sequential builds should increment the table suffix."

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

'@TestMethod("DiseaseSheet")
Public Sub TestBuildRejectsInvalidInputs()
    CustomTestSetTitles Assert, "DiseaseSheet", "TestBuildRejectsInvalidInputs"

    Dim diseaseWksh As Worksheet

    On Error GoTo Fail

    Assert.AreEqual ProjectError.InvalidArgument, BuildExpectingError(vbNullString), _
                 "Empty disease names should raise invalid argument errors."

    Assert.AreEqual ProjectError.InvalidArgument, BuildExpectingError("Variables"), _
                 "Reserved disease names should be rejected."

    Set diseaseWksh = Builder.Build("Beta")
    Assert.IsTrue Not diseaseWksh Is Nothing, "Control build should succeed for unique names."

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

    Set Builder = DiseaseSheet.Create(ThisWorkbook, Dropdowns, VariablesManager)
End Sub

Private Sub CleanupEnvironment()
    DeleteWorksheetSafe "Zeta"
    DeleteWorksheetSafe "Eta"
    DeleteWorksheetSafe "Alpha"
    DeleteWorksheetSafe "Beta"
    DeleteWorksheetSafe "Gamma"
    DeleteWorksheetSafe DROPDOWN_SHEET
    ClearWorksheetSafe ANCHOR_SHEET

    DeleteNameSafe VARIABLE_NAME_RANGE
    DeleteNameSafe MARKER_NAME_PREFIX & "001"
    DeleteNameSafe MARKER_NAME_PREFIX & "002"
    DeleteNameSafe MARKER_NAME_PREFIX & "003"

    Set Builder = Nothing
    Set Dropdowns = Nothing
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
    Dim store As HiddenNames

    Set store = HiddenNames.Create(ThisWorkbook)
    store.SetListObjectHeader VARIABLE_NAME_RANGE, lo, "Variable Name"
End Sub

Private Sub AddDropdownList(ByVal target As DropdownLists, ByVal listName As String, ByVal values As Variant)
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
