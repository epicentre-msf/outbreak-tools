Attribute VB_Name = "TestMasterSetupImportService"
Attribute VB_Description = "Tests proving a disease exported for a setup folds back into the master setup"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests proving a disease exported for a setup folds back into the master setup")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const VARIABLES_SHEET As String = "Variables"
Private Const CHOICES_SHEET As String = "Choices"
Private Const TRANSLATIONS_SHEET As String = "Translations"
Private Const DROPDOWN_SHEET As String = "__dropdowns"
Private Const STAGING_SHEET As String = "__dis_import"
Private Const DISEASE_NAME As String = "Alpha"
Private Const LANGUAGES_LIST As String = "__data_languages"
Private Const STATUS_LIST As String = "__var_status"
Private Const CHOICES_LIST As String = "__lst_choices"
Private Const PROHIBITED_LIST As String = "__prohibited_diseases_list"
Private Const DISEASES_LIST As String = "__diseases_list"
Private Const VARIABLE_NAME_RANGE As String = "__Col__Variables"
Private Const MARKER_NAME_PREFIX As String = "DISSHEET"
Private Const CHOICES_HEADER_ROW As Long = 4

Private Assert As CustomTest
Private Builder As DiseaseSheet
Private Exporter As DiseaseExporter
Private ExportManager As DiseaseExportWorkbook
Private Dropdowns As DropdownLists
Private VariablesManager As MasterSetupVariables
Private MasterChoices As LLChoices
Private TranslationTable As ListObject
Private Service As MasterSetupImportService
Private ExportBook As Workbook

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestMasterSetupImportService"
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
    Set Exporter = Nothing
    Set ExportManager = Nothing
    Set Dropdowns = Nothing
    Set VariablesManager = Nothing
    Set MasterChoices = Nothing
    Set TranslationTable = Nothing
    Set Service = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    CleanupEnvironment
    PrepareEnvironment
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'The assertions of a test reach the results sheet only once flushed.
    Assert.Flush
    CleanupEnvironment
End Sub

'@section Tests
'===============================================================================

'@TestMethod("MasterSetupImportService")
Public Sub TestSetupExportRebuildsTheDiseaseSheet()
    CustomTestSetTitles Assert, "MasterSetupImportService", "TestSetupExportRebuildsTheDiseaseSheet"

    Dim diseaseWksh As Worksheet
    Dim rebuiltTable As ListObject
    Dim summary As DiseaseImportSummary
    Dim logger As DiseaseLogger
    Dim manager As DiseaseWorksheetManager

    On Error GoTo Fail

    Set diseaseWksh = Builder.Build(DISEASE_NAME)
    FillDiseaseTable diseaseWksh.ListObjects(1), Array( _
        Array(1, "var_a", "demographics", "Age", "choice_age", "0 to 4 | 5 to 14", "core"), _
        Array(2, "var_b", "symptoms", "Fever", "choice_fever", "yes | no", "core") _
    )
    Set ExportBook = Exporter.BuildDiseaseWorkbook(diseaseWksh, TranslationTable, DISEASE_NAME, "ENG", "DISSHEET001")

    'The disease goes away, then comes back off the exported file.
    Set manager = DiseaseWorksheetManager.Create()
    manager.RemoveWorksheet ThisWorkbook, DISEASE_NAME
    Set logger = DiseaseLogger.Create()

    Set summary = Service.ImportSetupExport(ExportBook, logger)

    Assert.IsTrue WorksheetExists(DISEASE_NAME), "The disease worksheet should be rebuilt"
    Assert.AreEqual DISEASE_NAME, HiddenNames.Create(ExportBook).ValueAsString("__Var_DISNAME"), "The export tags the workbook with the disease name"
    Assert.AreEqual DISEASE_NAME, Service.ReadDiseaseName(ExportBook), "The name is read off the tag"
    Assert.AreEqual DISEASE_NAME, Service.DiseaseName, "The disease name comes from the tag"
    Assert.AreEqual "ENG", Service.LanguageTag, "The language comes from the Metadata sheet"
    Assert.AreEqual "DISSHEET001", Service.DiseaseCode, "The code comes from the Metadata sheet"
    Assert.AreEqual "ENG", ThisWorkbook.Worksheets(DISEASE_NAME).Cells(2, 2).Value, "The rebuilt sheet takes the language of the file"

    Set rebuiltTable = ThisWorkbook.Worksheets(DISEASE_NAME).ListObjects(1)
    Assert.AreEqual 2, rebuiltTable.ListRows.Count, "Both variables of the file should land"
    Assert.AreEqual "var_a", rebuiltTable.DataBodyRange.Cells(1, 2).Value, "The first variable should land on the first line"
    Assert.AreEqual "symptoms", rebuiltTable.DataBodyRange.Cells(2, 3).Value, "The section should travel with the variable"
    Assert.AreEqual "choice_fever", rebuiltTable.DataBodyRange.Cells(2, 5).Value, "The control should land as the choice"
    Assert.AreEqual "core", rebuiltTable.DataBodyRange.Cells(2, 7).Value, "The status should travel with the variable"
    Assert.IsTrue Left$(rebuiltTable.DataBodyRange.Cells(1, 4).Formula, 1) = "=", "The label column carries its formula again"
    Assert.IsTrue Left$(rebuiltTable.DataBodyRange.Cells(2, 6).Formula, 1) = "=", "The choice values column carries its formula again"

    Assert.IsFalse summary Is Nothing, "The import answers the summary of the merge"
    Assert.AreEqual 2, summary.AppendedVariables.Length, "A rebuilt sheet takes every line as appended"
    Assert.AreEqual 0, Service.AddedVariables.Length, "Every variable was already in the Variables table"
    Assert.AreEqual 0, Service.AddedChoices.Length, "Every list was already on the Choices sheet"
    Assert.IsFalse WorksheetExists(STAGING_SHEET), "The staging sheet goes away"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSetupExportRebuildsTheDiseaseSheet", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupImportService")
Public Sub TestSetupExportAddsMissingVariablesAndLists()
    CustomTestSetTitles Assert, "MasterSetupImportService", "TestSetupExportAddsMissingVariablesAndLists"

    Dim diseaseWksh As Worksheet
    Dim diseaseTable As ListObject
    Dim summary As DiseaseImportSummary
    Dim logger As DiseaseLogger

    On Error GoTo Fail

    'var_new is on no Variables line and choice_new on no Choices line.
    Set diseaseWksh = Builder.Build(DISEASE_NAME)
    Set diseaseTable = diseaseWksh.ListObjects(1)
    FillDiseaseTable diseaseTable, Array( _
        Array(1, "var_a", "demographics", "Age", "choice_age", "0 to 4 | 5 to 14", "core"), _
        Array(2, "var_new", "history", "New label", "choice_new", "low | high", "optional") _
    )
    Set ExportBook = Exporter.BuildDiseaseWorkbook(diseaseWksh, TranslationTable, DISEASE_NAME, "ENG", "DISSHEET001")

    'The second line leaves the sheet; the import has to bring it back.
    diseaseTable.ListRows(2).Delete
    Set logger = DiseaseLogger.Create()

    Set summary = Service.ImportSetupExport(ExportBook, logger)

    Assert.AreEqual 2, diseaseTable.ListRows.Count, "The merge appends the line the sheet lost"
    Assert.AreEqual "var_new", diseaseTable.DataBodyRange.Cells(2, 2).Value, "The appended line carries the variable"
    Assert.AreEqual 1, summary.AppendedVariables.Length, "One variable is appended on the sheet"
    Assert.AreEqual 1, summary.UpdatedVariables.Length, "The line both carry is updated"

    Assert.IsTrue VariablesManager.HasVariable("var_new"), "The Variables table takes the variable it lacked"
    Assert.AreEqual "history", VariablesManager.SectionFor("var_new"), "The section comes from the dictionary"
    Assert.AreEqual "New label", VariablesManager.LabelFor("var_new"), "The label comes from the dictionary"
    Assert.AreEqual "choice_new", VariablesManager.DefaultChoiceFor("var_new"), "The control becomes the default choice"
    Assert.AreEqual "optional", VariablesManager.DefaultStatusFor("var_new"), "The status becomes the default status"
    Assert.AreEqual 1, Service.AddedVariables.Length, "One variable is added to the Variables table"
    Assert.AreEqual "var_new", Service.AddedVariables.Item(1), "The added variable is named"

    Assert.IsTrue MasterChoices.ChoiceExists("choice_new"), "The Choices sheet takes the list it lacked"
    Assert.AreEqual "low", ChoiceLabel("choice_new", 1), "An English file keeps its labels"
    Assert.AreEqual "high", ChoiceLabel("choice_new", 2), "Every label of the list travels"
    Assert.AreEqual 1, Service.AddedChoices.Length, "One list is added to the Choices sheet"
    Assert.AreEqual "choice_new", Service.AddedChoices.Item(1), "The added list is named"
    Assert.IsTrue ChoicesTableReaches("choice_new"), "The choices table stretches over the added rows"
    Assert.IsTrue logger.HasEntries, "The additions are logged"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSetupExportAddsMissingVariablesAndLists", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupImportService")
Public Sub TestForeignLanguageAddsListsWithEmptyLabels()
    CustomTestSetTitles Assert, "MasterSetupImportService", "TestForeignLanguageAddsListsWithEmptyLabels"

    Dim diseaseWksh As Worksheet
    Dim logger As DiseaseLogger

    On Error GoTo Fail

    Set diseaseWksh = Builder.Build(DISEASE_NAME, "FRA")
    FillDiseaseTable diseaseWksh.ListObjects(1), Array( _
        Array(1, "var_a", "demographics", "Age", "choice_fra", "bas | haut", "core") _
    )
    Set ExportBook = Exporter.BuildDiseaseWorkbook(diseaseWksh, TranslationTable, DISEASE_NAME, "FRA", "DISSHEET001")
    Set logger = DiseaseLogger.Create()

    Service.ImportSetupExport ExportBook, logger

    Assert.AreEqual "FRA", Service.LanguageTag, "The language of the file is French"
    Assert.IsTrue MasterChoices.ChoiceExists("choice_fra"), "The list lands under its name"
    Assert.AreEqual vbNullString, ChoiceLabel("choice_fra", 1), "A French file leaves the first label empty"
    Assert.AreEqual vbNullString, ChoiceLabel("choice_fra", 2), "A French file leaves the second label empty"
    Assert.AreEqual 1, Service.AddedChoices.Length, "The list still counts as added"
    Assert.AreEqual "FRA", ThisWorkbook.Worksheets(DISEASE_NAME).Cells(2, 2).Value, "The existing sheet keeps its language"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestForeignLanguageAddsListsWithEmptyLabels", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupImportService")
Public Sub TestFileWithoutNameTakesTheGivenOneAndOnlyItsColumns()
    CustomTestSetTitles Assert, "MasterSetupImportService", "TestFileWithoutNameTakesTheGivenOneAndOnlyItsColumns"

    Dim diseaseWksh As Worksheet
    Dim dictionaryTable As ListObject
    Dim extraColumn As ListColumn
    Dim builtTable As ListObject
    Dim alerts As Boolean

    On Error GoTo Fail

    Set diseaseWksh = Builder.Build(DISEASE_NAME)
    FillDiseaseTable diseaseWksh.ListObjects(1), Array( _
        Array(1, "var_a", "demographics", "Age", "choice_age", "0 to 4 | 5 to 14", "core") _
    )
    Set ExportBook = Exporter.BuildDiseaseWorkbook(diseaseWksh, TranslationTable, DISEASE_NAME, "ENG", "DISSHEET001")

    'A setup export: no tag, no Metadata sheet, and a wider dictionary.
    HiddenNames.Create(ExportBook).RemoveName "__Var_DISNAME"
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ExportBook.Worksheets("Metadata").Delete
    Application.DisplayAlerts = alerts

    Set dictionaryTable = ExportBook.Worksheets("Dictionary").ListObjects("Tab_Dictionary")
    Set extraColumn = dictionaryTable.ListColumns.Add
    extraColumn.Name = "Sheet Name"
    extraColumn.DataBodyRange.Value = "sheet1"
    Set extraColumn = dictionaryTable.ListColumns.Add
    extraColumn.Name = "Sub Section"
    extraColumn.DataBodyRange.Value = "sub"

    Assert.AreEqual vbNullString, Service.ReadDiseaseName(ExportBook), "A file with no tag and no Metadata names no disease"
    Assert.IsFalse Service.DiseaseNameIsFree(DISEASE_NAME), "A name already a worksheet is refused"
    Assert.IsFalse Service.DiseaseNameIsFree("Variables"), "A reserved name is refused"
    Assert.IsFalse Service.DiseaseNameIsFree(""), "An empty name is refused"
    Assert.IsTrue Service.DiseaseNameIsFree("Beta"), "A new name is free"

    Service.ImportSetupExport ExportBook, diseaseName:="Beta"

    Assert.IsTrue WorksheetExists("Beta"), "The given name takes the disease worksheet"
    Set builtTable = ThisWorkbook.Worksheets("Beta").ListObjects(1)
    Assert.AreEqual 7, builtTable.ListColumns.Count, "Only the columns of a disease table are taken"
    Assert.AreEqual "var_a", builtTable.DataBodyRange.Cells(1, 2).Value, "The variable lands"
    Assert.AreEqual "demographics", builtTable.DataBodyRange.Cells(1, 3).Value, "The section lands in its column"
    Assert.AreEqual "core", builtTable.DataBodyRange.Cells(1, 7).Value, "The status lands in its column"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestFileWithoutNameTakesTheGivenOneAndOnlyItsColumns", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

Private Sub PrepareEnvironment()
    Dim variablesSheet As Worksheet
    Dim dropdownSheet As Worksheet
    Dim translationSheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim variableTable As ListObject
    Dim data As Variant

    'Variables sheet: the master table, in the eight columns the master carries.
    Set variablesSheet = EnsureWorksheet(VARIABLES_SHEET)
    ClearWorksheet variablesSheet
    data = RowsToMatrix(Array( _
        Array("Variable Order", "Variable Section", "Variable Name", "Label", "Default Choice", "Choices Values", "Default Status", "Comments"), _
        Array(1, "demographics", "var_a", "Age", "choice_age", "0 to 4 | 5 to 14", "core", ""), _
        Array(2, "symptoms", "var_b", "Fever", "choice_fever", "yes | no", "core", "") _
    ))
    WriteMatrix variablesSheet.Range("A1"), data
    Set variableTable = variablesSheet.ListObjects.Add(xlSrcRange, variablesSheet.Range("A1").Resize(3, 8), _
                                                       XlListObjectHasHeaders:=xlYes)
    variableTable.Name = "TST_ImportVariables"

    Set VariablesManager = MasterSetupVariables.Create(variableTable)
    RegisterVariableName variableTable

    Set dropdownSheet = EnsureWorksheet(DROPDOWN_SHEET)
    ClearWorksheet dropdownSheet

    Set Dropdowns = DropdownLists.Create(dropdownSheet)
    AddDropdownList Dropdowns, LANGUAGES_LIST, Array("ENG", "FRA", "ESP")
    AddDropdownList Dropdowns, STATUS_LIST, Array("core", "optional", "hidden")
    AddDropdownList Dropdowns, CHOICES_LIST, Array("choice_age", "choice_fever")
    AddDropdownList Dropdowns, PROHIBITED_LIST, Array("Variables", "Translations", "Choices")
    AddDropdownList Dropdowns, DISEASES_LIST, Array("", "")

    'Translations sheet: one language column per language the file can carry.
    Set translationSheet = EnsureWorksheet(TRANSLATIONS_SHEET)
    ClearWorksheet translationSheet
    data = RowsToMatrix(Array(Array("tag", "ENG", "FRA"), Array("hello", "Hello", "Bonjour")))
    WriteMatrix translationSheet.Range("A1"), data
    translationSheet.ListObjects.Add SourceType:=xlSrcRange, Source:=translationSheet.Range("A1").Resize(2, 3), _
                                      XlListObjectHasHeaders:=xlYes
    Set TranslationTable = translationSheet.ListObjects(1)

    'Choices sheet: headers on row 4, the way the master file carries them.
    'The values are picked so that Excel stores none of them as a date.
    Set choicesSheet = EnsureWorksheet(CHOICES_SHEET)
    ClearWorksheet choicesSheet
    data = RowsToMatrix(Array( _
        Array("list name", "ordering list", "translated label", "label", "short label"), _
        Array("choice_age", 1, "", "0 to 4", "0 to 4"), _
        Array("choice_age", 2, "", "5 to 14", "5 to 14"), _
        Array("choice_fever", 1, "", "yes", "yes"), _
        Array("choice_fever", 2, "", "no", "no") _
    ))
    WriteMatrix choicesSheet.Cells(CHOICES_HEADER_ROW, 1), data
    choicesSheet.ListObjects.Add SourceType:=xlSrcRange, _
                                 Source:=choicesSheet.Cells(CHOICES_HEADER_ROW, 1).Resize(5, 5), _
                                 XlListObjectHasHeaders:=xlYes
    Set MasterChoices = LLChoices.Create(choicesSheet, CHOICES_HEADER_ROW, 1)

    Set Builder = DiseaseSheet.Create(ThisWorkbook, Dropdowns, VariablesManager)
    Set ExportManager = DiseaseExportWorkbook.Create()
    Set Exporter = DiseaseExporter.Create(ExportManager, ApplicationState.Create(Application))
    Set Service = MasterSetupImportService.Create(ThisWorkbook, Builder, Dropdowns, VariablesManager, MasterChoices)
End Sub

Private Sub CleanupEnvironment()
    On Error Resume Next
        If Not ExportBook Is Nothing Then ExportBook.Close SaveChanges:=False
    On Error GoTo 0
    Set ExportBook = Nothing

    DeleteWorksheetSafe DISEASE_NAME
    DeleteWorksheetSafe "Beta"
    DeleteWorksheetSafe DROPDOWN_SHEET
    DeleteWorksheetSafe TRANSLATIONS_SHEET
    DeleteWorksheetSafe CHOICES_SHEET
    DeleteWorksheetSafe STAGING_SHEET
    ClearWorksheetSafe VARIABLES_SHEET

    DeleteNameSafe VARIABLE_NAME_RANGE
    DeleteNameSafe MARKER_NAME_PREFIX & "001"
    DeleteNameSafe MARKER_NAME_PREFIX & "002"
    DeleteNameSafe MARKER_NAME_PREFIX & "003"
End Sub

'@description Write the lines of a disease table over the line the builder left.
Private Sub FillDiseaseTable(ByVal table As ListObject, ByVal rows As Variant)
    Dim rowIndex As Long

    For rowIndex = LBound(rows) To UBound(rows)
        If table.ListRows.Count < rowIndex - LBound(rows) + 1 Then table.ListRows.Add
        table.ListRows(rowIndex - LBound(rows) + 1).Range.Value = rows(rowIndex)
    Next rowIndex
End Sub

'@description The label written for the nth row of a list on the master Choices sheet.
'@details Read cell by cell down the list name column, so rows AddChoice
'wrote outside the choices table are read too.
Private Function ChoiceLabel(ByVal listName As String, ByVal position As Long) As String
    Dim choicesSheet As Worksheet
    Dim rowIndex As Long
    Dim seen As Long
    Dim lastRow As Long

    Set choicesSheet = ThisWorkbook.Worksheets(CHOICES_SHEET)
    lastRow = choicesSheet.Cells(choicesSheet.Rows.Count, 1).End(xlUp).Row

    For rowIndex = CHOICES_HEADER_ROW + 1 To lastRow
        If StrComp(CStr(choicesSheet.Cells(rowIndex, 1).Value), listName, vbTextCompare) = 0 Then
            seen = seen + 1
            If seen = position Then
                'The third column is the translated label, where LLChoices
                'writes the label when that column exists.
                ChoiceLabel = CStr(choicesSheet.Cells(rowIndex, 3).Value)
                Exit Function
            End If
        End If
    Next rowIndex

    ChoiceLabel = "<missing>"
End Function

'@description Whether the choices ListObject covers every row naming the list.
Private Function ChoicesTableReaches(ByVal listName As String) As Boolean
    Dim choicesSheet As Worksheet
    Dim table As ListObject
    Dim lastRow As Long
    Dim rowIndex As Long

    Set choicesSheet = ThisWorkbook.Worksheets(CHOICES_SHEET)
    Set table = choicesSheet.ListObjects(1)
    lastRow = choicesSheet.Cells(choicesSheet.Rows.Count, 1).End(xlUp).Row

    For rowIndex = lastRow To CHOICES_HEADER_ROW + 1 Step -1
        If StrComp(CStr(choicesSheet.Cells(rowIndex, 1).Value), listName, vbTextCompare) = 0 Then
            ChoicesTableReaches = Not Intersect(table.Range, choicesSheet.Cells(rowIndex, 1)) Is Nothing
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub DeleteWorksheetSafe(ByVal sheetName As String)
    On Error Resume Next
        ThisWorkbook.Worksheets(sheetName).Delete
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

Private Sub DeleteNameSafe(ByVal nameValue As String)
    On Error Resume Next
        ThisWorkbook.Names(nameValue).Delete
    On Error GoTo 0
End Sub

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet

    For Each sheet In ThisWorkbook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            WorksheetExists = True
            Exit Function
        End If
    Next sheet
End Function

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
