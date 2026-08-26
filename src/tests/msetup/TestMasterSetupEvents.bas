Attribute VB_Name = "TestMasterSetupEvents"
Attribute VB_Description = "Tests covering the master setup event structure end to end"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests covering the master setup event structure end to end")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const VARIABLES_SHEET As String = "Variables"
Private Const CHOICES_SHEET As String = "Choices"
Private Const TRANSLATIONS_SHEET As String = "Translations"
Private Const DROPDOWNS_SHEET As String = "__dropdowns"
Private Const PASSWORDS_SHEET As String = "__pass"
Private Const DISEASE_NAME As String = "EvtAlpha"
Private Const CHOICES_LIST As String = "__lst_choices"
Private Const STATUS_LIST As String = "__var_status"

Private Assert As CustomTest

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestMasterSetupEvents"
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

'@TestMethod("MasterSetupEvents")
Public Sub TestPreparationPutsDropdownsOnTheRightColumns()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestPreparationPutsDropdownsOnTheRightColumns"

    Dim variablesTable As ListObject
    Dim choiceValidationText As String
    Dim statusValidationText As String

    On Error GoTo Fail

    'The service reads the managers off the prepared fixture sheets. A
    'Property Get cannot stand alone as a statement; the assignment is what
    'resolves them.
    Dim preparedVariables As MasterSetupVariables
    Set preparedVariables = MasterSetupEventsManager.MasterSetupService.Variables

    Set variablesTable = ThisWorkbook.Worksheets(VARIABLES_SHEET).ListObjects(1)

    choiceValidationText = variablesTable.ListColumns("Default Choice").DataBodyRange.Cells(1, 1).Validation.Formula1
    statusValidationText = variablesTable.ListColumns("Default Status").DataBodyRange.Cells(1, 1).Validation.Formula1

    Assert.IsTrue InStr(1, choiceValidationText, CHOICES_LIST, vbTextCompare) > 0, _
                  "The Default Choice column should validate against the choices list"
    Assert.IsTrue InStr(1, statusValidationText, STATUS_LIST, vbTextCompare) > 0, _
                  "The Default Status column should validate against the status list"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestPreparationPutsDropdownsOnTheRightColumns", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupEvents")
Public Sub TestVariablePickFillsTheDiseaseLine()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestVariablePickFillsTheDiseaseLine"

    Dim diseaseWksh As Worksheet
    Dim table As ListObject
    Dim nameCell As Range
    Dim choiceCell As Range

    On Error GoTo Fail

    Set diseaseWksh = BuildDiseaseFixture()
    Set table = diseaseWksh.ListObjects(1)
    Set nameCell = table.DataBodyRange.Cells(1, 2)
    Set choiceCell = table.DataBodyRange.Cells(1, 5)

    nameCell.Value = "var_age"
    MasterSetupEventsManager.MsSheetChanged diseaseWksh, nameCell

    Assert.AreEqual "demographics", table.DataBodyRange.Cells(1, 3).Value, "The section should fill from the Variables sheet"
    Assert.AreEqual "choice_age", choiceCell.Value, "The default choice should fill from the Variables sheet"
    Assert.AreEqual "core", table.DataBodyRange.Cells(1, 7).Value, "The default status should fill from the Variables sheet"
    Assert.IsTrue InStr(1, choiceCell.Validation.Formula1, CHOICES_LIST, vbTextCompare) > 0, _
                  "The Choice cell should carry the choices dropdown"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestVariablePickFillsTheDiseaseLine", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupEvents")
'@sub-title A pick on a protected disease sheet fills its line while the Variables sheet is protected too.
'@details The shipped file is protected and its dropdown sheet is very
'hidden. The service reads the managers off the sheets and writes nothing
'of its own, so the handler's own unprotect is the one write path. This is
'the guard of the fault the owner met on the mock.
Public Sub TestVariablePickFillsAProtectedDiseaseLine()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestVariablePickFillsAProtectedDiseaseLine"

    Dim diseaseWksh As Worksheet
    Dim variablesSheet As Worksheet
    Dim table As ListObject
    Dim nameCell As Range
    Dim pass As Passwords

    On Error GoTo Fail

    Set diseaseWksh = BuildDiseaseFixture()
    Set table = diseaseWksh.ListObjects(1)
    Set nameCell = table.DataBodyRange.Cells(1, 2)
    Set variablesSheet = ThisWorkbook.Worksheets(VARIABLES_SHEET)

    'The service is dropped so the managers are built on the protected
    'file, the way a reset VBA project builds them inside a change event.
    MasterSetupEventsManager.DisposeMasterSetup

    PreparePasswordsFixture PASSWORDS_SHEET
    Set pass = Passwords.Create(ThisWorkbook.Worksheets(PASSWORDS_SHEET))
    pass.DisplayPrompts = False
    pass.Protect variablesSheet.Name
    pass.Protect diseaseWksh.Name
    ThisWorkbook.Worksheets(DROPDOWNS_SHEET).Visible = xlSheetVeryHidden

    nameCell.Value = "var_age"
    MasterSetupEventsManager.MsSheetChanged diseaseWksh, nameCell

    Assert.AreEqual "demographics", table.DataBodyRange.Cells(1, 3).Value, "The section should fill on a protected file"
    Assert.AreEqual "choice_age", table.DataBodyRange.Cells(1, 5).Value, "The default choice should fill on a protected file"
    Assert.AreEqual "core", table.DataBodyRange.Cells(1, 7).Value, "The default status should fill on a protected file"
    Assert.IsTrue diseaseWksh.ProtectContents, "The disease sheet should be protected again after the fill"
    Assert.IsTrue variablesSheet.ProtectContents, "The Variables sheet should stay protected through the fill"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestVariablePickFillsAProtectedDiseaseLine", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupEvents")
Public Sub TestClearedVariableCleansTheDiseaseLine()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestClearedVariableCleansTheDiseaseLine"

    Dim diseaseWksh As Worksheet
    Dim table As ListObject
    Dim nameCell As Range
    Dim choiceCell As Range
    Dim validationGone As Boolean
    Dim probe As String

    On Error GoTo Fail

    Set diseaseWksh = BuildDiseaseFixture()
    Set table = diseaseWksh.ListObjects(1)
    Set nameCell = table.DataBodyRange.Cells(1, 2)
    Set choiceCell = table.DataBodyRange.Cells(1, 5)

    nameCell.Value = "var_age"
    MasterSetupEventsManager.MsSheetChanged diseaseWksh, nameCell

    nameCell.Value = vbNullString
    MasterSetupEventsManager.MsSheetChanged diseaseWksh, nameCell

    Assert.AreEqual vbNullString, CStr(table.DataBodyRange.Cells(1, 3).Value), "The section should clear with the name"
    Assert.AreEqual vbNullString, CStr(choiceCell.Value), "The choice should clear with the name"
    Assert.AreEqual vbNullString, CStr(table.DataBodyRange.Cells(1, 7).Value), "The status should clear with the name"

    'A deleted validation makes the probe raise; that raise is the assertion.
    On Error Resume Next
        probe = choiceCell.Validation.Formula1
        validationGone = (Err.Number <> 0)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue validationGone, "The choice dropdown should go with the cleared line"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestClearedVariableCleansTheDiseaseLine", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupEvents")
Public Sub TestDefaultChoicePickFillsChoicesValues()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestDefaultChoicePickFillsChoicesValues"

    Dim variablesTable As ListObject
    Dim choiceCell As Range
    Dim valuesCell As Range

    On Error GoTo Fail

    'The service resolves the managers once over the prepared fixture
    'sheets. A Property Get cannot stand alone as a statement; the
    'assignment is what resolves them.
    Dim preparedVariables As MasterSetupVariables
    Set preparedVariables = MasterSetupEventsManager.MasterSetupService.Variables

    Set variablesTable = ThisWorkbook.Worksheets(VARIABLES_SHEET).ListObjects(1)
    Set choiceCell = variablesTable.ListColumns("Default Choice").DataBodyRange.Cells(1, 1)
    Set valuesCell = variablesTable.ListColumns("Choices Values").DataBodyRange.Cells(1, 1)

    choiceCell.Value = "choice_age"
    MasterSetupEventsManager.MsSheetChanged ThisWorkbook.Worksheets(VARIABLES_SHEET), choiceCell

    Assert.AreEqual "0-4 | 5-14", valuesCell.Value, "The choices values should fill, joined with a pipe"

    choiceCell.Value = vbNullString
    MasterSetupEventsManager.MsSheetChanged ThisWorkbook.Worksheets(VARIABLES_SHEET), choiceCell

    Assert.AreEqual vbNullString, CStr(valuesCell.Value), "An unselected choice should empty the values"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDefaultChoicePickFillsChoicesValues", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupEvents")
Public Sub TestTranslatedLabelEditRecalculatesLabels()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestTranslatedLabelEditRecalculatesLabels"

    Dim choicesSheet As Worksheet
    Dim table As ListObject
    Dim translatedCell As Range
    Dim labelCell As Range

    On Error GoTo Fail

    Set choicesSheet = ThisWorkbook.Worksheets(CHOICES_SHEET)
    Set table = choicesSheet.ListObjects(1)
    Set translatedCell = table.ListColumns("translated label").DataBodyRange.Cells(1, 1)
    Set labelCell = table.ListColumns("label").DataBodyRange.Cells(1, 1)

    'Manual calculation holds the formula still; only the handler's
    'Calculate moves it.
    translatedCell.Value = "Bonjour"
    Assert.IsFalse CStr(labelCell.Value) = "Bonjour", "Manual calculation should hold the label still"

    MasterSetupEventsManager.MsSheetChanged choicesSheet, translatedCell

    Assert.AreEqual "Bonjour", labelCell.Value, "The label should recalculate from the translated label"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTranslatedLabelEditRecalculatesLabels", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupEvents")
Public Sub TestChoiceEditRefreshesJoinedValues()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestChoiceEditRefreshesJoinedValues"

    Dim variablesSheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim variablesTable As ListObject
    Dim choiceCell As Range
    Dim valuesCell As Range
    Dim translatedCell As Range

    On Error GoTo Fail

    'A Property Get cannot stand alone as a statement; the assignment
    'is what resolves the managers over the prepared fixture sheets.
    Dim preparedVariables As MasterSetupVariables
    Set preparedVariables = MasterSetupEventsManager.MasterSetupService.Variables

    Set variablesSheet = ThisWorkbook.Worksheets(VARIABLES_SHEET)
    Set variablesTable = variablesSheet.ListObjects(1)
    Set choiceCell = variablesTable.ListColumns("Default Choice").DataBodyRange.Cells(1, 1)
    Set valuesCell = variablesTable.ListColumns("Choices Values").DataBodyRange.Cells(1, 1)

    choiceCell.Value = "choice_age"
    MasterSetupEventsManager.MsSheetChanged variablesSheet, choiceCell
    Assert.AreEqual "0-4 | 5-14", valuesCell.Value, "The joined values should start from the labels"

    Set choicesSheet = ThisWorkbook.Worksheets(CHOICES_SHEET)
    Set translatedCell = choicesSheet.Range("C5")
    translatedCell.Value = "zero-four"
    MasterSetupEventsManager.MsSheetChanged choicesSheet, translatedCell

    Assert.AreEqual "zero-four | 5-14", valuesCell.Value, "A choices edit should rejoin the values on the Variables sheet"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestChoiceEditRefreshesJoinedValues", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupEvents")
Public Sub TestDeletedChoiceEmptiesReferences()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestDeletedChoiceEmptiesReferences"

    Dim variablesSheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim variablesTable As ListObject
    Dim diseaseWksh As Worksheet
    Dim diseaseTable As ListObject
    Dim nameCell As Range
    Dim choiceCell As Range
    Dim valuesCell As Range

    On Error GoTo Fail

    Set diseaseWksh = BuildDiseaseFixture()
    Set diseaseTable = diseaseWksh.ListObjects(1)
    Set nameCell = diseaseTable.DataBodyRange.Cells(1, 2)

    nameCell.Value = "var_age"
    MasterSetupEventsManager.MsSheetChanged diseaseWksh, nameCell
    Assert.AreEqual "choice_age", diseaseTable.DataBodyRange.Cells(1, 5).Value, "The disease line should start with the default choice"

    Set variablesSheet = ThisWorkbook.Worksheets(VARIABLES_SHEET)
    Set variablesTable = variablesSheet.ListObjects(1)
    Set choiceCell = variablesTable.ListColumns("Default Choice").DataBodyRange.Cells(1, 1)
    Set valuesCell = variablesTable.ListColumns("Choices Values").DataBodyRange.Cells(1, 1)
    choiceCell.Value = "choice_age"
    MasterSetupEventsManager.MsSheetChanged variablesSheet, choiceCell

    'The choice goes away: its rows take another list name.
    Set choicesSheet = ThisWorkbook.Worksheets(CHOICES_SHEET)
    choicesSheet.Range("A5").Value = "choice_gone"
    choicesSheet.Range("A6").Value = "choice_gone"
    MasterSetupEventsManager.MsSheetChanged choicesSheet, choicesSheet.Range("A5:A6")

    Assert.AreEqual vbNullString, CStr(choiceCell.Value), "A deleted choice should empty the variables default"
    Assert.AreEqual vbNullString, CStr(valuesCell.Value), "A deleted choice should empty the joined values"
    Assert.AreEqual vbNullString, CStr(diseaseTable.DataBodyRange.Cells(1, 5).Value), "A deleted choice should empty the disease line"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDeletedChoiceEmptiesReferences", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupEvents")
Public Sub TestVariableDefaultChangePropagatesToDiseases()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestVariableDefaultChangePropagatesToDiseases"

    Dim variablesSheet As Worksheet
    Dim variablesTable As ListObject
    Dim diseaseWksh As Worksheet
    Dim diseaseTable As ListObject
    Dim nameCell As Range
    Dim choiceCell As Range

    On Error GoTo Fail

    Set diseaseWksh = BuildDiseaseFixture()
    Set diseaseTable = diseaseWksh.ListObjects(1)
    Set nameCell = diseaseTable.DataBodyRange.Cells(1, 2)

    nameCell.Value = "var_age"
    MasterSetupEventsManager.MsSheetChanged diseaseWksh, nameCell
    Assert.AreEqual "choice_age", diseaseTable.DataBodyRange.Cells(1, 5).Value, "The disease line should start with the default choice"

    Set variablesSheet = ThisWorkbook.Worksheets(VARIABLES_SHEET)
    Set variablesTable = variablesSheet.ListObjects(1)
    Set choiceCell = variablesTable.ListColumns("Default Choice").DataBodyRange.Cells(1, 1)

    choiceCell.Value = "choice_sex"
    MasterSetupEventsManager.MsSheetChanged variablesSheet, choiceCell

    Assert.AreEqual "choice_sex", diseaseTable.DataBodyRange.Cells(1, 5).Value, "The new default should land on the disease line"
    Assert.AreEqual "M | F", variablesTable.ListColumns("Choices Values").DataBodyRange.Cells(1, 1).Value, _
                    "The joined values should follow the new default"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestVariableDefaultChangePropagatesToDiseases", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupEvents")
'@sub-title Leaving the Choices sheet rebuilds the choices dropdown from the lists it carries.
Public Sub TestLeavingTheChoicesSheetRefreshesTheChoicesDropdown()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestLeavingTheChoicesSheetRefreshesTheChoicesDropdown"

    Dim choicesSheet As Worksheet
    Dim names As BetterArray

    On Error GoTo Fail

    Set names = MasterSetupEventsManager.MasterSetupService.Dropdowns.Values(CHOICES_LIST)
    Assert.IsFalse names Is Nothing, "The prepared fixture should carry the choices dropdown"
    Assert.IsFalse names.Includes("choice_new"), "The dropdown should start without the new list"

    'A new list typed under the others, the table grown over it.
    Set choicesSheet = ThisWorkbook.Worksheets(CHOICES_SHEET)
    choicesSheet.Range("A9:E9").Value = Array("choice_new", 1, "", "Yes", "Yes")
    choicesSheet.ListObjects(1).Resize choicesSheet.Range("A4:E9")

    MasterSetupEventsManager.MsSheetDeactivated choicesSheet

    Set names = MasterSetupEventsManager.MasterSetupService.Dropdowns.Values(CHOICES_LIST)
    Assert.IsTrue names.Includes("choice_new"), "Leaving the sheet should put the new list in the dropdown"
    Assert.IsTrue names.Includes("choice_age"), "The lists already there should stay"

    'Leaving any other sheet writes nothing.
    choicesSheet.Range("A10:E10").Value = Array("choice_late", 1, "", "No", "No")
    choicesSheet.ListObjects(1).Resize choicesSheet.Range("A4:E10")
    MasterSetupEventsManager.MsSheetDeactivated ThisWorkbook.Worksheets(VARIABLES_SHEET)

    Set names = MasterSetupEventsManager.MasterSetupService.Dropdowns.Values(CHOICES_LIST)
    Assert.IsFalse names.Includes("choice_late"), "Leaving the Variables sheet should leave the dropdown alone"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLeavingTheChoicesSheetRefreshesTheChoicesDropdown", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupEvents")
Public Sub TestEnsureLanguagesAddsAColumnToTheMasterTable()
    CustomTestSetTitles Assert, "MasterSetupEvents", "TestEnsureLanguagesAddsAColumnToTheMasterTable"

    Dim translationsSheet As Worksheet
    Dim table As ListObject
    Dim manager As SetupTranslationsTable

    On Error GoTo Fail

    'The Add Language door opens with this same Create over the master
    'table, so the fixture has to carry the shape the class wants.
    Set translationsSheet = ThisWorkbook.Worksheets(TRANSLATIONS_SHEET)
    Set table = translationsSheet.ListObjects(1)
    Set manager = SetupTranslationsTable.Create(table)

    manager.EnsureLanguages "ESP"

    Assert.AreEqual 3, table.ListColumns.Count, "The table should gain one language column"
    Assert.AreEqual "ESP", table.ListColumns(3).Name, "The new column should carry the language name"
    Assert.AreEqual "ENG", table.ListColumns(1).Name, "The default language column should stay first"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEnsureLanguagesAddsAColumnToTheMasterTable", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

Private Sub PrepareEnvironment()
    Dim targetSheet As Worksheet
    Dim data As Variant
    Dim preparation As MasterSetupPreparation

    'Variables sheet: the master table the managers read.
    Set targetSheet = EnsureWorksheet(VARIABLES_SHEET)
    ClearWorksheet targetSheet
    data = RowsToMatrix(Array( _
        Array("Variable Order", "Variable Section", "Variable Name", "Variable Label", "Default Choice", "Choices Values", "Default Status", "Comments"), _
        Array(1, "demographics", "var_age", "Age", "choice_age", "", "core", ""), _
        Array(2, "symptoms", "var_fever", "Fever", "", "", "core", "") _
    ))
    'The Choices Values column is text, so a joined value such as "0-4 | 5-14"
    'stays as typed.
    targetSheet.Range("F1:F3").NumberFormat = "@"
    WriteMatrix targetSheet.Range("A1"), data
    targetSheet.ListObjects.Add SourceType:=xlSrcRange, Source:=targetSheet.Range("A1").Resize(3, 8), _
                                XlListObjectHasHeaders:=xlYes

    'Choices sheet: headers at row 4, the way the master file carries them,
    'with the ordering list column: AllChoices lists nothing without it, and
    'the cascade then reads every default choice as gone.
    'The label column carries a formula, the way the setup choices sheet does.
    Set targetSheet = EnsureWorksheet(CHOICES_SHEET)
    ClearWorksheet targetSheet
    data = RowsToMatrix(Array( _
        Array("list name", "ordering list", "translated label", "label", "short label"), _
        Array("choice_age", 1, "", "0-4", "0-4"), _
        Array("choice_age", 2, "", "5-14", "5-14"), _
        Array("choice_sex", 1, "", "M", "M"), _
        Array("choice_sex", 2, "", "F", "F") _
    ))
    'The two typed label columns are text before the write: Excel reads
    '"5-14" as a date in a General cell. Column D keeps the General format,
    'since a text cell stores a formula as its own text.
    targetSheet.Range("C4:C8").NumberFormat = "@"
    targetSheet.Range("E4:E8").NumberFormat = "@"
    WriteMatrix targetSheet.Range("A4"), data
    targetSheet.ListObjects.Add SourceType:=xlSrcRange, Source:=targetSheet.Range("A4").Resize(5, 5), _
                                XlListObjectHasHeaders:=xlYes
    targetSheet.Range("D5:D8").Formula = "=IF(C5="""", E5, C5)"
    targetSheet.Calculate

    'Translations sheet: the language columns feed the languages dropdown.
    'The table sits one column right of its helper tag column, the way
    'SetupTranslationsTable wants it.
    Set targetSheet = EnsureWorksheet(TRANSLATIONS_SHEET)
    ClearWorksheet targetSheet
    targetSheet.Range("A1").Value = "__TagInternal__"
    data = RowsToMatrix(Array(Array("ENG", "FRA"), Array("Age", "Âge")))
    WriteMatrix targetSheet.Range("B1"), data
    targetSheet.ListObjects.Add SourceType:=xlSrcRange, Source:=targetSheet.Range("B1").Resize(2, 2), _
                                XlListObjectHasHeaders:=xlYes

    'The developer presses Initialize once before a file ships: the
    'dropdown tables and the Variables table take their shape here, and
    'the events then read prepared sheets.
    Set preparation = MasterSetupPreparation.Create(ThisWorkbook)
    preparation.EnsureDropdowns
    preparation.EnsureVariables
End Sub

'@description Build one disease sheet over the fixture managers.
Private Function BuildDiseaseFixture() As Worksheet
    Dim builder As DiseaseSheet

    Set builder = DiseaseSheet.Create(ThisWorkbook, _
                                      MasterSetupEventsManager.MasterSetupService.Dropdowns, _
                                      MasterSetupEventsManager.MasterSetupService.Variables)

    Set BuildDiseaseFixture = builder.Build(DISEASE_NAME)
End Function

Private Sub CleanupEnvironment()
    MasterSetupEventsManager.DisposeMasterSetup
    ResetMasterSetupFunctionCaches

    DeleteWorksheetSafe DISEASE_NAME
    DeleteWorksheetSafe VARIABLES_SHEET
    DeleteWorksheetSafe CHOICES_SHEET
    DeleteWorksheetSafe TRANSLATIONS_SHEET
    DeleteWorksheetSafe DROPDOWNS_SHEET
    DeleteWorksheetSafe PASSWORDS_SHEET

    DeleteNameSafe "__Col__Variables"
    DeleteNameSafe "DISSHEET001"
    DeleteNameSafe "DISSHEET002"

    'The names the passwords fixture seeds.
    DeleteNameSafe "RNG_PublicKey"
    DeleteNameSafe "RNG_PrivateKey"
    DeleteNameSafe "RNG_DebuggingPassword"
    DeleteNameSafe "RNG_DebugMode"
    DeleteNameSafe "RNG_Version"
    DeleteNameSafe "RNG_LabPublicKey"
    DeleteNameSafe "RNG_LabPrivateKey"
    DeleteNameSafe "Passwords_ProtectedSheets"
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
