Attribute VB_Name = "TestMasterSetupPreparation"
Attribute VB_Description = "Unit tests for the MasterSetupPreparation orchestration helper"

Option Explicit

'@Folder("CustomTests.MasterSetup")
'@ModuleDescription("Validates dropdown registration and variables initialisation for MasterSetupPreparation")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

Private Assert As CustomTest
Private Subject As MasterSetupPreparation
Private FixtureWorkbook As Workbook
Private DropdownSheet As Worksheet
Private VariablesSheet As Worksheet
Private TranslationsSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const DROPDOWNS_SHEET_NAME As String = "__dropdowns"
Private Const VARIABLES_SHEET_NAME As String = "Variables"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const STATUS_DROPDOWN As String = "__var_status"
Private Const YESNO_DROPDOWN As String = "__yesno"
Private Const LANGUAGES_DROPDOWN As String = "__data_languages"
Private Const VARIABLE_COLUMN_NAME As String = "__Col__Variables"
Private Const CHOICES_DROPDOWN As String = "__lst_choices"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestMasterSetupPreparation"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    RestoreApp
End Sub


'@section Legacy adoption
'===============================================================================
'@TestMethod("MasterSetupPreparation")
Public Sub TestPrepareAdoptsLegacyDiseaseSheets()
    CustomTestSetTitles Assert, "MasterSetupPreparation", "TestPrepareAdoptsLegacyDiseaseSheets"

    Dim legacySheet As Worksheet
    Dim store As HiddenNames

    On Error GoTo Fail

    Set legacySheet = EnsureWorksheet("LegacyDisease", FixtureWorkbook)
    legacySheet.Cells(2, 2).Value = "ENG"
    legacySheet.Cells(2, 3).Value = "DISSHEET004"
    legacySheet.Cells(2, 4).Value = "DISSHEET"

    Subject.Prepare

    Set store = HiddenNames.Create(legacySheet)
    Assert.AreEqual "disease", store.ValueAsString("sheetTag"), "The legacy marker should become the hidden tag"
    Assert.AreEqual "ENG", store.ValueAsString("__Var_DISLANG"), "The language should move into the hidden names"
    Assert.AreEqual "DISSHEET004", store.ValueAsString("__Var_DISCODE"), "The code should move into the hidden names"
    Assert.AreEqual 4&, store.ValueAsLong("__Var_DISINDEX", 0), "The index should parse off the code"
    Assert.AreEqual vbNullString, CStr(legacySheet.Cells(2, 4).Value), "The marker cell should retire"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestPrepareAdoptsLegacyDiseaseSheets", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupPreparation")
Public Sub TestPrepareGrowsTheReportSheetVeryHidden()
    CustomTestSetTitles Assert, "MasterSetupPreparation", "TestPrepareGrowsTheReportSheetVeryHidden"

    Dim reportSheet As Worksheet

    On Error GoTo Fail

    Subject.Prepare

    Set reportSheet = FixtureSheet("__compRep")
    Assert.IsFalse reportSheet Is Nothing, "A workbook without __compRep should grow it on preparation"
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(reportSheet.Visible), "The report sheet should be very hidden"

    'A second preparation leaves the sheet as it stands.
    reportSheet.Visible = xlSheetVisible
    Subject.Prepare
    Assert.AreEqual CLng(xlSheetVisible), CLng(reportSheet.Visible), "An existing report sheet should keep its visibility"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestPrepareGrowsTheReportSheetVeryHidden", Err.Number, Err.Description
End Sub

'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    BusyApp

    Set FixtureWorkbook = NewWorkbook
    Set DropdownSheet = EnsureWorksheet(DROPDOWNS_SHEET_NAME, FixtureWorkbook)
    Set VariablesSheet = EnsureWorksheet(VARIABLES_SHEET_NAME, FixtureWorkbook)
    Set TranslationsSheet = EnsureWorksheet(TRANSLATIONS_SHEET_NAME, FixtureWorkbook)

    PrepareTranslationsFixture TranslationsSheet

    Set Subject = MasterSetupPreparation.Create(FixtureWorkbook)
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set Subject = Nothing
    Set TranslationsSheet = Nothing
    Set VariablesSheet = Nothing
    Set DropdownSheet = Nothing
    Set FixtureWorkbook = Nothing

    RestoreApp
End Sub


'@section Tests
'===============================================================================
'@TestMethod("MasterSetupPreparation")
Public Sub TestPrepareRegistersDropdowns()
    CustomTestSetTitles Assert, "MasterSetupPreparation", "TestPrepareRegistersDropdowns"
    On Error GoTo Fail

    Subject.Prepare Application

    Dim statuses As BetterArray
    Dim diseases As BetterArray

    Set statuses = Subject.Dropdowns.Values(STATUS_DROPDOWN)
    Assert.IsFalse statuses Is Nothing, "Status dropdown should be registered"
    Assert.IsTrue ContainsValue(statuses, "active"), "Status dropdown should contain 'active'"
    Assert.IsTrue ContainsValue(statuses, "inactive"), "Status dropdown should contain 'inactive'"

    Set diseases = Subject.Dropdowns.Values("__diseases_list")
    Assert.IsFalse diseases Is Nothing, "Diseases dropdown should be registered"
    Assert.IsFalse ContainsValue(diseases, "Variables"), "The diseases list carries tagged sheets alone"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareRegistersDropdowns"
End Sub

'@TestMethod("MasterSetupPreparation")
Public Sub TestPrepareInitialisesVariablesTable()
    CustomTestSetTitles Assert, "MasterSetupPreparation", "TestPrepareInitialisesVariablesTable"
    On Error GoTo Fail

    Subject.Prepare Application

    Dim manager As MasterSetupVariables
    Dim table As ListObject
    Dim statusRange As Range

    Set manager = Subject.Variables
    Assert.IsFalse manager Is Nothing, "Variables manager should be created"
    Assert.IsTrue manager.Initialised, "Variables manager should be initialised after preparation"

    Set table = manager.Table
    Assert.IsFalse table Is Nothing, "Variables table should exist after preparation"
    Assert.AreEqual 8&, table.ListColumns.Count, "Variables table should expose the eight expected columns"
    Assert.AreEqual "Default Status", table.ListColumns(7).Name, "Default Status column should exist"

    Set statusRange = table.ListColumns("Default Status").DataBodyRange
    Assert.IsFalse statusRange Is Nothing, "Default Status column should expose a data range"
    Assert.AreEqual xlValidateList, statusRange.Validation.Type, "Default Status should apply list validation"
    Assert.IsTrue InStr(1, statusRange.Validation.Formula1, STATUS_DROPDOWN, vbTextCompare) > 0, _
                 "Default Status validation should reference the status dropdown"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareInitialisesVariablesTable"
End Sub

'@TestMethod("MasterSetupPreparation")
Public Sub TestEnsureVariablesPublishesWorkbookRange()
    Dim manager As MasterSetupVariables
    Dim definedName As Name
    Dim expectedRange As Range
    Dim actualRange As Range

    CustomTestSetTitles Assert, "MasterSetupPreparation", "TestEnsureVariablesPublishesWorkbookRange"
    On Error GoTo Fail

    Subject.EnsureVariables

    Set manager = Subject.Variables
    Set expectedRange = manager.DataRange("Variable Name")

    Assert.IsFalse expectedRange Is Nothing, "Variables manager should expose the Variable Name data range."

    On Error Resume Next
        Set definedName = FixtureWorkbook.Names(VARIABLE_COLUMN_NAME)
    On Error GoTo 0
    On Error GoTo Fail

    Assert.IsFalse definedName Is Nothing, "Workbook hidden name should be created for the Variable Name column."

    Set actualRange = definedName.RefersToRange
    Assert.IsFalse actualRange Is Nothing, "Workbook hidden name should resolve to a valid range."
    Assert.AreEqual expectedRange.Address(True, True, xlA1, True), _
                     actualRange.Address(True, True, xlA1, True), _
                     "Workbook name should target the same cells as the manager data range."
    Exit Sub

Fail:
    ReportTestFailure "TestEnsureVariablesPublishesWorkbookRange"
End Sub

'@TestMethod("MasterSetupPreparation")
'@sub-title A second Prepare over an initialised workbook leaves one duplicate rule over the grown name column.
Public Sub TestPrepareTwiceKeepsOneDuplicateRuleOverTheNameColumn()
    Dim table As ListObject
    Dim nameRange As Range
    Dim ruleRange As Range

    CustomTestSetTitles Assert, "MasterSetupPreparation", "TestPrepareTwiceKeepsOneDuplicateRuleOverTheNameColumn"
    On Error GoTo Fail

    Subject.Prepare Application
    Assert.IsTrue Subject.Variables.Initialised, "The first preparation should initialise the variables table"

    'The table grows between the two preparations, the way a master setup
    'grows between two saves.
    Set table = Subject.Variables.Table
    table.ListRows.Add
    table.ListRows.Add

    Subject.Prepare Application

    Set nameRange = Subject.Variables.DataRange("Variable Name")
    Assert.IsFalse nameRange Is Nothing, "The name column should expose a data range after the second preparation"
    Assert.AreEqual 1&, nameRange.FormatConditions.Count, "The second preparation should leave exactly one rule on the name column"

    Set ruleRange = nameRange.FormatConditions(1).AppliesTo
    Assert.AreEqual nameRange.Address(False, False), ruleRange.Address(False, False), _
                    "The duplicate rule should cover the whole name column as it stands"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareTwiceKeepsOneDuplicateRuleOverTheNameColumn"
End Sub

'@TestMethod("MasterSetupPreparation")
'@sub-title An initialised table that lost its Default Choice dropdown gets it back on the next preparation.
'@details The Initialised flag is a hidden name that survives every save,
'so the validation written once by Initialise stood for good. The refresh
'runs outside the flag, the way the duplicate rule does.
Public Sub TestPrepareTwiceGivesTheDefaultChoiceColumnItsDropdown()
    Dim choiceRange As Range

    CustomTestSetTitles Assert, "MasterSetupPreparation", "TestPrepareTwiceGivesTheDefaultChoiceColumnItsDropdown"
    On Error GoTo Fail

    Subject.Prepare Application
    Assert.IsTrue Subject.Variables.Initialised, "The first preparation should initialise the variables table"

    'The table stands initialised and its column has no dropdown, the way
    'a table initialised over an empty choices list stood on the mock.
    Subject.Variables.DataRange("Default Choice").Validation.Delete
    Assert.IsFalse ColumnHasListValidation(Subject.Variables.DataRange("Default Choice")), _
                   "The Default Choice column should start this test without a dropdown"

    Subject.Prepare Application

    Set choiceRange = Subject.Variables.DataRange("Default Choice")
    Assert.IsTrue ColumnHasListValidation(choiceRange), "The second preparation should put the dropdown on the Default Choice column"
    Assert.IsTrue InStr(1, choiceRange.Validation.Formula1, CHOICES_DROPDOWN, vbTextCompare) > 0, _
                  "The Default Choice dropdown should reference the choices list"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareTwiceGivesTheDefaultChoiceColumnItsDropdown"
End Sub

'@TestMethod("MasterSetupPreparation")
Public Sub TestEnsureDropdownsLoadsLanguages()
    CustomTestSetTitles Assert, "MasterSetupPreparation", "TestEnsureDropdownsLoadsLanguages"
    On Error GoTo Fail

    Subject.EnsureDropdowns

    Dim languages As BetterArray
    Set languages = Subject.Dropdowns.Values(LANGUAGES_DROPDOWN)

    Assert.IsFalse languages Is Nothing, "Languages dropdown should exist after EnsureDropdowns"
    Assert.IsTrue ContainsValue(languages, "en"), "Languages dropdown should include English header from translations table"
    Assert.IsTrue ContainsValue(languages, "fr"), "Languages dropdown should include French header from translations table"
    Exit Sub

Fail:
    ReportTestFailure "TestEnsureDropdownsLoadsLanguages"
End Sub


'@section Helpers
'===============================================================================
Private Sub PrepareTranslationsFixture(ByVal wsTrans As Worksheet)
    Dim lo As ListObject

    If wsTrans Is Nothing Then Exit Sub

    wsTrans.Cells.Clear
    wsTrans.Range("A1").Value = "key"
    wsTrans.Range("B1").Value = "en"
    wsTrans.Range("C1").Value = "fr"
    wsTrans.Range("A2").Value = "greeting"
    wsTrans.Range("B2").Value = "Hello"
    wsTrans.Range("C2").Value = "Bonjour"

    Set lo = wsTrans.ListObjects.Add(xlSrcRange, wsTrans.Range("A1:C2"), , xlYes)
    lo.Name = "TST_MasterTranslations"
End Sub

'@description Whether a range carries a list validation over all its cells.
Private Function ColumnHasListValidation(ByVal target As Range) As Boolean
    Dim validationType As Long

    If target Is Nothing Then Exit Function

    'A range with no validation raises on the read; the raise is the answer.
    On Error Resume Next
        validationType = target.Validation.Type
        ColumnHasListValidation = (Err.Number = 0)
        Err.Clear
    On Error GoTo 0

    If ColumnHasListValidation Then ColumnHasListValidation = (validationType = xlValidateList)
End Function

'@description A worksheet of the fixture workbook by name, or Nothing.
Private Function FixtureSheet(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    For Each sh In FixtureWorkbook.Worksheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            Set FixtureSheet = sh
            Exit Function
        End If
    Next sh
End Function

Private Function ContainsValue(ByVal items As BetterArray, ByVal expected As String) As Boolean
    Dim idx As Long
    Dim candidate As Variant

    If items Is Nothing Then Exit Function

    For idx = items.LowerBound To items.UpperBound
        candidate = items.Item(idx)
        If NormalizeText(CStr(candidate)) = NormalizeText(expected) Then
            ContainsValue = True
            Exit Function
        End If
    Next idx
End Function

Private Function NormalizeText(ByVal valueText As String) As String
    NormalizeText = LCase$(Trim$(valueText))
End Function

Private Sub ReportTestFailure(ByVal context As String)
    Dim message As String

    If Assert Is Nothing Then Exit Sub

    message = context & " failed with error " & Err.Number & " (" & Err.Source & "): " & Err.Description
    Assert.LogFailure message
    Err.Clear
End Sub
