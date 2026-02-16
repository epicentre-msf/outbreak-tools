Attribute VB_Name = "TestEventMasterSetup"
Attribute VB_Description = "Unit tests for the EventMasterSetup orchestrator"

Option Explicit

'@Folder("CustomTests.MasterSetup")
'@ModuleDescription("Validates master setup event scaffolding covering dropdowns, variables, choices, and translations.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

Private Assert As ICustomTest
Private Subject As IEventMasterSetup
Private FixtureWorkbook As Workbook
Private DropdownSheet As Worksheet
Private VariablesSheet As Worksheet
Private ChoicesSheet As Worksheet
Private TranslationsSheet As Worksheet
Private RibbonSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const DROPDOWNS_SHEET_NAME As String = "__dropdowns"
Private Const VARIABLES_SHEET_NAME As String = "Variables"
Private Const CHOICES_SHEET_NAME As String = "Choices"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const RIBBON_TRANSLATION_SHEET As String = "__ribbonTranslation"
Private Const TRANSLATION_TABLE_NAME As String = "Tab_Translations"
Private Const RIBBON_TABLE_NAME As String = "TabTransId"
Private Const LANGUAGE_RANGE_NAME As String = "RNG_FileLang"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    TestHelpers.BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestEventMasterSetup"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then Assert.PrintResults TEST_OUTPUT_SHEET
    On Error GoTo 0
    Set Assert = Nothing
    TestHelpers.RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    TestHelpers.BusyApp

    Set FixtureWorkbook = TestHelpers.NewWorkbook
    Set DropdownSheet = TestHelpers.EnsureWorksheet(DROPDOWNS_SHEET_NAME, FixtureWorkbook)
    Set VariablesSheet = TestHelpers.EnsureWorksheet(VARIABLES_SHEET_NAME, FixtureWorkbook)
    Set ChoicesSheet = TestHelpers.EnsureWorksheet(CHOICES_SHEET_NAME, FixtureWorkbook)
    Set TranslationsSheet = TestHelpers.EnsureWorksheet(TRANSLATIONS_SHEET_NAME, FixtureWorkbook)
    Set RibbonSheet = TestHelpers.EnsureWorksheet(RIBBON_TRANSLATION_SHEET, FixtureWorkbook)

    PrepareChoicesFixture ChoicesSheet
    PrepareTranslationsFixture TranslationsSheet
    PrepareRibbonFixture RibbonSheet

    Set Subject = EventMasterSetup.Create(FixtureWorkbook)
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then Assert.Flush

    On Error Resume Next
        TestHelpers.DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set Subject = Nothing
    Set RibbonSheet = Nothing
    Set TranslationsSheet = Nothing
    Set ChoicesSheet = Nothing
    Set VariablesSheet = Nothing
    Set DropdownSheet = Nothing
    Set FixtureWorkbook = Nothing

    TestHelpers.RestoreApp
End Sub


'@section Tests
'===============================================================================
'@TestMethod("EventMasterSetup")
Public Sub TestCreateRequiresWorkbook()
    CustomTestSetTitles Assert, "EventMasterSetup", "TestCreateRequiresWorkbook"

    On Error Resume Next
        Dim target As IEventMasterSetup
        Set target = EventMasterSetup.Create(Nothing)
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, "Create should require a workbook reference"
    Err.Clear
    On Error GoTo 0
End Sub

'@TestMethod("EventMasterSetup")
Public Sub TestOnWorkbookOpenInitialisesDependencies()
    CustomTestSetTitles Assert, "EventMasterSetup", "TestOnWorkbookOpenInitialisesDependencies"
    On Error GoTo Fail

    Subject.OnWorkbookOpen Application

    Dim drops As IDropdownLists
    Dim vars As IMasterSetupVariables
    Dim choices As ILLChoices
    Dim trads As ITranslationObject
    Dim ribTrads As ITranslationObject

    Set drops = Subject.Dropdowns
    Assert.IsFalse drops Is Nothing, "Dropdown manager should be resolved"

    Set vars = Subject.Variables
    Assert.IsFalse vars Is Nothing, "Variables manager should be cached"
    Assert.IsTrue vars.Initialised, "Variables manager should be initialised after OnWorkbookOpen"

    Set choices = Subject.Choices
    Assert.IsFalse choices Is Nothing, "Choices helper should be created"

    Set trads = Subject.Translations
    Assert.IsFalse trads Is Nothing, "Workbook translations should be resolved"

    Set ribTrads = Subject.RibbonTranslations
    Assert.IsFalse ribTrads Is Nothing, "Ribbon translations should be resolved"
    Exit Sub

Fail:
    ReportTestFailure "TestOnWorkbookOpenInitialisesDependencies"
End Sub

'@TestMethod("EventMasterSetup")
Public Sub TestRefreshTranslationsResetsCaches()
    CustomTestSetTitles Assert, "EventMasterSetup", "TestRefreshTranslationsResetsCaches"
    On Error GoTo Fail

    Subject.OnWorkbookOpen Application

    Dim firstWorkbookTrads As ITranslationObject
    Dim firstRibbonTrads As ITranslationObject
    Dim firstChoices As ILLChoices

    Set firstWorkbookTrads = Subject.Translations
    Set firstRibbonTrads = Subject.RibbonTranslations
    Set firstChoices = Subject.Choices

    Subject.RefreshTranslations

    Assert.IsFalse firstWorkbookTrads Is Subject.Translations, "General translations should refresh after RefreshTranslations"
    Assert.IsFalse firstRibbonTrads Is Subject.RibbonTranslations, "Ribbon translations should refresh after RefreshTranslations"
    Assert.IsFalse firstChoices Is Subject.Choices, "Choices helper should refresh when translations change"
    Exit Sub

Fail:
    ReportTestFailure "TestRefreshTranslationsResetsCaches"
End Sub


'@section Helpers
'===============================================================================
Private Sub PrepareChoicesFixture(ByVal targetSheet As Worksheet)
    If targetSheet Is Nothing Then Exit Sub

    targetSheet.Cells.Clear
    targetSheet.Range("A4").Value = "list name"
    targetSheet.Range("B4").Value = "label"
    targetSheet.Range("C4").Value = "short label"
    targetSheet.Range("D4").Value = "ordering list"
    targetSheet.Range("A5").Value = "demo_list"
    targetSheet.Range("B5").Value = "Demo label"
    targetSheet.Range("C5").Value = "Demo"
    targetSheet.Range("D5").Value = 1
End Sub

Private Sub PrepareTranslationsFixture(ByVal targetSheet As Worksheet)
    Dim lo As ListObject

    If targetSheet Is Nothing Then Exit Sub

    targetSheet.Cells.Clear
    targetSheet.Range("A1").Value = "key"
    targetSheet.Range("B1").Value = "en"
    targetSheet.Range("C1").Value = "fr"
    targetSheet.Range("A2").Value = "askConfirm"
    targetSheet.Range("B2").Value = "Confirm"
    targetSheet.Range("C2").Value = "Confirmer"

    Set lo = targetSheet.ListObjects.Add(xlSrcRange, targetSheet.Range("A1:C2"), , xlYes)
    lo.Name = TRANSLATION_TABLE_NAME
End Sub

Private Sub PrepareRibbonFixture(ByVal targetSheet As Worksheet)
    Dim lo As ListObject
    Dim langCell As Range

    If targetSheet Is Nothing Then Exit Sub

    targetSheet.Cells.Clear
    targetSheet.Range("A1").Value = "key"
    targetSheet.Range("B1").Value = "en"
    targetSheet.Range("C1").Value = "fr"
    targetSheet.Range("A2").Value = "askConfirm"
    targetSheet.Range("B2").Value = "Confirm"
    targetSheet.Range("C2").Value = "Confirmer"

    Set lo = targetSheet.ListObjects.Add(xlSrcRange, targetSheet.Range("A1:C2"), , xlYes)
    lo.Name = RIBBON_TABLE_NAME

    Set langCell = targetSheet.Range("E2")
    langCell.Value = "en"

    FixtureWorkbook.Names.Add Name:=LANGUAGE_RANGE_NAME, RefersTo:=langCell
End Sub

Private Sub ReportTestFailure(ByVal context As String)
    If Assert Is Nothing Then Exit Sub
    Assert.LogFailure context & " failed with error " & Err.Number & ": " & Err.Description
    Err.Clear
End Sub
