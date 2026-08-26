Attribute VB_Name = "TestEventMasterSetup"
Attribute VB_Description = "Unit tests for the EventMasterSetup orchestrator"

Option Explicit

'@Folder("CustomTests.MasterSetup")
'@ModuleDescription("Validates master setup event scaffolding covering dropdowns, variables, choices, and translations.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

Private Assert As CustomTest
Private Subject As EventMasterSetup
Private FixtureWorkbook As Workbook
Private DropdownSheet As Worksheet
Private VariablesSheet As Worksheet
Private ChoicesSheet As Worksheet
Private TranslationsSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const DROPDOWNS_SHEET_NAME As String = "__dropdowns"
Private Const VARIABLES_SHEET_NAME As String = "Variables"
Private Const CHOICES_SHEET_NAME As String = "Choices"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const TRANSLATION_TABLE_NAME As String = "Tab_Translations"
Private Const VARIABLES_TABLE_NAME As String = "Tab_Variables"
Private Const REGISTRY_SHEET_NAME As String = "__updated"
Private Const PASSWORDS_SHEET_NAME As String = "__pass"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestEventMasterSetup"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then Assert.PrintResults TEST_OUTPUT_SHEET
    On Error GoTo 0
    Set Assert = Nothing
    RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    BusyApp

    Set FixtureWorkbook = NewWorkbook
    Set DropdownSheet = EnsureWorksheet(DROPDOWNS_SHEET_NAME, FixtureWorkbook)
    Set VariablesSheet = EnsureWorksheet(VARIABLES_SHEET_NAME, FixtureWorkbook)
    Set ChoicesSheet = EnsureWorksheet(CHOICES_SHEET_NAME, FixtureWorkbook)
    Set TranslationsSheet = EnsureWorksheet(TRANSLATIONS_SHEET_NAME, FixtureWorkbook)

    PrepareVariablesFixture VariablesSheet
    PrepareChoicesFixture ChoicesSheet
    PrepareTranslationsFixture TranslationsSheet

    Set Subject = EventMasterSetup.Create(FixtureWorkbook)
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then Assert.Flush

    On Error Resume Next
        DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set Subject = Nothing
    Set TranslationsSheet = Nothing
    Set ChoicesSheet = Nothing
    Set VariablesSheet = Nothing
    Set DropdownSheet = Nothing
    Set FixtureWorkbook = Nothing

    RestoreApp
End Sub


'@section Tests
'===============================================================================
'@TestMethod("EventMasterSetup")
Public Sub TestCreateRequiresWorkbook()
    CustomTestSetTitles Assert, "EventMasterSetup", "TestCreateRequiresWorkbook"

    On Error Resume Next
        Dim target As EventMasterSetup
        Set target = EventMasterSetup.Create(Nothing)
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, "Create should require a workbook reference"
    Err.Clear
    On Error GoTo 0
End Sub

'@TestMethod("EventMasterSetup")
'@sub-title After the open, the managers are built from the sheets and held across two reads.
Public Sub TestOnWorkbookOpenInitialisesDependencies()
    CustomTestSetTitles Assert, "EventMasterSetup", "TestOnWorkbookOpenInitialisesDependencies"
    On Error GoTo Fail

    Subject.OnWorkbookOpen Application

    Dim drops As DropdownLists
    Dim vars As MasterSetupVariables
    Dim choices As LLChoices

    Set drops = Subject.Dropdowns
    Assert.IsFalse drops Is Nothing, "Dropdown manager should be resolved"
    Assert.IsTrue drops Is Subject.Dropdowns, "The dropdown manager is held across two reads"

    Set vars = Subject.Variables
    Assert.IsFalse vars Is Nothing, "Variables manager should be resolved"
    Assert.IsTrue vars Is Subject.Variables, "The variables manager is held across two reads"

    Set choices = Subject.Choices
    Assert.IsFalse choices Is Nothing, "Choices helper should be created"

    Exit Sub

Fail:
    ReportTestFailure "TestOnWorkbookOpenInitialisesDependencies"
End Sub

'@TestMethod("EventMasterSetup")
'@sub-title The open writes nothing on the sheets: the Variables table stays uninitialised.
Public Sub TestOnWorkbookOpenLeavesTheSheetsAlone()
    CustomTestSetTitles Assert, "EventMasterSetup", "TestOnWorkbookOpenLeavesTheSheetsAlone"
    On Error GoTo Fail

    Subject.OnWorkbookOpen Application

    Assert.IsFalse Subject.Variables.Initialised, "The open runs no preparation over the Variables table"
    Assert.AreEqual 0&, DropdownSheet.ListObjects.Count, "The open writes no dropdown table"

    Exit Sub

Fail:
    ReportTestFailure "TestOnWorkbookOpenLeavesTheSheetsAlone"
End Sub

'@TestMethod("EventMasterSetup")
'@sub-title The open switches every flag of the __updated registry to yes, the way the setup does.
Public Sub TestOnWorkbookOpenSwitchesTheRegistryTagsToYes()
    CustomTestSetTitles Assert, "EventMasterSetup", "TestOnWorkbookOpenSwitchesTheRegistryTagsToYes"

    Dim registryTable As ListObject
    Dim flagCell As Range

    On Error GoTo Fail

    'A registry built over one tagged column, its flag at no, the way a
    'file stands after an Update Translations.
    ChoicesSheet.Range("B3").Value = "translate as text"
    ChoicesSheet.ListObjects.Add(xlSrcRange, ChoicesSheet.Range("A4:D5"), , xlYes).Name = "Tab_Choices"
    MasterSetupPreparation.Create(FixtureWorkbook).EnsureUpdatedRegistry

    Set registryTable = FixtureWorkbook.Worksheets(REGISTRY_SHEET_NAME).ListObjects(1)
    Set flagCell = registryTable.ListColumns("updated").DataBodyRange.Cells(1, 1)
    Subject.Updater.SwitchTagsToNo
    Assert.AreEqual "no", LCase$(Trim$(CStr(flagCell.Value))), "The flag should start this test at no"

    Subject.OnWorkbookOpen Application

    Assert.AreEqual "yes", LCase$(Trim$(CStr(flagCell.Value))), "The open should switch the flag to yes"

    Exit Sub

Fail:
    ReportTestFailure "TestOnWorkbookOpenSwitchesTheRegistryTagsToYes"
End Sub

'@TestMethod("EventMasterSetup")
'@sub-title The open ends a debug session and puts the protection matrix back.
'@details A file saved in debug mode reopens with its flag set and its sheets
'open; the open clears the flag and protects every sheet the table lists.
Public Sub TestOnWorkbookOpenLeavesDebugModeAndProtects()
    CustomTestSetTitles Assert, "EventMasterSetup", "TestOnWorkbookOpenLeavesDebugModeAndProtects"

    Dim passSheet As Worksheet
    Dim pass As Passwords

    On Error GoTo Fail

    PreparePasswordsFixture PASSWORDS_SHEET_NAME, FixtureWorkbook
    Set passSheet = FixtureWorkbook.Worksheets(PASSWORDS_SHEET_NAME)
    Set pass = Passwords.Create(passSheet)
    pass.DisplayPrompts = False

    'The Variables sheet is listed in the protection table, then the
    'workbook goes into debug mode: every sheet open, the flag set.
    pass.Protect VariablesSheet.Name
    pass.EnterDebugMode FixtureWorkbook
    Assert.IsFalse VariablesSheet.ProtectContents, "Debug mode should leave the Variables sheet open"
    Assert.IsTrue pass.IsInDebugMode, "The debug flag should be set before the open"

    Subject.OnWorkbookOpen Application

    Assert.IsTrue VariablesSheet.ProtectContents, "The open should protect the Variables sheet again"
    Set pass = Passwords.Create(passSheet)
    Assert.IsFalse pass.IsInDebugMode, "The open should clear the debug flag"

    Exit Sub

Fail:
    ReportTestFailure "TestOnWorkbookOpenLeavesDebugModeAndProtects"
End Sub

'@TestMethod("EventMasterSetup")
'@sub-title A workbook missing the sheets or the table answers Nothing for its managers.
Public Sub TestManagersAnswerNothingWithoutTheirSheets()
    CustomTestSetTitles Assert, "EventMasterSetup", "TestManagersAnswerNothingWithoutTheirSheets"
    On Error GoTo Fail

    VariablesSheet.ListObjects(1).Unlist
    DropdownSheet.Delete

    Assert.IsTrue Subject.Variables Is Nothing, "A Variables sheet without a table answers no manager"
    Assert.IsTrue Subject.Dropdowns Is Nothing, "A workbook without __dropdowns answers no manager"

    Exit Sub

Fail:
    ReportTestFailure "TestManagersAnswerNothingWithoutTheirSheets"
End Sub

'@TestMethod("EventMasterSetup")
Public Sub TestRefreshTranslationsResetsCaches()
    CustomTestSetTitles Assert, "EventMasterSetup", "TestRefreshTranslationsResetsCaches"
    On Error GoTo Fail

    Subject.OnWorkbookOpen Application

    Dim firstChoices As LLChoices
    Dim firstVariables As MasterSetupVariables

    Set firstChoices = Subject.Choices
    Set firstVariables = Subject.Variables

    Subject.RefreshTranslations

    Assert.IsFalse firstChoices Is Subject.Choices, "Choices helper should refresh when translations change"
    Assert.IsFalse firstVariables Is Subject.Variables, "The variables manager is rebuilt after a refresh"
    Exit Sub

Fail:
    ReportTestFailure "TestRefreshTranslationsResetsCaches"
End Sub


'@section Helpers
'===============================================================================
'@description The eight columns of the master table, one blank line under them.
Private Sub PrepareVariablesFixture(ByVal targetSheet As Worksheet)
    Dim lo As ListObject

    If targetSheet Is Nothing Then Exit Sub

    targetSheet.Cells.Clear
    targetSheet.Range("A1:H1").Value = Array("Variable Order", "Variable Section", "Variable Name", "Label", _
                                             "Default Choice", "Choices Values", "Default Status", "Comments")

    Set lo = targetSheet.ListObjects.Add(xlSrcRange, targetSheet.Range("A1:H2"), , xlYes)
    lo.Name = VARIABLES_TABLE_NAME
End Sub

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

    'The table sits one column right of its helper tag column, the way
    'SetupTranslationsTable wants it; the headers are the languages.
    targetSheet.Cells.Clear
    targetSheet.Range("A1").Value = "__TagInternal__"
    targetSheet.Range("B1").Value = "en"
    targetSheet.Range("C1").Value = "fr"
    targetSheet.Range("B2").Value = "Confirm"
    targetSheet.Range("C2").Value = "Confirmer"

    Set lo = targetSheet.ListObjects.Add(xlSrcRange, targetSheet.Range("B1:C2"), , xlYes)
    lo.Name = TRANSLATION_TABLE_NAME
End Sub

Private Sub ReportTestFailure(ByVal context As String)
    If Assert Is Nothing Then Exit Sub
    Assert.LogFailure context & " failed with error " & Err.Number & ": " & Err.Description
    Err.Clear
End Sub
