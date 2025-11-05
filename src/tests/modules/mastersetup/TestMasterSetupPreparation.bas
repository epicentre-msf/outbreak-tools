Attribute VB_Name = "TestMasterSetupPreparation"
Attribute VB_Description = "Unit tests for the MasterSetupPreparation orchestration helper"

Option Explicit

'@Folder("CustomTests.MasterSetup")
'@ModuleDescription("Validates dropdown registration and variables initialisation for MasterSetupPreparation")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

Private Assert As ICustomTest
Private Subject As IMasterSetupPreparation
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
Private Const LANGUAGES_DROPDOWN As String = "__languages"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    TestHelpers.BusyApp
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
    Set TranslationsSheet = TestHelpers.EnsureWorksheet(TRANSLATIONS_SHEET_NAME, FixtureWorkbook)

    PrepareTranslationsFixture TranslationsSheet

    Set Subject = MasterSetupPreparation.Create(FixtureWorkbook)
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        TestHelpers.DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set Subject = Nothing
    Set TranslationsSheet = Nothing
    Set VariablesSheet = Nothing
    Set DropdownSheet = Nothing
    Set FixtureWorkbook = Nothing

    TestHelpers.RestoreApp
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
    Assert.IsTrue ContainsValue(diseases, "Variables"), "Diseases dropdown should include core sheets"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareRegistersDropdowns"
End Sub

'@TestMethod("MasterSetupPreparation")
Public Sub TestPrepareInitialisesVariablesTable()
    CustomTestSetTitles Assert, "MasterSetupPreparation", "TestPrepareInitialisesVariablesTable"
    On Error GoTo Fail

    Subject.Prepare Application

    Dim manager As IMasterSetupVariables
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
