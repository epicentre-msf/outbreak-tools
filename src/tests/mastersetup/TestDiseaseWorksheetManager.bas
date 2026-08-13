Attribute VB_Name = "TestDiseaseWorksheetManager"
Attribute VB_Description = "Tests ensuring DiseaseWorksheetManager safely removes worksheets"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests ensuring DiseaseWorksheetManager safely removes worksheets")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TEMP_SHEET_NAME As String = "DiseaseRemovalFixture"

Private Assert As CustomTest
Private Manager As IDiseaseWorksheetManager

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseWorksheetManager"
    Set Manager = New DiseaseWorksheetManager
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        DeleteWorksheet TEMP_SHEET_NAME
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
    Set Manager = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    PrepareFixtureSheet
End Sub

'@TestCleanup
Private Sub TestCleanup()
    DeleteWorksheet TEMP_SHEET_NAME
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseWorksheetManager")
Public Sub TestRemoveWorksheetDeletesSheetOnce()
    CustomTestSetTitles Assert, "DiseaseWorksheetManager", "TestRemoveWorksheetDeletesSheetOnce"

    Dim originalCount As Long
    Dim removed As Boolean
    Dim targetBook As Workbook
    Dim originalAlerts As Boolean

    On Error GoTo Fail

    Set targetBook = ThisWorkbook
    originalAlerts = targetBook.Application.DisplayAlerts
    originalCount = targetBook.Worksheets.Count

    removed = Manager.RemoveWorksheet(targetBook, TEMP_SHEET_NAME)

    Assert.IsTrue removed, "RemoveWorksheet should return True when sheet exists"
    Assert.IsFalse WorksheetExists(TEMP_SHEET_NAME, targetBook), "Worksheet should be removed"
    Assert.AreEqual originalCount - 1, targetBook.Worksheets.Count, "Worksheet count should decrease by one"
    Assert.AreEqual originalAlerts, targetBook.Application.DisplayAlerts, "DisplayAlerts should be restored"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveWorksheetDeletesSheetOnce", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseWorksheetManager")
Public Sub TestRemoveWorksheetReturnsFalseWhenMissing()
    CustomTestSetTitles Assert, "DiseaseWorksheetManager", "TestRemoveWorksheetReturnsFalseWhenMissing"

    Dim targetBook As Workbook
    Dim removed As Boolean

    On Error GoTo Fail

    Set targetBook = ThisWorkbook
    DeleteWorksheet TEMP_SHEET_NAME

    removed = Manager.RemoveWorksheet(targetBook, TEMP_SHEET_NAME)
    Assert.IsFalse removed, "RemoveWorksheet should return False when sheet is missing"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveWorksheetReturnsFalseWhenMissing", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

Private Sub PrepareFixtureSheet()
    Dim fixtureSheet As Worksheet

    DeleteWorksheet TEMP_SHEET_NAME
    Set fixtureSheet = EnsureWorksheet(TEMP_SHEET_NAME)
    fixtureSheet.Range("A1").Value = "Fixture"
End Sub
