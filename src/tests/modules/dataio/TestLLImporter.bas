Attribute VB_Name = "TestLLImporter"
Attribute VB_Description = "Unit tests for LLImporter"

'@Folder("Tests.DataIO")
'@ModuleDescription("Unit tests for LLImporter")
'@TestModule

Option Explicit
Option Private Module

Private Assert As ICustomTest


'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CustomTest.Create()
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    '
End Sub

'@TestCleanup
Public Sub TestCleanup()
    '
End Sub


'@section Factory
'===============================================================================

'@TestMethod("LLImporter")
Public Sub FactoryCreatesWithWorkbook()
    CustomTestSetTitles Assert, "LLImporter", "FactoryCreatesWithWorkbook"
    On Error GoTo TestFail

    Dim imp As ILLImporter
    Set imp = LLImporter.Create(ThisWorkbook)
    Assert.IsNotNothing imp, "Factory should return a valid object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryCreatesWithWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("LLImporter")
Public Sub FactoryRejectsNothingWorkbook()
    CustomTestSetTitles Assert, "LLImporter", "FactoryRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim imp As ILLImporter
    On Error Resume Next
    Set imp = LLImporter.Create(Nothing)
    Assert.IsTrue Err.Number <> 0, "Factory should raise error for Nothing workbook"
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingWorkbook", Err.Number, Err.Description
End Sub


'@section Report
'===============================================================================

'@TestMethod("LLImporter")
Public Sub NeedReportFalseByDefault()
    CustomTestSetTitles Assert, "LLImporter", "NeedReportFalseByDefault"
    On Error GoTo TestFail

    Dim imp As ILLImporter
    Set imp = LLImporter.Create(ThisWorkbook)
    Assert.IsFalse imp.NeedReport, "NeedReport should be False before any import"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "NeedReportFalseByDefault", Err.Number, Err.Description
End Sub

'@TestMethod("LLImporter")
Public Sub ClearReportResetsState()
    CustomTestSetTitles Assert, "LLImporter", "ClearReportResetsState"
    On Error GoTo TestFail

    Dim imp As ILLImporter
    Set imp = LLImporter.Create(ThisWorkbook)
    imp.ClearReport
    Assert.IsFalse imp.NeedReport, "NeedReport should be False after ClearReport"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "ClearReportResetsState", Err.Number, Err.Description
End Sub

'@TestMethod("LLImporter")
Public Sub ReportSheetsEmptyByDefault()
    CustomTestSetTitles Assert, "LLImporter", "ReportSheetsEmptyByDefault"
    On Error GoTo TestFail

    Dim imp As ILLImporter
    Dim sheets As BetterArray

    Set imp = LLImporter.Create(ThisWorkbook)
    imp.ClearReport
    Set sheets = imp.ReportSheets(ImportReportNotImported)
    Assert.AreEqual CLng(0), sheets.Length, _
                    "ReportSheets should be empty before import"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "ReportSheetsEmptyByDefault", Err.Number, Err.Description
End Sub

'@TestMethod("LLImporter")
Public Sub ReportVariablesEmptyByDefault()
    CustomTestSetTitles Assert, "LLImporter", "ReportVariablesEmptyByDefault"
    On Error GoTo TestFail

    Dim imp As ILLImporter
    Dim vars As BetterArray

    Set imp = LLImporter.Create(ThisWorkbook)
    imp.ClearReport
    Set vars = imp.ReportVariables(ImportReportNotTouched)
    Assert.AreEqual CLng(0), vars.Length, _
                    "ReportVariables should be empty before import"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "ReportVariablesEmptyByDefault", Err.Number, Err.Description
End Sub

'@TestMethod("LLImporter")
Public Sub ReportSheetsInvalidScopeReturnsEmpty()
    CustomTestSetTitles Assert, "LLImporter", "ReportSheetsInvalidScopeReturnsEmpty"
    On Error GoTo TestFail

    Dim imp As ILLImporter
    Dim sheets As BetterArray

    Set imp = LLImporter.Create(ThisWorkbook)
    imp.ClearReport
    Set sheets = imp.ReportSheets(99)
    Assert.AreEqual CLng(0), sheets.Length, _
                    "Invalid scope should return empty BetterArray"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "ReportSheetsInvalidScopeReturnsEmpty", Err.Number, Err.Description
End Sub
