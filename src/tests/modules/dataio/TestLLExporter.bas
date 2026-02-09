Attribute VB_Name = "TestLLExporter"
Attribute VB_Description = "Unit tests for LLExporter"

'@Folder("Tests.DataIO")
'@ModuleDescription("Unit tests for LLExporter")
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

'@TestMethod("LLExporter")
Public Sub FactoryCreatesWithWorkbook()
    CustomTestSetTitles Assert, "LLExporter", "FactoryCreatesWithWorkbook"
    On Error GoTo TestFail

    Dim exporter As ILLExporter
    Set exporter = LLExporter.Create(ThisWorkbook)
    Assert.IsNotNothing exporter, "Factory should return a valid object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryCreatesWithWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("LLExporter")
Public Sub FactoryRejectsNothingWorkbook()
    CustomTestSetTitles Assert, "LLExporter", "FactoryRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim exporter As ILLExporter
    On Error Resume Next
    Set exporter = LLExporter.Create(Nothing)
    Assert.IsTrue Err.Number <> 0, "Factory should raise error for Nothing workbook"
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("LLExporter")
Public Sub LastExportPasswordEmptyByDefault()
    CustomTestSetTitles Assert, "LLExporter", "LastExportPasswordEmptyByDefault"
    On Error GoTo TestFail

    Dim exporter As ILLExporter
    Set exporter = LLExporter.Create(ThisWorkbook)
    Assert.AreEqual vbNullString, exporter.LastExportPassword, _
                    "Password should be empty before any export"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "LastExportPasswordEmptyByDefault", Err.Number, Err.Description
End Sub

'@TestMethod("LLExporter")
Public Sub CloseAllDoesNotError()
    CustomTestSetTitles Assert, "LLExporter", "CloseAllDoesNotError"
    On Error GoTo TestFail

    Dim exporter As ILLExporter
    Set exporter = LLExporter.Create(ThisWorkbook)
    exporter.CloseAll
    Assert.IsTrue True, "CloseAll should not raise errors when no workbook is open"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "CloseAllDoesNotError", Err.Number, Err.Description
End Sub
