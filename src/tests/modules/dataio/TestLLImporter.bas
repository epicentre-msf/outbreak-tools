Attribute VB_Name = "TestLLImporter"
Attribute VB_Description = "Unit tests for LLImporter"

'@Folder("Tests.DataIO")
'@ModuleDescription("Unit tests for LLImporter")
'@TestModule

'@description
'Validates the LLImporter class, which manages linelist data import operations
'including data import from HList/VList worksheets, custom dropdown import,
'migration metadata import, and geobase import. Tests cover factory
'initialisation (Create with valid and Nothing workbook arguments), the import
'report subsystem (NeedReport, ClearReport, ReportSheets, ReportVariables),
'and default state verification. The fixture uses ThisWorkbook as the host
'workbook for lightweight instantiation without requiring full linelist
'infrastructure.
'@depends LLImporter, ILLImporter, BetterArray, CustomTest

Option Explicit
Option Private Module

Private Assert As ICustomTest


'@section Module Lifecycle
'===============================================================================

'@sub-title Initialise the test module before any tests run.
'@details
'Creates the CustomTest assertion object used by all test methods in this
'module. Called once before the first test executes.
'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CustomTest.Create()
End Sub

'@sub-title Tear down the module after all tests complete.
'@details
'Releases the assertion object to avoid memory leaks. Called once after
'the last test finishes.
'@ModuleCleanup
Public Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@sub-title Reset state before each individual test.
'@TestInitialize
Public Sub TestInitialize()
    '
End Sub

'@sub-title Clean up after each individual test.
'@TestCleanup
Public Sub TestCleanup()
    '
End Sub


'@section Factory
'===============================================================================

'@sub-title Verify Create returns a valid ILLImporter instance for a real workbook.
'@details
'Acts by calling LLImporter.Create with ThisWorkbook. Asserts that the
'returned object is not Nothing, confirming the factory accepts a valid
'workbook and produces a usable importer instance.
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

'@sub-title Verify Create raises an error when the workbook argument is Nothing.
'@details
'Acts by calling LLImporter.Create with Nothing under On Error Resume Next.
'Asserts that a non-zero error number was raised, confirming the guard
'clause rejects invalid input and prevents creation of an uninitialised
'importer.
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

'@sub-title Verify NeedReport defaults to False on a fresh importer instance.
'@details
'Arranges by creating an importer from ThisWorkbook without performing any
'import operation. Acts by reading the NeedReport property. Asserts that it
'is False, confirming the initial report state is clean before any import
'has been executed.
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

'@sub-title Verify ClearReport resets the report state to clean.
'@details
'Arranges by creating an importer from ThisWorkbook. Acts by calling
'ClearReport and then reading NeedReport. Asserts that NeedReport is False
'after the clear, confirming ClearReport resets any accumulated import
'diagnostics back to the initial state.
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

'@sub-title Verify ReportSheets returns an empty BetterArray before any import.
'@details
'Arranges by creating an importer and calling ClearReport to ensure a
'pristine state. Acts by calling ReportSheets with ImportReportNotImported
'scope. Asserts that the returned BetterArray has zero length, confirming
'no sheet-level diagnostics exist before any import operation.
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

'@sub-title Verify ReportVariables returns an empty BetterArray before any import.
'@details
'Arranges by creating an importer and calling ClearReport to ensure a
'pristine state. Acts by calling ReportVariables with ImportReportNotTouched
'scope. Asserts that the returned BetterArray has zero length, confirming
'no variable-level diagnostics exist before any import operation.
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

'@sub-title Verify ReportSheets returns an empty BetterArray for an invalid scope value.
'@details
'Arranges by creating an importer and calling ClearReport. Acts by calling
'ReportSheets with scope value 99, which does not correspond to any valid
'ImportReportScope enum member. Asserts that the returned BetterArray has
'zero length, confirming the method handles out-of-range scope values
'gracefully by returning an empty collection rather than raising an error.
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
