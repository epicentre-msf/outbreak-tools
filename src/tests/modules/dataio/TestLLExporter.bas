Attribute VB_Name = "TestLLExporter"
Attribute VB_Description = "Unit tests for LLExporter"

'@Folder("Tests.DataIO")
'@ModuleDescription("Unit tests for LLExporter")
'@TestModule

'@description
'Validates the LLExporter class, which manages linelist data export operations
'including custom filtered exports, migration exports, analysis exports, and
'geobase exports. Tests cover factory initialisation (Create with valid and
'Nothing workbook arguments), default state of the LastExportPassword property,
'and safe invocation of CloseAll when no output workbooks are open. The fixture
'uses ThisWorkbook as the source workbook for lightweight instantiation without
'requiring full linelist infrastructure.
'@depends LLExporter, ILLExporter, CustomTest

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

'@sub-title Verify Create returns a valid ILLExporter instance for a real workbook.
'@details
'Acts by calling LLExporter.Create with ThisWorkbook. Asserts that the
'returned object is not Nothing, confirming the factory accepts a valid
'workbook and produces a usable exporter instance.
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

'@sub-title Verify Create raises an error when the workbook argument is Nothing.
'@details
'Acts by calling LLExporter.Create with Nothing under On Error Resume Next.
'Asserts that a non-zero error number was raised, confirming the guard
'clause rejects invalid input and prevents creation of an uninitialised
'exporter.
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

'@sub-title Verify LastExportPassword defaults to an empty string on a fresh instance.
'@details
'Arranges by creating an exporter from ThisWorkbook without performing any
'export operation. Acts by reading the LastExportPassword property. Asserts
'that it equals vbNullString, confirming the initial state is clean before
'any export has been executed.
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

'@sub-title Verify CloseAll completes without error when no output workbooks exist.
'@details
'Arranges by creating a fresh exporter that has not opened any output
'workbooks. Acts by calling CloseAll. Asserts that the call succeeds
'without raising an error, confirming CloseAll is safe to invoke even
'when there is nothing to close.
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
