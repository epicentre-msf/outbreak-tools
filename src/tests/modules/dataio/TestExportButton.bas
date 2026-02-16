Attribute VB_Name = "TestExportButton"
Attribute VB_Description = "Unit tests for ExportButton"

'@Folder("Tests.DataIO")
'@ModuleDescription("Unit tests for ExportButton")
'@TestModule
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument

'@description
'Validates the ExportButton class, which wraps an MSForms.CommandButton and an
'optional MSForms.CheckBox to drive filtered custom exports from the linelist.
'Tests cover factory initialisation (Create with valid arguments, rejection of
'Nothing workbook, translations, and button), the ExportNumber property that
'parses the numeric suffix from the button name (e.g. "CMDExport3" yields 3),
'the UseFilter property that reads and writes the companion checkbox state,
'and interface delegation via IExportButton. The fixture creates temporary
'worksheets with OLEObject controls for each test and tears them down in
'TestCleanup to ensure isolation.
'@depends ExportButton, IExportButton, ITranslationObject, LinelistSpecsTranslationStub, MSForms, CustomTest

Option Explicit
Option Private Module

Private Assert As ICustomTest
Private testSheet As Worksheet


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
'@details
'Deletes the temporary worksheet created during the test to remove any
'OLEObject controls and ensure the next test starts with a clean workbook.
'@TestCleanup
Public Sub TestCleanup()
    CleanupTestSheet
End Sub


'@section Helpers
'===============================================================================

'@sub-title Create a temporary worksheet for hosting OLEObject controls.
'@details
'Adds a new worksheet to ThisWorkbook, stores it in the module-level
'testSheet variable for cleanup, and returns a reference to the caller.
Private Function CreateTestSheet() As Worksheet
    Set testSheet = ThisWorkbook.Worksheets.Add
    Set CreateTestSheet = testSheet
End Function

'@sub-title Delete the temporary worksheet if it exists.
'@details
'Checks the module-level testSheet variable and, when set, deletes the
'worksheet with alerts suppressed to avoid user confirmation dialogs.
'Resets the reference to Nothing so repeated calls are safe.
Private Sub CleanupTestSheet()
    If Not testSheet Is Nothing Then
        Application.DisplayAlerts = False
        testSheet.Delete
        Application.DisplayAlerts = True
        Set testSheet = Nothing
    End If
End Sub

'@sub-title Create a CommandButton OLEObject on the given worksheet.
'@details
'Adds a Forms.CommandButton.1 OLE control to the worksheet, assigns
'the requested name to the underlying MSForms.CommandButton, and returns
'it. The button name drives ExportNumber parsing (e.g. "CMDExport3").
'@param sh Worksheet. The host worksheet for the OLEObject.
'@param buttonName String. The name to assign to the button control.
'@return MSForms.CommandButton. The newly created button.
Private Function CreateButton(ByVal sh As Worksheet, _
                               ByVal buttonName As String) As MSForms.CommandButton
    Dim ole As OLEObject
    Set ole = sh.OLEObjects.Add(ClassType:="Forms.CommandButton.1")
    Dim btn As MSForms.CommandButton
    Set btn = ole.Object
    btn.Name = buttonName
    Set CreateButton = btn
End Function

'@sub-title Create a CheckBox OLEObject on the given worksheet.
'@details
'Adds a Forms.CheckBox.1 OLE control to the worksheet and returns the
'underlying MSForms.CheckBox. Used to test the UseFilter property which
'reads and writes the companion checkbox state.
'@param sh Worksheet. The host worksheet for the OLEObject.
'@return MSForms.CheckBox. The newly created checkbox.
Private Function CreateCheckBox(ByVal sh As Worksheet) As MSForms.CheckBox
    Dim ole As OLEObject
    Set ole = sh.OLEObjects.Add(ClassType:="Forms.CheckBox.1")
    Set CreateCheckBox = ole.Object
End Function

'@sub-title Create a stub ITranslationObject for factory calls.
'@details
'Instantiates a LinelistSpecsTranslationStub and initialises it with
'an arbitrary name. The stub satisfies the Create factory's non-Nothing
'translation requirement without needing a full linelist dictionary.
'@return ITranslationObject. A lightweight translation stub.
Private Function CreateTranslationStub() As ITranslationObject
    Dim stub As New LinelistSpecsTranslationStub
    stub.Initialise "ExportTestStub"
    Set CreateTranslationStub = stub
End Function


'@section Factory Validation
'===============================================================================

'@sub-title Verify Create returns a valid ExportButton for valid arguments.
'@details
'Arranges a temporary worksheet with a CommandButton named "CMDExport1".
'Acts by calling ExportButton.Create with ThisWorkbook, a translation stub,
'and the button. Asserts that the returned object is not Nothing, confirming
'the factory accepts valid arguments and produces a usable instance.
'@TestMethod("ExportButton")
Public Sub FactoryCreatesWithValidArgs()
    CustomTestSetTitles Assert, "ExportButton", "FactoryCreatesWithValidArgs"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport1")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)
    Assert.IsNotNothing sut, "Factory should return a valid object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryCreatesWithValidArgs", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises an error when the workbook argument is Nothing.
'@details
'Acts by calling ExportButton.Create with Nothing as the workbook under
'On Error Resume Next. Asserts that a non-zero error number was raised,
'confirming the guard clause rejects a missing workbook.
'@TestMethod("ExportButton")
Public Sub FactoryRejectsNothingWorkbook()
    CustomTestSetTitles Assert, "ExportButton", "FactoryRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim sut As ExportButton
    On Error Resume Next
    Set sut = ExportButton.Create(Nothing, CreateTranslationStub(), Nothing)
    Assert.IsTrue Err.Number <> 0, "Should raise error for Nothing workbook"
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingWorkbook", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises an error when the translations argument is Nothing.
'@details
'Acts by calling ExportButton.Create with Nothing as the translations under
'On Error Resume Next. Asserts that a non-zero error number was raised,
'confirming the guard clause rejects a missing translation object.
'@TestMethod("ExportButton")
Public Sub FactoryRejectsNothingTranslations()
    CustomTestSetTitles Assert, "ExportButton", "FactoryRejectsNothingTranslations"
    On Error GoTo TestFail

    Dim sut As ExportButton
    On Error Resume Next
    Set sut = ExportButton.Create(ThisWorkbook, Nothing, Nothing)
    Assert.IsTrue Err.Number <> 0, "Should raise error for Nothing translations"
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingTranslations", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises an error when the button argument is Nothing.
'@details
'Acts by calling ExportButton.Create with a valid workbook and translation
'stub but Nothing as the button under On Error Resume Next. Asserts that a
'non-zero error number was raised, confirming the guard clause rejects a
'missing button control.
'@TestMethod("ExportButton")
Public Sub FactoryRejectsNothingButton()
    CustomTestSetTitles Assert, "ExportButton", "FactoryRejectsNothingButton"
    On Error GoTo TestFail

    Dim sut As ExportButton
    On Error Resume Next
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), Nothing)
    Assert.IsTrue Err.Number <> 0, "Should raise error for Nothing button"
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingButton", Err.Number, Err.Description
End Sub


'@section ExportNumber
'===============================================================================

'@sub-title Verify ExportNumber parses the numeric suffix from the button name.
'@details
'Arranges a button named "CMDExport3" on a temporary worksheet. Acts by
'creating an ExportButton and reading ExportNumber. Asserts that the value
'is 3, confirming the parsing logic strips the "CMDExport" prefix and
'converts the remaining characters to a Long.
'@TestMethod("ExportButton")
Public Sub ExportNumberParsesButtonName()
    CustomTestSetTitles Assert, "ExportButton", "ExportNumberParsesButtonName"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport3")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)
    Assert.AreEqual 3&, sut.ExportNumber, _
                    "ExportNumber should parse '3' from CMDExport3"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "ExportNumberParsesButtonName", Err.Number, Err.Description
End Sub


'@section UseFilter
'===============================================================================

'@sub-title Verify UseFilter returns False when no checkbox is bound.
'@details
'Arranges a button without a companion checkbox. Acts by creating an
'ExportButton and reading UseFilter. Asserts that the value is False,
'confirming the property defaults safely when the optional checkbox
'parameter was omitted during factory creation.
'@TestMethod("ExportButton")
Public Sub UseFilterFalseWithoutCheckbox()
    CustomTestSetTitles Assert, "ExportButton", "UseFilterFalseWithoutCheckbox"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport1")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)
    Assert.IsFalse sut.UseFilter, _
                   "UseFilter should be False when no checkbox bound"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "UseFilterFalseWithoutCheckbox", Err.Number, Err.Description
End Sub

'@sub-title Verify UseFilter reads the checkbox value when one is bound.
'@details
'Arranges a button with a companion checkbox whose Value is set to True.
'Acts by creating an ExportButton with the checkbox and reading UseFilter.
'Asserts that UseFilter is True, confirming the property delegates to the
'checkbox control's current state.
'@TestMethod("ExportButton")
Public Sub UseFilterReadsCheckboxValue()
    CustomTestSetTitles Assert, "ExportButton", "UseFilterReadsCheckboxValue"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport1")

    Dim chk As MSForms.CheckBox
    Set chk = CreateCheckBox(sh)
    chk.Value = True

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn, chk)
    Assert.IsTrue sut.UseFilter, _
                  "UseFilter should reflect checkbox True value"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "UseFilterReadsCheckboxValue", Err.Number, Err.Description
End Sub

'@sub-title Verify UseFilter Let updates the underlying checkbox control.
'@details
'Arranges a button with a companion checkbox initially set to True. Acts
'by setting UseFilter to False on the ExportButton. Asserts that the
'checkbox Value is now False, confirming the Property Let propagates the
'new state back to the underlying OLEObject control.
'@TestMethod("ExportButton")
Public Sub UseFilterLetUpdatesCheckbox()
    CustomTestSetTitles Assert, "ExportButton", "UseFilterLetUpdatesCheckbox"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport1")

    Dim chk As MSForms.CheckBox
    Set chk = CreateCheckBox(sh)
    chk.Value = True

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn, chk)

    sut.UseFilter = False
    Assert.IsFalse chk.Value, _
                   "Setting UseFilter to False should uncheck the checkbox"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "UseFilterLetUpdatesCheckbox", Err.Number, Err.Description
End Sub


'@section Interface
'===============================================================================

'@sub-title Verify IExportButton.ExportNumber delegates to the concrete property.
'@details
'Arranges a button named "CMDExport2" and creates an ExportButton instance.
'Acts by casting the instance to IExportButton and reading ExportNumber.
'Asserts that the interface property returns 2, confirming the delegation
'stub correctly forwards to the public ExportNumber implementation.
'@TestMethod("ExportButton")
Public Sub InterfaceExposesExportNumber()
    CustomTestSetTitles Assert, "ExportButton", "InterfaceExposesExportNumber"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport2")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)

    Dim iface As IExportButton
    Set iface = sut
    Assert.AreEqual 2&, iface.ExportNumber, _
                    "IExportButton.ExportNumber should delegate to ExportNumber"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "InterfaceExposesExportNumber", Err.Number, Err.Description
End Sub
