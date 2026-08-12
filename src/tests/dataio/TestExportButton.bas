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
'parses the numeric suffix from the button name (e.g. "CMDExport3" yields 3)
'and answers 0 for a name it cannot place, and the UseFilter property, which
'reads the companion checkbox on every ask.
'The fixture builds its controls on a UserForm and takes them off again in
'TestCleanup, so the order the tests run in decides nothing.
'
'WHAT STAYS OUT OF REACH
'-------------------------------------------------------------------------------
'RunExport opens a folder picker and ends in a message box, so the click path
'itself runs under no headless test. What it does with the answers is measured
'here through ExportNumber and UseFilter, and the sync it drives is measured in
'TestFilteredData.
'@depends ExportButton, TranslationObject, TestHelpersLite, MSForms, CustomTest

Option Explicit
Option Private Module

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As CustomTest

'The names of the controls this module put on the host form. Only these come
'off again, so the controls the form was designed with stay where they are.
Private addedControls As Collection


'@section Module Lifecycle
'===============================================================================

'@sub-title Initialise the test module before any tests run.
'@details
'Ensures the output worksheet exists and creates the CustomTest assertion
'object used by all test methods in this module. Called once before the
'first test executes.
'@ModuleInitialize
Public Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
End Sub

'@sub-title Tear down the module after all tests complete.
'@details
'Renders the collected results to the output worksheet, which is what the
'headless runner harvests into the results file, then restores the
'application state and releases the assertion object. Called once after the
'last test finishes.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Clear the host form before each individual test.
'@TestInitialize
Public Sub TestInitialize()
    ClearHostForm
End Sub

'@sub-title Flush the assert state and take the fixture controls off.
'@details
'Flush persists the assertions of the test that just ran into the results
'buffer. PrintResults renders that buffer alone, so a test that skips the
'flush leaves no trace in the results file.
'
'The controls go with it. A form instance keeps everything added to it until
'the instance is unloaded, and Controls.Add refuses a name already in use, so
'a test that left its button behind would break the next one.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then Assert.Flush
    ClearHostForm
End Sub


'@section Helpers
'===============================================================================

'@sub-title The form the fixture controls are built on.
'@details
'Excel for Mac carries no ActiveX worksheet controls, so OLEObjects.Add
'answered "Unable to get the Add property" for every Forms.CommandButton.1
'the old fixture asked for, and all seven tests errored on the arrange. A
'UserForm builds the same controls on every host, and it is the surface the
'class meets in production: SetupExportForm adds each export button to
'F_Export the same way.
'
'DraftForm is the spare form of the driver workbook. It carries no code of
'its own and the import sweep leaves it in place, so it is a host this suite
'gets for free.
'@return Object. The host form.
Private Function HostForm() As Object
    Set HostForm = DraftForm
End Function

'@sub-title Take every control this module added off the host form.
'@details
'Only the names this module recorded come off, so a control the form was
'designed with stays on it. Remove refuses a designed control anyway, and the
'handler carries that.
Private Sub ClearHostForm()
    Dim controlName As Variant

    If Not addedControls Is Nothing Then
        For Each controlName In addedControls
            On Error Resume Next
            HostForm().Controls.Remove CStr(controlName)
            On Error GoTo 0
        Next controlName
    End If

    Set addedControls = New Collection
End Sub

'@sub-title Create a CommandButton on the host form.
'@details
'Adds a Forms.CommandButton.1 to the host form under the requested name and
'records that name for cleanup. The button name drives ExportNumber parsing
'(e.g. "CMDExport3").
'@param buttonName String. The name to give the button control.
'@return MSForms.CommandButton. The newly created button.
Private Function CreateButton(ByVal buttonName As String) As MSForms.CommandButton
    Dim btn As MSForms.CommandButton

    Set btn = HostForm().Controls.Add("Forms.CommandButton.1", buttonName, True)
    addedControls.Add buttonName
    Set CreateButton = btn
End Function

'@sub-title Create a CheckBox on the host form.
'@details
'Adds a Forms.CheckBox.1 to the host form and records its name for cleanup.
'Used to test the UseFilter property, which reads the companion checkbox on
'every ask.
'@return MSForms.CheckBox. The newly created checkbox.
Private Function CreateCheckBox() As MSForms.CheckBox
    Dim chk As MSForms.CheckBox
    Dim controlName As String

    controlName = "CHKFilterFixture"
    Set chk = HostForm().Controls.Add("Forms.CheckBox.1", controlName, True)
    addedControls.Add controlName
    Set CreateCheckBox = chk
End Function

'@sub-title Create a stub TranslationObject for factory calls.
'@details
'Instantiates a TranslationObject and initialises it with
'an arbitrary name. The stub satisfies the Create factory's non-Nothing
'translation requirement without needing a full linelist dictionary.
'@return TranslationObject. A lightweight translation stub.
Private Function CreateTranslationStub() As TranslationObject
    Set CreateTranslationStub = BuildTranslationObject(ThisWorkbook, "ENG", Array())
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

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton("CMDExport1")

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

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton("CMDExport3")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)
    Assert.AreEqual 3&, sut.ExportNumber, _
                    "ExportNumber should parse '3' from CMDExport3"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "ExportNumberParsesButtonName", Err.Number, Err.Description
End Sub

'@sub-title Verify ExportNumber answers 0 for a name without the tag.
'@details
'Arranges a button named "Export", which carries no "CMDExport" prefix. Acts
'by reading ExportNumber. Asserts the answer is 0, confirming the read places
'no number on a name it does not recognise. The read runs on the click path,
'so answering 0 is what keeps a stray name off the raise route.
'@TestMethod("ExportButton")
Public Sub ExportNumberZeroWithoutTheTag()
    CustomTestSetTitles Assert, "ExportButton", "ExportNumberZeroWithoutTheTag"
    On Error GoTo TestFail

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton("Export")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)
    Assert.AreEqual 0&, sut.ExportNumber, _
                    "ExportNumber should answer 0 for a name without the tag"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "ExportNumberZeroWithoutTheTag", Err.Number, Err.Description
End Sub

'@sub-title Verify ExportNumber answers 0 when the tag is followed by letters.
'@details
'Arranges a button named "CMDExportABC". Acts by reading ExportNumber.
'Asserts the answer is 0. The old read handed the remainder to CLng, which
'raises 13 on a word, and the raise came out of a property read inside the
'click handler.
'@TestMethod("ExportButton")
Public Sub ExportNumberZeroForLettersAfterTheTag()
    CustomTestSetTitles Assert, "ExportButton", "ExportNumberZeroForLettersAfterTheTag"
    On Error GoTo TestFail

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton("CMDExportABC")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)
    Assert.AreEqual 0&, sut.ExportNumber, _
                    "ExportNumber should answer 0 when no digits follow the tag"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "ExportNumberZeroForLettersAfterTheTag", Err.Number, Err.Description
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

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton("CMDExport1")

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

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton("CMDExport1")

    Dim chk As MSForms.CheckBox
    Set chk = CreateCheckBox()
    chk.Value = True

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn, chk)
    Assert.IsTrue sut.UseFilter, _
                  "UseFilter should reflect checkbox True value"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "UseFilterReadsCheckboxValue", Err.Number, Err.Description
End Sub

'@sub-title Verify UseFilter answers the checkbox on every read.
'@details
'Arranges a button with a companion checkbox. Acts by reading UseFilter,
'changing the control, and reading it again. Asserts that both reads match
'the control, which is what lets one box serve every export button on the
'form: the class keeps no copy of the answer.
'@TestMethod("ExportButton")
Public Sub UseFilterFollowsTheCheckbox()
    CustomTestSetTitles Assert, "ExportButton", "UseFilterFollowsTheCheckbox"
    On Error GoTo TestFail

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton("CMDExport1")

    Dim chk As MSForms.CheckBox
    Set chk = CreateCheckBox()
    chk.Value = True

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn, chk)

    Assert.IsTrue sut.UseFilter, _
                  "UseFilter should be True while the checkbox is ticked"

    chk.Value = False
    Assert.IsFalse sut.UseFilter, _
                   "UseFilter should follow the checkbox once it is unticked"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "UseFilterFollowsTheCheckbox", Err.Number, Err.Description
End Sub
