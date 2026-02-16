Attribute VB_Name = "TestButtons"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the Buttons class")

'@description
'Validates the Buttons class, which creates and formats Excel worksheet
'shapes used as clickable action buttons in linelist sheets. Tests cover
'factory initialisation (anchor range, code name, scope), shape creation
'via Add, duplicate-button checking logging, and visual formatting
'against an LLFormat design template. The fixture creates a temporary
'worksheet for each test and cleans up all shapes and format sheets on
'teardown to guarantee isolation.
'@depends Buttons, IButtons, LLFormat, ILLFormat, Checking, IChecking, CustomTest, TestHelpers, LLFormatTestFixture

Private Const BUTTONS_SHEET As String = "ButtonsFixture"
Private Const DEFAULT_BUTTON_NAME As String = "FixtureButton"
Private Const FORMAT_TEMPLATE_SHEET As String = "LLFormatTemplate"
Private Const BUTTON_FORMAT_SHEET As String = "LLFormatFixture_Buttons"
Private Const FIXTURE_DEFAULT_DESIGN As String = "design 1"
Private Const LABEL_BUTTON_INTERIOR As String = "button default interior color"
Private Const LABEL_BUTTON_FONT As String = "button default font color"

Private Assert As ICustomTest
Private Fakes As Object
Private FixtureSheet As Worksheet

'@section Helpers
'===============================================================================

'@sub-title Create or clear the fixture worksheet for button tests.
Private Sub ResetButtonsSheet()
    Set FixtureSheet = EnsureWorksheet(BUTTONS_SHEET)
End Sub

'@sub-title Return the default anchor cell for button placement.
Private Function AnchorCell() As Range
    Set AnchorCell = FixtureSheet.Range("B2")
End Function

'@sub-title Build a Buttons instance with sensible defaults.
'@details
'Creates a Buttons object using the fixture sheet anchor. The scope and
'name can be overridden for specific tests; when no anchor is supplied the
'default AnchorCell is used.
Private Function BuildButton(Optional ByVal scope As ButtonScope = ButtonScopeSmall, _
                             Optional ByVal buttonName As String = DEFAULT_BUTTON_NAME, _
                             Optional ByVal anchor As Range) As IButtons
    If anchor Is Nothing Then
        Set anchor = AnchorCell()
    End If
    Set BuildButton = Buttons.Create(anchor, buttonName, scope)
End Function

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestButtons"
    ResetButtonsSheet
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet BUTTONS_SHEET
    LLFormatTestFixture.DeleteLLFormatFixture BUTTON_FORMAT_SHEET
    RestoreApp
    Set FixtureSheet = Nothing
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    ResetButtonsSheet
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    On Error Resume Next
        LLFormatTestFixture.DeleteLLFormatFixture BUTTON_FORMAT_SHEET
    On Error GoTo 0
    Set FixtureSheet = Nothing
End Sub

'@section Tests
'===============================================================================

'@sub-title Verify Create stores anchor, name, and scope.
'@details
'Creates a Buttons instance with an explicit anchor, code name, and
'ButtonScopeLarge. Asserts that OutputRange, Name, and ShapeScope all
'reflect the input values. Also verifies that passing Nothing as the
'anchor raises ObjectNotInitialized.
'@TestMethod("Buttons")
Public Sub TestCreateInitialisesState()
    CustomTestSetTitles Assert, "Buttons", "TestCreateInitialisesState"
    Dim buttonHelper As IButtons
    Dim anchor As Range

    On Error GoTo Fail

    Set anchor = AnchorCell()
    Set buttonHelper = Buttons.Create(anchor, "CreateButton", ButtonScopeLarge)

    Assert.AreEqual "Buttons", TypeName(buttonHelper), "Create should return a Buttons implementation"
    Assert.AreEqual anchor.Address(False, False), buttonHelper.OutputRange.Address(False, False), _
                   "OutputRange should match the provided anchor"
    Assert.AreEqual "CreateButton", buttonHelper.Name, "Name should match the provided code name"
    Assert.AreEqual ButtonScopeLarge, buttonHelper.ShapeScope, "ShapeScope should persist the provided scope"

    On Error Resume Next
        Set anchor = Nothing
        Err.Clear
        '@Ignore AssignmentNotUsed
        Set buttonHelper = Buttons.Create(anchor, "Invalid", ButtonScopeSmall)
        Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                        "Create should raise when the anchor range is missing"
    On Error GoTo Fail

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateInitialisesState", Err.Number, Err.Description
End Sub

'@sub-title Verify Add places exactly one shape on the worksheet.
'@details
'Creates a button and calls Add with a macro name and label. Asserts that
'the fixture sheet contains exactly one shape, its OnAction matches the
'provided command, and the text label is applied. Also checks that no
'checking entries are logged for the first creation.
'@TestMethod("Buttons")
Public Sub TestAddCreatesShape()
    CustomTestSetTitles Assert, "Buttons", "TestAddCreatesShape"
    Dim buttonHelper As IButtons
    Dim createdShape As Shape

    On Error GoTo Fail

    Set buttonHelper = BuildButton(ButtonScopeSmall, "AddButton")
    buttonHelper.Add actionCommand:="TestMacro", shapeLabel:="Press me"

    Assert.AreEqual 1, FixtureSheet.Shapes.Count, "Add should create exactly one shape"

    Set createdShape = FixtureSheet.Shapes("AddButton")
    Assert.AreEqual "TestMacro", createdShape.OnAction, "Add should assign the action command"
    Assert.AreEqual "Press me", createdShape.TextFrame2.TextRange.Characters.Text, _
                   "Add should populate the provided label"

    Assert.IsFalse buttonHelper.HasCheckings, "Creating the button should not log checkings"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddCreatesShape", Err.Number, Err.Description
End Sub

'@sub-title Verify duplicate Add logs a checking instead of creating a second shape.
'@details
'Calls Add twice with the same button name. Asserts that only one shape
'exists on the worksheet, HasCheckings is True, and CheckingValues contains
'exactly one entry whose label describes the duplicate button.
'@TestMethod("Buttons")
Public Sub TestAddExistingRecordsCheckings()
    CustomTestSetTitles Assert, "Buttons", "TestAddExistingRecordsCheckings"
    Dim buttonHelper As IButtons
    Dim logs As IChecking

    On Error GoTo Fail

    Set buttonHelper = BuildButton(ButtonScopeSmall, "ExistingButton")
    buttonHelper.Add actionCommand:="ExistingMacro", shapeLabel:="Run"
    buttonHelper.Add actionCommand:="ExistingMacro", shapeLabel:="Run"

    Assert.AreEqual 1, FixtureSheet.Shapes.Count, "Repeated Add should not duplicate the shape"
    Assert.IsTrue buttonHelper.HasCheckings, "Repeated Add should log a checking entry"

    Set logs = buttonHelper.CheckingValues
    Assert.IsTrue (Not logs Is Nothing), "CheckingValues should yield an IChecking instance"
    Assert.AreEqual 1, logs.Length, "Only one checking entry should be recorded"
    Assert.AreEqual "Button ExistingButton already exists; skipping creation", _
                   logs.ValueOf("0", checkingLabel), _
                   "Checking message should explain the duplicate shape"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddExistingRecordsCheckings", Err.Number, Err.Description
End Sub

'@sub-title Verify Format applies design colours from an LLFormat template.
'@details
'Prepares an LLFormat fixture with known button interior and font colours,
'creates a button, then calls Format with the design. Asserts that the
'shape fill and text font colour match the expected values read from the
'template sheet.
'@TestMethod("Buttons")
Public Sub TestFormatAppliesScopeUsingWorkbookDesign()
    CustomTestSetTitles Assert, "Buttons", "TestFormatAppliesScopeUsingWorkbookDesign"
    Dim buttonHelper As IButtons
    Dim design As ILLFormat
    Dim createdShape As Shape
    Dim formatSheet As Worksheet
    Dim templateSheet As Worksheet
    Dim expectedFillColor As Long
    Dim expectedFontColor As Long

    On Error GoTo Fail

    Set templateSheet = LLFormatTestFixture.PrepareLLFormatFixture(FORMAT_TEMPLATE_SHEET)
    expectedFillColor = CLng(LLFormatTestFixture.FixtureCell(templateSheet, _
                                     LABEL_BUTTON_INTERIOR, FIXTURE_DEFAULT_DESIGN).Interior.Color)
    expectedFontColor = CLng(LLFormatTestFixture.FixtureCell(templateSheet, _
                                     LABEL_BUTTON_FONT, FIXTURE_DEFAULT_DESIGN).Interior.Color)

    Set formatSheet = LLFormatTestFixture.PrepareLLFormatFixture(BUTTON_FORMAT_SHEET)
    formatSheet.Range("DESIGNTYPE").Value = FIXTURE_DEFAULT_DESIGN

    Set buttonHelper = BuildButton(ButtonScopeSmall, "FormatButton")
    buttonHelper.Add shapeLabel:="Format"

    Set design = LLFormat.Create(formatSheet)
    buttonHelper.Format design

    Set createdShape = FixtureSheet.Shapes("FormatButton")
    Assert.AreEqual expectedFillColor, createdShape.Fill.ForeColor.RGB, _
                     "Button should adopt linelist button colour"
    Assert.AreEqual expectedFontColor, createdShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB, _
                     "Button text should adopt design font colour"

    GoTo Cleanup

Fail:
    CustomTestLogFailure Assert, "TestFormatAppliesScopeUsingWorkbookDesign", Err.Number, Err.Description
    Resume Cleanup

Cleanup:
    DeleteWorksheet FORMAT_TEMPLATE_SHEET
    DeleteWorksheet BUTTON_FORMAT_SHEET
    Exit Sub
End Sub
