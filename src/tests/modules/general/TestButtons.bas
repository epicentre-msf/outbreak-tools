Attribute VB_Name = "TestButtons"

Option Explicit
Option Private Module

'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")

Private Const BUTTONS_SHEET As String = "ButtonsFixture"
Private Const DEFAULT_BUTTON_NAME As String = "FixtureButton"
Private Const BUTTON_FORMAT_SHEET As String = "LLFormatFixture_Buttons"
Private Const FIXTURE_DEFAULT_DESIGN As String = "design 1"
Private Const LABEL_BUTTON_INTERIOR As String = "button default interior color"
Private Const LABEL_BUTTON_FONT As String = "button default font color"

Private Assert As Object
Private Fakes As Object
Private FixtureSheet As Worksheet

'@section Helpers
'===============================================================================

Private Sub ResetButtonsSheet()
    Set FixtureSheet = EnsureWorksheet(BUTTONS_SHEET)
End Sub

Private Function AnchorCell() As Range
    Set AnchorCell = FixtureSheet.Range("B2")
End Function

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
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    BusyApp
    ResetButtonsSheet
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
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
    On Error Resume Next
        LLFormatTestFixture.DeleteLLFormatFixture BUTTON_FORMAT_SHEET
    On Error GoTo 0
    Set FixtureSheet = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Buttons")
Private Sub TestCreateInitialisesState()
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
    FailUnexpectedError Assert, "TestCreateInitialisesState"
End Sub

'@TestMethod("Buttons")
Private Sub TestAddCreatesShape()
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
    FailUnexpectedError Assert, "TestAddCreatesShape"
End Sub

'@TestMethod("Buttons")
Private Sub TestAddExistingRecordsCheckings()
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
    FailUnexpectedError Assert, "TestAddExistingRecordsCheckings"
End Sub

'@TestMethod("Buttons")
Private Sub TestFormatAppliesScopeUsingWorkbookDesign()
    Dim buttonHelper As IButtons
    Dim design As ILLFormat
    Dim createdShape As Shape
    Dim formatSheet As Worksheet
    Dim templateSheet As Worksheet
    Dim expectedFillColor As Long
    Dim expectedFontColor As Long

    On Error GoTo Fail

    Set templateSheet = LLFormatTestFixture.LLFormatTemplate
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
    FailUnexpectedError Assert, "TestFormatAppliesScopeUsingWorkbookDesign"
    Resume Cleanup

Cleanup:
    LLFormatTestFixture.DeleteLLFormatFixture BUTTON_FORMAT_SHEET
    Exit Sub
End Sub
