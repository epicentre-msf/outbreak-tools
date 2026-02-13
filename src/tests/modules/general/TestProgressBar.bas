Attribute VB_Name = "TestProgressBar"
Attribute VB_Description = "Tests for ProgressBar class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for ProgressBar class")

' =============================================================================
' TestProgressBar
' =============================================================================
'
' @description
'   Comprehensive test suite for the ProgressBar class, which provides a
'   cell-range-based visual progress indicator with colour rendering.
'   Tests cover the full lifecycle of a progress bar: factory construction
'   with argument validation, value updates with boundary clamping, step-based
'   increments, completion and reset behaviour, optional status cell messaging,
'   dynamic maximum adjustment, value format configuration, and visual
'   rendering of completed/pending cells with formatted text output.
'
'   The suite is organised into six logical sections:
'     1. Factory and Attach   -- Create with explicit/default max, guard
'                                clauses for Nothing range and zero maximum,
'                                BarRange property after construction.
'     2. Update and StepBy    -- Setting values directly, clamping overflow
'                                and negative values, StepBy increments with
'                                explicit and default step sizes, Complete
'                                and Reset methods.
'     3. Status Cell          -- Attaching an optional single-cell range for
'                                status messages, verifying default Nothing
'                                state, and rejecting multi-cell ranges.
'     4. Maximum Property     -- Changing Maximum at runtime and verifying
'                                that existing Value is re-clamped accordingly.
'     5. ConfigureValueFormat -- Guard clause rejecting blank format patterns.
'     6. Rendering            -- Verifying that Update paints cells with the
'                                correct completed/pending colours and writes
'                                a formatted "current / max" string to the
'                                first cell of the bar range.
'
' @depends ProgressBar, IProgressBar, CustomTest, TestHelpers
' =============================================================================

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const PROGRESSBAR_SHEET As String = "ProgressBarFixture"

Private Assert As ICustomTest
Private FixtureSheet As Worksheet

'@section Module lifecycle
'===============================================================================

' @sub-title ModuleInitialize
' @details
'   Runs once before any test in this module. Suppresses screen updates via
'   BusyApp, ensures the shared test output sheet exists, and creates the
'   CustomTest harness bound to this module name.
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestProgressBar"
End Sub

' @sub-title ModuleCleanup
' @details
'   Runs once after all tests in this module have executed. Prints accumulated
'   test results to the output sheet, removes the fixture worksheet used for
'   progress bar ranges, restores normal application state, and releases
'   object references.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet PROGRESSBAR_SHEET
    RestoreApp
    Set FixtureSheet = Nothing
    Set Assert = Nothing
End Sub

' @sub-title TestInitialize
' @details
'   Runs before each individual test. Suppresses screen updates and creates
'   (or re-creates) a hidden fixture worksheet that provides cell ranges for
'   progress bar attachment without visual interference.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureSheet = EnsureWorksheet(PROGRESSBAR_SHEET, visibility:=xlSheetHidden)
End Sub

' @sub-title TestCleanup
' @details
'   Runs after each individual test. Flushes any pending assertions in the
'   harness and clears the fixture worksheet so the next test starts with
'   a clean range surface.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    If Not FixtureSheet Is Nothing Then
        ClearWorksheet FixtureSheet
    End If
End Sub

'@section Tests - Factory and Attach
'===============================================================================

' @sub-title TestCreateReturnsInitialisedBar
' @details
'   Verifies that ProgressBar.Create returns a properly initialised instance.
'   Arranges a five-cell range on the fixture sheet and calls Create with an
'   explicit maximum of 50. Asserts that the returned object is of type
'   ProgressBar, the Maximum property equals 50, the Value starts at zero,
'   and PercentComplete starts at zero. This is the primary happy-path test
'   for factory construction.
'@TestMethod("ProgressBar")
Public Sub TestCreateReturnsInitialisedBar()
    CustomTestSetTitles Assert, "ProgressBar", "TestCreateReturnsInitialisedBar"
    On Error GoTo TestFail

    Dim barRange As Range
    Set barRange = FixtureSheet.Range("A1:E1")

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(barRange, 50)

    Assert.AreEqual "ProgressBar", TypeName(sut), _
                     "Create should return a ProgressBar instance"
    Assert.AreEqual 50, sut.Maximum, _
                     "Create should set the maximum to the supplied value"
    Assert.AreEqual CLng(0), sut.Value, _
                     "Create should initialise the current value to zero"
    Assert.AreEqual 0#, sut.PercentComplete, _
                     "Create should initialise percent complete to zero"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsInitialisedBar", Err.Number, Err.Description
End Sub

' @sub-title TestCreateDefaultMaximum
' @details
'   Verifies that Create uses a default maximum of 100 when no explicit
'   maximum argument is supplied. Arranges a ten-cell range and calls Create
'   with only the range parameter. Asserts that Maximum equals 100. This
'   confirms the optional parameter default behaviour.
'@TestMethod("ProgressBar")
Public Sub TestCreateDefaultMaximum()
    CustomTestSetTitles Assert, "ProgressBar", "TestCreateDefaultMaximum"
    On Error GoTo TestFail

    Dim barRange As Range
    Set barRange = FixtureSheet.Range("A1:J1")

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(barRange)

    Assert.AreEqual CLng(100), sut.Maximum, _
                     "Create without explicit maximum should default to 100"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestCreateDefaultMaximum", Err.Number, Err.Description
End Sub

' @sub-title TestCreateRejectsNothingRange
' @details
'   Verifies the guard clause that rejects a Nothing range argument. Calls
'   Create with Nothing and expects an error to be raised. Uses the
'   ExpectError pattern: if execution reaches past Create, the test logs
'   a failure; otherwise the error handler asserts that the error number
'   matches ProjectError.InvalidArgument.
'@TestMethod("ProgressBar")
Public Sub TestCreateRejectsNothingRange()
    CustomTestSetTitles Assert, "ProgressBar", "TestCreateRejectsNothingRange"
    On Error GoTo ExpectError

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(Nothing)

    Assert.LogFailure "Create should raise when barRange is Nothing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Create should raise InvalidArgument for Nothing range"
End Sub

' @sub-title TestCreateRejectsZeroMaximum
' @details
'   Verifies the guard clause that rejects a zero maximum value, which would
'   cause division-by-zero in percent calculations. Arranges a valid range
'   and calls Create with maximum set to 0. Expects InvalidArgument to be
'   raised, confirming the factory validates its numeric parameter.
'@TestMethod("ProgressBar")
Public Sub TestCreateRejectsZeroMaximum()
    CustomTestSetTitles Assert, "ProgressBar", "TestCreateRejectsZeroMaximum"
    On Error GoTo ExpectError

    Dim barRange As Range
    Set barRange = FixtureSheet.Range("A1:E1")

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(barRange, 0)

    Assert.LogFailure "Create should raise when maximum is zero"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Create should raise InvalidArgument for zero maximum"
End Sub

' @sub-title TestAttachSetsBarRange
' @details
'   Verifies that the BarRange property correctly exposes the range that was
'   passed to Create. Arranges a five-cell range, creates the bar with
'   maximum 10, and asserts that the BarRange address matches the original
'   range address. This confirms the range reference is stored, not copied.
'@TestMethod("ProgressBar")
Public Sub TestAttachSetsBarRange()
    CustomTestSetTitles Assert, "ProgressBar", "TestAttachSetsBarRange"
    On Error GoTo TestFail

    Dim barRange As Range
    Set barRange = FixtureSheet.Range("A1:E1")

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(barRange, 10)

    Assert.AreEqual barRange.Address, sut.BarRange.Address, _
                     "BarRange should reference the attached range"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestAttachSetsBarRange", Err.Number, Err.Description
End Sub

'@section Tests - Update and StepBy
'===============================================================================

' @sub-title TestUpdateSetsValueAndPercent
' @details
'   Verifies the core Update method sets both Value and PercentComplete.
'   Creates a bar with maximum 50 on a ten-cell range, then calls Update
'   with 25. Asserts Value equals 25 and PercentComplete equals 0.5
'   (25/50). This is the primary happy-path test for the Update method.
'@TestMethod("ProgressBar")
Public Sub TestUpdateSetsValueAndPercent()
    CustomTestSetTitles Assert, "ProgressBar", "TestUpdateSetsValueAndPercent"
    On Error GoTo TestFail

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:J1"), 50)

    sut.Update 25

    Assert.AreEqual CLng(25), sut.Value, _
                     "Update should set the current value"
    Assert.AreEqual 0.5, sut.PercentComplete, _
                     "PercentComplete should reflect value/maximum ratio"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestUpdateSetsValueAndPercent", Err.Number, Err.Description
End Sub

' @sub-title TestUpdateClampsToMaximum
' @details
'   Verifies that Update clamps values that exceed the maximum. Creates a bar
'   with maximum 10 and updates with value 999. Asserts that Value is clamped
'   to 10 and PercentComplete is exactly 1.0. This covers the overflow edge
'   case to prevent the bar from exceeding 100%.
'@TestMethod("ProgressBar")
Public Sub TestUpdateClampsToMaximum()
    CustomTestSetTitles Assert, "ProgressBar", "TestUpdateClampsToMaximum"
    On Error GoTo TestFail

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:E1"), 10)

    sut.Update 999

    Assert.AreEqual CLng(10), sut.Value, _
                     "Update should clamp value to the maximum"
    Assert.AreEqual 1#, sut.PercentComplete, _
                     "PercentComplete should be 1 when clamped to maximum"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestUpdateClampsToMaximum", Err.Number, Err.Description
End Sub

' @sub-title TestUpdateClampsNegativeToZero
' @details
'   Verifies that Update clamps negative values to zero. Creates a bar with
'   maximum 10 and updates with value -5. Asserts that Value is clamped to 0.
'   This covers the underflow edge case to prevent negative progress display.
'@TestMethod("ProgressBar")
Public Sub TestUpdateClampsNegativeToZero()
    CustomTestSetTitles Assert, "ProgressBar", "TestUpdateClampsNegativeToZero"
    On Error GoTo TestFail

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:E1"), 10)

    sut.Update -5

    Assert.AreEqual CLng(0), sut.Value, _
                     "Update should clamp negative values to zero"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestUpdateClampsNegativeToZero", Err.Number, Err.Description
End Sub

' @sub-title TestStepByIncrementsValue
' @details
'   Verifies that StepBy accumulates value incrementally, both with an
'   explicit step size and with the default step of 1. Creates a bar with
'   maximum 100, then steps by 10, by default (1), and by 5 in sequence.
'   Asserts the cumulative values are 10, 11, and 16 respectively. This
'   confirms correct accumulation across multiple StepBy calls and that
'   the optional parameter defaults to 1.
'@TestMethod("ProgressBar")
Public Sub TestStepByIncrementsValue()
    CustomTestSetTitles Assert, "ProgressBar", "TestStepByIncrementsValue"
    On Error GoTo TestFail

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:J1"), 100)

    sut.StepBy 10
    Assert.AreEqual CLng(10), sut.Value, _
                     "StepBy 10 from 0 should yield 10"

    sut.StepBy
    Assert.AreEqual CLng(11), sut.Value, _
                     "StepBy without argument should increment by 1"

    sut.StepBy 5
    Assert.AreEqual CLng(16), sut.Value, _
                     "StepBy 5 from 11 should yield 16"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestStepByIncrementsValue", Err.Number, Err.Description
End Sub

' @sub-title TestCompleteReachesMaximum
' @details
'   Verifies that the Complete method sets Value to the Maximum regardless
'   of the current progress. Creates a bar with maximum 80, updates to 30,
'   then calls Complete. Asserts that Value equals 80 and PercentComplete
'   equals 1.0. This confirms Complete provides a shortcut to mark the
'   bar as fully done.
'@TestMethod("ProgressBar")
Public Sub TestCompleteReachesMaximum()
    CustomTestSetTitles Assert, "ProgressBar", "TestCompleteReachesMaximum"
    On Error GoTo TestFail

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:E1"), 80)

    sut.Update 30
    sut.Complete

    Assert.AreEqual CLng(80), sut.Value, _
                     "Complete should set value to maximum"
    Assert.AreEqual 1#, sut.PercentComplete, _
                     "PercentComplete should be 1 after Complete"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestCompleteReachesMaximum", Err.Number, Err.Description
End Sub

' @sub-title TestResetClearsValue
' @details
'   Verifies that Reset returns the bar to its initial zero state. Creates
'   a bar with maximum 50, updates to 25, then calls Reset. Asserts that
'   Value returns to 0 and PercentComplete returns to 0.0. This confirms
'   Reset is the inverse of Complete, restoring the bar for potential reuse.
'@TestMethod("ProgressBar")
Public Sub TestResetClearsValue()
    CustomTestSetTitles Assert, "ProgressBar", "TestResetClearsValue"
    On Error GoTo TestFail

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:E1"), 50)

    sut.Update 25
    sut.Reset

    Assert.AreEqual CLng(0), sut.Value, _
                     "Reset should set value back to zero"
    Assert.AreEqual 0#, sut.PercentComplete, _
                     "PercentComplete should be 0 after Reset"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestResetClearsValue", Err.Number, Err.Description
End Sub

'@section Tests - Status Cell
'===============================================================================

' @sub-title TestAttachStatusCellWritesMessages
' @details
'   Verifies that a status cell can be attached and that Update writes
'   status messages to it. Arranges a progress bar and attaches cell G1 as
'   the status cell. First asserts that the StatusCell address matches G1.
'   Then calls Update with value 5 and message "Processing..." and asserts
'   that G1 contains the expected text. This covers both the attachment and
'   the message-writing behaviour in a single flow.
'@TestMethod("ProgressBar")
Public Sub TestAttachStatusCellWritesMessages()
    CustomTestSetTitles Assert, "ProgressBar", "TestAttachStatusCellWritesMessages"
    On Error GoTo TestFail

    Dim statusCell As Range
    Set statusCell = FixtureSheet.Range("G1")

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:E1"), 10)
    sut.AttachStatusCell statusCell

    Assert.AreEqual statusCell.Address, sut.StatusCell.Address, _
                     "StatusCell should reference the attached cell"

    sut.Update 5, "Processing..."

    Assert.AreEqual "Processing...", CStr(statusCell.Value), _
                     "Update with status message should write to the status cell"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestAttachStatusCellWritesMessages", Err.Number, Err.Description
End Sub

' @sub-title TestStatusCellIsNothingByDefault
' @details
'   Verifies that StatusCell is Nothing when no cell has been explicitly
'   attached. Creates a bar without calling AttachStatusCell and asserts
'   that the StatusCell property returns Nothing. This confirms the safe
'   default state that prevents null reference errors during Update calls
'   when no status output is needed.
'@TestMethod("ProgressBar")
Public Sub TestStatusCellIsNothingByDefault()
    CustomTestSetTitles Assert, "ProgressBar", "TestStatusCellIsNothingByDefault"
    On Error GoTo TestFail

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:E1"), 10)

    Assert.IsTrue sut.StatusCell Is Nothing, _
                   "StatusCell should be Nothing when no cell is attached"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestStatusCellIsNothingByDefault", Err.Number, Err.Description
End Sub

' @sub-title TestAttachStatusCellRejectsMultiCellRange
' @details
'   Verifies the guard clause that rejects a multi-cell range as a status
'   cell. Creates a bar and attempts to attach a two-cell range (G1:H1).
'   Expects InvalidArgument to be raised. The status cell must be a single
'   cell because the bar writes a single text message to it, and a
'   multi-cell range would produce ambiguous output behaviour.
'@TestMethod("ProgressBar")
Public Sub TestAttachStatusCellRejectsMultiCellRange()
    CustomTestSetTitles Assert, "ProgressBar", "TestAttachStatusCellRejectsMultiCellRange"
    On Error GoTo ExpectError

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:E1"), 10)

    sut.AttachStatusCell FixtureSheet.Range("G1:H1")

    Assert.LogFailure "AttachStatusCell should raise for multi-cell range"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "AttachStatusCell should raise InvalidArgument for multi-cell range"
End Sub

'@section Tests - Maximum Property
'===============================================================================

' @sub-title TestMaximumSetterReclampsValue
' @details
'   Verifies that changing Maximum at runtime re-clamps the existing Value.
'   Creates a bar with maximum 100 and updates to 80. Then sets Maximum to 50,
'   which is below the current Value. Asserts that Maximum is now 50, Value
'   has been clamped down to 50, and PercentComplete is 1.0. This confirms
'   that the setter enforces the value-within-bounds invariant dynamically.
'@TestMethod("ProgressBar")
Public Sub TestMaximumSetterReclampsValue()
    CustomTestSetTitles Assert, "ProgressBar", "TestMaximumSetterReclampsValue"
    On Error GoTo TestFail

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:E1"), 100)

    sut.Update 80
    sut.Maximum = 50

    Assert.AreEqual CLng(50), sut.Maximum, _
                     "Maximum setter should update the maximum"
    Assert.AreEqual CLng(50), sut.Value, _
                     "Value should be clamped to the new maximum"
    Assert.AreEqual 1#, sut.PercentComplete, _
                     "PercentComplete should reflect new ratio after re-clamping"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestMaximumSetterReclampsValue", Err.Number, Err.Description
End Sub

'@section Tests - ConfigureValueFormat
'===============================================================================

' @sub-title TestConfigureValueFormatRejectsBlank
' @details
'   Verifies that ConfigureValueFormat rejects an empty format pattern string.
'   Creates a bar and calls ConfigureValueFormat with vbNullString. Expects
'   InvalidArgument to be raised. A blank pattern would produce empty or
'   meaningless rendered text in the bar cells, so the guard clause prevents
'   misconfiguration.
'@TestMethod("ProgressBar")
Public Sub TestConfigureValueFormatRejectsBlank()
    CustomTestSetTitles Assert, "ProgressBar", "TestConfigureValueFormatRejectsBlank"
    On Error GoTo ExpectError

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(FixtureSheet.Range("A1:E1"), 10)

    sut.ConfigureValueFormat vbNullString

    Assert.LogFailure "ConfigureValueFormat should raise for blank pattern"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "ConfigureValueFormat should raise InvalidArgument for blank pattern"
End Sub

'@section Tests - Rendering
'===============================================================================

' @sub-title TestRenderPaintsCompletedCells
' @details
'   Verifies that Update triggers visual rendering with correct cell colours.
'   Creates a bar on a five-cell range with maximum 5 and updates to 3. Then
'   iterates over each cell in the range, counting cells painted with the
'   completed colour (hex 4DB870, a green) and the pending colour (hex EEF1F5,
'   a light grey). Asserts 3 completed and 2 pending cells, confirming that
'   the rendering maps progress value to proportional cell colouring.
'@TestMethod("ProgressBar")
Public Sub TestRenderPaintsCompletedCells()
    CustomTestSetTitles Assert, "ProgressBar", "TestRenderPaintsCompletedCells"
    On Error GoTo TestFail

    Dim barRange As Range
    Set barRange = FixtureSheet.Range("A1:E1")

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(barRange, 5)

    sut.Update 3

    Dim completedCount As Long
    Dim pendingCount As Long
    Dim cell As Range

    For Each cell In barRange.Cells
        If cell.Interior.Color = &H4DB870 Then
            completedCount = completedCount + 1
        ElseIf cell.Interior.Color = &HEEF1F5 Then
            pendingCount = pendingCount + 1
        End If
    Next cell

    Assert.AreEqual CLng(3), completedCount, _
                     "Three cells should be painted with the completed colour"
    Assert.AreEqual CLng(2), pendingCount, _
                     "Two cells should remain in the pending colour"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestRenderPaintsCompletedCells", Err.Number, Err.Description
End Sub

' @sub-title TestRenderWritesFormattedValueToFirstCell
' @details
'   Verifies that Update writes a formatted "current / max" string into the
'   first cell of the bar range. Creates a bar with maximum 200 on a
'   five-cell range and updates to 50. Reads the text from the first cell
'   and asserts it equals "50 / 200". This confirms the default value
'   format pattern renders correctly in the visible cell.
'@TestMethod("ProgressBar")
Public Sub TestRenderWritesFormattedValueToFirstCell()
    CustomTestSetTitles Assert, "ProgressBar", "TestRenderWritesFormattedValueToFirstCell"
    On Error GoTo TestFail

    Dim barRange As Range
    Set barRange = FixtureSheet.Range("A1:E1")

    Dim sut As IProgressBar
    Set sut = ProgressBar.Create(barRange, 200)

    sut.Update 50

    Dim cellText As String
    cellText = CStr(barRange.Cells(1).Value)

    Assert.AreEqual "50 / 200", cellText, _
                     "First cell should display formatted current/max values"
    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestRenderWritesFormattedValueToFirstCell", Err.Number, Err.Description
End Sub
