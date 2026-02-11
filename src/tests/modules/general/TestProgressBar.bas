Attribute VB_Name = "TestProgressBar"
Attribute VB_Description = "Tests for ProgressBar class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for ProgressBar class")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const PROGRESSBAR_SHEET As String = "ProgressBarFixture"

Private Assert As ICustomTest
Private FixtureSheet As Worksheet

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestProgressBar"
End Sub

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

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureSheet = EnsureWorksheet(PROGRESSBAR_SHEET, visibility:=xlSheetHidden)
End Sub

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
