Attribute VB_Name = "TestCustomTest"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests for the CustomTest harness class")

'@description
'  Unit tests for the CustomTest class, which is the project's own test harness
'  replacing Rubberduck assertions with an IChecking-based result model. Tests
'  validate that AreEqual, AreNotEqual, IsTrue, and IsFalse correctly log both
'  successes and failures, that ObjectExists performs type checking and IsNothing
'  detects unset references, that Flush queues results and BeginTest respects the
'  resetNames flag, that direct LogFailure and LogSuccesses calls produce the
'  expected entries, and that PrintResults writes structured output to a worksheet
'  with filter integration and subtitle overrides.
'@depends CustomTest, ICustomTest, Checking, IChecking, CheckingOutput, BetterArray, TestHelpers

Private Const OUTPUT_SHEET_NAME As String = "HarnessOutput"
Private Const VISIBLE_COLUMN_COUNT As Long = 3
Private Const FIRST_VISIBLE_COLUMN_INDEX As Long = 3

Private Assert As Object
Private Harness As ICustomTest
Private HarnessWorkbook As Workbook
Private OutputSheet As Worksheet
Private Results As BetterArray

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set HarnessWorkbook = NewWorkbook()
    Set OutputSheet = EnsureWorksheet(OUTPUT_SHEET_NAME, HarnessWorkbook)
    Set Harness = CustomTest.Create(HarnessWorkbook, OUTPUT_SHEET_NAME)
    Harness.SetModuleName "TestCustomTest"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Harness Is Nothing Then Harness.ReleaseResources
    On Error GoTo 0

    'Delete the current custom test workbook
    DeleteWorkbook HarnessWorkbook
    Set Harness = Nothing
    Set OutputSheet = Nothing
    Set HarnessWorkbook = Nothing
    Set Assert = Nothing
    Set Results = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    If Not OutputSheet Is Nothing Then OutputSheet.Cells.Clear
    If Not Harness Is Nothing Then Harness.ResetInstance
    Set Results = Nothing
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Results = Nothing
End Sub

'@sub-title Verify AreEqual logs success for matching values and error for mismatched values
'@details Arranges two AreEqual assertions: one with identical numbers (42, 42) and one
'  with different strings ("alpha" vs "beta"). Flushes the current test to retrieve the
'  IChecking log, then asserts that the first key is marked Success and the second is
'  marked Error with the supplied failure message captured in the label.
'@TestMethod("Harness")
Private Sub TestAreEqualCapturesSuccessAndFailure()
    On Error GoTo Fail

    Dim checkLog As IChecking
    Dim keys As BetterArray
    Dim firstKey As String
    Dim secondKey As String

    Harness.SetTestName "Equality"
    Harness.BeginTest
    Harness.AreEqual 42, 42, "Matching numbers"
    Harness.AreEqual "alpha", "beta", "Mismatched text"

    Set Results = Harness.FlushCurrentTest(False)
    Assert.IsTrue (Results.Length = 1), "Flush should yield a single checking"

    Set checkLog = Results.Item(Results.LowerBound)
    Assert.AreEqual "Equality", checkLog.Heading
    Assert.AreEqual "test", checkLog.Heading(True)

    Set keys = checkLog.ListOfKeys
    Assert.IsTrue (keys.Length = 2), "Two assertions expected"

    firstKey = CStr(keys.Item(keys.LowerBound))
    Assert.IsTrue InStr(1, checkLog.ValueOf(firstKey, checkingType), "Success", vbTextCompare) > 0, _
        "Successful comparison should log as success"

    secondKey = CStr(keys.Item(keys.LowerBound + 1))
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingType), "Error", vbTextCompare) > 0, _
        "Failed comparison should log as error"
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingLabel), "Mismatched text", vbTextCompare) > 0, _
        "Failure entry should capture the supplied message"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAreEqualCapturesSuccessAndFailure"
End Sub

'@sub-title Verify AreNotEqual logs success for distinct values and error for matching values
'@details Arranges two AreNotEqual assertions: one with different numbers (10, 42) and one
'  with identical strings ("same", "same"). After flushing, confirms the first key is Success
'  and the second is Error, with the failure label referencing both the expected inequality
'  and the actual matched value.
'@TestMethod("Harness")
Private Sub TestAreNotEqualCapturesSuccessAndFailure()
    On Error GoTo Fail

    Dim checkLog As IChecking
    Dim keys As BetterArray
    Dim firstKey As String
    Dim secondKey As String

    Harness.SetTestName "Inequality"
    Harness.BeginTest
    Harness.AreNotEqual 10, 42, "Distinct values should pass"
    Harness.AreNotEqual "same", "same", "Matching values should fail"

    Set Results = Harness.FlushCurrentTest(False)
    Assert.IsTrue (Results.Length = 1), "Flush should yield a single checking"

    Set checkLog = Results.Item(Results.LowerBound)
    Set keys = checkLog.ListOfKeys
    Assert.IsTrue (keys.Length = 2), "Two assertions expected"

    firstKey = CStr(keys.Item(keys.LowerBound))
    Assert.IsTrue InStr(1, checkLog.ValueOf(firstKey, checkingType), "Success", vbTextCompare) > 0, _
        "Different values should log as success"

    secondKey = CStr(keys.Item(keys.LowerBound + 1))
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingType), "Error", vbTextCompare) > 0, _
        "Matching values should log as error"
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingLabel), "Values to differ", vbTextCompare) > 0, _
        "Failure entry should indicate the expected inequality"
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingLabel), "Values matched", vbTextCompare) > 0, _
        "Failure entry should capture the actual match"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAreNotEqualCapturesSuccessAndFailure"
End Sub

'@sub-title Verify IsFalse logs success for False conditions and error for True conditions
'@details Arranges two IsFalse assertions: one passing False (expected success) and one
'  passing True (expected failure). After flushing, confirms the first entry is logged as
'  Success and the second as Error, with the failure label containing the user-supplied
'  message describing the unexpected True condition.
'@TestMethod("Harness")
Private Sub TestIsFalseCapturesSuccessAndFailure()
    On Error GoTo Fail

    Dim checkLog As IChecking
    Dim keys As BetterArray
    Dim firstKey As String
    Dim secondKey As String

    Harness.SetTestName "IsFalse"
    Harness.BeginTest
    Harness.IsFalse False, "Condition must be false"
    Harness.IsFalse True, "Condition unexpectedly true"

    Set Results = Harness.FlushCurrentTest()
    Assert.IsTrue (Results.Length = 1), "Flush should yield a single checking"

    Set checkLog = Results.Item(Results.LowerBound)
    Set keys = checkLog.ListOfKeys
    Assert.IsTrue (keys.Length = 2), "Two assertions expected"

    firstKey = CStr(keys.Item(keys.LowerBound))
    Assert.IsTrue InStr(1, checkLog.ValueOf(firstKey, checkingType), "Success", vbTextCompare) > 0, _
        "False condition should log success"

    secondKey = CStr(keys.Item(keys.LowerBound + 1))
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingType), "Error", vbTextCompare) > 0, _
        "True condition should log as error"
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingLabel), "Condition unexpectedly true", vbTextCompare) > 0, _
        "Failure entry should include supplied message"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestIsFalseCapturesSuccessAndFailure"
End Sub

'@sub-title Verify SetTestSubtitle overrides the default subtitle in the checking heading
'@details Sets a custom subtitle via SetTestSubtitle before calling BeginTest and IsTrue.
'  After flushing, asserts that the checking Heading returns the test name for the main
'  heading and the custom subtitle string for the secondary heading.
'@TestMethod("Harness")
Private Sub TestSubtitleOverridesDefault()
    On Error GoTo Fail

    Dim checkLog As IChecking

    Harness.SetTestName "Boolean"
    Harness.SetTestSubtitle "Custom subtitle"
    Harness.BeginTest
    Harness.IsTrue True, vbNullString

    Set Results = Harness.FlushCurrentTest()
    Assert.IsTrue (Results.Length = 1)

    Set checkLog = Results.Item(Results.LowerBound)
    Assert.AreEqual "Boolean", checkLog.Heading
    Assert.AreEqual "Custom subtitle", checkLog.Heading(True)
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestSubtitleOverridesDefault"
End Sub

'@sub-title Verify ObjectExists succeeds for valid references and fails for Nothing and type mismatches
'@details Arranges three ObjectExists assertions: a live Worksheet (should pass), an unassigned
'  Worksheet reference (Nothing -- should fail), and a Collection instance tested against the
'  "Worksheet" type name (type mismatch -- should fail). After flushing, confirms the first
'  entry is Success, the second is Error referencing the expected type, and the third is Error
'  referencing the actual Collection type.
'@TestMethod("Harness")
Private Sub TestObjectExistsValidatesReferences()
    On Error GoTo Fail

    Dim checkLog As IChecking
    Dim keys As BetterArray
    Dim firstKey As String
    Dim secondKey As String
    Dim thirdKey As String
    '@Ignore VariableNotAssigned
    Dim missingSheet As Worksheet
    Dim otherInstance As Object

    Harness.SetTestName "Existence"
    Harness.BeginTest

    Harness.ObjectExists OutputSheet, "Worksheet", "Worksheet reference should exist"

    '@Ignore UnassignedVariableUsage
    Harness.ObjectExists missingSheet, "Worksheet", "Missing worksheet should fail"

    Set otherInstance = New Collection
    Harness.ObjectExists otherInstance, "Worksheet", "Type mismatch should fail"

    Set Results = Harness.FlushCurrentTest()
    Assert.IsTrue (Results.Length = 1), "Flush should yield a single checking"

    Set checkLog = Results.Item(Results.LowerBound)
    Set keys = checkLog.ListOfKeys
    Assert.IsTrue (keys.Length = 3), "Three assertions expected"

    firstKey = CStr(keys.Item(keys.LowerBound))
    Assert.IsTrue InStr(1, checkLog.ValueOf(firstKey, checkingType), "Success", vbTextCompare) > 0, _
        "Existing object should log success"

    secondKey = CStr(keys.Item(keys.LowerBound + 1))
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingType), "Error", vbTextCompare) > 0, _
        "Nothing reference should log as error"
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingLabel), "expected: Instance of type 'Worksheet'", vbTextCompare) > 0, _
        "Failure entry should reference expected type"

    thirdKey = CStr(keys.Item(keys.LowerBound + 2))
    Assert.IsTrue InStr(1, checkLog.ValueOf(thirdKey, checkingType), "Error", vbTextCompare) > 0, _
        "Type mismatch should log as error"
    Assert.IsTrue InStr(1, checkLog.ValueOf(thirdKey, checkingLabel), "actual: Instance of type 'Collection'", vbTextCompare) > 0, _
        "Failure entry should capture actual type"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestObjectExistsValidatesReferences"
End Sub

'@sub-title Verify IsNothing logs success for Nothing references and error for assigned references
'@details Arranges two IsNothing assertions: an unassigned Worksheet variable (should succeed)
'  and the same variable after assignment to OutputSheet (should fail). After flushing, asserts
'  that the first key is marked Success, and the second is Error with the label mentioning both
'  "Nothing" and "Worksheet" to indicate the expected vs actual state.
'@TestMethod("Harness")
Private Sub TestIsNothingCapturesSuccessAndFailure()
    On Error GoTo Fail

    Dim checkLog As IChecking
    Dim keys As BetterArray
    Dim firstKey As String
    Dim secondKey As String
    Dim populatedSheet As Worksheet

    Harness.SetTestName "IsNothing"
    Harness.BeginTest

    Harness.IsNothing populatedSheet, "Unassigned reference should be Nothing"

    Set populatedSheet = OutputSheet
    Harness.IsNothing populatedSheet, "Assigned reference should fail"

    Set Results = Harness.FlushCurrentTest()
    Assert.IsTrue (Results.Length = 1), "Flush should yield a single checking"

    Set checkLog = Results.Item(Results.LowerBound)
    Set keys = checkLog.ListOfKeys
    Assert.IsTrue (keys.Length = 2), "Two assertions expected"

    firstKey = CStr(keys.Item(keys.LowerBound))
    Assert.IsTrue InStr(1, checkLog.ValueOf(firstKey, checkingType), "Success", vbTextCompare) > 0, _
        "Nothing reference should log success"

    secondKey = CStr(keys.Item(keys.LowerBound + 1))
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingType), "Error", vbTextCompare) > 0, _
        "Assigned reference should log as error"
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingLabel), "Nothing", vbTextCompare) > 0, _
        "Failure entry should reference the Nothing expectation"
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingLabel), "Worksheet", vbTextCompare) > 0, _
        "Failure entry should capture the detected type"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestIsNothingCapturesSuccessAndFailure"
End Sub

'@sub-title Verify Flush persists the current checking to the results buffer without returning it
'@details Calls Flush (not FlushCurrentTest) after recording a single assertion, then inspects
'  the Results buffer to confirm the checking was persisted with correct heading and subtitle.
'  A second Flush call without an active checking verifies the no-op behavior by asserting the
'  results length remains unchanged.
'@TestMethod("Harness")
Private Sub TestFlushQueuesCurrentCheckingWithoutReturn()
    On Error GoTo Fail

    Dim resultsBuffer As BetterArray
    Dim persisted As IChecking
    Dim keys As BetterArray

    Harness.SetTestName "FlushOnly"
    Harness.SetTestSubtitle "One shot"
    Harness.BeginTest
    Harness.IsTrue True, "Single assertion"

    Harness.Flush

    Set resultsBuffer = Harness.Results
    Assert.IsTrue (resultsBuffer.Length = 1), "Flush should persist the current checking"

    Set persisted = resultsBuffer.Item(resultsBuffer.LowerBound)
    Assert.AreEqual "FlushOnly", persisted.Heading
    Assert.AreEqual "One shot", persisted.Heading(True)

    Set keys = persisted.ListOfKeys
    Assert.IsTrue (keys.Length = 1), "Single entry expected after flush"

    Harness.Flush
    Set resultsBuffer = Harness.Results
    Assert.IsTrue (resultsBuffer.Length = 1), "Flushing without an active checking should be a no-op"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestFlushQueuesCurrentCheckingWithoutReturn"
End Sub

'@sub-title Verify Flush and BeginTest respect the resetNames flag for sticky vs cleared names
'@details Performs three rounds of flush: two with resetNames=False (names should persist) and
'  one preceded by BeginTest(True) (names should reset to defaults). Asserts that the first two
'  checkings retain the originally configured name and subtitle, while the third checking reverts
'  both heading and subtitle to the default "test" value.
'@TestMethod("Harness")
Private Sub TestFlushRespectsResetNamesFlag()
    On Error GoTo Fail

    Dim resultsBuffer As BetterArray
    Dim firstLog As IChecking
    Dim secondLog As IChecking
    Dim thirdLog As IChecking

    Harness.SetTestName "Sticky"
    Harness.SetTestSubtitle "Subtitle"
    Harness.BeginTest
    Harness.IsTrue True, "Initial assertion"
    Harness.Flush False

    Harness.BeginTest
    Harness.IsTrue True, "Second assertion"
    Harness.Flush False

    Set resultsBuffer = Harness.Results
    Assert.IsTrue (resultsBuffer.Length = 2), "Two checkings expected after consecutive flushes"

    Set firstLog = resultsBuffer.Item(resultsBuffer.LowerBound)
    Set secondLog = resultsBuffer.Item(resultsBuffer.LowerBound + 1)
    Assert.AreEqual "Sticky", firstLog.Heading
    Assert.AreEqual "Sticky", secondLog.Heading
    Assert.AreEqual "Subtitle", secondLog.Heading(True)

    Harness.BeginTest True 'this will reset the names
    Harness.IsTrue True, "Third assertion"
    Harness.Flush

    Set resultsBuffer = Harness.Results
    Assert.IsTrue (resultsBuffer.Length = 3), "Third flush should append another checking"

    Set thirdLog = resultsBuffer.Item(resultsBuffer.LowerBound + 2)
    Assert.AreEqual "test", thirdLog.Heading, "Names should reset when resetNames defaults to True"
    Assert.AreEqual "test", thirdLog.Heading(True)
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestFlushRespectsResetNamesFlag"
End Sub

'@sub-title Verify LogFailure and LogSuccesses directly record typed entries in the checking log
'@details Calls LogSuccesses and LogFailure on the harness, then flushes and inspects the
'  resulting IChecking. Asserts that the first entry is marked Success with the correct label
'  text, and the second entry is marked Error with the corresponding failure message.
'@TestMethod("Harness")
Private Sub TestLogFailureAndLogSuccesses()
    On Error GoTo Fail

    Dim checkLog As IChecking
    Dim keys As BetterArray
    Dim firstKey As String
    Dim secondKey As String

    Harness.SetTestName "Direct logging"
    Harness.BeginTest
    Harness.LogSuccesses "Operation completed"
    Harness.LogFailure "Operation failed"

    Set Results = Harness.FlushCurrentTest()
    Assert.IsTrue (Results.Length = 1), "Flush should yield a single checking"

    Set checkLog = Results.Item(Results.LowerBound)
    Set keys = checkLog.ListOfKeys
    Assert.IsTrue (keys.Length = 2), "Two log entries expected"

    firstKey = CStr(keys.Item(keys.LowerBound))
    Assert.IsTrue InStr(1, checkLog.ValueOf(firstKey, checkingType), "Success", vbTextCompare) > 0, _
        "LogSuccesses should mark entry as success"
    Assert.IsTrue InStr(1, checkLog.ValueOf(firstKey, checkingLabel), "Operation completed", vbTextCompare) > 0, _
        "Success entry should include supplied message"

    secondKey = CStr(keys.Item(keys.LowerBound + 1))
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingType), "Error", vbTextCompare) > 0, _
        "LogFailure should mark entry as error"
    Assert.IsTrue InStr(1, checkLog.ValueOf(secondKey, checkingLabel), "Operation failed", vbTextCompare) > 0, _
        "Failure entry should include supplied message"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestLogFailureAndLogSuccesses"
End Sub

'@sub-title Verify PrintResults clears the worksheet, writes batch output, and integrates with filtering
'@details Queues two test checkings (one passing, one failing) to the harness while a sentinel
'  value occupies an unrelated cell. After calling PrintResults, asserts the sentinel was cleared,
'  the filter header was written, and both test names appear exactly once. Then exercises
'  CheckingOutput.FilterWorksheet to confirm Success rows toggle visibility when the filter
'  value switches between "Without Successes" and "Successes".
'@TestMethod("Harness")
Private Sub TestPrintResultsClearsWorksheetAndWritesBatch()
    On Error GoTo Fail

    Dim sh As Worksheet
    Dim successCell As Range

    Set sh = OutputSheet
    sh.Cells.Clear
    sh.Cells(10, 5).Value = "Sentinel"

    Harness.SetTestName "First"
    Harness.SetTestSubtitle "Basics"
    Harness.BeginTest
    Harness.IsTrue True, "Passing assertion"
    Set Results = Harness.FlushCurrentTest()
    Assert.IsTrue (Results.Length = 1), "First flush should queue single checking"

    Harness.SetTestName "Second"
    Harness.SetTestSubtitle "Failure"
    Harness.BeginTest
    Harness.IsTrue False, "Forced failure"
    Set Results = Harness.FlushCurrentTest()
    Assert.IsTrue (Results.Length = 2), "Second flush should accumulate results"

    Assert.AreEqual "Sentinel", sh.Cells(10, 5).Value, _
        "Worksheet must remain untouched until PrintResults executes"

    Harness.PrintResults OUTPUT_SHEET_NAME

    Assert.AreEqual "Status:", sh.Cells(1, 2).Value, "Filter header should be written"
    Assert.AreEqual vbNullString, CStr(sh.Cells(10, 5).Value), "Worksheet should be cleared before writing"
    Assert.IsTrue (CountOccurrences(sh, "First") = 1), "First test title must be written once"
    Assert.IsTrue (CountOccurrences(sh, "Second") = 1), "Second test title must be written once"

    Set successCell = sh.Cells.Find(What:="Success", LookIn:=xlValues, LookAt:=xlPart)
    If Not successCell Is Nothing Then
        sh.Range("C1").Value = "Without Successes"
        CheckingOutput.Create(sh).FilterWorksheet "Without Successes"
        Assert.IsTrue sh.Rows(successCell.Row).Hidden, "Success rows should hide when filtering out successes"

        sh.Range("C1").Value = "Successes"
        CheckingOutput.Create(sh).FilterWorksheet "Successes"
        Assert.IsFalse sh.Rows(successCell.Row).Hidden, "Success rows should show when filtering successes"
    Else
        Assert.Fail "Success entry not found after printing"
    End If
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestPrintResultsClearsWorksheetAndWritesBatch"
End Sub

'@sub-title Count occurrences of a text value in the visible (or full) used range of a worksheet
Private Function CountOccurrences(ByVal sh As Worksheet, ByVal textValue As String, _
                                  Optional ByVal includeHiddenColumns As Boolean = False) As Long
    Dim searchRange As Range
    Dim lastRow As Long

    If includeHiddenColumns Then
        Set searchRange = sh.UsedRange
    Else
        lastRow = sh.Cells(sh.Rows.Count, FIRST_VISIBLE_COLUMN_INDEX).End(xlUp).Row
        If lastRow < 1 Then lastRow = 1
        Set searchRange = sh.Range(sh.Cells(1, FIRST_VISIBLE_COLUMN_INDEX), _
                                   sh.Cells(lastRow, FIRST_VISIBLE_COLUMN_INDEX + VISIBLE_COLUMN_COUNT - 1))
    End If

    On Error Resume Next
        CountOccurrences = Application.WorksheetFunction.CountIf(searchRange, textValue)
    On Error GoTo 0
End Function
