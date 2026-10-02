Attribute VB_Name = "TestMessenger"
Attribute VB_Description = "Tests for the Messenger class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the Messenger class")

'@description
'Drives Messenger, the one place a message box is shown from. An armed
'messenger shows nothing, writes the text down and answers with the answer the
'call site named; a disarmed one calls MsgBox.
'
'EVERY TEST HERE ARMS FIRST
'-------------------------------------------------------------------------------
'Show on a disarmed messenger opens a real box, and a box nobody clicks holds
'the whole run until the transport gives up. So no test in this suite calls
'Show while the messenger is disarmed. The disarmed path is one MsgBox call
'with the arguments it was handed, and it is checked by hand in Excel.
'
'THE RECORD IS SHARED
'-------------------------------------------------------------------------------
'Messenger has no Create and every caller uses the default instance, so the
'record survives from one test to the next. TestInitialize and TestCleanup both
'call Reset, which empties the record, disarms and drops the force state.
'
'THE STORED SWITCH IS WRITTEN ON THIS WORKBOOK
'-------------------------------------------------------------------------------
'ReadStoredSwitch takes a Workbook, so the name has to sit at workbook level.
'The suite writes it on ThisWorkbook through HiddenNames, the way the three
'workbooks will, and removes it before and after every test.
'@depends Messenger, HiddenNames, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'What the class says it stores the switch under. The tests pin the string
'itself so a rename cannot pass unnoticed on either side of the wire.
Private Const EXPECTED_SWITCH_NAME As String = "__OBT__SILENT_OPERATIONS__"

Private Assert As CustomTest


'@section Module lifecycle
'===============================================================================

'@sub-title Set up the assertion harness.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestMessenger"
End Sub

'@sub-title Print results and leave the messenger and the workbook clean.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    ClearStoredSwitch
    Messenger.Reset
    Set Assert = Nothing
    RestoreApp
End Sub

'@sub-title Empty the record and remove the stored switch before each test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Messenger.Reset
    ClearStoredSwitch
End Sub

'@sub-title Flush assert state and leave nothing armed.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    Messenger.Reset
    ClearStoredSwitch
End Sub


'@section Helper routines
'===============================================================================

'@sub-title Write the silence switch on this workbook.
'@param value String. Yes or No.
Private Sub SetStoredSwitch(ByVal value As String)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(ThisWorkbook)
    store.EnsureName EXPECTED_SWITCH_NAME, value, HiddenNameTypeString
    store.SetValue EXPECTED_SWITCH_NAME, value
End Sub

'@sub-title Remove the silence switch from this workbook.
'@details
'A test that never wrote the name still calls this, and removing a name that
'was never added is an ordinary answer rather than a fault.
Private Sub ClearStoredSwitch()
    Dim store As HiddenNames

    On Error Resume Next
        Set store = HiddenNames.Create(ThisWorkbook)
        store.RemoveName EXPECTED_SWITCH_NAME
    On Error GoTo 0
End Sub

'@sub-title How many lines the record holds.
'@return Long. The number of swallowed messages.
Private Function RecordedLineCount() As Long
    Dim recorded As String

    recorded = Messenger.Messages()
    If LenB(recorded) = 0 Then Exit Function

    RecordedLineCount = UBound(Split(recorded, vbNewLine)) + 1
End Function


'@section Arming and disarming
'===============================================================================

'@sub-title Arm turns the boxes off and Disarm turns them back on.
'@TestMethod("Messenger")
Public Sub TestArmAndDisarmToggleTheBoxes()
    CustomTestSetTitles Assert, "Messenger", "TestArmAndDisarmToggleTheBoxes"
    On Error GoTo TestFail

    Assert.IsFalse Messenger.Armed, "A messenger that was reset is disarmed"

    Messenger.Arm ThisWorkbook
    Assert.IsTrue Messenger.Armed, "Arm turns the boxes off"

    Messenger.Disarm
    Assert.IsFalse Messenger.Armed, "Disarm turns the boxes back on"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestArmAndDisarmToggleTheBoxes", Err.Number, Err.Description
End Sub

'@sub-title Arm remembers the workbook the run is about.
'@TestMethod("Messenger")
Public Sub TestArmRemembersItsHostWorkbook()
    CustomTestSetTitles Assert, "Messenger", "TestArmRemembersItsHostWorkbook"
    On Error GoTo TestFail

    Assert.IsNothing Messenger.HostBook, "A messenger that was reset holds no workbook"

    Messenger.Arm ThisWorkbook

    Assert.IsNotNothing Messenger.HostBook, "Arm keeps the workbook it was given"
    Assert.AreEqual ThisWorkbook.Name, Messenger.HostBook.Name, _
                    "The workbook kept is the one Arm was given"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestArmRemembersItsHostWorkbook", Err.Number, Err.Description
End Sub

'@sub-title Arm without a workbook refuses.
'@details
'The wrapper always has the workbook it is working on, and a messenger armed
'over Nothing would hide every box of that run with no run to attach them to.
'@TestMethod("Messenger")
Public Sub TestArmWithoutAWorkbookRaises()
    CustomTestSetTitles Assert, "Messenger", "TestArmWithoutAWorkbookRaises"
    On Error GoTo ExpectedError

    Messenger.Arm Nothing

    Assert.Fail "Arm must refuse a run with no workbook behind it"
    Exit Sub

ExpectedError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), CLng(Err.Number), _
                    "Arm with no workbook raises ObjectNotInitialized"
    Assert.IsFalse Messenger.Armed, "A refused Arm leaves the boxes on"
End Sub

'@sub-title Arm empties the record of the run before it.
'@details
'Arm opens a run. A wrapper reading Messages after its own run must never see
'a line the run before swallowed.
'@TestMethod("Messenger")
Public Sub TestArmEmptiesTheRecordOfTheRunBefore()
    CustomTestSetTitles Assert, "Messenger", "TestArmEmptiesTheRecordOfTheRunBefore"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook
    Messenger.Show "the first run said this", vbOK
    Messenger.Disarm

    Assert.IsTrue Messenger.HasMessages(), "The first run swallowed one message"

    Messenger.Arm ThisWorkbook

    Assert.IsFalse Messenger.HasMessages(), "Arm opens the second run with an empty record"
    Assert.AreEqual vbNullString, Messenger.Messages(), _
                    "Messages answers nothing at the start of a run"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestArmEmptiesTheRecordOfTheRunBefore", Err.Number, Err.Description
End Sub

'@sub-title Disarm keeps the record.
'@details
'The summary a wrapper answers is built after the run is over, so the lines
'have to outlive the silence.
'@TestMethod("Messenger")
Public Sub TestDisarmKeepsTheRecord()
    CustomTestSetTitles Assert, "Messenger", "TestDisarmKeepsTheRecord"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook
    Messenger.Show "the export folder was not found", vbOK
    Messenger.Disarm

    Assert.IsTrue Messenger.HasMessages(), "The record outlives Disarm"
    Assert.AreEqual "the export folder was not found", Messenger.Messages(), _
                    "The text is still readable once the boxes are back on"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDisarmKeepsTheRecord", Err.Number, Err.Description
End Sub

'@sub-title Reset puts the whole messenger back to its opening state.
'@TestMethod("Messenger")
Public Sub TestResetPutsEverythingBack()
    CustomTestSetTitles Assert, "Messenger", "TestResetPutsEverythingBack"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook, force:=True
    Messenger.Show "something happened", vbOK

    Messenger.Reset

    Assert.IsFalse Messenger.Armed, "Reset turns the boxes back on"
    Assert.IsFalse Messenger.HasMessages(), "Reset empties the record"
    Assert.IsNothing Messenger.HostBook, "Reset forgets the workbook"
    Assert.AreEqual CLng(vbNo), CLng(Messenger.CarryOn()), "Reset drops the force state"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestResetPutsEverythingBack", Err.Number, Err.Description
End Sub


'@section The force state
'===============================================================================

'@sub-title A run armed without force stops at a warning.
'@TestMethod("Messenger")
Public Sub TestCarryOnRefusesWhenTheRunWasNotForced()
    CustomTestSetTitles Assert, "Messenger", "TestCarryOnRefusesWhenTheRunWasNotForced"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook

    Assert.AreEqual CLng(vbNo), CLng(Messenger.CarryOn()), _
                    "A run that was not forced answers no to a warning"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCarryOnRefusesWhenTheRunWasNotForced", _
                         Err.Number, Err.Description
End Sub

'@sub-title A run armed with force pushes past a warning.
'@TestMethod("Messenger")
Public Sub TestCarryOnAgreesWhenTheRunWasForced()
    CustomTestSetTitles Assert, "Messenger", "TestCarryOnAgreesWhenTheRunWasForced"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook, force:=True

    Assert.AreEqual CLng(vbYes), CLng(Messenger.CarryOn()), _
                    "A forced run answers yes to a warning"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCarryOnAgreesWhenTheRunWasForced", _
                         Err.Number, Err.Description
End Sub

'@sub-title Disarm drops the force state.
'@details
'Force belongs to one call. The wrapper after it starts from no.
'@TestMethod("Messenger")
Public Sub TestDisarmDropsTheForceState()
    CustomTestSetTitles Assert, "Messenger", "TestDisarmDropsTheForceState"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook, force:=True
    Messenger.Disarm

    Assert.AreEqual CLng(vbNo), CLng(Messenger.CarryOn()), _
                    "The force state is gone once the run ends"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDisarmDropsTheForceState", Err.Number, Err.Description
End Sub

'@sub-title A second Arm without force clears the force of the one before.
'@TestMethod("Messenger")
Public Sub TestArmWithoutForceClearsTheForceBefore()
    CustomTestSetTitles Assert, "Messenger", "TestArmWithoutForceClearsTheForceBefore"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook, force:=True
    Messenger.Arm ThisWorkbook

    Assert.AreEqual CLng(vbNo), CLng(Messenger.CarryOn()), _
                    "The second run was not forced and answers no"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestArmWithoutForceClearsTheForceBefore", _
                         Err.Number, Err.Description
End Sub


'@section Showing while armed
'===============================================================================

'@sub-title Show answers the silent answer the call site named.
'@details
'Three call sites of the sweep want three different answers out of the same
'class: vbOK for a box with one button, vbYes for the import question R itself
'asked for, vbNo for the offer to open a report nobody is there to read.
'@TestMethod("Messenger")
Public Sub TestShowAnswersTheSilentAnswerItWasGiven()
    CustomTestSetTitles Assert, "Messenger", "TestShowAnswersTheSilentAnswerItWasGiven"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook

    Assert.AreEqual CLng(vbOK), CLng(Messenger.Show("the tags were updated", vbOK)), _
                    "A box with one button answers vbOK"
    Assert.AreEqual CLng(vbYes), _
                    CLng(Messenger.Show("import this file?", vbYes, vbYesNo)), _
                    "A question told to answer yes answers vbYes"
    Assert.AreEqual CLng(vbNo), _
                    CLng(Messenger.Show("open the report now?", vbNo, vbYesNo)), _
                    "A question told to answer no answers vbNo"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestShowAnswersTheSilentAnswerItWasGiven", _
                         Err.Number, Err.Description
End Sub

'@sub-title Show writes down what it swallowed.
'@TestMethod("Messenger")
Public Sub TestShowRecordsTheTextItSwallowed()
    CustomTestSetTitles Assert, "Messenger", "TestShowRecordsTheTextItSwallowed"
    On Error GoTo TestFail

    Assert.IsFalse Messenger.HasMessages(), "Nothing is recorded before the run starts"

    Messenger.Arm ThisWorkbook
    Messenger.Show "the geobase was imported", vbOK

    Assert.IsTrue Messenger.HasMessages(), "A swallowed box leaves a line behind"
    Assert.AreEqual "the geobase was imported", Messenger.Messages(), _
                    "The line is the text of the box"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestShowRecordsTheTextItSwallowed", Err.Number, Err.Description
End Sub

'@sub-title Every swallowed box takes one line of the record.
'@TestMethod("Messenger")
Public Sub TestEverySwallowedBoxTakesOneLine()
    CustomTestSetTitles Assert, "Messenger", "TestEverySwallowedBoxTakesOneLine"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook
    Messenger.Show "first", vbOK
    Messenger.Show "second", vbOK
    Messenger.Show "third", vbOK

    Assert.AreEqual CLng(3), RecordedLineCount(), "Three boxes leave three lines"
    Assert.AreEqual "first" & vbNewLine & "second" & vbNewLine & "third", _
                    Messenger.Messages(), _
                    "The lines read back in the order they were swallowed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEverySwallowedBoxTakesOneLine", Err.Number, Err.Description
End Sub

'@sub-title A title goes in front of the message on the recorded line.
'@TestMethod("Messenger")
Public Sub TestTheTitleGoesInFrontOfTheMessage()
    CustomTestSetTitles Assert, "Messenger", "TestTheTitleGoesInFrontOfTheMessage"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook
    Messenger.Show "the file has no metadata", vbOK, vbOKOnly, "Import"

    Assert.AreEqual "Import: the file has no metadata", Messenger.Messages(), _
                    "The recorded line reads title then message"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheTitleGoesInFrontOfTheMessage", _
                         Err.Number, Err.Description
End Sub

'@sub-title A message carrying line breaks still takes one line.
'@details
'A reader counts the lines of the record to count what a run swallowed, so a
'box whose text runs over three lines must not read as three boxes.
'@TestMethod("Messenger")
Public Sub TestALineBreakInsideAMessageIsFlattened()
    CustomTestSetTitles Assert, "Messenger", "TestALineBreakInsideAMessageIsFlattened"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook
    Messenger.Show "the import stopped" & vbNewLine & "three rows were refused", vbOK

    Assert.AreEqual CLng(1), RecordedLineCount(), "One box leaves one line"
    Assert.AreEqual "the import stopped three rows were refused", Messenger.Messages(), _
                    "Every break in the text became a space"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALineBreakInsideAMessageIsFlattened", _
                         Err.Number, Err.Description
End Sub

'@sub-title A run that swallowed nothing answers an empty record.
'@TestMethod("Messenger")
Public Sub TestARunThatSwallowedNothingAnswersEmpty()
    CustomTestSetTitles Assert, "Messenger", "TestARunThatSwallowedNothingAnswersEmpty"
    On Error GoTo TestFail

    Messenger.Arm ThisWorkbook
    Messenger.Disarm

    Assert.IsFalse Messenger.HasMessages(), "A quiet run has no messages"
    Assert.AreEqual vbNullString, Messenger.Messages(), _
                    "A quiet run answers an empty string"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARunThatSwallowedNothingAnswersEmpty", _
                         Err.Number, Err.Description
End Sub


'@section The stored switch
'===============================================================================

'@sub-title The switch name and its two values are the settled ones.
'@details
'The R package writes this exact name on the copy it works with. A rename here
'alone leaves the R side writing a name this class reads nothing from.
'@TestMethod("Messenger")
Public Sub TestTheSwitchNameIsTheSettledOne()
    CustomTestSetTitles Assert, "Messenger", "TestTheSwitchNameIsTheSettledOne"
    On Error GoTo TestFail

    Assert.AreEqual EXPECTED_SWITCH_NAME, Messenger.SwitchName, _
                    "The switch is stored under __OBT__SILENT_OPERATIONS__"
    Assert.AreEqual "Yes", Messenger.SwitchOnValue, "Yes means silent"
    Assert.AreEqual "No", Messenger.SwitchOffValue, "No means the boxes behave as they always have"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSwitchNameIsTheSettledOne", Err.Number, Err.Description
End Sub

'@sub-title A workbook holding no such name reads as No.
'@details
'This is what every linelist built before the switch existed answers, and it
'is why nothing about those workbooks changes.
'@TestMethod("Messenger")
Public Sub TestAMissingSwitchReadsAsNo()
    CustomTestSetTitles Assert, "Messenger", "TestAMissingSwitchReadsAsNo"
    On Error GoTo TestFail

    Assert.IsFalse Messenger.ReadStoredSwitch(ThisWorkbook), _
                   "A workbook with no switch on it shows its boxes"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAMissingSwitchReadsAsNo", Err.Number, Err.Description
End Sub

'@sub-title A switch holding Yes reads as silent.
'@TestMethod("Messenger")
Public Sub TestAStoredYesReadsAsSilent()
    CustomTestSetTitles Assert, "Messenger", "TestAStoredYesReadsAsSilent"
    On Error GoTo TestFail

    SetStoredSwitch "Yes"

    Assert.IsTrue Messenger.ReadStoredSwitch(ThisWorkbook), _
                  "A workbook whose switch holds Yes is silent"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAStoredYesReadsAsSilent", Err.Number, Err.Description
End Sub

'@sub-title A switch holding No reads as loud.
'@TestMethod("Messenger")
Public Sub TestAStoredNoReadsAsLoud()
    CustomTestSetTitles Assert, "Messenger", "TestAStoredNoReadsAsLoud"
    On Error GoTo TestFail

    SetStoredSwitch "No"

    Assert.IsFalse Messenger.ReadStoredSwitch(ThisWorkbook), _
                   "A workbook whose switch holds No shows its boxes"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAStoredNoReadsAsLoud", Err.Number, Err.Description
End Sub

'@sub-title The stored value is matched without regard to case.
'@details
'The R side writes the string, and a workbook edited by hand may hold "YES".
'@TestMethod("Messenger")
Public Sub TestTheStoredValueIsMatchedWithoutCase()
    CustomTestSetTitles Assert, "Messenger", "TestTheStoredValueIsMatchedWithoutCase"
    On Error GoTo TestFail

    SetStoredSwitch "YES"
    Assert.IsTrue Messenger.ReadStoredSwitch(ThisWorkbook), "YES reads as silent"

    SetStoredSwitch "yes"
    Assert.IsTrue Messenger.ReadStoredSwitch(ThisWorkbook), "yes reads as silent"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheStoredValueIsMatchedWithoutCase", _
                         Err.Number, Err.Description
End Sub

'@sub-title Reading the switch without a workbook answers No.
'@TestMethod("Messenger")
Public Sub TestReadingTheSwitchWithoutAWorkbookAnswersNo()
    CustomTestSetTitles Assert, "Messenger", "TestReadingTheSwitchWithoutAWorkbookAnswersNo"
    On Error GoTo TestFail

    Assert.IsFalse Messenger.ReadStoredSwitch(Nothing), _
                   "A read with no workbook answers as loud"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestReadingTheSwitchWithoutAWorkbookAnswersNo", _
                         Err.Number, Err.Description
End Sub

'@sub-title Reading the switch changes nothing about the messenger.
'@details
'The read answers a value and the open path decides what to do with it, so a
'workbook whose switch says Yes is still loud until something arms.
'@TestMethod("Messenger")
Public Sub TestReadingTheSwitchLeavesTheMessengerAlone()
    CustomTestSetTitles Assert, "Messenger", "TestReadingTheSwitchLeavesTheMessengerAlone"
    On Error GoTo TestFail

    SetStoredSwitch "Yes"

    Assert.IsTrue Messenger.ReadStoredSwitch(ThisWorkbook), "The switch reads as silent"
    Assert.IsFalse Messenger.Armed, "Reading the switch does not arm the messenger"

    Messenger.Arm ThisWorkbook
    Assert.IsTrue Messenger.ReadStoredSwitch(ThisWorkbook), _
                  "The read answers the same value while armed"
    Assert.IsTrue Messenger.Armed, "Reading the switch does not disarm the messenger"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestReadingTheSwitchLeavesTheMessengerAlone", _
                         Err.Number, Err.Description
End Sub
