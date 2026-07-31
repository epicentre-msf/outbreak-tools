Attribute VB_Name = "TestCustomLinelistFunctions"
Attribute VB_Description = "Tests for the linelist worksheet functions"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for the linelist worksheet functions")

'@description
'Drives the worksheet functions a generated linelist carries. Every one of them
'is called from a cell formula, so a raise out of one is a #VALUE! in front of
'the user and nothing in a log.
'
'WHY THIS MODULE EXISTS
'-------------------------------------------------------------------------------
'Issue 340 reported EPIWEEK answering 53 for 29 to 31 December 2025. The
'implementation that behaved that way was replaced in February 2026 and nothing
'ever proved the replacement, because src/modules/linelist/ had no registry row
'at all. A function can be correct for five months with an open bug report
'against it when no test ever calls it.
'
'THE ANSWERS BELOW ARE ISO 8601
'-------------------------------------------------------------------------------
'Week 1 of a year is the week holding 1 January when 1 January falls on day 1 to
'4 of the week, and it starts seven days later otherwise. The week starts on
'Monday here, which is the default the module falls back to.
'
'    28 Dec 2025   W52-2025
'    29 Dec 2025   W1-2026     1 Jan 2026 is a Thursday, day 4
'    31 Dec 2025   W1-2026
'    1  Jan 2026   W1-2026
'    1  Jan 2027   W53-2026    1 Jan 2027 is a Friday, day 5
'
'THE TWO HIDDEN NAMES THE FUNCTION READS
'-------------------------------------------------------------------------------
'RNG_EpiWeekStart carries the first day of the week and RNG_Week the letter the
'answer starts with. Both are user data by the time a linelist is delivered, and
'both are absent in this workbook, so the tests that need one create it and
'delete it again.
'@depends CustomLinelistFunctions, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "CustomLinelistFunctions"

' The two workbook-level names the epiweek arithmetic reads.
Private Const WEEK_START_NAME As String = "RNG_EpiWeekStart"
Private Const WEEK_TAG_NAME As String = "RNG_Week"

Private Assert As CustomTest


'@section Lifecycle
'===============================================================================

'@sub-title Prepare the assertion harness.
'@details
'This routine is Public because the harness reaches it through Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCustomLinelistFunctions"
    DropWorkbookName WEEK_START_NAME
    DropWorkbookName WEEK_TAG_NAME
End Sub

'@sub-title Print the results and drop the two names the suite may have added.
'@details
'This routine is Public because the harness reaches it through Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    DropWorkbookName WEEK_START_NAME
    DropWorkbookName WEEK_TAG_NAME

    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updates before each test.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush the results of the test that just ran.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Helpers
'===============================================================================

'@sub-title Delete one workbook-level name when it is there.
'@param nameId String. The name to remove.
Private Sub DropWorkbookName(ByVal nameId As String)
    On Error Resume Next
    ThisWorkbook.Names(nameId).Delete
    On Error GoTo 0
End Sub

'@sub-title Write one workbook-level name the way HiddenNames stores a string.
'@param nameId String. The name to write.
'@param value String. The value it carries.
Private Sub SetWorkbookName(ByVal nameId As String, ByVal value As String)
    DropWorkbookName nameId
    ThisWorkbook.Names.Add Name:=nameId, RefersTo:="=" & Chr$(34) & value & Chr$(34)
End Sub


'@section Epiweek
'===============================================================================

'@sub-title The reported dates of issue 340 answer week 1 of the next year.
'@details
'29 to 31 December 2025 belong to the week that holds 1 January 2026, and that
'day is a Thursday, so the week is week 1 of 2026. The report said 53.
'@TestMethod("CustomLinelistFunctions")
Public Sub TestTheLastDaysOf2025BelongToWeekOneOf2026()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheLastDaysOf2025BelongToWeekOneOf2026"
    On Error GoTo TestFail

    Assert.AreEqual "W1-2026", Epiweek(CLng(DateSerial(2025, 12, 29))), _
                    "29 December 2025 is in week 1 of 2026"
    Assert.AreEqual "W1-2026", Epiweek(CLng(DateSerial(2025, 12, 30))), _
                    "and so is 30 December 2025"
    Assert.AreEqual "W1-2026", Epiweek(CLng(DateSerial(2025, 12, 31))), _
                    "and so is 31 December 2025"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheLastDaysOf2025BelongToWeekOneOf2026", _
                         Err.Number, Err.Description
End Sub

'@sub-title The days either side of that boundary answer what ISO 8601 says.
'@details
'28 December 2025 is the last day of week 52 of 2025, 1 January 2026 opens week
'1 of 2026, and 1 January 2027 is a Friday, so it is still week 53 of 2026.
'@TestMethod("CustomLinelistFunctions")
Public Sub TestTheDaysAroundTheYearBoundaryFollowTheMajorityRule()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheDaysAroundTheYearBoundaryFollowTheMajorityRule"
    On Error GoTo TestFail

    Assert.AreEqual "W52-2025", Epiweek(CLng(DateSerial(2025, 12, 28))), _
                    "28 December 2025 closes week 52 of 2025"
    Assert.AreEqual "W1-2026", Epiweek(CLng(DateSerial(2026, 1, 1))), _
                    "1 January 2026 falls on day 4 of its week, so it opens week 1"
    Assert.AreEqual "W53-2026", Epiweek(CLng(DateSerial(2027, 1, 1))), _
                    "1 January 2027 falls on day 5, so its week belongs to 2026"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheDaysAroundTheYearBoundaryFollowTheMajorityRule", _
                         Err.Number, Err.Description
End Sub

'@sub-title A blank date answers nothing at all.
'@details
'A blank cell reaches the Long parameter as 0, which is 30 December 1899, and
'the arithmetic used to answer an 1899 epi-year for it. The formulas in the
'field guard the call with ISBLANK; the ones that do not used to get a date
'nobody typed.
'@TestMethod("CustomLinelistFunctions")
Public Sub TestABlankDateAnswersAnEmptyString()
    CustomTestSetTitles Assert, TESTMODULE, "TestABlankDateAnswersAnEmptyString"
    On Error GoTo TestFail

    Assert.AreEqual vbNullString, Epiweek(0), _
                    "A blank cell answers nothing rather than an 1899 epi-year"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestABlankDateAnswersAnEmptyString", _
                         Err.Number, Err.Description
End Sub

'@sub-title A week start that is not a number falls back to Monday.
'@details
'RNG_EpiWeekStart is user data in a delivered linelist. Converting it raised 13
'out of a worksheet function, which the user reads as #VALUE! with nothing
'saying why.
'@TestMethod("CustomLinelistFunctions")
Public Sub TestANonNumericWeekStartFallsBackToMonday()
    CustomTestSetTitles Assert, TESTMODULE, "TestANonNumericWeekStartFallsBackToMonday"
    On Error GoTo TestFail

    SetWorkbookName WEEK_START_NAME, "lundi"

    Assert.AreEqual "W1-2026", Epiweek(CLng(DateSerial(2025, 12, 29))), _
                    "A week start that is not a number is read as Monday " & _
                    "rather than raising out of the cell"

    DropWorkbookName WEEK_START_NAME

    Exit Sub
TestFail:
    DropWorkbookName WEEK_START_NAME
    CustomTestLogFailure Assert, "TestANonNumericWeekStartFallsBackToMonday", _
                         Err.Number, Err.Description
End Sub

'@sub-title A Sunday week start moves the boundary, and the majority rule follows it.
'@details
'The caller can override the first day of the week, and 0 is Sunday. The rule is
'the same one, read against a week that now runs Sunday to Saturday: 1 January
'2026 is a Thursday, which is day 5 of a Sunday week, so fewer than four days of
'that week fall in January and week 1 of 2026 starts seven days later, on 4
'January. Every day before it belongs to week 53 of 2025.
'
'This is the pair that says the week start is read rather than assumed. With
'Monday weeks the same 29 December 2025 answers week 1 of 2026.
'@TestMethod("CustomLinelistFunctions")
Public Sub TestASundayWeekStartMovesTheBoundary()
    CustomTestSetTitles Assert, TESTMODULE, "TestASundayWeekStartMovesTheBoundary"
    On Error GoTo TestFail

    Assert.AreEqual "W1-2026", Epiweek(CLng(DateSerial(2026, 1, 4)), 0), _
                    "With Sunday weeks, 4 January 2026 opens week 1 of 2026"
    Assert.AreEqual "W53-2025", Epiweek(CLng(DateSerial(2026, 1, 3)), 0), _
                    "and the day before it closes week 53 of 2025"
    Assert.AreEqual "W53-2025", Epiweek(CLng(DateSerial(2025, 12, 29)), 0), _
                    "so 29 December 2025 is in 2025 under Sunday weeks, where " & _
                    "Monday weeks put it in week 1 of 2026"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASundayWeekStartMovesTheBoundary", _
                         Err.Number, Err.Description
End Sub

'@sub-title The week letter comes from the translated hidden name.
'@details
'RNG_Week is what makes the answer read S1-2026 in French. An absent name falls
'back to W, which is what every other test in this module reads.
'@TestMethod("CustomLinelistFunctions")
Public Sub TestTheWeekLetterComesFromTheHiddenName()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheWeekLetterComesFromTheHiddenName"
    On Error GoTo TestFail

    SetWorkbookName WEEK_TAG_NAME, "S"

    Assert.AreEqual "S1-2026", Epiweek(CLng(DateSerial(2025, 12, 29))), _
                    "The week letter is read from RNG_Week"

    DropWorkbookName WEEK_TAG_NAME

    Assert.AreEqual "W1-2026", Epiweek(CLng(DateSerial(2025, 12, 29))), _
                    "and it falls back to W when the name is gone"

    Exit Sub
TestFail:
    DropWorkbookName WEEK_TAG_NAME
    CustomTestLogFailure Assert, "TestTheWeekLetterComesFromTheHiddenName", _
                         Err.Number, Err.Description
End Sub
