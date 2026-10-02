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
'@depends CustomLinelistFunctions, DropdownLists, BetterArray, CustomTest,
'  TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "CustomLinelistFunctions"

' The two workbook-level names the epiweek arithmetic reads.
Private Const WEEK_START_NAME As String = "RNG_EpiWeekStart"
Private Const WEEK_TAG_NAME As String = "RNG_Week"

' The shared time unit dropdown, and the sheet the suite parks it on.
Private Const TIMEUNIT_LIST_NAME As String = "dropdown_time_unit"
Private Const TIMEUNIT_SHEET As String = "CLFTimeUnit"

' The sheet holding the lookup table the VALUE_OF tests read.
Private Const VALUEOF_SHEET As String = "CLFValueOf"

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


'@section Aggregation
'===============================================================================
'GetAgg is Private, so these drive it through FindLastDay, which is the function
'that turns the chosen label into the last day of a period.

'@sub-title The shared time unit list decides the aggregation, on any sheet.
'@details
'The five labels are one workbook-level dropdown that Linelist.Prepare adds
'through DropdownLists, and the reader names it and nothing else. The reader
'used to test the tag of ActiveSheet first and answer "week" for every sheet
'that was not an analysis sheet, so a recalculation fired while the user sat
'anywhere else computed a monthly table as weeks. No test in this module opens
'an analysis sheet, so the answers below are the proof that the tag is gone.
'
'It also proves the read itself. The name refers to a structured reference over
'a ListObject column, and RefersToRange is what has to resolve it.
'@TestMethod("CustomLinelistFunctions")
Public Sub TestTheTimeUnitListDecidesTheAggregation()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheTimeUnitListDecidesTheAggregation"
    On Error GoTo TestFail

    Dim someDay As Long

    DropWorkbookName WEEK_START_NAME
    AddTimeUnitList

    someDay = CLng(DateSerial(2026, 3, 10))

    Assert.AreEqual CLng(DateSerial(2026, 3, 10)), FindLastDay("Day", someDay), _
                    "A daily aggregation ends on the day itself"
    Assert.AreEqual CLng(DateSerial(2026, 3, 31)), FindLastDay("Month", someDay), _
                    "A monthly aggregation ends on the last day of the month"
    Assert.AreEqual CLng(DateSerial(2026, 3, 31)), FindLastDay("Quarter", someDay), _
                    "A quarterly aggregation ends on the last day of the quarter"
    Assert.AreEqual CLng(DateSerial(2026, 12, 31)), FindLastDay("Year", someDay), _
                    "A yearly aggregation ends on the last day of the year"

    DropTimeUnitList
    Exit Sub
TestFail:
    DropTimeUnitList
    CustomTestLogFailure Assert, "TestTheTimeUnitListDecidesTheAggregation", _
                         Err.Number, Err.Description
End Sub

'@sub-title A label the list does not carry falls back to the week.
'@details
'The dropdown holds a user to the five labels, so a label outside them means a
'workbook built by an older version or a cell written by hand. The same answer
'covers a workbook whose list is missing, which is what a read that raises
'leaves behind.
'@TestMethod("CustomLinelistFunctions")
Public Sub TestAnUnknownTimeUnitFallsBackToTheWeek()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnUnknownTimeUnitFallsBackToTheWeek"
    On Error GoTo TestFail

    Dim someDay As Long
    Dim weekEnd As Long

    ' 10 March 2026 is a Tuesday, so the Monday week holding it ends on the 15th.
    someDay = CLng(DateSerial(2026, 3, 10))
    weekEnd = CLng(DateSerial(2026, 3, 15))

    DropWorkbookName WEEK_START_NAME
    AddTimeUnitList

    Assert.AreEqual weekEnd, FindLastDay("Fortnight", someDay), _
                    "A label outside the list is aggregated as a week"

    DropTimeUnitList

    Assert.AreEqual weekEnd, FindLastDay("Month", someDay), _
                    "and a missing list is aggregated as a week as well"

    Exit Sub
TestFail:
    DropTimeUnitList
    CustomTestLogFailure Assert, "TestAnUnknownTimeUnitFallsBackToTheWeek", _
                         Err.Number, Err.Description
End Sub


'@section VALUE_OF
'===============================================================================

'@sub-title A matched value keeps its own type, and a blank one answers nothing.
'@details
'The value column of a lookup table holds dates, numbers and blanks side by
'side. A blank cell used to come back as an Empty variant, which the formula
'cell shows as 0, so every unfilled row of the lookup table read as a zero. A
'stored 0 still answers 0, and a date stays a date.
'@TestMethod("CustomLinelistFunctions")
Public Sub TestValueOfKeepsTheTypeAndAnswersNothingOnBlank()
    CustomTestSetTitles Assert, TESTMODULE, "TestValueOfKeepsTheTypeAndAnswersNothingOnBlank"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim keyRng As Range
    Dim result As Variant

    Set sh = BuildValueOfTable()
    Set keyRng = sh.Range("F1")

    keyRng.Value = "with_date"
    result = VALUE_OF(keyRng, VALUEOF_SHEET, 1, 2)
    Assert.IsTrue IsDate(result), "A stored date comes back as a date"
    Assert.AreEqual CLng(DateSerial(2026, 3, 10)), CLng(CDate(result)), _
                    "and it is the stored one"

    keyRng.Value = "with_zero"
    result = VALUE_OF(keyRng, VALUEOF_SHEET, 1, 2)
    Assert.IsTrue IsNumeric(result), "A stored 0 comes back as a number"
    Assert.AreEqual 0#, CDbl(result), "and it is 0"

    keyRng.Value = "with_blank"
    result = VALUE_OF(keyRng, VALUEOF_SHEET, 1, 2)
    Assert.AreEqual vbNullString, result, _
                    "A blank value cell answers an empty string. It used to " & _
                    "come back Empty, which the formula cell shows as 0"

    keyRng.Value = "no_such_key"
    result = VALUE_OF(keyRng, VALUEOF_SHEET, 1, 2)
    Assert.AreEqual vbNullString, result, _
                    "A key with no match answers an empty string"

    DropValueOfTable
    Exit Sub
TestFail:
    DropValueOfTable
    CustomTestLogFailure Assert, "TestValueOfKeepsTheTypeAndAnswersNothingOnBlank", _
                         Err.Number, Err.Description
End Sub

'@sub-title Build the lookup table the VALUE_OF tests read.
'@details
'One ListObject on its own sheet: a key column and a value column holding a
'date, a 0 and a blank. The cache slot of the sheet is dropped first, so a
'previous test run cannot answer for this one.
'@return Worksheet. The sheet holding the table.
Private Function BuildValueOfTable() As Worksheet
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(VALUEOF_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    sh.Range("A1").Value = "key"
    sh.Range("B1").Value = "value"
    sh.Range("A2").Value = "with_date"
    sh.Range("B2").Value = DateSerial(2026, 3, 10)
    sh.Range("A3").Value = "with_zero"
    sh.Range("B3").Value = 0
    sh.Range("A4").Value = "with_blank"

    sh.ListObjects.Add(xlSrcRange, sh.Range("A1:B4"), , xlYes).Name = "CLFValueOfTable"

    CustomLinelistFunctions.ResetValueOfCache VALUEOF_SHEET

    Set BuildValueOfTable = sh
End Function

'@sub-title Take the lookup table and its worksheet away again.
Private Sub DropValueOfTable()
    CustomLinelistFunctions.ResetValueOfCache VALUEOF_SHEET
    DeleteWorksheet VALUEOF_SHEET
End Sub


'@section Aggregation helpers
'===============================================================================

'@sub-title Add the shared time unit dropdown the way a linelist carries it.
Private Sub AddTimeUnitList()
    Dim drop As DropdownLists
    Dim listValues As BetterArray

    DropTimeUnitList

    Set drop = DropdownLists.Create( _
        EnsureWorksheet(TIMEUNIT_SHEET, clearSheet:=True, visibility:=xlSheetHidden), _
        "dropdown_")

    Set listValues = New BetterArray
    listValues.LowerBound = 1
    listValues.Push "Day", "Week", "Month", "Quarter", "Year"

    drop.Add listValues, "time_unit"
End Sub

'@sub-title Take the dropdown and its worksheet away again.
Private Sub DropTimeUnitList()
    DropWorkbookName TIMEUNIT_LIST_NAME
    DeleteWorksheet TIMEUNIT_SHEET
End Sub
