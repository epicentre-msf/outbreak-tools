Attribute VB_Name = "TestLLExporter"
Attribute VB_Description = "Unit tests for LLExporter"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Unit tests for LLExporter")

Option Explicit

'@description
'Drives LLExporter, the class that writes one standalone workbook per export.
'The four modes each build a real file, so what this module reaches is the two
'reads every mode depends on: the selection an export makes before it writes
'anything, and the epiweek value that goes into the file's Metadata sheet.
'
'ONE WORKBOOK PER MODULE
'-------------------------------------------------------------------------------
'The fixture workbook carries a Dictionary worksheet and an Exports worksheet
'and is built once in ModuleInitialize. Every test here reads, so nothing has
'to be rebuilt between them.
'
'THE FIXTURE FAILS QUIETLY AND EVERY TEST SAYS SO
'-------------------------------------------------------------------------------
'An error escaping ModuleInitialize reaches the VBE as a modal dialog and the
'whole headless run comes back with no results file. The setup captures its
'error into two module fields and FixtureReady reports it as each test's own
'failure.
'
'WHAT THE FIXTURE CARRIES
'-------------------------------------------------------------------------------
'The shared dictionary fixture, with one main label rewritten to the geo p-code
'tag so the p-code branch has a row to find:
'
'  hid_beg_v1    main label "geo_pcode_adm1", so it is a p-code and nothing else
'  hid_end_v1    personal identifier "yes", and it is the only such row
'
'The Exports worksheet holds three rows, one per case:
'
'  export 1      p-codes off, identifiers off
'  export 2      both on
'  export 3      p-codes on, identifiers off
'@depends LLExporter, LLdictionary, LLExport, HiddenNames, CustomTest

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private SetupError As Long
Private SetupMessage As String

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "LLExporter"
Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const EXPORTS_SHEET As String = "Exports"

Private Const EPIWEEK_NAME As String = "RNG_EpiWeekStart"
Private Const EPIWEEK_VALUE As String = "3"

'The two fixture rows the exclusion has to find
Private Const PCODE_VARIABLE As String = "hid_beg_v1"
Private Const PCODE_TAG As String = "geo_pcode_adm1"
Private Const IDENTIFIER_VARIABLE As String = "hid_end_v1"


'@section Lifecycle
'===============================================================================

'@sub-title Build the assertion harness and the fixture workbook.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()

    Dim store As HiddenNames

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLExporter"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set FixtureWorkbook = NewWorkbook()
        DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, FixtureWorkbook
        TagOnePcodeRow
        BuildExportsSheet

        Set store = HiddenNames.Create(FixtureWorkbook)
        store.EnsureName EPIWEEK_NAME, EPIWEEK_VALUE, HiddenNameTypeString

        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print the results and drop the fixture workbook.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()

    'The fixture workbook goes first. PrintResults writes validation and named
    'ranges onto the output sheet of ThisWorkbook, and those writes want
    'ThisWorkbook to be the active book.
    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set FixtureWorkbook = Nothing

    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Put the application into its test state.
'@details
'There is no BeginTest call here on purpose. BeginTest opens the checking with
'whatever titles are pending, and the Flush in TestCleanup has just reset those
'to the default, so every result of the module would be filed under the default
'label.
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


'@section Factory
'===============================================================================

'@sub-title Verify Create returns a valid LLExporter instance for a real workbook.
'@TestMethod("LLExporter")
Public Sub FactoryCreatesWithWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "FactoryCreatesWithWorkbook"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Set exporter = LLExporter.Create(ThisWorkbook)
    Assert.IsNotNothing exporter, "Factory should return a valid object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryCreatesWithWorkbook", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises an error when the workbook argument is Nothing.
'@TestMethod("LLExporter")
Public Sub FactoryRejectsNothingWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "FactoryRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    On Error Resume Next
    Set exporter = LLExporter.Create(Nothing)
    Assert.IsTrue Err.Number <> 0, "Factory should raise error for Nothing workbook"
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingWorkbook", Err.Number, Err.Description
End Sub

'@sub-title Verify LastExportPassword defaults to an empty string on a fresh instance.
'@TestMethod("LLExporter")
Public Sub LastExportPasswordEmptyByDefault()
    CustomTestSetTitles Assert, TESTMODULE, "LastExportPasswordEmptyByDefault"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Set exporter = LLExporter.Create(ThisWorkbook)
    Assert.AreEqual vbNullString, exporter.LastExportPassword, _
                    "Password should be empty before any export"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "LastExportPasswordEmptyByDefault", Err.Number, Err.Description
End Sub

'@sub-title Verify CloseAll completes without error when no output workbooks exist.
'@TestMethod("LLExporter")
Public Sub CloseAllDoesNotError()
    CustomTestSetTitles Assert, TESTMODULE, "CloseAllDoesNotError"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Set exporter = LLExporter.Create(ThisWorkbook)
    exporter.CloseAll
    Assert.IsTrue True, "CloseAll should not raise errors when no workbook is open"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "CloseAllDoesNotError", Err.Number, Err.Description
End Sub


'@section The selection an export makes
'===============================================================================

'@sub-title An export with both flags off drops the p-code and the identifier.
'@TestMethod("LLExporter")
Public Sub TestBothFlagsOffDropThePcodeAndTheIdentifier()
    CustomTestSetTitles Assert, TESTMODULE, "TestBothFlagsOffDropThePcodeAndTheIdentifier"
    If Not FixtureReady("TestBothFlagsOffDropThePcodeAndTheIdentifier") Then Exit Sub
    On Error GoTo TestFail

    Dim excluded As BetterArray

    Set excluded = FixtureExporter().ExcludedVariablesFor(1)

    Assert.IsTrue excluded.Length >= 2, _
                  "Both flags off should drop at least the two fixture rows"
    Assert.IsTrue excluded.Includes(PCODE_VARIABLE), _
                  "A variable whose main label is a geo p-code tag is dropped"
    Assert.IsTrue excluded.Includes(IDENTIFIER_VARIABLE), _
                  "A variable flagged as a personal identifier is dropped"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBothFlagsOffDropThePcodeAndTheIdentifier", Err.Number, Err.Description
End Sub

'@sub-title An export that keeps everything drops nothing.
'@TestMethod("LLExporter")
Public Sub TestAnExportThatKeepsEverythingDropsNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnExportThatKeepsEverythingDropsNothing"
    If Not FixtureReady("TestAnExportThatKeepsEverythingDropsNothing") Then Exit Sub
    On Error GoTo TestFail

    Dim excluded As BetterArray

    Set excluded = FixtureExporter().ExcludedVariablesFor(2)

    Assert.AreEqual CLng(0), excluded.Length, _
                    "Both flags on should drop nothing at all"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnExportThatKeepsEverythingDropsNothing", Err.Number, Err.Description
End Sub

'@sub-title The two flags are read one at a time.
'@details
'Export 3 keeps its p-codes and drops its identifiers, so the p-code row has to
'survive and the identifier row has to go. A single flag standing in for both
'would fail this one.
'@TestMethod("LLExporter")
Public Sub TestTheIdentifierFlagActsOnItsOwn()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheIdentifierFlagActsOnItsOwn"
    If Not FixtureReady("TestTheIdentifierFlagActsOnItsOwn") Then Exit Sub
    On Error GoTo TestFail

    Dim excluded As BetterArray

    Set excluded = FixtureExporter().ExcludedVariablesFor(3)

    Assert.IsTrue excluded.Includes(IDENTIFIER_VARIABLE), _
                  "The identifier is dropped when its own flag is off"
    Assert.IsFalse excluded.Includes(PCODE_VARIABLE), _
                   "The p-code stays when the p-code flag is on"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheIdentifierFlagActsOnItsOwn", Err.Number, Err.Description
End Sub

'@sub-title Reading the selection leaves the linelist dictionary as it was.
'@details
'This is the whole point of the change. The exclusion used to be done by
'clearing cells of the export column on the linelist's own Dictionary
'worksheet, one way and for good, so a user who turned the flag back on got a
'file that was still missing those variables and a dictionary that no longer
'matched the setup it came from.
'@TestMethod("LLExporter")
Public Sub TestTheSourceDictionaryIsNeverWrittenTo()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheSourceDictionaryIsNeverWrittenTo"
    If Not FixtureReady("TestTheSourceDictionaryIsNeverWrittenTo") Then Exit Sub
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim excluded As BetterArray
    Dim before As Variant
    Dim after As Variant
    Dim counter As Long
    Dim differences As Long

    before = ExportColumnValues()

    Set exporter = FixtureExporter()
    Set excluded = exporter.ExcludedVariablesFor(1)

    after = ExportColumnValues()

    Assert.IsTrue excluded.Length > 0, _
                  "The export under test does drop something"

    Assert.AreEqual UBound(before), UBound(after), _
                    "The export column should keep every row it had"

    For counter = LBound(before) To UBound(before)
        If CStr(before(counter)) <> CStr(after(counter)) Then
            differences = differences + 1
        End If
    Next counter

    Assert.AreEqual CLng(0), differences, _
                    "Reading the selection must not change one cell of the export column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSourceDictionaryIsNeverWrittenTo", Err.Number, Err.Description
End Sub


'@section The epiweek value
'===============================================================================

'@sub-title The epiweek value comes back from the hidden name store.
'@TestMethod("LLExporter")
Public Sub TestEpiWeekStartReadsTheHiddenNameStore()
    CustomTestSetTitles Assert, TESTMODULE, "TestEpiWeekStartReadsTheHiddenNameStore"
    If Not FixtureReady("TestEpiWeekStartReadsTheHiddenNameStore") Then Exit Sub
    On Error GoTo TestFail

    Assert.AreEqual EPIWEEK_VALUE, FixtureExporter().EpiWeekStart(), _
                    "The stored epiweek start is what the metadata tag carries"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEpiWeekStartReadsTheHiddenNameStore", Err.Number, Err.Description
End Sub

'@sub-title The hidden name has no range behind it, which is why the old read failed.
'@details
'HiddenNames stores a string as a quoted literal in the name definition. The
'name is there and Names(...) finds it; asking that name for a range raises
'1004, and the export code used to do exactly that under a swallow, so every
'file it wrote carried a blank epiweek tag.
'@TestMethod("LLExporter")
Public Sub TestTheEpiWeekNameHasNoRangeBehindIt()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheEpiWeekNameHasNoRangeBehindIt"
    If Not FixtureReady("TestTheEpiWeekNameHasNoRangeBehindIt") Then Exit Sub
    On Error GoTo TestFail

    Dim rangeErrNumber As Long
    Dim storedName As Excel.Name
    Dim rangeAddress As String

    Set storedName = Nothing
    On Error Resume Next
    Set storedName = FixtureWorkbook.Names(EPIWEEK_NAME)
    On Error GoTo 0

    Assert.IsNotNothing storedName, "The workbook does carry the name"

    On Error Resume Next
    rangeAddress = storedName.RefersToRange.Address
    rangeErrNumber = Err.Number
    Err.Clear
    On Error GoTo 0

    Assert.IsTrue rangeErrNumber <> 0, _
                  "Asking a string-valued hidden name for a range raises"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheEpiWeekNameHasNoRangeBehindIt", Err.Number, Err.Description
End Sub

'@sub-title A workbook with no such name answers an empty string.
'@TestMethod("LLExporter")
Public Sub TestEpiWeekStartIsEmptyWhenTheNameIsAbsent()
    CustomTestSetTitles Assert, TESTMODULE, "TestEpiWeekStartIsEmptyWhenTheNameIsAbsent"
    On Error GoTo TestFail

    Dim plainBook As Workbook

    Set plainBook = NewWorkbook()
    Assert.AreEqual vbNullString, LLExporter.Create(plainBook).EpiWeekStart(), _
                    "A workbook that never stored the value answers an empty string"

    DeleteWorkbook plainBook

    Exit Sub
TestFail:
    On Error Resume Next
    If Not plainBook Is Nothing Then DeleteWorkbook plainBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestEpiWeekStartIsEmptyWhenTheNameIsAbsent", Err.Number, Err.Description
End Sub


'@section Fixture helpers
'===============================================================================

'@fun-title Report a fixture that could not be built as this test's failure.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is usable.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If SetupError = 0 And Not FixtureWorkbook Is Nothing Then
        FixtureReady = True
        Exit Function
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The fixture could not be built - " & SetupMessage
End Function

'@fun-title An exporter over the fixture workbook.
'@return LLExporter. A fresh instance, which costs nothing to build.
Private Function FixtureExporter() As LLExporter
    Set FixtureExporter = LLExporter.Create(FixtureWorkbook)
End Function

'@sub-title Give one fixture row a geo p-code main label.
'@details
'No row of the shared dictionary fixture carries one, and the p-code branch
'matches on the main label. hid_beg_v1 is chosen because it is not a personal
'identifier, so the two branches keep separate rows to find.
Private Sub TagOnePcodeRow()
    Dim sh As Worksheet
    Dim labelColumn As Long

    Set sh = FixtureWorkbook.Worksheets(DICTIONARY_SHEET)
    labelColumn = DictionaryTestFixture.DictionaryHeaderIndex("Main Label") + 1
    sh.Cells(2, labelColumn).Value = PCODE_TAG
End Sub

'@sub-title Write the three-row Exports worksheet the tests read.
Private Sub BuildExportsSheet()
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(EXPORTS_SHEET, FixtureWorkbook, clearSheet:=True)

    WriteRow sh.Cells(1, 1), "export number", "status", "label button", _
                             "file format", "file name", "password", _
                             "include personal identifiers", "include p-codes", _
                             "header format", "export metadata sheets", _
                             "export analyses sheets"

    WriteRow sh.Cells(2, 1), 1, "active", "both off", "xlsx", "one", "no", _
                             "no", "no", "default", "no", "no"
    WriteRow sh.Cells(3, 1), 2, "active", "both on", "xlsx", "two", "no", _
                             "yes", "yes", "default", "no", "no"
    WriteRow sh.Cells(4, 1), 3, "active", "identifiers off", "xlsx", "three", "no", _
                             "no", "yes", "default", "no", "no"
End Sub

'@fun-title Read the export 1 column of the fixture dictionary.
'@return Variant. A 1-based array of the cell values.
Private Function ExportColumnValues() As Variant
    Dim sh As Worksheet
    Dim column As Long
    Dim lastRow As Long
    Dim values() As Variant
    Dim counter As Long

    Set sh = FixtureWorkbook.Worksheets(DICTIONARY_SHEET)
    column = DictionaryTestFixture.DictionaryHeaderIndex("Export 1") + 1
    lastRow = DictionaryTestFixture.DictionaryFixtureRowCount() + 1

    ReDim values(2 To lastRow)
    For counter = 2 To lastRow
        values(counter) = sh.Cells(counter, column).Value
    Next counter

    ExportColumnValues = values
End Function
