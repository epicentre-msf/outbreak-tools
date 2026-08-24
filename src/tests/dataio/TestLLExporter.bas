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
'
'THE OPEN-FROM-FILE SEAM
'-------------------------------------------------------------------------------
'CreateFromFile backs the other-linelist export of F_ExportMig. The tests here
'save small workbooks beside the driver and open them back: a missing path and
'a workbook without the linelist sheets are refused, a file the factory opened
'belongs to it and CloseAll closes it, and a workbook already open in the
'session is used as it stands and stays open.
'@depends LLExporter, LLdictionary, LLExport, HiddenNames, ChoiceKeys, CustomTest, Checking

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

'The three sheets a linelist shell needs on top of Dictionary and Exports
Private Const PASS_SHEET As String = "__pass"
Private Const GEO_SHEET As String = "Geo"
Private Const TEMP_SHEET As String = "__temp"
'The name a linelist generated before the internal-sheet rename gave it
Private Const TEMP_SHEET_OLD As String = "temp__"

'Where the open-from-file tests put their saved workbooks
Private Const FILES_FOLDER As String = "ExporterFiles"

'The two fixture rows the exclusion has to find
Private Const PCODE_VARIABLE As String = "hid_beg_v1"
Private Const PCODE_TAG As String = "geo_pcode_adm1"
Private Const IDENTIFIER_VARIABLE As String = "hid_end_v1"

'The fixture row a geo control puts on an admin level. hid_beg_v1 carries the
'main label geo_pcode_adm1 and sits on level 1, so the two are dropped by
'different levels and a test can keep one of them.
Private Const GEO_LEVEL3_VARIABLE As String = "geo_h2"
Private Const GEO_LEVEL3_CONTROL As String = "geo3"


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
        TagOneGeoLevelRow
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

    'Four tests write workbooks into FILES_FOLDER through BuildTempFolder, which
    'makes it with MkDir, and each one kills its own file. The folder itself is
    'this module's to remove, and here is the one place every test has finished.
    'The path is built by hand because BuildTempFolder creates the folder it
    'answers. RmDir fails quietly when the folder is absent or holds a file.
    On Error Resume Next
        RmDir ThisWorkbook.Path & Application.PathSeparator & FILES_FOLDER
    On Error GoTo 0

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
'@details
'Export 2 keeps both flags on and leaves its admin levels cell empty, so all
'three rules stay quiet. The empty cell is the one worth naming: the p-code row
'and the geo3 row of the fixture both sit on a level, and a reader that took an
'empty cell for "no level at all" would drop them here.
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

'@sub-title An export naming one admin level drops the geo rows of the others.
'@details
'Export 4 asks for admin1 and keeps both flags on, so the level is the only
'rule that can drop anything. geo_h2 sits on level 3 and goes; hid_beg_v1
'carries the main label geo_pcode_adm1, sits on level 1 and stays.
'@TestMethod("LLExporter")
Public Sub TestOneAdminLevelDropsTheGeoRowsOfTheOthers()
    CustomTestSetTitles Assert, TESTMODULE, "TestOneAdminLevelDropsTheGeoRowsOfTheOthers"
    If Not FixtureReady("TestOneAdminLevelDropsTheGeoRowsOfTheOthers") Then Exit Sub
    On Error GoTo TestFail

    Dim excluded As BetterArray

    Set excluded = FixtureExporter().ExcludedVariablesFor(4)

    Assert.IsTrue excluded.Includes(GEO_LEVEL3_VARIABLE), _
                  "A geo3 row is dropped by an export that asks for admin1"
    Assert.IsFalse excluded.Includes(PCODE_VARIABLE), _
                   "The level the export asks for keeps its rows"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestOneAdminLevelDropsTheGeoRowsOfTheOthers", Err.Number, Err.Description
End Sub

'@sub-title The level named is the level kept, whichever one it is.
'@details
'Export 5 asks for admin3, so the two rows of export 4 change places. A rule
'that always kept level 1 would pass that test and fail this one.
'@TestMethod("LLExporter")
Public Sub TestTheLevelNamedIsTheLevelKept()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheLevelNamedIsTheLevelKept"
    If Not FixtureReady("TestTheLevelNamedIsTheLevelKept") Then Exit Sub
    On Error GoTo TestFail

    Dim excluded As BetterArray

    Set excluded = FixtureExporter().ExcludedVariablesFor(5)

    Assert.IsTrue excluded.Includes(PCODE_VARIABLE), _
                  "A level 1 row is dropped by an export that asks for admin3"
    Assert.IsFalse excluded.Includes(GEO_LEVEL3_VARIABLE), _
                   "The geo3 row stays when admin3 is the level asked for"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheLevelNamedIsTheLevelKept", Err.Number, Err.Description
End Sub

'@sub-title A cell the reader cannot make sense of keeps every level.
'@details
'Export 6 holds `admin1+adm2`, and `adm2` names no level. The whole cell is
'left alone rather than half honoured, so a typo costs a user no geo column.
'SetupErrors reports that same cell on the setup side.
'@TestMethod("LLExporter")
Public Sub TestAnAdminLevelsCellThatCannotBeReadKeepsEveryLevel()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnAdminLevelsCellThatCannotBeReadKeepsEveryLevel"
    If Not FixtureReady("TestAnAdminLevelsCellThatCannotBeReadKeepsEveryLevel") Then Exit Sub
    On Error GoTo TestFail

    Dim excluded As BetterArray

    Set excluded = FixtureExporter().ExcludedVariablesFor(6)

    Assert.AreEqual CLng(0), excluded.Length, _
                    "A cell holding a chunk that names no level drops nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnAdminLevelsCellThatCannotBeReadKeepsEveryLevel", Err.Number, Err.Description
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


'@section The options line a file carries about itself
'===============================================================================

'@sub-title Every option of the Exports row reaches the line.
'@details
'The Exports worksheet stays in the linelist, so a file has no way of saying how
'it was written unless the export writes it down. The line is what the import
'side splits back apart.
'@TestMethod("LLExporter")
Public Sub TestTheOptionsLineCarriesTheExportRow()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheOptionsLineCarriesTheExportRow"
    If Not FixtureReady("TestTheOptionsLineCarriesTheExportRow") Then Exit Sub
    On Error GoTo TestFail

    Dim options As String

    options = FixtureExporter().ExportOptionsFor(2)

    Assert.IsTrue InStr(1, options, "include personal identifiers: yes") > 0, _
                  "The identifiers flag of export 2 reaches the line"
    Assert.IsTrue InStr(1, options, "include p-codes: yes") > 0, _
                  "And the p-codes flag"
    Assert.IsTrue InStr(1, options, "file format: xlsx") > 0, _
                  "And the file format"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheOptionsLineCarriesTheExportRow", Err.Number, Err.Description
End Sub

'@sub-title The admin levels reach the line, and an empty cell reads as `all`.
'@details
'Every other option left empty is written `no`, and `no` on this one would read
'as no level at all. A file written with the cell empty carries every level.
'@TestMethod("LLExporter")
Public Sub TestTheOptionsLineNamesTheAdminLevels()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheOptionsLineNamesTheAdminLevels"
    If Not FixtureReady("TestTheOptionsLineNamesTheAdminLevels") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue InStr(1, FixtureExporter().ExportOptionsFor(4), "admin levels: admin1") > 0, _
                  "The levels of export 4 reach the line"
    Assert.IsTrue InStr(1, FixtureExporter().ExportOptionsFor(2), "admin levels: all") > 0, _
                  "An empty cell is written as all"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheOptionsLineNamesTheAdminLevels", Err.Number, Err.Description
End Sub

'@sub-title The pairs are joined with a pipe, so the import can split them.
'@TestMethod("LLExporter")
Public Sub TestTheOptionsLineSplitsOnThePipe()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheOptionsLineSplitsOnThePipe"
    If Not FixtureReady("TestTheOptionsLineSplitsOnThePipe") Then Exit Sub
    On Error GoTo TestFail

    Dim pairs As Variant

    pairs = Split(FixtureExporter().ExportOptionsFor(1), "|")

    Assert.IsTrue UBound(pairs) > 5, _
                  "The line splits into one pair per option"
    Assert.IsTrue InStr(1, CStr(pairs(0)), ":") > 0, _
                  "And each pair carries its own colon"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheOptionsLineSplitsOnThePipe", Err.Number, Err.Description
End Sub

'@sub-title A default header format says the file carries its variable names.
'@details
'The fixture export rows use "default", which is anything other than the two
'label formats, so the file holds its variable names on row 1 and reads back.
'@TestMethod("LLExporter")
Public Sub TestADefaultHeaderFormatKeepsTheVariableNames()
    CustomTestSetTitles Assert, TESTMODULE, "TestADefaultHeaderFormatKeepsTheVariableNames"
    If Not FixtureReady("TestADefaultHeaderFormatKeepsTheVariableNames") Then Exit Sub
    On Error GoTo TestFail

    Dim options As String

    options = FixtureExporter().ExportOptionsFor(1)

    Assert.IsTrue InStr(1, options, "variable names: yes") > 0, _
                  "The file carries its variable names"
    Assert.IsTrue InStr(1, options, "header row: 1") > 0, _
                  "On the first row"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestADefaultHeaderFormatKeepsTheVariableNames", Err.Number, Err.Description
End Sub

'@sub-title A label header format says the file carries no variable name.
'@details
'The one case that cannot be read back at all. Every column of such a file used
'to land in "not imported" with nothing said, because the import took row 1 as
'the header row whatever the file was.
'@TestMethod("LLExporter")
Public Sub TestALabelHeaderFormatSaysThereAreNoNames()
    CustomTestSetTitles Assert, TESTMODULE, "TestALabelHeaderFormatSaysThereAreNoNames"
    If Not FixtureReady("TestALabelHeaderFormatSaysThereAreNoNames") Then Exit Sub
    On Error GoTo TestFail

    Dim options As String

    SetHeaderFormat 1, "variables labels"
    options = FixtureExporter().ExportOptionsFor(1)
    SetHeaderFormat 1, "default"

    Assert.IsTrue InStr(1, options, "variable names: no") > 0, _
                  "The file carries no variable name anywhere"
    Assert.IsTrue InStr(1, options, "variable labels: yes") > 0, _
                  "Its header row holds the labels a user reads"

    Exit Sub
TestFail:
    On Error Resume Next
    SetHeaderFormat 1, "default"
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestALabelHeaderFormatSaysThereAreNoNames", Err.Number, Err.Description
End Sub

'@sub-title A file carrying both puts its variable names on row 2.
'@details
'This one CAN be read back, and the options line is what tells the import where
'to look. Refusing it would cost the user a file that works.
'@TestMethod("LLExporter")
Public Sub TestNamesAndLabelsPutTheNamesOnRowTwo()
    CustomTestSetTitles Assert, TESTMODULE, "TestNamesAndLabelsPutTheNamesOnRowTwo"
    If Not FixtureReady("TestNamesAndLabelsPutTheNamesOnRowTwo") Then Exit Sub
    On Error GoTo TestFail

    Dim options As String

    SetHeaderFormat 1, "variable names + variable labels"
    options = FixtureExporter().ExportOptionsFor(1)
    SetHeaderFormat 1, "default"

    Assert.IsTrue InStr(1, options, "variable names: yes") > 0, _
                  "The file does carry its variable names"
    Assert.IsTrue InStr(1, options, "header row: 2") > 0, _
                  "Under the label row, on the second row"

    Exit Sub
TestFail:
    On Error Resume Next
    SetHeaderFormat 1, "default"
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestNamesAndLabelsPutTheNamesOnRowTwo", Err.Number, Err.Description
End Sub

'@sub-title A migration export always carries its variable names on row 1.
'@TestMethod("LLExporter")
Public Sub TestAMigrationAlwaysCarriesItsNames()
    CustomTestSetTitles Assert, TESTMODULE, "TestAMigrationAlwaysCarriesItsNames"
    If Not FixtureReady("TestAMigrationAlwaysCarriesItsNames") Then Exit Sub
    On Error GoTo TestFail

    Dim options As String

    options = FixtureExporter().MigrationOptions()

    Assert.IsTrue InStr(1, options, "variable names: yes") > 0, _
                  "A migration file carries its variable names"
    Assert.IsTrue InStr(1, options, "header row: 1") > 0, _
                  "On the first row, whatever the Exports rows say"
    Assert.IsTrue InStr(1, options, "filtered data: no") > 0, _
                  "And it carries the whole data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAMigrationAlwaysCarriesItsNames", Err.Number, Err.Description
End Sub

'@sub-title An option the Exports row leaves empty is written as no.
'@details
'So the reader on the other side never has to tell an absent option from a
'false one.
'@TestMethod("LLExporter")
Public Sub TestAnEmptyOptionIsWrittenAsNo()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnEmptyOptionIsWrittenAsNo"
    If Not FixtureReady("TestAnEmptyOptionIsWrittenAsNo") Then Exit Sub
    On Error GoTo TestFail

    Dim options As String

    SetHeaderFormat 1, vbNullString
    options = FixtureExporter().ExportOptionsFor(1)
    SetHeaderFormat 1, "default"

    Assert.IsTrue InStr(1, options, "header format: variable names") > 0, _
                  "An empty header format reads as the variable names format"
    Assert.IsTrue InStr(1, options, "variable names: yes") > 0, _
                  "So the file carries its names"

    Exit Sub
TestFail:
    On Error Resume Next
    SetHeaderFormat 1, "default"
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestAnEmptyOptionIsWrittenAsNo", Err.Number, Err.Description
End Sub

'@sub-title The filtered flag reaches the line.
'@TestMethod("LLExporter")
Public Sub TestTheFilteredFlagReachesTheLine()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFilteredFlagReachesTheLine"
    If Not FixtureReady("TestTheFilteredFlagReachesTheLine") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue InStr(1, FixtureExporter().ExportOptionsFor(1, True), _
                        "filtered data: yes") > 0, _
                  "A filtered export says so in the file"
    Assert.IsTrue InStr(1, FixtureExporter().ExportOptionsFor(1, False), _
                        "filtered data: no") > 0, _
                  "And an unfiltered one says that"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFilteredFlagReachesTheLine", Err.Number, Err.Description
End Sub


'@section The custom dropdown key
'===============================================================================

'@sub-title The export builds its choice key from the registry name.
'@details
'The export side of a round trip that has never once worked. It used to read the
'header cell of the ListObject, which holds the name with every space turned
'into an underscore, so a dropdown called `contact type` went into the file
'under a name no DropdownLists call would take back.
'@TestMethod("LLExporter")
Public Sub TestTheExportKeyIsTheRegistryName()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheExportKeyIsTheRegistryName"
    On Error GoTo TestFail

    Assert.AreEqual "__choice_custom_contact type", _
                    ChoiceKeys.Create().CustomChoiceName("contact type"), _
                    "Both ends of the round trip name the choice this way"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheExportKeyIsTheRegistryName", Err.Number, Err.Description
End Sub


'@section The open-from-file seam
'===============================================================================

'@sub-title A path with no file behind it is refused before any open.
'@TestMethod("LLExporter")
Public Sub TestCreateFromFileRefusesAMissingPath()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateFromFileRefusesAMissingPath"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim missingPath As String
    Dim raisedNumber As Long

    missingPath = BuildTempFolder(ThisWorkbook, FILES_FOLDER) & _
                  Application.PathSeparator & "no_such_linelist.xlsb"

    On Error Resume Next
    Set exporter = LLExporter.CreateFromFile(missingPath)
    raisedNumber = Err.Number
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.ElementNotFound), raisedNumber, _
                    "A path with no file behind it raises ElementNotFound"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateFromFileRefusesAMissingPath", Err.Number, Err.Description
End Sub

'@sub-title A workbook without the linelist sheets is refused and closed again.
'@details
'The factory opened the file to look inside, so the refusal has to close what
'it opened: the user picks a wrong file, reads the message and picks again,
'and a copy of the wrong file must not be left standing in the session.
'@TestMethod("LLExporter")
Public Sub TestCreateFromFileRefusesANonLinelist()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateFromFileRefusesANonLinelist"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim plainBook As Workbook
    Dim filePath As String
    Dim savedName As String
    Dim raisedNumber As Long
    Dim raisedMessage As String

    filePath = BuildWorkbookPath(BuildTempFolder(ThisWorkbook, FILES_FOLDER), _
                                 "not_a_linelist")

    Set plainBook = NewWorkbook()
    plainBook.SaveAs fileName:=filePath, fileFormat:=xlExcel12
    savedName = plainBook.Name
    plainBook.Close savechanges:=False
    Set plainBook = Nothing

    On Error Resume Next
    Set exporter = LLExporter.CreateFromFile(filePath)
    raisedNumber = Err.Number
    raisedMessage = Err.Description
    On Error GoTo TestFail

    'The description does not survive the class boundary, so the number is
    'what is asserted and the text rides along in the report
    '(.obt/gotchas/err-description-at-class-boundary.md).
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), raisedNumber, _
                    "A workbook missing a linelist sheet raises unexpected state" & _
                    " - description was [" & raisedMessage & "]"
    Assert.IsFalse WorkbookOpenByName(savedName), _
                   "The refused workbook is closed again"

    Kill filePath
    Exit Sub
TestFail:
    On Error Resume Next
    If Not plainBook Is Nothing Then plainBook.Close savechanges:=False
    If LenB(savedName) > 0 And WorkbookOpenByName(savedName) Then _
        Application.Workbooks(savedName).Close savechanges:=False
    If LenB(filePath) > 0 Then
        If Dir$(filePath) <> vbNullString Then Kill filePath
    End If
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestCreateFromFileRefusesANonLinelist", Err.Number, Err.Description
End Sub

'@sub-title A file the factory opened belongs to it, and CloseAll closes it.
'@TestMethod("LLExporter")
Public Sub TestCreateFromFileOwnsTheFileItOpened()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateFromFileOwnsTheFileItOpened"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim filePath As String
    Dim savedName As String

    filePath = BuildWorkbookPath(BuildTempFolder(ThisWorkbook, FILES_FOLDER), _
                                 "owned_linelist")
    savedName = SaveLinelistShell(filePath)

    Set exporter = LLExporter.CreateFromFile(filePath)

    Assert.IsTrue exporter.OpenedFromFile, _
                  "An instance that opened its file says so"
    Assert.IsTrue WorkbookOpenByName(savedName), _
                  "The linelist file is open in the session"

    exporter.CloseAll

    Assert.IsFalse WorkbookOpenByName(savedName), _
                   "CloseAll closes the source it opened"

    Kill filePath
    Exit Sub
TestFail:
    On Error Resume Next
    If LenB(savedName) > 0 And WorkbookOpenByName(savedName) Then _
        Application.Workbooks(savedName).Close savechanges:=False
    If LenB(filePath) > 0 Then
        If Dir$(filePath) <> vbNullString Then Kill filePath
    End If
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestCreateFromFileOwnsTheFileItOpened", Err.Number, Err.Description
End Sub

'@sub-title A workbook already open in the session is used as it stands.
'@details
'This is also what protects the current linelist: a path already open is
'answered with the open workbook, ownership stays False, and CloseAll leaves
'it standing.
'@TestMethod("LLExporter")
Public Sub TestCreateFromFileLeavesAnOpenWorkbookStanding()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateFromFileLeavesAnOpenWorkbookStanding"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim shellBook As Workbook
    Dim filePath As String

    filePath = BuildWorkbookPath(BuildTempFolder(ThisWorkbook, FILES_FOLDER), _
                                 "open_linelist")

    Set shellBook = LinelistShellWorkbook()
    shellBook.SaveAs fileName:=filePath, fileFormat:=xlExcel12

    Set exporter = LLExporter.CreateFromFile(filePath)

    Assert.IsFalse exporter.OpenedFromFile, _
                   "An instance on a workbook the session already held owns nothing"

    exporter.CloseAll

    Assert.IsTrue WorkbookOpenByName(shellBook.Name), _
                  "CloseAll leaves the workbook it never opened standing"

    shellBook.Close savechanges:=False
    Kill filePath
    Exit Sub
TestFail:
    On Error Resume Next
    If Not shellBook Is Nothing Then shellBook.Close savechanges:=False
    If LenB(filePath) > 0 Then
        If Dir$(filePath) <> vbNullString Then Kill filePath
    End If
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestCreateFromFileLeavesAnOpenWorkbookStanding", Err.Number, Err.Description
End Sub

'@sub-title A linelist saved before the internal-sheet rename still opens.
'@details
'The temp sheet of such a file carries the old trailing name, and the factory
'reads it through the same fallback the export walks use.
'@TestMethod("LLExporter")
Public Sub TestCreateFromFileReadsTheOldTempName()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateFromFileReadsTheOldTempName"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim shellBook As Workbook
    Dim filePath As String
    Dim savedName As String

    filePath = BuildWorkbookPath(BuildTempFolder(ThisWorkbook, FILES_FOLDER), _
                                 "old_named_linelist")

    Set shellBook = NewWorkbook()
    EnsureWorksheet DICTIONARY_SHEET, shellBook
    EnsureWorksheet EXPORTS_SHEET, shellBook
    EnsureWorksheet PASS_SHEET, shellBook
    EnsureWorksheet GEO_SHEET, shellBook
    EnsureWorksheet TEMP_SHEET_OLD, shellBook
    shellBook.SaveAs fileName:=filePath, fileFormat:=xlExcel12
    savedName = shellBook.Name
    shellBook.Close savechanges:=False
    Set shellBook = Nothing

    Set exporter = LLExporter.CreateFromFile(filePath)

    Assert.IsTrue exporter.OpenedFromFile, _
                  "The factory accepts the old temp sheet name"

    exporter.CloseAll

    Assert.IsFalse WorkbookOpenByName(savedName), _
                   "CloseAll closes the file it opened"

    Kill filePath
    Exit Sub
TestFail:
    On Error Resume Next
    If Not shellBook Is Nothing Then shellBook.Close savechanges:=False
    If Not exporter Is Nothing Then exporter.CloseAll
    If LenB(savedName) > 0 And WorkbookOpenByName(savedName) Then _
        Application.Workbooks(savedName).Close savechanges:=False
    If LenB(filePath) > 0 Then
        If Dir$(filePath) <> vbNullString Then Kill filePath
    End If
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestCreateFromFileReadsTheOldTempName", Err.Number, Err.Description
End Sub


'@section What a failed geo export says about itself
'===============================================================================

'@sub-title A geo export that fails keeps the error in words of its own.
'@details
'The description of an error raised inside the class does not survive the way
'out, so the form that calls the export reads the name of the method and
'nothing of what refused. Acts by exporting the geobase of a workbook with no
'Geo worksheet, which is the first step of the walk. Asserts the raise still
'reaches the caller and that LastFailure names the step and carries the error
'number.
'@TestMethod("LLExporter")
Public Sub TestAFailedGeoExportSaysWhatItFailedOn()
    CustomTestSetTitles Assert, TESTMODULE, "TestAFailedGeoExportSaysWhatItFailedOn"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim plainBook As Workbook
    Dim folderPath As String
    Dim raisedNumber As Long
    Dim failure As String

    folderPath = BuildTempFolder(ThisWorkbook, FILES_FOLDER)

    Set plainBook = NewWorkbook()
    Set exporter = LLExporter.Create(plainBook)

    On Error Resume Next
    exporter.ExportGeo folderPath
    raisedNumber = Err.Number
    On Error GoTo TestFail

    failure = exporter.LastFailure
    DeleteWorkbook plainBook
    Set plainBook = Nothing

    Assert.IsTrue raisedNumber <> 0, _
                  "A geo export with no Geo worksheet still raises to its caller"
    Assert.IsTrue LenB(failure) > 0, _
                  "A failed geo export leaves an account of the failure behind"
    Assert.IsTrue InStr(1, failure, "reading the Geo sheet") > 0, _
                  "The account names the step the export was on"
    Assert.IsTrue InStr(1, failure, CStr(raisedNumber)) > 0, _
                  "The account carries the number the caller cannot read from the text"

    Exit Sub
TestFail:
    On Error Resume Next
    If Not plainBook Is Nothing Then DeleteWorkbook plainBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestAFailedGeoExportSaysWhatItFailedOn", Err.Number, Err.Description
End Sub

'@sub-title An export that worked leaves nothing behind to report.
'@details
'LastFailure is read whenever a walk lands on its error label, so a value left
'over from an earlier export would be logged as the cause of a later one.
'Asserts a fresh exporter has nothing to report.
'@TestMethod("LLExporter")
Public Sub TestAnExporterWithNoFailureReportsNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnExporterWithNoFailureReportsNothing"
    On Error GoTo TestFail

    Dim exporter As LLExporter

    Set exporter = LLExporter.Create(ThisWorkbook)

    Assert.AreEqual vbNullString, exporter.LastFailure, _
                    "An exporter that has run nothing has no failure to report"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnExporterWithNoFailureReportsNothing", Err.Number, Err.Description
End Sub


'@section Fixture helpers
'===============================================================================

'@sub-title Write one header format into the Exports worksheet of the fixture.
'@param exportNumber Long. Which export row to change (1-based).
'@param headerFormat String. The value to write.
Private Sub SetHeaderFormat(ByVal exportNumber As Long, ByVal headerFormat As String)
    FixtureWorkbook.Worksheets(EXPORTS_SHEET).Cells(exportNumber + 1, 9).Value = headerFormat
End Sub

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

'@sub-title Put one fixture row on an admin level.
'@details
'The shared dictionary fixture holds the RAW geo row, control `geo`, the way a
'setup file writes it. A linelist dictionary holds what LLdictionary.Prepare
'made of it, and the level branch of the exporter reads that: `geo1` to `geo4`.
'This writes the prepared control onto the one geo row of the fixture, so the
'fixture answers the question the exporter asks of a real linelist.
Private Sub TagOneGeoLevelRow()
    Dim sh As Worksheet
    Dim nameColumn As Long
    Dim controlColumn As Long
    Dim rowCounter As Long

    Set sh = FixtureWorkbook.Worksheets(DICTIONARY_SHEET)
    nameColumn = DictionaryTestFixture.DictionaryHeaderIndex("Variable Name") + 1
    controlColumn = DictionaryTestFixture.DictionaryHeaderIndex("Control") + 1

    For rowCounter = 2 To DictionaryTestFixture.DictionaryFixtureRowCount() + 1
        If StrComp(CStr(sh.Cells(rowCounter, nameColumn).Value), _
                   GEO_LEVEL3_VARIABLE, vbTextCompare) = 0 Then
            sh.Cells(rowCounter, controlColumn).Value = GEO_LEVEL3_CONTROL
            Exit For
        End If
    Next rowCounter
End Sub

'@sub-title Write the six-row Exports worksheet the tests read.
'@details
'"admin levels" is written LAST on the header row on purpose. SetHeaderFormat
'reaches the header format by its column number, and a column inserted before
'it would send that write somewhere else.
Private Sub BuildExportsSheet()
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(EXPORTS_SHEET, FixtureWorkbook, clearSheet:=True)

    WriteRow sh.Cells(1, 1), "export number", "status", "label button", _
                             "file format", "file name", "password", _
                             "include personal identifiers", "include p-codes", _
                             "header format", "export metadata sheets", _
                             "export analyses sheets", "admin levels"

    WriteRow sh.Cells(2, 1), 1, "active", "both off", "xlsx", "one", "no", _
                             "no", "no", "default", "no", "no", vbNullString
    WriteRow sh.Cells(3, 1), 2, "active", "both on", "xlsx", "two", "no", _
                             "yes", "yes", "default", "no", "no", vbNullString
    WriteRow sh.Cells(4, 1), 3, "active", "identifiers off", "xlsx", "three", "no", _
                             "no", "yes", "default", "no", "no", vbNullString

    'The three level rows keep both flags on, so the levels are the only rule
    'left to answer for what they drop.
    WriteRow sh.Cells(5, 1), 4, "active", "level one", "xlsx", "four", "no", _
                             "yes", "yes", "default", "no", "no", "admin1"
    WriteRow sh.Cells(6, 1), 5, "active", "level three", "xlsx", "five", "no", _
                             "yes", "yes", "default", "no", "no", "admin3"
    WriteRow sh.Cells(7, 1), 6, "active", "level misspelt", "xlsx", "six", "no", _
                             "yes", "yes", "default", "no", "no", "admin1+adm2"
End Sub

'@fun-title A workbook carrying the five sheets CreateFromFile looks for.
'@details
'Dictionary, Exports, __pass, Geo and __temp, all empty: the factory checks
'presence and reads nothing, so a shell is enough.
'@return Workbook. The new workbook, open and unsaved.
Private Function LinelistShellWorkbook() As Workbook
    Dim shellBook As Workbook

    Set shellBook = NewWorkbook()
    EnsureWorksheet DICTIONARY_SHEET, shellBook
    EnsureWorksheet EXPORTS_SHEET, shellBook
    EnsureWorksheet PASS_SHEET, shellBook
    EnsureWorksheet GEO_SHEET, shellBook
    EnsureWorksheet TEMP_SHEET, shellBook

    Set LinelistShellWorkbook = shellBook
End Function

'@fun-title Save a linelist shell to a path and close it.
'@param filePath String. Where the file lands.
'@return String. The file name the saved workbook carries.
Private Function SaveLinelistShell(ByVal filePath As String) As String
    Dim shellBook As Workbook

    Set shellBook = LinelistShellWorkbook()
    shellBook.SaveAs fileName:=filePath, fileFormat:=xlExcel12
    SaveLinelistShell = shellBook.Name
    shellBook.Close savechanges:=False
End Function

'@fun-title Whether a workbook of that name is open in the session.
'@param bookName String. The workbook name, extension included.
'@return Boolean. True when the session holds it.
Private Function WorkbookOpenByName(ByVal bookName As String) As Boolean
    Dim wkb As Workbook

    'An absent name raises on the read, and Nothing is the answer read back.
    On Error Resume Next
    Set wkb = Application.Workbooks(bookName)
    On Error GoTo 0

    WorkbookOpenByName = Not wkb Is Nothing
End Function

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
