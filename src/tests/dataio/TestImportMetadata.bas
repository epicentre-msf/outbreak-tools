Attribute VB_Name = "TestImportMetadata"
Attribute VB_Description = "Unit tests for ImportMetadata"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Unit tests for ImportMetadata")

Option Explicit

'@description
'Drives ImportMetadata, the reader of the Metadata worksheet of a file being
'imported. Every answer of the class is asserted against a real worksheet.
'
'WHAT THESE TESTS ARE GUARDING
'-------------------------------------------------------------------------------
'Four places used to build a DataSheet over this worksheet and run a Range.Find
'with LookIn, SearchOrder and SearchDirection unset, which Excel fills from the
'last Find of the session including one a user ran by hand. And the answer to
'"is this file in our language" was kept on the importer, so a caller that
'skipped the check got False by default and three metadata imports did nothing
'in silence.
'
'The options line is the other half. It is what tells an import whether a file
'can be read back at all, and a file written before the line existed has to fall
'back to the shape every export used to have.
'
'THREE WORKBOOKS PER MODULE
'-------------------------------------------------------------------------------
'A full one, a bare one with no Metadata worksheet, and one whose Metadata
'worksheet carries no `variable` header. All three are built once.
'@depends ImportMetadata, CustomTest

Private Assert As CustomTest
Private FullWorkbook As Workbook
Private BareWorkbook As Workbook
Private HeaderlessWorkbook As Workbook
Private SetupError As Long
Private SetupMessage As String

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ImportMetadata"
Private Const METADATA_SHEET As String = "Metadata"

Private Const LANGUAGE_VALUE As String = "English"
Private Const EXPORT_TYPE_VALUE As String = "custom export: export 2"
Private Const EPIWEEK_VALUE As String = "5"

'The options line carries a value holding a colon of its own, so the split on
'the first colon is what has to be right.
Private Const OPTIONS_LINE As String = _
    "header format: variable names + variable labels | variable names: yes | " & _
    "variable labels: yes | header row: 2 | filtered data: no | " & _
    "file name: report: weekly"


'@section Lifecycle
'===============================================================================

'@sub-title Build the assertion harness and the three fixture workbooks.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()

    Dim sh As Worksheet

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestImportMetadata"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set FullWorkbook = NewWorkbook()
        Set sh = EnsureWorksheet(METADATA_SHEET, FullWorkbook, clearSheet:=True)
        WriteRow sh.Cells(1, 1), "variable", "value"
        WriteRow sh.Cells(2, 1), "language", LANGUAGE_VALUE
        WriteRow sh.Cells(3, 1), "export_type", EXPORT_TYPE_VALUE
        WriteRow sh.Cells(4, 1), "export_options", OPTIONS_LINE
        WriteRow sh.Cells(5, 1), "epiweekstart", EPIWEEK_VALUE
        WriteRow sh.Cells(6, 1), "show_hide", "update on import"
        WriteRow sh.Cells(7, 1), "editable_labels", "no update on import"

        Set BareWorkbook = NewWorkbook()

        Set HeaderlessWorkbook = NewWorkbook()
        Set sh = EnsureWorksheet(METADATA_SHEET, HeaderlessWorkbook, clearSheet:=True)
        WriteRow sh.Cells(1, 1), "something", "else"
        WriteRow sh.Cells(2, 1), "language", "German"

        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print the results and drop the fixture workbooks.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()

    On Error Resume Next
        If Not HeaderlessWorkbook Is Nothing Then DeleteWorkbook HeaderlessWorkbook
        If Not BareWorkbook Is Nothing Then DeleteWorkbook BareWorkbook
        If Not FullWorkbook Is Nothing Then DeleteWorkbook FullWorkbook
    On Error GoTo 0

    Set HeaderlessWorkbook = Nothing
    Set BareWorkbook = Nothing
    Set FullWorkbook = Nothing

    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Put the application into its test state.
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

'@sub-title Create raises when the workbook argument is Nothing.
'@TestMethod("ImportMetadata")
Public Sub FactoryRejectsNothingWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "FactoryRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim meta As ImportMetadata
    On Error Resume Next
    Set meta = ImportMetadata.Create(Nothing)
    Assert.IsTrue Err.Number <> 0, "Factory should raise for a Nothing workbook"
    Err.Clear
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingWorkbook", Err.Number, Err.Description
End Sub

'@sub-title A workbook carrying a Metadata worksheet is read.
'@TestMethod("ImportMetadata")
Public Sub TestAFileWithMetadataIsRead()
    CustomTestSetTitles Assert, TESTMODULE, "TestAFileWithMetadataIsRead"
    If Not FixtureReady("TestAFileWithMetadataIsRead") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue FullMetadata().Exists, "The worksheet was found and read"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFileWithMetadataIsRead", Err.Number, Err.Description
End Sub

'@sub-title A workbook with no Metadata worksheet answers empty, and never raises.
'@details
'A file exported with the metadata sheets switched off carries none, so this is
'a real case and it is answered rather than raised.
'@TestMethod("ImportMetadata")
Public Sub TestAFileWithNoMetadataAnswersEmpty()
    CustomTestSetTitles Assert, TESTMODULE, "TestAFileWithNoMetadataAnswersEmpty"
    If Not FixtureReady("TestAFileWithNoMetadataAnswersEmpty") Then Exit Sub
    On Error GoTo TestFail

    Dim meta As ImportMetadata
    Set meta = ImportMetadata.Create(BareWorkbook)

    Assert.IsFalse meta.Exists, "The worksheet is absent"
    Assert.AreEqual vbNullString, meta.Language, "So there is no language"
    Assert.AreEqual vbNullString, meta.ExportType, "And no export type"
    Assert.AreEqual vbNullString, meta.EpiWeekStart, "And no epiweek value"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFileWithNoMetadataAnswersEmpty", Err.Number, Err.Description
End Sub

'@sub-title A worksheet named Metadata with no variable header is not one of ours.
'@details
'DataRange raises on the missing column. That is answered as a file carrying no
'metadata rather than as a failed import.
'@TestMethod("ImportMetadata")
Public Sub TestAMetadataSheetWithNoVariableHeaderIsAnswered()
    CustomTestSetTitles Assert, TESTMODULE, "TestAMetadataSheetWithNoVariableHeaderIsAnswered"
    If Not FixtureReady("TestAMetadataSheetWithNoVariableHeaderIsAnswered") Then Exit Sub
    On Error GoTo TestFail

    Dim meta As ImportMetadata
    Set meta = ImportMetadata.Create(HeaderlessWorkbook)

    Assert.IsFalse meta.Exists, "A worksheet with no variable header carries no metadata"
    Assert.AreEqual vbNullString, meta.Language, "And answers nothing rather than raising"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAMetadataSheetWithNoVariableHeaderIsAnswered", Err.Number, Err.Description
End Sub


'@section The tags
'===============================================================================

'@sub-title The three named tags answer what the worksheet holds.
'@TestMethod("ImportMetadata")
Public Sub TestTheNamedTagsAnswerTheWorksheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheNamedTagsAnswerTheWorksheet"
    If Not FixtureReady("TestTheNamedTagsAnswerTheWorksheet") Then Exit Sub
    On Error GoTo TestFail

    Dim meta As ImportMetadata
    Set meta = FullMetadata()

    Assert.AreEqual LANGUAGE_VALUE, meta.Language, "The language tag"
    Assert.AreEqual EXPORT_TYPE_VALUE, meta.ExportType, "The export type tag"
    Assert.AreEqual EPIWEEK_VALUE, meta.EpiWeekStart, "The epiweek tag"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheNamedTagsAnswerTheWorksheet", Err.Number, Err.Description
End Sub

'@sub-title A tag name is matched without regard to case.
'@details
'The Range.Find this replaced ran with MatchCase False, so the same files have to
'keep answering.
'@TestMethod("ImportMetadata")
Public Sub TestATagNameIsMatchedWithoutCase()
    CustomTestSetTitles Assert, TESTMODULE, "TestATagNameIsMatchedWithoutCase"
    If Not FixtureReady("TestATagNameIsMatchedWithoutCase") Then Exit Sub
    On Error GoTo TestFail

    Assert.AreEqual LANGUAGE_VALUE, FullMetadata().Tag("LANGUAGE"), _
                    "A tag asked for in capitals answers"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestATagNameIsMatchedWithoutCase", Err.Number, Err.Description
End Sub

'@sub-title A tag the file does not carry answers an empty string.
'@TestMethod("ImportMetadata")
Public Sub TestAnAbsentTagAnswersEmpty()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnAbsentTagAnswersEmpty"
    If Not FixtureReady("TestAnAbsentTagAnswersEmpty") Then Exit Sub
    On Error GoTo TestFail

    Assert.AreEqual vbNullString, FullMetadata().Tag("no_such_tag"), _
                    "A tag the file has none of answers nothing"
    Assert.AreEqual vbNullString, FullMetadata().Tag(vbNullString), _
                    "And so does an empty tag name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnAbsentTagAnswersEmpty", Err.Number, Err.Description
End Sub

'@sub-title The two update tags are read one at a time.
'@details
'`show_hide` asks for an update here and `editable_labels` does not, so a single
'flag standing in for both would fail this one.
'@TestMethod("ImportMetadata")
Public Sub TestTheTwoUpdateTagsAreReadSeparately()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheTwoUpdateTagsAreReadSeparately"
    If Not FixtureReady("TestTheTwoUpdateTagsAreReadSeparately") Then Exit Sub
    On Error GoTo TestFail

    Dim meta As ImportMetadata
    Set meta = FullMetadata()

    Assert.IsTrue meta.UpdatesOnImport("show_hide"), _
                  "The show/hide tag asks for an update"
    Assert.IsFalse meta.UpdatesOnImport("editable_labels"), _
                   "And the editable labels tag does not"
    Assert.IsFalse meta.UpdatesOnImport("no_such_tag"), _
                   "A tag the file has none of asks for nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheTwoUpdateTagsAreReadSeparately", Err.Number, Err.Description
End Sub

'@sub-title A migration export is told apart from a custom one.
'@TestMethod("ImportMetadata")
Public Sub TestAMigrationExportIsToldApart()
    CustomTestSetTitles Assert, TESTMODULE, "TestAMigrationExportIsToldApart"
    If Not FixtureReady("TestAMigrationExportIsToldApart") Then Exit Sub
    On Error GoTo TestFail

    Dim otherBook As Workbook
    Dim sh As Worksheet

    Assert.IsFalse FullMetadata().IsMigrationExport, _
                   "The fixture file is a custom export"

    Set otherBook = NewWorkbook()
    Set sh = EnsureWorksheet(METADATA_SHEET, otherBook, clearSheet:=True)
    WriteRow sh.Cells(1, 1), "variable", "value"
    WriteRow sh.Cells(2, 1), "export_type", "export for migrations"

    Assert.IsTrue ImportMetadata.Create(otherBook).IsMigrationExport, _
                  "And a file naming a migration is one"

    DeleteWorkbook otherBook

    Exit Sub
TestFail:
    On Error Resume Next
    If Not otherBook Is Nothing Then DeleteWorkbook otherBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestAMigrationExportIsToldApart", Err.Number, Err.Description
End Sub


'@section The export options line
'===============================================================================

'@sub-title The options line is split back into its pairs.
'@TestMethod("ImportMetadata")
Public Sub TestTheOptionsLineIsSplitBack()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheOptionsLineIsSplitBack"
    If Not FixtureReady("TestTheOptionsLineIsSplitBack") Then Exit Sub
    On Error GoTo TestFail

    Dim meta As ImportMetadata
    Set meta = FullMetadata()

    Assert.AreEqual "variable names + variable labels", meta.ExportOption("header format"), _
                    "A value carrying its own spaces comes back whole"
    Assert.AreEqual "yes", meta.ExportOption("variable names"), _
                    "And so does a plain one"
    Assert.AreEqual "no", meta.ExportOption("filtered data"), _
                    "And the last pair of the line"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheOptionsLineIsSplitBack", Err.Number, Err.Description
End Sub

'@sub-title A value carrying a colon of its own survives.
'@details
'The pair is split on the FIRST colon, so `file name: report: weekly` answers
'`report: weekly` rather than `report`.
'@TestMethod("ImportMetadata")
Public Sub TestAValueCarryingAColonSurvives()
    CustomTestSetTitles Assert, TESTMODULE, "TestAValueCarryingAColonSurvives"
    If Not FixtureReady("TestAValueCarryingAColonSurvives") Then Exit Sub
    On Error GoTo TestFail

    Assert.AreEqual "report: weekly", FullMetadata().ExportOption("file name"), _
                    "The split takes the first colon and leaves the rest alone"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAValueCarryingAColonSurvives", Err.Number, Err.Description
End Sub

'@sub-title An option the file does not carry answers an empty string.
'@TestMethod("ImportMetadata")
Public Sub TestAnAbsentOptionAnswersEmpty()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnAbsentOptionAnswersEmpty"
    If Not FixtureReady("TestAnAbsentOptionAnswersEmpty") Then Exit Sub
    On Error GoTo TestFail

    Assert.AreEqual vbNullString, FullMetadata().ExportOption("no such option"), _
                    "An option the line has none of answers nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnAbsentOptionAnswersEmpty", Err.Number, Err.Description
End Sub


'@section What the options decide
'===============================================================================

'@sub-title A file carrying both names and labels is readable, from row 2.
'@TestMethod("ImportMetadata")
Public Sub TestAFileCarryingBothIsReadFromRowTwo()
    CustomTestSetTitles Assert, TESTMODULE, "TestAFileCarryingBothIsReadFromRowTwo"
    If Not FixtureReady("TestAFileCarryingBothIsReadFromRowTwo") Then Exit Sub
    On Error GoTo TestFail

    Dim meta As ImportMetadata
    Set meta = FullMetadata()

    Assert.IsTrue meta.CarriesVariableNames, _
                  "The file does carry its variable names"
    Assert.AreEqual CLng(2), meta.HeaderRow, _
                    "And they are on the second row, under the labels"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFileCarryingBothIsReadFromRowTwo", Err.Number, Err.Description
End Sub

'@sub-title A file whose header row holds labels alone carries no variable names.
'@details
'The one case that cannot be read back at all, and the reason the import refuses
'it rather than reporting every column as unmatched.
'@TestMethod("ImportMetadata")
Public Sub TestALabelHeaderCarriesNoVariableNames()
    CustomTestSetTitles Assert, TESTMODULE, "TestALabelHeaderCarriesNoVariableNames"
    If Not FixtureReady("TestALabelHeaderCarriesNoVariableNames") Then Exit Sub
    On Error GoTo TestFail

    Dim meta As ImportMetadata

    Set meta = MetadataForOptions("header format: variables labels | " & _
                                  "variable names: no | header row: 1")

    Assert.IsFalse meta.CarriesVariableNames, _
                   "No column of this file can be matched to a column of a linelist"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALabelHeaderCarriesNoVariableNames", Err.Number, Err.Description
End Sub

'@sub-title A file written before the options line existed falls back to row 1.
'@details
'Every export used to put the variable names on row 1, so that is what a file
'carrying no options line is read as. This is what keeps every file already in
'the field readable.
'@TestMethod("ImportMetadata")
Public Sub TestAFileWithNoOptionsFallsBackToRowOne()
    CustomTestSetTitles Assert, TESTMODULE, "TestAFileWithNoOptionsFallsBackToRowOne"
    If Not FixtureReady("TestAFileWithNoOptionsFallsBackToRowOne") Then Exit Sub
    On Error GoTo TestFail

    Dim meta As ImportMetadata
    Set meta = ImportMetadata.Create(BareWorkbook)

    Assert.IsTrue meta.CarriesVariableNames, _
                  "A file that says nothing is taken as carrying its names"
    Assert.AreEqual CLng(1), meta.HeaderRow, _
                    "On the first row, which is what every export used to write"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFileWithNoOptionsFallsBackToRowOne", Err.Number, Err.Description
End Sub

'@sub-title A header row that is not a number falls back to row 1.
'@TestMethod("ImportMetadata")
Public Sub TestAJunkHeaderRowFallsBackToRowOne()
    CustomTestSetTitles Assert, TESTMODULE, "TestAJunkHeaderRowFallsBackToRowOne"
    If Not FixtureReady("TestAJunkHeaderRowFallsBackToRowOne") Then Exit Sub
    On Error GoTo TestFail

    Assert.AreEqual CLng(1), MetadataForOptions("header row: somewhere").HeaderRow, _
                    "A value that is not a row number is read as row 1"
    Assert.AreEqual CLng(1), MetadataForOptions("header row: 0").HeaderRow, _
                    "And so is a row number below 1"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAJunkHeaderRowFallsBackToRowOne", Err.Number, Err.Description
End Sub


'@section Fixture helpers
'===============================================================================

'@fun-title Report a fixture that could not be built as this test's failure.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is usable.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If SetupError = 0 Then
        If Not FullWorkbook Is Nothing Then
            If Not BareWorkbook Is Nothing Then
                FixtureReady = True
                Exit Function
            End If
        End If
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The fixture could not be built - " & SetupMessage
End Function

'@fun-title A reader over the full fixture workbook.
'@return ImportMetadata. A fresh instance.
Private Function FullMetadata() As ImportMetadata
    Set FullMetadata = ImportMetadata.Create(FullWorkbook)
End Function

'@fun-title A reader over a throwaway workbook carrying one options line.
'@details
'The workbook is left open for the length of the module and dropped with the
'others; building one per call costs nothing and keeps each test independent.
'@param options String. The options line to write.
'@return ImportMetadata. A reader over that line.
Private Function MetadataForOptions(ByVal options As String) As ImportMetadata
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(METADATA_SHEET, FullWorkbook, clearSheet:=False)
    sh.Cells(4, 2).Value = options

    Set MetadataForOptions = ImportMetadata.Create(FullWorkbook)

    'Put the fixture line back, so the tests that read it are unaffected.
    sh.Cells(4, 2).Value = OPTIONS_LINE
End Function
