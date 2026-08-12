Attribute VB_Name = "TestLLImporter"
Attribute VB_Description = "Unit tests for LLImporter"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Unit tests for LLImporter")

Option Explicit

'@description
'Drives LLImporter, the class that reads an exported workbook back into a
'linelist. Every import path is driven against real worksheets: the data lands
'in a real ListObject, the custom dropdown is a real DropdownLists registry, the
'epiweek value goes into a real hidden name store, and the report is read back
'off a real worksheet.
'
'THREE WORKBOOKS PER MODULE
'-------------------------------------------------------------------------------
'FixtureWorkbook stands in for the linelist being imported into,
'ImportWorkbook for the file being read, and GeobaseWorkbook for the geobase
'file the geo import reads. All three are built once in ModuleInitialize.
'Every test that writes into the linelist puts the sheet it touches back
'first, so the order the tests run in decides nothing.
'
'THE FIXTURE FAILS QUIETLY AND EVERY TEST SAYS SO
'-------------------------------------------------------------------------------
'An error escaping ModuleInitialize reaches the VBE as a modal dialog and the
'whole headless run comes back with no results file. The setup captures its
'error into two module fields and FixtureReady reports it as each test's own
'failure.
'
'WHAT THE LINELIST FIXTURE CARRIES
'-------------------------------------------------------------------------------
'  Dictionary          the shared fixture, three sheets over 75 variables
'  Choices             one list, list_correct_order, with three values
'  __pass              the shared passwords fixture
'  hlist2D-sheet1      a three-column ListObject, sheet_type HList
'  Custom dropdown     one dropdown named "contact type", registry-backed
'  RNG_DictionaryLanguage, RNG_CustomDrop, RNG_EpiWeekStart
'
'WHAT THE IMPORT FIXTURE CARRIES
'-------------------------------------------------------------------------------
'  Metadata            language, export type, options, epiweek, the two tags
'  Dictionary          the same fixture with one main label rewritten
'  Choices             list_correct_order with a fourth value, and the dropdown
'  hlist2D-sheet1      the three columns, a fourth the linelist lacks, two rows
'  ghost-sheet         a worksheet the linelist has none of
'
'WHAT THE GEOBASE FIXTURE CARRIES
'-------------------------------------------------------------------------------
'The linelist takes its Geo worksheet from GeoTestFixture, which fills T_NAMES
'with Province, District, Commune, Village and Health Facility in an EN column.
'GeobaseWorkbook holds the three sheets LLGeo.Import reads by name -- ADM1,
'ADM2 and NAMES -- and its NAMES sheet gives admin1 and admin2 labels of their
'own, so an import moves the five workbook-scoped level labels.
'
'Each geo test seeds the Geo worksheet again, which drops the level labels and
'the translated flag and puts Province back. That is what lets both tests read
'the same before-and-after pair whatever order they run in.
'@depends LLImporter, ImportMetadata, ImportReport, ChoiceKeys, LLChoices
'@depends DropdownLists, HiddenNames, Passwords, CustomTest, LLGeo, GeoTestFixture

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private ImportWorkbook As Workbook
Private GeobaseWorkbook As Workbook
Private SetupError As Long
Private SetupMessage As String

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "LLImporter"

Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const CHOICES_SHEET As String = "Choices"
Private Const METADATA_SHEET As String = "Metadata"
Private Const PASSWORD_SHEET As String = "__pass"
Private Const DROPDOWN_SHEET As String = "Custom dropdown"

Private Const HLIST_SHEET As String = "hlist2D-sheet1"
Private Const GHOST_SHEET As String = "ghost-sheet"
Private Const HLIST_TABLE As String = "table1"

'The geobase worksheet of the linelist and the two tables the geo tests read
Private Const GEO_SHEET As String = "Geo"
Private Const GEO_ADM1_TABLE As String = "T_ADM1"
Private Const GEO_ADM2_TABLE As String = "T_ADM2"

'The workbook-scoped hidden name holding the admin1 label
Private Const ADM1_LEVEL_NAME As String = "RNG_ADM1NAME"

'The admin1 label of the fixture geobase, and the one the file brings
Private Const ADM1_LABEL_BEFORE As String = "Province"
Private Const ADM1_LABEL_IN_FILE As String = "Region"
Private Const ADM2_LABEL_IN_FILE As String = "Zone"

'The first admin1 value of the file, and how many rows each of its two admin
'sheets carries
Private Const ADM1_FIRST_VALUE As String = "Alpha"
Private Const GEOBASE_ADM1_ROWS As Long = 2
Private Const GEOBASE_ADM2_ROWS As Long = 4

'The one geo-controlled variable of the dictionary fixture. UpdateDict writes
'the level label over its main label, and the cell goes back after each geo
'test.
Private Const HF_VARIABLE As String = "hf_h2"
Private Const HF_VARIABLE_LABEL As String = "HF on hlist2D"

Private Const LANGUAGE_NAME As String = "RNG_DictionaryLanguage"
Private Const LANGUAGE_VALUE As String = "English"
Private Const CUSTOMDROP_NAME As String = "RNG_CustomDrop"
Private Const EPIWEEK_NAME As String = "RNG_EpiWeekStart"
Private Const EPIWEEK_BEFORE As String = "1"
Private Const EPIWEEK_IN_FILE As String = "5"

'The custom dropdown is named with a space on purpose. The key the two ends
'built used to turn that space into an underscore on one side and leave a
'workbook counter on the other, so either of the old shapes fails these tests.
Private Const DROPDOWN_NAME As String = "contact type"
Private Const DROPDOWN_NEW_VALUE As String = "referral"

'Three real variable names off the hlist2D-sheet1 rows of the dictionary
'fixture, and one the fixture has never carried.
Private Const VAR_ONE As String = "date_h2"
Private Const VAR_TWO As String = "int_h2"
Private Const VAR_THREE As String = "choi_h2"
Private Const VAR_ABSENT As String = "not_a_variable_h2"

'The variable whose main label the file disagrees about
Private Const RELABELLED_IN_FILE As String = "A label the file disagrees about"

'The value the first imported row carries. It is plain text on purpose: Excel
'parses anything that looks like a date and reads it back in the host's own
'format, so a date literal here asserts the locale rather than the import.
Private Const FIRST_ROW_VALUE As String = "row-one"


'@section Lifecycle
'===============================================================================

'@sub-title Build the assertion harness and the two fixture workbooks.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLImporter"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set FixtureWorkbook = NewWorkbook()
        BuildLinelistFixture

        Set ImportWorkbook = NewWorkbook()
        BuildImportFixture

        Set GeobaseWorkbook = NewWorkbook()
        BuildGeobaseFixture

        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print the results and drop the fixture workbooks.
'@details
'The three fixture workbooks go first. PrintResults writes validation and named
'ranges onto the output sheet of ThisWorkbook, and those writes want ThisWorkbook
'to be the active book.
'
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()

    On Error Resume Next
        If Not GeobaseWorkbook Is Nothing Then DeleteWorkbook GeobaseWorkbook
        If Not ImportWorkbook Is Nothing Then DeleteWorkbook ImportWorkbook
        If Not FixtureWorkbook Is Nothing Then DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set GeobaseWorkbook = Nothing
    Set ImportWorkbook = Nothing
    Set FixtureWorkbook = Nothing

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

'@sub-title Create answers an importer for a real workbook.
'@TestMethod("LLImporter")
Public Sub FactoryCreatesWithWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "FactoryCreatesWithWorkbook"
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Set impObj = LLImporter.Create(ThisWorkbook)
    Assert.IsNotNothing impObj, "Factory should return a valid object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryCreatesWithWorkbook", Err.Number, Err.Description
End Sub

'@sub-title Create raises when the workbook argument is Nothing.
'@TestMethod("LLImporter")
Public Sub FactoryRejectsNothingWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "FactoryRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim impObj As LLImporter
    On Error Resume Next
    Set impObj = LLImporter.Create(Nothing)
    Assert.IsTrue Err.Number <> 0, "Factory should raise for a Nothing workbook"
    Err.Clear
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingWorkbook", Err.Number, Err.Description
End Sub


'@section The language check
'===============================================================================

'@sub-title The linelist answers the language it was built in.
'@TestMethod("LLImporter")
Public Sub TestCurrentLanguageReadsTheHiddenNameStore()
    CustomTestSetTitles Assert, TESTMODULE, "TestCurrentLanguageReadsTheHiddenNameStore"
    If Not FixtureReady("TestCurrentLanguageReadsTheHiddenNameStore") Then Exit Sub
    On Error GoTo TestFail

    Assert.AreEqual LANGUAGE_VALUE, FixtureImporter().CurrentLanguage, _
                    "The stored dictionary language is what the check compares against"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCurrentLanguageReadsTheHiddenNameStore", Err.Number, Err.Description
End Sub

'@sub-title A file in the same language is accepted.
'@TestMethod("LLImporter")
Public Sub TestTheSameLanguageIsAccepted()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheSameLanguageIsAccepted"
    If Not FixtureReady("TestTheSameLanguageIsAccepted") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue FixtureImporter().HasSameLanguage(ImportMetadataOfFile()), _
                  "The fixture file carries the language the linelist was built in"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSameLanguageIsAccepted", Err.Number, Err.Description
End Sub

'@sub-title A file with no metadata is refused, and the caller can tell it apart.
'@details
'"No Metadata sheet at all" and "a Metadata sheet naming no language" used to
'give the user the same message. They answer the same Boolean still; the
'metadata object is what tells the caller which of the two happened, so the
'three-way message can be built from it.
'@TestMethod("LLImporter")
Public Sub TestNoMetadataIsToldApartFromNoLanguage()
    CustomTestSetTitles Assert, TESTMODULE, "TestNoMetadataIsToldApartFromNoLanguage"
    If Not FixtureReady("TestNoMetadataIsToldApartFromNoLanguage") Then Exit Sub
    On Error GoTo TestFail

    Dim bareBook As Workbook
    Dim bareMeta As ImportMetadata

    Set bareBook = NewWorkbook()
    Set bareMeta = ImportMetadata.Create(bareBook)

    Assert.IsFalse FixtureImporter().HasSameLanguage(bareMeta), _
                   "A workbook with no Metadata sheet is not in our language"
    Assert.IsFalse bareMeta.Exists, _
                   "And the metadata object says the sheet is absent"
    Assert.IsTrue ImportMetadataOfFile().Exists, _
                  "While the real file does carry a Metadata sheet"

    DeleteWorkbook bareBook

    Exit Sub
TestFail:
    On Error Resume Next
    If Not bareBook Is Nothing Then DeleteWorkbook bareBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestNoMetadataIsToldApartFromNoLanguage", Err.Number, Err.Description
End Sub

'@sub-title A file in another language is refused, and both languages are known.
'@TestMethod("LLImporter")
Public Sub TestADifferentLanguageIsRefused()
    CustomTestSetTitles Assert, TESTMODULE, "TestADifferentLanguageIsRefused"
    If Not FixtureReady("TestADifferentLanguageIsRefused") Then Exit Sub
    On Error GoTo TestFail

    Dim otherBook As Workbook
    Dim otherMeta As ImportMetadata

    Set otherBook = NewWorkbook()
    WriteMetadataSheet otherBook, "French", "export for migrations", _
                       "header format: variable names | variable names: yes | header row: 1", _
                       EPIWEEK_IN_FILE
    Set otherMeta = ImportMetadata.Create(otherBook)

    Assert.IsFalse FixtureImporter().HasSameLanguage(otherMeta), _
                   "A file written in French does not match an English linelist"
    Assert.AreEqual "French", otherMeta.Language, _
                    "And the caller can name both languages to the user"

    DeleteWorkbook otherBook

    Exit Sub
TestFail:
    On Error Resume Next
    If Not otherBook Is Nothing Then DeleteWorkbook otherBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestADifferentLanguageIsRefused", Err.Number, Err.Description
End Sub


'@section Reading the file over before importing it
'===============================================================================

'@sub-title A complete file passes every check.
'@TestMethod("LLImporter")
Public Sub TestACompleteFileIsAccepted()
    CustomTestSetTitles Assert, TESTMODULE, "TestACompleteFileIsAccepted"
    If Not FixtureReady("TestACompleteFileIsAccepted") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue FixtureImporter().CheckImportFile(ImportWorkbook, ImportMetadataOfFile()), _
                  "A file carrying its variable names can be read"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestACompleteFileIsAccepted", Err.Number, Err.Description
End Sub

'@sub-title Every worksheet the file is missing is filed.
'@TestMethod("LLImporter")
Public Sub TestAFileMissingItsMetadataSheetsIsFiled()
    CustomTestSetTitles Assert, TESTMODULE, "TestAFileMissingItsMetadataSheetsIsFiled"
    If Not FixtureReady("TestAFileMissingItsMetadataSheetsIsFiled") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Dim bareBook As Workbook
    Dim canRead As Boolean

    Set bareBook = NewWorkbook()
    Set impObj = FixtureImporter()

    canRead = impObj.CheckImportFile(bareBook, ImportMetadata.Create(bareBook))

    Assert.IsTrue canRead, _
                  "A file with no metadata at all is still read, from row 1"
    Assert.IsTrue impObj.HasCheckings, _
                  "And every missing worksheet is filed"
    Assert.IsTrue impObj.CheckingValues.KeyExists("dictionary sheet"), _
                  "The missing Dictionary worksheet is named"
    Assert.IsTrue impObj.CheckingValues.KeyExists("choices sheet"), _
                  "The missing Choices worksheet is named"
    Assert.IsTrue impObj.CheckingValues.KeyExists("metadata sheet"), _
                  "The missing Metadata worksheet is named"

    DeleteWorkbook bareBook

    Exit Sub
TestFail:
    On Error Resume Next
    If Not bareBook Is Nothing Then DeleteWorkbook bareBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestAFileMissingItsMetadataSheetsIsFiled", Err.Number, Err.Description
End Sub

'@sub-title A file whose header row carries labels alone is refused.
'@details
'A custom export written with `header format = variables labels` holds the
'labels a user reads on row 1 and no variable name anywhere, so no column of it
'can be matched to a column of this linelist. Every row of such a file used to
'land in "not imported" with nothing said.
'@TestMethod("LLImporter")
Public Sub TestAFileWithNoVariableNamesIsRefused()
    CustomTestSetTitles Assert, TESTMODULE, "TestAFileWithNoVariableNamesIsRefused"
    If Not FixtureReady("TestAFileWithNoVariableNamesIsRefused") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Dim labelledBook As Workbook

    Set labelledBook = NewWorkbook()
    WriteMetadataSheet labelledBook, LANGUAGE_VALUE, "custom export: export 2", _
                       "header format: variables labels | variable names: no | " & _
                       "variable labels: yes | header row: 1", _
                       EPIWEEK_IN_FILE

    Set impObj = FixtureImporter()

    Assert.IsFalse impObj.CheckImportFile(labelledBook, ImportMetadata.Create(labelledBook)), _
                   "A file carrying no variable name cannot be read back"
    Assert.IsTrue impObj.CheckingValues.KeyExists("header row"), _
                  "And the reason is filed where the user reads it"

    DeleteWorkbook labelledBook

    Exit Sub
TestFail:
    On Error Resume Next
    If Not labelledBook Is Nothing Then DeleteWorkbook labelledBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestAFileWithNoVariableNamesIsRefused", Err.Number, Err.Description
End Sub


'@section Importing the data
'===============================================================================

'@sub-title Data lands in the table, and the sheet the linelist lacks is reported.
'@TestMethod("LLImporter")
Public Sub TestImportDataWritesTheRows()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportDataWritesTheRows"
    If Not FixtureReady("TestImportDataWritesTheRows") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Dim lo As ListObject
    Dim notImported As BetterArray

    ResetHListSheet

    Set impObj = FixtureImporter()
    impObj.ImportData ImportWorkbook, False, ImportMetadataOfFile()

    Set lo = FixtureWorkbook.Worksheets(HLIST_SHEET).ListObjects(HLIST_TABLE)

    Assert.IsNotNothing lo.DataBodyRange, "The table holds rows after the import"
    Assert.AreEqual FIRST_ROW_VALUE, _
                    CStr(lo.ListColumns(VAR_ONE).DataBodyRange.Cells(1, 1).Value), _
                    "The first data row of the file lands in the first column"
    Assert.AreEqual "11", _
                    CStr(lo.ListColumns(VAR_TWO).DataBodyRange.Cells(1, 1).Value), _
                    "And in the second"

    Set notImported = impObj.ReportSheets(ImportReportNotImported)
    Assert.IsTrue notImported.Includes(GHOST_SHEET), _
                  "A worksheet of the file the linelist has none of is reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestImportDataWritesTheRows", Err.Number, Err.Description
End Sub

'@sub-title A column of the file the linelist has no variable for is reported.
'@TestMethod("LLImporter")
Public Sub TestAColumnWithNoVariableIsReported()
    CustomTestSetTitles Assert, TESTMODULE, "TestAColumnWithNoVariableIsReported"
    If Not FixtureReady("TestAColumnWithNoVariableIsReported") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Dim notImported As BetterArray
    Dim counter As Long
    Dim found As Boolean
    Dim entry As Variant

    ResetHListSheet

    Set impObj = FixtureImporter()
    impObj.ImportData ImportWorkbook, False, ImportMetadataOfFile()

    Set notImported = impObj.ReportVariables(ImportReportNotImported)

    For counter = notImported.LowerBound To notImported.UpperBound
        entry = notImported.Item(counter)
        If CStr(entry(LBound(entry))) = VAR_ABSENT Then found = True
    Next counter

    Assert.IsTrue found, _
                  "The file's fourth column matches no column of the table and is named"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAColumnWithNoVariableIsReported", Err.Number, Err.Description
End Sub

'@sub-title Importing at the bottom keeps the rows already there.
'@details
'The fact the pasting question was built on. An import with pasteAtBottom False
'blanks the table first, so a user who answers "keep my data" has to reach the
'other branch. That question stopped being asked and False was passed as a
'literal, so three weeks of entered cases went with a colleague's file.
'@TestMethod("LLImporter")
Public Sub TestPasteAtBottomKeepsTheRowsAlreadyThere()
    CustomTestSetTitles Assert, TESTMODULE, "TestPasteAtBottomKeepsTheRowsAlreadyThere"
    If Not FixtureReady("TestPasteAtBottomKeepsTheRowsAlreadyThere") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Dim lo As ListObject
    Dim rowsAfterFirst As Long
    Dim rowsAfterSecond As Long

    ResetHListSheet

    Set impObj = FixtureImporter()
    impObj.ImportData ImportWorkbook, False, ImportMetadataOfFile()

    Set lo = FixtureWorkbook.Worksheets(HLIST_SHEET).ListObjects(HLIST_TABLE)
    rowsAfterFirst = lo.DataBodyRange.Rows.Count

    impObj.ImportData ImportWorkbook, True, ImportMetadataOfFile()

    Set lo = FixtureWorkbook.Worksheets(HLIST_SHEET).ListObjects(HLIST_TABLE)
    rowsAfterSecond = lo.DataBodyRange.Rows.Count

    Assert.IsTrue rowsAfterSecond > rowsAfterFirst, _
                  "Importing at the bottom adds rows under the ones already there"

    ResetHListSheet

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPasteAtBottomKeepsTheRowsAlreadyThere", Err.Number, Err.Description
End Sub

'@sub-title The linelist knows whether it holds data the user typed.
'@details
'HasData compares the filled cells of each row against blank_row_count, the
'CountA of an untouched row taken when the table was built. It had no caller at
'all while every import blanked the tables from row 1 with no question asked.
'@TestMethod("LLImporter")
Public Sub TestHasDataAnswersTheRowsTheUserTyped()
    CustomTestSetTitles Assert, TESTMODULE, "TestHasDataAnswersTheRowsTheUserTyped"
    If Not FixtureReady("TestHasDataAnswersTheRowsTheUserTyped") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim lo As ListObject

    ResetHListSheet

    Assert.IsFalse FixtureImporter().HasData, _
                   "A table of blank rows carries no user data"

    'The two facts the write depends on are asserted rather than assumed. A
    'raise here reaches the harness as the test name and nothing else, so what
    'the sheet was in gets stated while it can still be read.
    UnprotectHListSheet
    Set sh = FixtureWorkbook.Worksheets(HLIST_SHEET)
    Assert.IsFalse sh.ProtectContents, _
                   "The data entry sheet is open for writing"

    Set lo = sh.ListObjects(HLIST_TABLE)
    Assert.IsNotNothing lo.DataBodyRange, _
                        "And the table holds the one blank row a reset leaves"

    lo.ListColumns(1).DataBodyRange.Cells(1, 1).Value = "typed by hand"

    Assert.IsTrue FixtureImporter().HasData, _
                  "One filled cell past the blank baseline is user data"

    ResetHListSheet

    Exit Sub
TestFail:
    On Error Resume Next
    ResetHListSheet
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestHasDataAnswersTheRowsTheUserTyped", Err.Number, Err.Description
End Sub


'@section Importing the custom dropdowns
'===============================================================================

'@sub-title The values of a custom dropdown arrive.
'@details
'The single highest-value test of this module. This round trip has never once
'worked: the import cut five characters off the ListObject name, which leaves
'the workbook counter attached, and the export read the header cell, which is
'the name with every space turned into an underscore. The dropdown is called
'"contact type" on purpose, so either of the old shapes fails this test.
'@TestMethod("LLImporter")
Public Sub TestImportingACustomDropdownBringsItsValues()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportingACustomDropdownBringsItsValues"
    If Not FixtureReady("TestImportingACustomDropdownBringsItsValues") Then Exit Sub
    On Error GoTo TestFail

    Dim drop As DropdownLists
    Dim valuesBefore As BetterArray
    Dim valuesAfter As BetterArray

    ResetCustomDropdown

    Set drop = DropdownLists.Create(FixtureWorkbook.Worksheets(DROPDOWN_SHEET), vbNullString)
    Set valuesBefore = drop.Values(DROPDOWN_NAME)

    Assert.IsFalse valuesBefore.Includes(DROPDOWN_NEW_VALUE), _
                   "The linelist does not offer the new value before the import"

    FixtureImporter().ImportCustomDropdown ImportWorkbook, False

    Set drop = DropdownLists.Create(FixtureWorkbook.Worksheets(DROPDOWN_SHEET), vbNullString)
    Set valuesAfter = drop.Values(DROPDOWN_NAME)

    Assert.IsTrue valuesAfter.Includes(DROPDOWN_NEW_VALUE), _
                  "The value the file carries arrives in the dropdown"

    ResetCustomDropdown

    Exit Sub
TestFail:
    On Error Resume Next
    ResetCustomDropdown
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestImportingACustomDropdownBringsItsValues", Err.Number, Err.Description
End Sub

'@sub-title The key the two ends build is the registry name.
'@details
'The proof the export side and the import side agree. The choice the file
'carries answers under exactly the name DropdownLists.AllDropdowns holds.
'@TestMethod("LLImporter")
Public Sub TestTheChoiceKeyIsTheRegistryName()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheChoiceKeyIsTheRegistryName"
    If Not FixtureReady("TestTheChoiceKeyIsTheRegistryName") Then Exit Sub
    On Error GoTo TestFail

    Dim drop As DropdownLists
    Dim choi As LLChoices
    Dim registered As BetterArray
    Dim categories As BetterArray

    Set drop = DropdownLists.Create(FixtureWorkbook.Worksheets(DROPDOWN_SHEET), vbNullString)
    Set registered = drop.AllDropdowns

    Assert.IsTrue registered.Includes(DROPDOWN_NAME), _
                  "The registry holds the dropdown under the name with its space"

    Set choi = LLChoices.Create(ImportWorkbook.Worksheets(CHOICES_SHEET), 1, 1)
    Set categories = choi.Categories(ChoiceKeys.Create().CustomChoiceName(DROPDOWN_NAME))

    Assert.IsTrue categories.Length > 0, _
                  "And the file's choices answer under the key built from that name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheChoiceKeyIsTheRegistryName", Err.Number, Err.Description
End Sub

'@sub-title A file with no Choices worksheet leaves the dropdowns alone and says so.
'@TestMethod("LLImporter")
Public Sub TestAFileWithNoChoicesFilesAChecking()
    CustomTestSetTitles Assert, TESTMODULE, "TestAFileWithNoChoicesFilesAChecking"
    If Not FixtureReady("TestAFileWithNoChoicesFilesAChecking") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Dim bareBook As Workbook

    Set bareBook = NewWorkbook()
    Set impObj = FixtureImporter()

    impObj.ImportCustomDropdown bareBook, False

    Assert.IsTrue impObj.CheckingValues.KeyExists("custom dropdowns"), _
                  "A step that decided to do nothing records that it did"

    DeleteWorkbook bareBook

    Exit Sub
TestFail:
    On Error Resume Next
    If Not bareBook Is Nothing Then DeleteWorkbook bareBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestAFileWithNoChoicesFilesAChecking", Err.Number, Err.Description
End Sub


'@section Importing the single values
'===============================================================================

'@sub-title The epiweek start value lands in the hidden name store.
'@details
'The write used to go through Names("RNG_EpiWeekStart").RefersToRange, and a
'HiddenNames string is a quoted literal with no range behind it, so 1004 fell
'into the swallow above it and the value has never once landed.
'@TestMethod("LLImporter")
Public Sub TestTheEpiWeekValueLands()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheEpiWeekValueLands"
    If Not FixtureReady("TestTheEpiWeekValueLands") Then Exit Sub
    On Error GoTo TestFail

    Dim store As HiddenNames

    Set store = HiddenNames.Create(FixtureWorkbook)
    store.SetValue EPIWEEK_NAME, EPIWEEK_BEFORE

    Assert.AreEqual EPIWEEK_BEFORE, store.ValueAsString(EPIWEEK_NAME), _
                    "The linelist starts on the value it was built with"

    FixtureImporter().ImportSingleValues ImportMetadataOfFile()

    Set store = HiddenNames.Create(FixtureWorkbook)
    Assert.AreEqual EPIWEEK_IN_FILE, store.ValueAsString(EPIWEEK_NAME), _
                    "And carries the value the file held after the import"

    store.SetValue EPIWEEK_NAME, EPIWEEK_BEFORE

    Exit Sub
TestFail:
    On Error Resume Next
    HiddenNames.Create(FixtureWorkbook).SetValue EPIWEEK_NAME, EPIWEEK_BEFORE
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestTheEpiWeekValueLands", Err.Number, Err.Description
End Sub

'@sub-title The old write path could never have worked.
'@details
'States the fault as a test, so the shape is pinned rather than remembered. The
'mirror of TestLLExporter.TestTheEpiWeekNameHasNoRangeBehindIt, on the write
'side of the same round trip.
'@TestMethod("LLImporter")
Public Sub TestTheEpiWeekNameHasNoRangeToWriteInto()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheEpiWeekNameHasNoRangeToWriteInto"
    If Not FixtureReady("TestTheEpiWeekNameHasNoRangeToWriteInto") Then Exit Sub
    On Error GoTo TestFail

    Dim storedName As Excel.Name
    Dim rangeErrNumber As Long

    Set storedName = Nothing
    On Error Resume Next
    Set storedName = FixtureWorkbook.Names(EPIWEEK_NAME)
    On Error GoTo 0

    Assert.IsNotNothing storedName, "The workbook does carry the name"

    On Error Resume Next
    storedName.RefersToRange.Value = "4"
    rangeErrNumber = Err.Number
    Err.Clear
    On Error GoTo 0

    Assert.IsTrue rangeErrNumber <> 0, _
                  "Writing through a string-valued hidden name raises"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheEpiWeekNameHasNoRangeToWriteInto", Err.Number, Err.Description
End Sub

'@sub-title A file carrying no epiweek value leaves the linelist as it was.
'@TestMethod("LLImporter")
Public Sub TestNoEpiWeekValueLeavesTheStoreAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestNoEpiWeekValueLeavesTheStoreAlone"
    If Not FixtureReady("TestNoEpiWeekValueLeavesTheStoreAlone") Then Exit Sub
    On Error GoTo TestFail

    Dim bareBook As Workbook
    Dim store As HiddenNames

    Set store = HiddenNames.Create(FixtureWorkbook)
    store.SetValue EPIWEEK_NAME, EPIWEEK_BEFORE

    Set bareBook = NewWorkbook()
    FixtureImporter().ImportSingleValues ImportMetadata.Create(bareBook)

    Set store = HiddenNames.Create(FixtureWorkbook)
    Assert.AreEqual EPIWEEK_BEFORE, store.ValueAsString(EPIWEEK_NAME), _
                    "A file naming no value changes nothing"

    DeleteWorkbook bareBook

    Exit Sub
TestFail:
    On Error Resume Next
    If Not bareBook Is Nothing Then DeleteWorkbook bareBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestNoEpiWeekValueLeavesTheStoreAlone", Err.Number, Err.Description
End Sub


'@section The geobase
'===============================================================================

'@sub-title A geobase import lands the rows and the new level labels.
'@details
'ImportGeobase had no test of its own, and the stale level captions reported
'from the field came through this path. What the import owes the linelist is
'both halves of the geobase: the rows of the admin tables, and the five
'workbook-scoped level labels every reader of the class asks for.
'
'The admin1 value is read out of the data body by position. The import ends in
'a Translate, so by the time the assertion runs the header of that column
'carries the level label of the file.
'@TestMethod("LLImporter")
Public Sub TestImportGeobaseLandsTheRowsAndTheLabels()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportGeobaseLandsTheRowsAndTheLabels"
    If Not FixtureReady("TestImportGeobaseLandsTheRowsAndTheLabels") Then Exit Sub
    On Error GoTo TestFail

    Dim geoSheet As Worksheet
    Dim landed As LLGeo
    Dim store As HiddenNames

    Set geoSheet = SeedLinelistGeobase()

    FixtureImporter().ImportGeobase GeobaseWorkbook
    ResetGeoVariableLabel

    Assert.AreEqual GEOBASE_ADM1_ROWS, _
                    geoSheet.ListObjects(GEO_ADM1_TABLE).ListRows.Count, _
                    "The admin1 table holds the rows of the file"
    Assert.AreEqual GEOBASE_ADM2_ROWS, _
                    geoSheet.ListObjects(GEO_ADM2_TABLE).ListRows.Count, _
                    "And so does the admin2 table"
    Assert.AreEqual ADM1_FIRST_VALUE, _
                    CStr(geoSheet.ListObjects(GEO_ADM1_TABLE).DataBodyRange.Cells(1, 1).Value), _
                    "The first admin1 value of the file is the first row of the table"

    Set store = HiddenNames.Create(FixtureWorkbook)
    Assert.AreEqual ADM1_LABEL_IN_FILE, store.ValueAsString(ADM1_LEVEL_NAME), _
                    "The admin1 label of the file reaches the workbook store"

    Set landed = LLGeo.Create(geoSheet)
    Assert.AreEqual ADM1_LABEL_IN_FILE, landed.GeoNames("adm1_name"), _
                    "A manager built after the import reads the admin1 label of the file"
    Assert.AreEqual ADM2_LABEL_IN_FILE, landed.GeoNames("adm2_name"), _
                    "And the admin2 label with it"

    Exit Sub
TestFail:
    On Error Resume Next
    ResetGeoVariableLabel
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestImportGeobaseLandsTheRowsAndTheLabels", Err.Number, Err.Description
End Sub

'@sub-title A manager read before the import still answers the old label.
'@details
'This is the contract every caller of a geobase import owes, and it is the half
'of the wiring no headless run can reach: HandleImportGeobase is private to a
'form module that carries no registry row, and it follows the import with
'ResetEventCaches on both its paths.
'
'LLGeo reads the five level labels once per instance, and the import runs
'through an instance of its own. A manager held from before the import
'therefore keeps the label of the geobase that has gone, and only a caller that
'drops it reads the new one.
'@TestMethod("LLImporter")
Public Sub TestAManagerHeldAcrossTheImportKeepsItsOldLabel()
    CustomTestSetTitles Assert, TESTMODULE, "TestAManagerHeldAcrossTheImportKeepsItsOldLabel"
    If Not FixtureReady("TestAManagerHeldAcrossTheImportKeepsItsOldLabel") Then Exit Sub
    On Error GoTo TestFail

    Dim geoSheet As Worksheet
    Dim held As LLGeo
    Dim beforeImport As String

    Set geoSheet = SeedLinelistGeobase()

    'The read is what loads the label cache of the held manager, which is the
    'state the import then moves the store under.
    Set held = LLGeo.Create(geoSheet)
    held.Translate rawNames:=False
    beforeImport = held.GeoNames("adm1_name")

    FixtureImporter().ImportGeobase GeobaseWorkbook
    ResetGeoVariableLabel

    Assert.AreEqual ADM1_LABEL_BEFORE, beforeImport, _
                    "The manager cached the label the geobase carried at the start"
    Assert.AreEqual ADM1_LABEL_BEFORE, held.GeoNames("adm1_name"), _
                    "It still answers that label once the import has moved the store"
    Assert.AreEqual ADM1_LABEL_IN_FILE, LLGeo.Create(geoSheet).GeoNames("adm1_name"), _
                    "A manager built after the import answers the label of the file"

    Exit Sub
TestFail:
    On Error Resume Next
    ResetGeoVariableLabel
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestAManagerHeldAcrossTheImportKeepsItsOldLabel", Err.Number, Err.Description
End Sub


'@section Comparing the two files
'===============================================================================

'@sub-title A label the two files disagree about is filed.
'@TestMethod("LLImporter")
Public Sub TestADifferingLabelIsFiled()
    CustomTestSetTitles Assert, TESTMODULE, "TestADifferingLabelIsFiled"
    If Not FixtureReady("TestADifferingLabelIsFiled") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter

    ResetHListSheet

    Set impObj = FixtureImporter()
    impObj.ImportData ImportWorkbook, False, ImportMetadataOfFile()
    impObj.CompareWithImportFile ImportWorkbook

    Assert.IsTrue impObj.CheckingValues.KeyExists("label of " & VAR_TWO), _
                  "The variable whose label the file rewrote is named"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestADifferingLabelIsFiled", Err.Number, Err.Description
End Sub

'@sub-title A choice list the two files disagree about is filed.
'@details
'The check that matters most. A data validation dropdown does not block a
'programmatic write, so a value from the file's vocabulary lands in the cell and
'stays there: the cell shows something no dropdown of this linelist can produce,
'the next hand edit of it is rejected, and every analysis grouping by that
'variable gains a category the setup never declared.
'@TestMethod("LLImporter")
Public Sub TestADifferingChoiceListIsFiled()
    CustomTestSetTitles Assert, TESTMODULE, "TestADifferingChoiceListIsFiled"
    If Not FixtureReady("TestADifferingChoiceListIsFiled") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter

    ResetHListSheet

    Set impObj = FixtureImporter()
    impObj.ImportData ImportWorkbook, False, ImportMetadataOfFile()
    impObj.CompareWithImportFile ImportWorkbook

    Assert.IsTrue impObj.CheckingValues.KeyExists("choices of " & VAR_THREE), _
                  "A choice variable whose list the file disagrees about is named"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestADifferingChoiceListIsFiled", Err.Number, Err.Description
End Sub

'@sub-title A file with no dictionary is compared against nothing and says so.
'@TestMethod("LLImporter")
Public Sub TestComparingAgainstNoDictionaryDoesNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestComparingAgainstNoDictionaryDoesNothing"
    If Not FixtureReady("TestComparingAgainstNoDictionaryDoesNothing") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Dim bareBook As Workbook

    Set bareBook = NewWorkbook()
    Set impObj = FixtureImporter()

    impObj.CompareWithImportFile bareBook

    Assert.IsFalse impObj.HasCheckings, _
                   "There is nothing to compare and nothing to say about it"

    DeleteWorkbook bareBook

    Exit Sub
TestFail:
    On Error Resume Next
    If Not bareBook Is Nothing Then DeleteWorkbook bareBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestComparingAgainstNoDictionaryDoesNothing", Err.Number, Err.Description
End Sub


'@section The report
'===============================================================================

'@sub-title NeedReport is False before anything runs.
'@TestMethod("LLImporter")
Public Sub NeedReportFalseByDefault()
    CustomTestSetTitles Assert, TESTMODULE, "NeedReportFalseByDefault"
    On Error GoTo TestFail

    Assert.IsFalse LLImporter.Create(ThisWorkbook).NeedReport, _
                   "NeedReport should be False before any import"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "NeedReportFalseByDefault", Err.Number, Err.Description
End Sub

'@sub-title ClearReport resets everything, the diagnostics included.
'@TestMethod("LLImporter")
Public Sub TestClearReportResetsTheDiagnosticsToo()
    CustomTestSetTitles Assert, TESTMODULE, "TestClearReportResetsTheDiagnosticsToo"
    If Not FixtureReady("TestClearReportResetsTheDiagnosticsToo") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Dim bareBook As Workbook

    Set bareBook = NewWorkbook()
    Set impObj = FixtureImporter()

    impObj.CheckImportFile bareBook, ImportMetadata.Create(bareBook)
    Assert.IsTrue impObj.HasCheckings, "The check did file something"

    impObj.ClearReport

    Assert.IsFalse impObj.NeedReport, "NeedReport is False after ClearReport"
    Assert.IsFalse impObj.HasCheckings, "And the diagnostics go with it"
    Assert.AreEqual CLng(0), impObj.ReportSheets(ImportReportNotImported).Length, _
                    "And the report lists are empty"

    DeleteWorkbook bareBook

    Exit Sub
TestFail:
    On Error Resume Next
    If Not bareBook Is Nothing Then DeleteWorkbook bareBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestClearReportResetsTheDiagnosticsToo", Err.Number, Err.Description
End Sub

'@sub-title An unknown scope answers an empty list on both sides.
'@TestMethod("LLImporter")
Public Sub ReportSheetsInvalidScopeReturnsEmpty()
    CustomTestSetTitles Assert, TESTMODULE, "ReportSheetsInvalidScopeReturnsEmpty"
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Set impObj = LLImporter.Create(ThisWorkbook)

    Assert.AreEqual CLng(0), impObj.ReportSheets(99).Length, _
                    "An unknown scope answers an empty list"
    Assert.AreEqual CLng(0), impObj.ReportVariables(99).Length, _
                    "And so does the variable side"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "ReportSheetsInvalidScopeReturnsEmpty", Err.Number, Err.Description
End Sub

'@sub-title The report is kept in the workbook and survives the import object.
'@details
'The four lists used to live in memory and die with the import object, while the
'four ListObjects they belong on sat empty. The store is what lets F_ImportRep
'open the last report days later. This test reads them back through a second
'store, so what is asserted is what reached the worksheet.
'@TestMethod("LLImporter")
Public Sub TestFinalizeReportKeepsTheReportInTheWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "TestFinalizeReportKeepsTheReportInTheWorkbook"
    If Not FixtureReady("TestFinalizeReportKeepsTheReportInTheWorkbook") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter
    Dim store As ImportReport

    ResetHListSheet

    Set impObj = FixtureImporter()
    impObj.ImportData ImportWorkbook, False, ImportMetadataOfFile()
    impObj.FinalizeReport

    Set store = ImportReport.Create(FixtureWorkbook)

    Assert.IsTrue store.SheetNames(ImportReportNotTouched).Length > 0, _
                  "The sheets that took no data are on the worksheet after the import"
    Assert.IsTrue store.SheetNames(ImportReportNotImported).Includes(GHOST_SHEET), _
                  "And so is the sheet of the file the linelist has none of"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFinalizeReportKeepsTheReportInTheWorkbook", Err.Number, Err.Description
End Sub

'@sub-title The four report lists reach the checkings the user reads.
'@TestMethod("LLImporter")
Public Sub TestTheReportListsAreFiledAsCheckings()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheReportListsAreFiledAsCheckings"
    If Not FixtureReady("TestTheReportListsAreFiledAsCheckings") Then Exit Sub
    On Error GoTo TestFail

    Dim impObj As LLImporter

    ResetHListSheet

    Set impObj = FixtureImporter()
    impObj.ImportData ImportWorkbook, False, ImportMetadataOfFile()
    impObj.FinalizeReport

    Assert.IsTrue impObj.HasCheckings, _
                  "Everything the import found reaches one place"
    Assert.IsTrue impObj.CheckingValues.KeyExists("sheet not imported: " & GHOST_SHEET), _
                  "The sheet the linelist has none of is one of them"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheReportListsAreFiledAsCheckings", Err.Number, Err.Description
End Sub


'@section Fixture helpers
'===============================================================================

'@fun-title Report a fixture that could not be built as this test's failure.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is usable.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If SetupError = 0 Then
        If Not FixtureWorkbook Is Nothing Then
            If Not ImportWorkbook Is Nothing Then
                If Not GeobaseWorkbook Is Nothing Then
                    FixtureReady = True
                    Exit Function
                End If
            End If
        End If
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The fixture could not be built - " & SetupMessage
End Function

'@fun-title An importer over the linelist fixture.
'@return LLImporter. A fresh instance, which costs nothing to build.
Private Function FixtureImporter() As LLImporter
    Set FixtureImporter = LLImporter.Create(FixtureWorkbook)
End Function

'@fun-title The metadata of the import fixture.
'@return ImportMetadata. A fresh reader over the file's Metadata sheet.
Private Function ImportMetadataOfFile() As ImportMetadata
    Set ImportMetadataOfFile = ImportMetadata.Create(ImportWorkbook)
End Function


'@section Fixture helpers - the linelist
'===============================================================================

'@sub-title Build every worksheet the linelist fixture carries.
Private Sub BuildLinelistFixture()

    Dim store As HiddenNames

    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, FixtureWorkbook
    PasswordsTestFixture.PreparePasswordsFixture PASSWORD_SHEET, FixtureWorkbook

    BuildLinelistChoices
    BuildHListSheet
    BuildCustomDropdown

    Set store = HiddenNames.Create(FixtureWorkbook)
    store.EnsureName LANGUAGE_NAME, LANGUAGE_VALUE, HiddenNameTypeString
    store.SetValue LANGUAGE_NAME, LANGUAGE_VALUE
    store.EnsureName CUSTOMDROP_NAME, DROPDOWN_SHEET, HiddenNameTypeString
    store.SetValue CUSTOMDROP_NAME, DROPDOWN_SHEET
    store.EnsureName EPIWEEK_NAME, EPIWEEK_BEFORE, HiddenNameTypeString
    store.SetValue EPIWEEK_NAME, EPIWEEK_BEFORE
End Sub

'@sub-title Write the Choices worksheet of the linelist.
'@details
'choi_h2 is a choice_manual bound to list_correct_order in the dictionary
'fixture, and the file offers one value for that list this one does not, which is
'what the choice comparison has to find.
Private Sub BuildLinelistChoices()
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(CHOICES_SHEET, FixtureWorkbook, clearSheet:=True)

    WriteRow sh.Cells(1, 1), "list name", "label"
    WriteRow sh.Cells(2, 1), "list_correct_order", "A"
    WriteRow sh.Cells(3, 1), "list_correct_order", "B"
    WriteRow sh.Cells(4, 1), "list_correct_order", "C"
End Sub

'@sub-title Write the HList data entry worksheet of the linelist.
'@details
'A three-column ListObject over real variable names of the dictionary fixture,
'with the three worksheet-level names an import reads: sheet_type, table_name
'and blank_row_count. The last of them is 0, so one filled cell of a row counts
'as data the user typed.
Private Sub BuildHListSheet()
    Dim sh As Worksheet
    Dim store As HiddenNames

    Set sh = EnsureWorksheet(HLIST_SHEET, FixtureWorkbook, clearSheet:=True)

    WriteRow sh.Cells(1, 1), VAR_ONE, VAR_TWO, VAR_THREE
    sh.ListObjects.Add(xlSrcRange, sh.Range(sh.Cells(1, 1), sh.Cells(2, 3)), , xlYes) _
      .Name = HLIST_TABLE

    Set store = HiddenNames.Create(sh)
    store.EnsureName "sheet_type", "HList", HiddenNameTypeString
    store.SetValue "sheet_type", "HList"
    store.EnsureName "table_name", HLIST_TABLE, HiddenNameTypeString
    store.SetValue "table_name", HLIST_TABLE
    store.EnsureName "blank_row_count", 0, HiddenNameTypeLong
    store.SetValue "blank_row_count", 0
End Sub

'@sub-title Empty the HList table back to one blank row.
'@details
'Called by every test that writes into it, so the order the tests run in decides
'nothing.
Private Sub ResetHListSheet()
    Dim lo As ListObject
    Dim sh As Worksheet

    Set sh = FixtureWorkbook.Worksheets(HLIST_SHEET)

    UnprotectHListSheet

    Set lo = sh.ListObjects(HLIST_TABLE)

    On Error Resume Next
    lo.DataBodyRange.ClearContents
    On Error GoTo 0

    lo.Resize sh.Range(sh.Cells(1, 1), sh.Cells(2, 3))

    'CustomTable.Import with pasteAtBottom False cleans the table first, and a
    'table it has emptied can come back with no data body at all. A resize onto
    'two rows does not always put one back, so the row is added when it is
    'missing and the reset always leaves the same shape.
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

'@sub-title Open the HList worksheet for writing.
'@details
'ImportData protects the sheet on its way out, and a protected sheet takes
'neither a ClearContents nor a Resize nor a cell write. Both routes are tried:
'the Passwords object is what protected it, and the raw Unprotect covers the
'case where the passwords fixture itself could not be built.
Private Sub UnprotectHListSheet()
    Dim sh As Worksheet
    Dim pass As Passwords

    Set sh = FixtureWorkbook.Worksheets(HLIST_SHEET)

    On Error Resume Next
        Set pass = Passwords.Create(FixtureWorkbook.Worksheets(PASSWORD_SHEET))
        If Not pass Is Nothing Then pass.UnProtect sh.Name
        If sh.ProtectContents Then sh.Unprotect
    On Error GoTo 0
End Sub

'@sub-title Build the custom dropdown worksheet and register one dropdown.
Private Sub BuildCustomDropdown()
    Dim sh As Worksheet
    Dim drop As DropdownLists

    Set sh = EnsureWorksheet(DROPDOWN_SHEET, FixtureWorkbook, clearSheet:=True)

    Set drop = DropdownLists.Create(sh, vbNullString)
    drop.Add StartingDropdownValues(), DROPDOWN_NAME
End Sub

'@sub-title Put the custom dropdown back to the two values it starts with.
Private Sub ResetCustomDropdown()
    Dim drop As DropdownLists

    Set drop = DropdownLists.Create(FixtureWorkbook.Worksheets(DROPDOWN_SHEET), vbNullString)
    drop.Update StartingDropdownValues(), DROPDOWN_NAME, pasteAtBottom:=False
End Sub

'@fun-title The two values the custom dropdown starts and is put back to.
'@return BetterArray. The two values.
Private Function StartingDropdownValues() As BetterArray
    Dim values As BetterArray

    Set values = New BetterArray
    values.LowerBound = 1
    values.Push "family", "colleague"

    Set StartingDropdownValues = values
End Function


'@section Fixture helpers - the file being imported
'===============================================================================

'@sub-title Build every worksheet the import fixture carries.
Private Sub BuildImportFixture()

    WriteMetadataSheet ImportWorkbook, LANGUAGE_VALUE, "export for migrations", _
                       "header format: variable names | variable names: yes | " & _
                       "variable labels: no | header row: 1 | filtered data: no", _
                       EPIWEEK_IN_FILE

    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, ImportWorkbook
    MarkImportDictionaryPrepared
    RewriteOneLabel

    BuildImportChoices
    BuildImportDataSheet
    BuildGhostSheet
End Sub

'@sub-title Write a Metadata worksheet onto a workbook.
'@param wb Workbook. The workbook to write into.
'@param language String. The language tag.
'@param exportType String. The export type tag.
'@param options String. The export options line.
'@param epiWeek String. The epiweek start tag.
Private Sub WriteMetadataSheet(ByVal wb As Workbook, ByVal language As String, _
                               ByVal exportType As String, ByVal options As String, _
                               ByVal epiWeek As String)
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(METADATA_SHEET, wb, clearSheet:=True)

    WriteRow sh.Cells(1, 1), "variable", "value"
    WriteRow sh.Cells(2, 1), "language", language
    WriteRow sh.Cells(3, 1), "export_type", exportType
    WriteRow sh.Cells(4, 1), "export_options", options
    WriteRow sh.Cells(5, 1), "epiweekstart", epiWeek
    WriteRow sh.Cells(6, 1), "show_hide", "update on import"
    WriteRow sh.Cells(7, 1), "editable_labels", "update on import"
End Sub

'@sub-title Make the import dictionary read as prepared.
'@details
'CompareWithImportFile exits on a dictionary that was never prepared, and the
'shared fixture is the raw setup shape. LLdictionary.Prepared asks for six
'columns a preparation adds and for the cell under the last data row to be blue,
'so both are written here.
Private Sub MarkImportDictionaryPrepared()
    Dim sh As Worksheet
    Dim headers As Variant
    Dim firstFreeColumn As Long
    Dim lastRow As Long
    Dim counter As Long

    Set sh = ImportWorkbook.Worksheets(DICTIONARY_SHEET)

    headers = Array("table name", "column index", "visibility", _
                    "crf index", "crf choices", "crf status")
    firstFreeColumn = DictionaryTestFixture.DictionaryFixtureColumnCount() + 1

    For counter = LBound(headers) To UBound(headers)
        sh.Cells(1, firstFreeColumn + counter - LBound(headers)).Value = CStr(headers(counter))
    Next counter

    lastRow = DictionaryTestFixture.DictionaryFixtureRowCount() + 1
    sh.Cells(lastRow + 1, 1).Font.Color = vbBlue
End Sub

'@sub-title Give one variable of the import dictionary a different main label.
Private Sub RewriteOneLabel()
    Dim sh As Worksheet
    Dim labelColumn As Long
    Dim nameColumn As Long
    Dim counter As Long
    Dim lastRow As Long

    Set sh = ImportWorkbook.Worksheets(DICTIONARY_SHEET)
    nameColumn = DictionaryTestFixture.DictionaryHeaderIndex("Variable Name") + 1
    labelColumn = DictionaryTestFixture.DictionaryHeaderIndex("Main Label") + 1
    lastRow = DictionaryTestFixture.DictionaryFixtureRowCount() + 1

    For counter = 2 To lastRow
        If CStr(sh.Cells(counter, nameColumn).Value) = VAR_TWO Then
            sh.Cells(counter, labelColumn).Value = RELABELLED_IN_FILE
            Exit For
        End If
    Next counter
End Sub

'@sub-title Write the Choices worksheet of the file.
'@details
'It carries the custom dropdown under the key ChoiceKeys builds, with one value
'the linelist does not hold, and a third value for list_uni so the choice
'comparison has a difference to find.
Private Sub BuildImportChoices()
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(CHOICES_SHEET, ImportWorkbook, clearSheet:=True)

    WriteRow sh.Cells(1, 1), "list name", "label"
    WriteRow sh.Cells(2, 1), "list_correct_order", "A"
    WriteRow sh.Cells(3, 1), "list_correct_order", "B"
    WriteRow sh.Cells(4, 1), "list_correct_order", "C"
    WriteRow sh.Cells(5, 1), "list_correct_order", "D"
    WriteRow sh.Cells(6, 1), "__choice_custom_" & DROPDOWN_NAME, "family"
    WriteRow sh.Cells(7, 1), "__choice_custom_" & DROPDOWN_NAME, "colleague"
    WriteRow sh.Cells(8, 1), "__choice_custom_" & DROPDOWN_NAME, DROPDOWN_NEW_VALUE
End Sub

'@sub-title Write the data worksheet of the file.
'@details
'Three columns the linelist knows and one it does not, so the import has both a
'match to make and a column to report.
Private Sub BuildImportDataSheet()
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(HLIST_SHEET, ImportWorkbook, clearSheet:=True)

    WriteRow sh.Cells(1, 1), VAR_ONE, VAR_TWO, VAR_THREE, VAR_ABSENT
    WriteRow sh.Cells(2, 1), FIRST_ROW_VALUE, 11, "A", "dropped"
    WriteRow sh.Cells(3, 1), "row-two", 12, "B", "dropped"
End Sub

'@sub-title Write a worksheet the linelist has no match for.
Private Sub BuildGhostSheet()
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(GHOST_SHEET, ImportWorkbook, clearSheet:=True)
    WriteRow sh.Cells(1, 1), "some_variable", "another_variable"
    WriteRow sh.Cells(2, 1), "a", "b"
End Sub


'@section Fixture helpers - the geobase
'===============================================================================

'@fun-title Build the geobase worksheet of the linelist again.
'@details
'PrepareGeoFixture clears the worksheet, which drops the sheet-scoped
'translated flag with it, and it drops the five workbook-scoped level labels by
'hand. Both geo tests start from Province whatever ran before them.
'
'LLGeo.Create refuses a worksheet missing one of its nine tables, so the
'fixture is built whole, and it carries data because the class reads T_NAMES to
'resolve the labels.
'@return Worksheet. The geobase worksheet of the linelist fixture.
Private Function SeedLinelistGeobase() As Worksheet
    Set SeedLinelistGeobase = GeoTestFixture.PrepareGeoFixture(GEO_SHEET, _
                                                               FixtureWorkbook, _
                                                               withData:=True)
End Function

'@sub-title Write the three worksheets the geobase file carries.
'@details
'LLGeo.Import reads a source worksheet only when its name is one of the table
'names it knows, and it matches the columns of that sheet to the columns of the
'table by header. The headers are the raw level names, because Import reverts
'the geobase headers before it copies anything in.
'
'The file names its own admin1 and admin2 levels, which is what makes the two
'tests able to tell the labels of the file apart from the labels of the
'fixture. It carries no METADATA sheet, so the language code of the linelist
'stays EN and the NAMES column read is the EN one.
Private Sub BuildGeobaseFixture()
    Dim sh As Worksheet

    Set sh = EnsureWorksheet("ADM1", GeobaseWorkbook, clearSheet:=True)
    WriteRow sh.Cells(1, 1), "adm1_name"
    WriteRow sh.Cells(2, 1), ADM1_FIRST_VALUE
    WriteRow sh.Cells(3, 1), "Beta"

    Set sh = EnsureWorksheet("ADM2", GeobaseWorkbook, clearSheet:=True)
    WriteRow sh.Cells(1, 1), "adm1_name", "adm2_name"
    WriteRow sh.Cells(2, 1), ADM1_FIRST_VALUE, "Alpha North"
    WriteRow sh.Cells(3, 1), ADM1_FIRST_VALUE, "Alpha South"
    WriteRow sh.Cells(4, 1), "Beta", "Beta North"
    WriteRow sh.Cells(5, 1), "Beta", "Beta South"

    Set sh = EnsureWorksheet("NAMES", GeobaseWorkbook, clearSheet:=True)
    WriteRow sh.Cells(1, 1), "level", "EN"
    WriteRow sh.Cells(2, 1), "adm1_name", ADM1_LABEL_IN_FILE
    WriteRow sh.Cells(3, 1), "adm2_name", ADM2_LABEL_IN_FILE
    WriteRow sh.Cells(4, 1), "adm3_name", "Commune"
    WriteRow sh.Cells(5, 1), "adm4_name", "Village"
    WriteRow sh.Cells(6, 1), "hf_name", "Health Post"
End Sub

'@sub-title Put the dictionary label a geobase import overwrites back.
'@details
'ImportGeobase ends in LLGeo.UpdateDict, which writes the level label of the
'geobase over the main label of every geo-controlled variable of the
'dictionary. The fixture carries one, hf_h2, and the dictionary is shared with
'every other test of this module, so the cell goes back the way ResetHListSheet
'puts the data table back.
Private Sub ResetGeoVariableLabel()
    Dim sh As Worksheet
    Dim nameColumn As Long
    Dim labelColumn As Long
    Dim lastRow As Long
    Dim counter As Long

    Set sh = FixtureWorkbook.Worksheets(DICTIONARY_SHEET)
    nameColumn = DictionaryTestFixture.DictionaryHeaderIndex("Variable Name") + 1
    labelColumn = DictionaryTestFixture.DictionaryHeaderIndex("Main Label") + 1
    lastRow = DictionaryTestFixture.DictionaryFixtureRowCount() + 1

    For counter = 2 To lastRow
        If CStr(sh.Cells(counter, nameColumn).Value) = HF_VARIABLE Then
            sh.Cells(counter, labelColumn).Value = HF_VARIABLE_LABEL
            Exit For
        End If
    Next counter
End Sub
