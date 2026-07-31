Attribute VB_Name = "TestLLSheets"
Attribute VB_Description = "Tests for the LLSheets class"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@Folder("CustomTests")

'@ModuleDescription("Tests for the LLSheets class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'Validates the core behaviour of the LLSheets class, which provides
'sheet-level metadata derived from the linelist dictionary. Tests cover
'factory creation guard clauses, sheet containment checks, row-index
'lookups, data-bounds validation, sheet-info error paths, control
'detection, variable-count guards, and variable-address preparation
'requirements. Each test builds an LLSheets instance from a dictionary
'fixture and exercises one public method or error condition.
'@depends LLSheets, LLdictionary, CustomTest, DictionaryTestFixture, TestHelpersLite

Private Const DICT_SHEET As String = "LLSheetsDict"
Private Const SHEET_VERTICAL As String = "vlist1D-sheet1"
Private Const SHEET_HORIZONTAL As String = "hlist2D-sheet1"
Private Const KNOWN_VARIABLE As String = "choi_v1"

Private Assert As CustomTest
Private Dictionary As LLdictionary
Private Sheets As LLSheets

'@section Fixture Lifecycle
'===============================================================================

'@sub-title Reset the dictionary fixture worksheet to a known state
Private Sub ResetDictionarySheet()
    PrepareDictionaryFixture DICT_SHEET
End Sub

'@section Module Lifecycle
'===============================================================================

'@sub-title Initialise the test module and prepare shared resources
'@details
'Public, so the headless runner reaches it: the runner calls every lifecycle
'hook through Application.Run, which cannot see a Private Sub. The handler
'matters as much. An error escaping a lifecycle hook aborts the WHOLE module,
'so a fixture problem here would drop every test and the run would report no
'failures, because nothing ran.
'@ModuleInitialize
Public Sub ModuleInitialize()
    On Error GoTo Fail
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLSheets"
    ResetDictionarySheet
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "ModuleInitialize", Err.Number, Err.Description
End Sub

'@sub-title Tear down module-level resources after all tests complete
'@details
'Every step before RestoreApp is wrapped, because RestoreApp has to run
'whatever happened above: the hooks here call BusyApp, which puts Excel in
'manual calculation, and the next module in the run would inherit that.
'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        DeleteWorksheet DICT_SHEET
    On Error GoTo 0

    RestoreApp

    Set Sheets = Nothing
    Set Dictionary = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Create fresh Dictionary and Sheets instances before each test
'@details
'Resets the dictionary fixture worksheet, constructs a new LLdictionary
'from the fixture, and wraps it in a new LLSheets instance. This ensures
'every test starts with an unmodified dictionary and a cleanly initialised
'Sheets object so that tests remain independent of one another.
'@TestInitialize
Public Sub TestInitialize()
    On Error GoTo Fail
    BusyApp
    ResetDictionarySheet
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    Set Sheets = LLSheets.Create(Dictionary)
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInitialize", Err.Number, Err.Description
End Sub

'@sub-title Release per-test objects and flush assertion state
'@details
'Flushes any buffered assertion results to the output sheet and releases
'the Sheets and Dictionary references. Runs after each individual test
'method completes.
'@TestCleanup
Public Sub TestCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.Flush
        End If
    On Error GoTo 0

    Set Sheets = Nothing
    Set Dictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@sub-title Verify that Create raises when given a Nothing dictionary
'@details
'Arranges by calling LLSheets.Create with Nothing as the dictionary
'argument. Acts by attempting the creation, which should raise an error.
'Asserts that the error number equals ProjectError.ObjectNotInitialized,
'confirming the factory method guards against null dictionary input.
'@TestMethod("LLSheets")
Public Sub TestCreateRejectsNullDictionary()
    CustomTestSetTitles Assert, "LLSheets", "TestCreateRejectsNullDictionary"
    On Error GoTo ExpectError

    Dim invalid As LLSheets
    '@Ignore AssignmentNotUsed
    Set invalid = LLSheets.Create(Nothing)
    Assert.LogFailure "Create should raise when dictionary is Nothing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "Create should flag missing dictionary as ObjectNotInitialized"
    Err.Clear
End Sub

'@sub-title Verify that Contains detects known fixture sheets and rejects unknown names
'@details
'Arranges by using the module-level Sheets instance built from the
'dictionary fixture. Acts by calling Contains with the vertical and
'horizontal fixture sheet names, plus a non-existent name. Asserts that
'Contains returns True for both known sheets and False for the unknown
'sheet name, confirming the lookup works correctly for both matches and
'misses.
'@TestMethod("LLSheets")
Public Sub TestContainsRecognisesFixtureSheets()
    CustomTestSetTitles Assert, "LLSheets", "TestContainsRecognisesFixtureSheets"
    On Error GoTo Fail

    Assert.IsTrue Sheets.Contains(SHEET_VERTICAL), "Expected fixture sheet to be present"
    Assert.IsTrue Sheets.Contains(SHEET_HORIZONTAL), "Expected horizontal fixture sheet to be present"
    Assert.IsFalse Sheets.Contains("missing-sheet"), "Contains should return False for unknown sheet"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestContainsRecognisesFixtureSheets", Err.Number, Err.Description
End Sub

'@sub-title Verify that a sheet name is matched whatever case it is asked in
'@details
'Arranges by upper-casing and lower-casing the two fixture sheet names.
'Acts by calling Contains and RowIndex with those spellings. Asserts that
'both answer as they do for the stored spelling. Excel worksheet names are
'case-insensitive, so a sheet stored as "vlist1D-sheet1" and asked for as
'"VLIST1D-SHEET1" is the same sheet.
'@TestMethod("LLSheets")
Public Sub TestSheetLookupIgnoresCase()
    CustomTestSetTitles Assert, "LLSheets", "TestSheetLookupIgnoresCase"
    On Error GoTo Fail

    Dim storedRow As Long

    storedRow = Sheets.RowIndex(SHEET_VERTICAL)

    Assert.IsTrue Sheets.Contains(UCase$(SHEET_VERTICAL)), _
                  "Contains should answer True for the upper-cased sheet name"
    Assert.IsTrue Sheets.Contains(LCase$(SHEET_HORIZONTAL)), _
                  "Contains should answer True for the lower-cased sheet name"
    Assert.AreEqual storedRow, Sheets.RowIndex(UCase$(SHEET_VERTICAL)), _
                     "RowIndex should answer the same row whatever case is asked for"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSheetLookupIgnoresCase", Err.Number, Err.Description
End Sub

'@sub-title Verify that RowIndex returns a positive worksheet row for a known sheet
'@details
'Arranges by using the module-level Sheets instance with the vertical
'fixture sheet name. Acts by calling RowIndex to retrieve the worksheet
'row number. Asserts that the returned index is greater than zero,
'confirming that RowIndex successfully resolves a known sheet name to a
'valid row position in the dictionary.
'@TestMethod("LLSheets")
Public Sub TestRowIndexReturnsWorksheetRow()
    CustomTestSetTitles Assert, "LLSheets", "TestRowIndexReturnsWorksheetRow"
    On Error GoTo Fail

    Dim idx As Long
    idx = Sheets.RowIndex(SHEET_VERTICAL)
    Assert.IsTrue (idx > 0), "RowIndex should return a positive worksheet row"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRowIndexReturnsWorksheetRow", Err.Number, Err.Description
End Sub

'@sub-title Verify that RowIndex answers 0 for the header row
'@details
'Arranges by using the module-level Sheets instance, whose sheet-name range
'is captured with its header row. Acts by calling RowIndex with the header
'text. Asserts that the answer is 0. LLDataEntry places rows from this
'answer, so a header row handed back as a sheet row would write over the
'dictionary head.
'@TestMethod("LLSheets")
Public Sub TestRowIndexRejectsHeaderRow()
    CustomTestSetTitles Assert, "LLSheets", "TestRowIndexRejectsHeaderRow"
    On Error GoTo Fail

    Assert.AreEqual 0&, Sheets.RowIndex("Sheet Name"), _
                     "RowIndex should answer 0 when passed the header text"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRowIndexRejectsHeaderRow", Err.Number, Err.Description
End Sub

'@sub-title Verify that a held sheet row is checked before it is handed out
'@details
'Arranges by reading the row of the vertical fixture sheet, which makes the
'object hold it, then writing another name into that very cell. Acts by
'asking for the row again. Asserts that the answer is a different row that
'still holds the sheet name. A held row that was handed back without the
'check would answer for a row that now belongs to another sheet.
'@TestMethod("LLSheets")
Public Sub TestHeldSheetRowIsCheckedBeforeUse()
    CustomTestSetTitles Assert, "LLSheets", "TestHeldSheetRowIsCheckedBeforeUse"
    On Error GoTo Fail

    Dim firstRow As Long
    Dim laterRow As Long
    Dim nameColumn As Long

    firstRow = Sheets.RowIndex(SHEET_VERTICAL)
    nameColumn = Dictionary.Data.ColumnIndex("sheet name", matchCase:=False)
    ThisWorkbook.Worksheets(DICT_SHEET).Cells(firstRow, nameColumn).Value = "renamed-sheet"

    laterRow = Sheets.RowIndex(SHEET_VERTICAL)

    Assert.IsTrue (laterRow > 0), "RowIndex should still find the sheet on its other rows"
    Assert.IsTrue (laterRow <> firstRow), "RowIndex should leave the row that now holds another name"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestHeldSheetRowIsCheckedBeforeUse", Err.Number, Err.Description
End Sub

'@sub-title Verify that DataBounds raises for an unsupported selector value
'@details
'Arranges by passing selector value 99 (which does not correspond to any
'SheetBound enum member) to DataBounds for a known sheet. Acts by calling
'DataBounds, which should raise an error. Asserts that the error number
'equals ProjectError.InvalidArgument, confirming that the method validates
'its selector parameter and rejects out-of-range values.
'@TestMethod("LLSheets")
Public Sub TestDataBoundsRejectsUnknownSelector()
    CustomTestSetTitles Assert, "LLSheets", "TestDataBoundsRejectsUnknownSelector"
    On Error GoTo ExpectError

    Dim unused As Long
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.DataBounds(SHEET_VERTICAL, 99)
    Assert.LogFailure "DataBounds should raise for unsupported selectors"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Invalid selectors should return InvalidArgument - Description " & Err.Description
    Err.Clear
End Sub

'@sub-title Verify that SheetInfo raises when the table-name column is absent
'@details
'Arranges by requesting SheetInfoSheetTable from the fixture, which does
'not include the required table-name column. Acts by calling SheetInfo
'with the SheetInfoSheetTable selector. Asserts that the error number
'equals ProjectError.ElementNotFound, confirming that SheetInfo detects
'and reports a missing table column rather than returning invalid data.
'@TestMethod("LLSheets")
Public Sub TestSheetInfoRaisesWhenTableColumnMissing()
    CustomTestSetTitles Assert, "LLSheets", "TestSheetInfoRaisesWhenTableColumnMissing"
    On Error GoTo ExpectError

    Dim unused As String
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.SheetInfo(SHEET_VERTICAL, SheetInfoType.SheetInfoSheetTable)
    Assert.LogFailure "SheetInfo should raise when table name column is missing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing table column should raise ElementNotFound"
    Err.Clear
End Sub

'@sub-title Verify that ContainsControl detects formula controls and rejects missing ones
'@details
'Arranges by using the module-level Sheets instance with the vertical
'fixture sheet. Acts by calling ContainsControl twice: once with "formula"
'as the control type, and once with "__missing__". Asserts that the first
'call returns True (confirming formula controls exist in the fixture) and
'the second returns False (confirming non-existent control types are
'correctly rejected).
'@TestMethod("LLSheets")
Public Sub TestContainsControlDetectsFormulaControls()
    CustomTestSetTitles Assert, "LLSheets", "TestContainsControlDetectsFormulaControls"
    On Error GoTo Fail

    Assert.IsTrue Sheets.ContainsControl(SHEET_VERTICAL, "formula", colName:="Control"), _
                  "Expected the fixture sheet to include formula controls"
    Assert.IsFalse Sheets.ContainsControl(SHEET_VERTICAL, "__missing__"), _
                   "Non-existent control types should return False"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestContainsControlDetectsFormulaControls", Err.Number, Err.Description
End Sub

'@sub-title Verify that NumberOfVars raises for an unknown sheet name
'@details
'Arranges by calling NumberOfVars with "unknown-sheet", which does not
'exist in the dictionary fixture. Acts by executing the call, which
'should raise an error. Asserts that the error number equals
'ProjectError.ElementNotFound, confirming that the method validates
'sheet existence and raises an appropriate error for missing sheets.
'@TestMethod("LLSheets")
Public Sub TestNumberOfVarsRaisesWhenSheetMissing()
    CustomTestSetTitles Assert, "LLSheets", "TestNumberOfVarsRaisesWhenSheetMissing"
    On Error GoTo ExpectError

    Dim unused As Long
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.NumberOfVars("unknown-sheet")
    Assert.LogFailure "NumberOfVars should raise when the sheet is absent"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing sheets should raise ElementNotFound"
    Err.Clear
End Sub

'@sub-title Verify that VariableAddress raises when the dictionary is not prepared
'@details
'Arranges by using the module-level Sheets instance whose underlying
'dictionary has not been prepared. Acts by calling VariableAddress with a
'known variable name. Asserts that the error number equals
'ProjectError.ObjectNotInitialized, confirming that VariableAddress
'enforces a preparation prerequisite and refuses to resolve addresses
'against an unprepared dictionary.
'@TestMethod("LLSheets")
Public Sub TestVariableAddressRequiresPreparedDictionary()
    CustomTestSetTitles Assert, "LLSheets", "TestVariableAddressRequiresPreparedDictionary"
    On Error GoTo ExpectError

    Dim unused As String
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.VariableAddress(KNOWN_VARIABLE)
    Assert.LogFailure "VariableAddress should require a prepared dictionary"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "VariableAddress should signal missing preparation"
    Err.Clear
End Sub
