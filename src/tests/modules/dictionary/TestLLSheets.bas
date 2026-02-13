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
'@depends LLSheets, ILLSheets, LLdictionary, ILLdictionary, CustomTest, ICustomTest

Private Const DICT_SHEET As String = "LLSheetsDict"
Private Const SHEET_VERTICAL As String = "vlist1D-sheet1"
Private Const SHEET_HORIZONTAL As String = "hlist2D-sheet1"
Private Const KNOWN_VARIABLE As String = "choi_v1"

Private Assert As ICustomTest
Private Dictionary As ILLdictionary
Private Sheets As ILLSheets

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
'Creates the test-output worksheet if it does not already exist, builds
'the CustomTest assertion object, registers the module name for reporting,
'and resets the dictionary fixture to a clean baseline. Runs once before
'any test method in this module executes.
'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLSheets"
    ResetDictionarySheet
End Sub

'@sub-title Tear down module-level resources after all tests complete
'@details
'Prints accumulated test results to the output sheet, releases object
'references for the Sheets, Dictionary, and Assert instances, and deletes
'the temporary dictionary worksheet. Runs once after every test method in
'this module has finished.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Sheets = Nothing
    Set Dictionary = Nothing
    Set Assert = Nothing
    DeleteWorksheet DICT_SHEET
End Sub

'@sub-title Create fresh Dictionary and Sheets instances before each test
'@details
'Resets the dictionary fixture worksheet, constructs a new LLdictionary
'from the fixture, and wraps it in a new LLSheets instance. This ensures
'every test starts with an unmodified dictionary and a cleanly initialised
'Sheets object so that tests remain independent of one another.
'@TestInitialize
Private Sub TestInitialize()
    ResetDictionarySheet
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    Set Sheets = LLSheets.Create(Dictionary)
End Sub

'@sub-title Release per-test objects and flush assertion state
'@details
'Flushes any buffered assertion results to the output sheet and releases
'the Sheets and Dictionary references. Runs after each individual test
'method completes.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
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

    Dim invalid As ILLSheets
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
