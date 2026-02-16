Attribute VB_Name = "TestLLSheetsExtra"
Attribute VB_Description = "Additional tests for the LLSheets class"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'@Folder("CustomTests")

'@ModuleDescription("Additional tests for the LLSheets class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'Provides supplementary tests for the LLSheets class that exercise
'scenarios not covered by the primary TestLLSheets module. Tests validate
'that Contains rejects header names, that DataBounds returns correct
'row and column boundaries for both vertical and horizontal sheet layouts,
'that ContainsControl raises when the control column is removed from the
'dictionary, and that VariableAddress resolves correct cell references
'for horizontal and vertical variables after dictionary preparation.
'@depends LLSheets, ILLSheets, LLdictionary, ILLdictionary, LLVariables, ILLVariables, CustomTest, ICustomTest

Private Const DICT_SHEET As String = "LLSheetsExtraDict"

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
    Assert.SetModuleName "TestLLSheetsExtra"
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

'@sub-title Verify that Contains returns False when given the header name itself
'@details
'Arranges by using the module-level Sheets instance built from the
'dictionary fixture. Acts by calling Contains with the literal column
'header string "Sheet Name". Asserts that Contains returns False,
'confirming that the method treats the header row value as a non-match
'and does not confuse it with an actual sheet entry.
'@TestMethod("LLSheetsExtra")
Public Sub TestContainsRejectsHeaderName()
    CustomTestSetTitles Assert, "LLSheets", "TestContainsRejectsHeaderName"
    On Error GoTo Fail

    Assert.IsFalse Sheets.Contains("Sheet Name"), "Contains should return False when passed the header name"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestContainsRejectsHeaderName", Err.Number, Err.Description
End Sub

'@sub-title Verify that DataBounds returns correct boundaries for vertical and horizontal layouts
'@details
'Arranges by using the module-level Sheets instance and retrieving all
'four boundary values (RowStart, RowEnd, ColStart, ColEnd) for both the
'vertical fixture sheet and the horizontal fixture sheet, along with the
'variable count for each. Acts by calling DataBounds and NumberOfVars for
'both sheet names. Asserts that the vertical layout has top row 4, left
'and right column both at 5, and bottom row equal to top plus count minus
'one, and that the horizontal layout has top row 8, left column 1, bottom
'row at top plus 201, and right column equal to left plus count minus one.
'@TestMethod("LLSheetsExtra")
Public Sub TestDataBoundsForBothLayouts()
    CustomTestSetTitles Assert, "LLSheets", "TestDataBoundsForBothLayouts"
    On Error GoTo Fail

    Dim vTop As Long, vBottom As Long, vLeft As Long, vRight As Long
    Dim hTop As Long, hBottom As Long, hLeft As Long, hRight As Long
    Dim vCount As Long, hCount As Long

    vTop = Sheets.DataBounds("vlist1D-sheet1", SheetBound.RowSart)
    vBottom = Sheets.DataBounds("vlist1D-sheet1", SheetBound.RowEnd)
    vLeft = Sheets.DataBounds("vlist1D-sheet1", SheetBound.ColStart)
    vRight = Sheets.DataBounds("vlist1D-sheet1", SheetBound.ColEnd)

    hTop = Sheets.DataBounds("hlist2D-sheet1", SheetBound.RowSart)
    hBottom = Sheets.DataBounds("hlist2D-sheet1", SheetBound.RowEnd)
    hLeft = Sheets.DataBounds("hlist2D-sheet1", SheetBound.ColStart)
    hRight = Sheets.DataBounds("hlist2D-sheet1", SheetBound.ColEnd)

    vCount = Sheets.NumberOfVars("vlist1D-sheet1")
    hCount = Sheets.NumberOfVars("hlist2D-sheet1")

    Assert.AreEqual 4, vTop, "Vertical layout top row should be 4"
    Assert.AreEqual 5, vLeft, "Vertical layout left column should be 5"
    Assert.AreEqual 5, vRight, "Vertical layout right column equals left column"
    Assert.AreEqual vTop + IIf(vCount > 0, vCount - 1, 0), vBottom, _
                     "Vertical bottom should match top + count - 1"

    Assert.AreEqual 8, hTop, "Horizontal layout top row should be 8"
    Assert.AreEqual 1, hLeft, "Horizontal layout left column should be 1"
    Assert.AreEqual 8 + 201, hBottom, "Horizontal layout bottom should be top + 201 rows"
    Assert.AreEqual hLeft + IIf(hCount > 0, hCount - 1, 0), hRight, _
                     "Horizontal right should match left + count - 1"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDataBoundsForBothLayouts", Err.Number, Err.Description
End Sub

'@sub-title Verify that ContainsControl raises when the control column is removed
'@details
'Arranges by explicitly removing the "control" column from the dictionary
'so that the underlying lookup cannot find it. Acts by calling
'ContainsControl with a valid sheet name and control type, which should
'raise an error due to the missing column. Asserts that the error number
'equals ProjectError.ElementNotFound, confirming that ContainsControl
'validates the existence of the control column before attempting a lookup.
'@TestMethod("LLSheetsExtra")
Public Sub TestContainsControlRaisesWhenControlColumnMissing()
    CustomTestSetTitles Assert, "LLSheets", "TestContainsControlRaisesWhenControlColumnMissing"
    On Error GoTo ExpectError

    Dictionary.RemoveColumn "control"
    Dim hasControl As Boolean
    '@Ignore VariableNotUsed, AssignmentNotUsed
    hasControl = Sheets.ContainsControl("vlist1D-sheet1", "formula")
    Assert.LogFailure "ContainsControl should raise when control column is missing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing control column should raise ElementNotFound"
    Err.Clear
End Sub



'@sub-title Verify that VariableAddress resolves correct addresses for both layouts
'@details
'Arranges by preparing the dictionary and seeding column-index values for
'two representative variables via LLVariables: "num_valid_h2" with column
'index 3 on a horizontal sheet and "choi_v1" with column index 10 on a
'vertical sheet. Acts by calling VariableAddress for each variable. Asserts
'that the horizontal variable returns "C9" (a relative same-sheet address
'computed from column index 3 and top row 8) and that the vertical variable
'returns "'vlist1D-sheet1'!$E$10" (an absolute cross-sheet reference using
'column E and the supplied row index).
'@TestMethod("LLSheetsExtra")
Public Sub TestVariableAddressHorizontalAndVertical()
    CustomTestSetTitles Assert, "LLSheets", "TestVariableAddressHorizontalAndVertical"

    On Error GoTo Fail

    'Prepare the dictionary minimally so Prepared() is True
    Dictionary.Prepare

    'Seed column index values for two representative variables
    Dim vars As ILLVariables
    Set vars = LLVariables.Create(Dictionary)

    vars.SetValue "num_valid_h2", "column index", 3 'Horizontal sheet
    vars.SetValue  "choi_v1", "column index", 10     'Vertical sheet

    'When on the same sheet, horizontal address should be relative and omit prefix
    Dim addrH As String
    addrH = Sheets.VariableAddress("num_valid_h2", onSheet:="hlist2D-sheet1")
    Assert.AreEqual "C9", addrH, "Expected horizontal variable address to be B9 given index=3 and top=8"

    'Vertical address should include sheet prefix and be absolute in A1 style
    Dim addrV As String
    addrV = Sheets.VariableAddress("choi_v1")
    Assert.AreEqual "'vlist1D-sheet1'!$E$10", addrV, _
                     "Expected vertical variable address to target column E and supplied row index"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestVariableAddressHorizontalAndVertical", Err.Number, Err.Description
End Sub
