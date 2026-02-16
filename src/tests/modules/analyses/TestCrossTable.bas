Attribute VB_Name = "TestCrossTable"
Attribute VB_Description = "Tests for CrossTable class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for CrossTable class")

'@description
'Validates the CrossTable class, which lays out analysis cross-tables on an
'output worksheet by creating named ranges, section headers, and column
'structures for univariate and global-summary table scopes. Tests cover
'factory rejection of Nothing parameters, property passthrough for
'Specifications and Wksh, Build behaviour for GlobalSummary and Univariate
'scopes (named-range creation, position tracking, column counting, section
'headers), and NamedRangesList population. The fixture writes a minimal
'TableSpecs layout on a hidden worksheet and uses stubs for the dictionary,
'linelist data, and translation dependencies.
'@depends CrossTable, ICrossTable, TableSpecs, ITableSpecs, TableSpecsLinelistStub,
'  AnalysisDictionaryStub, LinelistSpecsTranslationStub, ILLdictionary,
'  BetterArray, CustomTest, TestHelpers

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "CrossTableFixture"
Private Const OUTPUT_SHEET As String = "CrossTableOutput"

' Column layout for the fixture (must match TableSpecs header order)
Private Const COL_SECTION As Long = 1
Private Const COL_ROW As Long = 2
Private Const COL_COLUMN As Long = 3
Private Const COL_TOTAL As Long = 4
Private Const COL_PERCENTAGE As Long = 5
Private Const COL_MISSING As Long = 6
Private Const COL_GRAPH As Long = 7
Private Const COL_LABEL As Long = 8
Private Const COL_FUNCTION As Long = 9
Private Const COL_NGEO As Long = 10
Private Const NUM_COLUMNS As Long = 10

Private Assert As ICustomTest
Private dictStub As ILLdictionary
Private lDataStub As TableSpecsLinelistStub
Private transStub As LinelistSpecsTranslationStub

'@section Helpers
'===============================================================================

'@sub-title Populate the fixture worksheet with a table-scope header and data rows.
'@details
'Creates or clears the fixture sheet, writes the table scope name in cell A1,
'places the ten-column header row at row 3, and writes any supplied data rows
'starting at row 4. The resulting layout mirrors a minimal analysis
'dictionary sheet that TableSpecs.Create can consume.
Private Sub BuildFixture(ByVal tableScopeName As String, dataRows As Variant)
    Dim sh As Worksheet
    Dim headerArray As Variant
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    sh.Cells(1, 1).Value = tableScopeName
    headerArray = Array( _
        Array("section", "row", "column", "total", "percentage", _
              "missing", "graph", "label", "function", "n geo"))
    headerMatrix = RowsToMatrix(headerArray)
    WriteMatrix sh.Cells(3, 1), headerMatrix

    If Not IsEmpty(dataRows) Then
        dataMatrix = RowsToMatrix(dataRows)
        WriteMatrix sh.Cells(4, 1), dataMatrix
    End If
End Sub

'@sub-title Return the header range of the fixture sheet.
Private Function FixtureHeaderRange() As Range
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    Set FixtureHeaderRange = sh.Range(sh.Cells(3, 1), sh.Cells(3, NUM_COLUMNS))
End Function

'@sub-title Return a single data row from the fixture sheet by index.
Private Function FixtureDataRange(ByVal dataRowIndex As Long) As Range
    Dim sh As Worksheet
    Dim rowNum As Long
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    rowNum = 3 + dataRowIndex
    Set FixtureDataRange = sh.Range(sh.Cells(rowNum, 1), sh.Cells(rowNum, NUM_COLUMNS))
End Function

'@sub-title Build a TableSpecs instance from the fixture header and a data row.
Private Function CreateSpecs(ByVal dataRowIndex As Long) As ITableSpecs
    Set CreateSpecs = TableSpecs.Create( _
        FixtureHeaderRange(), _
        FixtureDataRange(dataRowIndex), _
        dictStub)
End Function

'@sub-title Return a clean output worksheet for cross-table builds.
Private Function OutputSheet() As Worksheet
    Set OutputSheet = EnsureWorksheet(OUTPUT_SHEET, clearSheet:=True, visibility:=xlSheetHidden)
End Function

'@sub-title Check whether a worksheet-scoped named range exists.
Private Function RangeExistsOnSheet(ByVal sh As Worksheet, ByVal rngName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExistsOnSheet = (Not rng Is Nothing)
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Set up the test harness, stubs, and translation entries.
'@details
'Suppresses screen updates, creates the CustomTest assertion object, and
'initialises the three stubs: AnalysisDictionaryStub for dictionary
'lookups, LinelistSpecsTranslationStub with standard message keys
'(MSG_NA, MSG_Total, MSG_Percent, etc.), and TableSpecsLinelistStub
'wired to both. These stubs remain alive for all tests in the module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCrossTable"
    Set dictStub = New AnalysisDictionaryStub
    Set lDataStub = New TableSpecsLinelistStub
    Set transStub = New LinelistSpecsTranslationStub
    transStub.Initialise "TestTrans"
    transStub.SetTranslation "MSG_NA", "Missing"
    transStub.SetTranslation "MSG_Total", "Total"
    transStub.SetTranslation "MSG_Percent", "%"
    transStub.SetTranslation "MSG_AllData", "All Data"
    transStub.SetTranslation "MSG_FilteredData", "Filtered Data"
    transStub.SetTranslation "MSG_GlobalSummary", "Global Summary"
    transStub.SetTranslation "MSG_Period", "Period"
    lDataStub.SetDictionary dictStub
    lDataStub.SetTranslation transStub
End Sub

'@sub-title Print test results and tear down all fixture resources.
'@details
'Outputs the accumulated test results to the output sheet, deletes both
'fixture and output worksheets, restores the Excel application state,
'and releases all stub and assertion references.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet FIXTURE_SHEET
    DeleteWorksheet OUTPUT_SHEET
    RestoreApp
    Set dictStub = Nothing
    Set lDataStub = Nothing
    Set transStub = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updates before each test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush assertion state after each test.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory validation tests
'===============================================================================

'@sub-title Verify Create rejects Nothing for the specs parameter.
'@details
'Arranges a valid output worksheet and passes Nothing as the specs
'argument to CrossTable.Create. Asserts that the returned object is
'Nothing, confirming the factory guards against missing specifications.
'@TestMethod("CrossTable")
Public Sub TestCreateRejectsNothingSpecs()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateRejectsNothingSpecs"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = OutputSheet()

    On Error Resume Next
    Dim ct As ICrossTable
    Set ct = CrossTable.Create(Nothing, sh, lDataStub)
    On Error GoTo 0

    Assert.IsTrue (ct Is Nothing), _
                  "Create with Nothing specs should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingSpecs", Err.Number, Err.Description
End Sub

'@sub-title Verify Create rejects Nothing for the worksheet parameter.
'@details
'Arranges valid TableSpecs via the fixture and passes Nothing as the
'worksheet argument to CrossTable.Create. Asserts that the returned
'object is Nothing, confirming the factory guards against a missing
'output worksheet.
'@TestMethod("CrossTable")
Public Sub TestCreateRejectsNothingWorksheet()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateRejectsNothingWorksheet"
    On Error GoTo TestFail

    ' Build a valid specs fixture
    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    On Error Resume Next
    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, Nothing, lDataStub)
    On Error GoTo 0

    Assert.IsTrue (ct Is Nothing), _
                  "Create with Nothing worksheet should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingWorksheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Create rejects Nothing for the linelist data parameter.
'@details
'Arranges valid TableSpecs and an output worksheet, then passes Nothing
'as the linelist data argument to CrossTable.Create. Asserts that the
'returned object is Nothing, confirming the factory guards against
'missing linelist data.
'@TestMethod("CrossTable")
Public Sub TestCreateRejectsNothingLData()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateRejectsNothingLData"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    On Error Resume Next
    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, Nothing)
    On Error GoTo 0

    Assert.IsTrue (ct Is Nothing), _
                  "Create with Nothing linelist data should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingLData", Err.Number, Err.Description
End Sub

'@sub-title Verify Create returns a valid CrossTable with all valid arguments.
'@details
'Arranges valid TableSpecs, an output worksheet, and linelist data stub,
'then calls CrossTable.Create. Asserts that the returned object is not
'Nothing, confirming successful factory initialisation.
'@TestMethod("CrossTable")
Public Sub TestCreateReturnsValidObject()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateReturnsValidObject"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, lDataStub)

    Assert.IsTrue (Not ct Is Nothing), _
                  "Create with valid params should succeed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsValidObject", Err.Number, Err.Description
End Sub

'@section Property tests
'===============================================================================

'@sub-title Verify Specifications property returns the specs passed at creation.
'@details
'Creates a CrossTable with known TableSpecs and asserts that the
'Specifications property returns an object whose TableId matches the
'original specs. Confirms that the factory correctly stores and exposes
'the specifications reference.
'@TestMethod("CrossTable")
Public Sub TestSpecificationsProperty()
    CustomTestSetTitles Assert, "CrossTable", "TestSpecificationsProperty"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, lDataStub)

    Assert.AreEqual specs.TableId, ct.Specifications.TableId, _
                    "Specifications should return the specs passed at creation"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpecificationsProperty", Err.Number, Err.Description
End Sub

'@sub-title Verify Wksh property returns the output worksheet passed at creation.
'@details
'Creates a CrossTable targeting a known output worksheet and asserts that
'the Wksh property returns a worksheet whose Name matches. Confirms the
'factory correctly stores the output worksheet reference.
'@TestMethod("CrossTable")
Public Sub TestWkshProperty()
    CustomTestSetTitles Assert, "CrossTable", "TestWkshProperty"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, lDataStub)

    Assert.AreEqual sh.Name, ct.Wksh.Name, _
                    "Wksh should return the output worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWkshProperty", Err.Number, Err.Description
End Sub

'@section Build tests — GlobalSummary
'===============================================================================

'@sub-title Verify Build creates ROWGS_SET named range for global summary scope.
'@details
'Arranges a fixture with the "Add or remove rows of Global Summary" scope
'and a single data row with label "Total Cases". Calls Build, then asserts
'that the ROWGS_SET named range exists on the output sheet, confirming that
'the global-summary row header range is created.
'@TestMethod("CrossTable")
Public Sub TestBuildGlobalSummaryCreatesRowGsSet()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildGlobalSummaryCreatesRowGsSet"
    On Error GoTo TestFail

    BuildFixture "Add or remove rows of Global Summary", Array( _
        Array("S1", "", "", "", "", "", "", "Total Cases", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, lDataStub)
    ct.Build

    Assert.IsTrue RangeExistsOnSheet(sh, "ROWGS_SET"), _
                  "Build should create ROWGS_SET named range for global summary"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildGlobalSummaryCreatesRowGsSet", Err.Number, Err.Description
End Sub

'@sub-title Verify Build creates COLGS_SET named range for global summary scope.
'@details
'Arranges the same global-summary fixture and calls Build. Asserts that the
'COLGS_SET named range exists on the output sheet, confirming that the
'global-summary column header range is created alongside the row header.
'@TestMethod("CrossTable")
Public Sub TestBuildGlobalSummaryCreatesColGsSet()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildGlobalSummaryCreatesColGsSet"
    On Error GoTo TestFail

    BuildFixture "Add or remove rows of Global Summary", Array( _
        Array("S1", "", "", "", "", "", "", "Total Cases", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, lDataStub)
    ct.Build

    Assert.IsTrue RangeExistsOnSheet(sh, "COLGS_SET"), _
                  "Build should create COLGS_SET column header for global summary"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildGlobalSummaryCreatesColGsSet", Err.Number, Err.Description
End Sub

'@section Build tests — Univariate
'===============================================================================

'@sub-title Verify Build creates essential named ranges for a univariate table.
'@details
'Arranges a univariate-analysis fixture with a single row variable and
'calls Build. Asserts that the key named ranges are created on the output
'sheet: ENDTABLE_, ROW_CATEGORIES_, VALUES_COL_1_, INTERIOR_VALUES_, and
'TITLE_, each suffixed by the table identifier. These ranges are required
'for formula writing and formatting downstream.
'@TestMethod("CrossTable")
Public Sub TestBuildUnivariateCreatesNamedRanges()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildUnivariateCreatesNamedRanges"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, lDataStub)
    ct.Build

    Dim tabId As String
    tabId = specs.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "ENDTABLE_" & tabId), _
                  "Build should create ENDTABLE_ named range"
    Assert.IsTrue RangeExistsOnSheet(sh, "ROW_CATEGORIES_" & tabId), _
                  "Build should create ROW_CATEGORIES_ named range"
    Assert.IsTrue RangeExistsOnSheet(sh, "VALUES_COL_1_" & tabId), _
                  "Build should create VALUES_COL_1_ named range"
    Assert.IsTrue RangeExistsOnSheet(sh, "INTERIOR_VALUES_" & tabId), _
                  "Build should create INTERIOR_VALUES_ named range"
    Assert.IsTrue RangeExistsOnSheet(sh, "TITLE_" & tabId), _
                  "Build should create TITLE_ named range"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildUnivariateCreatesNamedRanges", Err.Number, Err.Description
End Sub

'@sub-title Verify EndRow and EndColumn are set after a univariate Build.
'@details
'Builds a univariate cross-table and asserts that both EndRow and
'EndColumn are greater than zero. These position properties are used by
'subsequent tables to know where to start writing on the output sheet.
'@TestMethod("CrossTable")
Public Sub TestBuildUnivariateEndRowSet()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildUnivariateEndRowSet"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, lDataStub)
    ct.Build

    Assert.IsTrue ct.EndRow > 0, _
                  "EndRow should be set after Build"
    Assert.IsTrue ct.EndColumn > 0, _
                  "EndColumn should be set after Build"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildUnivariateEndRowSet", Err.Number, Err.Description
End Sub

'@sub-title Verify a univariate table reports exactly one data column.
'@details
'Builds a univariate cross-table (no column variable specified) and
'asserts that NumberOfColumns equals 1. A univariate table has a single
'values column, in contrast to bivariate tables that expand by category.
'@TestMethod("CrossTable")
Public Sub TestBuildUnivariateNumberOfColumns()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildUnivariateNumberOfColumns"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, lDataStub)
    ct.Build

    Assert.AreEqual 1&, ct.NumberOfColumns, _
                    "Univariate table should have 1 data column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildUnivariateNumberOfColumns", Err.Number, Err.Description
End Sub

'@section Build tests — Univariate with section
'===============================================================================

'@sub-title Verify Build creates a SECTION_ named range when the specs define a section.
'@details
'Arranges a univariate fixture whose data row includes a section value
'("S1") and calls Build. Asserts that a SECTION_ named range suffixed by
'the table identifier exists on the output sheet, confirming that section
'header rows are written for tables that start a new analysis section.
'@TestMethod("CrossTable")
Public Sub TestBuildUnivariateNewSectionCreatesSection()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildUnivariateNewSectionCreatesSection"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, lDataStub)
    ct.Build

    Dim tabId As String
    tabId = specs.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "SECTION_" & tabId), _
                  "Build should create SECTION_ named range for new section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildUnivariateNewSectionCreatesSection", Err.Number, Err.Description
End Sub

'@section NamedRangesList tests
'===============================================================================

'@sub-title Verify NamedRangesList is populated after Build.
'@details
'Builds a univariate cross-table and retrieves the NamedRangesList
'BetterArray. Asserts that the array length is greater than zero,
'confirming that Build tracks all named ranges it creates. This list
'is consumed downstream for cleanup and formula-writing operations.
'@TestMethod("CrossTable")
Public Sub TestNamedRangesListPopulatedAfterBuild()
    CustomTestSetTitles Assert, "CrossTable", "TestNamedRangesListPopulatedAfterBuild"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, lDataStub)
    ct.Build

    Dim rangeNames As BetterArray
    Set rangeNames = ct.NamedRangesList

    Assert.IsTrue rangeNames.Length > 0, _
                  "NamedRangesList should contain entries after Build"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNamedRangesListPopulatedAfterBuild", Err.Number, Err.Description
End Sub
