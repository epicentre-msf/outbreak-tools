Attribute VB_Name = "TestTableSpecs"
Attribute VB_Description = "Tests for TableSpecs class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for TableSpecs class")

'@description
'Validates the TableSpecs class, which parses a single row from an analysis
'specification table and exposes computed properties such as TableScope,
'TableId, HasTotal, HasPercentage, HasMissing, HasGraph, and section
'navigation (IsNewSection, Previous, NextSpecs, TableSectionId). Tests cover
'factory guard-clause validation, scope resolution for all six analysis types
'(time series, global summary, univariate, bivariate, spatial, and
'spatio-temporal), flag computation that depends on scope and column presence,
'the TotalRequested vs HasTotal distinction (Bug #1 fix), and forward/backward
'row navigation. The fixture builds a hidden worksheet with a typed header row
'and data rows, backed by a dictionary fixture for variable validation.
'@depends TableSpecs, ITableSpecs, LLdictionary, ILLdictionary, CustomTest, TestHelpers

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "TableSpecsFixture"
Private Const DICT_SHEET As String = "TableSpecsDict"

' Column layout for the fixture (must match header order)
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
Private dict As ILLdictionary

'@section Helpers
'===============================================================================

'@sub-title Build a fixture sheet with type label, header, and data rows.
'@details
'Creates or clears the hidden FIXTURE_SHEET worksheet and populates it with
'the analysis type label in row 1, a standard 10-column header in row 3,
'and caller-supplied data rows starting at row 4. The layout matches the
'structure TableSpecs expects: type label at hRng.Cells(-1, 1), header
'range spanning NUM_COLUMNS columns, and contiguous data rows below.
Private Sub BuildFixture(ByVal tableScopeName As String, dataRows As Variant)
    Dim sh As Worksheet
    Dim headerArray As Variant
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    ' Row 1: type label (read by TableScope via hRng.Cells(-1, 1))
    sh.Cells(1, 1).Value = tableScopeName

    ' Row 3: header row
    headerArray = Array( _
        Array("section", "row", "column", "total", "percentage", _
              "missing", "graph", "label", "function", "n geo"))
    headerMatrix = RowsToMatrix(headerArray)
    WriteMatrix sh.Cells(3, 1), headerMatrix

    ' Row 4+: data rows
    If Not IsEmpty(dataRows) Then
        dataMatrix = RowsToMatrix(dataRows)
        WriteMatrix sh.Cells(4, 1), dataMatrix
    End If
End Sub

'@sub-title Return the header range (row 3, NUM_COLUMNS wide).
Private Function FixtureHeaderRange() As Range
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    Set FixtureHeaderRange = sh.Range(sh.Cells(3, 1), sh.Cells(3, NUM_COLUMNS))
End Function

'@sub-title Return a data row range by 1-based index.
'@details
'Computes the absolute row number as 3 + dataRowIndex (header at row 3,
'first data row at row 4) and returns a range spanning NUM_COLUMNS columns.
Private Function FixtureDataRange(ByVal dataRowIndex As Long) As Range
    Dim sh As Worksheet
    Dim rowNum As Long

    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    rowNum = 3 + dataRowIndex
    Set FixtureDataRange = sh.Range(sh.Cells(rowNum, 1), sh.Cells(rowNum, NUM_COLUMNS))
End Function

'@sub-title Create an ITableSpecs from a fixture data row index.
Private Function CreateSpecs(ByVal dataRowIndex As Long) As ITableSpecs
    Set CreateSpecs = TableSpecs.Create( _
        FixtureHeaderRange(), _
        FixtureDataRange(dataRowIndex), _
        dict)
End Function

'@sub-title Build the standard three-row time series fixture data.
'@details
'Returns an array of three row arrays using real dictionary variable names
'so that ValidTable resolves correctly during Create. Row 1: section S1,
'date_v1/choi_v1, all flags on (total=yes, percentage=row, missing=yes,
'graph=yes). Row 2: same section S1, same variables, all flags off. Row 3:
'new section S2, date_v1, no column variable, all flags off.
Private Function TimeSeriesDataRows() As Variant
    TimeSeriesDataRows = Array( _
        Array("S1", "date_v1", "choi_v1", "yes", "row", "yes", "yes", "", "", ""), _
        Array("S1", "date_v1", "choi_v1", "no", "no", "no", "no", "", "", ""), _
        Array("S2", "date_v1", "", "no", "no", "no", "no", "", "", ""))
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Set up module-level fixtures for all TableSpecs tests.
'@details
'Suppresses screen updating, ensures the test output sheet exists, creates
'the CustomTest assert object, prepares a dictionary fixture sheet with
'known variable definitions (via PrepareDictionaryFixture), and wraps it
'in an LLdictionary instance used by all tests.
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestTableSpecs"
    PrepareDictionaryFixture DICT_SHEET
    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
End Sub

'@sub-title Print results and tear down module-level fixtures.
'@details
'Prints accumulated test results to the output sheet, deletes both the
'fixture and dictionary worksheets, restores Excel application state,
'and releases object references.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet FIXTURE_SHEET
    DeleteWorksheet DICT_SHEET
    RestoreApp
    Set dict = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updating before each test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush assert state after each test.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory validation tests
'===============================================================================

'@sub-title Verify Create returns Nothing when the header range is Nothing.
'@details
'Arranges a fixture sheet with a data range but passes Nothing as the header
'range to TableSpecs.Create. Asserts that the returned specs object is
'Nothing, confirming the factory guard clause rejects invalid input.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsNothingHeader()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsNothingHeader"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    Dim dataRng As Range
    Set dataRng = sh.Range(sh.Cells(4, 1), sh.Cells(4, NUM_COLUMNS))

    On Error Resume Next
    Dim specs As ITableSpecs
    Set specs = TableSpecs.Create(Nothing, dataRng, dict)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "Create with Nothing header should fail (return Nothing)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingHeader", Err.Number, Err.Description
End Sub

'@sub-title Verify Create returns Nothing when the data range is Nothing.
'@details
'Arranges a fixture sheet with a header range but passes Nothing as the data
'range to TableSpecs.Create. Asserts that the returned specs object is
'Nothing.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsNothingRange()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsNothingRange"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    Dim hRng As Range
    Set hRng = sh.Range(sh.Cells(3, 1), sh.Cells(3, NUM_COLUMNS))

    On Error Resume Next
    Dim specs As ITableSpecs
    Set specs = TableSpecs.Create(hRng, Nothing, dict)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "Create with Nothing data range should fail (return Nothing)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingRange", Err.Number, Err.Description
End Sub

'@sub-title Verify Create returns Nothing when the dictionary is Nothing.
'@details
'Arranges a fixture sheet with valid header and data ranges but passes
'Nothing as the dictionary to TableSpecs.Create. Asserts that the returned
'specs object is Nothing.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsNothingDict()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsNothingDict"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    Dim hRng As Range
    Dim dataRng As Range
    Set hRng = sh.Range(sh.Cells(3, 1), sh.Cells(3, NUM_COLUMNS))
    Set dataRng = sh.Range(sh.Cells(4, 1), sh.Cells(4, NUM_COLUMNS))

    On Error Resume Next
    Dim specs As ITableSpecs
    Set specs = TableSpecs.Create(hRng, dataRng, Nothing)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "Create with Nothing dictionary should fail (return Nothing)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingDict", Err.Number, Err.Description
End Sub

'@sub-title Verify Create returns Nothing when header and data column counts differ.
'@details
'Arranges a header range with 10 columns and a data range with only 5
'columns. Asserts that TableSpecs.Create rejects the mismatched widths
'and returns Nothing.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsMismatchedColumns()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsMismatchedColumns"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    ' Header with 10 columns, data with 5 columns
    Dim hRng As Range
    Dim dataRng As Range
    Set hRng = sh.Range(sh.Cells(3, 1), sh.Cells(3, NUM_COLUMNS))
    Set dataRng = sh.Range(sh.Cells(4, 1), sh.Cells(4, 5))

    On Error Resume Next
    Dim specs As ITableSpecs
    Set specs = TableSpecs.Create(hRng, dataRng, dict)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "Create with mismatched column count should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsMismatchedColumns", Err.Number, Err.Description
End Sub

'@section TableScope tests
'===============================================================================

'@sub-title Verify time series analysis scope is parsed correctly.
'@details
'Builds a fixture with the type label "time series analysis" and standard
'data rows. Creates a TableSpecs from the first data row and asserts that
'TableScope equals ScopeTimeSeries.
'@TestMethod("TableSpecs")
Public Sub TestTableScopeTimeSeries()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableScopeTimeSeries"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(ScopeTimeSeries), CLng(specs.TableScope), _
                    "Should parse 'time series analysis' as ScopeTimeSeries"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableScopeTimeSeries", Err.Number, Err.Description
End Sub

'@sub-title Verify global summary scope is parsed correctly.
'@details
'Builds a fixture with the type label "global summary" and a single data
'row containing label and function columns. Asserts that TableScope equals
'ScopeGlobalSummary.
'@TestMethod("TableSpecs")
Public Sub TestTableScopeGlobalSummary()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableScopeGlobalSummary"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("", "", "", "", "", "", "", "Count of cases", "sum", ""))
    BuildFixture "global summary", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(ScopeGlobalSummary), CLng(specs.TableScope), _
                    "Should parse 'global summary' as ScopeGlobalSummary"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableScopeGlobalSummary", Err.Number, Err.Description
End Sub

'@sub-title Verify univariate analysis scope is parsed correctly.
'@details
'Builds a fixture with the type label "univariate analysis" and a single
'data row with a row variable and flags. Asserts that TableScope equals
'ScopeUnivariate.
'@TestMethod("TableSpecs")
Public Sub TestTableScopeUnivariate()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableScopeUnivariate"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "choi_v1", "", "", "yes", "yes", "yes", "", "", ""))
    BuildFixture "univariate analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(ScopeUnivariate), CLng(specs.TableScope), _
                    "Should parse 'univariate analysis' as ScopeUnivariate"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableScopeUnivariate", Err.Number, Err.Description
End Sub

'@sub-title Verify bivariate analysis scope is parsed correctly.
'@details
'Builds a fixture with the type label "bivariate analysis" and a data row
'containing both row and column variables. Asserts that TableScope equals
'ScopeBivariate.
'@TestMethod("TableSpecs")
Public Sub TestTableScopeBivariate()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableScopeBivariate"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "choi_v1", "choi_h2", "", "row", "row", "values", "", "", ""))
    BuildFixture "bivariate analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(ScopeBivariate), CLng(specs.TableScope), _
                    "Should parse 'bivariate analysis' as ScopeBivariate"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableScopeBivariate", Err.Number, Err.Description
End Sub

'@sub-title Verify spatial analysis scope is parsed correctly.
'@details
'Builds a fixture with the type label "spatial analysis" and a data row
'with a geo row variable and a choice column variable. Asserts that
'TableScope equals ScopeSpatial.
'@TestMethod("TableSpecs")
Public Sub TestTableScopeSpatial()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableScopeSpatial"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "geo_h2", "choi_v1", "", "yes", "yes", "yes", "", "", ""))
    BuildFixture "spatial analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(ScopeSpatial), CLng(specs.TableScope), _
                    "Should parse 'spatial analysis' as ScopeSpatial"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableScopeSpatial", Err.Number, Err.Description
End Sub

'@sub-title Verify spatio-temporal analysis scope is parsed correctly.
'@details
'Builds a fixture with the type label "spatio-temporal analysis" and a data
'row with a date row variable, geo column variable, and n geo value. Asserts
'that TableScope equals ScopeSpatioTemporal.
'@TestMethod("TableSpecs")
Public Sub TestTableScopeSpatioTemporal()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableScopeSpatioTemporal"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "date_v1", "geo_h2", "", "", "", "yes", "", "", "5"))
    BuildFixture "spatio-temporal analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(ScopeSpatioTemporal), CLng(specs.TableScope), _
                    "Should parse 'spatio-temporal analysis' as ScopeSpatioTemporal"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableScopeSpatioTemporal", Err.Number, Err.Description
End Sub

'@section TableId tests
'===============================================================================

'@sub-title Verify TableId uses scope prefix and row offset for time series.
'@details
'Builds a time series fixture with three data rows. Creates specs from data
'rows 1 and 2. Asserts that TableId follows the pattern "TS_tab<offset>"
'where offset is the data range row minus the header row. Row 1 produces
'"TS_tab1" and row 2 produces "TS_tab2".
'@TestMethod("TableSpecs")
Public Sub TestTableIdTimeSeries()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableIdTimeSeries"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()

    ' First data row at sheet row 4, header at row 3: offset = 4 - 3 = 1
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual "TS_tab1", specs.TableId, _
                    "Time series table id should be TS_tab<row offset>"

    ' Second data row at sheet row 5: offset = 5 - 3 = 2
    Set specs = CreateSpecs(2)
    Assert.AreEqual "TS_tab2", specs.TableId, _
                    "Second row table id should use row offset 2"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableIdTimeSeries", Err.Number, Err.Description
End Sub

'@sub-title Verify TableId uses GS prefix for global summary scope.
'@details
'Builds a global summary fixture with one data row. Asserts that TableId
'starts with "GS_" followed by the row offset identifier.
'@TestMethod("TableSpecs")
Public Sub TestTableIdGlobalSummary()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableIdGlobalSummary"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("", "", "", "", "", "", "", "Count", "sum", ""))
    BuildFixture "global summary", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual "GS_tab1", specs.TableId, _
                    "Global summary table id should start with GS"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableIdGlobalSummary", Err.Number, Err.Description
End Sub

'@section Value tests
'===============================================================================

'@sub-title Verify Value returns the correct cell content for known columns.
'@details
'Builds a time series fixture and creates specs from the first data row.
'Asserts that Value returns the expected cell content for each of the
'standard column names: section, row, column, total, and percentage.
'@TestMethod("TableSpecs")
Public Sub TestValueReturnsColumnData()
    CustomTestSetTitles Assert, "TableSpecs", "TestValueReturnsColumnData"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual "S1", specs.Value("section"), _
                    "Value('section') should return the section name"
    Assert.AreEqual "date_v1", specs.Value("row"), _
                    "Value('row') should return the row variable"
    Assert.AreEqual "choi_v1", specs.Value("column"), _
                    "Value('column') should return the column variable"
    Assert.AreEqual "yes", specs.Value("total"), _
                    "Value('total') should return 'yes'"
    Assert.AreEqual "row", specs.Value("percentage"), _
                    "Value('percentage') should return 'row'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueReturnsColumnData", Err.Number, Err.Description
End Sub

'@sub-title Verify Value returns an empty string for an unknown column name.
'@details
'Builds a time series fixture and queries Value with a column name that
'does not exist in the header. Asserts that the result is vbNullString
'rather than raising an error.
'@TestMethod("TableSpecs")
Public Sub TestValueReturnsEmptyForUnknownColumn()
    CustomTestSetTitles Assert, "TableSpecs", "TestValueReturnsEmptyForUnknownColumn"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual vbNullString, specs.Value("nonexistent_column"), _
                    "Value for missing column should return empty string"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueReturnsEmptyForUnknownColumn", Err.Number, Err.Description
End Sub

'@section IsNewSection tests
'===============================================================================

'@sub-title Verify the first data row is always flagged as a new section.
'@details
'Builds a time series fixture and creates specs from data row 1. The
'previous row is the header, so the section value will not match. Asserts
'that IsNewSection returns True.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionFirstRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionFirstRow"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    ' First data row: previous row is the header, so section won't match
    Assert.IsTrue specs.IsNewSection, _
                  "First data row should always be a new section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionFirstRow", Err.Number, Err.Description
End Sub

'@sub-title Verify a row with the same section as its predecessor is not new.
'@details
'Builds a time series fixture and creates specs from data row 2, which has
'section "S1" identical to row 1. Asserts that IsNewSection returns False.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionSameSection()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionSameSection"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs

    ' Row 2 has section "S1" same as row 1
    Set specs = CreateSpecs(2)
    Assert.IsFalse specs.IsNewSection, _
                   "Row in same section as previous should not be a new section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionSameSection", Err.Number, Err.Description
End Sub

'@sub-title Verify a row with a different section than its predecessor is new.
'@details
'Builds a time series fixture and creates specs from data row 3, which has
'section "S2" whereas row 2 has "S1". Asserts that IsNewSection returns True.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionDifferentSection()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionDifferentSection"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs

    ' Row 3 has section "S2", different from row 2 "S1"
    Set specs = CreateSpecs(3)
    Assert.IsTrue specs.IsNewSection, _
                  "Row with different section than previous should be a new section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionDifferentSection", Err.Number, Err.Description
End Sub

'@sub-title Verify global summary scope never reports a new section.
'@details
'Builds a global summary fixture with a section value present in the data.
'Asserts that IsNewSection returns False regardless, because global summary
'tables are not sectioned.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionGlobalSummaryAlwaysFalse()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionGlobalSummaryAlwaysFalse"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "", "", "", "", "", "", "Count", "sum", ""))
    BuildFixture "global summary", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsFalse specs.IsNewSection, _
                   "GlobalSummary should never be a new section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionGlobalSummaryAlwaysFalse", Err.Number, Err.Description
End Sub

'@section HasTotal tests
'===============================================================================

'@sub-title Verify time series HasTotal is True when total is "yes" and column exists.
'@details
'Uses the standard time series fixture row 1 where total="yes" and
'column="choi_v1". Asserts that HasTotal returns True.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalTimeSeriesWithTotalYes()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalTimeSeriesWithTotalYes"
    On Error GoTo TestFail

    ' Row with total=yes and column variable present
    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)  ' total=yes, column=choi_v1

    Assert.IsTrue specs.HasTotal, _
                  "Time series with total='yes' and column present should have total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalTimeSeriesWithTotalYes", Err.Number, Err.Description
End Sub

'@sub-title Verify time series HasTotal is True when driven by percentage=row.
'@details
'Builds a custom fixture where total="no" but percentage="row" and a column
'variable is present. HasTotal must return True because percentage
'computation requires a total row for the denominator.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalTimeSeriesPercentageDriven()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalTimeSeriesPercentageDriven"
    On Error GoTo TestFail

    ' Row with total=no but percentage=row and column present
    Dim rows As Variant
    rows = Array(Array("S1", "date_v1", "choi_v1", "no", "row", "no", "no", "", "", ""))
    BuildFixture "time series analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsTrue specs.HasTotal, _
                  "Time series with percentage='row' should have total for computation"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalTimeSeriesPercentageDriven", Err.Number, Err.Description
End Sub

'@sub-title Verify time series HasTotal is True when driven by percentage=column.
'@details
'Builds a custom fixture where total="no" but percentage="column" and a
'column variable is present. Asserts that HasTotal returns True because
'column-based percentage also requires a total row.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalTimeSeriesPercentageColumnDriven()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalTimeSeriesPercentageColumnDriven"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "date_v1", "choi_v1", "no", "column", "no", "no", "", "", ""))
    BuildFixture "time series analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsTrue specs.HasTotal, _
                  "Time series with percentage='column' should have total for computation"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalTimeSeriesPercentageColumnDriven", Err.Number, Err.Description
End Sub

'@sub-title Verify time series HasTotal is False when no column variable exists.
'@details
'Uses the standard time series fixture row 3 which has no column variable
'and total="no". Asserts that HasTotal returns False because a total row
'only applies when there is a column variable to subtotal.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalTimeSeriesNoColumnNoTotal()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalTimeSeriesNoColumnNoTotal"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs

    ' Row 3: no column, total=no, percentage=no
    Set specs = CreateSpecs(3)

    Assert.IsFalse specs.HasTotal, _
                   "Time series without column variable should not have total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalTimeSeriesNoColumnNoTotal", Err.Number, Err.Description
End Sub

'@sub-title Verify global summary scope never reports HasTotal.
'@details
'Builds a global summary fixture with total="yes" in the data row. Asserts
'that HasTotal returns False regardless, because global summary tables do
'not use total rows.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalGlobalSummaryAlwaysFalse()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalGlobalSummaryAlwaysFalse"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("", "", "", "yes", "", "", "", "Count", "sum", ""))
    BuildFixture "global summary", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsFalse specs.HasTotal, _
                   "GlobalSummary should never have total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalGlobalSummaryAlwaysFalse", Err.Number, Err.Description
End Sub

'@sub-title Verify univariate scope always reports HasTotal.
'@details
'Builds a univariate fixture with total="no" in the data. Asserts that
'HasTotal returns True regardless, because univariate tables always include
'a total row by design.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalUnivariateAlwaysTrue()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalUnivariateAlwaysTrue"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "choi_v1", "", "no", "no", "no", "no", "", "", ""))
    BuildFixture "univariate analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsTrue specs.HasTotal, _
                  "Univariate should always have total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalUnivariateAlwaysTrue", Err.Number, Err.Description
End Sub

'@section TotalRequested tests (Bug #1 fix)
'===============================================================================

'@sub-title Verify TotalRequested is True when the user explicitly set total to "yes".
'@details
'Uses the standard time series fixture row 1 where total="yes". Asserts
'that TotalRequested returns True. This property distinguishes user intent
'from computed HasTotal.
'@TestMethod("TableSpecs")
Public Sub TestTotalRequestedTrueWhenExplicit()
    CustomTestSetTitles Assert, "TableSpecs", "TestTotalRequestedTrueWhenExplicit"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)  ' total=yes

    Assert.IsTrue specs.TotalRequested, _
                  "TotalRequested should be True when total='yes'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTotalRequestedTrueWhenExplicit", Err.Number, Err.Description
End Sub

'@sub-title Verify TotalRequested is False when total is driven by percentage only.
'@details
'Builds a fixture where total="no" but percentage="row" with a column
'present, so HasTotal is True (computed need). Asserts that TotalRequested
'is False because the user did not explicitly request a total row. This is
'the Bug #1 fix: the rendering layer uses TotalRequested to decide whether
'to display the total label, while HasTotal controls row allocation.
'@TestMethod("TableSpecs")
Public Sub TestTotalRequestedFalseWhenPercentageDriven()
    CustomTestSetTitles Assert, "TableSpecs", "TestTotalRequestedFalseWhenPercentageDriven"
    On Error GoTo TestFail

    ' HasTotal will be True (percentage=row + column present) but TotalRequested should be False
    Dim rows As Variant
    rows = Array(Array("S1", "date_v1", "choi_v1", "no", "row", "no", "no", "", "", ""))
    BuildFixture "time series analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsTrue specs.HasTotal, _
                  "HasTotal should be True due to percentage"
    Assert.IsFalse specs.TotalRequested, _
                   "TotalRequested should be False when total is not 'yes'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTotalRequestedFalseWhenPercentageDriven", Err.Number, Err.Description
End Sub

'@sub-title Verify TotalRequested is False when total is explicitly "no".
'@details
'Uses the standard time series fixture row 2 where total="no" and
'percentage="no". Asserts that TotalRequested returns False.
'@TestMethod("TableSpecs")
Public Sub TestTotalRequestedFalseWhenNoTotal()
    CustomTestSetTitles Assert, "TableSpecs", "TestTotalRequestedFalseWhenNoTotal"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(2)  ' total=no, percentage=no

    Assert.IsFalse specs.TotalRequested, _
                   "TotalRequested should be False when total='no'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTotalRequestedFalseWhenNoTotal", Err.Number, Err.Description
End Sub

'@section HasPercentage tests
'===============================================================================

'@sub-title Verify time series HasPercentage is True with percentage=row and HasTotal.
'@details
'Uses the standard time series fixture row 1 where percentage="row",
'total="yes", and a column variable is present. Asserts HasPercentage
'returns True because both the percentage flag and a total row are present.
'@TestMethod("TableSpecs")
Public Sub TestHasPercentageTimeSeriesRowWithTotal()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasPercentageTimeSeriesRowWithTotal"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)  ' percentage=row, total=yes, column present

    Assert.IsTrue specs.HasPercentage, _
                  "Time series with percentage='row' and HasTotal should have percentage"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasPercentageTimeSeriesRowWithTotal", Err.Number, Err.Description
End Sub

'@sub-title Verify time series HasPercentage is False when HasTotal is False.
'@details
'Builds a fixture with percentage="row" but no column variable, which means
'HasTotal is False. Asserts that HasPercentage is also False because
'percentage computation requires a total row.
'@TestMethod("TableSpecs")
Public Sub TestHasPercentageTimeSeriesNoTotal()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasPercentageTimeSeriesNoTotal"
    On Error GoTo TestFail

    ' percentage=row but no column -> HasTotal=False -> HasPercentage=False
    Dim rows As Variant
    rows = Array(Array("S1", "date_v1", "", "no", "row", "no", "no", "", "", ""))
    BuildFixture "time series analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsFalse specs.HasPercentage, _
                   "Time series percentage needs HasTotal (which needs column variable)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasPercentageTimeSeriesNoTotal", Err.Number, Err.Description
End Sub

'@sub-title Verify spatio-temporal scope never reports HasPercentage.
'@details
'Builds a spatio-temporal fixture with percentage="row" in the data.
'Asserts that HasPercentage returns False regardless, because spatio-temporal
'tables do not support percentage display.
'@TestMethod("TableSpecs")
Public Sub TestHasPercentageSpatioTemporalAlwaysFalse()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasPercentageSpatioTemporalAlwaysFalse"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "date_v1", "geo_h2", "", "row", "", "yes", "", "", "5"))
    BuildFixture "spatio-temporal analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsFalse specs.HasPercentage, _
                   "SpatioTemporal should never have percentage"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasPercentageSpatioTemporalAlwaysFalse", Err.Number, Err.Description
End Sub

'@sub-title Verify bivariate HasPercentage is True with percentage=row.
'@details
'Builds a bivariate fixture with percentage="row" and both row and column
'variables present. Asserts that HasPercentage returns True.
'@TestMethod("TableSpecs")
Public Sub TestHasPercentageBivariateRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasPercentageBivariateRow"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "choi_v1", "choi_h2", "", "row", "", "", "", "", ""))
    BuildFixture "bivariate analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsTrue specs.HasPercentage, _
                  "Bivariate with percentage='row' should have percentage"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasPercentageBivariateRow", Err.Number, Err.Description
End Sub

'@section HasMissing tests
'===============================================================================

'@sub-title Verify time series HasMissing is True with missing=yes and column present.
'@details
'Uses the standard time series fixture row 1 where missing="yes" and
'column="choi_v1". Asserts that HasMissing returns True.
'@TestMethod("TableSpecs")
Public Sub TestHasMissingTimeSeriesWithColumn()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasMissingTimeSeriesWithColumn"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)  ' missing=yes, column=choi_v1

    Assert.IsTrue specs.HasMissing, _
                  "Time series with missing='yes' and column present should have missing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasMissingTimeSeriesWithColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify time series HasMissing is False when no column variable exists.
'@details
'Builds a fixture with missing="yes" but an empty column field. Asserts
'that HasMissing returns False because the missing row is only meaningful
'when a column variable is present to compute missing counts against.
'@TestMethod("TableSpecs")
Public Sub TestHasMissingTimeSeriesNoColumn()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasMissingTimeSeriesNoColumn"
    On Error GoTo TestFail

    ' missing=yes but no column
    Dim rows As Variant
    rows = Array(Array("S1", "date_v1", "", "no", "no", "yes", "no", "", "", ""))
    BuildFixture "time series analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsFalse specs.HasMissing, _
                   "Time series missing requires column variable"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasMissingTimeSeriesNoColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify bivariate HasMissing is True with missing=all.
'@details
'Builds a bivariate fixture with missing="all" and both row and column
'variables. Asserts that HasMissing returns True, confirming that the
'"all" keyword is accepted as a truthy missing value.
'@TestMethod("TableSpecs")
Public Sub TestHasMissingBivariateAll()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasMissingBivariateAll"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "choi_v1", "choi_h2", "", "", "all", "", "", "", ""))
    BuildFixture "bivariate analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsTrue specs.HasMissing, _
                  "Bivariate with missing='all' should have missing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasMissingBivariateAll", Err.Number, Err.Description
End Sub

'@section HasGraph tests
'===============================================================================

'@sub-title Verify time series HasGraph is True with graph=yes.
'@details
'Uses the standard time series fixture row 1 where graph="yes". Asserts
'that HasGraph returns True.
'@TestMethod("TableSpecs")
Public Sub TestHasGraphTimeSeriesYes()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasGraphTimeSeriesYes"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)  ' graph=yes

    Assert.IsTrue specs.HasGraph, _
                  "Time series with graph='yes' should have graph"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasGraphTimeSeriesYes", Err.Number, Err.Description
End Sub

'@sub-title Verify time series HasGraph is False with graph=no.
'@details
'Uses the standard time series fixture row 2 where graph="no". Asserts
'that HasGraph returns False.
'@TestMethod("TableSpecs")
Public Sub TestHasGraphTimeSeriesNo()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasGraphTimeSeriesNo"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(2)  ' graph=no

    Assert.IsFalse specs.HasGraph, _
                   "Time series with graph='no' should not have graph"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasGraphTimeSeriesNo", Err.Number, Err.Description
End Sub

'@sub-title Verify global summary scope never reports HasGraph.
'@details
'Builds a global summary fixture with graph="yes" in the data row. Asserts
'that HasGraph returns False regardless, because global summary tables do
'not support graph output.
'@TestMethod("TableSpecs")
Public Sub TestHasGraphGlobalSummaryAlwaysFalse()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasGraphGlobalSummaryAlwaysFalse"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("", "", "", "", "", "", "yes", "Count", "sum", ""))
    BuildFixture "global summary", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsFalse specs.HasGraph, _
                   "GlobalSummary should never have graph"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasGraphGlobalSummaryAlwaysFalse", Err.Number, Err.Description
End Sub

'@sub-title Verify bivariate HasGraph is True with graph=values.
'@details
'Builds a bivariate fixture with graph="values" and both row and column
'variables. Asserts that HasGraph returns True, confirming that "values"
'is accepted as a truthy graph setting for bivariate tables.
'@TestMethod("TableSpecs")
Public Sub TestHasGraphBivariateValues()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasGraphBivariateValues"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "choi_v1", "choi_h2", "", "", "", "values", "", "", ""))
    BuildFixture "bivariate analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsTrue specs.HasGraph, _
                  "Bivariate with graph='values' should have graph"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasGraphBivariateValues", Err.Number, Err.Description
End Sub

'@section Navigation tests (Previous, NextSpecs, TableSectionId)
'===============================================================================

'@sub-title Verify Previous returns the spec for the preceding data row.
'@details
'Builds a time series fixture and creates specs from data row 2 (same
'section S1 as row 1). Calls Previous and asserts that the returned
'ITableSpecs has TableId "TS_tab1", confirming backward navigation works.
'@TestMethod("TableSpecs")
Public Sub TestPreviousReturnsPriorRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestPreviousReturnsPriorRow"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(2)  ' Same section S1 as row 1

    Dim prevSpec As ITableSpecs
    Set prevSpec = specs.Previous

    Assert.IsTrue (Not prevSpec Is Nothing), _
                  "Previous should return a table spec"
    Assert.AreEqual "TS_tab1", prevSpec.TableId, _
                    "Previous of row 2 should be row 1"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreviousReturnsPriorRow", Err.Number, Err.Description
End Sub

'@sub-title Verify Previous raises an error on a new-section row.
'@details
'Builds a time series fixture and creates specs from data row 1, which is
'the first row and therefore a new section. Calls Previous under On Error
'Resume Next and asserts that a non-zero error number is raised, because
'there is no valid predecessor within the same section.
'@TestMethod("TableSpecs")
Public Sub TestPreviousThrowsOnNewSection()
    CustomTestSetTitles Assert, "TableSpecs", "TestPreviousThrowsOnNewSection"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)  ' First row = new section

    Dim prevSpec As ITableSpecs
    On Error Resume Next
    Set prevSpec = specs.Previous
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo 0

    Assert.IsTrue (errNum <> 0), _
                  "Previous on new section table should raise an error"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreviousThrowsOnNewSection", Err.Number, Err.Description
End Sub

'@sub-title Verify NextSpecs returns the spec for the following data row.
'@details
'Builds a time series fixture and creates specs from data row 1. Calls
'NextSpecs with an anchor range pointing to the last data row (row 3).
'Asserts that the returned ITableSpecs has TableId "TS_tab2", confirming
'forward navigation works within the anchor boundary.
'@TestMethod("TableSpecs")
Public Sub TestNextSpecsReturnsNextRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestNextSpecsReturnsNextRow"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    ' Anchor range = last data row (row 6 = row 3 + 3 data rows)
    Dim anchorRng As Range
    Set anchorRng = FixtureDataRange(3)

    Dim nextSpec As ITableSpecs
    Set nextSpec = specs.NextSpecs(anchorRng)

    Assert.IsTrue (Not nextSpec Is Nothing), _
                  "NextSpecs should return a table spec"
    Assert.AreEqual "TS_tab2", nextSpec.TableId, _
                    "NextSpecs of row 1 should be row 2"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNextSpecsReturnsNextRow", Err.Number, Err.Description
End Sub

'@sub-title Verify NextSpecs returns Nothing when beyond the anchor boundary.
'@details
'Builds a time series fixture and creates specs from the last data row
'(row 3). Sets the anchor range to the same row so there is no room for
'a next row. Asserts that NextSpecs returns Nothing.
'@TestMethod("TableSpecs")
Public Sub TestNextSpecsNothingBeyondAnchor()
    CustomTestSetTitles Assert, "TableSpecs", "TestNextSpecsNothingBeyondAnchor"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(3)  ' Last data row

    ' Anchor range = same row (no room for next)
    Dim anchorRng As Range
    Set anchorRng = FixtureDataRange(3)

    Dim nextSpec As ITableSpecs
    Set nextSpec = specs.NextSpecs(anchorRng)

    Assert.IsTrue (nextSpec Is Nothing), _
                  "NextSpecs beyond anchor should return Nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNextSpecsNothingBeyondAnchor", Err.Number, Err.Description
End Sub

'@sub-title Verify TableSectionId equals TableId for the first table in a section.
'@details
'Builds a time series fixture and creates specs from data row 1, which is
'a new section. Asserts that TableSectionId equals TableId because the
'first table in a section defines the section identifier.
'@TestMethod("TableSpecs")
Public Sub TestTableSectionIdFirstInSection()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableSectionIdFirstInSection"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)  ' New section

    Assert.AreEqual specs.TableId, specs.TableSectionId, _
                    "First table in section should be its own section id"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableSectionIdFirstInSection", Err.Number, Err.Description
End Sub

'@sub-title Verify TableSectionId inherits from the first table in the same section.
'@details
'Builds a time series fixture and creates specs from data row 2, which
'shares section "S1" with row 1. Asserts that TableSectionId is "TS_tab1"
'(the TableId of the section's first table), not "TS_tab2".
'@TestMethod("TableSpecs")
Public Sub TestTableSectionIdSubsequentInSection()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableSectionIdSubsequentInSection"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(2)  ' Same section S1

    Assert.AreEqual "TS_tab1", specs.TableSectionId, _
                    "Subsequent table should inherit section id from first table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableSectionIdSubsequentInSection", Err.Number, Err.Description
End Sub
