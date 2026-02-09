Attribute VB_Name = "TestTableSpecs"
Attribute VB_Description = "Tests for TableSpecs class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for TableSpecs class")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "TableSpecsFixture"

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
Private lDataStub As TableSpecsLinelistStub

'@section Helpers
'===============================================================================

' @description Build a fixture sheet with type label, header, and data rows.
'              Layout: Row 1 = type label, Row 2 = empty, Row 3 = header, Row 4+ = data.
' @param tableTypeName The analysis type label (e.g. "time series analysis")
' @param dataRows Array of row arrays, each containing NUM_COLUMNS values
Private Sub BuildFixture(ByVal tableTypeName As String, dataRows As Variant)
    Dim sh As Worksheet
    Dim headerArray As Variant
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    ' Row 1: type label (read by TableType via hRng.Cells(-1, 1))
    sh.Cells(1, 1).Value = tableTypeName

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

' @description Get the header range (row 3, NUM_COLUMNS wide)
Private Function FixtureHeaderRange() As Range
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    Set FixtureHeaderRange = sh.Range(sh.Cells(3, 1), sh.Cells(3, NUM_COLUMNS))
End Function

' @description Get a data row range (row 3 + offset, NUM_COLUMNS wide)
' @param dataRowIndex 1-based index of the data row (1 = first data row at row 4)
Private Function FixtureDataRange(ByVal dataRowIndex As Long) As Range
    Dim sh As Worksheet
    Dim rowNum As Long

    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    rowNum = 3 + dataRowIndex
    Set FixtureDataRange = sh.Range(sh.Cells(rowNum, 1), sh.Cells(rowNum, NUM_COLUMNS))
End Function

' @description Create an ITableSpecs from a fixture data row index
Private Function CreateSpecs(ByVal dataRowIndex As Long) As ITableSpecs
    Set CreateSpecs = TableSpecs.Create( _
        FixtureHeaderRange(), _
        FixtureDataRange(dataRowIndex), _
        lDataStub)
End Function

' @description Standard time series data rows for most tests.
'              Row 1: S1, date_var, choice_var, yes, row, yes, yes (all flags on)
'              Row 2: S1, date_var, choice_var, no, no, no, no (same section, all flags off)
'              Row 3: S2, date_var, "", no, no, no, no (new section, no column)
Private Function TimeSeriesDataRows() As Variant
    TimeSeriesDataRows = Array( _
        Array("S1", "date_var", "choice_var", "yes", "row", "yes", "yes", "", "", ""), _
        Array("S1", "date_var", "choice_var", "no", "no", "no", "no", "", "", ""), _
        Array("S2", "date_var", "", "no", "no", "no", "no", "", "", ""))
End Function

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestTableSpecs"
    Set lDataStub = New TableSpecsLinelistStub
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet FIXTURE_SHEET
    RestoreApp
    Set lDataStub = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory validation tests
'===============================================================================

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
    Set specs = TableSpecs.Create(Nothing, dataRng, lDataStub)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "Create with Nothing header should fail (return Nothing)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingHeader", Err.Number, Err.Description
End Sub

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
    Set specs = TableSpecs.Create(hRng, Nothing, lDataStub)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "Create with Nothing data range should fail (return Nothing)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingRange", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsNothingLData()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsNothingLData"
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
                  "Create with Nothing lData should fail (return Nothing)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingLData", Err.Number, Err.Description
End Sub

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
    Set specs = TableSpecs.Create(hRng, dataRng, lDataStub)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "Create with mismatched column count should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsMismatchedColumns", Err.Number, Err.Description
End Sub

'@section TableType tests
'===============================================================================

'@TestMethod("TableSpecs")
Public Sub TestTableTypeTimeSeries()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableTypeTimeSeries"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(TypeTimeSeries), CLng(specs.TableType), _
                    "Should parse 'time series analysis' as TypeTimeSeries"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableTypeTimeSeries", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestTableTypeGlobalSummary()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableTypeGlobalSummary"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("", "", "", "", "", "", "", "Count of cases", "sum", ""))
    BuildFixture "global summary", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(TypeGlobalSummary), CLng(specs.TableType), _
                    "Should parse 'global summary' as TypeGlobalSummary"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableTypeGlobalSummary", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestTableTypeUnivariate()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableTypeUnivariate"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "var1", "", "", "yes", "yes", "yes", "", "", ""))
    BuildFixture "univariate analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(TypeUnivariate), CLng(specs.TableType), _
                    "Should parse 'univariate analysis' as TypeUnivariate"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableTypeUnivariate", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestTableTypeBivariate()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableTypeBivariate"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "var1", "var2", "", "row", "row", "values", "", "", ""))
    BuildFixture "bivariate analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(TypeBivariate), CLng(specs.TableType), _
                    "Should parse 'bivariate analysis' as TypeBivariate"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableTypeBivariate", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestTableTypeSpatial()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableTypeSpatial"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "geo_var", "choice_var", "", "yes", "yes", "yes", "", "", ""))
    BuildFixture "spatial analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(TypeSpatial), CLng(specs.TableType), _
                    "Should parse 'spatial analysis' as TypeSpatial"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableTypeSpatial", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestTableTypeSpatioTemporal()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableTypeSpatioTemporal"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "date_var", "geo_var", "", "", "", "yes", "", "", "5"))
    BuildFixture "spatio-temporal analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(TypeSpatioTemporal), CLng(specs.TableType), _
                    "Should parse 'spatio-temporal analysis' as TypeSpatioTemporal"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableTypeSpatioTemporal", Err.Number, Err.Description
End Sub

'@section TableId tests
'===============================================================================

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

'@TestMethod("TableSpecs")
Public Sub TestValueReturnsColumnData()
    CustomTestSetTitles Assert, "TableSpecs", "TestValueReturnsColumnData"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual "S1", specs.Value("section"), _
                    "Value('section') should return the section name"
    Assert.AreEqual "date_var", specs.Value("row"), _
                    "Value('row') should return the row variable"
    Assert.AreEqual "choice_var", specs.Value("column"), _
                    "Value('column') should return the column variable"
    Assert.AreEqual "yes", specs.Value("total"), _
                    "Value('total') should return 'yes'"
    Assert.AreEqual "row", specs.Value("percentage"), _
                    "Value('percentage') should return 'row'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueReturnsColumnData", Err.Number, Err.Description
End Sub

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

'@TestMethod("TableSpecs")
Public Sub TestHasTotalTimeSeriesWithTotalYes()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalTimeSeriesWithTotalYes"
    On Error GoTo TestFail

    ' Row with total=yes and column variable present
    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)  ' total=yes, column=choice_var

    Assert.IsTrue specs.HasTotal, _
                  "Time series with total='yes' and column present should have total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalTimeSeriesWithTotalYes", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestHasTotalTimeSeriesPercentageDriven()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalTimeSeriesPercentageDriven"
    On Error GoTo TestFail

    ' Row with total=no but percentage=row and column present
    Dim rows As Variant
    rows = Array(Array("S1", "date_var", "choice_var", "no", "row", "no", "no", "", "", ""))
    BuildFixture "time series analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsTrue specs.HasTotal, _
                  "Time series with percentage='row' should have total for computation"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalTimeSeriesPercentageDriven", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestHasTotalTimeSeriesPercentageColumnDriven()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalTimeSeriesPercentageColumnDriven"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "date_var", "choice_var", "no", "column", "no", "no", "", "", ""))
    BuildFixture "time series analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsTrue specs.HasTotal, _
                  "Time series with percentage='column' should have total for computation"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalTimeSeriesPercentageColumnDriven", Err.Number, Err.Description
End Sub

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

'@TestMethod("TableSpecs")
Public Sub TestHasTotalUnivariateAlwaysTrue()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalUnivariateAlwaysTrue"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "var1", "", "no", "no", "no", "no", "", "", ""))
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

'@TestMethod("TableSpecs")
Public Sub TestTotalRequestedFalseWhenPercentageDriven()
    CustomTestSetTitles Assert, "TableSpecs", "TestTotalRequestedFalseWhenPercentageDriven"
    On Error GoTo TestFail

    ' HasTotal will be True (percentage=row + column present) but TotalRequested should be False
    Dim rows As Variant
    rows = Array(Array("S1", "date_var", "choice_var", "no", "row", "no", "no", "", "", ""))
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

'@TestMethod("TableSpecs")
Public Sub TestHasPercentageTimeSeriesNoTotal()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasPercentageTimeSeriesNoTotal"
    On Error GoTo TestFail

    ' percentage=row but no column → HasTotal=False → HasPercentage=False
    Dim rows As Variant
    rows = Array(Array("S1", "date_var", "", "no", "row", "no", "no", "", "", ""))
    BuildFixture "time series analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsFalse specs.HasPercentage, _
                   "Time series percentage needs HasTotal (which needs column variable)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasPercentageTimeSeriesNoTotal", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestHasPercentageSpatioTemporalAlwaysFalse()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasPercentageSpatioTemporalAlwaysFalse"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "date_var", "geo_var", "", "row", "", "yes", "", "", "5"))
    BuildFixture "spatio-temporal analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsFalse specs.HasPercentage, _
                   "SpatioTemporal should never have percentage"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasPercentageSpatioTemporalAlwaysFalse", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestHasPercentageBivariateRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasPercentageBivariateRow"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "var1", "var2", "", "row", "", "", "", "", ""))
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

'@TestMethod("TableSpecs")
Public Sub TestHasMissingTimeSeriesWithColumn()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasMissingTimeSeriesWithColumn"
    On Error GoTo TestFail

    BuildFixture "time series analysis", TimeSeriesDataRows()
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)  ' missing=yes, column=choice_var

    Assert.IsTrue specs.HasMissing, _
                  "Time series with missing='yes' and column present should have missing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasMissingTimeSeriesWithColumn", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestHasMissingTimeSeriesNoColumn()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasMissingTimeSeriesNoColumn"
    On Error GoTo TestFail

    ' missing=yes but no column
    Dim rows As Variant
    rows = Array(Array("S1", "date_var", "", "no", "no", "yes", "no", "", "", ""))
    BuildFixture "time series analysis", rows

    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsFalse specs.HasMissing, _
                   "Time series missing requires column variable"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasMissingTimeSeriesNoColumn", Err.Number, Err.Description
End Sub

'@TestMethod("TableSpecs")
Public Sub TestHasMissingBivariateAll()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasMissingBivariateAll"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "var1", "var2", "", "", "all", "", "", "", ""))
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

'@TestMethod("TableSpecs")
Public Sub TestHasGraphBivariateValues()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasGraphBivariateValues"
    On Error GoTo TestFail

    Dim rows As Variant
    rows = Array(Array("S1", "var1", "var2", "", "", "", "values", "", "", ""))
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
