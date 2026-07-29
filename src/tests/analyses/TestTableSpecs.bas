Attribute VB_Name = "TestTableSpecs"
Attribute VB_Description = "Tests for TableSpecs class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for TableSpecs class")

'@description
'Validates the TableSpecs class, which parses a single row from an analysis
'specification table and exposes computed properties such as TableScope,
'TableId, HasTotal, HasPercentage, HasMissing, HasGraph, GeoCount,
'SpatialTableScopes and section navigation (IsNewSection, Previous,
'NextSpecs, TableSectionId).
'
'THE FIXTURE IS A REAL LISTOBJECT WITH THE REAL HEADER
'-------------------------------------------------------------------------------
'The scope of a specification row is the name of the ListObject it sits in,
'so every fixture here builds a real ListObject and names it the way the
'setup workbook names it. The seven header rows are the ones measured on
'2026-07-29 in .mock/setup_mock.xlsb, src/bin/setup/setup.xlsb and
'releases/latest/OBT-main-latest/setup_main-2026-06-11.xlsb, which all agree.
'An earlier fixture wrote a type label four rows above the header, which
'reproduced the mock layout and passed against a workbook shape the field
'never sees.
'
'THE FLAG VOCABULARIES ARE COPIED FROM SetupPreparation
'-------------------------------------------------------------------------------
'TestFlagVocabularyContract drives every entry of each registered dropdown
'through the flag property bound to it. The vocabularies are written out
'below, taken from SetupPreparation.RegisterAllDropdowns and its
'SetValidation calls. A vocabulary changed there needs the copy here changed
'with it; the test then says which flag stopped answering.
'@depends TableSpecs, LLdictionary, CustomTest, TestHelpersLite, DictionaryTestFixture

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "TableSpecsFixture"
Private Const DICT_SHEET As String = "TableSpecsDict"

' The header row of every fixture table. Data rows start immediately below,
' so data row N sits at HEADER_ROW + N and TableId reads "<prefix>_tabN".
Private Const HEADER_ROW As Long = 5

' The seven analysis ListObject names, spelled the way the setup workbook
' spells them.
Private Const TABLE_GLOBAL_SUMMARY As String = "Tab_Global_Summary"
Private Const TABLE_UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const TABLE_BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TABLE_TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const TABLE_SPATIAL As String = "Tab_Spatial_Analysis"
Private Const TABLE_SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const TABLE_GRAPH_TIMESERIES As String = "Tab_Graph_TimeSeries"

' A geo variable the shared dictionary fixture has no equivalent of.
' AppendGeoLines writes adm1_ rows during Prepare, and the fixture is a raw
' dictionary, so this row is added by ModuleInitialize.
Private Const GEO_ROW_VARIABLE As String = "zone"
Private Const GEO_PREFIXED_VARIABLE As String = "adm1_zone"

Private Assert As CustomTest
Private dict As LLdictionary

'@section Fixture headers
'===============================================================================
'@description The seven analysis header rows, measured in the setup workbook.

'@sub-title Header of Tab_Global_Summary, 3 columns.
Private Function GlobalSummaryHeader() As Variant
    GlobalSummaryHeader = Array("Summary label", "Summary function", "Format")
End Function

'@sub-title Header of Tab_Univariate_Analysis, 10 columns.
Private Function UnivariateHeader() As Variant
    UnivariateHeader = Array( _
        "Section", "Table title", "Group by variable (row)", "Add missing data", _
        "Summary function", "Summary label", "Format", "Add percentage", _
        "Add graph", "Flip coordinates")
End Function

'@sub-title Header of Tab_Bivariate_Analysis, 11 columns.
Private Function BivariateHeader() As Variant
    BivariateHeader = Array( _
        "Section", "Table title", "Group by variable (row)", _
        "Group by variable (column)", "Add missing data", "Summary function", _
        "Summary label", "Format", "Add percentage", "Add graph", _
        "Flip coordinates")
End Function

'@sub-title Header of Tab_TimeSeries_Analysis, 12 columns.
'@details
'This block carries no column holding the word "graph", which is why
'HasGraph answers False for every time series row. Time series charts come
'from Tab_Graph_TimeSeries.
Private Function TimeSeriesHeader() As Variant
    TimeSeriesHeader = Array( _
        "Series ID", "Section", "Time variable (row)", _
        "Group by variable (column)", "Title (header)", "Add missing data", _
        "Summary function", "Summary label", "Format", "Add percentage", _
        "Add total", "Table order")
End Function

'@sub-title Header of Tab_Graph_TimeSeries, 12 columns.
Private Function GraphTimeSeriesHeader() As Variant
    GraphTimeSeriesHeader = Array( _
        "Graph title (select)", "Series title (select)", "Graph ID", _
        "Series ID", "Graph order", "Time variable (row)", _
        "Group by variable (column)", "Choices", "Label", _
        "Plot values or percentages", "Chart type", "Y-Axis")
End Function

'@sub-title Header of Tab_Spatial_Analysis, 12 columns.
Private Function SpatialHeader() As Variant
    SpatialHeader = Array( _
        "Section", "Table title", "Geo/HF variable (row)", "N geo max", _
        "Group by variable (column)", "Add missing data", "Summary function", _
        "Summary label", "Format", "Add percentage", "Add graph", _
        "Flip coordinates")
End Function

'@sub-title Header of Tab_SpatioTemporal_Analysis, 10 columns.
Private Function SpatioTemporalHeader() As Variant
    SpatioTemporalHeader = Array( _
        "Section (select)", "Time variable (row)", "Geo/HF variable (column)", _
        "N geo max", "Title (header)", "Spatial type", "Summary function", _
        "Summary label", "Format", "Add graph")
End Function

'@section Fixture helpers
'===============================================================================

'@sub-title Drop every ListObject on the fixture sheet.
'@details
'A ListObject name is unique across the workbook, so the table built by the
'previous test has to go before the next one takes the same name. Unlist
'turns the table back into an ordinary range and frees the name.
'@param sh Worksheet. The fixture worksheet.
Private Sub RemoveFixtureTables(ByVal sh As Worksheet)
    Dim idx As Long

    For idx = sh.ListObjects.Count To 1 Step -1
        sh.ListObjects(idx).Unlist
    Next idx
End Sub

'@sub-title Build a fixture ListObject with a header row and data rows.
'@details
'Clears the fixture sheet, writes the header at HEADER_ROW and the data rows
'below it, then wraps both in a ListObject carrying the name the setup
'workbook uses for that analysis block. TableSpecs reads the scope from that
'name.
'@param tableName String. The ListObject name, one of the seven Tab_ names.
'@param headerRow Variant. A one-dimensional array of column names.
'@param dataRows Variant. An array of row arrays, each as wide as headerRow.
Private Sub BuildFixture(ByVal tableName As String, _
                         ByVal headerRow As Variant, _
                         ByVal dataRows As Variant)
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim tableRng As Range
    Dim rowCount As Long
    Dim colCount As Long

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)
    RemoveFixtureTables sh
    sh.Cells.Clear

    colCount = UBound(headerRow) - LBound(headerRow) + 1
    rowCount = UBound(dataRows) - LBound(dataRows) + 1

    WriteMatrix sh.Cells(HEADER_ROW, 1), RowsToMatrix(Array(headerRow))
    WriteMatrix sh.Cells(HEADER_ROW + 1, 1), RowsToMatrix(dataRows)

    Set tableRng = sh.Range(sh.Cells(HEADER_ROW, 1), _
                            sh.Cells(HEADER_ROW + rowCount, colCount))
    Set lo = sh.ListObjects.Add(xlSrcRange, tableRng, , xlYes)
    lo.Name = tableName
End Sub

'@sub-title Return the fixture header range.
Private Function FixtureHeaderRange() As Range
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    Set FixtureHeaderRange = sh.ListObjects(1).HeaderRowRange
End Function

'@sub-title Return a data row range by 1-based index.
Private Function FixtureDataRange(ByVal dataRowIndex As Long) As Range
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    Set FixtureDataRange = sh.ListObjects(1).ListRows(dataRowIndex).Range
End Function

'@sub-title Return the cell of one column on one fixture data row.
'@param dataRowIndex Long. The 1-based data row.
'@param columnIndex Long. The 1-based column within the table.
Private Function FixtureCell(ByVal dataRowIndex As Long, _
                             ByVal columnIndex As Long) As Range
    Set FixtureCell = FixtureDataRange(dataRowIndex).Cells(1, columnIndex)
End Function

'@sub-title Create a TableSpecs from a fixture data row index.
Private Function CreateSpecs(ByVal dataRowIndex As Long) As TableSpecs
    Set CreateSpecs = TableSpecs.Create( _
        FixtureHeaderRange(), _
        FixtureDataRange(dataRowIndex), _
        dict)
End Function

'@sub-title The standard three-row time series fixture data.
'@details
'Row 1: section S1, date_v1 by choi_v1, total=yes, percentage=row,
'missing=yes. Row 2: same section, every flag off. Row 3: section S2, no
'column variable, every flag off. All three are valid tables, because
'date_v1 is a date variable and choi_v1 is a choice variable.
Private Function TimeSeriesDataRows() As Variant
    TimeSeriesDataRows = Array( _
        Array("Series 1", "S1", "date_v1", "choi_v1", "First table", "yes", _
              "", "", "", "row", "yes", "1"), _
        Array("Series 2", "S1", "date_v1", "choi_v1", "Second table", "no", _
              "", "", "", "no", "no", "2"), _
        Array("Series 3", "S2", "date_v1", "", "Third table", "no", _
              "", "", "", "no", "no", "3"))
End Function

'@sub-title One univariate data row with the flags the caller passes.
Private Function UnivariateRow(ByVal rowVar As String, _
                               ByVal missing As String, _
                               ByVal percentage As String, _
                               ByVal graph As String) As Variant
    UnivariateRow = Array( _
        Array("S1", "A univariate table", rowVar, missing, "", "", "", _
              percentage, graph, "no"))
End Function

'@sub-title One bivariate data row with the flags the caller passes.
Private Function BivariateRow(ByVal rowVar As String, _
                              ByVal colVar As String, _
                              ByVal missing As String, _
                              ByVal percentage As String, _
                              ByVal graph As String) As Variant
    BivariateRow = Array( _
        Array("S1", "A bivariate table", rowVar, colVar, missing, "", "", "", _
              percentage, graph, "no"))
End Function

'@sub-title One spatial data row with the flags the caller passes.
Private Function SpatialRow(ByVal rowVar As String, _
                            ByVal colVar As String, _
                            ByVal geoMax As String, _
                            ByVal missing As String, _
                            ByVal percentage As String, _
                            ByVal graph As String) As Variant
    SpatialRow = Array( _
        Array("S1", "A spatial table", rowVar, geoMax, colVar, missing, "", _
              "", "", percentage, graph, "no"))
End Function

'@sub-title One spatio-temporal data row.
Private Function SpatioTemporalRow(ByVal rowVar As String, _
                                   ByVal colVar As String, _
                                   ByVal geoMax As String, _
                                   ByVal graph As String) As Variant
    SpatioTemporalRow = Array( _
        Array("S1", rowVar, colVar, geoMax, "A spatio-temporal table", "geo", _
              "", "", "", graph))
End Function

'@sub-title One graph data row.
Private Function GraphRow(ByVal graphTitle As String, _
                          ByVal seriesTitle As String, _
                          ByVal plotValues As String) As Variant
    GraphRow = Array( _
        Array(graphTitle, seriesTitle, "G1", "Series 1", "1", "date_v1", _
              "choi_v1", "", "A label", plotValues, "line", "left"))
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Set up module-level fixtures for all TableSpecs tests.
'@details
'Suppresses screen updating, ensures the test output sheet exists, creates
'the CustomTest assert object, prepares a dictionary fixture sheet with
'known variable definitions, appends the one geo variable the shared fixture
'has no equivalent of, and wraps the sheet in an LLdictionary instance used
'by all tests. This routine is Public because the harness calls it by name
'through Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    Dim sh As Worksheet
    Dim appendRow As Long

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestTableSpecs"

    PrepareDictionaryFixture DICT_SHEET

    ' adm1_zone gives SpatialTableScopes a "geo" answer to find, and gives
    ' the spatial ValidTable rule a geo variable to accept.
    Set sh = ThisWorkbook.Worksheets(DICT_SHEET)
    appendRow = 1 + DictionaryFixtureRowCount() + 1
    sh.Cells(appendRow, 1).Value = GEO_PREFIXED_VARIABLE

    Set dict = LLdictionary.Create(sh, 1, 1)
End Sub

'@sub-title Print results and tear down module-level fixtures.
'@details
'Prints accumulated test results to the output sheet, deletes both the
'fixture and dictionary worksheets, restores Excel application state, and
'releases object references. This routine is Public because the harness
'calls it by name through Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
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

'@sub-title Verify Create raises InvalidArgument when the header range is Nothing.
'@details
'Passes Nothing as the header range and asserts on the error number. An
'assertion on "specs Is Nothing" alone passes whenever Create fails for any
'reason at all, including the reasons this test is meant to rule out.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsNothingHeader()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsNothingHeader"
    On Error GoTo TestFail

    Dim specs As TableSpecs
    Dim errNumber As Long

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Dim dataRng As Range
    Set dataRng = FixtureDataRange(1)

    On Error Resume Next
    Set specs = TableSpecs.Create(Nothing, dataRng, dict)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                    "A Nothing header range should raise InvalidArgument"
    Assert.IsTrue (specs Is Nothing), _
                  "Create with a Nothing header should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingHeader", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises InvalidArgument when the data range is Nothing.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsNothingRange()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsNothingRange"
    On Error GoTo TestFail

    Dim specs As TableSpecs
    Dim errNumber As Long

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Dim hRng As Range
    Set hRng = FixtureHeaderRange()

    On Error Resume Next
    Set specs = TableSpecs.Create(hRng, Nothing, dict)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                    "A Nothing specification range should raise InvalidArgument"
    Assert.IsTrue (specs Is Nothing), _
                  "Create with a Nothing data range should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingRange", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises InvalidArgument when the dictionary is Nothing.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsNothingDict()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsNothingDict"
    On Error GoTo TestFail

    Dim specs As TableSpecs
    Dim errNumber As Long

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Dim hRng As Range
    Dim dataRng As Range
    Set hRng = FixtureHeaderRange()
    Set dataRng = FixtureDataRange(1)

    On Error Resume Next
    Set specs = TableSpecs.Create(hRng, dataRng, Nothing)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                    "A Nothing dictionary should raise InvalidArgument"
    Assert.IsTrue (specs Is Nothing), _
                  "Create with a Nothing dictionary should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingDict", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises InvalidArgument on mismatched column counts.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsMismatchedColumns()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsMismatchedColumns"
    On Error GoTo TestFail

    Dim specs As TableSpecs
    Dim errNumber As Long

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Dim hRng As Range
    Dim dataRng As Range
    Set hRng = FixtureHeaderRange()
    Set dataRng = FixtureDataRange(1)
    Set dataRng = dataRng.Resize(1, 5)

    On Error Resume Next
    Set specs = TableSpecs.Create(hRng, dataRng, dict)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                    "Two ranges of different widths should raise InvalidArgument"
    Assert.IsTrue (specs Is Nothing), _
                  "Create with mismatched widths should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsMismatchedColumns", Err.Number, Err.Description
End Sub

'@sub-title Verify Create rejects ranges that start at different columns.
'@details
'ComputeIsNewSection resolves the section index against the header range and
'then indexes the specification range with it. Two ranges starting at
'different columns make every section read land on the wrong column, so
'Create refuses the pair.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsColumnOffsetMismatch()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsColumnOffsetMismatch"
    On Error GoTo TestFail

    Dim specs As TableSpecs
    Dim errNumber As Long
    Dim sh As Worksheet

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)

    Dim hRng As Range
    Dim dataRng As Range
    Set hRng = sh.Range(sh.Cells(HEADER_ROW, 1), sh.Cells(HEADER_ROW, 5))
    Set dataRng = sh.Range(sh.Cells(HEADER_ROW + 1, 2), sh.Cells(HEADER_ROW + 1, 6))

    On Error Resume Next
    Set specs = TableSpecs.Create(hRng, dataRng, dict)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                    "Ranges starting at different columns should raise InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsColumnOffsetMismatch", Err.Number, Err.Description
End Sub

'@sub-title Verify Create rejects a specification range spanning several rows.
'@details
'Every read in the class assumes a single row, so a two-row range would be
'read as its first row and the rest would go missing without a word.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsMultiRowSpecification()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsMultiRowSpecification"
    On Error GoTo TestFail

    Dim specs As TableSpecs
    Dim errNumber As Long

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Dim hRng As Range
    Dim dataRng As Range
    Set hRng = FixtureHeaderRange()
    Set dataRng = FixtureDataRange(1).Resize(2)

    On Error Resume Next
    Set specs = TableSpecs.Create(hRng, dataRng, dict)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                    "A two-row specification range should raise InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsMultiRowSpecification", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises when the row sits outside any ListObject.
'@details
'The scope of a row is the name of the table it sits in, so a row on a bare
'worksheet has no scope to resolve and Create says so.
'@TestMethod("TableSpecs")
Public Sub TestCreateRejectsRowOutsideAnalysisTable()
    CustomTestSetTitles Assert, "TableSpecs", "TestCreateRejectsRowOutsideAnalysisTable"
    On Error GoTo TestFail

    Dim specs As TableSpecs
    Dim errNumber As Long
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)
    RemoveFixtureTables sh
    sh.Cells.Clear
    WriteMatrix sh.Cells(HEADER_ROW, 1), RowsToMatrix(Array(TimeSeriesHeader()))

    Dim hRng As Range
    Dim dataRng As Range
    Set hRng = sh.Range(sh.Cells(HEADER_ROW, 1), sh.Cells(HEADER_ROW, 12))
    Set dataRng = sh.Range(sh.Cells(HEADER_ROW + 1, 1), sh.Cells(HEADER_ROW + 1, 12))

    On Error Resume Next
    Set specs = TableSpecs.Create(hRng, dataRng, dict)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.ErrorUnexpectedState, errNumber, _
                    "A row outside a ListObject should raise ErrorUnexpectedState"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsRowOutsideAnalysisTable", Err.Number, Err.Description
End Sub

'@sub-title Verify InitCache refuses to run a second time.
'@details
'InitCache builds the column and value caches that every derived answer in
'the class is computed from, so a second call would rebuild half the state
'and leave the other half pointing at the old data.
'@TestMethod("TableSpecs")
Public Sub TestInitCacheRefusesASecondCall()
    CustomTestSetTitles Assert, "TableSpecs", "TestInitCacheRefusesASecondCall"
    On Error GoTo TestFail

    Dim specs As TableSpecs
    Dim errNumber As Long

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Set specs = CreateSpecs(1)

    On Error Resume Next
    specs.InitCache
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.SomethingWentWrong, errNumber, _
                    "A second InitCache call should raise SomethingWentWrong"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestInitCacheRefusesASecondCall", Err.Number, Err.Description
End Sub

'@section TableScope tests
'===============================================================================

'@sub-title Verify each analysis ListObject name resolves to its own scope.
'@details
'Builds all seven analysis blocks in turn, each as a real ListObject with
'the name the setup workbook uses, and asserts the scope. Tab_Graph_TimeSeries
'is the one that used to raise: the scope came from a label reading
'"Graph on Time Series" while the code looked for "time series graphs", so
'Create threw for every row of that block and its checks validated nothing.
'@TestMethod("TableSpecs")
Public Sub TestTableScopeComesFromTheListObjectName()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableScopeComesFromTheListObjectName"
    On Error GoTo TestFail

    Dim specs As TableSpecs

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 Array(Array("Count of cases", "sum", ""))
    Set specs = CreateSpecs(1)
    Assert.AreEqual CLng(ScopeGlobalSummary), CLng(specs.TableScope), _
                    "Tab_Global_Summary should resolve to ScopeGlobalSummary"

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("choi_v1", "yes", "yes", "yes")
    Set specs = CreateSpecs(1)
    Assert.AreEqual CLng(ScopeUnivariate), CLng(specs.TableScope), _
                    "Tab_Univariate_Analysis should resolve to ScopeUnivariate"

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRow("choi_v1", "choi_h2", "row", "row", "values")
    Set specs = CreateSpecs(1)
    Assert.AreEqual CLng(ScopeBivariate), CLng(specs.TableScope), _
                    "Tab_Bivariate_Analysis should resolve to ScopeBivariate"

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Set specs = CreateSpecs(1)
    Assert.AreEqual CLng(ScopeTimeSeries), CLng(specs.TableScope), _
                    "Tab_TimeSeries_Analysis should resolve to ScopeTimeSeries"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow(GEO_ROW_VARIABLE, "choi_v1", "5", "yes", "yes", "yes")
    Set specs = CreateSpecs(1)
    Assert.AreEqual CLng(ScopeSpatial), CLng(specs.TableScope), _
                    "Tab_Spatial_Analysis should resolve to ScopeSpatial"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "5", "yes")
    Set specs = CreateSpecs(1)
    Assert.AreEqual CLng(ScopeSpatioTemporal), CLng(specs.TableScope), _
                    "Tab_SpatioTemporal_Analysis should resolve to ScopeSpatioTemporal"

    BuildFixture TABLE_GRAPH_TIMESERIES, GraphTimeSeriesHeader(), _
                 GraphRow("A graph", "A series", "values")
    Set specs = CreateSpecs(1)
    Assert.AreEqual CLng(ScopeTimeSeriesGraph), CLng(specs.TableScope), _
                    "Tab_Graph_TimeSeries should resolve to ScopeTimeSeriesGraph"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableScopeComesFromTheListObjectName", Err.Number, Err.Description
End Sub

'@section TableId tests
'===============================================================================

'@sub-title Verify every scope has its own TableId prefix.
'@details
'Seven scopes, seven prefixes. ScopeTimeSeriesGraph had none, so a graph row
'would have produced the id "_tab1" the moment the scope became reachable.
'@TestMethod("TableSpecs")
Public Sub TestTableIdPrefixPerScope()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableIdPrefixPerScope"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 Array(Array("Count of cases", "sum", ""))
    Assert.AreEqual "GS_tab1", CreateSpecs(1).TableId, _
                    "Global summary table id should start with GS"

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("choi_v1", "no", "no", "no")
    Assert.AreEqual "UA_tab1", CreateSpecs(1).TableId, _
                    "Univariate table id should start with UA"

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRow("choi_v1", "choi_h2", "no", "no", "no")
    Assert.AreEqual "BA_tab1", CreateSpecs(1).TableId, _
                    "Bivariate table id should start with BA"

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Assert.AreEqual "TS_tab1", CreateSpecs(1).TableId, _
                    "Time series table id should start with TS"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow(GEO_ROW_VARIABLE, "", "5", "no", "no", "no")
    Assert.AreEqual "SA_tab1", CreateSpecs(1).TableId, _
                    "Spatial table id should start with SA"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "5", "no")
    Assert.AreEqual "SPT_tab1", CreateSpecs(1).TableId, _
                    "Spatio-temporal table id should start with SPT"

    BuildFixture TABLE_GRAPH_TIMESERIES, GraphTimeSeriesHeader(), _
                 GraphRow("A graph", "A series", "values")
    Assert.AreEqual "TSG_tab1", CreateSpecs(1).TableId, _
                    "Time series graph table id should start with TSG"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableIdPrefixPerScope", Err.Number, Err.Description
End Sub

'@sub-title Verify TableId counts rows from the header row.
'@details
'The first data row sits one row below the header, so its id ends in tab1
'and the second row's ends in tab2.
'@TestMethod("TableSpecs")
Public Sub TestTableIdUsesTheRowOffset()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableIdUsesTheRowOffset"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Assert.AreEqual "TS_tab1", CreateSpecs(1).TableId, _
                    "The first data row should be tab1"
    Assert.AreEqual "TS_tab2", CreateSpecs(2).TableId, _
                    "The second data row should be tab2"
    Assert.AreEqual "TS_tab3", CreateSpecs(3).TableId, _
                    "The third data row should be tab3"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableIdUsesTheRowOffset", Err.Number, Err.Description
End Sub

'@section Value tests
'===============================================================================

'@sub-title Verify Value returns the cell content of each named column.
'@TestMethod("TableSpecs")
Public Sub TestValueReturnsColumnData()
    CustomTestSetTitles Assert, "TableSpecs", "TestValueReturnsColumnData"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Dim specs As TableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual "S1", specs.Value("section"), _
                    "Value('section') should read the Section column"
    Assert.AreEqual "date_v1", specs.Value("row"), _
                    "Value('row') should read Time variable (row)"
    Assert.AreEqual "choi_v1", specs.Value("column"), _
                    "Value('column') should read Group by variable (column)"
    Assert.AreEqual "yes", specs.Value("total"), _
                    "Value('total') should read Add total"
    Assert.AreEqual "row", specs.Value("percentage"), _
                    "Value('percentage') should read Add percentage"
    Assert.AreEqual "yes", specs.Value("missing"), _
                    "Value('missing') should read Add missing data"
    Assert.AreEqual "First table", specs.Value("title"), _
                    "Value('title') should read Title (header)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueReturnsColumnData", Err.Number, Err.Description
End Sub

'@sub-title Verify Value returns an empty string for an unknown column name.
'@TestMethod("TableSpecs")
Public Sub TestValueReturnsEmptyForUnknownColumn()
    CustomTestSetTitles Assert, "TableSpecs", "TestValueReturnsEmptyForUnknownColumn"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Dim specs As TableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual vbNullString, specs.Value("nonexistent_column"), _
                    "A column absent from the header should read as empty"
    Assert.AreEqual vbNullString, specs.Value("graph"), _
                    "The time series block has no column holding the word graph"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueReturnsEmptyForUnknownColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify the default search takes the first header holding the term.
'@details
'"row" is a substring of both "Table row order" and "Time variable (row)",
'and the default search takes whichever comes first. Adding a column to the
'left of the intended one silently rebinds every lookup, which is why
'strictSearch exists.
'@TestMethod("TableSpecs")
Public Sub TestValueSubstringSearchTakesTheFirstMatch()
    CustomTestSetTitles Assert, "TableSpecs", "TestValueSubstringSearchTakesTheFirstMatch"
    On Error GoTo TestFail

    Dim header As Variant
    header = Array("Series ID", "Section", "Table row order", _
                   "Time variable (row)", "Group by variable (column)", _
                   "Title (header)", "Add missing data", "Summary function", _
                   "Summary label", "Format", "Add percentage", "Add total")

    BuildFixture TABLE_TIMESERIES, header, _
                 Array(Array("Series 1", "S1", "7", "date_v1", "choi_v1", _
                             "A title", "no", "", "", "", "no", "no"))

    Dim specs As TableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual "7", specs.Value("row"), _
                    "The default search takes Table row order, the first match"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueSubstringSearchTakesTheFirstMatch", Err.Number, Err.Description
End Sub

'@sub-title Verify strictSearch compares against the whole header.
'@TestMethod("TableSpecs")
Public Sub TestValueStrictSearchMatchesTheWholeHeader()
    CustomTestSetTitles Assert, "TableSpecs", "TestValueStrictSearchMatchesTheWholeHeader"
    On Error GoTo TestFail

    Dim header As Variant
    header = Array("Series ID", "Section", "Table row order", _
                   "Time variable (row)", "Group by variable (column)", _
                   "Title (header)", "Add missing data", "Summary function", _
                   "Summary label", "Format", "Add percentage", "Add total")

    BuildFixture TABLE_TIMESERIES, header, _
                 Array(Array("Series 1", "S1", "7", "date_v1", "choi_v1", _
                             "A title", "no", "", "", "", "no", "no"))

    Dim specs As TableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual "date_v1", specs.Value("Time variable (row)", strictSearch:=True), _
                    "A strict search should reach the column it names"
    Assert.AreEqual vbNullString, specs.Value("row", strictSearch:=True), _
                    "A strict search for a partial name should find nothing"
    Assert.AreEqual "7", specs.Value("table row order", strictSearch:=True), _
                    "A strict search should ignore the case of the header"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueStrictSearchMatchesTheWholeHeader", Err.Number, Err.Description
End Sub

'@sub-title Verify an error-valued cell reads as an empty string.
'@details
'The analysis sheet carries formula columns, so a cell can hold #DIV/0! after
'an ordinary edit. Coercing that to text used to raise a type mismatch inside
'the factory, which aborted the whole analysis run.
'@TestMethod("TableSpecs")
Public Sub TestErrorValuedCellReadsAsEmpty()
    CustomTestSetTitles Assert, "TableSpecs", "TestErrorValuedCellReadsAsEmpty"
    On Error GoTo TestFail

    Dim specs As TableSpecs

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    FixtureCell(1, 5).Formula = "=1/0"

    Set specs = CreateSpecs(1)

    Assert.IsTrue (Not specs Is Nothing), _
                  "An error value in the row should not stop Create"
    Assert.AreEqual vbNullString, specs.Value("title"), _
                    "An error-valued cell should read as an empty string"
    Assert.AreEqual "date_v1", specs.Value("row"), _
                    "The other columns of the row should still read"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestErrorValuedCellReadsAsEmpty", Err.Number, Err.Description
End Sub

'@section GeoCount tests
'===============================================================================

'@sub-title Verify GeoCount clamps whatever the free-text cell holds.
'@details
'"N geo max" has no dropdown, so the cell can hold anything. An empty cell
'and text both answer the default of 5, a zero answers 1, and a huge number
'is capped at 20. Without the clamp a zero built a table with no columns and
'a huge number built a table that wide.
'@TestMethod("TableSpecs")
Public Sub TestGeoCountClampsTheFreeTextValue()
    CustomTestSetTitles Assert, "TableSpecs", "TestGeoCountClampsTheFreeTextValue"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "", "no")
    Assert.AreEqual CLng(5), CreateSpecs(1).GeoCount, _
                    "An empty N geo max should answer the default of 5"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "abc", "no")
    Assert.AreEqual CLng(5), CreateSpecs(1).GeoCount, _
                    "Text in N geo max should answer the default of 5"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "0", "no")
    Assert.AreEqual CLng(1), CreateSpecs(1).GeoCount, _
                    "A zero N geo max should answer the floor of 1"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "5", "no")
    Assert.AreEqual CLng(5), CreateSpecs(1).GeoCount, _
                    "A plain 5 should answer 5"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "100000", "no")
    Assert.AreEqual CLng(20), CreateSpecs(1).GeoCount, _
                    "A huge N geo max should answer the ceiling of 20"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoCountClampsTheFreeTextValue", Err.Number, Err.Description
End Sub

'@section SpatialTableScopes tests
'===============================================================================

'@sub-title Verify the spatial sub-type is read from the dictionary prefix.
'@details
'A spatial row names a variable, and the dictionary holds the prefixed copy
'that Prepare wrote. hf_ wins over adm1_, and a variable with neither prefix
'answers an empty string.
'@TestMethod("TableSpecs")
Public Sub TestSpatialTableScopesReadsThePrefix()
    CustomTestSetTitles Assert, "TableSpecs", "TestSpatialTableScopesReadsThePrefix"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow("h2", "", "5", "no", "no", "no")
    Assert.AreEqual "hf", CreateSpecs(1).SpatialTableScopes, _
                    "A variable with an hf_ copy in the dictionary reads as hf"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow(GEO_ROW_VARIABLE, "", "5", "no", "no", "no")
    Assert.AreEqual "geo", CreateSpecs(1).SpatialTableScopes, _
                    "A variable with an adm1_ copy in the dictionary reads as geo"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow("choi_v1", "", "5", "no", "no", "no")
    Assert.AreEqual vbNullString, CreateSpecs(1).SpatialTableScopes, _
                    "A variable with neither prefix reads as an empty string"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialTableScopesReadsThePrefix", Err.Number, Err.Description
End Sub

'@sub-title Verify a spatio-temporal table reads its spatial variable from the column.
'@TestMethod("TableSpecs")
Public Sub TestSpatialTableScopesUsesColumnForSpatioTemporal()
    CustomTestSetTitles Assert, "TableSpecs", "TestSpatialTableScopesUsesColumnForSpatioTemporal"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "h2", "5", "no")

    Assert.AreEqual "hf", CreateSpecs(1).SpatialTableScopes, _
                    "Spatio-temporal reads the spatial variable from the column field"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialTableScopesUsesColumnForSpatioTemporal", Err.Number, Err.Description
End Sub

'@section ValidTable tests
'===============================================================================

'@sub-title Verify global summary validity needs a label and a function.
'@TestMethod("TableSpecs")
Public Sub TestValidTableGlobalSummary()
    CustomTestSetTitles Assert, "TableSpecs", "TestValidTableGlobalSummary"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 Array(Array("Count of cases", "sum", ""))
    Assert.IsTrue CreateSpecs(1).ValidTable, _
                  "A global summary row with a label and a function is valid"

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 Array(Array("", "sum", ""))
    Assert.IsFalse CreateSpecs(1).ValidTable, _
                   "A global summary row with no label is invalid"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValidTableGlobalSummary", Err.Number, Err.Description
End Sub

'@sub-title Verify univariate validity needs a choice row variable.
'@TestMethod("TableSpecs")
Public Sub TestValidTableUnivariate()
    CustomTestSetTitles Assert, "TableSpecs", "TestValidTableUnivariate"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("choi_v1", "no", "no", "no")
    Assert.IsTrue CreateSpecs(1).ValidTable, _
                  "A univariate row grouped by a choice variable is valid"

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("date_v1", "no", "no", "no")
    Assert.IsFalse CreateSpecs(1).ValidTable, _
                   "A univariate row grouped by a date variable is invalid"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValidTableUnivariate", Err.Number, Err.Description
End Sub

'@sub-title Verify bivariate validity needs two choice variables.
'@TestMethod("TableSpecs")
Public Sub TestValidTableBivariate()
    CustomTestSetTitles Assert, "TableSpecs", "TestValidTableBivariate"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRow("choi_v1", "choi_h2", "no", "no", "no")
    Assert.IsTrue CreateSpecs(1).ValidTable, _
                  "A bivariate row crossing two choice variables is valid"

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRow("choi_v1", "date_v1", "no", "no", "no")
    Assert.IsFalse CreateSpecs(1).ValidTable, _
                   "A bivariate row crossing a date variable is invalid"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValidTableBivariate", Err.Number, Err.Description
End Sub

'@sub-title Verify time series validity needs a date row variable.
'@TestMethod("TableSpecs")
Public Sub TestValidTableTimeSeries()
    CustomTestSetTitles Assert, "TableSpecs", "TestValidTableTimeSeries"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Assert.IsTrue CreateSpecs(1).ValidTable, _
                  "A time series row over a date variable is valid"
    Assert.IsTrue CreateSpecs(3).ValidTable, _
                  "A time series row with no column variable is valid"

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array(Array("Series 1", "S1", "choi_v1", "choi_v1", "A title", _
                             "no", "", "", "", "no", "no", "1"))
    Assert.IsFalse CreateSpecs(1).ValidTable, _
                   "A time series row over a choice variable is invalid"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValidTableTimeSeries", Err.Number, Err.Description
End Sub

'@sub-title Verify spatial validity needs a geo or hf variable.
'@TestMethod("TableSpecs")
Public Sub TestValidTableSpatial()
    CustomTestSetTitles Assert, "TableSpecs", "TestValidTableSpatial"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow(GEO_ROW_VARIABLE, "", "5", "no", "no", "no")
    Assert.IsTrue CreateSpecs(1).ValidTable, _
                  "A spatial row over a geo variable is valid"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow("h2", "choi_v1", "5", "no", "no", "no")
    Assert.IsTrue CreateSpecs(1).ValidTable, _
                  "A spatial row over an hf variable crossed by a choice is valid"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow("choi_v1", "", "5", "no", "no", "no")
    Assert.IsFalse CreateSpecs(1).ValidTable, _
                   "A spatial row over a variable with no geo prefix is invalid"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValidTableSpatial", Err.Number, Err.Description
End Sub

'@sub-title Verify spatio-temporal validity needs a date row and a geo column.
'@TestMethod("TableSpecs")
Public Sub TestValidTableSpatioTemporal()
    CustomTestSetTitles Assert, "TableSpecs", "TestValidTableSpatioTemporal"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "5", "no")
    Assert.IsTrue CreateSpecs(1).ValidTable, _
                  "A spatio-temporal row over a date and an hf control is valid"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", GEO_ROW_VARIABLE, "5", "no")
    Assert.IsTrue CreateSpecs(1).ValidTable, _
                  "A spatio-temporal row over a geo variable is valid"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("choi_v1", "hf_h2", "5", "no")
    Assert.IsFalse CreateSpecs(1).ValidTable, _
                   "A spatio-temporal row whose row variable is not a date is invalid"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValidTableSpatioTemporal", Err.Number, Err.Description
End Sub

'@sub-title Verify a graph row is valid once a graph and a series are picked.
'@details
'The graph block's remaining columns are formulas derived from those two
'dropdowns, so the two titles are what the user actually fills in. Until the
'scope resolved, this block was skipped whole and no graph row was ever
'checked.
'@TestMethod("TableSpecs")
Public Sub TestValidTableTimeSeriesGraph()
    CustomTestSetTitles Assert, "TableSpecs", "TestValidTableTimeSeriesGraph"
    On Error GoTo TestFail

    BuildFixture TABLE_GRAPH_TIMESERIES, GraphTimeSeriesHeader(), _
                 GraphRow("A graph", "A series", "values")
    Assert.IsTrue CreateSpecs(1).ValidTable, _
                  "A graph row naming a graph and a series is valid"

    BuildFixture TABLE_GRAPH_TIMESERIES, GraphTimeSeriesHeader(), _
                 GraphRow("A graph", "", "values")
    Assert.IsFalse CreateSpecs(1).ValidTable, _
                   "A graph row with no series title is invalid"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValidTableTimeSeriesGraph", Err.Number, Err.Description
End Sub

'@section IsNewSection tests
'===============================================================================

'@sub-title Verify the first data row is always flagged as a new section.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionFirstRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionFirstRow"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Assert.IsTrue CreateSpecs(1).IsNewSection, _
                  "The first data row sits under the header, so it starts a section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionFirstRow", Err.Number, Err.Description
End Sub

'@sub-title Verify a row repeating its predecessor's section is not new.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionSameSection()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionSameSection"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Assert.IsFalse CreateSpecs(2).IsNewSection, _
                   "Row 2 repeats section S1, so it continues the section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionSameSection", Err.Number, Err.Description
End Sub

'@sub-title Verify a row changing section is flagged as new.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionDifferentSection()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionDifferentSection"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Assert.IsTrue CreateSpecs(3).IsNewSection, _
                  "Row 3 moves from S1 to S2, so it starts a section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionDifferentSection", Err.Number, Err.Description
End Sub

'@sub-title Verify a capitalisation slip keeps two rows in one section.
'@details
'The section column is free text. Comparing it with regard to case split one
'section into two and every section-keyed named range followed the split.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionIgnoresCase()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionIgnoresCase"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array( _
                     Array("Series 1", "S1", "date_v1", "choi_v1", "T1", "no", _
                           "", "", "", "no", "no", "1"), _
                     Array("Series 2", "s1", "date_v1", "choi_v1", "T2", "no", _
                           "", "", "", "no", "no", "2"))

    Assert.IsFalse CreateSpecs(2).IsNewSection, _
                   "S1 and s1 name one section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionIgnoresCase", Err.Number, Err.Description
End Sub

'@sub-title Verify global summary rows never start a section.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionGlobalSummaryAlwaysFalse()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionGlobalSummaryAlwaysFalse"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 Array(Array("Count", "sum", ""))

    Assert.IsFalse CreateSpecs(1).IsNewSection, _
                   "Each global summary row stands alone, so it starts no section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionGlobalSummaryAlwaysFalse", Err.Number, Err.Description
End Sub

'@section HasTotal tests
'===============================================================================

'@sub-title Verify time series HasTotal follows total="yes" with a column.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalTimeSeriesWithTotalYes()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalTimeSeriesWithTotalYes"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Assert.IsTrue CreateSpecs(1).HasTotal, _
                  "total='yes' with a column variable gives a total"
    Assert.IsFalse CreateSpecs(2).HasTotal, _
                   "total='no' with no percentage gives no total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalTimeSeriesWithTotalYes", Err.Number, Err.Description
End Sub

'@sub-title Verify time series HasTotal is driven by percentage=row or column.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalTimeSeriesPercentageDriven()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalTimeSeriesPercentageDriven"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array(Array("Series 1", "S1", "date_v1", "choi_v1", "T1", "no", _
                             "", "", "", "row", "no", "1"))
    Assert.IsTrue CreateSpecs(1).HasTotal, _
                  "percentage='row' needs the total as its denominator"

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array(Array("Series 1", "S1", "date_v1", "choi_v1", "T1", "no", _
                             "", "", "", "column", "no", "1"))
    Assert.IsTrue CreateSpecs(1).HasTotal, _
                  "percentage='column' needs the total as its denominator"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalTimeSeriesPercentageDriven", Err.Number, Err.Description
End Sub

'@sub-title Verify time series HasTotal needs a column variable.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalTimeSeriesNoColumnNoTotal()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalTimeSeriesNoColumnNoTotal"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Assert.IsFalse CreateSpecs(3).HasTotal, _
                   "A time series row with no column variable has nothing to total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalTimeSeriesNoColumnNoTotal", Err.Number, Err.Description
End Sub

'@sub-title Verify HasTotal per scope for the scopes that answer a constant.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalConstantScopes()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalConstantScopes"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 Array(Array("Count", "sum", ""))
    Assert.IsFalse CreateSpecs(1).HasTotal, _
                   "Global summary never has a total"

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("choi_v1", "no", "no", "no")
    Assert.IsTrue CreateSpecs(1).HasTotal, _
                  "Univariate always has a total"

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRow("choi_v1", "choi_h2", "no", "no", "no")
    Assert.IsTrue CreateSpecs(1).HasTotal, _
                  "Bivariate always has a total"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "5", "no")
    Assert.IsFalse CreateSpecs(1).HasTotal, _
                   "Spatio-temporal never has a total"

    BuildFixture TABLE_GRAPH_TIMESERIES, GraphTimeSeriesHeader(), _
                 GraphRow("A graph", "A series", "values")
    Assert.IsFalse CreateSpecs(1).HasTotal, _
                   "The graph block never has a total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalConstantScopes", Err.Number, Err.Description
End Sub

'@sub-title Verify spatial HasTotal follows the presence of a column variable.
'@TestMethod("TableSpecs")
Public Sub TestHasTotalSpatialFollowsTheColumn()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasTotalSpatialFollowsTheColumn"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow(GEO_ROW_VARIABLE, "choi_v1", "5", "no", "no", "no")
    Assert.IsTrue CreateSpecs(1).HasTotal, _
                  "A spatial row with a column variable has a total"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow(GEO_ROW_VARIABLE, "", "5", "no", "no", "no")
    Assert.IsFalse CreateSpecs(1).HasTotal, _
                   "A spatial row with no column variable has no total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasTotalSpatialFollowsTheColumn", Err.Number, Err.Description
End Sub

'@section TotalRequested tests
'===============================================================================

'@sub-title Verify TotalRequested tracks what the user wrote.
'@details
'A time series row can carry a total the user never asked for, because
'percentage needs one as its denominator. TotalRequested is what tells the
'rendering layer whether to show it.
'@TestMethod("TableSpecs")
Public Sub TestTotalRequestedTracksTheUserChoice()
    CustomTestSetTitles Assert, "TableSpecs", "TestTotalRequestedTracksTheUserChoice"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Assert.IsTrue CreateSpecs(1).TotalRequested, _
                  "total='yes' is a total the user asked for"
    Assert.IsFalse CreateSpecs(2).TotalRequested, _
                   "total='no' is no request"

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array(Array("Series 1", "S1", "date_v1", "choi_v1", "T1", "no", _
                             "", "", "", "row", "no", "1"))
    Dim specs As TableSpecs
    Set specs = CreateSpecs(1)
    Assert.IsTrue specs.HasTotal, _
                  "percentage='row' builds a total"
    Assert.IsFalse specs.TotalRequested, _
                   "A total built for percentage is no request"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTotalRequestedTracksTheUserChoice", Err.Number, Err.Description
End Sub

'@sub-title Verify a pasted flag is read whatever its case and spacing.
'@details
'The flag columns carry a dropdown, and a pasted value goes past the
'validation. Every flag comparison trims and lower-cases first.
'@TestMethod("TableSpecs")
Public Sub TestFlagValueIgnoresCaseAndSpacing()
    CustomTestSetTitles Assert, "TableSpecs", "TestFlagValueIgnoresCaseAndSpacing"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array(Array("Series 1", "S1", "date_v1", "choi_v1", "T1", " YES ", _
                             "", "", "", "no", " Yes", "1"))
    Dim specs As TableSpecs
    Set specs = CreateSpecs(1)

    Assert.IsTrue specs.TotalRequested, _
                  "' Yes' in the total column is a total request"
    Assert.IsTrue specs.HasMissing, _
                  "' YES ' in the missing column shows missing data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFlagValueIgnoresCaseAndSpacing", Err.Number, Err.Description
End Sub

'@section HasPercentage tests
'===============================================================================

'@sub-title Verify time series HasPercentage needs a total behind it.
'@TestMethod("TableSpecs")
Public Sub TestHasPercentageTimeSeries()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasPercentageTimeSeries"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Assert.IsTrue CreateSpecs(1).HasPercentage, _
                  "percentage='row' with a total shows percentages"

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array(Array("Series 1", "S1", "date_v1", "", "T1", "no", _
                             "", "", "", "row", "no", "1"))
    Assert.IsFalse CreateSpecs(1).HasPercentage, _
                   "With no column variable there is no total, so no percentages"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasPercentageTimeSeries", Err.Number, Err.Description
End Sub

'@sub-title Verify univariate and spatial HasPercentage follow a yes.
'@TestMethod("TableSpecs")
Public Sub TestHasPercentageUnivariateAndSpatial()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasPercentageUnivariateAndSpatial"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("choi_v1", "no", "yes", "no")
    Assert.IsTrue CreateSpecs(1).HasPercentage, _
                  "A univariate row with percentage='yes' shows percentages"

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("choi_v1", "no", "no", "no")
    Assert.IsFalse CreateSpecs(1).HasPercentage, _
                   "A univariate row with percentage='no' shows none"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow(GEO_ROW_VARIABLE, "choi_v1", "5", "no", "yes", "no")
    Assert.IsTrue CreateSpecs(1).HasPercentage, _
                  "A spatial row with a column and percentage='yes' shows percentages"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow(GEO_ROW_VARIABLE, "", "5", "no", "yes", "no")
    Assert.IsFalse CreateSpecs(1).HasPercentage, _
                   "A spatial row with no column has no total, so no percentages"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasPercentageUnivariateAndSpatial", Err.Number, Err.Description
End Sub

'@sub-title Verify the scopes that never show percentages.
'@TestMethod("TableSpecs")
Public Sub TestHasPercentageConstantScopes()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasPercentageConstantScopes"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 Array(Array("Count", "sum", ""))
    Assert.IsFalse CreateSpecs(1).HasPercentage, _
                   "Global summary never shows percentages"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "5", "no")
    Assert.IsFalse CreateSpecs(1).HasPercentage, _
                   "Spatio-temporal never shows percentages"

    BuildFixture TABLE_GRAPH_TIMESERIES, GraphTimeSeriesHeader(), _
                 GraphRow("A graph", "A series", "percentages")
    Assert.IsFalse CreateSpecs(1).HasPercentage, _
                   "The graph block never shows percentages on a table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasPercentageConstantScopes", Err.Number, Err.Description
End Sub

'@section HasMissing tests
'===============================================================================

'@sub-title Verify time series HasMissing needs a column variable.
'@TestMethod("TableSpecs")
Public Sub TestHasMissingTimeSeries()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasMissingTimeSeries"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Assert.IsTrue CreateSpecs(1).HasMissing, _
                  "missing='yes' with a column variable shows missing data"

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array(Array("Series 1", "S1", "date_v1", "", "T1", "yes", _
                             "", "", "", "no", "no", "1"))
    Assert.IsFalse CreateSpecs(1).HasMissing, _
                   "With no column variable there is no axis to show missing on"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasMissingTimeSeries", Err.Number, Err.Description
End Sub

'@sub-title Verify univariate and spatial HasMissing follow a yes.
'@TestMethod("TableSpecs")
Public Sub TestHasMissingUnivariateAndSpatial()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasMissingUnivariateAndSpatial"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("choi_v1", "yes", "no", "no")
    Assert.IsTrue CreateSpecs(1).HasMissing, _
                  "A univariate row with missing='yes' shows missing data"

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("choi_v1", "no", "no", "no")
    Assert.IsFalse CreateSpecs(1).HasMissing, _
                   "A univariate row with missing='no' shows none"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow(GEO_ROW_VARIABLE, "choi_v1", "5", "yes", "no", "no")
    Assert.IsTrue CreateSpecs(1).HasMissing, _
                  "A spatial row with a column and missing='yes' shows missing data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasMissingUnivariateAndSpatial", Err.Number, Err.Description
End Sub

'@sub-title Verify the scopes that never show missing data.
'@TestMethod("TableSpecs")
Public Sub TestHasMissingConstantScopes()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasMissingConstantScopes"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 Array(Array("Count", "sum", ""))
    Assert.IsFalse CreateSpecs(1).HasMissing, _
                   "Global summary never shows missing data"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "5", "no")
    Assert.IsFalse CreateSpecs(1).HasMissing, _
                   "Spatio-temporal never shows missing data"

    BuildFixture TABLE_GRAPH_TIMESERIES, GraphTimeSeriesHeader(), _
                 GraphRow("A graph", "A series", "values")
    Assert.IsFalse CreateSpecs(1).HasMissing, _
                   "The graph block never shows missing data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasMissingConstantScopes", Err.Number, Err.Description
End Sub

'@section HasGraph tests
'===============================================================================

'@sub-title Verify a bivariate graph is drawn for percentages and for values.
'@details
'The dropdown bound to "Add graph" offers "percentages" and "values". The
'class used to compare against the singular "percentage", so a user who
'picked percentages got no chart and no message.
'@TestMethod("TableSpecs")
Public Sub TestHasGraphBivariateAcceptsTheDropdownValues()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasGraphBivariateAcceptsTheDropdownValues"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRow("choi_v1", "choi_h2", "no", "no", "percentages")
    Assert.IsTrue CreateSpecs(1).HasGraph, _
                  "A bivariate row with graph='percentages' draws a chart"

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRow("choi_v1", "choi_h2", "no", "no", "values")
    Assert.IsTrue CreateSpecs(1).HasGraph, _
                  "A bivariate row with graph='values' draws a chart"

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRow("choi_v1", "choi_h2", "no", "no", "no")
    Assert.IsFalse CreateSpecs(1).HasGraph, _
                   "A bivariate row with graph='no' draws none"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasGraphBivariateAcceptsTheDropdownValues", Err.Number, Err.Description
End Sub

'@sub-title Verify univariate, spatial and spatio-temporal graphs follow a yes.
'@TestMethod("TableSpecs")
Public Sub TestHasGraphYesNoScopes()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasGraphYesNoScopes"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("choi_v1", "no", "no", "yes")
    Assert.IsTrue CreateSpecs(1).HasGraph, _
                  "A univariate row with graph='yes' draws a chart"

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRow("choi_v1", "no", "no", "no")
    Assert.IsFalse CreateSpecs(1).HasGraph, _
                   "A univariate row with graph='no' draws none"

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRow(GEO_ROW_VARIABLE, "", "5", "no", "no", "yes")
    Assert.IsTrue CreateSpecs(1).HasGraph, _
                  "A spatial row with graph='yes' draws a chart"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "5", "yes")
    Assert.IsTrue CreateSpecs(1).HasGraph, _
                  "A spatio-temporal row with graph='yes' draws a chart"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasGraphYesNoScopes", Err.Number, Err.Description
End Sub

'@sub-title Verify global summary and the graph block draw no table chart.
'@details
'The time series header carries no column holding the word "graph", so the
'time series branch of HasGraph answers False for every real row. Time series
'charts come from Tab_Graph_TimeSeries, and every row of that block is
'already a graph.
'@TestMethod("TableSpecs")
Public Sub TestHasGraphIsFalseForTheBlocksThatCarryNoFlag()
    CustomTestSetTitles Assert, "TableSpecs", "TestHasGraphIsFalseForTheBlocksThatCarryNoFlag"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 Array(Array("Count", "sum", ""))
    Assert.IsFalse CreateSpecs(1).HasGraph, _
                   "Global summary never draws a chart"

    BuildFixture TABLE_GRAPH_TIMESERIES, GraphTimeSeriesHeader(), _
                 GraphRow("A graph", "A series", "values")
    Assert.IsFalse CreateSpecs(1).HasGraph, _
                   "The graph block draws its charts through its own writer"

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()
    Assert.IsFalse CreateSpecs(1).HasGraph, _
                   "The time series header carries no graph column, so it draws none"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasGraphIsFalseForTheBlocksThatCarryNoFlag", Err.Number, Err.Description
End Sub

'@section Flag vocabulary contract
'===============================================================================

'@sub-title Drive every registered dropdown entry through the flag it feeds.
'@details
'Each analysis column carries a dropdown registered by SetupPreparation, and
'the flag properties compare against strings. This test walks every entry of
'every vocabulary through the property bound to it, so a value the user can
'pick that the class answers nothing for shows up here. That is the shape of
'the bug that left bivariate percentage charts undrawn.
'@TestMethod("TableSpecs")
Public Sub TestFlagVocabularyContract()
    CustomTestSetTitles Assert, "TableSpecs", "TestFlagVocabularyContract"
    On Error GoTo TestFail

    Dim entry As Variant

    ' __yesno on the univariate "Add missing data", "Add percentage" and
    ' "Add graph" columns.
    For Each entry In Array("yes", "no")
        BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                     UnivariateRow("choi_v1", CStr(entry), CStr(entry), CStr(entry))
        Dim univariate As TableSpecs
        Set univariate = CreateSpecs(1)
        Assert.AreEqual (CStr(entry) = "yes"), univariate.HasMissing, _
                        "__yesno entry '" & entry & "' should drive univariate HasMissing"
        Assert.AreEqual (CStr(entry) = "yes"), univariate.HasPercentage, _
                        "__yesno entry '" & entry & "' should drive univariate HasPercentage"
        Assert.AreEqual (CStr(entry) = "yes"), univariate.HasGraph, _
                        "__yesno entry '" & entry & "' should drive univariate HasGraph"
    Next entry

    ' __missing_ba on the bivariate "Add missing data" column.
    For Each entry In Array("no", "row", "column", "all")
        BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                     BivariateRow("choi_v1", "choi_h2", CStr(entry), "no", "no")
        Assert.AreEqual (CStr(entry) <> "no"), CreateSpecs(1).HasMissing, _
                        "__missing_ba entry '" & entry & "' should drive bivariate HasMissing"
    Next entry

    ' __percentage_ba on the bivariate "Add percentage" column.
    For Each entry In Array("no", "row", "column", "total")
        BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                     BivariateRow("choi_v1", "choi_h2", "no", CStr(entry), "no")
        Assert.AreEqual (CStr(entry) <> "no"), CreateSpecs(1).HasPercentage, _
                        "__percentage_ba entry '" & entry & "' should drive bivariate HasPercentage"
    Next entry

    ' __perc_val on the bivariate "Add graph" column.
    For Each entry In Array("percentages", "values")
        BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                     BivariateRow("choi_v1", "choi_h2", "no", "no", CStr(entry))
        Assert.IsTrue CreateSpecs(1).HasGraph, _
                      "__perc_val entry '" & entry & "' should draw a bivariate chart"
    Next entry

    ' __percentage_ta on the time series "Add percentage" column, with a
    ' column variable present so the total exists.
    For Each entry In Array("no", "row", "column")
        BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                     Array(Array("Series 1", "S1", "date_v1", "choi_v1", "T1", "no", _
                                 "", "", "", CStr(entry), "no", "1"))
        Assert.AreEqual (CStr(entry) <> "no"), CreateSpecs(1).HasPercentage, _
                        "__percentage_ta entry '" & entry & "' should drive time series HasPercentage"
    Next entry

    ' __yesno on the spatial "Add missing data", "Add percentage" and
    ' "Add graph" columns, with a column variable present.
    For Each entry In Array("yes", "no")
        BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                     SpatialRow(GEO_ROW_VARIABLE, "choi_v1", "5", CStr(entry), _
                                CStr(entry), CStr(entry))
        Dim spatial As TableSpecs
        Set spatial = CreateSpecs(1)
        Assert.AreEqual (CStr(entry) = "yes"), spatial.HasMissing, _
                        "__yesno entry '" & entry & "' should drive spatial HasMissing"
        Assert.AreEqual (CStr(entry) = "yes"), spatial.HasPercentage, _
                        "__yesno entry '" & entry & "' should drive spatial HasPercentage"
        Assert.AreEqual (CStr(entry) = "yes"), spatial.HasGraph, _
                        "__yesno entry '" & entry & "' should drive spatial HasGraph"
    Next entry

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFlagVocabularyContract", Err.Number, Err.Description
End Sub

'@section Navigation tests
'===============================================================================

'@sub-title Verify Previous returns the spec of the preceding data row.
'@TestMethod("TableSpecs")
Public Sub TestPreviousReturnsPriorRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestPreviousReturnsPriorRow"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Dim prevSpec As TableSpecs
    Set prevSpec = CreateSpecs(2).Previous

    Assert.IsTrue (Not prevSpec Is Nothing), _
                  "Row 2 continues section S1, so it has a previous table"
    Assert.AreEqual "TS_tab1", prevSpec.TableId, _
                    "The previous table of row 2 is row 1"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreviousReturnsPriorRow", Err.Number, Err.Description
End Sub

'@sub-title Verify Previous answers Nothing on a row that starts a section.
'@details
'Previous used to raise InvalidArgument here, which forced every caller to
'wrap the read in On Error Resume Next and made "this row starts a section"
'look the same as "something else broke".
'@TestMethod("TableSpecs")
Public Sub TestPreviousNothingOnNewSection()
    CustomTestSetTitles Assert, "TableSpecs", "TestPreviousNothingOnNewSection"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Dim prevSpec As TableSpecs
    Set prevSpec = CreateSpecs(1).Previous

    Assert.IsTrue (prevSpec Is Nothing), _
                  "The first row starts a section, so it has no previous table"

    Set prevSpec = CreateSpecs(3).Previous
    Assert.IsTrue (prevSpec Is Nothing), _
                  "Row 3 opens section S2, so it has no previous table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreviousNothingOnNewSection", Err.Number, Err.Description
End Sub

'@sub-title Verify a boundary carried by an invalid row still opens the section.
'@details
'Rows are S1 valid, S2 invalid, S2 valid. The build skips the invalid row, so
'row 3 is the first row of S2 that is drawn and it has to open the section.
'Comparing against the row physically above read "S2 follows S2", row 3
'reported no new section, and section S2 never opened at all.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionSeesABoundaryCarriedByAnInvalidRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionSeesABoundaryCarriedByAnInvalidRow"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array( _
                     Array("Series 1", "S1", "date_v1", "choi_v1", "T1", "no", _
                           "", "", "", "no", "no", "1"), _
                     Array("Series 2", "S2", "choi_v1", "choi_v1", "T2", "no", _
                           "", "", "", "no", "no", "2"), _
                     Array("Series 3", "S2", "date_v1", "choi_v1", "T3", "no", _
                           "", "", "", "no", "no", "3"))

    Dim specs As TableSpecs
    Set specs = CreateSpecs(3)

    Assert.IsTrue specs.IsNewSection, _
                  "Row 3 is the first drawn row of S2, so it opens the section"
    Assert.IsTrue (specs.Previous Is Nothing), _
                  "A row that opens a section has no previous table"
    Assert.AreEqual "TS_tab3", specs.TableSectionId, _
                    "Row 3 owns the section id of S2"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionSeesABoundaryCarriedByAnInvalidRow", Err.Number, Err.Description
End Sub

'@sub-title Verify the row below an invalid section anchor becomes the anchor.
'@details
'Rows are S1 invalid, S1 valid, S1 valid. The build skips row 1, so row 2 is
'the first row of S1 that is drawn and it has to carry the section header and
'the date controls. Row 2 used to report no new section, then every row of
'the section appended itself to infrastructure nothing had built. This is
'issue #183.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionWhenTheSectionAnchorIsInvalid()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionWhenTheSectionAnchorIsInvalid"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array( _
                     Array("Series 1", "S1", "choi_v1", "choi_v1", "T1", "no", _
                           "", "", "", "no", "no", "1"), _
                     Array("Series 2", "S1", "date_v1", "choi_v1", "T2", "no", _
                           "", "", "", "no", "no", "2"), _
                     Array("Series 3", "S1", "date_v1", "choi_v1", "T3", "no", _
                           "", "", "", "no", "no", "3"))

    Assert.IsFalse CreateSpecs(1).ValidTable, _
                   "Row 1 groups by a choice variable, so the build skips it"
    Assert.IsTrue CreateSpecs(2).IsNewSection, _
                  "Row 2 is the first drawn row of S1, so it opens the section"
    Assert.AreEqual "TS_tab2", CreateSpecs(2).TableSectionId, _
                    "Row 2 owns the section id"
    Assert.IsFalse CreateSpecs(3).IsNewSection, _
                   "Row 3 follows a drawn row of the same section"
    Assert.AreEqual "TS_tab2", CreateSpecs(3).TableSectionId, _
                    "Row 3 joins the section row 2 opened"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionWhenTheSectionAnchorIsInvalid", Err.Number, Err.Description
End Sub

'@sub-title Verify a run of invalid rows leaves the first drawn row as anchor.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionAfterARunOfInvalidRows()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionAfterARunOfInvalidRows"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array( _
                     Array("Series 1", "S1", "choi_v1", "choi_v1", "T1", "no", _
                           "", "", "", "no", "no", "1"), _
                     Array("Series 2", "S1", "choi_v1", "choi_v1", "T2", "no", _
                           "", "", "", "no", "no", "2"), _
                     Array("Series 3", "S1", "date_v1", "choi_v1", "T3", "no", _
                           "", "", "", "no", "no", "3"))

    Assert.IsTrue CreateSpecs(3).IsNewSection, _
                  "With two invalid rows above it, row 3 opens the section"
    Assert.IsTrue (CreateSpecs(3).Previous Is Nothing), _
                  "There is no drawn table above row 3"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionAfterARunOfInvalidRows", Err.Number, Err.Description
End Sub

'@sub-title Verify the same rule reaches the non-temporal blocks.
'@details
'ComputeIsNewSection serves every scope, so an invalid section anchor changes
'the grouping of univariate and bivariate tables too.
'@TestMethod("TableSpecs")
Public Sub TestIsNewSectionAnchorRuleReachesEveryScope()
    CustomTestSetTitles Assert, "TableSpecs", "TestIsNewSectionAnchorRuleReachesEveryScope"
    On Error GoTo TestFail

    ' Univariate: row 1 groups by a date variable, so the build skips it.
    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 Array( _
                     Array("S1", "First", "date_v1", "no", "", "", "", "no", "no", "no"), _
                     Array("S1", "Second", "choi_v1", "no", "", "", "", "no", "no", "no"))
    Assert.IsFalse CreateSpecs(1).ValidTable, _
                   "A univariate row over a date variable is skipped"
    Assert.IsTrue CreateSpecs(2).IsNewSection, _
                  "The first drawn univariate row opens the section"

    ' Bivariate: row 1 crosses a date variable, so the build skips it.
    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 Array( _
                     Array("S1", "First", "choi_v1", "date_v1", "no", "", "", "", _
                           "no", "no", "no"), _
                     Array("S1", "Second", "choi_v1", "choi_h2", "no", "", "", "", _
                           "no", "no", "no"))
    Assert.IsFalse CreateSpecs(1).ValidTable, _
                   "A bivariate row crossing a date variable is skipped"
    Assert.IsTrue CreateSpecs(2).IsNewSection, _
                  "The first drawn bivariate row opens the section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsNewSectionAnchorRuleReachesEveryScope", Err.Number, Err.Description
End Sub

'@sub-title Verify Previous walks over an invalid row inside a section.
'@TestMethod("TableSpecs")
Public Sub TestPreviousWalksOverAnInvalidRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestPreviousWalksOverAnInvalidRow"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array( _
                     Array("Series 1", "S1", "date_v1", "choi_v1", "T1", "no", _
                           "", "", "", "no", "no", "1"), _
                     Array("Series 2", "S1", "choi_v1", "choi_v1", "T2", "no", _
                           "", "", "", "no", "no", "2"), _
                     Array("Series 3", "S1", "date_v1", "choi_v1", "T3", "no", _
                           "", "", "", "no", "no", "3"))

    Dim prevSpec As TableSpecs
    Set prevSpec = CreateSpecs(3).Previous

    Assert.IsTrue (Not prevSpec Is Nothing), _
                  "Section S1 holds a valid table above row 3"
    Assert.AreEqual "TS_tab1", prevSpec.TableId, _
                    "The walk steps over the invalid row 2 and reaches row 1"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreviousWalksOverAnInvalidRow", Err.Number, Err.Description
End Sub

'@sub-title Verify NextSpecs returns the spec of the following data row.
'@TestMethod("TableSpecs")
Public Sub TestNextSpecsReturnsNextRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestNextSpecsReturnsNextRow"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Dim nextSpec As TableSpecs
    Set nextSpec = CreateSpecs(1).NextSpecs(FixtureDataRange(3))

    Assert.IsTrue (Not nextSpec Is Nothing), _
                  "Row 1 has a following table inside the anchor"
    Assert.AreEqual "TS_tab2", nextSpec.TableId, _
                    "The table following row 1 is row 2"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNextSpecsReturnsNextRow", Err.Number, Err.Description
End Sub

'@sub-title Verify NextSpecs returns Nothing beyond the anchor row.
'@TestMethod("TableSpecs")
Public Sub TestNextSpecsNothingBeyondAnchor()
    CustomTestSetTitles Assert, "TableSpecs", "TestNextSpecsNothingBeyondAnchor"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Dim nextSpec As TableSpecs
    Set nextSpec = CreateSpecs(3).NextSpecs(FixtureDataRange(3))

    Assert.IsTrue (nextSpec Is Nothing), _
                  "The anchor is the last row, so nothing follows it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNextSpecsNothingBeyondAnchor", Err.Number, Err.Description
End Sub

'@sub-title Verify NextSpecs walks over an invalid row.
'@TestMethod("TableSpecs")
Public Sub TestNextSpecsSkipsAnInvalidRow()
    CustomTestSetTitles Assert, "TableSpecs", "TestNextSpecsSkipsAnInvalidRow"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 Array( _
                     Array("Series 1", "S1", "date_v1", "choi_v1", "T1", "no", _
                           "", "", "", "no", "no", "1"), _
                     Array("Series 2", "S1", "choi_v1", "choi_v1", "T2", "no", _
                           "", "", "", "no", "no", "2"), _
                     Array("Series 3", "S1", "date_v1", "choi_v1", "T3", "no", _
                           "", "", "", "no", "no", "3"))

    Dim nextSpec As TableSpecs
    Set nextSpec = CreateSpecs(1).NextSpecs(FixtureDataRange(3))

    Assert.IsTrue (Not nextSpec Is Nothing), _
                  "A valid table follows row 1 inside the anchor"
    Assert.AreEqual "TS_tab3", nextSpec.TableId, _
                    "The walk steps over the invalid row 2 and reaches row 3"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNextSpecsSkipsAnInvalidRow", Err.Number, Err.Description
End Sub

'@sub-title Verify the first table of a section owns the section id.
'@TestMethod("TableSpecs")
Public Sub TestTableSectionIdFirstInSection()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableSectionIdFirstInSection"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Dim specs As TableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual specs.TableId, specs.TableSectionId, _
                    "The first table of a section carries its own id as the section id"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableSectionIdFirstInSection", Err.Number, Err.Description
End Sub

'@sub-title Verify a later table inherits the section id of the first one.
'@TestMethod("TableSpecs")
Public Sub TestTableSectionIdSubsequentInSection()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableSectionIdSubsequentInSection"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesDataRows()

    Assert.AreEqual "TS_tab1", CreateSpecs(2).TableSectionId, _
                    "Row 2 continues S1, so it carries row 1's id"
    Assert.AreEqual "TS_tab3", CreateSpecs(3).TableSectionId, _
                    "Row 3 opens S2, so it carries its own id"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableSectionIdSubsequentInSection", Err.Number, Err.Description
End Sub

'@sub-title Verify a global summary row is its own section.
'@TestMethod("TableSpecs")
Public Sub TestTableSectionIdGlobalSummary()
    CustomTestSetTitles Assert, "TableSpecs", "TestTableSectionIdGlobalSummary"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 Array(Array("Count", "sum", ""), Array("Mean age", "mean", ""))

    Assert.AreEqual "GS_tab2", CreateSpecs(2).TableSectionId, _
                    "Each global summary row stands alone as its own section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableSectionIdGlobalSummary", Err.Number, Err.Description
End Sub

'@section Category tests
'===============================================================================

'@sub-title Verify categories are empty without a LinelistSpecs to read.
'@details
'The setup checking context builds a TableSpecs with no linelist behind it,
'so both category getters answer an empty array there.
'@TestMethod("TableSpecs")
Public Sub TestCategoriesAreEmptyWithoutLinelistSpecs()
    CustomTestSetTitles Assert, "TableSpecs", "TestCategoriesAreEmptyWithoutLinelistSpecs"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRow("choi_v1", "choi_h2", "no", "no", "no")

    Dim specs As TableSpecs
    Set specs = CreateSpecs(1)

    Assert.AreEqual CLng(0), specs.RowCategories(Nothing).Length, _
                    "Row categories are empty with no linelist to read them from"
    Assert.AreEqual CLng(0), specs.ColumnCategories(Nothing).Length, _
                    "Column categories are empty with no linelist to read them from"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCategoriesAreEmptyWithoutLinelistSpecs", Err.Number, Err.Description
End Sub

'@sub-title Verify spatio-temporal column categories are GeoCount placeholders.
'@details
'The geographic names are filled in at runtime, so the setup only decides how
'many columns to leave room for.
'@TestMethod("TableSpecs")
Public Sub TestSpatioTemporalColumnCategoriesArePlaceholders()
    CustomTestSetTitles Assert, "TableSpecs", "TestSpatioTemporalColumnCategoriesArePlaceholders"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "3", "no")

    Dim categories As BetterArray
    Set categories = CreateSpecs(1).ColumnCategories(Nothing)

    Assert.AreEqual CLng(3), categories.Length, _
                    "N geo max of 3 leaves room for three geographic columns"
    Assert.AreEqual vbNullString, CStr(categories.Item(1)), _
                    "Each placeholder is an empty string"

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRow("date_v1", "hf_h2", "", "no")
    Assert.AreEqual CLng(5), CreateSpecs(1).ColumnCategories(Nothing).Length, _
                    "An empty N geo max leaves room for the default five"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalColumnCategoriesArePlaceholders", Err.Number, Err.Description
End Sub
