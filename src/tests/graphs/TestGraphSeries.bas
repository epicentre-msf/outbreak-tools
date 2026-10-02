Attribute VB_Name = "TestGraphSeries"
Attribute VB_Description = "Tests for GraphSeries class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for GraphSeries class")

'@description
'Drives GraphSeries, which answers the chart series of one built cross-table:
'univariate, bivariate, spatial or spatio-temporal. It came out of GraphSpecs,
'which held this half and the time series half in one type with eight branches
'discriminating between them.
'
'THE FIXTURES ARE REAL LISTOBJECTS WITH THE REAL HEADERS
'-------------------------------------------------------------------------------
'TableSpecs reads the analysis scope off the name of the ListObject a row sits
'in, so the fixture names its tables Tab_Univariate_Analysis,
'Tab_Bivariate_Analysis, Tab_SpatioTemporal_Analysis and
'Tab_TimeSeries_Analysis. The headers are the ones measured on 2026-07-31 in
'.mock/setup_mock.xlsb, src/bin/setup/setup.xlsb and
'releases/main/setup/setup_main-2026-06-11.xlsb, which all agree.
'
'NO BUILD IS NEEDED, AND THAT IS WHAT KEEPS THIS SUITE CHEAP
'-------------------------------------------------------------------------------
'The builder reads the scope, the flags and the geographic unit count off
'TableSpecs, and Table.NumberOfColumns, which is 0 until Build runs. So the
'univariate and spatio-temporal paths need no Build at all: a spatio-temporal
'table takes its column count from GeoCount. WriteSpatioTemporalGraph drives the
'pair the same way in production.
'@depends GraphSeries, SeriesBuffer, CrossTable, TableSpecs, LinelistSpecs, LLdictionary, Checking, BetterArray, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SIMPLE_SHEET As String = "GSeriesSimple"
Private Const OUTPUT_SHEET As String = "GSeriesOutput"
Private Const DICT_SHEET As String = "GSeriesDict"

' The simple-mode blocks, named the way the setup workbook names them.
Private Const TABLE_UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const TABLE_BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TABLE_SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const TABLE_TIMESERIES As String = "Tab_TimeSeries_Analysis"

' The header row of a fixture. Data row N sits at HEADER_ROW + N, so TableId
' reads "<prefix>_tabN".
Private Const SIMPLE_HEADER_ROW As Long = 5

Private Assert As CustomTest
Private Dict As LLdictionary

'@section Header rows measured in the setup workbooks
'===============================================================================

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

'@sub-title Header of Tab_SpatioTemporal_Analysis, 10 columns.
'@details This block carries no "flip coordinates" column, so FlipCoordinates
'answers False for every row of it.
Private Function SpatioTemporalHeader() As Variant
    SpatioTemporalHeader = Array( _
        "Section (select)", "Time variable (row)", "Geo/HF variable (column)", _
        "N geo max", "Title (header)", "Spatial type", "Summary function", _
        "Summary label", "Format", "Add graph")
End Function

'@sub-title Header of Tab_TimeSeries_Analysis, 12 columns.
Private Function TimeSeriesHeader() As Variant
    TimeSeriesHeader = Array( _
        "Series ID", "Section", "Time variable (row)", _
        "Group by variable (column)", "Title (header)", "Add missing data", _
        "Summary function", "Summary label", "Format", "Add percentage", _
        "Add total", "Table order")
End Function

'@section Fixture helpers
'===============================================================================

'@sub-title Free a ListObject name held anywhere in the workbook.
'@details
'A ListObject name is unique across the workbook, and four other suites build
'fixtures under these same names. Unlist turns the table back into an ordinary
'range and frees the name.
'@param tableName String. The ListObject name to release.
Private Sub ReleaseTableName(ByVal tableName As String)
    Dim sh As Worksheet
    Dim idx As Long

    For Each sh In ThisWorkbook.Worksheets
        For idx = sh.ListObjects.Count To 1 Step -1
            If StrComp(sh.ListObjects(idx).Name, tableName, vbTextCompare) = 0 Then
                sh.ListObjects(idx).Unlist
            End If
        Next idx
    Next sh
End Sub

'@sub-title Drop every ListObject on a fixture sheet.
'@param sh Worksheet. The fixture worksheet.
Private Sub RemoveFixtureTables(ByVal sh As Worksheet)
    Dim idx As Long

    For idx = sh.ListObjects.Count To 1 Step -1
        sh.ListObjects(idx).Unlist
    Next idx
End Sub

'@sub-title Delete every workbook name that points at one worksheet.
'@details
'The names these fixtures write are workbook-scoped and outlive a Cells.Clear.
'The owning worksheet is read off RefersToRange, which answers whatever the
'sheet is called. A name holding a value has no range and raises, so the read is
'guarded.
'@param sh Worksheet. The worksheet whose names go.
Private Sub ClearSheetNames(ByVal sh As Worksheet)
    Dim idx As Long
    Dim nameItem As Name
    Dim owner As Worksheet

    For idx = ThisWorkbook.Names.Count To 1 Step -1
        Set nameItem = ThisWorkbook.Names(idx)
        Set owner = Nothing

        On Error Resume Next
        Set owner = nameItem.RefersToRange.Worksheet
        Err.Clear
        On Error GoTo 0

        If Not owner Is Nothing Then
            If StrComp(owner.Name, sh.Name, vbTextCompare) = 0 Then nameItem.Delete
        End If
    Next idx
End Sub

'@sub-title Empty a fixture worksheet without clearing the whole sheet.
'@details
'EnsureWorksheet is asked NOT to clear. Its ClearWorksheet runs Cells.Clear over
'the whole 1,048,576 by 16,384 sheet, which costs seconds a call and took a
'green analyses run past the runner's cap.
'@param sheetName String. The worksheet to reset.
'@return Worksheet. The empty worksheet.
Private Function ResetFixtureSheet(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(sheetName, clearSheet:=False, visibility:=xlSheetHidden)
    RemoveFixtureTables sh
    ClearSheetNames sh
    sh.Range("A1:AZ60").Clear

    Set ResetFixtureSheet = sh
End Function

'@sub-title A LinelistSpecs carrying nothing but a dictionary.
'@details
'CrossTable reads the translation object and the categories off it, and
'TableSpecs asks for the dictionary. Dictionary() hands back the injected
'instance without asking whether Prepare has run.
'@return LinelistSpecs. The collaborator CrossTable requires.
Private Function LinelistData() As LinelistSpecs
    Dim stub As LinelistSpecs

    Set stub = New LinelistSpecs
    stub.TestAssignDictionary Dict
    Set LinelistData = stub
End Function

'@sub-title Build a fixture table and answer the specification of its one row.
'@param tableName String. One of the analysis ListObject names.
'@param headerRow Variant. The header of that block.
'@param dataRow Variant. One data row, as wide as the header.
'@return TableSpecs. The specification of that row.
Private Function SimpleSpecs(ByVal tableName As String, _
                             ByVal headerRow As Variant, _
                             ByVal dataRow As Variant) As TableSpecs
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim colCount As Long

    ReleaseTableName tableName

    Set sh = ResetFixtureSheet(SIMPLE_SHEET)

    colCount = UBound(headerRow) - LBound(headerRow) + 1

    WriteMatrix sh.Cells(SIMPLE_HEADER_ROW, 1), RowsToMatrix(Array(headerRow))
    WriteMatrix sh.Cells(SIMPLE_HEADER_ROW + 1, 1), RowsToMatrix(Array(dataRow))

    Set lo = sh.ListObjects.Add(xlSrcRange, _
                                sh.Range(sh.Cells(SIMPLE_HEADER_ROW, 1), _
                                         sh.Cells(SIMPLE_HEADER_ROW + 1, colCount)), _
                                , xlYes)
    lo.Name = tableName

    Set SimpleSpecs = TableSpecs.Create(lo.HeaderRowRange, lo.ListRows(1).Range, Dict)
End Function

'@sub-title Build a cross-table over a fixture row.
'@param tableName String. One of the analysis ListObject names.
'@param headerRow Variant. The header of that block.
'@param dataRow Variant. One data row.
'@return CrossTable. A cross-table created for its specification alone.
Private Function FixtureTable(ByVal tableName As String, _
                              ByVal headerRow As Variant, _
                              ByVal dataRow As Variant) As CrossTable
    Dim specs As TableSpecs
    Dim outSh As Worksheet

    Set specs = SimpleSpecs(tableName, headerRow, dataRow)
    Set outSh = ResetFixtureSheet(OUTPUT_SHEET)

    Set FixtureTable = CrossTable.Create(specs, outSh, LinelistData())
End Function

'@sub-title Build the series builder of a fixture row.
'@param tableName String. One of the analysis ListObject names.
'@param headerRow Variant. The header of that block.
'@param dataRow Variant. One data row.
'@return GraphSeries. A builder before Series() is read.
Private Function BuildGraphSeries(ByVal tableName As String, _
                                  ByVal headerRow As Variant, _
                                  ByVal dataRow As Variant) As GraphSeries
    Set BuildGraphSeries = GraphSeries.Create(FixtureTable(tableName, headerRow, dataRow))
End Function

'@sub-title One univariate data row with the flags the caller passes.
Private Function UnivariateRow(ByVal percentage As String, _
                               ByVal graph As String, _
                               ByVal flip As String) As Variant
    UnivariateRow = Array("S1", "A univariate table", "choi_v1", "no", "", "", _
                          "", percentage, graph, flip)
End Function

'@sub-title One bivariate data row with the graph setting the caller passes.
Private Function BivariateRow(ByVal graph As String) As Variant
    BivariateRow = Array("S1", "A bivariate table", "choi_v1", "choi_h2", "no", _
                         "", "", "", "no", graph, "no")
End Function

'@sub-title One spatio-temporal data row with the geo count and graph setting.
Private Function SpatioTemporalRow(ByVal geoMax As String, _
                                   ByVal graph As String) As Variant
    SpatioTemporalRow = Array("S1", "date_v1", "adm1_zone", geoMax, _
                              "A spatio-temporal table", "geo", "", "", "", graph)
End Function

'@sub-title One time series data row.
Private Function TimeSeriesRow() As Variant
    TimeSeriesRow = Array("Series 1", "S1", "date_v1", "choi_v1", "First table", _
                          "no", "", "", "", "no", "yes", "1")
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Build the assertion harness and the dictionary the specifications need.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run. A Private lifecycle hook is the trap that has cost five
'modules a run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    Dim dictSheet As Worksheet

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestGraphSeries"

    Set dictSheet = ResetFixtureSheet(DICT_SHEET)
    WriteRow dictSheet.Range("A1"), "variable name", "control", "control details"
    WriteRow dictSheet.Range("A2"), "choi_v1", "choice_manual", "list_manual"
    WriteRow dictSheet.Range("A3"), "choi_h2", "choice_manual", "list_manual"
    WriteRow dictSheet.Range("A4"), "date_v1", "date", ""
    WriteRow dictSheet.Range("A5"), "adm1_zone", "geo", ""

    Set Dict = LLdictionary.Create(dictSheet, 1, 1, 1)
End Sub

'@sub-title Print the results and drop every fixture this module made.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    ReleaseTableName TABLE_UNIVARIATE
    ReleaseTableName TABLE_BIVARIATE
    ReleaseTableName TABLE_SPATIOTEMPORAL
    ReleaseTableName TABLE_TIMESERIES

    DeleteWorksheets SIMPLE_SHEET, OUTPUT_SHEET, DICT_SHEET

    RestoreApp
    Set Dict = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updates before each test.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush the results of each test.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory validation
'===============================================================================

'@sub-title Verify Create refuses a Nothing cross-table.
'@TestMethod("GraphSeries")
Public Sub TestCreateRejectsNothingTable()
    CustomTestSetTitles Assert, "GraphSeries", "TestCreateRejectsNothingTable"
    On Error GoTo TestFail

    Dim builder As GraphSeries
    Dim errNumber As Long

    On Error Resume Next
    Set builder = GraphSeries.Create(Nothing)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (builder Is Nothing), "Create with a Nothing cross-table gives nothing back"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "Create with a Nothing cross-table raises InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingTable", Err.Number, Err.Description
End Sub

'@sub-title Verify Create refuses a time series table.
'@details
'A time series table is drawn by TimeSeriesGraphs, from the three graph setup
'tables. This is the one branch the two halves used to share a type for.
'@TestMethod("GraphSeries")
Public Sub TestCreateRejectsATimeSeriesTable()
    CustomTestSetTitles Assert, "GraphSeries", "TestCreateRejectsATimeSeriesTable"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim builder As GraphSeries
    Dim errNumber As Long

    Set tabl = FixtureTable(TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRow())

    On Error Resume Next
    Set builder = GraphSeries.Create(tabl)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (builder Is Nothing), "A time series table gives nothing back"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A time series table raises InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsATimeSeriesTable", Err.Number, Err.Description
End Sub

'@section What the builder puts on a chart
'===============================================================================

'@sub-title Verify a univariate table gives one bar series.
'@TestMethod("GraphSeries")
Public Sub TestUnivariateBarGivesOneSeries()
    CustomTestSetTitles Assert, "GraphSeries", "TestUnivariateBarGivesOneSeries"
    On Error GoTo TestFail

    Dim builder As GraphSeries
    Dim buffer As SeriesBuffer

    Set builder = BuildGraphSeries(TABLE_UNIVARIATE, UnivariateHeader(), _
                                   UnivariateRow("no", "yes", "no"))
    Set buffer = builder.Series()

    Assert.AreEqual 1&, buffer.Count, "One series for a table with no percentage"
    Assert.AreEqual "VALUES_COL_1_UA_tab1", buffer.RangeName(1), "It plots the value column"
    Assert.AreEqual "bar", buffer.ChartType(1), "Vertical bars by default"
    Assert.AreEqual "left", buffer.AxisSide(1), "On the primary axis"
    Assert.AreEqual "ROW_CATEGORIES_UA_tab1", buffer.CategoryRange(1), _
                    "Labelled by the row categories of its own table"
    Assert.AreEqual "LABEL_COL_1_UA_tab1", buffer.LegendRange(1), _
                    "And named by its column label"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateBarGivesOneSeries", Err.Number, Err.Description
End Sub

'@sub-title Verify a percentage column adds a point series on the right axis.
'@TestMethod("GraphSeries")
Public Sub TestUnivariateWithPercentageAddsPointOnRight()
    CustomTestSetTitles Assert, "GraphSeries", "TestUnivariateWithPercentageAddsPointOnRight"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer

    Set buffer = BuildGraphSeries(TABLE_UNIVARIATE, UnivariateHeader(), _
                                  UnivariateRow("yes", "yes", "no")).Series()

    Assert.AreEqual 2&, buffer.Count, "The percentage overlay is a second series"
    Assert.AreEqual "PERC_COL_1_UA_tab1", buffer.RangeName(2), "It plots the percentage column"
    Assert.AreEqual "point", buffer.ChartType(2), "Drawn as points"
    Assert.AreEqual "right", buffer.AxisSide(2), "On the secondary axis"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateWithPercentageAddsPointOnRight", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify horizontal bars drop the percentage overlay.
'@details
'The flag is read through TableSpecs, so a cell holding " Yes " flips the chart
'the same way a cell holding "yes" does. Comparing the raw cell let a pasted
'value open the gate in AnalysisOutput and match nothing in the builder.
'@TestMethod("GraphSeries")
Public Sub TestUnivariateFlippedSkipsThePercentageOverlay()
    CustomTestSetTitles Assert, "GraphSeries", "TestUnivariateFlippedSkipsThePercentageOverlay"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer

    Set buffer = BuildGraphSeries(TABLE_UNIVARIATE, UnivariateHeader(), _
                                  UnivariateRow("yes", "yes", " Yes ")).Series()

    Assert.AreEqual 1&, buffer.Count, "A flipped chart carries the values alone"
    Assert.AreEqual "hbar", buffer.ChartType(1), "A cell holding "" Yes "" still flips the chart"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateFlippedSkipsThePercentageOverlay", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a spatio-temporal table builds one series per geographic unit.
'@details
'The column count comes from TableSpecs.GeoCount, which carries the default of
'five and clamps to 1 to 20. Reading the cell raw gave zero series for a "0" and
'fifty for a "50", each phantom one mislabelling the last real series.
'@TestMethod("GraphSeries")
Public Sub TestSpatioTemporalSeriesCountFollowsTheGeoCount()
    CustomTestSetTitles Assert, "GraphSeries", "TestSpatioTemporalSeriesCountFollowsTheGeoCount"
    On Error GoTo TestFail

    Assert.AreEqual 3&, BuildGraphSeries(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                         SpatioTemporalRow("3", "yes")).Series().Count, _
                    "Three geographic units give three series"

    Assert.AreEqual 1&, BuildGraphSeries(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                         SpatioTemporalRow("0", "yes")).Series().Count, _
                    "A count of zero is clamped up to one"

    Assert.AreEqual 20&, BuildGraphSeries(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                          SpatioTemporalRow("50", "yes")).Series().Count, _
                    "A count of fifty is clamped down to twenty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalSeriesCountFollowsTheGeoCount", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a spatio-temporal chart labels its rows from the section.
'@TestMethod("GraphSeries")
Public Sub TestSpatioTemporalUsesSectionRowCategories()
    CustomTestSetTitles Assert, "GraphSeries", "TestSpatioTemporalUsesSectionRowCategories"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer

    Set buffer = BuildGraphSeries(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                  SpatioTemporalRow("2", "yes")).Series()

    Assert.AreEqual 2&, buffer.Count, "Two geographic units give two series"
    Assert.AreEqual "ROW_CATEGORIES_SPT_tab1", buffer.CategoryRange(1), _
                    "The row categories are the ones of the section, which the first row starts"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalUsesSectionRowCategories", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a graph setting with outer spaces still builds the series.
'@TestMethod("GraphSeries")
Public Sub TestGraphSettingWithSpacesStillBuilds()
    CustomTestSetTitles Assert, "GraphSeries", "TestGraphSettingWithSpacesStillBuilds"
    On Error GoTo TestFail

    Dim builder As GraphSeries

    Set builder = BuildGraphSeries(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                   SpatioTemporalRow("4", " Yes "))

    Assert.AreEqual 4&, builder.Series().Count, _
                    "A cell holding "" Yes "" builds the same four series as ""yes"""
    Assert.IsTrue (Not builder.HasCheckings), "And nothing is reported against it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGraphSettingWithSpacesStillBuilds", Err.Number, Err.Description
End Sub

'@section What is skipped, and what says so
'===============================================================================

'@sub-title Verify a graph setting nobody recognises is reported.
'@details
'The Select Case had no Case Else, so an unknown setting pushed no series, said
'nothing, and left AnalysisOutput to build a chart Graphs.Format then raised
'1004 on.
'@TestMethod("GraphSeries")
Public Sub TestUnknownGraphSettingIsReportedAndBuildsNothing()
    CustomTestSetTitles Assert, "GraphSeries", "TestUnknownGraphSettingIsReportedAndBuildsNothing"
    On Error GoTo TestFail

    Dim builder As GraphSeries

    Set builder = BuildGraphSeries(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                   SpatioTemporalRow("4", "sometimes"))

    Assert.AreEqual 0&, builder.Series().Count, "An unknown setting builds no series"
    Assert.IsTrue builder.HasCheckings, "And it is reported"
    Assert.IsTrue (Not builder.CheckingValues Is Nothing), "The report is handed over"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnknownGraphSettingIsReportedAndBuildsNothing", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a percentage chart on a table with no percentages is reported.
'@TestMethod("GraphSeries")
Public Sub TestPercentageChartWithoutPercentageColumnIsReported()
    CustomTestSetTitles Assert, "GraphSeries", "TestPercentageChartWithoutPercentageColumnIsReported"
    On Error GoTo TestFail

    Dim builder As GraphSeries

    Set builder = BuildGraphSeries(TABLE_BIVARIATE, BivariateHeader(), _
                                   BivariateRow("percentages"))

    Assert.AreEqual 0&, builder.Series().Count, "No percentage column, no series"
    Assert.IsTrue builder.HasCheckings, "And the reason is reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPercentageChartWithoutPercentageColumnIsReported", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the series are built once and kept.
'@details
'Series() builds on the first read, so a caller has one call to make and no
'order to remember. Reading it twice gives the same buffer.
'@TestMethod("GraphSeries")
Public Sub TestTheSeriesAreBuiltOnceAndKept()
    CustomTestSetTitles Assert, "GraphSeries", "TestTheSeriesAreBuiltOnceAndKept"
    On Error GoTo TestFail

    Dim builder As GraphSeries
    Dim firstRead As SeriesBuffer
    Dim secondRead As SeriesBuffer

    Set builder = BuildGraphSeries(TABLE_UNIVARIATE, UnivariateHeader(), _
                                   UnivariateRow("yes", "yes", "no"))
    Set firstRead = builder.Series()
    Set secondRead = builder.Series()

    Assert.AreEqual 2&, firstRead.Count, "The first read builds the series"
    Assert.IsTrue (firstRead Is secondRead), "The second read answers the same buffer"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSeriesAreBuiltOnceAndKept", Err.Number, Err.Description
End Sub

'@section The checking keys
'===============================================================================

'@sub-title Verify two builders can file into one report.
'@details
'Checking.Add raises ElementShouldNotExists on a duplicate key and
'Checking.Append replays every key into the target, so a key made of a bare
'counter takes the whole generation down on the second collaborator that files
'anything. AnalysisOutput merges every table of a sheet into one report.
'@TestMethod("GraphSeries")
Public Sub TestTwoBuildersProduceDistinctCheckingKeys()
    CustomTestSetTitles Assert, "GraphSeries", "TestTwoBuildersProduceDistinctCheckingKeys"
    On Error GoTo TestFail

    Dim firstBuilder As GraphSeries
    Dim secondBuilder As GraphSeries
    Dim builtSeries As SeriesBuffer
    Dim report As Checking
    Dim errNumber As Long

    Set firstBuilder = BuildGraphSeries(TABLE_BIVARIATE, BivariateHeader(), _
                                        BivariateRow("sometimes"))
    Set builtSeries = firstBuilder.Series()

    Set secondBuilder = BuildGraphSeries(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                         SpatioTemporalRow("4", "sometimes"))
    Set builtSeries = secondBuilder.Series()

    Assert.IsTrue firstBuilder.HasCheckings, "The bivariate table filed an entry"
    Assert.IsTrue secondBuilder.HasCheckings, "So did the spatio-temporal one"

    Set report = Checking.Create("Analysis output")

    On Error Resume Next
    report.Append firstBuilder.CheckingValues
    report.Append secondBuilder.CheckingValues
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual 0&, errNumber, _
                    "Two builders file into one report without a key collision"
    Assert.IsTrue (report.Length > 0), "And the report carries their entries"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTwoBuildersProduceDistinctCheckingKeys", _
                         Err.Number, Err.Description
End Sub
