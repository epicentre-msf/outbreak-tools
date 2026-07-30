Attribute VB_Name = "TestGraphSpecs"
Attribute VB_Description = "Tests for GraphSpecs class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for GraphSpecs class")

'@description
'Drives GraphSpecs, which decides what goes on every chart a linelist draws. It
'works in two modes: simple, from one built CrossTable, and complex, from the
'three time series setup tables. This module had never been imported, compiled
'or run before 2026-07-31.
'
'THE FIXTURES ARE REAL LISTOBJECTS WITH THE REAL HEADERS
'-------------------------------------------------------------------------------
'CreateRangeSpecs recognises its three tables by name, so the fixture names them
'Tab_Graph_TimeSeries, Tab_TimeSeries_Analysis and Tab_Label_TSGraph. The
'headers are the ones measured on 2026-07-31 in .mock/setup_mock.xlsb,
'src/bin/setup/setup.xlsb and releases/main/setup/setup_main-2026-06-11.xlsb,
'which all agree. The class finds most of its columns by a partial match --
'"axis" against "Y-Axis", "type" against "Chart type" -- so a fixture with tidy
'invented headers would exercise none of that.
'
'THE LINELIST SPECIFICATIONS ARE A BARE INSTANCE WITH A DICTIONARY IN IT
'-------------------------------------------------------------------------------
'GraphSpecs reads exactly one thing off LinelistSpecs: Dictionary(), which it
'hands to TableSpecs.Create. TestAssignDictionary fills that field, so the suite
'needs no designer workbook and none of the seven sheets LinelistSpecs.Create
'asks for.
'@depends GraphSpecs, CrossTable, TableSpecs, LinelistSpecs, LLdictionary, Checking, BetterArray, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "GraphSpecsFixture"
Private Const SIMPLE_SHEET As String = "GraphSpecsSimple"
Private Const OUTPUT_SHEET As String = "GraphSpecsOutput"
Private Const SECOND_OUTPUT_SHEET As String = "GraphSpecsOutput2"
Private Const DICT_SHEET As String = "GraphSpecsDict"

' The three setup tables of the time series graph block, named the way the
' setup workbook names them.
Private Const TABLE_GRAPH_TIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TABLE_TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const TABLE_LABEL_TSGRAPH As String = "Tab_Label_TSGraph"

' The simple-mode blocks.
Private Const TABLE_UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const TABLE_BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TABLE_SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"

' Where each fixture table starts on the complex fixture sheet.
Private Const GRAPH_HEADER_ROW As Long = 2
Private Const TIMESERIES_HEADER_ROW As Long = 8
Private Const TITLES_HEADER_ROW As Long = 14

' The header row of a simple-mode fixture. Data row N sits at HEADER_ROW + N,
' so TableId reads "<prefix>_tabN".
Private Const SIMPLE_HEADER_ROW As Long = 5

' The identifiers the fixtures produce. The first data row of a block is
' "<prefix>_tab1", and a first row always starts its own section.
Private Const TS_TABLE_ID As String = "TS_tab1"
Private Const SERIES_ID As String = "Series 1"
Private Const GRAPH_ID As String = "G1"
Private Const CATEGORY_CHOICE As String = "C1"
Private Const TOTAL_CHOICE As String = "Total"

Private Assert As CustomTest
Private Dict As LLdictionary

'@section Header rows measured in the setup workbooks
'===============================================================================

'@sub-title Header of Tab_Graph_TimeSeries, 12 columns.
Private Function GraphTableHeader() As Variant
    GraphTableHeader = Array( _
        "Graph title (select)", "Series title (select)", "Graph ID", _
        "Series ID", "Graph order", "Time variable (row)", _
        "Group by variable (column)", "Choices", "Label", _
        "Plot values or percentages", "Chart type", "Y-Axis")
End Function

'@sub-title Header of Tab_TimeSeries_Analysis, 12 columns.
Private Function TimeSeriesHeader() As Variant
    TimeSeriesHeader = Array( _
        "Series ID", "Section", "Time variable (row)", _
        "Group by variable (column)", "Title (header)", "Add missing data", _
        "Summary function", "Summary label", "Format", "Add percentage", _
        "Add total", "Table order")
End Function

'@sub-title Header of Tab_Label_TSGraph, 3 columns.
Private Function TitlesHeader() As Variant
    TitlesHeader = Array("Graph title", "Graph order", "Graph ID")
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

'@sub-title Header of Tab_SpatioTemporal_Analysis, 10 columns.
'@details This block carries no "flip coordinates" column, so FlipCoordinates
'answers False for every row of it.
Private Function SpatioTemporalHeader() As Variant
    SpatioTemporalHeader = Array( _
        "Section (select)", "Time variable (row)", "Geo/HF variable (column)", _
        "N geo max", "Title (header)", "Spatial type", "Summary function", _
        "Summary label", "Format", "Add graph")
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
'The names this suite writes are workbook-scoped and outlive a Cells.Clear, and
'TimeSeriesColumnName asks the sheet for them, so a name left behind changes the
'answer for the next test. The owning worksheet is read off RefersToRange. That
'read holds whatever the sheet is called, while the RefersTo text quotes a name
'carrying a space. A name holding a value has no range and raises, so the read is
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
'EnsureWorksheet is asked NOT to clear. Its ClearWorksheet runs Cells.Clear
'over the whole 1,048,576 by 16,384 sheet, which costs seconds a call and took
'a green analyses run past the runner's cap. Every fixture here lives inside the
'first fifty rows, so a bounded block is enough.
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

'@sub-title Build the three time series setup tables.
'@details
'One graph row, one time series row and one title row, so the identifiers are
'known: the time series table id is TS_tab1 and it starts its own section.
'@param graphId String. The "Graph ID" cell of the graph row.
'@param choiceValue String. The "Choices" cell of the graph row.
'@param percValue String. The "Plot values or percentages" cell.
'@param seriesId String. The "Series ID" cell of the graph row.
'@param titleGraphId String. The "Graph ID" cell of the title row.
'@return BetterArray. The three ListObjects, in the order CreateRangeSpecs wants.
Private Function BuildComplexFixture(ByVal graphId As String, _
                                     ByVal choiceValue As String, _
                                     ByVal percValue As String, _
                                     ByVal seriesId As String, _
                                     ByVal titleGraphId As String) As BetterArray
    Dim sh As Worksheet
    Dim loTable As BetterArray
    Dim graphLo As ListObject
    Dim tsLo As ListObject
    Dim titleLo As ListObject

    ReleaseTableName TABLE_GRAPH_TIMESERIES
    ReleaseTableName TABLE_TIMESERIES
    ReleaseTableName TABLE_LABEL_TSGRAPH

    Set sh = ResetFixtureSheet(FIXTURE_SHEET)

    WriteMatrix sh.Cells(GRAPH_HEADER_ROW, 1), RowsToMatrix(Array(GraphTableHeader()))
    WriteMatrix sh.Cells(GRAPH_HEADER_ROW + 1, 1), RowsToMatrix(Array( _
        Array("A graph", "A series", graphId, seriesId, "1", "date_v1", _
              "choi_v1", choiceValue, "Series label", percValue, "line", "left")))
    Set graphLo = sh.ListObjects.Add(xlSrcRange, _
                                     sh.Range(sh.Cells(GRAPH_HEADER_ROW, 1), _
                                              sh.Cells(GRAPH_HEADER_ROW + 1, 12)), _
                                     , xlYes)
    graphLo.Name = TABLE_GRAPH_TIMESERIES

    WriteMatrix sh.Cells(TIMESERIES_HEADER_ROW, 1), RowsToMatrix(Array(TimeSeriesHeader()))
    WriteMatrix sh.Cells(TIMESERIES_HEADER_ROW + 1, 1), RowsToMatrix(Array( _
        Array(SERIES_ID, "S1", "date_v1", "choi_v1", "First table", "no", _
              "", "", "", "no", "yes", "1")))
    Set tsLo = sh.ListObjects.Add(xlSrcRange, _
                                  sh.Range(sh.Cells(TIMESERIES_HEADER_ROW, 1), _
                                           sh.Cells(TIMESERIES_HEADER_ROW + 1, 12)), _
                                  , xlYes)
    tsLo.Name = TABLE_TIMESERIES

    WriteMatrix sh.Cells(TITLES_HEADER_ROW, 1), RowsToMatrix(Array(TitlesHeader()))
    WriteMatrix sh.Cells(TITLES_HEADER_ROW + 1, 1), RowsToMatrix(Array( _
        Array("A graph title", "1", titleGraphId)))
    Set titleLo = sh.ListObjects.Add(xlSrcRange, _
                                     sh.Range(sh.Cells(TITLES_HEADER_ROW, 1), _
                                              sh.Cells(TITLES_HEADER_ROW + 1, 3)), _
                                     , xlYes)
    titleLo.Name = TABLE_LABEL_TSGRAPH

    Set loTable = New BetterArray
    loTable.LowerBound = 1
    loTable.Push graphLo, tsLo, titleLo

    Set BuildComplexFixture = loTable
End Function

'@sub-title The fixture the resolution tests share.
'@param choiceValue String. The "Choices" cell of the graph row.
'@param percValue String. The "Plot values or percentages" cell.
'@return BetterArray. The three ListObjects.
Private Function StandardComplexFixture(ByVal choiceValue As String, _
                                        ByVal percValue As String) As BetterArray
    Set StandardComplexFixture = BuildComplexFixture(GRAPH_ID, choiceValue, _
                                                     percValue, SERIES_ID, GRAPH_ID)
End Function

'@sub-title Build the output worksheet the time series charts are drawn on.
'@details
'When the categories are wanted, the sheet carries a COLUMN_CATEGORIES_ range
'over two headed cells, each with a LABEL_COL_ name of its own. That is the
'shape CrossTable leaves behind, and it is what TimeSeriesColumnName walks.
'@param withCategories Boolean. False leaves the sheet with no named range,
'which is the state a table skipped during the table pass leaves.
'@param sheetName String. The worksheet to build.
'@return Worksheet. The output worksheet.
Private Function BuildOutputSheet(ByVal withCategories As Boolean, _
                                  Optional ByVal sheetName As String = OUTPUT_SHEET) As Worksheet
    Dim sh As Worksheet

    Set sh = ResetFixtureSheet(sheetName)

    If withCategories Then
        sh.Range("B2").Value = CATEGORY_CHOICE
        sh.Range("C2").Value = "C2"
        sh.Range("B2:C2").Name = "COLUMN_CATEGORIES_" & TS_TABLE_ID
        sh.Range("B2").Name = "LABEL_COL_1_" & TS_TABLE_ID
        sh.Range("C2").Name = "LABEL_COL_2_" & TS_TABLE_ID
    End If

    Set BuildOutputSheet = sh
End Function

'@sub-title A LinelistSpecs carrying nothing but a dictionary.
'@details
'GraphSpecs reads Dictionary() off it and nothing else, and Dictionary() hands
'back the injected instance without asking whether Prepare has run.
'@return LinelistSpecs. The collaborator CreateRangeSpecs requires.
Private Function LinelistData() As LinelistSpecs
    Dim stub As LinelistSpecs

    Set stub = New LinelistSpecs
    stub.TestAssignDictionary Dict
    Set LinelistData = stub
End Function

'@sub-title Build a simple-mode fixture table and answer its specification.
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

'@sub-title Build a GraphSpecs in simple mode over a fixture row.
'@details
'The cross-table is created for its specification alone. The series builder
'reads the scope, the flags and the geographic unit count, none of which needs
'the geometry a Build would have set. That is how WriteSpatioTemporalGraph
'drives the pair too.
'@param tableName String. One of the analysis ListObject names.
'@param headerRow Variant. The header of that block.
'@param dataRow Variant. One data row.
'@return GraphSpecs. An instance in simple mode, before CreateSeries.
Private Function SimpleGraphSpecs(ByVal tableName As String, _
                                  ByVal headerRow As Variant, _
                                  ByVal dataRow As Variant) As GraphSpecs
    Dim specs As TableSpecs
    Dim outSh As Worksheet
    Dim tabl As CrossTable

    Set specs = SimpleSpecs(tableName, headerRow, dataRow)
    Set outSh = BuildOutputSheet(withCategories:=False)
    Set tabl = CrossTable.Create(specs, outSh, LinelistData())

    Set SimpleGraphSpecs = GraphSpecs.Create(tabl)
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

'@section Reading what was built
'===============================================================================

'@sub-title The series names of the first graph of a complex build.
Private Function FirstGraphSeriesNames(ByVal specs As GraphSpecs) As BetterArray
    Dim lists As BetterArray
    Dim entry As BetterArray

    Set FirstGraphSeriesNames = New BetterArray

    Set lists = specs.SpecsLists()
    If lists.Length = 0 Then Exit Function

    Set entry = lists.Item(lists.LowerBound)
    Set FirstGraphSeriesNames = entry.Item(1)
End Function

'@sub-title The first series name of the first graph, or an empty string.
Private Function FirstSeriesName(ByVal specs As GraphSpecs) As String
    Dim names As BetterArray

    Set names = FirstGraphSeriesNames(specs)
    If names.Length = 0 Then Exit Function

    FirstSeriesName = CStr(names.Item(names.LowerBound))
End Function

'@sub-title The row category name of the first series of a simple build.
Private Function FirstRowCategory(ByVal specs As GraphSpecs) As String
    If specs.NumberOfSeries() = 0 Then Exit Function
    FirstRowCategory = specs.SeriesLabel(1)
End Function

'@sub-title The title of the first graph of a complex build.
Private Function FirstGraphTitle(ByVal specs As GraphSpecs) As String
    Dim lists As BetterArray
    Dim entry As BetterArray

    Set lists = specs.SpecsLists()
    If lists.Length = 0 Then Exit Function

    Set entry = lists.Item(lists.LowerBound)
    FirstGraphTitle = CStr(entry.Item(7))
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
    Assert.SetModuleName "TestGraphSpecs"

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

    ReleaseTableName TABLE_GRAPH_TIMESERIES
    ReleaseTableName TABLE_TIMESERIES
    ReleaseTableName TABLE_LABEL_TSGRAPH
    ReleaseTableName TABLE_UNIVARIATE
    ReleaseTableName TABLE_BIVARIATE
    ReleaseTableName TABLE_SPATIOTEMPORAL

    DeleteWorksheets FIXTURE_SHEET, SIMPLE_SHEET, OUTPUT_SHEET, _
                     SECOND_OUTPUT_SHEET, DICT_SHEET

    RestoreApp
    Set Dict = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updates before each test.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run. The first assertion of each test opens the checking, which
'picks up the titles set a line above it. Calling BeginTest here would open it
'with whatever titles are pending and file every result under the default label.
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
'@TestMethod("GraphSpecs")
Public Sub TestCreateRejectsNothingTable()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRejectsNothingTable"
    On Error GoTo TestFail

    Dim specs As GraphSpecs
    Dim errNumber As Long

    On Error Resume Next
    Set specs = GraphSpecs.Create(Nothing)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (specs Is Nothing), "Create with a Nothing cross-table gives nothing back"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "Create with a Nothing cross-table raises InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingTable", Err.Number, Err.Description
End Sub

'@sub-title Verify CreateRangeSpecs refuses a Nothing ListObject collection.
'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsNothingLoTable()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsNothingLoTable"
    On Error GoTo TestFail

    Dim specs As GraphSpecs
    Dim outSh As Worksheet

    Set outSh = BuildOutputSheet(withCategories:=False)

    On Error Resume Next
    Set specs = GraphSpecs.CreateRangeSpecs(Nothing, outSh, LinelistData())
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (specs Is Nothing), "CreateRangeSpecs refuses a Nothing collection"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRangeSpecsRejectsNothingLoTable", Err.Number, Err.Description
End Sub

'@sub-title Verify CreateRangeSpecs refuses a Nothing output worksheet.
'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsNothingSheet()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsNothingSheet"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim specs As GraphSpecs

    Set loTable = StandardComplexFixture(CATEGORY_CHOICE, "values")

    On Error Resume Next
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, Nothing, LinelistData())
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (specs Is Nothing), "CreateRangeSpecs refuses a Nothing worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRangeSpecsRejectsNothingSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify CreateRangeSpecs refuses a Nothing linelist specification.
'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsNothingLData()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsNothingLData"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    Set loTable = StandardComplexFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=True)

    On Error Resume Next
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, Nothing)
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (specs Is Nothing), "CreateRangeSpecs refuses Nothing linelist specifications"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRangeSpecsRejectsNothingLData", Err.Number, Err.Description
End Sub

'@sub-title Verify CreateRangeSpecs refuses a collection of the wrong size.
'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsWrongCount()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsWrongCount"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim shortList As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs
    Dim errNumber As Long

    Set loTable = StandardComplexFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=True)

    Set shortList = New BetterArray
    shortList.LowerBound = 1
    shortList.Push loTable.Item(1)

    On Error Resume Next
    Set specs = GraphSpecs.CreateRangeSpecs(shortList, outSh, LinelistData())
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (specs Is Nothing), "CreateRangeSpecs refuses one ListObject"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A collection of the wrong size raises InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRangeSpecsRejectsWrongCount", Err.Number, Err.Description
End Sub

'@sub-title Verify the three tables are recognised by their names.
'@details
'The order used to be read off a caption two rows above each header row. That
'gap is two rows in a release setup file and four rows in every development one,
'so the check refused a whole family of workbooks. The names are identical in
'all six files that were measured.
'@TestMethod("GraphSpecs")
Public Sub TestTheSetupTableNamesAreWhatIdentifiesEachTable()
    CustomTestSetTitles Assert, "GraphSpecs", "TestTheSetupTableNamesAreWhatIdentifiesEachTable"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    ' The fixture writes no caption above any of the three headers.
    Set loTable = StandardComplexFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=True)

    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())

    Assert.IsTrue (Not specs Is Nothing), _
                  "The three tables are accepted on their names alone"
    Assert.AreEqual outSh.Name, specs.Wksh.Name, _
                    "Wksh answers the output worksheet in complex mode"
    Assert.AreEqual 0&, specs.NumberOfGraphs, _
                    "No graph is built before CreateSeries runs"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSetupTableNamesAreWhatIdentifiesEachTable", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the three tables in the wrong order are refused.
'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsTablesInTheWrongOrder()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsTablesInTheWrongOrder"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim swapped As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs
    Dim errNumber As Long

    Set loTable = StandardComplexFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=True)

    Set swapped = New BetterArray
    swapped.LowerBound = 1
    swapped.Push loTable.Item(2), loTable.Item(1), loTable.Item(3)

    On Error Resume Next
    Set specs = GraphSpecs.CreateRangeSpecs(swapped, outSh, LinelistData())
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (specs Is Nothing), "The time series table cannot stand in for the graph table"
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), errNumber, _
                    "The wrong order raises ErrorUnexpectedState"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRangeSpecsRejectsTablesInTheWrongOrder", _
                         Err.Number, Err.Description
End Sub

'@section Complex mode - resolving a series to a named range
'===============================================================================

'@sub-title Verify a category choice resolves to the VALUES_COL_ range.
'@TestMethod("GraphSpecs")
Public Sub TestCategoryChoiceResolvesToValuesCol()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCategoryChoiceResolvesToValuesCol"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    Set loTable = StandardComplexFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())
    specs.CreateSeries

    Assert.AreEqual 1&, specs.NumberOfGraphs, "One graph identifier gives one graph"
    Assert.AreEqual "VALUES_COL_1_" & TS_TABLE_ID, FirstSeriesName(specs), _
                    "A category choice plots the values of its own column"
    Assert.AreEqual "A graph title", FirstGraphTitle(specs), _
                    "The graph carries the title of its row in the titles table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCategoryChoiceResolvesToValuesCol", Err.Number, Err.Description
End Sub

'@sub-title Verify a category choice asking for percentages resolves to PERC_COL_.
'@TestMethod("GraphSpecs")
Public Sub TestCategoryPercentageResolvesToPercCol()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCategoryPercentageResolvesToPercCol"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    Set loTable = StandardComplexFixture(CATEGORY_CHOICE, "percentages")
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())
    specs.CreateSeries

    Assert.AreEqual "PERC_COL_1_" & TS_TABLE_ID, FirstSeriesName(specs), _
                    "A category asking for percentages plots its percentage column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCategoryPercentageResolvesToPercCol", Err.Number, Err.Description
End Sub

'@sub-title Verify the total choice resolves to TOTAL_COL_VALUES_.
'@TestMethod("GraphSpecs")
Public Sub TestTotalChoiceResolvesToTotalColValues()
    CustomTestSetTitles Assert, "GraphSpecs", "TestTotalChoiceResolvesToTotalColValues"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    Set loTable = StandardComplexFixture(TOTAL_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())
    specs.CreateSeries

    Assert.AreEqual "TOTAL_COL_VALUES_" & TS_TABLE_ID, FirstSeriesName(specs), _
                    "The total plots the values of the total column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTotalChoiceResolvesToTotalColValues", Err.Number, Err.Description
End Sub

'@sub-title Verify a percentage total resolves to a range CrossTable creates.
'@details
'It used to build TOTAL_LABEL_PERC_ by replacing COL with PERC in the label
'name. No code anywhere creates that name, so every percentage total in a time
'series chart handed Graphs a name that did not resolve, and AddSeries then let
'AddLabels relabel the series before it. TOTAL_PERC_VALUES_ is the one
'CrossTable writes.
'@TestMethod("GraphSpecs")
Public Sub TestTotalPercentageResolvesToTotalPercValues()
    CustomTestSetTitles Assert, "GraphSpecs", "TestTotalPercentageResolvesToTotalPercValues"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs
    Dim seriesName As String

    Set loTable = StandardComplexFixture(TOTAL_CHOICE, "percentages")
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())
    specs.CreateSeries

    seriesName = FirstSeriesName(specs)

    Assert.AreEqual "TOTAL_PERC_VALUES_" & TS_TABLE_ID, seriesName, _
                    "A percentage total plots the percentage column of the total"
    Assert.IsTrue (InStr(1, seriesName, "TOTAL_LABEL_PERC_") = 0), _
                  "The name no code creates is gone"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTotalPercentageResolvesToTotalPercValues", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the total is matched without regard to case and spaces.
'@TestMethod("GraphSpecs")
Public Sub TestTotalChoiceIsMatchedWithoutRegardToCase()
    CustomTestSetTitles Assert, "GraphSpecs", "TestTotalChoiceIsMatchedWithoutRegardToCase"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    Set loTable = StandardComplexFixture(" total ", "values")
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())
    specs.CreateSeries

    Assert.AreEqual "TOTAL_COL_VALUES_" & TS_TABLE_ID, FirstSeriesName(specs), _
                    "A cell holding "" total "" still names the total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTotalChoiceIsMatchedWithoutRegardToCase", _
                         Err.Number, Err.Description
End Sub

'@section Complex mode - what is skipped, and what says so
'===============================================================================

'@sub-title Verify a table with no column categories is skipped and reported.
'@details
'The table row may have been skipped during the table pass, and then
'COLUMN_CATEGORIES_ was never created. Reading it raised 1004, and
'WriteTimeSeriesGraphs has no handler around the call that triggers the build,
'so every remaining chart of the sheet went with it.
'@TestMethod("GraphSpecs")
Public Sub TestMissingColumnCategoriesLogsAndSkips()
    CustomTestSetTitles Assert, "GraphSpecs", "TestMissingColumnCategoriesLogsAndSkips"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    Set loTable = StandardComplexFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=False)
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())
    specs.CreateSeries

    Assert.AreEqual 0&, specs.NumberOfGraphs, _
                    "A graph whose only series cannot be resolved is left out"
    Assert.IsTrue specs.HasCheckings, _
                  "The missing column categories are reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestMissingColumnCategoriesLogsAndSkips", Err.Number, Err.Description
End Sub

'@sub-title Verify a series id absent from the time series table is reported.
'@TestMethod("GraphSpecs")
Public Sub TestUnknownSeriesIdLogsAndSkips()
    CustomTestSetTitles Assert, "GraphSpecs", "TestUnknownSeriesIdLogsAndSkips"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    Set loTable = BuildComplexFixture(GRAPH_ID, CATEGORY_CHOICE, "values", _
                                      "A series nothing defines", GRAPH_ID)
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())
    specs.CreateSeries

    Assert.AreEqual 0&, specs.NumberOfGraphs, "A graph with no resolvable series is left out"
    Assert.IsTrue specs.HasCheckings, "The unknown series identifier is reported"
    Assert.IsTrue (Not specs.CheckingValues Is Nothing), _
                  "The report is handed over when there is one"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnknownSeriesIdLogsAndSkips", Err.Number, Err.Description
End Sub

'@sub-title Verify a blank graph identifier column builds no graph at all.
'@details
'BetterArray hands out a one-element array holding Empty when it is asked for
'the items of an empty array, so cloning an empty list gave a length of 1. That
'one phantom identifier produced one phantom graph, a chart with no series, and
'a 1004 out of Graphs.Format with no handler above it.
'@TestMethod("GraphSpecs")
Public Sub TestEmptyGraphIdColumnBuildsNoGraph()
    CustomTestSetTitles Assert, "GraphSpecs", "TestEmptyGraphIdColumnBuildsNoGraph"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    Set loTable = BuildComplexFixture(vbNullString, CATEGORY_CHOICE, "values", _
                                      SERIES_ID, GRAPH_ID)
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())
    specs.CreateSeries

    Assert.AreEqual 0&, specs.NumberOfGraphs, "A blank identifier column gives no graph"
    Assert.AreEqual 0&, specs.SpecsLists().Length, "And the collection handed out is empty"
    Assert.IsTrue specs.HasCheckings, "The empty setup table is reported"
    Assert.IsTrue (Not specs.Valid()), "Valid says so too"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEmptyGraphIdColumnBuildsNoGraph", Err.Number, Err.Description
End Sub

'@sub-title Verify a graph identifier is matched in the titles table whatever its case.
'@details
'Two tables typed by two hands. A case difference used to give a chart with an
'empty title and a navigation entry reading "Go to graph: ".
'@TestMethod("GraphSpecs")
Public Sub TestGraphTitleIsMatchedWithoutRegardToCase()
    CustomTestSetTitles Assert, "GraphSpecs", "TestGraphTitleIsMatchedWithoutRegardToCase"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    Set loTable = BuildComplexFixture(GRAPH_ID, CATEGORY_CHOICE, "values", _
                                      SERIES_ID, LCase$(GRAPH_ID))
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())
    specs.CreateSeries

    Assert.AreEqual "A graph title", FirstGraphTitle(specs), _
                    "The title row is found when only its case differs"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGraphTitleIsMatchedWithoutRegardToCase", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a graph with no title row is drawn and reported.
'@TestMethod("GraphSpecs")
Public Sub TestGraphWithNoTitleRowIsReported()
    CustomTestSetTitles Assert, "GraphSpecs", "TestGraphWithNoTitleRowIsReported"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim specs As GraphSpecs

    Set loTable = BuildComplexFixture(GRAPH_ID, CATEGORY_CHOICE, "values", _
                                      SERIES_ID, "Another graph")
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, outSh, LinelistData())
    specs.CreateSeries

    Assert.AreEqual 1&, specs.NumberOfGraphs, "The graph is still drawn"
    Assert.AreEqual vbNullString, FirstGraphTitle(specs), "And it carries no title"
    Assert.IsTrue specs.HasCheckings, "The missing title row is reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGraphWithNoTitleRowIsReported", Err.Number, Err.Description
End Sub

'@section Simple mode
'===============================================================================

'@sub-title Verify a univariate table gives one bar series.
'@TestMethod("GraphSpecs")
Public Sub TestUnivariateBarPushesOneSeries()
    CustomTestSetTitles Assert, "GraphSpecs", "TestUnivariateBarPushesOneSeries"
    On Error GoTo TestFail

    Dim specs As GraphSpecs

    Set specs = SimpleGraphSpecs(TABLE_UNIVARIATE, UnivariateHeader(), _
                                 UnivariateRow("no", "yes", "no"))
    specs.CreateSeries

    Assert.AreEqual 1&, specs.NumberOfSeries, "One series for a table with no percentage"
    Assert.AreEqual "VALUES_COL_1_UA_tab1", specs.SeriesName(1), "It plots the value column"
    Assert.AreEqual "bar", specs.SeriesType(1), "Vertical bars by default"
    Assert.AreEqual "left", specs.SeriesPos(1), "On the primary axis"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateBarPushesOneSeries", Err.Number, Err.Description
End Sub

'@sub-title Verify a percentage column adds a point series on the right axis.
'@TestMethod("GraphSpecs")
Public Sub TestUnivariateWithPercentageAddsPointOnRight()
    CustomTestSetTitles Assert, "GraphSpecs", "TestUnivariateWithPercentageAddsPointOnRight"
    On Error GoTo TestFail

    Dim specs As GraphSpecs

    Set specs = SimpleGraphSpecs(TABLE_UNIVARIATE, UnivariateHeader(), _
                                 UnivariateRow("yes", "yes", "no"))
    specs.CreateSeries

    Assert.AreEqual 2&, specs.NumberOfSeries, "The percentage overlay is a second series"
    Assert.AreEqual "PERC_COL_1_UA_tab1", specs.SeriesName(2), "It plots the percentage column"
    Assert.AreEqual "point", specs.SeriesType(2), "Drawn as points"
    Assert.AreEqual "right", specs.SeriesPos(2), "On the secondary axis"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateWithPercentageAddsPointOnRight", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify horizontal bars drop the percentage overlay.
'@details
'The flag is read through TableSpecs, so a cell holding " Yes " flips the chart
'the same way a cell holding "yes" does. Comparing the raw cell here let a
'pasted value open the gate in AnalysisOutput and match nothing in this class.
'@TestMethod("GraphSpecs")
Public Sub TestUnivariateFlippedSkipsThePercentageOverlay()
    CustomTestSetTitles Assert, "GraphSpecs", "TestUnivariateFlippedSkipsThePercentageOverlay"
    On Error GoTo TestFail

    Dim specs As GraphSpecs

    Set specs = SimpleGraphSpecs(TABLE_UNIVARIATE, UnivariateHeader(), _
                                 UnivariateRow("yes", "yes", " Yes "))
    specs.CreateSeries

    Assert.AreEqual 1&, specs.NumberOfSeries, "A flipped chart carries the values alone"
    Assert.AreEqual "hbar", specs.SeriesType(1), "A cell holding "" Yes "" still flips the chart"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateFlippedSkipsThePercentageOverlay", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a spatio-temporal table builds one series per geographic unit.
'@details
'The column count comes from TableSpecs.GeoCount, which carries the default of
'five and clamps to 1..20. Reading the cell raw here gave zero series for a "0"
'and fifty for a "50", each phantom one mislabelling the last real series.
'@TestMethod("GraphSpecs")
Public Sub TestSpatioTemporalSeriesCountFollowsTheGeoCount()
    CustomTestSetTitles Assert, "GraphSpecs", "TestSpatioTemporalSeriesCountFollowsTheGeoCount"
    On Error GoTo TestFail

    Dim specs As GraphSpecs

    Set specs = SimpleGraphSpecs(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                 SpatioTemporalRow("3", "yes"))
    specs.CreateSeries
    Assert.AreEqual 3&, specs.NumberOfSeries, "Three geographic units give three series"

    Set specs = SimpleGraphSpecs(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                 SpatioTemporalRow("0", "yes"))
    specs.CreateSeries
    Assert.AreEqual 1&, specs.NumberOfSeries, "A count of zero is clamped up to one"

    Set specs = SimpleGraphSpecs(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                 SpatioTemporalRow("50", "yes"))
    specs.CreateSeries
    Assert.AreEqual 20&, specs.NumberOfSeries, "A count of fifty is clamped down to twenty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalSeriesCountFollowsTheGeoCount", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a spatio-temporal chart labels its rows from the section.
'@TestMethod("GraphSpecs")
Public Sub TestSpatioTemporalUsesSectionRowCategories()
    CustomTestSetTitles Assert, "GraphSpecs", "TestSpatioTemporalUsesSectionRowCategories"
    On Error GoTo TestFail

    Dim specs As GraphSpecs

    Set specs = SimpleGraphSpecs(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                 SpatioTemporalRow("2", "yes"))
    specs.CreateSeries

    Assert.AreEqual 2&, specs.NumberOfSeries, "Two geographic units give two series"
    Assert.AreEqual "ROW_CATEGORIES_SPT_tab1", FirstRowCategory(specs), _
                    "The row categories are the ones of the section, which the first row starts"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalUsesSectionRowCategories", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a graph setting with outer spaces still builds the series.
'@TestMethod("GraphSpecs")
Public Sub TestGraphSettingWithSpacesStillBuilds()
    CustomTestSetTitles Assert, "GraphSpecs", "TestGraphSettingWithSpacesStillBuilds"
    On Error GoTo TestFail

    Dim specs As GraphSpecs

    Set specs = SimpleGraphSpecs(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                 SpatioTemporalRow("4", " Yes "))
    specs.CreateSeries

    Assert.AreEqual 4&, specs.NumberOfSeries, _
                    "A cell holding "" Yes "" builds the same four series as ""yes"""
    Assert.IsTrue (Not specs.HasCheckings), "And nothing is reported against it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGraphSettingWithSpacesStillBuilds", Err.Number, Err.Description
End Sub

'@sub-title Verify a graph setting nobody recognises is reported.
'@details
'The Select Case had no Case Else, so an unknown setting pushed no series, said
'nothing, and left AnalysisOutput to build a chart Graphs.Format then raised
'1004 on.
'@TestMethod("GraphSpecs")
Public Sub TestUnknownGraphSettingLogsAndBuildsNothing()
    CustomTestSetTitles Assert, "GraphSpecs", "TestUnknownGraphSettingLogsAndBuildsNothing"
    On Error GoTo TestFail

    Dim specs As GraphSpecs

    Set specs = SimpleGraphSpecs(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                 SpatioTemporalRow("4", "sometimes"))
    specs.CreateSeries

    Assert.AreEqual 0&, specs.NumberOfSeries, "An unknown setting builds no series"
    Assert.IsTrue specs.HasCheckings, "And it is reported"
    Assert.IsTrue (Not specs.Valid()), "Valid answers False for a build that produced nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnknownGraphSettingLogsAndBuildsNothing", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a percentage chart on a table with no percentages is reported.
'@TestMethod("GraphSpecs")
Public Sub TestPercentageChartWithoutPercentageColumnIsReported()
    CustomTestSetTitles Assert, "GraphSpecs", "TestPercentageChartWithoutPercentageColumnIsReported"
    On Error GoTo TestFail

    Dim specs As GraphSpecs

    Set specs = SimpleGraphSpecs(TABLE_BIVARIATE, BivariateHeader(), _
                                 BivariateRow("percentages"))
    specs.CreateSeries

    Assert.AreEqual 0&, specs.NumberOfSeries, "No percentage column, no series"
    Assert.IsTrue specs.HasCheckings, "And the reason is reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPercentageChartWithoutPercentageColumnIsReported", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify an out-of-range series index is refused.
'@TestMethod("GraphSpecs")
Public Sub TestSeriesIndexOutOfBoundsRaises()
    CustomTestSetTitles Assert, "GraphSpecs", "TestSeriesIndexOutOfBoundsRaises"
    On Error GoTo TestFail

    Dim specs As GraphSpecs
    Dim errNumber As Long
    Dim seriesName As String

    Set specs = SimpleGraphSpecs(TABLE_UNIVARIATE, UnivariateHeader(), _
                                 UnivariateRow("no", "yes", "no"))
    specs.CreateSeries

    On Error Resume Next
    seriesName = specs.SeriesName(2)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "Reading past the last series raises InvalidArgument"
    Assert.AreEqual vbNullString, seriesName, "And gives nothing back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSeriesIndexOutOfBoundsRaises", Err.Number, Err.Description
End Sub

'@section The checking keys
'===============================================================================

'@sub-title Verify two instances can file into one report.
'@details
'Checking.Add raises ElementShouldNotExists on a duplicate key and
'Checking.Append replays every key into the target, so a key made of a bare
'counter takes the whole generation down on the second collaborator that files
'anything. AnalysisOutput merges every table of a sheet into one report. This is
'the same fault the Graphs report records as its C2.
'@TestMethod("GraphSpecs")
Public Sub TestTwoInstancesProduceDistinctCheckingKeys()
    CustomTestSetTitles Assert, "GraphSpecs", "TestTwoInstancesProduceDistinctCheckingKeys"
    On Error GoTo TestFail

    Dim firstSpecs As GraphSpecs
    Dim secondSpecs As GraphSpecs
    Dim report As Checking
    Dim errNumber As Long

    Set firstSpecs = SimpleGraphSpecs(TABLE_BIVARIATE, BivariateHeader(), _
                                      BivariateRow("sometimes"))
    firstSpecs.CreateSeries

    Set secondSpecs = SimpleGraphSpecs(TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                                       SpatioTemporalRow("4", "sometimes"))
    secondSpecs.CreateSeries

    Assert.IsTrue firstSpecs.HasCheckings, "The bivariate table filed an entry"
    Assert.IsTrue secondSpecs.HasCheckings, "So did the spatio-temporal one"

    Set report = Checking.Create("Analysis output")

    On Error Resume Next
    report.Append firstSpecs.CheckingValues
    report.Append secondSpecs.CheckingValues
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual 0&, errNumber, _
                    "Two instances file into one report without a key collision"
    Assert.IsTrue (report.Length > 0), "And the report carries their entries"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTwoInstancesProduceDistinctCheckingKeys", _
                         Err.Number, Err.Description
End Sub
