Attribute VB_Name = "TestTimeSeriesGraphs"
Attribute VB_Description = "Tests for TimeSeriesGraphs class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for TimeSeriesGraphs class")

'@description
'Drives TimeSeriesGraphs, which reads the three time series graph setup tables
'and answers one entry per graph identifier: its series and its display title.
'It came out of GraphSpecs, where this half shared a type with the cross-table
'half and handed its answer over as a seven-element positional BetterArray that
'the caller destructured by index.
'
'THE FIXTURES ARE REAL LISTOBJECTS WITH THE REAL HEADERS
'-------------------------------------------------------------------------------
'The class recognises its three tables by name, so the fixture names them
'Tab_Graph_TimeSeries, Tab_TimeSeries_Analysis and Tab_Label_TSGraph. The
'headers are the ones measured on 2026-07-31 in .mock/setup_mock.xlsb,
'src/bin/setup/setup.xlsb and releases/main/setup/setup_main-2026-06-11.xlsb,
'which all agree. The class finds most of its columns by a partial match --
'"axis" against "Y-Axis", "type" against "Chart type" -- so a fixture with tidy
'invented headers would exercise none of that.
'
'THE LINELIST SPECIFICATIONS ARE A BARE INSTANCE WITH A DICTIONARY IN IT
'-------------------------------------------------------------------------------
'The class reads exactly one thing off LinelistSpecs: Dictionary(), which it
'hands to TableSpecs.Create. TestAssignDictionary fills that field, so the suite
'needs no designer workbook and none of the seven sheets LinelistSpecs.Create
'asks for.
'@depends TimeSeriesGraphs, SeriesBuffer, TableSpecs, LinelistSpecs, LLdictionary, Checking, BetterArray, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "TSGraphsFixture"
Private Const OUTPUT_SHEET As String = "TSGraphsOutput"
Private Const DICT_SHEET As String = "TSGraphsDict"

' The three setup tables of the time series graph block, named the way the
' setup workbook names them.
Private Const TABLE_GRAPH_TIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TABLE_TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const TABLE_LABEL_TSGRAPH As String = "Tab_Label_TSGraph"

' Where each fixture table starts on the fixture sheet.
Private Const GRAPH_HEADER_ROW As Long = 2
Private Const TIMESERIES_HEADER_ROW As Long = 8
Private Const TITLES_HEADER_ROW As Long = 14

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

'@section Fixture helpers
'===============================================================================

'@sub-title Free a ListObject name held anywhere in the workbook.
'@details
'A ListObject name is unique across the workbook, and other suites build
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

'@sub-title Build the three time series setup tables.
'@details
'One graph row, one time series row and one title row, so the identifiers are
'known: the time series table id is TS_tab1 and it starts its own section.
'@param graphId String. The "Graph ID" cell of the graph row.
'@param choiceValue String. The "Choices" cell of the graph row.
'@param percValue String. The "Plot values or percentages" cell.
'@param seriesId String. The "Series ID" cell of the graph row.
'@param titleGraphId String. The "Graph ID" cell of the title row.
'@return BetterArray. The three ListObjects, in the order Create wants.
Private Function BuildFixture(ByVal graphId As String, _
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

    Set BuildFixture = loTable
End Function

'@sub-title The fixture the resolution tests share.
'@param choiceValue String. The "Choices" cell of the graph row.
'@param percValue String. The "Plot values or percentages" cell.
'@return BetterArray. The three ListObjects.
Private Function StandardFixture(ByVal choiceValue As String, _
                                 ByVal percValue As String) As BetterArray
    Set StandardFixture = BuildFixture(GRAPH_ID, choiceValue, percValue, _
                                       SERIES_ID, GRAPH_ID)
End Function

'@sub-title Build the output worksheet the time series charts are drawn on.
'@details
'When the categories are wanted, the sheet carries a COLUMN_CATEGORIES_ range
'over two headed cells, each with a LABEL_COL_ name of its own. That is the
'shape CrossTable leaves behind, and it is what TimeSeriesColumnName walks.
'@param withCategories Boolean. False leaves the sheet with no named range,
'which is the state a table skipped during the table pass leaves.
'@return Worksheet. The output worksheet.
Private Function BuildOutputSheet(ByVal withCategories As Boolean) As Worksheet
    Dim sh As Worksheet

    Set sh = ResetFixtureSheet(OUTPUT_SHEET)

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
'@return LinelistSpecs. The collaborator Create requires.
Private Function LinelistData() As LinelistSpecs
    Dim stub As LinelistSpecs

    Set stub = New LinelistSpecs
    stub.TestAssignDictionary Dict
    Set LinelistData = stub
End Function

'@sub-title Build the graphs of a fixture, with categories on the output sheet.
'@param choiceValue String. The "Choices" cell of the graph row.
'@param percValue String. The "Plot values or percentages" cell.
'@return TimeSeriesGraphs. The builder, before Count is read.
Private Function StandardGraphs(ByVal choiceValue As String, _
                                ByVal percValue As String) As TimeSeriesGraphs
    Dim loTable As BetterArray
    Dim outSh As Worksheet

    Set loTable = StandardFixture(choiceValue, percValue)
    Set outSh = BuildOutputSheet(withCategories:=True)

    Set StandardGraphs = TimeSeriesGraphs.Create(loTable, outSh, LinelistData())
End Function

'@sub-title The first series name of the first graph, or an empty string.
'@param graphs TimeSeriesGraphs. The builder to read.
'@return String. The data range name of series 1 of graph 1.
Private Function FirstSeriesName(ByVal graphs As TimeSeriesGraphs) As String
    If graphs.Count = 0 Then Exit Function
    If graphs.Series(1).Count = 0 Then Exit Function
    FirstSeriesName = graphs.Series(1).RangeName(1)
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
    Assert.SetModuleName "TestTimeSeriesGraphs"

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

    DeleteWorksheets FIXTURE_SHEET, OUTPUT_SHEET, DICT_SHEET

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

'@sub-title Verify Create refuses a Nothing ListObject collection.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestCreateRejectsNothingLoTable()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestCreateRejectsNothingLoTable"
    On Error GoTo TestFail

    Dim graphs As TimeSeriesGraphs
    Dim outSh As Worksheet

    Set outSh = BuildOutputSheet(withCategories:=False)

    On Error Resume Next
    Set graphs = TimeSeriesGraphs.Create(Nothing, outSh, LinelistData())
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (graphs Is Nothing), "Create refuses a Nothing collection"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingLoTable", Err.Number, Err.Description
End Sub

'@sub-title Verify Create refuses a Nothing output worksheet.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestCreateRejectsNothingSheet()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestCreateRejectsNothingSheet"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim graphs As TimeSeriesGraphs

    Set loTable = StandardFixture(CATEGORY_CHOICE, "values")

    On Error Resume Next
    Set graphs = TimeSeriesGraphs.Create(loTable, Nothing, LinelistData())
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (graphs Is Nothing), "Create refuses a Nothing worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Create refuses Nothing linelist specifications.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestCreateRejectsNothingLData()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestCreateRejectsNothingLData"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim graphs As TimeSeriesGraphs

    Set loTable = StandardFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=True)

    On Error Resume Next
    Set graphs = TimeSeriesGraphs.Create(loTable, outSh, Nothing)
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (graphs Is Nothing), "Create refuses Nothing linelist specifications"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingLData", Err.Number, Err.Description
End Sub

'@sub-title Verify Create refuses a collection of the wrong size.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestCreateRejectsWrongCount()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestCreateRejectsWrongCount"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim shortList As BetterArray
    Dim outSh As Worksheet
    Dim graphs As TimeSeriesGraphs
    Dim errNumber As Long

    Set loTable = StandardFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=True)

    Set shortList = New BetterArray
    shortList.LowerBound = 1
    shortList.Push loTable.Item(1)

    On Error Resume Next
    Set graphs = TimeSeriesGraphs.Create(shortList, outSh, LinelistData())
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (graphs Is Nothing), "Create refuses one ListObject"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A collection of the wrong size raises InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsWrongCount", Err.Number, Err.Description
End Sub

'@sub-title Verify the three tables are recognised by their names.
'@details
'The order used to be read off a caption two rows above each header row. That
'gap is two rows in a release setup file and four rows in every development one,
'so the check refused a whole family of workbooks. The names are identical in
'all six files that were measured.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestTheSetupTableNamesAreWhatIdentifiesEachTable()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestTheSetupTableNamesAreWhatIdentifiesEachTable"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim graphs As TimeSeriesGraphs

    ' The fixture writes no caption above any of the three headers.
    Set loTable = StandardFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=True)

    Set graphs = TimeSeriesGraphs.Create(loTable, outSh, LinelistData())

    Assert.IsTrue (Not graphs Is Nothing), _
                  "The three tables are accepted on their names alone"
    Assert.AreEqual outSh.Name, graphs.Wksh.Name, _
                    "Wksh answers the output worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSetupTableNamesAreWhatIdentifiesEachTable", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the three tables in the wrong order are refused.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestCreateRejectsTablesInTheWrongOrder()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestCreateRejectsTablesInTheWrongOrder"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim swapped As BetterArray
    Dim outSh As Worksheet
    Dim graphs As TimeSeriesGraphs
    Dim errNumber As Long

    Set loTable = StandardFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=True)

    Set swapped = New BetterArray
    swapped.LowerBound = 1
    swapped.Push loTable.Item(2), loTable.Item(1), loTable.Item(3)

    On Error Resume Next
    Set graphs = TimeSeriesGraphs.Create(swapped, outSh, LinelistData())
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.IsTrue (graphs Is Nothing), "The time series table cannot stand in for the graph table"
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), errNumber, _
                    "The wrong order raises ErrorUnexpectedState"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsTablesInTheWrongOrder", _
                         Err.Number, Err.Description
End Sub

'@section Resolving a series to a named range
'===============================================================================

'@sub-title Verify a category choice resolves to the VALUES_COL_ range.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestCategoryChoiceResolvesToValuesCol()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestCategoryChoiceResolvesToValuesCol"
    On Error GoTo TestFail

    Dim graphs As TimeSeriesGraphs

    Set graphs = StandardGraphs(CATEGORY_CHOICE, "values")

    Assert.AreEqual 1&, graphs.Count, "One graph identifier gives one graph"
    Assert.AreEqual 1&, graphs.Series(1).Count, "And that graph carries one series"
    Assert.AreEqual "VALUES_COL_1_" & TS_TABLE_ID, FirstSeriesName(graphs), _
                    "A category choice plots the values of its own column"
    Assert.AreEqual "A graph title", graphs.Title(1), _
                    "The graph carries the title of its row in the titles table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCategoryChoiceResolvesToValuesCol", Err.Number, Err.Description
End Sub

'@sub-title Verify a category choice asking for percentages resolves to PERC_COL_.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestCategoryPercentageResolvesToPercCol()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestCategoryPercentageResolvesToPercCol"
    On Error GoTo TestFail

    Assert.AreEqual "PERC_COL_1_" & TS_TABLE_ID, _
                    FirstSeriesName(StandardGraphs(CATEGORY_CHOICE, "percentages")), _
                    "A category asking for percentages plots its percentage column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCategoryPercentageResolvesToPercCol", Err.Number, Err.Description
End Sub

'@sub-title Verify the total choice resolves to TOTAL_COL_VALUES_.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestTotalChoiceResolvesToTotalColValues()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestTotalChoiceResolvesToTotalColValues"
    On Error GoTo TestFail

    Assert.AreEqual "TOTAL_COL_VALUES_" & TS_TABLE_ID, _
                    FirstSeriesName(StandardGraphs(TOTAL_CHOICE, "values")), _
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
'@TestMethod("TimeSeriesGraphs")
Public Sub TestTotalPercentageResolvesToTotalPercValues()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestTotalPercentageResolvesToTotalPercValues"
    On Error GoTo TestFail

    Dim seriesName As String

    seriesName = FirstSeriesName(StandardGraphs(TOTAL_CHOICE, "percentages"))

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
'@TestMethod("TimeSeriesGraphs")
Public Sub TestTotalChoiceIsMatchedWithoutRegardToCase()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestTotalChoiceIsMatchedWithoutRegardToCase"
    On Error GoTo TestFail

    Assert.AreEqual "TOTAL_COL_VALUES_" & TS_TABLE_ID, _
                    FirstSeriesName(StandardGraphs(" total ", "values")), _
                    "A cell holding "" total "" still names the total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTotalChoiceIsMatchedWithoutRegardToCase", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a series carries the label and the prefix of its setup row.
'@details
'The series of a time series graph is labelled by its prefix alone, so the
'prefix is what the chart legend reads.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestASeriesCarriesTheLabelAndThePrefixOfItsRow()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestASeriesCarriesTheLabelAndThePrefixOfItsRow"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer

    Set buffer = StandardGraphs(CATEGORY_CHOICE, "values").Series(1)

    Assert.AreEqual "line", buffer.ChartType(1), "The chart type comes off the setup row"
    Assert.AreEqual "left", buffer.AxisSide(1), "And the axis side"
    Assert.AreEqual "Series label", buffer.LabelPrefix(1), "And the display label"
    Assert.AreEqual "ROW_CATEGORIES_" & TS_TABLE_ID, buffer.CategoryRange(1), _
                    "The categories are the ones of the section the table sits in"
    Assert.AreEqual "LABEL_COL_1_" & TS_TABLE_ID, buffer.LegendRange(1), _
                    "And the legend range is the column the choice resolved to"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASeriesCarriesTheLabelAndThePrefixOfItsRow", _
                         Err.Number, Err.Description
End Sub

'@section What is skipped, and what says so
'===============================================================================

'@sub-title Verify a table with no column categories is skipped and reported.
'@details
'The table row may have been skipped during the table pass, and then
'COLUMN_CATEGORIES_ was never created. Reading it raised 1004, and
'WriteTimeSeriesGraphs has no handler around the call that triggers the build,
'so every remaining chart of the sheet went with it.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestMissingColumnCategoriesIsReportedAndSkipped()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestMissingColumnCategoriesIsReportedAndSkipped"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim graphs As TimeSeriesGraphs

    Set loTable = StandardFixture(CATEGORY_CHOICE, "values")
    Set outSh = BuildOutputSheet(withCategories:=False)
    Set graphs = TimeSeriesGraphs.Create(loTable, outSh, LinelistData())

    Assert.AreEqual 0&, graphs.Count, _
                    "A graph whose only series cannot be resolved is left out"
    Assert.IsTrue graphs.HasCheckings, "The missing column categories are reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestMissingColumnCategoriesIsReportedAndSkipped", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a series id absent from the time series table is reported.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestUnknownSeriesIdIsReportedAndSkipped()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestUnknownSeriesIdIsReportedAndSkipped"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim graphs As TimeSeriesGraphs

    Set loTable = BuildFixture(GRAPH_ID, CATEGORY_CHOICE, "values", _
                               "A series nothing defines", GRAPH_ID)
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set graphs = TimeSeriesGraphs.Create(loTable, outSh, LinelistData())

    Assert.AreEqual 0&, graphs.Count, "A graph with no resolvable series is left out"
    Assert.IsTrue graphs.HasCheckings, "The unknown series identifier is reported"
    Assert.IsTrue (Not graphs.CheckingValues Is Nothing), _
                  "The report is handed over when there is one"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnknownSeriesIdIsReportedAndSkipped", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a blank graph identifier column builds no graph at all.
'@details
'BetterArray hands out a one-element array holding Empty when it is asked for
'the items of an empty array, so cloning an empty list gave a length of 1. That
'one phantom identifier produced one phantom graph, a chart with no series, and
'a 1004 out of Graphs.Format with no handler above it. The graphs are held as
'objects now and an empty identifier is never pushed.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestABlankGraphIdColumnBuildsNoGraph()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestABlankGraphIdColumnBuildsNoGraph"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim graphs As TimeSeriesGraphs

    Set loTable = BuildFixture(vbNullString, CATEGORY_CHOICE, "values", _
                               SERIES_ID, GRAPH_ID)
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set graphs = TimeSeriesGraphs.Create(loTable, outSh, LinelistData())

    Assert.AreEqual 0&, graphs.Count, "A blank identifier column gives no graph"
    Assert.IsTrue graphs.HasCheckings, "The empty setup table is reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestABlankGraphIdColumnBuildsNoGraph", Err.Number, Err.Description
End Sub

'@sub-title Verify a graph identifier is matched in the titles table whatever its case.
'@details
'Two tables typed by two hands. A case difference used to give a chart with an
'empty title and a navigation entry reading "Go to graph: ".
'@TestMethod("TimeSeriesGraphs")
Public Sub TestGraphTitleIsMatchedWithoutRegardToCase()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestGraphTitleIsMatchedWithoutRegardToCase"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim graphs As TimeSeriesGraphs

    Set loTable = BuildFixture(GRAPH_ID, CATEGORY_CHOICE, "values", _
                               SERIES_ID, LCase$(GRAPH_ID))
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set graphs = TimeSeriesGraphs.Create(loTable, outSh, LinelistData())

    Assert.AreEqual 1&, graphs.Count, "The graph is built"
    Assert.AreEqual "A graph title", graphs.Title(1), _
                    "The title row is found when only its case differs"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGraphTitleIsMatchedWithoutRegardToCase", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a graph with no title row is drawn and reported.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestAGraphWithNoTitleRowIsReported()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestAGraphWithNoTitleRowIsReported"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Dim outSh As Worksheet
    Dim graphs As TimeSeriesGraphs

    Set loTable = BuildFixture(GRAPH_ID, CATEGORY_CHOICE, "values", _
                               SERIES_ID, "Another graph")
    Set outSh = BuildOutputSheet(withCategories:=True)
    Set graphs = TimeSeriesGraphs.Create(loTable, outSh, LinelistData())

    Assert.AreEqual 1&, graphs.Count, "The graph is still drawn"
    Assert.AreEqual vbNullString, graphs.Title(1), "And it carries no title"
    Assert.IsTrue graphs.HasCheckings, "The missing title row is reported"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAGraphWithNoTitleRowIsReported", Err.Number, Err.Description
End Sub

'@sub-title Verify reading past the last graph is refused.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestReadingPastTheLastGraphRaises()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestReadingPastTheLastGraphRaises"
    On Error GoTo TestFail

    Dim graphs As TimeSeriesGraphs
    Dim graphTitle As String
    Dim errNumber As Long

    Set graphs = StandardGraphs(CATEGORY_CHOICE, "values")

    On Error Resume Next
    graphTitle = graphs.Title(2)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "Reading past the last graph raises InvalidArgument"
    Assert.AreEqual vbNullString, graphTitle, "And gives nothing back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestReadingPastTheLastGraphRaises", Err.Number, Err.Description
End Sub

'@section The checking keys
'===============================================================================

'@sub-title Verify two builders can file into one report.
'@details
'Checking.Add raises ElementShouldNotExists on a duplicate key and
'Checking.Append replays every key into the target, so a key made of a bare
'counter takes the whole generation down on the second collaborator that files
'anything. The key names the output worksheet, and one instance serves one.
'@TestMethod("TimeSeriesGraphs")
Public Sub TestTheCheckingKeysCarryTheWorksheetName()
    CustomTestSetTitles Assert, "TimeSeriesGraphs", "TestTheCheckingKeysCarryTheWorksheetName"
    On Error GoTo TestFail

    Dim graphs As TimeSeriesGraphs
    Dim report As Checking
    Dim errNumber As Long

    Set graphs = StandardGraphs(CATEGORY_CHOICE, "values")

    ' The fixture graph resolves, and its title row is present, so nothing is
    ' filed. A graph with no title row is what files an entry.
    Set graphs = Nothing
    Set graphs = TimeSeriesGraphs.Create( _
        BuildFixture(GRAPH_ID, CATEGORY_CHOICE, "values", SERIES_ID, "Another graph"), _
        BuildOutputSheet(withCategories:=True), LinelistData())

    Assert.AreEqual 1&, graphs.Count, "The graph is built"
    Assert.IsTrue graphs.HasCheckings, "And the missing title row is filed"

    Set report = Checking.Create("Analysis output")

    On Error Resume Next
    report.Append graphs.CheckingValues
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual 0&, errNumber, "The entries go into a shared report without a collision"
    Assert.IsTrue (report.Length > 0), "And the report carries them"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheCheckingKeysCarryTheWorksheetName", _
                         Err.Number, Err.Description
End Sub
