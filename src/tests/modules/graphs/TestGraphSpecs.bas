Attribute VB_Name = "TestGraphSpecs"
Attribute VB_Description = "Tests for GraphSpecs class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for GraphSpecs class")

'@description
'Validates the GraphSpecs class, which builds chart series specifications
'in two modes: simple (non-time-series from an ICrossTable) and complex
'(time series from analysis ListObjects with graph, time series, and title
'tables). Tests focus on factory validation and initial state since full
'series-building tests require ICrossTable (simple mode) or real analysis
'ListObjects with named ranges (complex mode), and are exercised through
'integration tests in TestAnalysisOutput. Tests verify: Create rejects
'Nothing cross-table (simple mode); CreateRangeSpecs rejects Nothing
'loTable, Nothing sheet, Nothing lData, and wrong ListObject count
'(complex mode); complex mode initial state has zero series and graphs
'before CreateSeries; Wksh returns the output sheet in complex mode.
'@depends GraphSpecs, IGraphSpecs, BetterArray, TableSpecsLinelistStub, AnalysisDictionaryStub, CustomTest, TestHelpers

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "GraphSpecsFixture"

Private Assert As ICustomTest

'@section Helpers
'===============================================================================

'@sub-title Create a minimal lData stub with a dictionary configured.
'@details
'GraphSpecs.CreateRangeSpecs stores lData and later passes
'lData.Dictionary() to TableSpecs.Create, so the dictionary must not be
'Nothing. Instantiates a TableSpecsLinelistStub and assigns a fresh
'AnalysisDictionaryStub to satisfy this requirement.
'@return TableSpecsLinelistStub. A stub with a valid dictionary reference.
Private Function CreateLDataStub() As TableSpecsLinelistStub
    Dim stub As TableSpecsLinelistStub
    Set stub = New TableSpecsLinelistStub
    stub.SetDictionary New AnalysisDictionaryStub
    Set CreateLDataStub = stub
End Function

'@sub-title Create a fixture worksheet with three ListObjects for complex mode testing.
'@details
'Builds a hidden worksheet named "GraphSpecsFixture" containing three
'ListObjects in the expected order: a graph table (tblGraphTS with columns
'graph id, series id, axis, percentages, type, choices, label), a time
'series table (tblTimeSeries with columns row, column, section, total,
'percentage, missing, graph), and a title table (tblGraphTitles with
'columns title, subtitle, graph id). Each table has one data row with
'realistic values. Returns a BetterArray (1-based) containing the three
'ListObject references.
'@return BetterArray. A 1-based array of [graphLo, tsLo, titleLo] ListObjects.
Private Function BuildComplexFixture() As BetterArray
    Dim sh As Worksheet
    Dim loTable As BetterArray
    Dim graphLo As ListObject
    Dim tsLo As ListObject
    Dim titleLo As ListObject
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    ' Graph listobject: starts at row 2 (row 1 = type label)
    sh.Cells(1, 1).Value = "graph on time series"
    headerMatrix = RowsToMatrix(Array( _
        Array("graph id", "series id", "axis", "percentages", "type", "choices", "label")))
    WriteMatrix sh.Cells(2, 1), headerMatrix
    dataMatrix = RowsToMatrix(Array( _
        Array("g1", "ts_row1", "left", "values", "bar", "choice_a", "Series A")))
    WriteMatrix sh.Cells(3, 1), dataMatrix
    Set graphLo = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                      Source:=sh.Range("A2:G3"), _
                                      XlListObjectHasHeaders:=xlYes)
    graphLo.Name = "tblGraphTS"

    ' Time series listobject: starts at row 6 (row 5 = type label)
    sh.Cells(5, 1).Value = "time series analysis"
    headerMatrix = RowsToMatrix(Array( _
        Array("row", "column", "section", "total", "percentage", "missing", "graph")))
    WriteMatrix sh.Cells(6, 1), headerMatrix
    dataMatrix = RowsToMatrix(Array( _
        Array("ts_row1", "choice_var", "S1", "yes", "no", "no", "yes")))
    WriteMatrix sh.Cells(7, 1), dataMatrix
    Set tsLo = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                   Source:=sh.Range("A6:G7"), _
                                   XlListObjectHasHeaders:=xlYes)
    tsLo.Name = "tblTimeSeries"

    ' Title listobject: starts at row 10 (row 9 = type label)
    sh.Cells(9, 1).Value = "labels for time series graphs"
    headerMatrix = RowsToMatrix(Array(Array("title", "subtitle", "graph id")))
    WriteMatrix sh.Cells(10, 1), headerMatrix
    dataMatrix = RowsToMatrix(Array(Array("Graph Title 1", "", "g1")))
    WriteMatrix sh.Cells(11, 1), dataMatrix
    Set titleLo = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                      Source:=sh.Range("A10:C11"), _
                                      XlListObjectHasHeaders:=xlYes)
    titleLo.Name = "tblGraphTitles"

    Set loTable = New BetterArray
    loTable.LowerBound = 1
    loTable.Push graphLo, tsLo, titleLo

    Set BuildComplexFixture = loTable
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Initialise the test module before any tests run.
'@details
'Suppresses screen updates via BusyApp, ensures the test output sheet
'exists, creates the CustomTest assertion object targeting that sheet,
'and sets the module name for result grouping.
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestGraphSpecs"
End Sub

'@sub-title Tear down the module after all tests complete.
'@details
'Prints accumulated test results to the output sheet, deletes the
'fixture worksheet, restores the application state via RestoreApp,
'and releases the assertion object.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet FIXTURE_SHEET
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Reset state before each individual test.
'@details
'Suppresses screen updates so worksheet operations during each test do
'not trigger flickering or event cascades.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Clean up after each individual test.
'@details
'Flushes any pending assertion results to the output sheet so each test's
'outcome is recorded before the next test begins.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory validation tests - Create (simple mode)
'===============================================================================

'@sub-title Verify Create returns Nothing when the cross-table argument is Nothing.
'@details
'Acts by calling GraphSpecs.Create with Nothing under On Error Resume
'Next. Asserts that the result is Nothing, confirming the guard clause
'rejects invalid input in simple mode without raising an unhandled error.
'@TestMethod("GraphSpecs")
Public Sub TestCreateRejectsNothingTable()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRejectsNothingTable"
    On Error GoTo TestFail

    On Error Resume Next
    Dim specs As IGraphSpecs
    Set specs = GraphSpecs.Create(Nothing)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "Create with Nothing cross-table should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingTable", Err.Number, Err.Description
End Sub

'@section Factory validation tests - CreateRangeSpecs (complex mode)
'===============================================================================

'@sub-title Verify CreateRangeSpecs returns Nothing when the loTable argument is Nothing.
'@details
'Arranges a fixture worksheet and a valid lData stub. Acts by calling
'GraphSpecs.CreateRangeSpecs with Nothing as the loTable under On Error
'Resume Next. Asserts that the result is Nothing, confirming the guard
'clause rejects a missing ListObject collection.
'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsNothingLoTable()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsNothingLoTable"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    Dim lDataStub As TableSpecsLinelistStub
    Set lDataStub = CreateLDataStub()

    On Error Resume Next
    Dim specs As IGraphSpecs
    Set specs = GraphSpecs.CreateRangeSpecs(Nothing, sh, lDataStub)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "CreateRangeSpecs with Nothing loTable should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRangeSpecsRejectsNothingLoTable", Err.Number, Err.Description
End Sub

'@sub-title Verify CreateRangeSpecs returns Nothing when the output sheet is Nothing.
'@details
'Arranges a valid complex fixture and lData stub. Acts by calling
'GraphSpecs.CreateRangeSpecs with Nothing as the sheet under On Error
'Resume Next. Asserts that the result is Nothing, confirming the guard
'clause rejects a missing output worksheet.
'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsNothingSheet()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsNothingSheet"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Set loTable = BuildComplexFixture()

    Dim lDataStub As TableSpecsLinelistStub
    Set lDataStub = CreateLDataStub()

    On Error Resume Next
    Dim specs As IGraphSpecs
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, Nothing, lDataStub)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "CreateRangeSpecs with Nothing sheet should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRangeSpecsRejectsNothingSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify CreateRangeSpecs returns Nothing when the lData argument is Nothing.
'@details
'Arranges a valid complex fixture and retrieves the fixture worksheet.
'Acts by calling GraphSpecs.CreateRangeSpecs with Nothing as lData under
'On Error Resume Next. Asserts that the result is Nothing, confirming the
'guard clause rejects a missing linelist data reference.
'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsNothingLData()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsNothingLData"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Set loTable = BuildComplexFixture()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)

    On Error Resume Next
    Dim specs As IGraphSpecs
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, sh, Nothing)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "CreateRangeSpecs with Nothing lData should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRangeSpecsRejectsNothingLData", Err.Number, Err.Description
End Sub

'@sub-title Verify CreateRangeSpecs returns Nothing when the loTable has the wrong count.
'@details
'Arranges a fixture worksheet with only one ListObject instead of the
'required three (graph, time series, title). Acts by calling
'GraphSpecs.CreateRangeSpecs with the undersized loTable under On Error
'Resume Next. Asserts that the result is Nothing, confirming the factory
'validates the ListObject count before proceeding.
'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsWrongCount()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsWrongCount"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    ' Only 1 listobject instead of 3
    sh.Cells(1, 1).Value = "graph on time series"
    WriteRow sh.Cells(2, 1), "graph id", "series id"
    WriteRow sh.Cells(3, 1), "g1", "ts1"
    Dim lo As ListObject
    Set lo = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                 Source:=sh.Range("A2:B3"), _
                                 XlListObjectHasHeaders:=xlYes)

    Dim loTable As BetterArray
    Set loTable = New BetterArray
    loTable.LowerBound = 1
    loTable.Push lo

    Dim lDataStub As TableSpecsLinelistStub
    Set lDataStub = CreateLDataStub()

    On Error Resume Next
    Dim specs As IGraphSpecs
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, sh, lDataStub)
    On Error GoTo 0

    Assert.IsTrue (specs Is Nothing), _
                  "CreateRangeSpecs with wrong listobject count should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRangeSpecsRejectsWrongCount", Err.Number, Err.Description
End Sub

'@section Initial state tests
'===============================================================================

'@sub-title Verify complex mode initial state reports zero series and zero graphs.
'@details
'Arranges a valid complex fixture with all three ListObjects. Acts by
'calling GraphSpecs.CreateRangeSpecs and reading NumberOfSeries and
'NumberOfGraphs before calling CreateSeries. Asserts that the factory
'succeeds (not Nothing) and that both counts are zero, confirming the
'class defers series population until CreateSeries is explicitly invoked.
'@TestMethod("GraphSpecs")
Public Sub TestComplexModeInitialState()
    CustomTestSetTitles Assert, "GraphSpecs", "TestComplexModeInitialState"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Set loTable = BuildComplexFixture()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)

    Dim lDataStub As TableSpecsLinelistStub
    Set lDataStub = CreateLDataStub()

    Dim specs As IGraphSpecs
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, sh, lDataStub)

    Assert.IsTrue (Not specs Is Nothing), _
                  "CreateRangeSpecs with valid fixture should succeed"

    ' Before CreateSeries is called, counts should be zero
    Assert.AreEqual 0&, specs.NumberOfSeries, _
                    "NumberOfSeries should be 0 for complex mode"
    Assert.AreEqual 0&, specs.NumberOfGraphs, _
                    "NumberOfGraphs should be 0 before CreateSeries"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestComplexModeInitialState", Err.Number, Err.Description
End Sub

'@sub-title Verify Wksh returns the output sheet in complex mode.
'@details
'Arranges a valid complex fixture. Acts by calling CreateRangeSpecs and
'reading the Wksh property. Asserts that the returned worksheet name
'matches the fixture sheet name, confirming that complex mode stores and
'exposes the output worksheet supplied during construction.
'@TestMethod("GraphSpecs")
Public Sub TestComplexModeWkshReturnsOutputSheet()
    CustomTestSetTitles Assert, "GraphSpecs", "TestComplexModeWkshReturnsOutputSheet"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Set loTable = BuildComplexFixture()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)

    Dim lDataStub As TableSpecsLinelistStub
    Set lDataStub = CreateLDataStub()

    Dim specs As IGraphSpecs
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, sh, lDataStub)

    Assert.AreEqual sh.Name, specs.Wksh.Name, _
                    "Wksh should return the output sheet for complex mode"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestComplexModeWkshReturnsOutputSheet", Err.Number, Err.Description
End Sub
