Attribute VB_Name = "TestGraphSpecs"
Attribute VB_Description = "Tests for GraphSpecs class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for GraphSpecs class")

' GraphSpecs tests focus on factory validation and initial state.
' Full series-building tests require ICrossTable (simple mode) or real
' analysis listobjects with named ranges (complex mode), and are exercised
' through integration tests in TestAnalysisOutput.

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "GraphSpecsFixture"

Private Assert As ICustomTest

'@section Helpers
'===============================================================================

' @description Create a fixture worksheet with 3 listobjects for complex mode testing.
'              Layout matches the expected order: graph table, time series table, titles table.
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

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestGraphSpecs"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet FIXTURE_SHEET
    RestoreApp
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

'@section Factory validation tests - Create (simple mode)
'===============================================================================

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

'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsNothingLoTable()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsNothingLoTable"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    Dim lDataStub As TableSpecsLinelistStub
    Set lDataStub = New TableSpecsLinelistStub

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

'@TestMethod("GraphSpecs")
Public Sub TestCreateRangeSpecsRejectsNothingSheet()
    CustomTestSetTitles Assert, "GraphSpecs", "TestCreateRangeSpecsRejectsNothingSheet"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Set loTable = BuildComplexFixture()

    Dim lDataStub As TableSpecsLinelistStub
    Set lDataStub = New TableSpecsLinelistStub

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
    Set lDataStub = New TableSpecsLinelistStub

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

'@TestMethod("GraphSpecs")
Public Sub TestComplexModeInitialState()
    CustomTestSetTitles Assert, "GraphSpecs", "TestComplexModeInitialState"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Set loTable = BuildComplexFixture()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)

    Dim lDataStub As TableSpecsLinelistStub
    Set lDataStub = New TableSpecsLinelistStub

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

'@TestMethod("GraphSpecs")
Public Sub TestComplexModeWkshReturnsOutputSheet()
    CustomTestSetTitles Assert, "GraphSpecs", "TestComplexModeWkshReturnsOutputSheet"
    On Error GoTo TestFail

    Dim loTable As BetterArray
    Set loTable = BuildComplexFixture()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)

    Dim lDataStub As TableSpecsLinelistStub
    Set lDataStub = New TableSpecsLinelistStub

    Dim specs As IGraphSpecs
    Set specs = GraphSpecs.CreateRangeSpecs(loTable, sh, lDataStub)

    Assert.AreEqual sh.Name, specs.Wksh.Name, _
                    "Wksh should return the output sheet for complex mode"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestComplexModeWkshReturnsOutputSheet", Err.Number, Err.Description
End Sub
