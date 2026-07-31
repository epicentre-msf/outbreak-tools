Attribute VB_Name = "TestSeriesBuffer"
Attribute VB_Description = "Tests for SeriesBuffer class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for SeriesBuffer class")

'@description
'Drives SeriesBuffer, the value object both graph builders fill and every
'consumer walks. It needs no worksheet, no chart and no cross-table, which is
'why this suite is the cheapest one in the folder.
'
'WHAT IT REPLACED
'-------------------------------------------------------------------------------
'Six parallel BetterArrays inside the builder, written through two routines that
'a doc comment asked the caller to pair up. The tests below pin the shape that
'makes the pairing impossible to get wrong: one Add carries all six values, and
'one entry is one series.
'@depends SeriesBuffer, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As CustomTest

'@section Module lifecycle
'===============================================================================

'@sub-title Build the assertion harness.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSeriesBuffer"
End Sub

'@sub-title Print the results.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updates before each test.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run. The first assertion of each test opens the checking, which
'picks up the titles set a line above it.
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

'@section Tests
'===============================================================================

'@sub-title Verify a fresh buffer holds nothing.
'@TestMethod("SeriesBuffer")
Public Sub TestAFreshBufferHoldsNothing()
    CustomTestSetTitles Assert, "SeriesBuffer", "TestAFreshBufferHoldsNothing"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer

    Set buffer = SeriesBuffer.Create()

    Assert.IsTrue (Not buffer Is Nothing), "Create answers a buffer"
    Assert.AreEqual 0&, buffer.Count, "And it holds no series"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFreshBufferHoldsNothing", Err.Number, Err.Description
End Sub

'@sub-title Verify one Add carries all six values of one series.
'@details
'The six values used to be pushed through two routines into six parallel
'arrays, so a caller who made one call and forgot the other left a series half
'written. One Add is what makes that impossible.
'@TestMethod("SeriesBuffer")
Public Sub TestOneAddCarriesTheWholeSeries()
    CustomTestSetTitles Assert, "SeriesBuffer", "TestOneAddCarriesTheWholeSeries"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer

    Set buffer = SeriesBuffer.Create()
    buffer.Add rangeName:="VALUES_COL_1_UA_tab1", _
               chartType:="bar", _
               axisSide:="left", _
               categoryRange:="ROW_CATEGORIES_UA_tab1", _
               legendRange:="LABEL_COL_1_UA_tab1", _
               labelPrefix:="Week 12"

    Assert.AreEqual 1&, buffer.Count, "One Add gives one series"
    Assert.AreEqual "VALUES_COL_1_UA_tab1", buffer.RangeName(1), "The data range comes back"
    Assert.AreEqual "bar", buffer.ChartType(1), "So does the chart type"
    Assert.AreEqual "left", buffer.AxisSide(1), "And the axis side"
    Assert.AreEqual "ROW_CATEGORIES_UA_tab1", buffer.CategoryRange(1), "And the category range"
    Assert.AreEqual "LABEL_COL_1_UA_tab1", buffer.LegendRange(1), "And the legend range"
    Assert.AreEqual "Week 12", buffer.LabelPrefix(1), "And the display prefix"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestOneAddCarriesTheWholeSeries", Err.Number, Err.Description
End Sub

'@sub-title Verify the prefix is empty when the caller leaves it out.
'@TestMethod("SeriesBuffer")
Public Sub TestThePrefixIsEmptyWhenItIsLeftOut()
    CustomTestSetTitles Assert, "SeriesBuffer", "TestThePrefixIsEmptyWhenItIsLeftOut"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer

    Set buffer = SeriesBuffer.Create()
    buffer.Add "VALUES_COL_1_x", "bar", "left", "ROW_CATEGORIES_x", "LABEL_COL_1_x"

    Assert.AreEqual vbNullString, buffer.LabelPrefix(1), _
                    "A series added with no prefix carries an empty one"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestThePrefixIsEmptyWhenItIsLeftOut", Err.Number, Err.Description
End Sub

'@sub-title Verify the entries keep the order they were added in.
'@details
'The order is the plot order. Excel draws series 1 first and the legend reads
'top to bottom in the same order.
'@TestMethod("SeriesBuffer")
Public Sub TestTheEntriesKeepTheOrderTheyWereAddedIn()
    CustomTestSetTitles Assert, "SeriesBuffer", "TestTheEntriesKeepTheOrderTheyWereAddedIn"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer
    Dim counter As Long

    Set buffer = SeriesBuffer.Create()
    For counter = 1 To 5
        buffer.Add "VALUES_COL_" & counter & "_x", "bar", "left", _
                   "ROW_CATEGORIES_x", "LABEL_COL_" & counter & "_x"
    Next counter

    Assert.AreEqual 5&, buffer.Count, "Five adds give five series"
    Assert.AreEqual "VALUES_COL_1_x", buffer.RangeName(1), "The first added is the first read"
    Assert.AreEqual "VALUES_COL_5_x", buffer.RangeName(5), "And the last is the last"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheEntriesKeepTheOrderTheyWereAddedIn", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify two series of one buffer stay apart.
'@details
'Six parallel arrays let one series read a value another one wrote. One entry
'per series is what stops that.
'@TestMethod("SeriesBuffer")
Public Sub TestTwoSeriesOfOneBufferStayApart()
    CustomTestSetTitles Assert, "SeriesBuffer", "TestTwoSeriesOfOneBufferStayApart"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer

    Set buffer = SeriesBuffer.Create()
    buffer.Add "VALUES_COL_1_x", "bar", "left", "ROW_CATEGORIES_x", "LABEL_COL_1_x"
    buffer.Add "PERC_COL_1_x", "point", "right", "ROW_CATEGORIES_x", "PERC_LABEL_COL_x", "over"

    Assert.AreEqual 2&, buffer.Count, "Both series are held"
    Assert.AreEqual "bar", buffer.ChartType(1), "The first keeps its own chart type"
    Assert.AreEqual "point", buffer.ChartType(2), "And the second keeps its own"
    Assert.AreEqual "left", buffer.AxisSide(1), "The first keeps its own axis"
    Assert.AreEqual "right", buffer.AxisSide(2), "And the second keeps its own"
    Assert.AreEqual vbNullString, buffer.LabelPrefix(1), "The first carries no prefix"
    Assert.AreEqual "over", buffer.LabelPrefix(2), "And the second carries the one it was given"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTwoSeriesOfOneBufferStayApart", Err.Number, Err.Description
End Sub

'@sub-title Verify reading past the last series is refused.
'@TestMethod("SeriesBuffer")
Public Sub TestReadingPastTheLastSeriesRaises()
    CustomTestSetTitles Assert, "SeriesBuffer", "TestReadingPastTheLastSeriesRaises"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer
    Dim rangeName As String
    Dim errNumber As Long

    Set buffer = SeriesBuffer.Create()
    buffer.Add "VALUES_COL_1_x", "bar", "left", "ROW_CATEGORIES_x", "LABEL_COL_1_x"

    On Error Resume Next
    rangeName = buffer.RangeName(2)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "Reading past the last series raises InvalidArgument"
    Assert.AreEqual vbNullString, rangeName, "And gives nothing back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestReadingPastTheLastSeriesRaises", Err.Number, Err.Description
End Sub

'@sub-title Verify reading an empty buffer is refused.
'@TestMethod("SeriesBuffer")
Public Sub TestReadingAnEmptyBufferRaises()
    CustomTestSetTitles Assert, "SeriesBuffer", "TestReadingAnEmptyBufferRaises"
    On Error GoTo TestFail

    Dim buffer As SeriesBuffer
    Dim chartType As String
    Dim errNumber As Long

    Set buffer = SeriesBuffer.Create()

    On Error Resume Next
    chartType = buffer.ChartType(1)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "Reading the first series of an empty buffer raises InvalidArgument"
    Assert.AreEqual vbNullString, chartType, "And gives nothing back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestReadingAnEmptyBufferRaises", Err.Number, Err.Description
End Sub

'@sub-title Verify two buffers hold their own series.
'@details
'TimeSeriesGraphs makes one buffer per graph and keeps them all, so a buffer
'that shared state with the next one would give every chart the same series.
'@TestMethod("SeriesBuffer")
Public Sub TestTwoBuffersHoldTheirOwnSeries()
    CustomTestSetTitles Assert, "SeriesBuffer", "TestTwoBuffersHoldTheirOwnSeries"
    On Error GoTo TestFail

    Dim firstBuffer As SeriesBuffer
    Dim secondBuffer As SeriesBuffer

    Set firstBuffer = SeriesBuffer.Create()
    firstBuffer.Add "VALUES_COL_1_a", "bar", "left", "ROW_CATEGORIES_a", "LABEL_COL_1_a"

    Set secondBuffer = SeriesBuffer.Create()
    secondBuffer.Add "VALUES_COL_1_b", "line", "right", "ROW_CATEGORIES_b", "LABEL_COL_1_b"
    secondBuffer.Add "VALUES_COL_2_b", "line", "right", "ROW_CATEGORIES_b", "LABEL_COL_2_b"

    Assert.AreEqual 1&, firstBuffer.Count, "The first buffer keeps its one series"
    Assert.AreEqual 2&, secondBuffer.Count, "The second keeps its two"
    Assert.AreEqual "VALUES_COL_1_a", firstBuffer.RangeName(1), _
                    "And the first still answers its own data range"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTwoBuffersHoldTheirOwnSeries", Err.Number, Err.Description
End Sub
