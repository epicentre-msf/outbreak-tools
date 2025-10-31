Attribute VB_Name = "TestEventSetup"
Attribute VB_Description = "Unit tests for the EventSetup workbook service"

Option Explicit

'@Folder("CustomTests.Setup")
'@ModuleDescription("Exercises the EventSetup class to ensure workbook events remain performant and reliable")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private Subject As IEventSetup

Private Const OUTPUT_SHEET As String = "testsOutputs"
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const SHEET_CHOICES As String = "Choices"
Private Const SHEET_ANALYSIS As String = "Analysis"
Private Const SHEET_UPDATED As String = "__updated"
Private Const SHEET_DROPDOWN As String = "__variables"
Private Const SHEET_PASSWORDS As String = "__pass"
Private Const COUNTER_NAME As String = "_SetupTranslationsCounter"
Private Const DICT_START_ROW As Long = 5
Private Const DICT_START_COLUMN As Long = 1
Private Const CHOICES_START_ROW As Long = 4
Private Const CHOICES_START_COLUMN As Long = 1
Private Const LIST_TS_DATA As String = "Tab_TimeSeries_Analysis"
Private Const LIST_GRAPH_TS As String = "Tab_Graph_TimeSeries"
Private Const LIST_SPATIO_TEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const LIST_SPATIO_TEMPORAL_SPECS As String = "Tab_SpatioTemporal_Specs"

'@ModuleInitialize
Public Sub ModuleInitialize()
    TestHelpers.BusyApp
    AssertSheetSetup
    Set Assert = CustomTest.Create(ThisWorkbook, OUTPUT_SHEET)
    Assert.SetModuleName "TestEventSetup"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    TestHelpers.RestoreApp
End Sub

'@TestInitialize
Public Sub TestInitialize()
    TestHelpers.BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    BuildFixtureWorkbook
    Set Subject = EventSetup.Create(FixtureWorkbook)
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        TestHelpers.DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set Subject = Nothing
    Set FixtureWorkbook = Nothing
End Sub

'@TestMethod("EventSetup")
Public Sub TestWorkbookOpenInitialisesHiddenCounter()
    CustomTestSetTitles Assert, "EventSetup", "Initialises hidden counter on workbook open"
    On Error GoTo Fail

    Subject.OnWorkbookOpen

    Dim namesManager As IHiddenNames
    Set namesManager = HiddenNames.Create(FixtureWorkbook.Worksheets(SHEET_UPDATED))

    Assert.AreEqual CLng(0), CLng(namesManager.Value(COUNTER_NAME)), "Workbook open should reset the translations counter"
    Exit Sub

Fail:
    ReportTestFailure "TestWorkbookOpenInitialisesHiddenCounter"
End Sub

'@TestMethod("EventSetup")
Public Sub TestUpdateAnalysisDropdownsBuildsGeoLists()
    CustomTestSetTitles Assert, "EventSetup", "Builds geo lists during dropdown refresh"
    On Error GoTo Fail

    Subject.UpdateAnalysisDropdowns True

    Dim dropInspector As IDropdownLists
    Set dropInspector = DropdownLists.Create(FixtureWorkbook.Worksheets(SHEET_DROPDOWN))

    Assert.IsTrue dropInspector.Exists("__geo_vars"), "Geo dropdown must be created after refresh"
    Assert.IsTrue dropInspector.Exists("__hfonly_vars"), "HF dropdown must be created after refresh"
    Assert.IsTrue dropInspector.Exists("__geoonly_vars"), "Geo only dropdown must be created after refresh"
    Assert.IsTrue dropInspector.Values("__geo_vars").Includes("geo_var"), "Geo list should contain geo variables"
    Assert.IsTrue dropInspector.Values("__hfonly_vars").Includes("hf_var"), "HF list should contain HF variables"
    Exit Sub

Fail:
    ReportTestFailure "TestUpdateAnalysisDropdownsBuildsGeoLists"
End Sub

'@TestMethod("EventSetup")
Public Sub TestResetCachesRebuildsGeoDropdowns()
    CustomTestSetTitles Assert, "EventSetup", "Rebuilds dropdowns after cache reset"
    On Error GoTo Fail

    Subject.UpdateAnalysisDropdowns True

    AppendGeoVariable "geo_var_new"

    Subject.ResetCaches
    Subject.UpdateAnalysisDropdowns True

    Dim dropInspector As IDropdownLists
    Set dropInspector = DropdownLists.Create(FixtureWorkbook.Worksheets(SHEET_DROPDOWN))

    Assert.IsTrue dropInspector.Values("__geo_vars").Includes("geo_var_new"), "Geo dropdown should include newly added variable after cache reset"
    Exit Sub

Fail:
    ReportTestFailure "TestResetCachesRebuildsGeoDropdowns"
End Sub

'@TestMethod("EventSetup")
Public Sub TestBuildTimeSeriesHeaderUsesDictionaryLabels()
    CustomTestSetTitles Assert, "EventSetup", "Builds header from dictionary labels"
    On Error GoTo Fail

    Dim header As String
    header = Subject.BuildTimeSeriesHeader("time_var", "geo_var", "Sum")

    Assert.AreEqual "Sum " & ChrW(9472) & " Time Label " & ChrW(9472) & " Geo Label", header, _
        "Header should combine summary, time and group labels"
    Exit Sub

Fail:
    ReportTestFailure "TestBuildTimeSeriesHeaderUsesDictionaryLabels"
End Sub

'@TestMethod("EventSetup")
Public Sub TestAnalysisGraphValueReturnsExpectedColumn()
    CustomTestSetTitles Assert, "EventSetup", "Retrieves values from graph table"
    On Error GoTo Fail

    Dim result As String
    result = Subject.AnalysisGraphValue("Series A", "Graph ID")

    Assert.AreEqual "GRAPH_5", result, "Graph value should match cached table content"
    Exit Sub

Fail:
    ReportTestFailure "TestAnalysisGraphValueReturnsExpectedColumn"
End Sub

'@TestMethod("EventSetup")
Public Sub TestAnalysisTimeSeriesValueReturnsExpectedColumn()
    CustomTestSetTitles Assert, "EventSetup", "Retrieves values from time series table"
    On Error GoTo Fail

    Dim result As String
    result = Subject.AnalysisTimeSeriesValue("Series A", "Series ID")

    Assert.AreEqual "SERIES_A", result, "Time series lookup should use cached table"
    Exit Sub

Fail:
    ReportTestFailure "TestAnalysisTimeSeriesValueReturnsExpectedColumn"
End Sub

'@TestMethod("EventSetup")
Public Sub TestSpatioTemporalSpecValueReturnsExpectedColumn()
    CustomTestSetTitles Assert, "EventSetup", "Retrieves values from spatio-temporal specs"
    On Error GoTo Fail

    Dim result As String
    result = Subject.SpatioTemporalSpecValue("Section A", "N geo max")

    Assert.AreEqual "5", result, "Spatio-temporal spec lookup should use cached table"
    Exit Sub

Fail:
    ReportTestFailure "TestSpatioTemporalSpecValueReturnsExpectedColumn"
End Sub

'@TestMethod("EventSetup")
Public Sub TestResetTranslationCounterMethod()
    CustomTestSetTitles Assert, "EventSetup", "Resets translation counter on demand"
    On Error GoTo Fail

    Dim namesManager As IHiddenNames
    Set namesManager = HiddenNames.Create(FixtureWorkbook.Worksheets(SHEET_UPDATED))
    namesManager.EnsureName COUNTER_NAME, CLng(0), HiddenNameTypeLong
    namesManager.SetValue COUNTER_NAME, CLng(42)

    Subject.ResetTranslationCounter

    Assert.AreEqual CLng(0), CLng(namesManager.Value(COUNTER_NAME)), "Manual reset should restore counter to zero"
    Exit Sub

Fail:
    ReportTestFailure "TestResetTranslationCounterMethod"
End Sub

'@TestMethod("EventSetup")
Public Sub TestRecalculateAnalysisEvaluatesFormulas()
    CustomTestSetTitles Assert, "EventSetup", "RecalculateAnalysis evaluates analysis tables"
    On Error GoTo Fail

    Dim analysis As Worksheet
    Set analysis = FixtureWorkbook.Worksheets(SHEET_ANALYSIS)

    analysis.Range("B4").Formula = "=1+1"
    analysis.Range("D7").Formula = "=2+2"
    analysis.Range("C11").Formula = "=3+3"
    analysis.Range("B15").Formula = "=4+4"

    Subject.RecalculateAnalysis

    Assert.AreEqual CDbl(2), CDbl(analysis.Range("B4").Value), "Time-series table should be recalculated"
    Assert.AreEqual CDbl(4), CDbl(analysis.Range("D7").Value), "Graph table should be recalculated"
    Assert.AreEqual CDbl(6), CDbl(analysis.Range("C11").Value), "Spatio-temporal table should be recalculated"
    Assert.AreEqual CDbl(8), CDbl(analysis.Range("B15").Value), "Spatio-temporal spec table should be recalculated"
    Exit Sub

Fail:
    ReportTestFailure "TestRecalculateAnalysisEvaluatesFormulas"
End Sub

'@section Fixture builders
'===============================================================================
Private Sub BuildFixtureWorkbook()
    PrepareUpdatedSheet
    PrepareDropdownSheet
    PreparePasswordSheet
    PrepareDictionarySheet
    PrepareChoicesSheet
    PrepareAnalysisSheet
End Sub

Private Sub PrepareUpdatedSheet()
    TestHelpers.EnsureWorksheet SHEET_UPDATED, FixtureWorkbook
End Sub

Private Sub PrepareDropdownSheet()
    Dim dropSheet As Worksheet
    Set dropSheet = TestHelpers.EnsureWorksheet(SHEET_DROPDOWN, FixtureWorkbook)
    ClearWorksheetFormatting dropSheet
End Sub

Private Sub ClearWorksheetFormatting(ByVal sh As Worksheet)
    sh.Cells.Clear
End Sub

Private Sub PreparePasswordSheet()
    PasswordsTestFixture.PreparePasswordsFixture SHEET_PASSWORDS, FixtureWorkbook
End Sub

Private Sub PrepareDictionarySheet()
    Dim dictSheet As Worksheet
    Dim targetRange As Range
    Dim tableRange As Range
    Dim headers As Variant
    Dim rows As Variant
    Dim lo As ListObject

    headers = Array("variable name", "Main Label", "control", "control details", "sheet type", "Variable Type")
    rows = Array( _
        Array("geo_var", "Geo Label", "geo", "geo_list", "hlist2D", "text"), _
        Array("hf_var", "HF Label", "hf", "hf_list", "hlist2D", "text"), _
        Array("time_var", "Time Label", "manual", "time_list", "hlist2D", "date"), _
        Array("column_choice", "Choice Column", "choice_manual", "choice_list", "hlist2D", "text") _
    )

    Set dictSheet = TestHelpers.EnsureWorksheet(SHEET_DICTIONARY, FixtureWorkbook)
    dictSheet.Cells.Clear

    dictSheet.Rows("1:4").Insert

    Set targetRange = dictSheet.Cells(DICT_START_ROW, DICT_START_COLUMN)
    WriteMatrix targetRange, RowsToMatrix(Array(headers))
    WriteMatrix targetRange.Offset(1, 0), RowsToMatrix(rows)

    Set tableRange = targetRange.Resize(UBound(rows) - LBound(rows) + 2, UBound(headers) - LBound(headers) + 1)
    On Error Resume Next
        dictSheet.ListObjects("Tab_Dictionary").Delete
    On Error GoTo 0

    Set lo = dictSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = "Tab_Dictionary"
End Sub

Private Sub PrepareChoicesSheet()
    Dim choicesSheet As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim tableRange As Range
    Dim lo As ListObject

    headerMatrix = RowsToMatrix(Array(Array("list name", "ordering list", "label", "short label")))
    dataMatrix = RowsToMatrix(Array( _
        Array("choice_list", 1, "Option A", "OptA"), _
        Array("choice_list", 2, "Option B", "OptB") _
    ))

    Set choicesSheet = TestHelpers.EnsureWorksheet(SHEET_CHOICES, FixtureWorkbook)
    choicesSheet.Cells.Clear

    WriteMatrix choicesSheet.Cells(CHOICES_START_ROW, CHOICES_START_COLUMN), headerMatrix
    WriteMatrix choicesSheet.Cells(CHOICES_START_ROW + 1, CHOICES_START_COLUMN), dataMatrix

    Set tableRange = choicesSheet.Range("A4").Resize(UBound(dataMatrix, 1) - LBound(dataMatrix, 1) + 2, UBound(headerMatrix, 2))
    On Error Resume Next
        choicesSheet.ListObjects(1).Delete
    On Error GoTo 0
    Set lo = choicesSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = "Tab_Choices"
End Sub

Private Sub PrepareAnalysisSheet()
    Dim analysis As Worksheet
    Dim tsRange As Range
    Dim graphRange As Range
    Dim spatioRange As Range
    Dim lo As ListObject

    Set analysis = TestHelpers.EnsureWorksheet(SHEET_ANALYSIS, FixtureWorkbook)
    analysis.Cells.Clear

    'Time series table
    WriteMatrix analysis.Range("A3"), RowsToMatrix(Array(Array("Title", "Series ID", "summary label", "add total")))
    WriteMatrix analysis.Range("A4"), RowsToMatrix(Array(Array("Series A", "SERIES_A", "Summary A", "no")))
    Set tsRange = analysis.Range("A3:D4")
    On Error Resume Next
        analysis.ListObjects(LIST_TS_DATA).Delete
    On Error GoTo 0
    Set lo = analysis.ListObjects.Add(SourceType:=xlSrcRange, Source:=tsRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = LIST_TS_DATA

    'Graph table
    WriteMatrix analysis.Range("A6"), RowsToMatrix(Array(Array("series title", "column", "Graph ID", "choice", "values or percentages")))
    WriteMatrix analysis.Range("A7"), RowsToMatrix(Array(Array("Series A", "column_choice", "GRAPH_5", vbNullString, vbNullString)))
    Set graphRange = analysis.Range("A6:E7")
    On Error Resume Next
        analysis.ListObjects(LIST_GRAPH_TS).Delete
    On Error GoTo 0
    Set lo = analysis.ListObjects.Add(SourceType:=xlSrcRange, Source:=graphRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = LIST_GRAPH_TS

    'Spatio temporal table
    WriteMatrix analysis.Range("A10"), RowsToMatrix(Array(Array("section", "spatial type", "geo")))
    WriteMatrix analysis.Range("A11"), RowsToMatrix(Array(Array("Section A", "geo", vbNullString)))
    Set spatioRange = analysis.Range("A10:C11")
    On Error Resume Next
        analysis.ListObjects(LIST_SPATIO_TEMPORAL).Delete
    On Error GoTo 0
    Set lo = analysis.ListObjects.Add(SourceType:=xlSrcRange, Source:=spatioRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = LIST_SPATIO_TEMPORAL

    'Spatio temporal specs table
    WriteMatrix analysis.Range("A14"), RowsToMatrix(Array(Array("Section", "N geo max")))
    WriteMatrix analysis.Range("A15"), RowsToMatrix(Array(Array("Section A", "5")))
    Set spatioRange = analysis.Range("A14:B15")
    On Error Resume Next
        analysis.ListObjects(LIST_SPATIO_TEMPORAL_SPECS).Delete
    On Error GoTo 0
    Set lo = analysis.ListObjects.Add(SourceType:=xlSrcRange, Source:=spatioRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = LIST_SPATIO_TEMPORAL_SPECS
End Sub

Private Sub AppendGeoVariable(ByVal variableName As String)
    Dim dictSheet As Worksheet
    Dim dictTable As ListObject
    Dim newRow As ListRow

    Set dictSheet = FixtureWorkbook.Worksheets(SHEET_DICTIONARY)
    Set dictTable = dictSheet.ListObjects("Tab_Dictionary")

    Set newRow = dictTable.ListRows.Add
    newRow.Range.Cells(1, 1).Value = variableName
    newRow.Range.Cells(1, 2).Value = variableName & " Label"
    newRow.Range.Cells(1, 3).Value = "geo"
    newRow.Range.Cells(1, 4).Value = "geo_list"
    newRow.Range.Cells(1, 5).Value = "hlist2D"
    newRow.Range.Cells(1, 6).Value = "text"
End Sub

'@section Utilities
'===============================================================================
Private Sub AssertSheetSetup()
    If Assert Is Nothing Then Exit Sub
End Sub

Private Sub ReportTestFailure(ByVal testName As String)
    Assert.ReportFailure testName
End Sub
