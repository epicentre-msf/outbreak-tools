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
Private Const SHEET_EXPORTS As String = "Exports"
Private Const SHEET_TRANSLATIONS As String = "Translations"
Private Const COUNTER_NAME As String = "_SetupTranslationsCounter"
Private Const DICT_START_ROW As Long = 5
Private Const DICT_START_COLUMN As Long = 1
Private Const CHOICES_START_ROW As Long = 4
Private Const CHOICES_START_COLUMN As Long = 1
Private Const LIST_TS_DATA As String = "Tab_TimeSeries_Analysis"
Private Const LIST_GRAPH_TS As String = "Tab_Graph_TimeSeries"
Private Const LIST_SPATIO_TEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const LIST_SPATIO_TEMPORAL_SPECS As String = "Tab_SpatioTemporal_Specs"
Private Const TAB_TRANSLATIONS As String = "Tab_Translations"

'@ModuleInitialize
Public Sub ModuleInitialize()
    TestHelpers.BusyApp
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
    CustomTestLogFailure Assert, "TestWorkbookOpenInitialisesHiddenCounter", Err.Number, Err.Description
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
    CustomTestLogFailure Assert, "TestUpdateAnalysisDropdownsBuildsGeoLists", Err.Number, Err.Description
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
    CustomTestLogFailure Assert, "TestResetCachesRebuildsGeoDropdowns", Err.Number, Err.Description
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
    CustomTestLogFailure Assert, "TestBuildTimeSeriesHeaderUsesDictionaryLabels", Err.Number, Err.Description
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
    CustomTestLogFailure Assert, "TestAnalysisGraphValueReturnsExpectedColumn", Err.Number, Err.Description
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
    CustomTestLogFailure Assert, "TestAnalysisTimeSeriesValueReturnsExpectedColumn", Err.Number, Err.Description
End Sub

'@TestMethod("EventSetup")
Public Sub TestSpatioTemporalSpecValueReturnsExpectedColumn()
    CustomTestSetTitles Assert, "EventSetup", "Retrieves values from spatio-temporal specs"

    Dim result As String
    result = Subject.SpatioTemporalSpecValue("Section A", "N geo max")

    Assert.AreEqual "5", result, "Spatio-temporal spec lookup should use cached table"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSpatioTemporalSpecValueReturnsExpectedColumn", Err.Number, Err.Description
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
    CustomTestLogFailure Assert, "TestResetTranslationCounterMethod", Err.Number, Err.Description
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
    CustomTestLogFailure Assert, "TestRecalculateAnalysisEvaluatesFormulas", Err.Number, Err.Description
End Sub

'@TestMethod("EventSetup")
Public Sub TestManageRowsAddsDictionaryRow()
    CustomTestSetTitles Assert, "EventSetup", "ManageRows adds dictionary rows"
    On Error GoTo Fail

    Dim dictTable As ListObject
    Set dictTable = FixtureWorkbook.Worksheets(SHEET_DICTIONARY).ListObjects("Tab_Dictionary")
    Dim initialCount As Long
    initialCount = dictTable.ListRows.Count

    Subject.ManageRows SHEET_DICTIONARY

    Assert.AreEqual initialCount + 5, dictTable.ListRows.Count, "ManageRows should append a new dictionary row when del is False"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestManageRowsAddsDictionaryRow", Err.Number, Err.Description
End Sub

'@TestMethod("EventSetup")
Public Sub TestManageRowsDeletesDictionaryRow()
    CustomTestSetTitles Assert, "EventSetup", "ManageRows deletes dictionary rows"
    On Error GoTo Fail

    Dim dictTable As ListObject
    Set dictTable = FixtureWorkbook.Worksheets(SHEET_DICTIONARY).ListObjects("Tab_Dictionary")
    Dim baseline As Long
    baseline = dictTable.ListRows.Count

    Subject.ManageRows SHEET_DICTIONARY
    Assert.AreEqual baseline + 5, dictTable.ListRows.Count, "Precondition failed: dictionary row was not added"

    Subject.ManageRows SHEET_DICTIONARY, True

    Assert.AreEqual baseline, dictTable.ListRows.Count, "ManageRows with del True should remove the previously added dictionary row"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestManageRowsDeletesDictionaryRow", Err.Number, Err.Description
End Sub

'@TestMethod("EventSetup")
Public Sub TestManageRowsAddsAnalysisRows()
    CustomTestSetTitles Assert, "EventSetup", "ManageRows adds rows in analysis tables"
    On Error GoTo Fail

    Dim analysisSheet As Worksheet
    Dim tsTable As ListObject
    Dim initialCount As Long

    Set analysisSheet = FixtureWorkbook.Worksheets(SHEET_ANALYSIS)
    analysisSheet.Cells(1, 1).Value = "Add or remove rows of time series analysis"

    Set tsTable = analysisSheet.ListObjects(LIST_TS_DATA)
    initialCount = tsTable.ListRows.Count

    Subject.ManageRows SHEET_ANALYSIS

    Assert.IsTrue tsTable.ListRows.Count > initialCount, "ManageRows should append rows to the selected analysis table"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestManageRowsAddsAnalysisRows", Err.Number, Err.Description
End Sub

'@TestMethod("EventSetup")
Public Sub TestInsertRowsAddsDictionaryEntries()
    CustomTestSetTitles Assert, "EventSetup", "InsertRows mirrors dictionary selections"
    On Error GoTo Fail

    Dim dictSheet As Worksheet
    Dim dictTable As ListObject
    Dim selectionRange As Range
    Dim baseline As Long

    Set dictSheet = FixtureWorkbook.Worksheets(SHEET_DICTIONARY)
    Set dictTable = dictSheet.ListObjects("Tab_Dictionary")

    baseline = dictTable.ListRows.Count

    Set selectionRange = dictTable.ListRows(2).Range
    Set selectionRange = selectionRange.Resize(2, dictTable.ListColumns.Count)

    Subject.InsertRows SHEET_DICTIONARY, selectionRange

    Assert.AreEqual baseline + 2, dictTable.ListRows.Count, _
        "InsertRows should add as many entries as selected rows"
    Assert.AreEqual vbNullString, CStr(dictTable.ListRows(2).Range.Cells(1, 1).Value), _
        "Inserted rows should start blank"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsAddsDictionaryEntries", Err.Number, Err.Description
End Sub

'@TestMethod("EventSetup")
Public Sub TestInsertRowsShiftsStackedAnalysisTables()
    CustomTestSetTitles Assert, "EventSetup", "InsertRows shifts stacked analysis tables"
    On Error GoTo Fail

    Dim analysisSheet As Worksheet
    Dim tsTable As ListObject
    Dim graphTable As ListObject
    Dim selectionRange As Range
    Dim baseline As Long
    Dim originalGraphHeader As Long

    Set analysisSheet = FixtureWorkbook.Worksheets(SHEET_ANALYSIS)
    Set tsTable = analysisSheet.ListObjects(LIST_TS_DATA)
    Set graphTable = analysisSheet.ListObjects(LIST_GRAPH_TS)

    baseline = tsTable.ListRows.Count
    originalGraphHeader = graphTable.HeaderRowRange.Row

    Set selectionRange = tsTable.ListRows(1).Range

    Subject.InsertRows SHEET_ANALYSIS, selectionRange

    Assert.AreEqual baseline + 1, tsTable.ListRows.Count, _
        "InsertRows should append a row to the selected analysis table"
    Assert.AreEqual originalGraphHeader + 1, graphTable.HeaderRowRange.Row, _
        "Worksheet insertion should shift stacked tables down"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsShiftsStackedAnalysisTables", Err.Number, Err.Description
End Sub

'@TestMethod("EventSetup")
Public Sub TestInsertRowsAddsExportRowsAndSyncsDictionary()
    CustomTestSetTitles Assert, "EventSetup", "InsertRows adds export rows and syncs dictionary"
    On Error GoTo Fail

    Dim exportSheet As Worksheet
    Dim exportTable As ListObject
    Dim selectionRange As Range
    Dim baseline As Long
    Dim dictSheet As Worksheet
    Dim dictTable As ListObject

    Set exportSheet = FixtureWorkbook.Worksheets(SHEET_EXPORTS)
    Set exportTable = exportSheet.ListObjects("Tab_Exports")
    Set dictSheet = FixtureWorkbook.Worksheets(SHEET_DICTIONARY)
    Set dictTable = dictSheet.ListObjects("Tab_Dictionary")

    EnsureDictionaryExportColumn dictTable, "Export 1"
    EnsureDictionaryExportColumn dictTable, "Export 2"
    dictTable.ListColumns("Export 1").DataBodyRange.Cells(1, 1).Value = "Alpha"
    dictTable.ListColumns("Export 2").DataBodyRange.Cells(1, 1).Value = "Beta"

    Subject.ResetCaches

    baseline = exportTable.ListRows.Count
    Set selectionRange = exportTable.ListRows(1).Range

    Subject.InsertRows SHEET_EXPORTS, selectionRange

    Assert.AreEqual baseline + 1, exportTable.ListRows.Count, _
        "InsertRows should add a new export definition"
    Assert.AreEqual "Alpha", CStr(dictTable.ListColumns("Export 1").DataBodyRange.Cells(1, 1).Value), _
        "Export 1 column should keep its original data"
    Assert.AreEqual "Beta", CStr(dictTable.ListColumns("Export 2").DataBodyRange.Cells(1, 1).Value), _
        "Existing export columns should remain untouched"
    Assert.AreEqual vbNullString, CStr(dictTable.ListColumns("Export 3").DataBodyRange.Cells(1, 1).Value), _
        "New export column should be blank after insertion"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsAddsExportRowsAndSyncsDictionary", Err.Number, Err.Description
End Sub

'@TestMethod("EventSetup")
Public Sub TestSortTablesOrdersAnalysisGraphRows()
    CustomTestSetTitles Assert, "EventSetup", "SortTables reorders analysis graph entries"
    On Error GoTo Fail

    Dim analysisSheet As Worksheet
    Dim graphTable As ListObject
    Dim orderColumn As Range

    Set analysisSheet = FixtureWorkbook.Worksheets(SHEET_ANALYSIS)
    Set graphTable = analysisSheet.ListObjects(LIST_GRAPH_TS)
    
    graphTable.ListRows.Add
    graphTable.ListRows.Add
    graphTable.ListRows.Add

    With graphTable.ListRows(2).Range
        .Cells(1, 1).Value = "Series B"
        .Cells(1, 2).Value = "column_choice"
        .Cells(1, 3).Value = "GRAPH_1"
        .Cells(1, 4).Value = 8
        .Cells(1, 5).Value = "test"
    End With


    With graphTable.ListRows(3).Range
        .Cells(2, 1).Value = "Series C"
        .Cells(2, 2).Value = "column_choice"
        .Cells(2, 3).Value = "GRAPH_5"
        .Cells(2, 4).Value = 1
        .Cells(2, 5).Value = "test"
    End With


    Subject.SortTables SHEET_ANALYSIS

    Set orderColumn = graphTable.ListColumns("Graph order").DataBodyRange

    Assert.AreEqual CLng(1), CLng(orderColumn.Cells(1, 1).Value), "Graph order should be sorted ascending after SortTables"
    Assert.AreEqual CLng(8), CLng(orderColumn.Cells(2, 1).Value), "Existing rows should follow ascending Graph order"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSortTablesOrdersAnalysisGraphRows", Err.Number, Err.Description
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
    PrepareExportsSheet
    PrepareTranslationsSheet
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

    headers = Array("Variable Name", "Main Label", "Control", "Control Details", "Sheet Type", "Sheet Name", "main section", "Variable Type")
    rows = Array( _
        Array("geo_var", "Geo Label", "geo", "geo_list", "hlist2D", "SheetGeo", "Section Geo", "text"), _
        Array("hf_var", "HF Label", "hf", "hf_list", "hlist2D", "SheetHF", "Section HF", "text"), _
        Array("time_var", "Time Label", "manual", "time_list", "hlist2D", "SheetTime", "Section Time", "date"), _
        Array("column_choice", "Choice Column", "choice_manual", "choice_list", "hlist2D", "SheetChoice", "Section Choice", "text") _
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

Private Sub EnsureDictionaryExportColumn(ByVal dictTable As ListObject, ByVal columnName As String)
    If DictionaryColumnExists(dictTable, columnName) Then Exit Sub
    dictTable.ListColumns.Add.Name = columnName
End Sub

Private Function DictionaryColumnExists(ByVal dictTable As ListObject, ByVal columnName As String) As Boolean
    On Error Resume Next
        DictionaryColumnExists = Not dictTable.ListColumns(columnName) Is Nothing
    On Error GoTo 0
End Function

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
    analysis.Cells(1, 1).Value = "Add or remove rows of all tables"

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
    WriteMatrix analysis.Range("A6"), RowsToMatrix(Array(Array("series title", "column", "Graph ID", "Graph order", "choice", "values or percentages")))
    WriteMatrix analysis.Range("A7"), RowsToMatrix(Array(Array("Series A", "column_choice", "GRAPH_5", 5, vbNullString, vbNullString)))
    Set graphRange = analysis.Range("A6:F7")
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

Private Sub PrepareExportsSheet()
    Dim exportsSheet As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim tableRange As Range
    Dim lo As ListObject

    headerMatrix = RowsToMatrix(Array(Array("export number", "export name")))
    dataMatrix = RowsToMatrix(Array( _
        Array(1, "Export Alpha"), _
        Array(2, "Export Beta")))

    Set exportsSheet = TestHelpers.EnsureWorksheet(SHEET_EXPORTS, FixtureWorkbook)
    exportsSheet.Cells.Clear

    WriteMatrix exportsSheet.Cells(4, 1), headerMatrix
    WriteMatrix exportsSheet.Cells(5, 1), dataMatrix

    Set tableRange = exportsSheet.Range("A4").Resize(UBound(dataMatrix, 1) - LBound(dataMatrix, 1) + 2, UBound(headerMatrix, 2))
    On Error Resume Next
        exportsSheet.ListObjects(1).Delete
    On Error GoTo 0
    Set lo = exportsSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = "Tab_Exports"
End Sub

Private Sub PrepareTranslationsSheet()
    Dim translationsSheet As Worksheet
    Dim tableRange As Range
    Dim lo As ListObject

    Set translationsSheet = TestHelpers.EnsureWorksheet(SHEET_TRANSLATIONS, FixtureWorkbook)
    translationsSheet.Cells.Clear

    translationsSheet.Cells(1, 1).Value = "TranslationTag"
    translationsSheet.Cells(1, 2).Value = "English"
    translationsSheet.Cells(2, 1).Value = "RNG_Demo__1"
    translationsSheet.Cells(2, 2).Value = "Hello"

    Set tableRange = translationsSheet.Range("A1:B2")
    On Error Resume Next
        translationsSheet.ListObjects(TAB_TRANSLATIONS).Delete
    On Error GoTo 0
    Set lo = translationsSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = TAB_TRANSLATIONS
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

