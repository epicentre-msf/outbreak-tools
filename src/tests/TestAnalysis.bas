Attribute VB_Name = "TestAnalysis"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private CoreAnalysis As IAnalysis
Private Translator As ITranslationObject

'@section Helpers
'===============================================================================

Private Sub ResetAnalysis(Optional ByVal sectionValue As String = "Initial Section")
    PrepareAnalysisSheet sectionValue
    Set CoreAnalysis = Analysis.Create(ThisWorkbook.Worksheets(ANALYSISSHEET))
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ResetAnalysis
    Set Translator = CreateAnalysisTranslator()
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet ANALYSISTRANSLATIONSHEET
    DeleteWorksheet ANALYSISSHEET

    Set Translator = Nothing
    Set CoreAnalysis = Nothing
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    ResetAnalysis
    Set Translator = CreateAnalysisTranslator()
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set CoreAnalysis = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Analysis")
Private Sub TestCreateInitialisesWorksheet()
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim sut As IAnalysis

    Set hostSheet = EnsureWorksheet("AnalysisCreate")
    ClearWorksheet hostSheet
    hostSheet.Cells(1, 1).Value = "Add or remove rows of Global Summary"
    BuildAnalysisTable hostSheet, "Create Section"

    Set sut = Analysis.Create(hostSheet)
    Assert.AreEqual hostSheet.Name, sut.Wksh.Name, "Create should bind provided worksheet"

    DeleteWorksheet "AnalysisCreate"
    Exit Sub

Fail:
    DeleteWorksheet "AnalysisCreate"
    FailUnexpectedError Assert, "TestCreateInitialisesWorksheet"
End Sub

'@TestMethod("Analysis")
Private Sub TestCreateRaisesWhenWorksheetMissing()
    On Error GoTo Handler

    Dim sut As IAnalysis

    '@Ignore AssignmentNotUsed
    Set sut = Analysis.Create(Nothing)
    Assert.Fail "Create should raise when worksheet is Nothing"
    Exit Sub

Handler:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                   "Create should raise ObjectNotInitialized for missing worksheet"
End Sub

'@TestMethod("Analysis")
Private Sub TestSelfReturnsSameInstance()
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim instance As Analysis
    Dim reference As IAnalysis

    Set hostSheet = EnsureWorksheet("AnalysisSelf")
    ClearWorksheet hostSheet
    BuildAnalysisTable hostSheet, "Self Section"

    Set instance = New Analysis
    Set instance.Wksh = hostSheet
    Set reference = instance.Self

    Assert.IsTrue (instance Is reference), "Self should expose the current instance"

    DeleteWorksheet "AnalysisSelf"
    Exit Sub

Fail:
    DeleteWorksheet "AnalysisSelf"
    FailUnexpectedError Assert, "TestSelfReturnsSameInstance"
End Sub

'@TestMethod("Analysis")
Private Sub TestWkshPropertyRoundtrips()
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim instance As Analysis

    Set hostSheet = EnsureWorksheet("AnalysisWksh")
    ClearWorksheet hostSheet
    BuildAnalysisTable hostSheet, "Worksheet Section"

    Set instance = New Analysis
    Set instance.Wksh = hostSheet

    Assert.AreEqual hostSheet.Name, instance.Wksh.Name, "Wksh getter should return assigned worksheet"

    DeleteWorksheet "AnalysisWksh"
    Exit Sub

Fail:
    DeleteWorksheet "AnalysisWksh"
    FailUnexpectedError Assert, "TestWkshPropertyRoundtrips"
End Sub

'@TestMethod("Analysis")
Private Sub TestHasCheckingsReturnsNothingWhenEmpty()
    On Error GoTo Fail

    Assert.IsFalse CoreAnalysis.HasCheckings, "Fresh analysis instance should have no checkings"
    Assert.IsTrue (CoreAnalysis.CheckingValues Is Nothing), "CheckingValues should be Nothing when no logs were recorded"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestHasCheckingsReturnsNothingWhenEmpty"
End Sub

'@TestMethod("Analysis")
Private Sub TestAddAndRemoveRows()
    On Error GoTo Fail

    Dim analysisTable As ListObject
    Dim initialRows As Long

    Set analysisTable = ThisWorkbook.Worksheets(ANALYSISSHEET).ListObjects(TAB_GLOBAL_SUMMARY)
    initialRows = analysisTable.DataBodyRange.Rows.Count

    CoreAnalysis.AddRows
    Assert.IsTrue (initialRows + 5 = analysisTable.DataBodyRange.Rows.Count), "AddRows should append default number of rows"

    CoreAnalysis.RemoveRows
    Assert.IsTrue (initialRows = analysisTable.DataBodyRange.Rows.Count), "RemoveRows should trim blank rows"
    Exit Sub

Fail:
FailUnexpectedError Assert, "TestAddAndRemoveRows"
End Sub

'@TestMethod("Analysis")
Private Sub TestAddRowsTargetsSelectedTable()
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim univariateTable As ListObject
    Dim summaryTable As ListObject
    Dim initialUnivariateRows As Long
    Dim initialSummaryRows As Long

    PrepareFullAnalysisWorksheet "Add or remove rows of Univariate Analysis"
    Set hostSheet = ThisWorkbook.Worksheets(ANALYSISSHEET)
    Set CoreAnalysis = Analysis.Create(hostSheet)

    Set univariateTable = hostSheet.ListObjects(TAB_UNIVARIATE)
    Set summaryTable = hostSheet.ListObjects(TAB_GLOBAL_SUMMARY)

    initialUnivariateRows = univariateTable.DataBodyRange.Rows.Count
    initialSummaryRows = summaryTable.DataBodyRange.Rows.Count

    CoreAnalysis.AddRows

    Set univariateTable = hostSheet.ListObjects(TAB_UNIVARIATE)
    Set summaryTable = hostSheet.ListObjects(TAB_GLOBAL_SUMMARY)

    Assert.AreEqual initialUnivariateRows + 5, univariateTable.DataBodyRange.Rows.Count, _
                   "AddRows should expand the targeted table"
    Assert.AreEqual initialSummaryRows, summaryTable.DataBodyRange.Rows.Count, _
                   "AddRows should not touch unrelated tables"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddRowsTargetsSelectedTable"
End Sub

'@TestMethod("Analysis")
Private Sub TestAddRowsAssignsIdsForTimeSeries()
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim timeSeriesTable As ListObject
    Dim idRange As Range

    PrepareFullAnalysisWorksheet "Add or remove rows of Time Series Analysis"
    Set hostSheet = ThisWorkbook.Worksheets(ANALYSISSHEET)
    Set CoreAnalysis = Analysis.Create(hostSheet)

    CoreAnalysis.AddRows

    Set timeSeriesTable = hostSheet.ListObjects(TAB_TIME_SERIES)
    Set idRange = timeSeriesTable.ListColumns("Series ID").DataBodyRange

    Assert.AreEqual 6, idRange.Rows.Count, "Time series table should hold existing row plus appended defaults"
    Assert.AreEqual "Series 1", idRange.Cells(1, 1).Value, "ID generation should restart at one"
    Assert.AreEqual "Series 6", idRange.Cells(idRange.Rows.Count, 1).Value, "IDs should fill new rows sequentially"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddRowsAssignsIdsForTimeSeries"
End Sub

'@TestMethod("Analysis")
Private Sub TestRemoveRowsPreservesMinimumForSpatioTemporal()
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim spatioTemporalTable As ListObject

    PrepareFullAnalysisWorksheet "Add or remove rows of Spatio-Temporal Analysis"
    Set hostSheet = ThisWorkbook.Worksheets(ANALYSISSHEET)
    Set CoreAnalysis = Analysis.Create(hostSheet)

    CoreAnalysis.RemoveRows

    Set spatioTemporalTable = hostSheet.ListObjects(TAB_SPATIO_TEMPORAL)

    Assert.AreEqual 3, spatioTemporalTable.DataBodyRange.Rows.Count, _
                   "RemoveRows should retain populated rows and prune trailing blanks"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestRemoveRowsPreservesMinimumForSpatioTemporal"
End Sub

'@TestMethod("Analysis")
Private Sub TestAddRowsLogsMissingListObject()
    On Error GoTo Fail

    Dim logs As IChecking
    Dim keys As BetterArray
    Dim idx As Long
    Dim foundWarning As Boolean

    ResetAnalysis
    ThisWorkbook.Worksheets(ANALYSISSHEET).Cells(1, 1).Value = "Add or remove rows of Spatial Analysis"

    CoreAnalysis.AddRows

    Assert.IsTrue CoreAnalysis.HasCheckings, "Missing tables should log a warning"
    Set logs = CoreAnalysis.CheckingValues
    Set keys = logs.ListOfKeys

    For idx = keys.LowerBound To keys.UpperBound
        If InStr(1, logs.ValueOf(CStr(keys.Item(idx)), checkingLabel), TAB_SPATIAL, vbTextCompare) > 0 Then
            foundWarning = True
            Assert.AreEqual "Warning", logs.ValueOf(CStr(keys.Item(idx)), checkingType), _
                             "Missing listobject should be recorded as a warning"
            Exit For
        End If
    Next idx

    Assert.IsTrue foundWarning, "AddRows should report the missing spatial analysis table"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddRowsLogsMissingListObject"
End Sub

'@TestMethod("Analysis")
Private Sub TestImportWithNothingLogsError()
    On Error GoTo Fail

    Dim logs As IChecking
    Dim keys As BetterArray
    Dim idx As Long
    Dim foundError As Boolean

    CoreAnalysis.Import Nothing

    Assert.IsTrue CoreAnalysis.HasCheckings, "Importing nothing should log diagnostics"
    Set logs = CoreAnalysis.CheckingValues
    Set keys = logs.ListOfKeys

    For idx = keys.LowerBound To keys.UpperBound
        If InStr(1, logs.ValueOf(CStr(keys.Item(idx)), checkingLabel), "source worksheet not provided", vbTextCompare) > 0 Then
            Assert.AreEqual "Error", logs.ValueOf(CStr(keys.Item(idx)), checkingType), _
                             "Missing worksheet should be logged as an error"
            foundError = True
            Exit For
        End If
    Next idx

    Assert.IsTrue foundError, "Import Nothing should record an error entry"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestImportWithNothingLogsError"
End Sub

'@TestMethod("Analysis")
Private Sub TestImportRecordsMissingTables()
    On Error GoTo Fail

    Dim sourceSheet As Worksheet
    Dim logs As IChecking
    Dim keys As BetterArray
    Dim idx As Long
    Dim missingLogged As Boolean

    Set sourceSheet = EnsureWorksheet("AnalysisSourceMinimal")
    ClearWorksheet sourceSheet
    sourceSheet.Cells(1, 1).Value = "Add or remove rows of Global Summary"
    BuildAnalysisTable sourceSheet, "Source Section"

    CoreAnalysis.Import sourceSheet

    Assert.IsTrue CoreAnalysis.HasCheckings, "Missing tables should record diagnostics"
    Set logs = CoreAnalysis.CheckingValues
    Set keys = logs.ListOfKeys

    For idx = keys.LowerBound To keys.UpperBound
        If InStr(1, logs.ValueOf(CStr(keys.Item(idx)), checkingLabel), TAB_UNIVARIATE, vbTextCompare) > 0 Then
            missingLogged = True
            Exit For
        End If
    Next idx

    Assert.IsTrue missingLogged, "Import should warn about absent analysis tables"

    DeleteWorksheet "AnalysisSourceMinimal"
    Exit Sub

Fail:
    DeleteWorksheet "AnalysisSourceMinimal"
    FailUnexpectedError Assert, "TestImportRecordsMissingTables"
End Sub

'@TestMethod("Analysis")
Private Sub TestTranslateWithoutTranslatorRecordsWarning()
    On Error GoTo Fail

    Dim logs As IChecking
    Dim keys As BetterArray
    Dim idx As Long
    Dim foundWarning As Boolean

    CoreAnalysis.Translate Nothing

    Assert.IsTrue CoreAnalysis.HasCheckings, "Missing translator should log diagnostics"
    Set logs = CoreAnalysis.CheckingValues
    Set keys = logs.ListOfKeys

    For idx = keys.LowerBound To keys.UpperBound
        If InStr(1, logs.ValueOf(CStr(keys.Item(idx)), checkingLabel), "translation object not provided", vbTextCompare) > 0 Then
            Assert.AreEqual "Warning", logs.ValueOf(CStr(keys.Item(idx)), checkingType), _
                             "Missing translator should be reported as a warning"
            foundWarning = True
            Exit For
        End If
    Next idx

    Assert.IsTrue foundWarning, "Translate should record missing translator warnings"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestTranslateWithoutTranslatorRecordsWarning"
End Sub

'@TestMethod("Analysis")
Private Sub TestSortOrdersGraphTables()
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim graphTable As ListObject
    Dim ids As Range

    PrepareFullAnalysisWorksheet
    Set hostSheet = ThisWorkbook.Worksheets(ANALYSISSHEET)
    Set CoreAnalysis = Analysis.Create(hostSheet)

    CoreAnalysis.Sort

    Set graphTable = hostSheet.ListObjects(TAB_GRAPH_TIME_SERIES)
    Set ids = graphTable.ListColumns("Graph ID").DataBodyRange
    Assert.AreEqual "Graph 2", ids.Cells(1, 1).Value, "Sort should order Graph IDs ascending"
    Assert.AreEqual "Graph 5", ids.Cells(2, 1).Value, "Sort should preserve remaining rows"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestSortOrdersGraphTables"
End Sub

'@TestMethod("Analysis")
Private Sub TestImportCopiesValues()
    On Error GoTo Fail

    Dim sourceSheet As Worksheet
    Dim analysisTable As ListObject

    Set sourceSheet = EnsureWorksheet("AnalysisSource")
    ClearWorksheet sourceSheet
    sourceSheet.Cells(1, 1).Value = "Add or remove rows of Global Summary"
    BuildAnalysisTable sourceSheet, "Imported Section"

    CoreAnalysis.Import sourceSheet

    Set analysisTable = ThisWorkbook.Worksheets(ANALYSISSHEET).ListObjects(TAB_GLOBAL_SUMMARY)
    Assert.AreEqual "Imported Section", analysisTable.DataBodyRange.Cells(1, 1).Value, "Import should replace table values"

    DeleteWorksheet "AnalysisSource"
    Exit Sub

Fail:
    DeleteWorksheet "AnalysisSource"
    FailUnexpectedError Assert, "TestImportCopiesValues"
End Sub

'@TestMethod("Analysis")
Private Sub TestExportCreatesSheet()
    On Error GoTo Fail

    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim sourceSheet As Worksheet

    Set exportBook = NewWorkbook()
    Set sourceSheet = ThisWorkbook.Worksheets(ANALYSISSHEET)

    CoreAnalysis.Export exportBook
    Set exportedSheet = exportBook.Worksheets(sourceSheet.Name)

    Assert.AreEqual sourceSheet.Range("A3").Value, exportedSheet.Range("A3").Value, "Export should copy sheet content"
    DeleteWorkbook exportBook
    Exit Sub

Fail:
    DeleteWorkbook exportBook
    FailUnexpectedError Assert, "TestExportCreatesSheet"
End Sub

'@TestMethod("Analysis")
Private Sub TestTranslateUpdatesValues()
    On Error GoTo Fail

    Dim analysisTable As ListObject

    Set analysisTable = ThisWorkbook.Worksheets(ANALYSISSHEET).ListObjects(TAB_GLOBAL_SUMMARY)
    analysisTable.DataBodyRange.Cells(1, 1).Value = "greeting"
    analysisTable.DataBodyRange.Cells(1, 2).Value = "farewell"

    CoreAnalysis.Translate Translator

    Assert.AreEqual "Bonjour", analysisTable.DataBodyRange.Cells(1, 1).Value, "Translate should localise section values"
    Assert.AreEqual "Au revoir", analysisTable.DataBodyRange.Cells(1, 2).Value, "Translate should localise table titles"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestTranslateUpdatesValues"
End Sub
