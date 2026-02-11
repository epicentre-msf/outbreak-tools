Attribute VB_Name = "TestAnalysis"

Option Explicit


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@Folder("Tests")

Private Assert As ICustomTest
Private CoreAnalysis As IAnalysis
Private Translator As ITranslationObject
Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@section Helpers
'===============================================================================

Private Sub ResetAnalysis(Optional ByVal sectionValue As String = "Initial Section")
    Dim hostSheet As Worksheet
    Set hostSheet = PrepareAnalysisSheet(sectionValue)
    Set CoreAnalysis = Analysis.Create(hostSheet)
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysis"
    ResetAnalysis
    Set Translator = CreateAnalysisTranslator()
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    ClearTestAnalysisSheets

    Set Translator = Nothing
    Set CoreAnalysis = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    BusyApp
    ResetAnalysis
    Set Translator = CreateAnalysisTranslator()
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set CoreAnalysis = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Analysis")
Public Sub TestCreateInitialisesWorksheet()
    CustomTestSetTitles Assert, "Analysis", "TestCreateInitialisesWorksheet"
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim sut As IAnalysis

    Set hostSheet = EnsureWorksheet("AnalysisCreate")
    hostSheet.Cells(1, 1).Value = "Add or remove rows of Global Summary"
    BuildAnalysisTable hostSheet, "Create Section"

    Set sut = Analysis.Create(hostSheet)
    Assert.AreEqual hostSheet.Name, sut.Wksh.Name, "Create should bind provided worksheet"

    DeleteWorksheet "AnalysisCreate"
    Exit Sub

Fail:
    DeleteWorksheet "AnalysisCreate"
    CustomTestLogFailure Assert, "TestCreateInitialisesWorksheet", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestCreateRaisesWhenWorksheetMissing()
    CustomTestSetTitles Assert, "Analysis", "TestCreateRaisesWhenWorksheetMissing"
    On Error GoTo Handler

    Dim sut As IAnalysis

    '@Ignore AssignmentNotUsed
    Set sut = Analysis.Create(Nothing)
    Assert.LogFailure "Create should raise when worksheet is Nothing"
    Exit Sub

Handler:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                   "Create should raise ObjectNotInitialized for missing worksheet"
End Sub

'@TestMethod("Analysis")
Public Sub TestSelfReturnsSameInstance()
    CustomTestSetTitles Assert, "Analysis", "TestSelfReturnsSameInstance"
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim instance As Analysis
    Dim reference As IAnalysis

    Set hostSheet = EnsureWorksheet("AnalysisSelf")
    BuildAnalysisTable hostSheet, "Self Section"

    Set instance = New Analysis
    Set instance.Wksh = hostSheet
    Set reference = instance.Self()

    Assert.IsTrue (instance Is reference), "Self should expose the current instance"

    DeleteWorksheet "AnalysisSelf"
    Exit Sub

Fail:
    DeleteWorksheet "AnalysisSelf"
    CustomTestLogFailure Assert, "TestSelfReturnsSameInstance", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestWkshPropertyRoundtrips()
    CustomTestSetTitles Assert, "Analysis", "TestWkshPropertyRoundtrips"
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim anaObj As IAnalysis
    Dim anaSheet As Worksheet

    Set hostSheet = EnsureWorksheet("AnalysisWksh")
    BuildAnalysisTable hostSheet, "Worksheet Section"

    Set anaObj = Analysis.Create(hostSheet)
    Set anaSheet = anaObj.Wksh()

    Assert.IsTrue (hostSheet Is anaSheet), "Wksh getter should return assigned worksheet"
    DeleteWorksheet "AnalysisWksh"
    Exit Sub

Fail:
    DeleteWorksheet "AnalysisWksh"
    CustomTestLogFailure Assert, "TestWkshPropertyRoundtrips", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestHasCheckingsReturnsNothingWhenEmpty()
    CustomTestSetTitles Assert, "Analysis", "TestHasCheckingsReturnsNothingWhenEmpty"
    On Error GoTo Fail

    Assert.IsFalse CoreAnalysis.HasCheckings, "Fresh analysis instance should have no checkings"
    Assert.IsTrue (CoreAnalysis.CheckingValues Is Nothing), "CheckingValues should be Nothing when no logs were recorded"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestHasCheckingsReturnsNothingWhenEmpty", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestAddAndRemoveRows()
    CustomTestSetTitles Assert, "Analysis", "TestAddAndRemoveRows"
    On Error GoTo Fail

    Dim anaLo As ListObject
    Dim initialRows As Long


    Set anaLo = AnalysisTestFixture.AnalysisTable("global summary")

    initialRows = anaLo.DataBodyRange.Rows.Count

    CoreAnalysis.AddRows
    Assert.IsTrue (initialRows + 5 = anaLo.DataBodyRange.Rows.Count), "AddRows should append default number of rows"

    CoreAnalysis.RemoveRows
    Assert.IsTrue (initialRows = anaLo.DataBodyRange.Rows.Count), "RemoveRows should trim blank rows"
    Exit Sub

Fail:
CustomTestLogFailure Assert, "TestAddAndRemoveRows", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestAddRowsTargetsSelectedTable()
    CustomTestSetTitles Assert, "Analysis", "TestAddRowsTargetsSelectedTable"
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim univariateTable As ListObject
    Dim summaryTable As ListObject
    Dim initialUnivariateRows As Long
    Dim initialSummaryRows As Long

    Set hostSheet = PrepareFullAnalysisWorksheet("ADD OR REMOVE ROWS OF UNIVARIATE ANALYSIS")
    Set CoreAnalysis = Analysis.Create(hostSheet)

    Set univariateTable = AnalysisTestFixture.AnalysisTable("univariate analysis", hostSheet)
    Set summaryTable = AnalysisTestFixture.AnalysisTable("global summary", hostSheet)

    initialUnivariateRows = univariateTable.DataBodyRange.Rows.Count
    initialSummaryRows = summaryTable.DataBodyRange.Rows.Count

    CoreAnalysis.AddRows

    Assert.AreEqual initialUnivariateRows + 5, univariateTable.DataBodyRange.Rows.Count, _
                   "AddRows should expand the targeted table"
    Assert.AreEqual initialSummaryRows, summaryTable.DataBodyRange.Rows.Count, _
                   "AddRows should not touch unrelated tables"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddRowsTargetsSelectedTable", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestAddRowsAssignsIdsForTimeSeries()
    CustomTestSetTitles Assert, "Analysis", "TestAddRowsAssignsIdsForTimeSeries"
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim timeSeriesTable As ListObject
    Dim idRange As Range

    Set hostSheet = PrepareFullAnalysisWorksheet("ADD OR REMOVE ROWS OF TIME SERIES ANALYSIS")
    Set CoreAnalysis = Analysis.Create(hostSheet)

    CoreAnalysis.AddRows

    Set timeSeriesTable = AnalysisTestFixture.AnalysisTable("time series analysis", hostSheet)
    Set idRange = timeSeriesTable.ListColumns("Series ID").DataBodyRange

    Assert.AreEqual 6, idRange.Rows.Count, "Time series table should hold existing row plus appended defaults"
    Assert.AreEqual "Series 1", idRange.Cells(1, 1).Value, "ID generation should restart at one"
    Assert.AreEqual "Series 6", idRange.Cells(idRange.Rows.Count, 1).Value, "IDs should fill new rows sequentially"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddRowsAssignsIdsForTimeSeries", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestInsertRowsUsesSelectionHeight()
    CustomTestSetTitles Assert, "Analysis", "TestInsertRowsUsesSelectionHeight"
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim sut As IAnalysis
    Dim summaryTable As ListObject
    Dim univariateTable As ListObject
    Dim selectionRange As Range
    Dim initialRows As Long
    Dim originalUnivariateHeader As Long

    Set hostSheet = PrepareFullAnalysisWorksheet("Add or remove rows of Global Summary")
    Set sut = Analysis.Create(hostSheet)

    Set summaryTable = AnalysisTestFixture.AnalysisTable("global summary", hostSheet)
    Set univariateTable = AnalysisTestFixture.AnalysisTable("univariate analysis", hostSheet)

    initialRows = summaryTable.ListRows.Count
    originalUnivariateHeader = univariateTable.HeaderRowRange.Row

    Set selectionRange = summaryTable.DataBodyRange.Rows("1:2")

    sut.InsertRows selectionRange

    Assert.AreEqual initialRows + 2, summaryTable.ListRows.Count, _
        "InsertRows should add rows matching the selection height"
    Assert.AreEqual originalUnivariateHeader + 2, univariateTable.HeaderRowRange.Row, _
        "InsertRows should shift stacked tables when worksheet insertion is enabled"
    Assert.AreEqual vbNullString, CStr(summaryTable.ListRows(1).Range.Cells(1, 1).Value), _
        "Inserted rows should clear the cells at the anchor"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestInsertRowsUsesSelectionHeight", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestDeleteRowsRemovesSelectedRows()
    CustomTestSetTitles Assert, "Analysis", "TestDeleteRowsRemovesSelectedRows"
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim sut As IAnalysis
    Dim summaryTable As ListObject
    Dim selectionRange As Range
    Dim baseline As Long

    Set hostSheet = PrepareFullAnalysisWorksheet("Add or remove rows of Global Summary")
    Set sut = Analysis.Create(hostSheet)

    Set summaryTable = AnalysisTestFixture.AnalysisTable("global summary", hostSheet)
    baseline = summaryTable.ListRows.Count

    Set selectionRange = summaryTable.ListRows(1).Range
    sut.DeleteRows selectionRange

    Assert.AreEqual baseline - 1, summaryTable.ListRows.Count, _
                     "DeleteRows should remove the targeted analysis rows"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDeleteRowsRemovesSelectedRows", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestRemoveRowsPreservesMinimumForSpatioTemporal()
    CustomTestSetTitles Assert, "Analysis", "TestRemoveRowsPreservesMinimumForSpatioTemporal"
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim spatioTemporalTable As ListObject

    Set hostSheet = PrepareFullAnalysisWorksheet("ADD OR REMOVE ROWS OF SPATIO-TEMPORAL ANALYSIS")
    Set CoreAnalysis = Analysis.Create(hostSheet)
    CoreAnalysis.RemoveRows
    Set spatioTemporalTable = AnalysisTestFixture.AnalysisTable("spatio-temporal analysis", hostSheet)

    Assert.AreEqual 4, spatioTemporalTable.DataBodyRange.Rows.Count, _
                   "RemoveRows should retain populated rows and prune trailing blanks"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemoveRowsPreservesMinimumForSpatioTemporal", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestAddRowsLogsMissingListObject()

    CustomTestSetTitles Assert, "Analysis", "TestAddRowsLogsMissingListObject"

    On Error GoTo Fail

    Dim logs As IChecking
    Dim keys As BetterArray
    Dim idx As Long
    Dim foundWarning As Boolean
    Dim hostsheet As Worksheet

    ResetAnalysis
    Set hostSheet = CoreAnalysis.Wksh()

    hostSheet.Cells(1, 1).Value = "Add or remove rows of Spatial Analysis"

    CoreAnalysis.AddRows

    Assert.IsTrue CoreAnalysis.HasCheckings, "Missing tables should log a warning"
    Set logs = CoreAnalysis.CheckingValues
    Set keys = logs.ListOfKeys

    For idx = keys.LowerBound To keys.UpperBound
        If InStr(1, logs.ValueOf(CStr(keys.Item(idx)), checkingLabel), "Tab_Spatial_Analysis", vbTextCompare) > 0 Then
            foundWarning = True
            Assert.IsTrue InStr(1, logs.ValueOf(CStr(keys.Item(idx)), checkingType), "Warning"), _
                             "Missing listobject should be recorded as a warning"
            Exit For
        End If
    Next idx

    Assert.IsTrue foundWarning, "AddRows should report the missing spatial analysis table"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddRowsLogsMissingListObject", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestImportWithNothingLogsError()
    CustomTestSetTitles Assert, "Analysis", "TestImportWithNothingLogsError"

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
            Assert.IsTrue InStr(1, logs.ValueOf(CStr(keys.Item(idx)), checkingType), "Error"), _
                             "Missing worksheet should be logged as an error"
            foundError = True
            Exit For
        End If
    Next idx

    Assert.IsTrue foundError, "Import Nothing should record an error entry"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportWithNothingLogsError", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestImportRecordsMissingTables()
    CustomTestSetTitles Assert, "Analysis", "TestImportRecordsMissingTables"
    On Error GoTo Fail

    Dim sourceSheet As Worksheet
    Dim logs As IChecking
    Dim keys As BetterArray
    Dim idx As Long
    Dim missingLogged As Boolean

    Set sourceSheet = EnsureWorksheet("AnalysisSourceMinimal")
    Set CoreAnalysis = Analysis.Create(PrepareFullAnalysisWorksheet())
    sourceSheet.Cells(1, 1).Value = "Add or remove rows of Global Summary"
    BuildAnalysisTable sourceSheet, "Source Section"

    CoreAnalysis.Import sourceSheet

    Assert.IsTrue CoreAnalysis.HasCheckings, "Missing tables should record diagnostics"
    Set logs = CoreAnalysis.CheckingValues
    Set keys = logs.ListOfKeys

    For idx = keys.LowerBound To keys.UpperBound
        If InStr(1, logs.ValueOf(CStr(keys.Item(idx)), checkingLabel), "Tab_Univariate_Analysis", vbTextCompare) > 0 Then
            missingLogged = True
            Exit For
        End If
    Next

    Assert.IsTrue missingLogged, "Import should warn about absent analysis tables"

    DeleteWorksheet "AnalysisSourceMinimal"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportRecordsMissingTables", Err.Number, Err.Description
    DeleteWorksheet "AnalysisSourceMinimal"
End Sub

'@TestMethod("Analysis")
Public Sub TestTranslateWithoutTranslatorRecordsWarning()
    CustomTestSetTitles Assert, "Analysis", "TestTranslateWithoutTranslatorRecordsWarning"
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
            Assert.IsTrue InStr(1, logs.ValueOf(CStr(keys.Item(idx)), checkingType), "Warning"), _
                             "Missing translator should be reported as a warning"
            foundWarning = True
            Exit For
        End If
    Next idx

    Assert.IsTrue foundWarning, "Translate should record missing translator warnings"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTranslateWithoutTranslatorRecordsWarning", Err.Number, Err.Description
End Sub

'@TestMethod("Analysis")
Public Sub TestSortOrdersGraphTables()
    CustomTestSetTitles Assert, "Analysis", "TestSortOrdersGraphTables"
    On Error GoTo Fail

    Dim hostSheet As Worksheet
    Dim graphTable As ListObject
    Dim ids As Range
    Dim sec As Range
    
    Set hostSheet = PrepareFullAnalysisWorksheet()
    Set CoreAnalysis = Analysis.Create(hostSheet)

    CoreAnalysis.Sort

    Set graphTable = AnalysisTestFixture.AnalysisTable("spatio-temporal analysis", hostSheet)
    Set ids = graphTable.ListColumns("Section").DataBodyRange
    Assert.AreEqual "Region A", ids.Cells(1, 1).Value, "Sort should order section ascending"
    Assert.AreEqual "Region A", ids.Cells(2, 1).Value, "Sort should order by keeping section close together"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSortOrdersGraphTables", Err.Number, Err.Description
End Sub


'@TestMethod("Analysis")
Public Sub TestExportCreatesSheet()
    CustomTestSetTitles Assert, "Analysis", "TestExportCreatesSheet"
    On Error GoTo Fail

    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet
    Dim sourceSheet As Worksheet

    Set exportBook = NewWorkbook()
    Set sourcesheet = PrepareFullAnalysisWorksheet()
    Set CoreAnalysis = Analysis.Create(sourcesheet)

    CoreAnalysis.Export exportBook
    Set exportedSheet = exportBook.Worksheets(sourceSheet.Name)

    Assert.AreEqual sourceSheet.Range("A3").Value, exportedSheet.Range("A3").Value, "Export should copy sheet content"
    DeleteWorkbook exportBook
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportCreatesSheet", Err.Number, Err.Description
    DeleteWorkbook exportBook
End Sub

'@TestMethod("Analysis")
Public Sub TestTranslateUpdatesValues()
    CustomTestSetTitles Assert, "Analysis", "TestTranslateUpdatesValues"
    On Error GoTo Fail

    Dim hostLo As ListObject

    Set hostLo = AnalysisTestFixture.AnalysisTable("global summary")
    hostLo.DataBodyRange.Cells(1, 1).Value = "greeting"
    hostLo.DataBodyRange.Cells(1, 2).Value = "farewell"

    CoreAnalysis.Translate Translator

    Assert.AreEqual "Bonjour", hostLo.DataBodyRange.Cells(1, 1).Value, "Translate should localise section values"
    Assert.AreEqual "Au revoir", hostLo.DataBodyRange.Cells(1, 2).Value, "Translate should localise table titles"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTranslateUpdatesValues", Err.Number, Err.Description
End Sub
