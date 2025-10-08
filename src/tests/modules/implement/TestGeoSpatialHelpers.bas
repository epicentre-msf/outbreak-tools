Attribute VB_Name = "TestGeoSpatialHelpers"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Implement")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private FixtureWorkbook As Workbook

Private Const GEO_SCOPE_ADMIN As Byte = 0
Private Const GEO_SCOPE_HF As Byte = 1
Private Const LEVEL_ADMIN1 As Byte = 0
Private Const LEVEL_ADMIN2 As Byte = 1
Private Const LEVEL_ADMIN3 As Byte = 2

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set FixtureWorkbook = TestHelpers.NewWorkbook
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not FixtureWorkbook Is Nothing Then
        TestHelpers.DeleteWorkbook FixtureWorkbook
        Set FixtureWorkbook = Nothing
    End If
    TestHelpers.RestoreApp
End Sub

'@section Translation Cache
'===============================================================================

'@TestMethod("GeoSpatialHelpers")
Private Sub TestTranslationCacheRequiresWorksheet()
    On Error GoTo ExpectError
        GeoHeaderTranslationCache.Create Nothing
        Assert.Fail "Translation cache should reject missing worksheet"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "Cache must validate worksheet argument"
End Sub

'@TestMethod("GeoSpatialHelpers")
Private Sub TestTranslationCacheResolveFallback()
    Dim cache As GeoHeaderTranslationCache
    Dim resolved As String

    Set cache = GeoHeaderTranslationCache.Create(FixtureWorkbook.Worksheets(1))
    cache.EnsureLoaded
    resolved = cache.Resolve("adm1_name")

    Assert.AreEqual "adm1_name", resolved, _
                     "When no translation is present, the cache should return the tag"
End Sub

'@section Column Resolver
'===============================================================================

'@TestMethod("GeoSpatialHelpers")
Private Sub TestColumnResolverRequiresDependencies()
    On Error GoTo ExpectCacheMissing
        GeoColumnResolver.Create Nothing, New BetterArray
        Assert.Fail "Resolver should reject missing cache"
        Exit Sub
ExpectCacheMissing:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number

    On Error GoTo ExpectTablesMissing
        GeoColumnResolver.Create GeoHeaderTranslationCache.Create(FixtureWorkbook.Worksheets(1)), Nothing
        Assert.Fail "Resolver should reject missing table list"
        Exit Sub
ExpectTablesMissing:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number
End Sub

'@TestMethod("GeoSpatialHelpers")
Private Sub TestColumnResolverFindsColumns()
    Dim cache As GeoHeaderTranslationCache
    Dim tables As BetterArray
    Dim resolver As GeoColumnResolver
    Dim sheet As Worksheet
    Dim lo As ListObject
    Dim hfTable As ListObject

    Set sheet = FixtureWorkbook.Worksheets(1)
    sheet.Range("A1:C2").Value = Array(Array("adm1_name", "adm2_name", "value"), _
                                       Array("ADM1", "ADM2", "VAL"))
    Set lo = sheet.ListObjects.Add(xlSrcRange, sheet.Range("A1:C2"), , xlYes)
    lo.Name = "T_ADM1"

    sheet.Range("E1:F2").Value = Array(Array("adm1_name", "hf_code"), _
                                       Array("ADM1", "HF1"))
    Set hfTable = sheet.ListObjects.Add(xlSrcRange, sheet.Range("E1:F2"), , xlYes)
    hfTable.Name = "T_HF"

    Set cache = GeoHeaderTranslationCache.Create(sheet)
    Set tables = New BetterArray
    tables.LowerBound = 1
    tables.Push "T_ADM1"

    Set resolver = GeoColumnResolver.Create(cache, tables)

    Assert.AreEqual 1, resolver.ResolveAdminColumn(1, "adm1_name"), _
                     "Admin resolver should locate the first column"
    Assert.AreEqual 2, resolver.ResolveAdminColumn(1, "adm2_name"), _
                     "Admin resolver should locate the second column"
    Assert.AreEqual 2, resolver.ResolveHFColumn("hf_code"), _
                     "HF resolver should locate the HF-specific column"
End Sub

'@section Filter Runner
'===============================================================================

'@TestMethod("GeoSpatialHelpers")
Private Sub TestFilterRunnerRequiresListObject()
    On Error GoTo ExpectError
        GeoFilterRunner.Create Nothing
        Assert.Fail "Filter runner should reject missing listobject"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number
End Sub

'@TestMethod("GeoSpatialHelpers")
Private Sub TestFilterRunnerFiltersRows()
    Dim sheet As Worksheet
    Dim tableObject As ListObject
    Dim runner As GeoFilterRunner
    Dim visible As Range

    Set sheet = FixtureWorkbook.Worksheets(1)
    sheet.Range("A1:B3").Value = Array(Array("Country", "Value"), _
                                       Array("A", 1), _
                                       Array("B", 2))
    Set tableObject = sheet.ListObjects.Add(xlSrcRange, sheet.Range("A1:B3"), , xlYes)

    Set runner = GeoFilterRunner.Create(tableObject)
    Set visible = runner.ApplyCriteria(1, "A")

    Assert.ObjectExists visible, "Range", "ApplyCriteria should return visible rows"
    Assert.AreEqual "A", CStr(visible.Cells(1, 1).Value), "Filtered row should match criteria"
End Sub

'@section Unique Extractor
'===============================================================================

'@TestMethod("GeoSpatialHelpers")
Private Sub TestUniqueExtractorReturnsDistinctValues()
    Dim extractor As GeoUniqueExtractor
    Dim sheet As Worksheet
    Dim result As BetterArray

    Set sheet = FixtureWorkbook.Worksheets(1)
    sheet.Range("A1:A5").Value = Application.Transpose(Array("A", "B", "A", "", "C"))

    Set extractor = New GeoUniqueExtractor
    Set result = extractor.Extract(sheet.Range("A1:A5"))

    Assert.AreEqual 3, result.Length, "Extractor should return non-blank distinct values"
End Sub

'@section Spatial Scanner
'===============================================================================

'@TestMethod("GeoSpatialHelpers")
Private Sub TestSpatialScannerRequiresWorksheet()
    On Error GoTo ExpectError
        SpatialVariableScanner.Create Nothing
        Assert.Fail "Scanner should reject missing worksheet"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number
End Sub

'@TestMethod("GeoSpatialHelpers")
Private Sub TestSpatialScannerReadsVariables()
    Dim sheet As Worksheet
    Dim lo As ListObject
    Dim scanner As SpatialVariableScanner
    Dim variables As BetterArray

    Set sheet = FixtureWorkbook.Worksheets(1)
    sheet.Range("A1:A3").Value = Array("listofgeovars", "varA", "varB")
    Set lo = sheet.ListObjects.Add(xlSrcRange, sheet.Range("A1:A3"), , xlYes)
    lo.Name = "listofgeovars"

    Set scanner = SpatialVariableScanner.Create(sheet)
    Set variables = scanner.ScanVariables

    Assert.AreEqual 2, variables.Length, "Scanner should return unique variables"
End Sub

'@section Spatial Updater
'===============================================================================

'@TestMethod("GeoSpatialHelpers")
Private Sub TestSpatialUpdaterRequiresDependencies()
    On Error GoTo ExpectSheet
        SpatialTableUpdater.Create Nothing, New GeoSpatialPerformanceTracker
        Assert.Fail "Updater should reject missing worksheet"
        Exit Sub
ExpectSheet:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number

    On Error GoTo ExpectTracker
        SpatialTableUpdater.Create FixtureWorkbook.Worksheets(1), Nothing
        Assert.Fail "Updater should reject missing tracker"
        Exit Sub
ExpectTracker:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number
End Sub

'@TestMethod("GeoSpatialHelpers")
Private Sub TestSpatialUpdaterTracksDuration()
    Dim tracker As GeoSpatialPerformanceTracker
    Dim updater As SpatialTableUpdater
    Dim metadata As BetterArray

    Set tracker = New GeoSpatialPerformanceTracker
    Set updater = SpatialTableUpdater.Create(FixtureWorkbook.Worksheets(1), tracker)
    Set metadata = New BetterArray
    metadata.LowerBound = 1

    updater.ApplyUpdates metadata
Assert.IsTrue tracker.DurationMilliseconds >= 0, _
                 "Updater should record elapsed time even for no-op updates"
End Sub

'@section Geo Legacy Adapter
'===============================================================================

'@TestMethod("GeoSpatialHelpers")
Private Sub TestGeoLegacyAdapterRetrievesColumnValues()
    Dim sheet As Worksheet
    Dim tableNames As BetterArray
    Dim adapter As GeoLegacyAdapter
    Dim values As BetterArray

    Set sheet = FixtureWorkbook.Worksheets(1)
    sheet.Range("A1:B3").Value = Array(Array("adm1_name", "adm2_name"), _
                                       Array("Alpha", "One"), _
                                       Array("Beta", "Two"))
    sheet.ListObjects.Add xlSrcRange, sheet.Range("A1:B3"), , xlYes
    sheet.ListObjects(1).Name = "T_ADM1"

    Set tableNames = New BetterArray
    tableNames.LowerBound = 1
    tableNames.Push "T_ADM1"

    Set adapter = GeoLegacyAdapter.Create(sheet, tableNames)
    Set values = adapter.GetColumnValues(1, "adm1_name")

    Assert.AreEqual 2, values.Length, "Adapter should return distinct adm1 values"
    Assert.AreEqual "adm1_name", adapter.ResolveHeader("adm1_name"), _
                     "ResolveHeader should fall back to original tag when no translation is available"

    Dim filters As BetterArray
    Set filters = New BetterArray
    filters.LowerBound = 1
    filters.Push Array("adm2_name", "One")

    Set values = adapter.GetFilteredValues(1, "adm1_name", filters)
    Assert.AreEqual 1, values.Length, "Filtering should narrow down results"
    Assert.AreEqual "Alpha", CStr(values.Item(values.LowerBound)), "Filtered value should match criteria"
End Sub

'@TestMethod("GeoSpatialHelpers")
Private Sub TestGeoLevelServiceFetchesHierarchy()
    Dim sheet As Worksheet
    Dim tableNames As BetterArray
    Dim adapter As GeoLegacyAdapter
    Dim service As GeoLevelService
    Dim parents As BetterArray
    Dim values As BetterArray

    Set sheet = FixtureWorkbook.Worksheets(1)

    sheet.Range("A1:A3").Value = Application.Transpose(Array("adm1_name", "Alpha", "Beta"))
    sheet.ListObjects.Add xlSrcRange, sheet.Range("A1:A3"), , xlYes
    sheet.ListObjects(1).Name = "T_ADM1"

    sheet.Range("C1:D4").Value = Array(Array("adm1_name", "adm2_name"), _
                                        Array("Alpha", "One"), _
                                        Array("Alpha", "Two"), _
                                        Array("Beta", "Three"))
    sheet.ListObjects.Add xlSrcRange, sheet.Range("C1:D4"), , xlYes
    sheet.ListObjects(2).Name = "T_ADM2"

    sheet.Range("F1:H3").Value = Array(Array("adm1_name", "adm2_name", "hf_name"), _
                                        Array("Alpha", "One", "HF-A"), _
                                        Array("Alpha", "Two", "HF-B"))
    sheet.ListObjects.Add xlSrcRange, sheet.Range("F1:H3"), , xlYes
    sheet.ListObjects(3).Name = "T_HF"

    Set tableNames = New BetterArray
    tableNames.LowerBound = 1
    tableNames.Push "T_ADM1", "T_ADM2", "T_ADM3", "T_ADM4", "T_HF"

    Set adapter = GeoLegacyAdapter.Create(sheet, tableNames)
    Set service = GeoLevelService.Create(adapter)

    Set values = service.FetchLevels(LEVEL_ADMIN1, GEO_SCOPE_ADMIN)
    Assert.AreEqual 2, values.Length, "Level 1 should expose all adm1 values"

    Set values = service.FetchLevels(LEVEL_ADMIN2, GEO_SCOPE_ADMIN, "Alpha")
    Assert.AreEqual 2, values.Length, "Level 2 should be filtered by parent"

    Set parents = New BetterArray
    parents.LowerBound = 1
    parents.Push "Alpha", "One"
    Set values = service.FetchLevels(LEVEL_ADMIN3, GEO_SCOPE_HF, parents)
    Assert.AreEqual 1, values.Length, "HF scope should honour hierarchical parents"
    Assert.AreEqual "HF-A", values.Item(values.LowerBound), "HF result should match filtered facility"
End Sub

'@section Spatial Legacy Adapter
'===============================================================================

'@TestMethod("GeoSpatialHelpers")
Private Sub TestSpatialLegacyAdapterRefresh()
    Dim configSheet As Worksheet
    Dim spatialSheet As Worksheet
    Dim adapter As SpatialLegacyAdapter
    Dim metadata As BetterArray

    Set configSheet = FixtureWorkbook.Worksheets(1)
    configSheet.Range("A1:A4").Value = Array("listofgeovars", "geo_var_a", "geo_var_b", "geo_var_a")
    configSheet.ListObjects.Add xlSrcRange, configSheet.Range("A1:A4"), , xlYes
    configSheet.ListObjects(1).Name = "listofgeovars"

    Set spatialSheet = FixtureWorkbook.Worksheets.Add

    Set adapter = SpatialLegacyAdapter.Create(configSheet, spatialSheet)
    Set metadata = adapter.RefreshSpatialTables

    Assert.AreEqual 3, metadata.Length, "Refresh should expose unique variable names"
    Assert.IsTrue adapter.DurationMilliseconds >= 0, "Adapter should track refresh duration"
End Sub

'@section Services Facade
'===============================================================================

'@TestMethod("GeoSpatialHelpers")
Private Sub TestServicesFactory()
    Dim services As GeoSpatialServices
    Dim cache As GeoHeaderTranslationCache
    Dim tracker As GeoSpatialPerformanceTracker

    Set services = New GeoSpatialServices
    Set cache = services.BuildTranslationCache(FixtureWorkbook.Worksheets(1))
    Set tracker = services.NewPerformanceTracker

    Assert.ObjectExists cache, "GeoHeaderTranslationCache", "Services should build translation cache"
    Assert.ObjectExists tracker, "GeoSpatialPerformanceTracker", "Services should expose performance tracker"
End Sub

End Module
