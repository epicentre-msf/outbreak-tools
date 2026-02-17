Attribute VB_Name = "TestLLGeo"
Attribute VB_Description = "Tests for LLGeo class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLGeo class")

'@description
'Validates the LLGeo class, which manages the geobase worksheet containing
'ListObjects for four admin levels (T_ADM1 through T_ADM4), health facilities
'(T_HF), translation names (T_NAMES), historic entries, and metadata. Tests
'focus on factory validation and lightweight property behaviour since full
'integration tests with geo data and AutoFilter require a real geobase
'worksheet with imported data. The fixture builds a minimal geobase worksheet
'with all required tables and named ranges but no data rows, then tears it
'down in ModuleCleanup to ensure isolation. Tests verify: factory rejects
'Nothing; factory rejects sheets missing required tables; factory succeeds
'with a full fixture; HasNoData reflects RNG_GeoUpdated state; GeoNames
'resolves from named range cache; GeoNames falls back to raw name for unknown
'keys; GeoLevel returns empty when geobase has no data.
'@depends LLGeo, ILLGeo, BetterArray, CustomTest, TestHelpers, HiddenNames, IHiddenNames

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const GEO_FIXTURE As String = "GeoFixture"

Private Assert As ICustomTest

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
    Assert.SetModuleName "TestLLGeo"
End Sub

'@sub-title Tear down the module after all tests complete.
'@details
'Prints accumulated test results to the output sheet, restores the
'application state via RestoreApp, releases the assertion object, and
'deletes all temporary worksheets created during the test run.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
    DeleteWorksheets GEO_FIXTURE, "GeoEmptyTest"
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

'@section Helpers
'===============================================================================

'@sub-title Build a minimal geobase fixture with all required tables and named ranges.
'@details
'Creates a hidden worksheet named "GeoFixture" containing nine ListObjects
'(T_ADM1 through T_ADM4, T_HF, T_NAMES, T_HISTOGEO, T_HISTOHF, T_METADATA)
'with correct header columns but no data rows. Also creates all required
'cell-based named ranges (RNG_PastingGeoCol, RNG_HFNAME, RNG_ADM1NAME
'through RNG_ADM4NAME) in row 5, and five HiddenNames (RNG_GeoUpdated,
'RNG_GeoName, RNG_GeoLangCode, RNG_FormLoaded, RNG_MetaLang) as constant
'values. Sets RNG_GeoUpdated to "empty" and RNG_GeoName to "test_geo".
'@return Worksheet. The fully prepared geobase fixture sheet.
Private Function BuildGeoFixture() As Worksheet
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim rng As Range
    Dim counter As Long
    Dim tblNames As Variant
    Dim tblCols As Variant
    Dim rngNames As Variant
    Dim startCol As Long

    Set sh = EnsureWorksheet(GEO_FIXTURE, clearSheet:=True, visibility:=xlSheetHidden)

    tblNames = Array("T_ADM1", "T_ADM2", "T_ADM3", "T_ADM4", "T_HF", _
                     "T_NAMES", "T_HISTOGEO", "T_HISTOHF", "T_METADATA")
    tblCols = Array(2, 3, 4, 5, 5, 2, 1, 1, 2)

    startCol = 1
    For counter = LBound(tblNames) To UBound(tblNames)
        Set rng = sh.Range(sh.Cells(1, startCol), _
                           sh.Cells(2, startCol + CLng(tblCols(counter)) - 1))

        Select Case CStr(tblNames(counter))
        Case "T_ADM1"
            sh.Cells(1, startCol).Value = "adm1_name"
            sh.Cells(1, startCol + 1).Value = "adm1_concat"
        Case "T_ADM2"
            sh.Cells(1, startCol).Value = "adm1_name"
            sh.Cells(1, startCol + 1).Value = "adm2_name"
            sh.Cells(1, startCol + 2).Value = "adm2_concat"
        Case "T_ADM3"
            sh.Cells(1, startCol).Value = "adm1_name"
            sh.Cells(1, startCol + 1).Value = "adm2_name"
            sh.Cells(1, startCol + 2).Value = "adm3_name"
            sh.Cells(1, startCol + 3).Value = "adm3_concat"
        Case "T_ADM4"
            sh.Cells(1, startCol).Value = "adm1_name"
            sh.Cells(1, startCol + 1).Value = "adm2_name"
            sh.Cells(1, startCol + 2).Value = "adm3_name"
            sh.Cells(1, startCol + 3).Value = "adm4_name"
            sh.Cells(1, startCol + 4).Value = "adm4_concat"
        Case "T_HF"
            sh.Cells(1, startCol).Value = "hf_name"
            sh.Cells(1, startCol + 1).Value = "hf_pcode"
            sh.Cells(1, startCol + 2).Value = "adm3_name"
            sh.Cells(1, startCol + 3).Value = "adm2_name"
            sh.Cells(1, startCol + 4).Value = "adm1_name"
        Case "T_NAMES"
            sh.Cells(1, startCol).Value = "level"
            sh.Cells(1, startCol + 1).Value = "EN"
        Case "T_HISTOGEO"
            sh.Cells(1, startCol).Value = "HistoGeo"
        Case "T_HISTOHF"
            sh.Cells(1, startCol).Value = "HistoFacility"
        Case "T_METADATA"
            sh.Cells(1, startCol).Value = "variable"
            sh.Cells(1, startCol + 1).Value = "value"
        End Select

        sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = CStr(tblNames(counter))
        startCol = startCol + CLng(tblCols(counter)) + 1
    Next counter

    'Place cell-based named ranges in row 5 (only the 6 that remain as ranges)
    rngNames = Array("RNG_PastingGeoCol", "RNG_HFNAME", "RNG_ADM1NAME", _
                     "RNG_ADM2NAME", "RNG_ADM3NAME", "RNG_ADM4NAME")

    For counter = LBound(rngNames) To UBound(rngNames)
        Set rng = sh.Cells(5, counter - LBound(rngNames) + 1)
        rng.Name = CStr(rngNames(counter))
    Next counter

    'Create the 5 HiddenNames (constant values, not cell references)
    Dim geoStore As IHiddenNames
    Set geoStore = HiddenNames.Create(sh)
    geoStore.EnsureName "RNG_GeoUpdated", "empty", HiddenNameTypeString
    geoStore.EnsureName "RNG_GeoName", "test_geo", HiddenNameTypeString
    geoStore.EnsureName "RNG_GeoLangCode", vbNullString, HiddenNameTypeString
    geoStore.EnsureName "RNG_FormLoaded", vbNullString, HiddenNameTypeString
    geoStore.EnsureName "RNG_MetaLang", vbNullString, HiddenNameTypeString

    Set BuildGeoFixture = sh
End Function

'@section Factory validation tests
'===============================================================================

'@sub-title Verify Create returns Nothing when the worksheet argument is Nothing.
'@details
'Acts by calling LLGeo.Create with Nothing under On Error Resume Next.
'Asserts that the result is Nothing, confirming the guard clause rejects
'invalid input without raising an unhandled error.
'@TestMethod("LLGeo")
Public Sub TestCreateRejectsNothing()
    CustomTestSetTitles Assert, "LLGeo", "TestCreateRejectsNothing"
    On Error GoTo TestFail

    On Error Resume Next
    Dim geo As ILLGeo
    Set geo = LLGeo.Create(Nothing)
    On Error GoTo 0

    Assert.IsTrue (geo Is Nothing), _
                  "Create with Nothing sheet should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothing", Err.Number, Err.Description
End Sub

'@sub-title Verify Create returns Nothing when the worksheet lacks required tables.
'@details
'Arranges a blank hidden worksheet named "GeoEmptyTest" with no ListObjects
'or named ranges. Acts by calling LLGeo.Create with that sheet under On
'Error Resume Next. Asserts that the result is Nothing, confirming the
'factory validates the presence of required geobase tables before returning
'an instance.
'@TestMethod("LLGeo")
Public Sub TestCreateRejectsSheetMissingTables()
    CustomTestSetTitles Assert, "LLGeo", "TestCreateRejectsSheetMissingTables"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet("GeoEmptyTest", clearSheet:=True, visibility:=xlSheetHidden)

    On Error Resume Next
    Dim geo As ILLGeo
    Set geo = LLGeo.Create(sh)
    On Error GoTo 0

    Assert.IsTrue (geo Is Nothing), _
                  "Create with sheet missing required tables should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsSheetMissingTables", Err.Number, Err.Description
End Sub

'@sub-title Verify Create succeeds with a fully populated geobase fixture.
'@details
'Arranges a complete geobase fixture via BuildGeoFixture containing all
'nine ListObjects and all required named ranges. Acts by calling
'LLGeo.Create with that sheet. Asserts that the result is not Nothing,
'confirming the factory accepts a well-formed geobase worksheet.
'@TestMethod("LLGeo")
Public Sub TestCreateSucceedsWithFullFixture()
    CustomTestSetTitles Assert, "LLGeo", "TestCreateSucceedsWithFullFixture"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()

    Dim geo As ILLGeo
    Set geo = LLGeo.Create(sh)

    Assert.IsNotNothing geo, _
                        "Create with full fixture should succeed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateSucceedsWithFullFixture", Err.Number, Err.Description
End Sub

'@section HasNoData tests
'===============================================================================

'@sub-title Verify HasNoData returns True when RNG_GeoUpdated reads "empty".
'@details
'Arranges a geobase fixture whose RNG_GeoUpdated is set to "empty" by
'default. Acts by creating an LLGeo instance and reading HasNoData.
'Asserts that the property is True, confirming the class correctly
'interprets the "empty" flag as having no imported data.
'@TestMethod("LLGeo")
Public Sub TestHasNoDataWhenEmpty()
    CustomTestSetTitles Assert, "LLGeo", "TestHasNoDataWhenEmpty"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()

    Dim geo As ILLGeo
    Set geo = LLGeo.Create(sh)

    Assert.IsTrue geo.HasNoData, _
                  "HasNoData should be True when RNG_GeoUpdated is 'empty'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasNoDataWhenEmpty", Err.Number, Err.Description
End Sub

'@sub-title Verify HasNoData returns False when RNG_GeoUpdated indicates data exists.
'@details
'Arranges a geobase fixture and overwrites RNG_GeoUpdated with
'"updated, not translated" to simulate imported data. Acts by creating
'an LLGeo instance and reading HasNoData. Asserts that the property is
'False, confirming the class detects the non-empty state.
'@TestMethod("LLGeo")
Public Sub TestHasNoDataFalseWhenUpdated()
    CustomTestSetTitles Assert, "LLGeo", "TestHasNoDataFalseWhenUpdated"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()
    Dim geoStore As IHiddenNames
    Set geoStore = HiddenNames.Create(sh)
    geoStore.SetValue "RNG_GeoUpdated", "updated, not translated"

    Dim geo As ILLGeo
    Set geo = LLGeo.Create(sh)

    Assert.IsFalse geo.HasNoData, _
                   "HasNoData should be False when geobase has data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasNoDataFalseWhenUpdated", Err.Number, Err.Description
End Sub

'@section GeoNames cache tests
'===============================================================================

'@sub-title Verify GeoNames resolves a translated name from the named range cache.
'@details
'Arranges a geobase fixture with RNG_ADM1NAME set to "Province" and
'RNG_GeoUpdated set to "updated, not translated" so the cache is
'populated. Acts by creating an LLGeo instance and calling
'GeoNames("adm1_name"). Asserts that the result is "Province",
'confirming the lazy-loaded BetterArray cache returns the translated
'display name from the matching named range.
'@TestMethod("LLGeo")
Public Sub TestGeoNamesResolvesFromCache()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoNamesResolvesFromCache"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()
    sh.Range("RNG_ADM1NAME").Value = "Province"
    Dim geoStore As IHiddenNames
    Set geoStore = HiddenNames.Create(sh)
    geoStore.SetValue "RNG_GeoUpdated", "updated, not translated"

    Dim geo As ILLGeo
    Set geo = LLGeo.Create(sh)

    Assert.AreEqual "Province", geo.GeoNames("adm1_name"), _
                    "GeoNames should return translated name from RNG_ADM1NAME"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoNamesResolvesFromCache", Err.Number, Err.Description
End Sub

'@sub-title Verify GeoNames falls back to the raw name for an unknown key.
'@details
'Arranges a geobase fixture with default named ranges. Acts by creating
'an LLGeo instance and calling GeoNames("unknown_field"). Asserts that
'the result equals "unknown_field", confirming the property returns the
'raw input unchanged when no translation mapping exists in the cache.
'@TestMethod("LLGeo")
Public Sub TestGeoNamesFallsBackToRawName()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoNamesFallsBackToRawName"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()

    Dim geo As ILLGeo
    Set geo = LLGeo.Create(sh)

    Assert.AreEqual "unknown_field", geo.GeoNames("unknown_field"), _
                    "GeoNames should return the raw name when no mapping exists"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoNamesFallsBackToRawName", Err.Number, Err.Description
End Sub

'@section UpdateLevelNames tests
'===============================================================================

'@sub-title Verify that UpdateLevelNames populates RNG_ADM*NAME from T_NAMES.
'@details
'Arranges a geobase fixture with T_NAMES containing 5 data rows mapping
'level identifiers (adm1_name through hf_name) to English translations, and
'RNG_GeoLangCode set to "EN". Acts by creating an LLGeo instance and calling
'Translate (which invokes UpdateLevelNames internally). Asserts that
'RNG_ADM1NAME is populated with the expected translated value "Province",
'confirming that the programmatic lookup from T_NAMES replaces the previous
'formula-based approach.
'@TestMethod("LLGeo")
Public Sub TestUpdateLevelNamesPopulatesFromTNAMES()
    CustomTestSetTitles Assert, "LLGeo", "TestUpdateLevelNamesPopulatesFromTNAMES"
    On Error GoTo TestFail

    'Arrange
    Dim sh As Worksheet
    Set sh = BuildGeoFixture()

    'Populate T_NAMES with level-to-translation data
    Dim loNames As ListObject
    Set loNames = sh.ListObjects("T_NAMES")
    Dim baseCol As Long
    baseCol = loNames.Range.Column

    sh.Cells(2, baseCol).Value = "adm1_name"
    sh.Cells(2, baseCol + 1).Value = "Province"
    sh.Cells(3, baseCol).Value = "adm2_name"
    sh.Cells(3, baseCol + 1).Value = "District"
    sh.Cells(4, baseCol).Value = "adm3_name"
    sh.Cells(4, baseCol + 1).Value = "Commune"
    sh.Cells(5, baseCol).Value = "adm4_name"
    sh.Cells(5, baseCol + 1).Value = "Village"
    sh.Cells(6, baseCol).Value = "hf_name"
    sh.Cells(6, baseCol + 1).Value = "Health Facility"

    'Resize T_NAMES to include all 5 data rows
    loNames.Resize sh.Range(sh.Cells(loNames.Range.Row, baseCol), _
                            sh.Cells(6, baseCol + 1))

    'Set the language code so UpdateLevelNames knows which column to read
    Dim geoStore As IHiddenNames
    Set geoStore = HiddenNames.Create(sh)
    geoStore.SetValue "RNG_GeoLangCode", "EN"
    geoStore.SetValue "RNG_GeoUpdated", "updated"

    'Act: Create LLGeo and call Translate which triggers UpdateLevelNames
    Dim geo As ILLGeo
    Set geo = LLGeo.Create(sh)
    geo.Translate rawNames:=False

    'Assert
    Assert.AreEqual "Province", sh.Range("RNG_ADM1NAME").Value, _
                    "RNG_ADM1NAME should be populated from T_NAMES EN column"
    Assert.AreEqual "District", sh.Range("RNG_ADM2NAME").Value, _
                    "RNG_ADM2NAME should be populated from T_NAMES EN column"
    Assert.AreEqual "Health Facility", sh.Range("RNG_HFNAME").Value, _
                    "RNG_HFNAME should be populated from T_NAMES EN column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateLevelNamesPopulatesFromTNAMES", Err.Number, Err.Description
End Sub

'@sub-title Verify that UpdateLevelNames handles empty language code gracefully.
'@details
'Arranges a geobase fixture with T_NAMES data but RNG_GeoLangCode set to empty
'string (default). Acts by creating an LLGeo instance and calling Translate.
'Asserts that RNG_ADM1NAME remains empty, confirming that UpdateLevelNames
'silently skips when no language code is configured.
'@TestMethod("LLGeo")
Public Sub TestUpdateLevelNamesSkipsWithNoLangCode()
    CustomTestSetTitles Assert, "LLGeo", "TestUpdateLevelNamesSkipsWithNoLangCode"
    On Error GoTo TestFail

    'Arrange
    Dim sh As Worksheet
    Set sh = BuildGeoFixture()

    'Populate T_NAMES with data but leave RNG_GeoLangCode empty (default)
    Dim loNames As ListObject
    Set loNames = sh.ListObjects("T_NAMES")
    Dim baseCol As Long
    baseCol = loNames.Range.Column

    sh.Cells(2, baseCol).Value = "adm1_name"
    sh.Cells(2, baseCol + 1).Value = "Province"
    loNames.Resize sh.Range(sh.Cells(loNames.Range.Row, baseCol), _
                            sh.Cells(2, baseCol + 1))

    Dim geoStore As IHiddenNames
    Set geoStore = HiddenNames.Create(sh)
    geoStore.SetValue "RNG_GeoUpdated", "updated"

    'Act
    Dim geo As ILLGeo
    Set geo = LLGeo.Create(sh)
    geo.Translate rawNames:=False

    'Assert: RNG_ADM1NAME should remain empty since no langCode is set
    Assert.AreEqual vbNullString, CStr(sh.Range("RNG_ADM1NAME").Value & ""), _
                    "RNG_ADM1NAME should remain empty when RNG_GeoLangCode is not set"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateLevelNamesSkipsWithNoLangCode", Err.Number, Err.Description
End Sub

'@section GeoLevel tests
'===============================================================================

'@sub-title Verify GeoLevel returns an empty BetterArray when the geobase has no data.
'@details
'Arranges a geobase fixture with RNG_GeoUpdated set to "empty" (default)
'and no data rows in any admin table. Acts by creating an LLGeo instance
'and calling GeoLevel(LevelAdmin1). Asserts that the returned BetterArray
'has zero length, confirming that GeoLevel handles empty tables gracefully
'by returning an empty collection rather than raising an error.
'@TestMethod("LLGeo")
Public Sub TestGeoLevelReturnsEmptyWhenNoData()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoLevelReturnsEmptyWhenNoData"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()

    Dim geo As ILLGeo
    Set geo = LLGeo.Create(sh)

    Dim result As BetterArray
    Set result = geo.GeoLevel(LevelAdmin1)

    Assert.AreEqual CLng(0), result.Length, _
                    "GeoLevel should return empty array when geobase has no data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoLevelReturnsEmptyWhenNoData", Err.Number, Err.Description
End Sub
