Attribute VB_Name = "TestLLGeo"
Attribute VB_Description = "Tests for LLGeo class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLGeo class")

' LLGeo tests focus on factory validation and lightweight property behavior.
' Full integration tests with geo data and AutoFilter require a real geobase
' worksheet with imported data — making them unsuitable for unit tests.
' These tests verify:
' - Factory rejects Nothing
' - Factory rejects sheets missing required tables
' - Factory succeeds with a full fixture
' - HasNoData reflects RNG_GeoUpdated state
' - GeoNames resolves from named range cache
' - GeoNames falls back to raw name for unknown keys
' - GeoLevel returns empty when geobase has no data
' - AdminCode resolves translated names correctly

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const GEO_FIXTURE As String = "GeoFixture"

Private Assert As ICustomTest

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLGeo"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
    DeleteWorksheets GEO_FIXTURE, "GeoEmptyTest"
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

'@section Helpers
'===============================================================================

' Build a minimal geobase fixture with all required tables and named ranges.
' Tables have correct header columns but no data rows — RNG_GeoUpdated = "empty".
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
            sh.Cells(1, startCol).Value = "variable"
            sh.Cells(1, startCol + 1).Value = "value"
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

    'Place named ranges in row 5 to avoid table overlap
    rngNames = Array("RNG_GeoName", "RNG_GeoUpdated", "RNG_PastingGeoCol", _
                     "RNG_GeoLangCode", "RNG_HFNAME", "RNG_ADM1NAME", _
                     "RNG_ADM2NAME", "RNG_ADM3NAME", "RNG_ADM4NAME", _
                     "RNG_FormLoaded", "RNG_MetaLang")

    For counter = LBound(rngNames) To UBound(rngNames)
        Set rng = sh.Cells(5, counter - LBound(rngNames) + 1)
        rng.Name = CStr(rngNames(counter))
    Next counter

    sh.Range("RNG_GeoUpdated").Value = "empty"
    sh.Range("RNG_GeoName").Value = "test_geo"

    Set BuildGeoFixture = sh
End Function

'@section Factory validation tests
'===============================================================================

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

'@TestMethod("LLGeo")
Public Sub TestHasNoDataFalseWhenUpdated()
    CustomTestSetTitles Assert, "LLGeo", "TestHasNoDataFalseWhenUpdated"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()
    sh.Range("RNG_GeoUpdated").Value = "updated, not translated"

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

'@TestMethod("LLGeo")
Public Sub TestGeoNamesResolvesFromCache()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoNamesResolvesFromCache"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()
    sh.Range("RNG_ADM1NAME").Value = "Province"
    sh.Range("RNG_GeoUpdated").Value = "updated, not translated"

    Dim geo As ILLGeo
    Set geo = LLGeo.Create(sh)

    Assert.AreEqual "Province", geo.GeoNames("adm1_name"), _
                    "GeoNames should return translated name from RNG_ADM1NAME"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoNamesResolvesFromCache", Err.Number, Err.Description
End Sub

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

'@section GeoLevel tests
'===============================================================================

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
