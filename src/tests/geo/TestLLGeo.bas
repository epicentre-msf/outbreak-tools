Attribute VB_Name = "TestLLGeo"
Attribute VB_Description = "Tests for LLGeo class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLGeo class")

'@description
'Validates the LLGeo class, which manages the geobase worksheet containing
'ListObjects for four admin levels (T_ADM1 through T_ADM4), health facilities
'(T_HF), translation names (T_NAMES), historic entries, and metadata.
'
'THE FIXTURE COMES IN TWO SHAPES
'-------------------------------------------------------------------------------
'BuildGeoFixture builds the nine tables with headers alone, which is what the
'factory and flag tests need. BuildGeoFixture(withData:=True) fills them with
'3 admin1 by 2 admin2 by 2 admin3 rows, a health facility table and the five
'level translations, which is what the cascade, the import, the translation
'round trip and the historic tests need. The third admin1 value is the number
'3, because two call sites in EventLinelist hand GeoLevel a raw cell value.
'
'T_HF CARRIES A COLUMN THE CLASS HAS NOTHING FOR
'-------------------------------------------------------------------------------
'Its columns are hf_name, hf_pcode, hf_extra, adm3_name, adm2_name, adm1_name.
'The extra column and the reversed level order are what pin the header
'translation to a lookup by name.
'
'THE FIVE LEVEL LABELS ARE WORKBOOK-SCOPED
'-------------------------------------------------------------------------------
'They are hidden names on the workbook, so they outlive a worksheet clear.
'Every fixture build drops them first, and ModuleCleanup leaves five empty
'ones behind for whatever else in the workbook resolves them.
'@depends LLGeo, BetterArray, CustomTest, TestHelpers, HiddenNames, Passwords

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const GEO_FIXTURE As String = "GeoFixture"
Private Const GEO_SHEET_NAME As String = "Geo"

Private Assert As CustomTest

'@section Module lifecycle
'===============================================================================

'@sub-title Initialise the test module before any tests run.
'@details
'Suppresses screen updates via BusyApp, ensures the test output sheet
'exists, creates the CustomTest assertion object targeting that sheet,
'and sets the module name for result grouping.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLGeo"
End Sub

'@sub-title Tear down the module after all tests complete.
'@details
'Prints accumulated test results to the output sheet, restores the
'application state via RestoreApp, releases the assertion object, deletes
'the temporary worksheets, and leaves the five level labels defined and
'empty.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
    DeleteWorksheets GEO_FIXTURE, "GeoEmptyTest"
    ResetLevelNames ThisWorkbook
End Sub

'@sub-title Reset state before each individual test.
'@details
'Suppresses screen updates so worksheet operations during each test do
'not trigger flickering or event cascades.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Clean up after each individual test.
'@details
'Flushes any pending assertion results to the output sheet so each test's
'outcome is recorded before the next test begins.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Fixture helpers
'===============================================================================

'@sub-title Drop the five workbook-scoped level labels.
'@details
'They survive a worksheet clear because a hidden name holding a constant
'refers to no sheet, so each fixture build has to drop them by hand.
'@param wb Workbook. The workbook holding them.
Private Sub DropLevelNames(ByVal wb As Workbook)
    Dim ids As Variant
    Dim counter As Long

    ids = LevelNameIds()

    For counter = LBound(ids) To UBound(ids)
        On Error Resume Next
        wb.Names(CStr(ids(counter))).Delete
        On Error GoTo 0
    Next
End Sub

'@sub-title Leave the five level labels defined and empty.
'@param wb Workbook. The workbook holding them.
Private Sub ResetLevelNames(ByVal wb As Workbook)
    Dim store As HiddenNames
    Dim ids As Variant
    Dim counter As Long

    DropLevelNames wb
    Set store = HiddenNames.Create(wb)
    ids = LevelNameIds()

    For counter = LBound(ids) To UBound(ids)
        store.EnsureName CStr(ids(counter)), vbNullString, HiddenNameTypeString
    Next
End Sub

'@sub-title The five hidden names holding the level labels.
'@return Variant. An array of the five identifiers.
Private Function LevelNameIds() As Variant
    LevelNameIds = Array("RNG_ADM1NAME", "RNG_ADM2NAME", "RNG_ADM3NAME", _
                         "RNG_ADM4NAME", "RNG_HFNAME")
End Function

'@sub-title Read one level label out of the workbook store.
'@param wb Workbook. The workbook holding the labels.
'@param nameId String. The hidden name identifier.
'@return String. The stored label.
Private Function LevelLabel(ByVal wb As Workbook, ByVal nameId As String) As String
    LevelLabel = HiddenNames.Create(wb).ValueAsString(nameId)
End Function

'@sub-title Write one cell of a fixture table.
'@param sh Worksheet. The fixture worksheet.
'@param Lo ListObject. The table to write into.
'@param rowOffset Long. Rows below the header row.
'@param colOffset Long. Columns right of the first column.
'@param cellValue Variant. The value to write.
Private Sub WriteTableCell(ByVal sh As Worksheet, ByVal Lo As ListObject, _
                           ByVal rowOffset As Long, ByVal colOffset As Long, _
                           ByVal cellValue As Variant)
    sh.Cells(Lo.Range.Row + rowOffset, Lo.Range.Column + colOffset).Value = cellValue
End Sub

'@sub-title Grow a fixture table to hold a number of data rows.
'@param sh Worksheet. The fixture worksheet.
'@param Lo ListObject. The table to grow.
'@param dataRows Long. The number of data rows wanted.
Private Sub ResizeTable(ByVal sh As Worksheet, ByVal Lo As ListObject, _
                        ByVal dataRows As Long)
    Dim firstRow As Long
    Dim firstCol As Long
    Dim lastCol As Long

    firstRow = Lo.Range.Row
    firstCol = Lo.Range.Column
    lastCol = firstCol + Lo.Range.Columns.Count - 1

    Lo.Resize sh.Range(sh.Cells(firstRow, firstCol), _
                       sh.Cells(firstRow + dataRows, lastCol))
End Sub

'@sub-title Build the nine geobase tables and the geo hidden names on a sheet.
'@details
'Creates T_ADM1 through T_ADM4, T_HF, T_NAMES, T_HISTOGEO, T_HISTOHF and
'T_METADATA side by side with one blank column between them, the
'RNG_PastingGeoCol cell, and the four hidden names the factory asks for.
'@param sh Worksheet. The worksheet to build on.
'@param withData Boolean. True to fill the tables with geographic data.
Private Sub BuildGeoTables(ByVal sh As Worksheet, ByVal withData As Boolean)
    Dim rng As Range
    Dim counter As Long
    Dim tblNames As Variant
    Dim tblCols As Variant
    Dim startCol As Long
    Dim geoStore As HiddenNames

    tblNames = Array("T_ADM1", "T_ADM2", "T_ADM3", "T_ADM4", "T_HF", _
                     "T_NAMES", "T_HISTOGEO", "T_HISTOHF", "T_METADATA")
    tblCols = Array(2, 3, 4, 5, 6, 2, 1, 1, 2)

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
            sh.Cells(1, startCol + 2).Value = "hf_extra"
            sh.Cells(1, startCol + 3).Value = "adm3_name"
            sh.Cells(1, startCol + 4).Value = "adm2_name"
            sh.Cells(1, startCol + 5).Value = "adm1_name"
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

    'The paste anchor stays a real cell
    sh.Cells(40, 1).Name = "RNG_PastingGeoCol"

    Set geoStore = HiddenNames.Create(sh)
    geoStore.EnsureName "RNG_GeoUpdated", "empty", HiddenNameTypeString
    geoStore.EnsureName "RNG_GeoName", "test_geo", HiddenNameTypeString
    geoStore.EnsureName "RNG_GeoLangCode", vbNullString, HiddenNameTypeString
    geoStore.EnsureName "RNG_MetaLang", vbNullString, HiddenNameTypeString

    DropLevelNames sh.Parent

    If withData Then FillGeoData sh, geoStore
End Sub

'@sub-title Fill the fixture tables with a small geobase.
'@details
'Three admin1 values, two admin2 under each and two admin3 under each of
'those, one admin4 per admin3, three health facilities and the five level
'translations in an EN column. The third admin1 value is the number 3.
'@param sh Worksheet. The fixture worksheet.
'@param geoStore HiddenNames. The hidden name store of that worksheet.
Private Sub FillGeoData(ByVal sh As Worksheet, ByVal geoStore As HiddenNames)
    Dim adm1Values As Variant
    Dim adm1Idx As Long
    Dim childIdx As Long
    Dim grandIdx As Long
    Dim adm2Row As Long
    Dim adm3Row As Long
    Dim adm1Value As Variant
    Dim adm2Value As String
    Dim adm3Value As String
    Dim loAdm1 As ListObject
    Dim loAdm2 As ListObject
    Dim loAdm3 As ListObject
    Dim loAdm4 As ListObject
    Dim loHF As ListObject
    Dim loNames As ListObject

    adm1Values = Array("P1", "P2", 3)

    Set loAdm1 = sh.ListObjects("T_ADM1")
    Set loAdm2 = sh.ListObjects("T_ADM2")
    Set loAdm3 = sh.ListObjects("T_ADM3")
    Set loAdm4 = sh.ListObjects("T_ADM4")
    Set loHF = sh.ListObjects("T_HF")
    Set loNames = sh.ListObjects("T_NAMES")

    For adm1Idx = 0 To 2
        adm1Value = adm1Values(adm1Idx)
        WriteTableCell sh, loAdm1, adm1Idx + 1, 0, adm1Value

        For childIdx = 1 To 2
            adm2Row = adm2Row + 1
            adm2Value = "D" & adm2Row
            WriteTableCell sh, loAdm2, adm2Row, 0, adm1Value
            WriteTableCell sh, loAdm2, adm2Row, 1, adm2Value

            For grandIdx = 1 To 2
                adm3Row = adm3Row + 1
                adm3Value = "C" & adm3Row

                WriteTableCell sh, loAdm3, adm3Row, 0, adm1Value
                WriteTableCell sh, loAdm3, adm3Row, 1, adm2Value
                WriteTableCell sh, loAdm3, adm3Row, 2, adm3Value

                WriteTableCell sh, loAdm4, adm3Row, 0, adm1Value
                WriteTableCell sh, loAdm4, adm3Row, 1, adm2Value
                WriteTableCell sh, loAdm4, adm3Row, 2, adm3Value
                WriteTableCell sh, loAdm4, adm3Row, 3, "V" & adm3Row
            Next grandIdx
        Next childIdx
    Next adm1Idx

    ResizeTable sh, loAdm1, 3
    ResizeTable sh, loAdm2, adm2Row
    ResizeTable sh, loAdm3, adm3Row
    ResizeTable sh, loAdm4, adm3Row

    'T_HF: hf_name, hf_pcode, hf_extra, adm3_name, adm2_name, adm1_name
    For childIdx = 1 To 3
        WriteTableCell sh, loHF, childIdx, 0, "HF" & childIdx
        WriteTableCell sh, loHF, childIdx, 1, "PC" & childIdx
        WriteTableCell sh, loHF, childIdx, 2, "X" & childIdx
        WriteTableCell sh, loHF, childIdx, 3, "C" & childIdx
        WriteTableCell sh, loHF, childIdx, 4, "D1"
        WriteTableCell sh, loHF, childIdx, 5, "P1"
    Next childIdx
    ResizeTable sh, loHF, 3

    WriteTableCell sh, loNames, 1, 0, "adm1_name"
    WriteTableCell sh, loNames, 1, 1, "Province"
    WriteTableCell sh, loNames, 2, 0, "adm2_name"
    WriteTableCell sh, loNames, 2, 1, "District"
    WriteTableCell sh, loNames, 3, 0, "adm3_name"
    WriteTableCell sh, loNames, 3, 1, "Commune"
    WriteTableCell sh, loNames, 4, 0, "adm4_name"
    WriteTableCell sh, loNames, 4, 1, "Village"
    WriteTableCell sh, loNames, 5, 0, "hf_name"
    WriteTableCell sh, loNames, 5, 1, "Health Facility"
    ResizeTable sh, loNames, 5

    geoStore.SetValue "RNG_GeoLangCode", "EN"
    geoStore.SetValue "RNG_GeoUpdated", "updated"
End Sub

'@sub-title Build a geobase fixture worksheet in the test workbook.
'@param withData Optional Boolean. True to fill the tables. Defaults to False.
'@return Worksheet. The fully prepared geobase fixture sheet.
Private Function BuildGeoFixture(Optional ByVal withData As Boolean = False) As Worksheet
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(GEO_FIXTURE, clearSheet:=True, visibility:=xlSheetHidden)
    BuildGeoTables sh, withData

    Set BuildGeoFixture = sh
End Function

'@sub-title Build a workbook of its own carrying a geobase worksheet.
'@details
'Update walks every worksheet of the workbook that owns the geo sheet, so
'the tests that drive it need a workbook whose sheets they control.
'@param withData Boolean. True to fill the tables.
'@return Workbook. A new workbook holding a Geo worksheet.
Private Function BuildGeoWorkbook(ByVal withData As Boolean) As Workbook
    Dim wb As Workbook
    Dim sh As Worksheet

    Set wb = NewWorkbook()
    Set sh = wb.Worksheets(1)
    sh.Name = GEO_SHEET_NAME
    BuildGeoTables sh, withData

    Set BuildGeoWorkbook = wb
End Function

'@sub-title Add an HList worksheet to a workbook the way a linelist carries one.
'@details
'The sheet gets a sheet_type of HList, a table_name, a _START cell holding a
'variable name, and one control hidden name written under the key VarWriter
'uses: "<variable> -- control".
'@param wb Workbook. The workbook to add it to.
'@param sheetName String. The name of the sheet.
'@param withPrintSheet Boolean. True to add the matching print_ partner.
'@return Worksheet. The HList worksheet.
Private Function AddHListSheet(ByVal wb As Workbook, ByVal sheetName As String, _
                               ByVal withPrintSheet As Boolean) As Worksheet
    Dim sh As Worksheet
    Dim printSh As Worksheet
    Dim store As HiddenNames

    Set sh = wb.Worksheets.Add
    sh.Name = sheetName

    'Row 7 holds the main label, row 8 the variable name and the _START cell
    sh.Cells(8, 2).Value = "province_var"
    sh.Cells(8, 2).Name = "TAB_ONE_START"
    sh.Cells(7, 2).Value = "Old label"

    Set store = HiddenNames.Create(sh)
    store.EnsureName "sheet_type", "HList", HiddenNameTypeString
    store.EnsureName "table_name", "TAB_ONE", HiddenNameTypeString
    store.EnsureName "province_var -- control", "geo1", HiddenNameTypeString

    If withPrintSheet Then
        Set printSh = wb.Worksheets.Add
        printSh.Name = "print_" & sheetName
    End If

    Set AddHListSheet = sh
End Function

'@section Factory validation tests
'===============================================================================

'@sub-title Verify Create refuses a Nothing worksheet with ObjectNotInitialized.
'@details
'Acts by calling LLGeo.Create with Nothing under On Error Resume Next and
'captures the error number. Asserts the number is ObjectNotInitialized, so
'the test fails when the factory raises anything else.
'@TestMethod("LLGeo")
Public Sub TestCreateRejectsNothingGeoSheet()
    CustomTestSetTitles Assert, "LLGeo", "TestCreateRejectsNothingGeoSheet"
    On Error GoTo TestFail

    Dim geo As LLGeo
    Dim errNumber As Long
    Dim errDescription As String

    On Error Resume Next
    Set geo = LLGeo.Create(Nothing)
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "Create with a Nothing sheet should raise ObjectNotInitialized - " & _
                    "description was [" & errDescription & "]"
    Assert.IsTrue (geo Is Nothing), _
                  "Create with Nothing sheet should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingGeoSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Create refuses a sheet with no tables and says which is missing.
'@details
'Arranges a blank hidden worksheet with no ListObjects. Acts by calling
'LLGeo.Create with it and captures the error number. Asserts the number is
'ElementNotFound, so a failure for any other reason fails the test.
'@TestMethod("LLGeo")
Public Sub TestCreateRejectsSheetMissingTables()
    CustomTestSetTitles Assert, "LLGeo", "TestCreateRejectsSheetMissingTables"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim errNumber As Long
    Dim errDescription As String

    Set sh = EnsureWorksheet("GeoEmptyTest", clearSheet:=True, visibility:=xlSheetHidden)

    On Error Resume Next
    Set geo = LLGeo.Create(sh)
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "Create with a sheet holding no table should raise ElementNotFound - " & _
                    "description was [" & errDescription & "]"
    Assert.IsTrue (geo Is Nothing), _
                  "Create with sheet missing required tables should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsSheetMissingTables", Err.Number, Err.Description
End Sub

'@sub-title Verify Create succeeds with a fully populated geobase fixture.
'@details
'Arranges a complete geobase fixture via BuildGeoFixture containing all
'nine ListObjects and the required hidden names. Acts by calling
'LLGeo.Create with that sheet. Asserts that the result is not Nothing.
'@TestMethod("LLGeo")
Public Sub TestCreateSucceedsWithFullFixture()
    CustomTestSetTitles Assert, "LLGeo", "TestCreateSucceedsWithFullFixture"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()

    Dim geo As LLGeo
    Set geo = LLGeo.Create(sh)

    Assert.IsNotNothing geo, _
                        "Create with full fixture should succeed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateSucceedsWithFullFixture", Err.Number, Err.Description
End Sub

'@sub-title Verify Create leaves the five level labels defined on the workbook.
'@details
'Arranges a fixture whose five level labels have been dropped. Acts by
'creating an LLGeo instance. Asserts the five hidden names exist afterwards,
'which is what CrossTable's "= RNG_ADM1NAME" formula resolves against.
'@TestMethod("LLGeo")
Public Sub TestCreateEnsuresTheFiveLevelNames()
    CustomTestSetTitles Assert, "LLGeo", "TestCreateEnsuresTheFiveLevelNames"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim store As HiddenNames
    Dim ids As Variant
    Dim counter As Long

    Set sh = BuildGeoFixture()
    Set geo = LLGeo.Create(sh)

    Set store = HiddenNames.Create(ThisWorkbook)
    ids = LevelNameIds()

    For counter = LBound(ids) To UBound(ids)
        Assert.IsTrue store.HasName(CStr(ids(counter))), _
                      "Create should leave " & CStr(ids(counter)) & " defined on the workbook"
    Next

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateEnsuresTheFiveLevelNames", Err.Number, Err.Description
End Sub

'@sub-title Verify a level label held in a cell moves into the hidden name.
'@details
'Arranges a fixture, then writes a cell-based named range RNG_ADM1NAME the
'way a geo sheet built before the move carries it. Acts by creating an LLGeo
'instance. Asserts GeoNames answers the value that was in the cell, which is
'the whole migration off the cells.
'@TestMethod("LLGeo")
Public Sub TestALevelLabelInACellMovesIntoTheHiddenName()
    CustomTestSetTitles Assert, "LLGeo", "TestALevelLabelInACellMovesIntoTheHiddenName"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo

    Set sh = BuildGeoFixture()
    sh.Cells(42, 1).Value = "Region"
    sh.Cells(42, 1).Name = "RNG_ADM1NAME"

    Set geo = LLGeo.Create(sh)

    Assert.AreEqual "Region", geo.GeoNames("adm1_name"), _
                    "A level label held in a cell should move into the hidden name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALevelLabelInACellMovesIntoTheHiddenName", Err.Number, Err.Description
End Sub

'@section HasNoData tests
'===============================================================================

'@sub-title Verify HasNoData returns True when RNG_GeoUpdated reads "empty".
'@details
'Arranges a geobase fixture whose RNG_GeoUpdated is set to "empty" by
'default. Acts by creating an LLGeo instance and reading HasNoData.
'Asserts that the property is True.
'@TestMethod("LLGeo")
Public Sub TestHasNoDataWhenEmpty()
    CustomTestSetTitles Assert, "LLGeo", "TestHasNoDataWhenEmpty"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()

    Dim geo As LLGeo
    Set geo = LLGeo.Create(sh)

    Assert.IsTrue geo.HasNoData, _
                  "HasNoData should be True when RNG_GeoUpdated is 'empty'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasNoDataWhenEmpty", Err.Number, Err.Description
End Sub

'@sub-title Verify HasNoData returns False when RNG_GeoUpdated indicates data exists.
'@details
'Arranges a geobase fixture and overwrites RNG_GeoUpdated with the combined
'string a workbook written before the flag split carries. Acts by creating
'an LLGeo instance and reading HasNoData. Asserts that the property is
'False, so the old string still answers the question it answered.
'@TestMethod("LLGeo")
Public Sub TestHasNoDataFalseWhenUpdated()
    CustomTestSetTitles Assert, "LLGeo", "TestHasNoDataFalseWhenUpdated"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()
    Dim geoStore As HiddenNames
    Set geoStore = HiddenNames.Create(sh)
    geoStore.SetValue "RNG_GeoUpdated", "updated, not translated"

    Dim geo As LLGeo
    Set geo = LLGeo.Create(sh)

    Assert.IsFalse geo.HasNoData, _
                   "HasNoData should be False when geobase has data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasNoDataFalseWhenUpdated", Err.Number, Err.Description
End Sub

'@sub-title Verify the old combined flag string is split into two hidden names.
'@details
'Arranges a fixture whose RNG_GeoUpdated reads "updated, translated", which
'is what a workbook written before the split carries. Acts by creating an
'LLGeo instance and reading HasNoData, which loads the flags. Asserts
'RNG_GeoUpdated now reads "updated" alone and RNG_GeoTranslated is True.
'@TestMethod("LLGeo")
Public Sub TestTheOldCombinedFlagSplitsIntoTwoNames()
    CustomTestSetTitles Assert, "LLGeo", "TestTheOldCombinedFlagSplitsIntoTwoNames"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim geoStore As HiddenNames

    Set sh = BuildGeoFixture()
    Set geoStore = HiddenNames.Create(sh)
    geoStore.SetValue "RNG_GeoUpdated", "updated, translated"

    Set geo = LLGeo.Create(sh)
    Assert.IsFalse geo.HasNoData, "A geobase reading 'updated, translated' holds data"

    Set geoStore = HiddenNames.Create(sh)
    Assert.AreEqual "updated", geoStore.ValueAsString("RNG_GeoUpdated"), _
                    "The data flag should hold the data fact alone"
    Assert.IsTrue geoStore.ValueAsBoolean("RNG_GeoTranslated"), _
                  "The translation fact should move to its own boolean name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheOldCombinedFlagSplitsIntoTwoNames", Err.Number, Err.Description
End Sub

'@sub-title Verify an export leaves HasNoData answering what it answered.
'@details
'Arranges an empty geobase and reads HasNoData. Acts by exporting to a new
'workbook, which reverts the headers and puts them back. Asserts HasNoData
'is True both before and after, so the export cannot make an empty geobase
'report that it holds data.
'@TestMethod("LLGeo")
Public Sub TestHasNoDataSurvivesAnExport()
    CustomTestSetTitles Assert, "LLGeo", "TestHasNoDataSurvivesAnExport"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim outWkb As Workbook

    Set sh = BuildGeoFixture()
    Set geo = LLGeo.Create(sh)
    Assert.IsTrue geo.HasNoData, "The fixture starts empty"

    Set outWkb = NewWorkbook()
    geo.ExportToFile outWkb
    DeleteWorkbook outWkb

    Assert.IsTrue geo.HasNoData, _
                  "An export should leave an empty geobase reporting that it is empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHasNoDataSurvivesAnExport", Err.Number, Err.Description
End Sub

'@section GeoNames cache tests
'===============================================================================

'@sub-title Verify GeoNames resolves a level label from the workbook store.
'@details
'Arranges a geobase fixture and writes "Province" into RNG_ADM1NAME on the
'workbook store. Acts by creating an LLGeo instance and calling
'GeoNames("adm1_name"). Asserts that the result is "Province".
'@TestMethod("LLGeo")
Public Sub TestGeoNamesResolvesFromCache()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoNamesResolvesFromCache"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim store As HiddenNames
    Dim geo As LLGeo

    Set sh = BuildGeoFixture()

    Set store = HiddenNames.Create(ThisWorkbook)
    store.EnsureName "RNG_ADM1NAME", "Province", HiddenNameTypeString
    store.SetValue "RNG_ADM1NAME", "Province"

    Set geo = LLGeo.Create(sh)

    Assert.AreEqual "Province", geo.GeoNames("adm1_name"), _
                    "GeoNames should return the label held in RNG_ADM1NAME"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoNamesResolvesFromCache", Err.Number, Err.Description
End Sub

'@sub-title Verify GeoNames falls back to the raw name for an unknown key.
'@details
'Arranges a geobase fixture. Acts by creating an LLGeo instance and calling
'GeoNames("unknown_field"). Asserts that the result equals "unknown_field".
'@TestMethod("LLGeo")
Public Sub TestGeoNamesFallsBackToRawName()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoNamesFallsBackToRawName"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()

    Dim geo As LLGeo
    Set geo = LLGeo.Create(sh)

    Assert.AreEqual "unknown_field", geo.GeoNames("unknown_field"), _
                    "GeoNames should return the raw name when no label is stored"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoNamesFallsBackToRawName", Err.Number, Err.Description
End Sub

'@sub-title Verify GeoBaseName answers what Import stored.
'@details
'Arranges a geobase fixture whose RNG_GeoName reads "test_geo". Acts by
'creating an LLGeo instance and reading GeoBaseName. Asserts the value comes
'back, which is what LLExporter puts in the file name of an exported
'geobase.
'@TestMethod("LLGeo")
Public Sub TestGeoBaseNameAnswersTheStoredName()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoBaseNameAnswersTheStoredName"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo

    Set sh = BuildGeoFixture()
    Set geo = LLGeo.Create(sh)

    Assert.AreEqual "test_geo", geo.GeoBaseName(), _
                    "GeoBaseName should answer the name held in RNG_GeoName"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoBaseNameAnswersTheStoredName", Err.Number, Err.Description
End Sub

'@section UpdateLevelNames tests
'===============================================================================

'@sub-title Verify Translate fills the five level labels from T_NAMES.
'@details
'Arranges a geobase fixture carrying the five translations in an EN column
'and a language code of EN. Acts by calling Translate. Asserts the three
'labels land in the workbook store.
'@TestMethod("LLGeo")
Public Sub TestUpdateLevelNamesPopulatesFromTNAMES()
    CustomTestSetTitles Assert, "LLGeo", "TestUpdateLevelNamesPopulatesFromTNAMES"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo

    Set sh = BuildGeoFixture(withData:=True)
    Set geo = LLGeo.Create(sh)
    geo.Translate rawNames:=False

    Assert.AreEqual "Province", LevelLabel(ThisWorkbook, "RNG_ADM1NAME"), _
                    "RNG_ADM1NAME should be filled from the T_NAMES EN column"
    Assert.AreEqual "District", LevelLabel(ThisWorkbook, "RNG_ADM2NAME"), _
                    "RNG_ADM2NAME should be filled from the T_NAMES EN column"
    Assert.AreEqual "Health Facility", LevelLabel(ThisWorkbook, "RNG_HFNAME"), _
                    "RNG_HFNAME should be filled from the T_NAMES EN column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateLevelNamesPopulatesFromTNAMES", Err.Number, Err.Description
End Sub

'@sub-title Verify Translate skips the level labels with no language code.
'@details
'Arranges a geobase fixture with T_NAMES data and an empty RNG_GeoLangCode.
'Acts by calling Translate. Asserts RNG_ADM1NAME stays empty.
'@TestMethod("LLGeo")
Public Sub TestUpdateLevelNamesSkipsWithNoLangCode()
    CustomTestSetTitles Assert, "LLGeo", "TestUpdateLevelNamesSkipsWithNoLangCode"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim geoStore As HiddenNames

    Set sh = BuildGeoFixture(withData:=True)
    Set geoStore = HiddenNames.Create(sh)
    geoStore.SetValue "RNG_GeoLangCode", vbNullString

    Set geo = LLGeo.Create(sh)
    geo.Translate rawNames:=False

    Assert.AreEqual vbNullString, LevelLabel(ThisWorkbook, "RNG_ADM1NAME"), _
                    "RNG_ADM1NAME should stay empty when RNG_GeoLangCode is not set"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateLevelNamesSkipsWithNoLangCode", Err.Number, Err.Description
End Sub

'@section Translate tests
'===============================================================================

'@sub-title Verify a translate round trip puts every header back as it was.
'@details
'Arranges a geobase fixture whose T_HF carries an extra column and holds its
'level columns in reverse order. Acts by translating and reverting. Asserts
'every level header of T_ADM4 and T_HF reads its raw name again and the
'extra column is untouched, which a write by fixed offset could not do.
'@TestMethod("LLGeo")
Public Sub TestTranslateRoundTripsEveryHeader()
    CustomTestSetTitles Assert, "LLGeo", "TestTranslateRoundTripsEveryHeader"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim hfHeader As Range
    Dim adm4Header As Range

    Set sh = BuildGeoFixture(withData:=True)
    Set geo = LLGeo.Create(sh)

    geo.Translate rawNames:=False

    Set hfHeader = sh.ListObjects("T_HF").HeaderRowRange
    Assert.AreEqual "Health Facility", CStr(hfHeader.Cells(1, 1).Value), _
                    "The first T_HF header should carry the facility label"
    Assert.AreEqual "Commune", CStr(hfHeader.Cells(1, 4).Value), _
                    "The T_HF admin 3 column should carry its label, found by name"
    Assert.AreEqual "hf_extra", CStr(hfHeader.Cells(1, 3).Value), _
                    "A T_HF column the class has nothing for should be left alone"

    geo.Translate rawNames:=True

    Set hfHeader = sh.ListObjects("T_HF").HeaderRowRange
    Set adm4Header = sh.ListObjects("T_ADM4").HeaderRowRange

    Assert.AreEqual "hf_name", CStr(hfHeader.Cells(1, 1).Value), _
                    "The first T_HF header should read its raw name again"
    Assert.AreEqual "hf_pcode", CStr(hfHeader.Cells(1, 2).Value), _
                    "The T_HF pcode column should read its raw name"
    Assert.AreEqual "hf_extra", CStr(hfHeader.Cells(1, 3).Value), _
                    "The extra T_HF column should read its raw name"
    Assert.AreEqual "adm3_name", CStr(hfHeader.Cells(1, 4).Value), _
                    "The T_HF admin 3 column should read its raw name again"
    Assert.AreEqual "adm1_name", CStr(hfHeader.Cells(1, 6).Value), _
                    "The T_HF admin 1 column should read its raw name again"
    Assert.AreEqual "adm4_name", CStr(adm4Header.Cells(1, 4).Value), _
                    "The T_ADM4 admin 4 header should read its raw name again"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTranslateRoundTripsEveryHeader", Err.Number, Err.Description
End Sub

'@sub-title Verify Clear puts the raw headers back before it empties the tables.
'@details
'Arranges a translated geobase. Acts by calling Clear on its own, the way
'EventsDesignerAdvanced does. Asserts the T_ADM1 header reads its raw name,
'so the column lookups of the class still find it.
'@TestMethod("LLGeo")
Public Sub TestClearPutsTheRawHeadersBack()
    CustomTestSetTitles Assert, "LLGeo", "TestClearPutsTheRawHeadersBack"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo

    Set sh = BuildGeoFixture(withData:=True)
    Set geo = LLGeo.Create(sh)
    geo.Translate rawNames:=False

    Assert.AreEqual "Province", CStr(sh.ListObjects("T_ADM1").HeaderRowRange.Cells(1, 1).Value), _
                    "The header carries the label before the clear"

    geo.Clear

    Assert.AreEqual "adm1_name", CStr(sh.ListObjects("T_ADM1").HeaderRowRange.Cells(1, 1).Value), _
                    "Clear should put the raw column name back on the header"
    Assert.IsTrue geo.HasNoData, "Clear should leave the geobase reporting no data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestClearPutsTheRawHeadersBack", Err.Number, Err.Description
End Sub

'@section GeoLevel tests
'===============================================================================

'@sub-title Verify GeoLevel returns an empty BetterArray when the geobase has no data.
'@details
'Arranges a geobase fixture with RNG_GeoUpdated set to "empty" and no data
'rows. Acts by calling GeoLevel(LevelAdmin1). Asserts the array is empty.
'@TestMethod("LLGeo")
Public Sub TestGeoLevelReturnsEmptyWhenNoData()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoLevelReturnsEmptyWhenNoData"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildGeoFixture()

    Dim geo As LLGeo
    Set geo = LLGeo.Create(sh)

    Dim result As BetterArray
    Set result = geo.GeoLevel(LevelAdmin1)

    Assert.AreEqual CLng(0), result.Length, _
                    "GeoLevel should return empty array when geobase has no data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoLevelReturnsEmptyWhenNoData", Err.Number, Err.Description
End Sub

'@sub-title Verify GeoLevel lists the whole first level.
'@details
'Arranges a geobase fixture holding three admin 1 values. Acts by calling
'GeoLevel(LevelAdmin1). Asserts all three come back.
'@TestMethod("LLGeo")
Public Sub TestGeoLevelListsTheFirstLevel()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoLevelListsTheFirstLevel"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim result As BetterArray

    Set sh = BuildGeoFixture(withData:=True)
    Set geo = LLGeo.Create(sh)
    Set result = geo.GeoLevel(LevelAdmin1)

    Assert.AreEqual CLng(3), result.Length, _
                    "The three admin 1 values should come back"
    Assert.IsTrue result.Includes("P1"), "P1 should be in the admin 1 list"
    Assert.IsTrue result.Includes("P2"), "P2 should be in the admin 1 list"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoLevelListsTheFirstLevel", Err.Number, Err.Description
End Sub

'@sub-title Verify GeoLevel returns exactly the children of one parent.
'@details
'Arranges a geobase fixture where P1 has two admin 2 children. Acts by
'calling GeoLevel(LevelAdmin2, GeoScopeAdmin, "P1"). Asserts the two
'children come back and nothing else, which pins the AutoFilter cascade and
'the in-memory dedupe together.
'@TestMethod("LLGeo")
Public Sub TestGeoLevelReturnsTheChildrenOfOneParent()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoLevelReturnsTheChildrenOfOneParent"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim result As BetterArray

    Set sh = BuildGeoFixture(withData:=True)
    Set geo = LLGeo.Create(sh)
    Set result = geo.GeoLevel(LevelAdmin2, GeoScopeAdmin, "P1")

    Assert.AreEqual CLng(2), result.Length, _
                    "P1 should answer with its two admin 2 children"
    Assert.IsTrue result.Includes("D1"), "D1 is a child of P1"
    Assert.IsTrue result.Includes("D2"), "D2 is a child of P1"
    Assert.IsFalse result.Includes("D3"), "D3 belongs to another parent"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoLevelReturnsTheChildrenOfOneParent", Err.Number, Err.Description
End Sub

'@sub-title Verify GeoLevel takes a numeric parent name.
'@details
'Arranges a geobase fixture whose third admin 1 value is the number 3, which
'is what a geobase with numeric admin codes carries and what a raw cell value
'from EventLinelist hands over. Acts by asking for its children. Asserts the
'two children come back, so the argument test no longer refuses an ordinary
'cell value and hands the caller an empty dropdown.
'@TestMethod("LLGeo")
Public Sub TestGeoLevelTakesANumericParentName()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoLevelTakesANumericParentName"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim result As BetterArray
    Dim parentValue As Variant

    Set sh = BuildGeoFixture(withData:=True)
    Set geo = LLGeo.Create(sh)

    parentValue = sh.ListObjects("T_ADM1").DataBodyRange.Cells(3, 1).Value
    Set result = geo.GeoLevel(LevelAdmin2, GeoScopeAdmin, parentValue)

    Assert.AreEqual CLng(2), result.Length, _
                    "A numeric admin 1 value should answer with its two children"
    Assert.IsTrue result.Includes("D5"), "D5 is a child of the numeric admin 1"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoLevelTakesANumericParentName", Err.Number, Err.Description
End Sub

'@sub-title Verify GeoLevel takes an empty parent name without raising.
'@details
'Arranges a geobase fixture. Acts by asking for the children of an empty
'cell value. Asserts no error escapes and the answer is an empty list, which
'is the honest answer for a parent nothing is filed under.
'@TestMethod("LLGeo")
Public Sub TestGeoLevelTakesAnEmptyParentName()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoLevelTakesAnEmptyParentName"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim result As BetterArray
    Dim errNumber As Long
    Dim parentValue As Variant

    Set sh = BuildGeoFixture(withData:=True)
    Set geo = LLGeo.Create(sh)

    parentValue = sh.Cells(45, 1).Value

    On Error Resume Next
    Set result = geo.GeoLevel(LevelAdmin2, GeoScopeAdmin, parentValue)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(0), errNumber, _
                    "An empty parent value should not raise"
    Assert.AreEqual CLng(0), result.Length, _
                    "An empty parent value should answer an empty list"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoLevelTakesAnEmptyParentName", Err.Number, Err.Description
End Sub

'@sub-title Verify a wrong parent table reaches the caller as InvalidArgument.
'@details
'Arranges a geobase fixture. Acts by asking for admin 3 with a table of one
'name where two are wanted. Asserts the caller sees InvalidArgument, which
'is what the old shape swallowed on its way to an empty dropdown.
'@TestMethod("LLGeo")
Public Sub TestGeoLevelRaisesOnAWrongParentTable()
    CustomTestSetTitles Assert, "LLGeo", "TestGeoLevelRaisesOnAWrongParentTable"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim result As BetterArray
    Dim parents As BetterArray
    Dim errNumber As Long
    Dim errDescription As String

    Set sh = BuildGeoFixture(withData:=True)
    Set geo = LLGeo.Create(sh)

    Set parents = New BetterArray
    parents.LowerBound = 1
    parents.Push "P1"

    On Error Resume Next
    Set result = geo.GeoLevel(LevelAdmin3, GeoScopeAdmin, parents)
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A parent table of the wrong length should raise InvalidArgument - " & _
                    "description was [" & errDescription & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoLevelRaisesOnAWrongParentTable", Err.Number, Err.Description
End Sub

'@sub-title Verify a failed GeoLevel leaves no AutoFilter on the table.
'@details
'Arranges a geobase fixture whose T_ADM4 header no longer names adm4_name,
'so the lookup fails after the three parent filters have been applied. Acts
'by asking for admin 4. Asserts the caller sees ElementNotFound and the
'table carries no filter, so the next reader of T_ADM4 sees every row.
'@TestMethod("LLGeo")
Public Sub TestAFailedGeoLevelLeavesNoAutoFilter()
    CustomTestSetTitles Assert, "LLGeo", "TestAFailedGeoLevelLeavesNoAutoFilter"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim result As BetterArray
    Dim parents As BetterArray
    Dim Lo As ListObject
    Dim errNumber As Long
    Dim errDescription As String
    Dim stillFiltered As Boolean

    Set sh = BuildGeoFixture(withData:=True)
    Set Lo = sh.ListObjects("T_ADM4")
    Lo.HeaderRowRange.Cells(1, 4).Value = "broken_name"

    Set geo = LLGeo.Create(sh)

    Set parents = New BetterArray
    parents.LowerBound = 1
    parents.Push "P1", "D1", "C1"

    On Error Resume Next
    Set result = geo.GeoLevel(LevelAdmin4, GeoScopeAdmin, parents)
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "A missing geo column should reach the caller as ElementNotFound - " & _
                    "description was [" & errDescription & "]"

    Set Lo = sh.ListObjects("T_ADM4")
    If Not Lo.AutoFilter Is Nothing Then stillFiltered = Lo.AutoFilter.FilterMode

    Assert.IsFalse stillFiltered, _
                   "A failed GeoLevel should leave the table with no filter on it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFailedGeoLevelLeavesNoAutoFilter", Err.Number, Err.Description
End Sub

'@section AdminCode tests
'===============================================================================

'@sub-title Verify AdminCode resolves each of the four levels.
'@details
'Arranges a translated geobase. Acts by asking for the code of each level
'label and of a label that names no level. Asserts the four codes and the
'documented answer for a miss.
'@TestMethod("LLGeo")
Public Sub TestAdminCodeResolvesEachLevel()
    CustomTestSetTitles Assert, "LLGeo", "TestAdminCodeResolvesEachLevel"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo

    Set sh = BuildGeoFixture(withData:=True)
    Set geo = LLGeo.Create(sh)
    geo.Translate rawNames:=False

    Assert.AreEqual "adm1", geo.AdminCode("Province"), "Province is admin 1"
    Assert.AreEqual "adm2", geo.AdminCode("District"), "District is admin 2"
    Assert.AreEqual "adm3", geo.AdminCode("Commune"), "Commune is admin 3"
    Assert.AreEqual "adm4", geo.AdminCode("Village"), "Village is admin 4"
    Assert.AreEqual "adm1", geo.AdminCode("Nowhere"), _
                    "A label naming no level answers admin 1, which is what EventLinelist reads"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAdminCodeResolvesEachLevel", Err.Number, Err.Description
End Sub

'@section Historic tests
'===============================================================================

'@sub-title Verify the row UpdateHistoric just wrote is de-duplicated and placed.
'@details
'Arranges an empty geobase. Acts by writing the same value twice and then a
'value that sorts before it. Asserts one row per value and the first row
'holding the one that sorts first, which the old shape could not do because
'it worked on the range as it stood before the row was written.
'@TestMethod("LLGeo")
Public Sub TestUpdateHistoricDeduplicatesAndSorts()
    CustomTestSetTitles Assert, "LLGeo", "TestUpdateHistoricDeduplicatesAndSorts"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim Lo As ListObject

    Set sh = BuildGeoFixture()
    Set geo = LLGeo.Create(sh)

    geo.UpdateHistoric "Zulu"
    geo.UpdateHistoric "Zulu"
    geo.UpdateHistoric "Alpha"

    Set Lo = sh.ListObjects("T_HISTOGEO")

    Assert.AreEqual CLng(2), CLng(Lo.ListRows.Count), _
                    "Writing the same value twice should leave one row for it"
    Assert.AreEqual "Alpha", CStr(Lo.DataBodyRange.Cells(1, 1).Value), _
                    "The table should be sorted, so Alpha comes first"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateHistoricDeduplicatesAndSorts", Err.Number, Err.Description
End Sub

'@sub-title Verify ClearHistoric empties only the table its scope names.
'@details
'Arranges a geobase holding one geo entry and one facility entry. Acts by
'clearing the facility scope alone. Asserts the geo table keeps its row.
'@TestMethod("LLGeo")
Public Sub TestClearHistoricEmptiesOneScope()
    CustomTestSetTitles Assert, "LLGeo", "TestClearHistoricEmptiesOneScope"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo

    Set sh = BuildGeoFixture()
    Set geo = LLGeo.Create(sh)

    geo.UpdateHistoric "Zone A", GeoScopeAdmin
    geo.UpdateHistoric "Facility A", GeoScopeHF

    geo.ClearHistoric GeoScopeHF

    Assert.AreEqual CLng(1), CLng(sh.ListObjects("T_HISTOGEO").ListRows.Count), _
                    "Clearing the facility scope should leave the geo history alone"
    Assert.AreEqual vbNullString, _
                    CStr(sh.ListObjects("T_HISTOHF").DataBodyRange.Cells(1, 1).Value), _
                    "The facility history should be empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestClearHistoricEmptiesOneScope", Err.Number, Err.Description
End Sub

'@section Import tests
'===============================================================================

'@sub-title Verify Import lands every row and fills the whole concat column.
'@details
'Arranges a source workbook with an ADM1 sheet of three rows and an ADM2
'sheet of four. Acts by importing it. Asserts both row counts and that the
'concat column of T_ADM2 is filled on every row, which the single-cell write
'left to an Excel option a user can switch off.
'@TestMethod("LLGeo")
Public Sub TestImportLandsTheRowsAndFillsTheConcatColumn()
    CustomTestSetTitles Assert, "LLGeo", "TestImportLandsTheRowsAndFillsTheConcatColumn"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim sourceWkb As Workbook
    Dim sourceSh As Worksheet
    Dim loAdm2 As ListObject
    Dim counter As Long
    Dim filledRows As Long

    Set sh = BuildGeoFixture()
    Set geo = LLGeo.Create(sh)

    Set sourceWkb = NewWorkbook()
    Set sourceSh = sourceWkb.Worksheets(1)
    sourceSh.Name = "ADM1"
    WriteColumn sourceSh.Cells(1, 1), "adm1_name", "P1", "P2", "P3"

    Set sourceSh = sourceWkb.Worksheets.Add
    sourceSh.Name = "ADM2"
    WriteColumn sourceSh.Cells(1, 1), "adm1_name", "P1", "P1", "P2", "P2"
    WriteColumn sourceSh.Cells(1, 2), "adm2_name", "D1", "D2", "D3", "D4"

    geo.Import sourceWkb
    DeleteWorkbook sourceWkb

    Set loAdm2 = sh.ListObjects("T_ADM2")

    Assert.AreEqual CLng(3), CLng(sh.ListObjects("T_ADM1").ListRows.Count), _
                    "The three admin 1 rows should land"
    Assert.AreEqual CLng(4), CLng(loAdm2.ListRows.Count), _
                    "The four admin 2 rows should land"

    For counter = 1 To loAdm2.ListRows.Count
        If LenB(CStr(loAdm2.DataBodyRange.Cells(counter, 3).Value)) > 0 Then
            filledRows = filledRows + 1
        End If
    Next

    Assert.AreEqual CLng(4), filledRows, _
                    "Every row of the concat column should carry the formula result"
    Assert.IsFalse geo.HasNoData, "A geobase that took an ADM1 table holds data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestImportLandsTheRowsAndFillsTheConcatColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify Import reads a source whose first rows are blank.
'@details
'Arranges a source workbook whose ADM1 header sits on row 3. Acts by
'importing it. Asserts all three data rows land, which the old row count
'read as a row number dropped the last of.
'@TestMethod("LLGeo")
Public Sub TestImportReadsASourceStartingBelowRowOne()
    CustomTestSetTitles Assert, "LLGeo", "TestImportReadsASourceStartingBelowRowOne"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim sourceWkb As Workbook
    Dim sourceSh As Worksheet

    Set sh = BuildGeoFixture()
    Set geo = LLGeo.Create(sh)

    Set sourceWkb = NewWorkbook()
    Set sourceSh = sourceWkb.Worksheets(1)
    sourceSh.Name = "ADM1"
    WriteColumn sourceSh.Cells(3, 1), "adm1_name", "P1", "P2", "P3"

    geo.Import sourceWkb
    DeleteWorkbook sourceWkb

    Assert.AreEqual CLng(3), CLng(sh.ListObjects("T_ADM1").ListRows.Count), _
                    "Every row should land when the source header is below row 1"
    Assert.AreEqual "P3", _
                    CStr(sh.ListObjects("T_ADM1").DataBodyRange.Cells(3, 1).Value), _
                    "The last record of the source should land"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestImportReadsASourceStartingBelowRowOne", Err.Number, Err.Description
End Sub

'@sub-title Verify Import leaves the flag alone when it imported nothing.
'@details
'Arranges a source workbook holding no sheet the geobase knows. Acts by
'importing it. Asserts the geobase still reports that it has no data, so a
'user is told nothing came across.
'@TestMethod("LLGeo")
Public Sub TestImportLeavesTheFlagAloneWhenNothingCameAcross()
    CustomTestSetTitles Assert, "LLGeo", "TestImportLeavesTheFlagAloneWhenNothingCameAcross"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim geo As LLGeo
    Dim sourceWkb As Workbook

    Set sh = BuildGeoFixture()
    Set geo = LLGeo.Create(sh)

    Set sourceWkb = NewWorkbook()
    sourceWkb.Worksheets(1).Name = "SomethingElse"

    geo.Import sourceWkb
    DeleteWorkbook sourceWkb

    Assert.IsTrue geo.HasNoData, _
                  "A workbook holding none of the geo sheets should leave the geobase empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestImportLeavesTheFlagAloneWhenNothingCameAcross", Err.Number, Err.Description
End Sub

'@section Update tests
'===============================================================================

'@sub-title Verify Update writes the level label into an HList header.
'@details
'Arranges a workbook holding a geobase, an HList worksheet whose one column
'is controlled by geo1, and its print partner. Acts by calling Update.
'Asserts the header cell above the data reads the admin 1 label, which is the
'write that has not reached an HList header since the control value moved
'into HiddenNames.
'@TestMethod("LLGeo")
Public Sub TestUpdateWritesTheLevelLabelIntoAnHListHeader()
    CustomTestSetTitles Assert, "LLGeo", "TestUpdateWritesTheLevelLabelIntoAnHListHeader"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim hlistSh As Worksheet
    Dim geo As LLGeo
    Dim pass As Passwords

    Set wb = BuildGeoWorkbook(True)
    Set sh = wb.Worksheets(GEO_SHEET_NAME)
    Set hlistSh = AddHListSheet(wb, "HLIST_ONE", withPrintSheet:=True)

    PreparePasswordsFixture "PassFixture", wb
    Set pass = Passwords.Create(wb.Worksheets("PassFixture"))

    Set geo = LLGeo.Create(sh)
    geo.Translate rawNames:=False
    geo.Update pass

    Assert.AreEqual "Province", CStr(hlistSh.Cells(7, 2).Value), _
                    "The HList header should carry the admin 1 label"
    Assert.AreEqual "Province", CStr(wb.Worksheets("print_HLIST_ONE").Cells(7, 2).Value), _
                    "The print sheet should carry the same label"
    Assert.AreEqual CLng(0), geo.SkippedSheets().Length, _
                    "A sheet with its print partner should not be reported as skipped"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateWritesTheLevelLabelIntoAnHListHeader", Err.Number, Err.Description
End Sub

'@sub-title Verify Update reports an HList sheet with no print partner.
'@details
'Arranges a workbook whose HList worksheet has no print_ partner. Acts by
'calling Update. Asserts the sheet is named in SkippedSheets, its own header
'is still written, and the sheet is protected again, which the old shape
'left open when the missing partner aborted the whole walk.
'@TestMethod("LLGeo")
Public Sub TestUpdateReportsAnHListSheetWithNoPrintPartner()
    CustomTestSetTitles Assert, "LLGeo", "TestUpdateReportsAnHListSheetWithNoPrintPartner"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim hlistSh As Worksheet
    Dim geo As LLGeo
    Dim pass As Passwords

    Set wb = BuildGeoWorkbook(True)
    Set sh = wb.Worksheets(GEO_SHEET_NAME)
    Set hlistSh = AddHListSheet(wb, "HLIST_LONE", withPrintSheet:=False)

    PreparePasswordsFixture "PassFixture", wb
    Set pass = Passwords.Create(wb.Worksheets("PassFixture"))

    Set geo = LLGeo.Create(sh)
    geo.Translate rawNames:=False
    geo.Update pass

    Assert.IsTrue geo.SkippedSheets().Includes("HLIST_LONE"), _
                  "A sheet with no print partner should be reported to the caller"
    Assert.AreEqual "Province", CStr(hlistSh.Cells(7, 2).Value), _
                    "The sheet itself should still take the label"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateReportsAnHListSheetWithNoPrintPartner", Err.Number, Err.Description
End Sub
