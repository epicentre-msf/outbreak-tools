Attribute VB_Name = "GeoTestFixture"
Attribute VB_Description = "Builds a geobase worksheet for the suites that need one"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Builds a geobase worksheet for the suites that need one")

Option Explicit

'@description
'Builds the nine geobase ListObjects LLGeo reads, and fills them with a small
'geography when a suite asks for one. It carries no test of its own: the runner
'imports it and runs it as a zero-test module, the same as the other fixture
'modules of this folder.
'
'WHY IT IS SHARED
'-------------------------------------------------------------------------------
'These routines were private to TestLLGeo. A geobase is what turns a geo
'variable into geo1 to geo4 columns -- LLdictionary.Prepare expands one only
'when it is handed an LLGeo -- so any suite that wants to drive the geo cascade
'of a generated linelist needs the same tables. TestEventLinelistSheets is the
'second caller and the reason this module exists.
'
'THE SHAPE OF THE FIXTURE
'-------------------------------------------------------------------------------
'BuildGeoFixtureTables with withData False writes the nine tables with their
'headers alone, which is what the factory and flag tests of TestLLGeo need.
'With withData True it fills them: 3 admin1, 2 admin2 under each, 2 admin3
'under each of those, one admin4 per admin3, three health facilities and the
'five level translations in an EN column.
'
'THE THIRD ADMIN1 VALUE IS THE NUMBER 3
'-------------------------------------------------------------------------------
'Two call sites in EventLinelist hand GeoLevel a raw cell value, so one level
'value of the fixture is a number rather than text.
'
'T_HF CARRIES A COLUMN THE CLASS HAS NOTHING FOR
'-------------------------------------------------------------------------------
'Its columns are hf_name, hf_pcode, hf_extra, adm3_name, adm2_name, adm1_name.
'The extra column and the reversed level order are what pin the header
'translation to a lookup by name.
'
'THE FIVE LEVEL LABELS ARE WORKBOOK-SCOPED
'-------------------------------------------------------------------------------
'They are hidden names on the workbook holding constants, so they refer to no
'worksheet and survive a worksheet clear. Every build drops them first.
'@depends HiddenNames, TestHelpersLite

'The paste anchor LLGeo resolves. It stays a real cell.
Private Const PASTING_ANCHOR As String = "RNG_PastingGeoCol"
Private Const PASTING_ANCHOR_CELL As String = "A40"


'@section Public entry points
'===============================================================================

'@fun-title Build a geobase worksheet, creating or clearing it first.
'@details
'The convenience form. A caller that already holds the worksheet it wants --
'because it renamed the first sheet of a fresh workbook, say -- calls
'BuildGeoFixtureTables instead and keeps its own sheet.
'@param sheetName String. The worksheet to build on.
'@param targetBook Optional Workbook. The host workbook. Defaults to ThisWorkbook.
'@param withData Optional Boolean. True to fill the tables. Defaults to False.
'@param visibility Optional Long. Sheet visibility. Defaults to xlSheetHidden.
'@return Worksheet. The prepared geobase worksheet.
Public Function PrepareGeoFixture(ByVal sheetName As String, _
                                  Optional ByVal targetBook As Workbook, _
                                  Optional ByVal withData As Boolean = False, _
                                  Optional ByVal visibility As Long = xlSheetHidden) As Worksheet
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(sheetName, targetBook, clearSheet:=True, visibility:=visibility)
    BuildGeoFixtureTables sh, withData

    Set PrepareGeoFixture = sh
End Function

'@sub-title Build the nine geobase tables and the geo hidden names on a sheet.
'@details
'Creates T_ADM1 through T_ADM4, T_HF, T_NAMES, T_HISTOGEO, T_HISTOHF and
'T_METADATA side by side with one blank column between them, the paste anchor
'cell, and the four hidden names the LLGeo factory asks for.
'@param sh Worksheet. The worksheet to build on.
'@param withData Boolean. True to fill the tables with geographic data.
Public Sub BuildGeoFixtureTables(ByVal sh As Worksheet, ByVal withData As Boolean)
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

        WriteTableHeaders sh, CStr(tblNames(counter)), startCol

        sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = CStr(tblNames(counter))
        startCol = startCol + CLng(tblCols(counter)) + 1
    Next counter

    sh.Range(PASTING_ANCHOR_CELL).Name = PASTING_ANCHOR

    Set geoStore = HiddenNames.Create(sh)
    geoStore.EnsureName "RNG_GeoUpdated", "empty", HiddenNameTypeString
    geoStore.EnsureName "RNG_GeoName", "test_geo", HiddenNameTypeString
    geoStore.EnsureName "RNG_GeoLangCode", vbNullString, HiddenNameTypeString
    geoStore.EnsureName "RNG_MetaLang", vbNullString, HiddenNameTypeString

    DropGeoLevelNames sh.Parent

    If withData Then FillGeoData sh, geoStore
End Sub

'@sub-title Drop the five workbook-scoped level labels.
'@details
'They survive a worksheet clear because a hidden name holding a constant refers
'to no sheet, so each build has to drop them by hand.
'@param wb Workbook. The workbook holding them.
Public Sub DropGeoLevelNames(ByVal wb As Workbook)
    Dim ids As Variant
    Dim counter As Long

    ids = GeoFixtureLevelNameIds()

    For counter = LBound(ids) To UBound(ids)
        On Error Resume Next
        wb.Names(CStr(ids(counter))).Delete
        On Error GoTo 0
    Next counter
End Sub

'@fun-title The five hidden names holding the level labels.
'@return Variant. An array of the five identifiers.
Public Function GeoFixtureLevelNameIds() As Variant
    GeoFixtureLevelNameIds = Array("RNG_ADM1NAME", "RNG_ADM2NAME", "RNG_ADM3NAME", _
                                   "RNG_ADM4NAME", "RNG_HFNAME")
End Function

'@sub-title Write one cell of a fixture table.
'@param sh Worksheet. The fixture worksheet.
'@param Lo ListObject. The table to write into.
'@param rowOffset Long. Rows below the header row.
'@param colOffset Long. Columns right of the first column.
'@param cellValue Variant. The value to write.
Public Sub GeoFixtureWriteCell(ByVal sh As Worksheet, ByVal Lo As ListObject, _
                               ByVal rowOffset As Long, ByVal colOffset As Long, _
                               ByVal cellValue As Variant)
    sh.Cells(Lo.Range.Row + rowOffset, Lo.Range.Column + colOffset).Value = cellValue
End Sub

'@sub-title Grow a fixture table to hold a number of data rows.
'@param sh Worksheet. The fixture worksheet.
'@param Lo ListObject. The table to grow.
'@param dataRows Long. The number of data rows wanted.
Public Sub GeoFixtureResizeTable(ByVal sh As Worksheet, ByVal Lo As ListObject, _
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


'@section Private
'===============================================================================

'@sub-title Write the header row of one geobase table.
'@param sh Worksheet. The fixture worksheet.
'@param tableName String. The table whose headers are wanted.
'@param startCol Long. The column the table starts in.
Private Sub WriteTableHeaders(ByVal sh As Worksheet, ByVal tableName As String, _
                              ByVal startCol As Long)
    Select Case tableName

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
End Sub

'@sub-title Fill the fixture tables with a small geobase.
'@details
'Three admin1 values, two admin2 under each and two admin3 under each of those,
'one admin4 per admin3, three health facilities and the five level translations
'in an EN column. The third admin1 value is the number 3.
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
        GeoFixtureWriteCell sh, loAdm1, adm1Idx + 1, 0, adm1Value

        For childIdx = 1 To 2
            adm2Row = adm2Row + 1
            adm2Value = "D" & adm2Row
            GeoFixtureWriteCell sh, loAdm2, adm2Row, 0, adm1Value
            GeoFixtureWriteCell sh, loAdm2, adm2Row, 1, adm2Value

            For grandIdx = 1 To 2
                adm3Row = adm3Row + 1
                adm3Value = "C" & adm3Row

                GeoFixtureWriteCell sh, loAdm3, adm3Row, 0, adm1Value
                GeoFixtureWriteCell sh, loAdm3, adm3Row, 1, adm2Value
                GeoFixtureWriteCell sh, loAdm3, adm3Row, 2, adm3Value

                GeoFixtureWriteCell sh, loAdm4, adm3Row, 0, adm1Value
                GeoFixtureWriteCell sh, loAdm4, adm3Row, 1, adm2Value
                GeoFixtureWriteCell sh, loAdm4, adm3Row, 2, adm3Value
                GeoFixtureWriteCell sh, loAdm4, adm3Row, 3, "V" & adm3Row
            Next grandIdx
        Next childIdx
    Next adm1Idx

    GeoFixtureResizeTable sh, loAdm1, 3
    GeoFixtureResizeTable sh, loAdm2, adm2Row
    GeoFixtureResizeTable sh, loAdm3, adm3Row
    GeoFixtureResizeTable sh, loAdm4, adm3Row

    'T_HF: hf_name, hf_pcode, hf_extra, adm3_name, adm2_name, adm1_name
    For childIdx = 1 To 3
        GeoFixtureWriteCell sh, loHF, childIdx, 0, "HF" & childIdx
        GeoFixtureWriteCell sh, loHF, childIdx, 1, "PC" & childIdx
        GeoFixtureWriteCell sh, loHF, childIdx, 2, "X" & childIdx
        GeoFixtureWriteCell sh, loHF, childIdx, 3, "C" & childIdx
        GeoFixtureWriteCell sh, loHF, childIdx, 4, "D1"
        GeoFixtureWriteCell sh, loHF, childIdx, 5, "P1"
    Next childIdx
    GeoFixtureResizeTable sh, loHF, 3

    GeoFixtureWriteCell sh, loNames, 1, 0, "adm1_name"
    GeoFixtureWriteCell sh, loNames, 1, 1, "Province"
    GeoFixtureWriteCell sh, loNames, 2, 0, "adm2_name"
    GeoFixtureWriteCell sh, loNames, 2, 1, "District"
    GeoFixtureWriteCell sh, loNames, 3, 0, "adm3_name"
    GeoFixtureWriteCell sh, loNames, 3, 1, "Commune"
    GeoFixtureWriteCell sh, loNames, 4, 0, "adm4_name"
    GeoFixtureWriteCell sh, loNames, 4, 1, "Village"
    GeoFixtureWriteCell sh, loNames, 5, 0, "hf_name"
    GeoFixtureWriteCell sh, loNames, 5, 1, "Health Facility"
    GeoFixtureResizeTable sh, loNames, 5

    geoStore.SetValue "RNG_GeoLangCode", "EN"
    geoStore.SetValue "RNG_GeoUpdated", "updated"
End Sub
