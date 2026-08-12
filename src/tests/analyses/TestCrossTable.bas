Attribute VB_Name = "TestCrossTable"
Attribute VB_Description = "Tests for CrossTable class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for CrossTable class")

'@description
'Validates CrossTable, the layout engine for all six analysis table scopes. It
'writes rows, columns, headers and about forty named ranges per table, and five
'other classes address the table only through those names, so most of what this
'suite asserts is the presence, the position and the size of a named range.
'
'THE FIXTURE IS A REAL LISTOBJECT
'-------------------------------------------------------------------------------
'TableSpecs reads the analysis scope from the name of the ListObject the
'specification row sits in, so every fixture wraps its header row and its data
'rows in a ListObject carrying the name the setup workbook uses
'(Tab_Univariate_Analysis and the other five). A row on a bare worksheet has no
'scope to resolve and TableSpecs.Create raises ErrorUnexpectedState.
'
'THE OUTPUT SHEET IS RESET BY NAME, NOT ONLY BY VALUE
'-------------------------------------------------------------------------------
'CrossTable creates its names with Cell.Name, which produces workbook-scoped
'names. Those outlive a Cells.Clear, and the class asks RangeExists whether it
'has already built a piece of structure, so a name left behind by an earlier
'test changes the answer for the next one. ResetOutputSheet deletes them, and it
'matches both spellings of RefersTo: a sheet name with no space is written
'without quotes and the sweep in ClearWorksheet looks only for the quoted form.
'@depends CrossTable, TableSpecs, LLdictionary, LLVariables, LLFormat,
'  TranslationObject, LinelistDataStub, BetterArray, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "CrossTableFixture"
Private Const OUTPUT_SHEET As String = "CrossTableOutput"
Private Const DICT_SHEET As String = "CrossTableDict"
Private Const TRANS_SHEET As String = "CrossTableTrans"
Private Const TRANS_TABLE As String = "T_CrossTableTranslation"
Private Const FORMAT_SHEET As String = "CrossTableFormat"

' The header row of every fixture table. Data rows start immediately below, so
' data row N sits at HEADER_ROW + N and TableId reads "<prefix>_tabN".
Private Const HEADER_ROW As Long = 5

' The blocks the two worksheets are reset over between tests. Both are bigger
' than anything this suite writes, and both exist so the reset does not touch
' every cell of a sheet. See OutputSheet for what that cost.
Private Const OUTPUT_BLOCK As String = "A1:AN200"
Private Const FIXTURE_BLOCK As String = "A1:T30"

' The analysis ListObject names, spelled the way the setup workbook spells them.
Private Const TABLE_GLOBAL_SUMMARY As String = "Tab_Global_Summary"
Private Const TABLE_UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const TABLE_BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TABLE_TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const TABLE_SPATIAL As String = "Tab_Spatial_Analysis"
Private Const TABLE_SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"

' Variables the shared dictionary fixture holds, and the two geo rows this
' suite appends to it. SpatialTableScopes probes hf_<var> then adm1_<var>.
Private Const ROW_CHOICE_VARIABLE As String = "choi_v1"
Private Const COL_CHOICE_VARIABLE As String = "choi_ord_v1"
Private Const DATE_VARIABLE As String = "date_v1"
Private Const GEO_VARIABLE As String = "zone"
Private Const GEO_PREFIXED_VARIABLE As String = "adm1_zone"
Private Const HF_VARIABLE As String = "center"
Private Const HF_PREFIXED_VARIABLE As String = "hf_center"

' The column CrossTable lays a table out from, and the fixed height of the
' temporal grid. Both are private constants of the class under test.
Private Const STANDARD_START_COL As Long = 3
Private Const NB_ROWS_TIME_SERIES As Long = 56

Private Assert As CustomTest
Private dict As LLdictionary
Private trans As TranslationObject
Private lData As LinelistDataStub

'@section Fixture headers
'===============================================================================
'@description The six analysis header rows, spelled as the setup workbook has
'them.

'@sub-title Header of Tab_Global_Summary, 3 columns.
Private Function GlobalSummaryHeader() As Variant
    GlobalSummaryHeader = Array("Summary label", "Summary function", "Format")
End Function

'@sub-title Header of Tab_Univariate_Analysis, 10 columns.
Private Function UnivariateHeader() As Variant
    UnivariateHeader = Array( _
        "Section", "Table title", "Group by variable (row)", "Add missing data", _
        "Summary function", "Summary label", "Format", "Add percentage", _
        "Add graph", "Flip coordinates")
End Function

'@sub-title Header of Tab_Bivariate_Analysis, 11 columns.
Private Function BivariateHeader() As Variant
    BivariateHeader = Array( _
        "Section", "Table title", "Group by variable (row)", _
        "Group by variable (column)", "Add missing data", "Summary function", _
        "Summary label", "Format", "Add percentage", "Add graph", _
        "Flip coordinates")
End Function

'@sub-title Header of Tab_TimeSeries_Analysis, 12 columns.
Private Function TimeSeriesHeader() As Variant
    TimeSeriesHeader = Array( _
        "Series ID", "Section", "Time variable (row)", _
        "Group by variable (column)", "Title (header)", "Add missing data", _
        "Summary function", "Summary label", "Format", "Add percentage", _
        "Add total", "Table order")
End Function

'@sub-title Header of Tab_Spatial_Analysis, 12 columns.
Private Function SpatialHeader() As Variant
    SpatialHeader = Array( _
        "Section", "Table title", "Geo/HF variable (row)", "N geo max", _
        "Group by variable (column)", "Add missing data", "Summary function", _
        "Summary label", "Format", "Add percentage", "Add graph", _
        "Flip coordinates")
End Function

'@sub-title Header of Tab_SpatioTemporal_Analysis, 10 columns.
Private Function SpatioTemporalHeader() As Variant
    SpatioTemporalHeader = Array( _
        "Section (select)", "Time variable (row)", "Geo/HF variable (column)", _
        "N geo max", "Title (header)", "Spatial type", "Summary function", _
        "Summary label", "Format", "Add graph")
End Function

'@section Fixture data rows
'===============================================================================

'@sub-title One global summary row.
'@param label String. The summary label written into the row cell.
Private Function GlobalSummaryRows(ByVal label As String) As Variant
    GlobalSummaryRows = Array(Array(label, "count", ""))
End Function

'@sub-title One univariate row with the flags the caller passes.
Private Function UnivariateRows(ByVal rowVar As String, _
                                ByVal missing As String, _
                                ByVal percentage As String) As Variant
    UnivariateRows = Array( _
        Array("S1", "A univariate table", rowVar, missing, "", "Cases", "", _
              percentage, "no", "no"))
End Function

'@sub-title One bivariate row with the flags the caller passes.
Private Function BivariateRows(ByVal rowVar As String, _
                               ByVal colVar As String, _
                               ByVal missing As String, _
                               ByVal percentage As String) As Variant
    BivariateRows = Array( _
        Array("S1", "A bivariate table", rowVar, colVar, missing, "", "Cases", _
              "", percentage, "no", "no"))
End Function

'@sub-title One time series row with no column variable.
'@details
'AddColumns builds the composite header from the summary label, the time
'variable and the column variable, and drops the last of the three when the
'row names no column variable. That arm needs a row with the column cell empty.
Private Function TimeSeriesRowsNoColumn(ByVal total As String) As Variant
    TimeSeriesRowsNoColumn = Array( _
        Array("Series 1", "S1", DATE_VARIABLE, "", _
              "First table", "no", "", "Cases", "", "", total, "1"))
End Function

'@sub-title Two global summary rows sharing one banner.
Private Function GlobalSummaryTwoRows() As Variant
    GlobalSummaryTwoRows = Array( _
        Array("First summary", "count", ""), _
        Array("Second summary", "count", ""))
End Function

'@sub-title Two time series rows in one section.
'@details
'Row 1 opens section S1 and row 2 continues it, so row 2 is the non-new-section
'temporal table whose start row is inherited from row 1.
Private Function TimeSeriesRows(ByVal total As String) As Variant
    TimeSeriesRows = Array( _
        Array("Series 1", "S1", DATE_VARIABLE, COL_CHOICE_VARIABLE, _
              "First table", "no", "", "Cases", "", "", total, "1"), _
        Array("Series 2", "S1", DATE_VARIABLE, COL_CHOICE_VARIABLE, _
              "Second table", "no", "", "Cases", "", "", total, "2"))
End Function

'@sub-title One spatial row with the row variable and geo count the caller passes.
Private Function SpatialRows(ByVal rowVar As String, _
                             ByVal geoMax As String) As Variant
    SpatialRows = Array( _
        Array("S1", "A spatial table", rowVar, geoMax, "", "no", "", "Cases", _
              "", "", "no", "no"))
End Function

'@sub-title One spatio-temporal row.
Private Function SpatioTemporalRows(ByVal colVar As String, _
                                    ByVal geoMax As String, _
                                    ByVal spatialType As String) As Variant
    SpatioTemporalRows = Array( _
        Array("S1", DATE_VARIABLE, colVar, geoMax, "A spatio-temporal table", _
              spatialType, "", "Cases", "", "no"))
End Function

'@section Fixture helpers
'===============================================================================

'@sub-title Drop every ListObject on the fixture sheet.
'@details
'A ListObject name is unique across the workbook, so the table the previous
'test built has to go before the next one takes the same name. Unlist turns the
'table back into an ordinary range and frees the name.
'@param sh Worksheet. The fixture worksheet.
Private Sub RemoveFixtureTables(ByVal sh As Worksheet)
    Dim idx As Long

    For idx = sh.ListObjects.Count To 1 Step -1
        sh.ListObjects(idx).Unlist
    Next idx
End Sub

'@sub-title Build a fixture ListObject with a header row and data rows.
'@param tableName String. The ListObject name, one of the six Tab_ names.
'@param headerRow Variant. A one-dimensional array of column names.
'@param dataRows Variant. An array of row arrays, each as wide as headerRow.
Private Sub BuildFixture(ByVal tableName As String, _
                         ByVal headerRow As Variant, _
                         ByVal dataRows As Variant)
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim tableRng As Range
    Dim rowCount As Long
    Dim colCount As Long

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=False, visibility:=xlSheetHidden)
    RemoveFixtureTables sh
    sh.Range(FIXTURE_BLOCK).Clear

    colCount = UBound(headerRow) - LBound(headerRow) + 1
    rowCount = UBound(dataRows) - LBound(dataRows) + 1

    WriteMatrix sh.Cells(HEADER_ROW, 1), RowsToMatrix(Array(headerRow))
    WriteMatrix sh.Cells(HEADER_ROW + 1, 1), RowsToMatrix(dataRows)

    Set tableRng = sh.Range(sh.Cells(HEADER_ROW, 1), _
                            sh.Cells(HEADER_ROW + rowCount, colCount))
    Set lo = sh.ListObjects.Add(xlSrcRange, tableRng, , xlYes)
    lo.Name = tableName
End Sub

'@sub-title Return the fixture header range.
Private Function FixtureHeaderRange() As Range
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    Set FixtureHeaderRange = sh.ListObjects(1).HeaderRowRange
End Function

'@sub-title Return a fixture data row range by 1-based index.
Private Function FixtureDataRange(ByVal dataRowIndex As Long) As Range
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    Set FixtureDataRange = sh.ListObjects(1).ListRows(dataRowIndex).Range
End Function

'@sub-title Create a TableSpecs from a fixture data row index.
Private Function CreateSpecs(ByVal dataRowIndex As Long) As TableSpecs
    Set CreateSpecs = TableSpecs.Create( _
        FixtureHeaderRange(), _
        FixtureDataRange(dataRowIndex), _
        dict)
End Function

'@sub-title Delete every workbook-scoped name that points at one worksheet.
'@details
'Excel writes RefersTo as ='Sheet name'!$C$10 when the sheet name needs quoting
'and as =Sheetname!$C$10 when it does not. Both spellings are matched here. The
'worksheet-scoped names go too, so the sheet starts each test with none.
'@param sh Worksheet. The worksheet whose names are removed.
Private Sub RemoveSheetNames(ByVal sh As Worksheet)
    Dim idx As Long
    Dim nm As Name
    Dim target As String

    For idx = sh.Names.Count To 1 Step -1
        sh.Names(idx).Delete
    Next idx

    For idx = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(idx)
        target = vbNullString
        On Error Resume Next
        target = nm.RefersTo
        On Error GoTo 0
        If (InStr(1, target, "'" & sh.Name & "'!", vbTextCompare) > 0) Or _
           (InStr(1, target, "=" & sh.Name & "!", vbTextCompare) > 0) Then
            nm.Delete
        End If
    Next idx
End Sub

'@sub-title Return an output worksheet with no content, no names and no merges.
'@details
'Format merges label cells and hides the total column, and neither survives a
'Cells.Clear, so both are undone here. Every build in this suite therefore
'starts from the same blank sheet and the start row is the same number.
'
'The reset is bounded to OUTPUT_BLOCK, and clearSheet is False for the same
'reason. Running UnMerge, the two unhide writes and Clear over every cell of
'the sheet cost about seven seconds a test, forty-six times, which took the
'whole run past the runner's ten-minute cap. The block is wider and taller than
'any table this suite builds: the tallest is a temporal grid of 56 rows under
'seventeen inserted header rows, and the rightmost cells are the shared unit
'lists in columns 26 and 27.
Private Function OutputSheet() As Worksheet
    Dim sh As Worksheet
    Dim block As Range

    Set sh = EnsureWorksheet(OUTPUT_SHEET, clearSheet:=False, visibility:=xlSheetHidden)
    RemoveSheetNames sh
    Set block = sh.Range(OUTPUT_BLOCK)
    block.UnMerge
    block.EntireRow.Hidden = False
    block.EntireColumn.Hidden = False
    block.Clear
    Set OutputSheet = sh
End Function

'@sub-title Build the translation table this suite reads its labels from.
'@details
'CrossTable asks for twenty-two message tags. A tag that is missing answers
'with its own name, so every tag is seeded here and an assertion can name the
'text it expects.
Private Function BuildTranslator() As TranslationObject
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim translationRows As Variant

    Set sh = EnsureWorksheet(TRANS_SHEET, clearSheet:=True, visibility:=xlSheetHidden)
    RemoveFixtureTables sh

    translationRows = Array( _
        Array("MSG_NA", "Missing"), _
        Array("MSG_Total", "Total"), _
        Array("MSG_Percent", "%"), _
        Array("MSG_AllData", "All data"), _
        Array("MSG_FilteredData", "Filtered data"), _
        Array("MSG_GlobalSummary", "Global summary"), _
        Array("MSG_Period", "Period"), _
        Array("MSG_InputGeoLevels", "Geo levels"), _
        Array("MSG_StartDate", "Start date"), _
        Array("MSG_TimeUnit", "Time unit"), _
        Array("MSG_EndDate", "End date"), _
        Array("MSG_Day", "Day"), _
        Array("MSG_Week", "Week"), _
        Array("MSG_Month", "Month"), _
        Array("MSG_Quarter", "Quarter"), _
        Array("MSG_Year", "Year"), _
        Array("MSG_NoDevide", "No divide"), _
        Array("MSG_Devide", "Divide"), _
        Array("MSG_SelectAdmin", "Select admin"), _
        Array("MSG_SelectPOPFACT", "Select factor"), _
        Array("MSG_MultiplyBy", "Multiply by"), _
        Array("MSG_HF", "Health facility"))

    WriteMatrix sh.Cells(1, 1), RowsToMatrix(Array(Array("tag", "English")))
    WriteMatrix sh.Cells(2, 1), RowsToMatrix(translationRows)

    Set lo = sh.ListObjects.Add(xlSrcRange, sh.Range("A1").CurrentRegion, , xlYes)
    lo.Name = TRANS_TABLE

    Set BuildTranslator = TranslationObject.Create(lo, "English")
End Function

'@sub-title Check whether a named range resolves on one worksheet.
'@param sh Worksheet. The worksheet to resolve against.
'@param rngName String. The name to look up.
Private Function RangeExistsOnSheet(ByVal sh As Worksheet, _
                                    ByVal rngName As String) As Boolean
    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExistsOnSheet = (Not rng Is Nothing)
End Function

'@sub-title Return the address of a named range, or an empty string.
'@param sh Worksheet. The worksheet to resolve against.
'@param rngName String. The name to look up.
Private Function NamedAddress(ByVal sh As Worksheet, _
                              ByVal rngName As String) As String
    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    If Not rng Is Nothing Then NamedAddress = rng.Address(False, False)
End Function

'@sub-title Build one cross-table from a fixture data row and return it.
'@param sh Worksheet. The output worksheet.
'@param dataRowIndex Long. The 1-based fixture data row.
Private Function BuildTable(ByVal sh As Worksheet, _
                            ByVal dataRowIndex As Long) As CrossTable
    Dim tabl As CrossTable

    Set tabl = CrossTable.Create(CreateSpecs(dataRowIndex), sh, lData)
    tabl.Build
    Set BuildTable = tabl
End Function

'@sub-title Create one cross-table from a fixture data row without building it.
'@details
'Build runs AddHeader, AddRows, AddColumns and NameRanges in that order, and the
'stage tests call those one at a time so a raise names the stage it came from.
'This hands back the unbuilt table those tests start from.
'@param sh Worksheet. The output worksheet.
'@param dataRowIndex Long. The 1-based fixture data row.
Private Function NewTable(ByVal sh As Worksheet, _
                          ByVal dataRowIndex As Long) As CrossTable
    Set NewTable = CrossTable.Create(CreateSpecs(dataRowIndex), sh, lData)
End Function

'@sub-title Read one cell of the output sheet as text.
'@param sh Worksheet. The worksheet to read.
'@param rw Long. The row to read.
'@param col Long. The column to read.
Private Function CellText(ByVal sh As Worksheet, _
                          ByVal rw As Long, _
                          ByVal col As Long) As String
    CellText = CStr(sh.Cells(rw, col).Value)
End Function

'@sub-title Read one dictionary column of one variable.
'@details
'The stage tests assert the labels CrossTable writes, and those labels come out
'of the dictionary rather than out of the specification row, so the expected
'value is read from the same place the class reads it.
'@param varName String. The variable to look up.
'@param colName String. The dictionary column to read.
Private Function DictLabel(ByVal varName As String, _
                           ByVal colName As String) As String
    DictLabel = LLVariables.Create(dict).Value(colName:=colName, varName:=varName)
End Function

'@sub-title Join one row of the output sheet into a single string.
'@details
'A layout assertion that names every cell it read localizes itself, and the
'headless runner keeps only what an assertion message carries.
'@param sh Worksheet. The worksheet to read.
'@param rw Long. The row to read.
'@param firstCol Long. The first column to read.
'@param lastCol Long. The last column to read.
Private Function RowText(ByVal sh As Worksheet, _
                         ByVal rw As Long, _
                         ByVal firstCol As Long, _
                         ByVal lastCol As Long) As String
    Dim col As Long
    Dim parts As String

    For col = firstCol To lastCol
        If col > firstCol Then parts = parts & "|"
        parts = parts & CellText(sh, rw, col)
    Next col

    RowText = parts
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Set up the dictionary, the translator and the linelist data stub.
'@details
'Builds the shared dictionary fixture, appends the two geo variables the
'fixture has no equivalent of, then seeds the stub with the translator and with
'the categories of the two choice variables the tables group by. This routine
'is Public because the harness calls it by name through Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    Dim sh As Worksheet
    Dim appendRow As Long

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCrossTable"

    PrepareDictionaryFixture DICT_SHEET

    ' adm1_zone gives the spatial scope a "geo" answer to find and hf_center
    ' gives it an "hf" answer. AppendGeoLines writes rows of this shape during
    ' a real Prepare, and the fixture is a raw dictionary.
    Set sh = ThisWorkbook.Worksheets(DICT_SHEET)
    appendRow = 1 + DictionaryFixtureRowCount() + 1
    sh.Cells(appendRow, 1).Value = GEO_PREFIXED_VARIABLE
    sh.Cells(appendRow, 2).Value = "Zone"
    sh.Cells(appendRow, 10).Value = "Zones"
    sh.Cells(appendRow + 1, 1).Value = HF_PREFIXED_VARIABLE
    sh.Cells(appendRow + 1, 2).Value = "Health centre"

    Set dict = LLdictionary.Create(sh, 1, 1)
    Set trans = BuildTranslator()

    Set lData = New LinelistDataStub
    lData.SetTransObject trans
    lData.SetCategories ROW_CHOICE_VARIABLE, BetterArrayFromList("A", "B", "C")
    lData.SetCategories COL_CHOICE_VARIABLE, BetterArrayFromList("X", "Y")
End Sub

'@sub-title Print results and tear down the fixtures.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet FIXTURE_SHEET
    DeleteWorksheet OUTPUT_SHEET
    DeleteWorksheet DICT_SHEET
    DeleteWorksheet TRANS_SHEET
    DeleteWorksheet FORMAT_SHEET
    RestoreApp
    Set dict = Nothing
    Set trans = Nothing
    Set lData = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updating before each test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush assert state after each test.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory
'===============================================================================

'@sub-title Verify Create raises InvalidArgument when the specifications are Nothing.
'@TestMethod("CrossTable")
Public Sub TestCreateRejectsNothingSpecs()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateRejectsNothingSpecs"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim errNumber As Long
    Dim sh As Worksheet

    Set sh = OutputSheet()

    On Error Resume Next
    Set tabl = CrossTable.Create(Nothing, sh, lData)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                    "Nothing specifications should raise InvalidArgument"
    Assert.IsTrue (tabl Is Nothing), _
                  "Create with Nothing specifications should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingSpecs", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises InvalidArgument when the worksheet is Nothing.
'@TestMethod("CrossTable")
Public Sub TestCreateRejectsNothingWorksheet()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateRejectsNothingWorksheet"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim specs As TableSpecs
    Dim errNumber As Long

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set specs = CreateSpecs(1)

    On Error Resume Next
    Set tabl = CrossTable.Create(specs, Nothing, lData)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                    "A Nothing output worksheet should raise InvalidArgument"
    Assert.IsTrue (tabl Is Nothing), _
                  "Create with a Nothing worksheet should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingWorksheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises InvalidArgument when the linelist data is Nothing.
'@TestMethod("CrossTable")
Public Sub TestCreateRejectsNothingLinelistData()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateRejectsNothingLinelistData"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim specs As TableSpecs
    Dim errNumber As Long
    Dim sh As Worksheet

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set specs = CreateSpecs(1)
    Set sh = OutputSheet()

    On Error Resume Next
    Set tabl = CrossTable.Create(specs, sh, Nothing)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                    "Nothing linelist data should raise InvalidArgument"
    Assert.IsTrue (tabl Is Nothing), _
                  "Create with Nothing linelist data should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingLinelistData", Err.Number, Err.Description
End Sub

'@sub-title Verify Create accepts a stub in place of the linelist specifications.
'@details
'The parameter is typed Object, so the object handed over needs the two members
'the class reads and nothing else. This is what keeps the LinelistSpecs closure
'out of the test workbook, so it is asserted rather than assumed.
'@TestMethod("CrossTable")
Public Sub TestCreateAcceptsAnyObjectAsLinelistData()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateAcceptsAnyObjectAsLinelistData"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()

    Set tabl = CrossTable.Create(CreateSpecs(1), sh, lData)

    Assert.IsTrue (Not tabl Is Nothing), _
                  "Create with valid arguments should hand back a table"
    Assert.AreEqual "LinelistDataStub", TypeName(tabl.LinelistData), _
                    "LinelistData should hand back the object passed at creation"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateAcceptsAnyObjectAsLinelistData", Err.Number, Err.Description
End Sub

'@sub-title Verify the instance is sealed against a second setup write.
'@TestMethod("CrossTable")
Public Sub TestSetupWriteAfterCreateRaises()
    CustomTestSetTitles Assert, "CrossTable", "TestSetupWriteAfterCreateRaises"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim errNumber As Long

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = CrossTable.Create(CreateSpecs(1), sh, lData)

    On Error Resume Next
    Set tabl.Wksh = sh
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.SomethingWentWrong, errNumber, _
                    "Assigning Wksh after Create should raise SomethingWentWrong"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetupWriteAfterCreateRaises", Err.Number, Err.Description
End Sub

'@section Properties
'===============================================================================

'@sub-title Verify Specifications and Wksh hand back what Create was given.
'@TestMethod("CrossTable")
Public Sub TestSpecificationsAndWkshPassThrough()
    CustomTestSetTitles Assert, "CrossTable", "TestSpecificationsAndWkshPassThrough"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim specs As TableSpecs
    Dim sh As Worksheet

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set specs = CreateSpecs(1)
    Set sh = OutputSheet()
    Set tabl = CrossTable.Create(specs, sh, lData)

    Assert.AreEqual specs.TableId, tabl.Specifications.TableId, _
                    "Specifications should hand back the specs passed at creation"
    Assert.AreEqual sh.Name, tabl.Wksh.Name, _
                    "Wksh should hand back the worksheet passed at creation"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpecificationsAndWkshPassThrough", Err.Number, Err.Description
End Sub

'@sub-title Verify Translations reads the translation object off the linelist data.
'@TestMethod("CrossTable")
Public Sub TestTranslationsComeFromLinelistData()
    CustomTestSetTitles Assert, "CrossTable", "TestTranslationsComeFromLinelistData"
    On Error GoTo TestFail

    Dim tabl As CrossTable

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set tabl = CrossTable.Create(CreateSpecs(1), OutputSheet(), lData)

    Assert.IsTrue (Not tabl.Translations Is Nothing), _
                  "Translations should resolve through the linelist data object"
    Assert.AreEqual "Total", tabl.Translations.TranslatedValue("MSG_Total"), _
                    "Translations should read the table this suite seeded"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTranslationsComeFromLinelistData", Err.Number, Err.Description
End Sub

'@sub-title Verify StartRow creates its marker and then reuses it.
'@details
'The first read scans column C upward and names the cell it lands on. Every
'later read is a lookup of that name, so two reads have to agree.
'@TestMethod("CrossTable")
Public Sub TestStartRowIsNamedAndStable()
    CustomTestSetTitles Assert, "CrossTable", "TestStartRowIsNamedAndStable"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim firstRead As Long

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = CrossTable.Create(CreateSpecs(1), sh, lData)

    firstRead = tabl.StartRow

    Assert.IsTrue firstRead > 1, _
                  "A first table should start below the first row"
    Assert.IsTrue RangeExistsOnSheet(sh, "STARTROW_" & tabl.Specifications.TableId), _
                  "Reading StartRow should create the STARTROW_ marker"
    Assert.AreEqual firstRead, tabl.StartRow, _
                    "A second read of StartRow should answer the same row"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestStartRowIsNamedAndStable", Err.Number, Err.Description
End Sub

'@section Global summary
'===============================================================================

'@sub-title Verify a global summary build creates the two shared markers.
'@TestMethod("CrossTable")
Public Sub TestGlobalSummaryCreatesSharedMarkers()
    CustomTestSetTitles Assert, "CrossTable", "TestGlobalSummaryCreatesSharedMarkers"
    On Error GoTo TestFail

    Dim sh As Worksheet

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 GlobalSummaryRows("Total cases")
    Set sh = OutputSheet()
    BuildTable sh, 1

    Assert.IsTrue RangeExistsOnSheet(sh, "ROWGS_SET"), _
                  "A global summary build should create ROWGS_SET"
    Assert.IsTrue RangeExistsOnSheet(sh, "COLGS_SET"), _
                  "A global summary build should create COLGS_SET"
    Assert.AreEqual "Global summary", sh.Range("ROWGS_SET").Value, _
                    "ROWGS_SET should carry the translated global summary label"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGlobalSummaryCreatesSharedMarkers", Err.Number, Err.Description
End Sub

'@sub-title Verify a global summary row writes its label and two columns.
'@TestMethod("CrossTable")
Public Sub TestGlobalSummaryWritesLabelAndTwoColumns()
    CustomTestSetTitles Assert, "CrossTable", "TestGlobalSummaryWritesLabelAndTwoColumns"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 GlobalSummaryRows("Total cases")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Assert.AreEqual CLng(2), tabl.NumberOfColumns, _
                    "A global summary table should report two columns"
    Assert.AreEqual "Total cases", sh.Cells(tabl.StartRow, STANDARD_START_COL).Value, _
                    "The summary label should be written at the start of the row"
    Assert.AreEqual "All data", sh.Range("COLGS_SET").Value, _
                    "The first shared column header should be the all-data label"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGlobalSummaryWritesLabelAndTwoColumns", Err.Number, Err.Description
End Sub

'@sub-title Verify NameRanges writes nothing for a global summary table.
'@details
'A global summary row has no row categories and no column structure beyond the
'two shared cells, so NameRanges leaves before it names anything.
'@TestMethod("CrossTable")
Public Sub TestGlobalSummaryNamesNoTableRanges()
    CustomTestSetTitles Assert, "CrossTable", "TestGlobalSummaryNamesNoTableRanges"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 GlobalSummaryRows("Total cases")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue Not RangeExistsOnSheet(sh, "ENDTABLE_" & tabId), _
                  "A global summary table should name no end-table marker"
    Assert.IsTrue Not RangeExistsOnSheet(sh, "INTERIOR_VALUES_" & tabId), _
                  "A global summary table should name no interior values"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGlobalSummaryNamesNoTableRanges", Err.Number, Err.Description
End Sub

'@sub-title Verify RowsCategoriesRange refuses a global summary table.
'@TestMethod("CrossTable")
Public Sub TestRowsCategoriesRangeRejectsGlobalSummary()
    CustomTestSetTitles Assert, "CrossTable", "TestRowsCategoriesRangeRejectsGlobalSummary"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim rng As Range
    Dim errNumber As Long

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 GlobalSummaryRows("Total cases")
    Set tabl = BuildTable(OutputSheet(), 1)

    On Error Resume Next
    Set rng = tabl.RowsCategoriesRange()
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.ErrorUnexpectedState, errNumber, _
                    "A global summary table should refuse to answer row categories"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRowsCategoriesRangeRejectsGlobalSummary", Err.Number, Err.Description
End Sub

'@section Univariate
'===============================================================================

'@sub-title Verify a univariate build creates the ranges its readers need.
'@details
'Five other classes address the table through these names, so each one is
'asserted by itself and a failure says which name is missing.
'@TestMethod("CrossTable")
Public Sub TestUnivariateCreatesNamedRanges()
    CustomTestSetTitles Assert, "CrossTable", "TestUnivariateCreatesNamedRanges"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "TITLE_" & tabId), _
                  "A univariate build should name the title cell"
    Assert.IsTrue RangeExistsOnSheet(sh, "SECTION_" & tabId), _
                  "A first table in a section should name the section cell"
    Assert.IsTrue RangeExistsOnSheet(sh, "ENDTABLE_" & tabId), _
                  "A univariate build should name the end-table marker"
    Assert.IsTrue RangeExistsOnSheet(sh, "ROW_CATEGORIES_" & tabId), _
                  "A univariate build should name the row categories"
    Assert.IsTrue RangeExistsOnSheet(sh, "LABEL_ROW_CATEGORIES_" & tabId), _
                  "A univariate build should name the row categories label"
    Assert.IsTrue RangeExistsOnSheet(sh, "VALUES_COL_1_" & tabId), _
                  "A univariate build should name its one value column"
    Assert.IsTrue RangeExistsOnSheet(sh, "LABEL_COL_1_" & tabId), _
                  "A univariate build should name the value column label"
    Assert.IsTrue RangeExistsOnSheet(sh, "INTERIOR_VALUES_" & tabId), _
                  "A univariate build should name the interior values"
    Assert.IsTrue RangeExistsOnSheet(sh, "COLUMN_CATEGORIES_" & tabId), _
                  "A univariate build should name the column categories"
    Assert.IsTrue RangeExistsOnSheet(sh, "TOTAL_ROW_" & tabId), _
                  "A univariate build should name the total row"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateCreatesNamedRanges", Err.Number, Err.Description
End Sub

'@sub-title Verify the row categories are the seeded categories plus a total.
'@TestMethod("CrossTable")
Public Sub TestUnivariateRowsAreCategoriesPlusTotal()
    CustomTestSetTitles Assert, "CrossTable", "TestUnivariateRowsAreCategoriesPlusTotal"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim startRw As Long

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    startRw = tabl.StartRow

    Assert.AreEqual "A", sh.Cells(startRw + 1, STANDARD_START_COL).Value, _
                    "The first row category should be the first seeded category"
    Assert.AreEqual "C", sh.Cells(startRw + 3, STANDARD_START_COL).Value, _
                    "The third row category should be the third seeded category"
    Assert.AreEqual "Total", sh.Cells(startRw + 4, STANDARD_START_COL).Value, _
                    "A total row should follow the categories"
    Assert.AreEqual startRw + 4, tabl.EndRow, _
                    "EndRow should be the last written row"
    Assert.AreEqual CLng(3), tabl.RowsCategoriesRange(onlyCategories:=True, includeHeaders:=False).Rows.Count, _
                    "Three seeded categories should give a three-row category range"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateRowsAreCategoriesPlusTotal", Err.Number, Err.Description
End Sub

'@sub-title Verify a missing row is added when the row asks for one.
'@TestMethod("CrossTable")
Public Sub TestUnivariateWithMissingAddsAMissingRow()
    CustomTestSetTitles Assert, "CrossTable", "TestUnivariateWithMissingAddsAMissingRow"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim startRw As Long

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "yes", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    startRw = tabl.StartRow

    Assert.AreEqual "Missing", sh.Cells(startRw + 4, STANDARD_START_COL).Value, _
                    "A missing row should follow the categories"
    Assert.AreEqual "Total", sh.Cells(startRw + 5, STANDARD_START_COL).Value, _
                    "The total row should follow the missing row"
    Assert.IsTrue RangeExistsOnSheet(sh, "MISSING_ROW_" & tabl.Specifications.TableId), _
                  "A missing row should be named"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateWithMissingAddsAMissingRow", Err.Number, Err.Description
End Sub

'@sub-title Verify a univariate table reports one data column.
'@TestMethod("CrossTable")
Public Sub TestUnivariateReportsOneDataColumn()
    CustomTestSetTitles Assert, "CrossTable", "TestUnivariateReportsOneDataColumn"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Assert.AreEqual CLng(1), tabl.NumberOfColumns, _
                    "A univariate table has one data column"
    Assert.AreEqual CLng(STANDARD_START_COL + 1), tabl.EndColumn, _
                    "One data column puts EndColumn one to the right of the start"
    Assert.AreEqual "Cases", sh.Cells(tabl.StartRow, STANDARD_START_COL + 1).Value, _
                    "The summary label should head the value column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateReportsOneDataColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify a percentage column is added and named.
'@TestMethod("CrossTable")
Public Sub TestUnivariateWithPercentageAddsAColumn()
    CustomTestSetTitles Assert, "CrossTable", "TestUnivariateWithPercentageAddsAColumn"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "yes")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.AreEqual CLng(STANDARD_START_COL + 2), tabl.EndColumn, _
                    "A percentage column widens the table by one"
    Assert.IsTrue RangeExistsOnSheet(sh, "PERC_COL_1_" & tabId), _
                  "The percentage column should be named"
    Assert.IsTrue RangeExistsOnSheet(sh, "PERC_LABEL_COL_" & tabId), _
                  "The percentage column label should be named"
    Assert.AreEqual "%", sh.Cells(tabl.StartRow, STANDARD_START_COL + 2).Value, _
                    "The percentage column should carry the translated percent label"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateWithPercentageAddsAColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify the interior values span the value column exactly.
'@TestMethod("CrossTable")
Public Sub TestUnivariateInteriorValuesSpanTheValueColumn()
    CustomTestSetTitles Assert, "CrossTable", "TestUnivariateInteriorValuesSpanTheValueColumn"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.AreEqual NamedAddress(sh, "VALUES_COL_1_" & tabId), _
                    NamedAddress(sh, "INTERIOR_VALUES_" & tabId), _
                    "One value column makes the interior values that same column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateInteriorValuesSpanTheValueColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify RowRange finds a category row and spans the table width.
'@TestMethod("CrossTable")
Public Sub TestRowRangeFindsACategoryRow()
    CustomTestSetTitles Assert, "CrossTable", "TestRowRangeFindsACategoryRow"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim rng As Range

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set tabl = BuildTable(OutputSheet(), 1)

    Set rng = tabl.RowRange("B")

    Assert.IsTrue (Not rng Is Nothing), _
                  "RowRange should find a seeded category"
    Assert.AreEqual CLng(STANDARD_START_COL), CLng(rng.Column), _
                    "RowRange with headers should start at the row label column"
    Assert.AreEqual tabl.EndColumn, CLng(rng.Column + rng.Columns.Count - 1), _
                    "RowRange should reach the last column of the table"
    Assert.IsTrue (tabl.RowRange("not a category") Is Nothing), _
                  "RowRange should hand back nothing for a label it cannot find"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRowRangeFindsACategoryRow", Err.Number, Err.Description
End Sub

'@sub-title Verify HeaderRange covers the column header cells.
'@TestMethod("CrossTable")
Public Sub TestHeaderRangeCoversTheColumnHeaders()
    CustomTestSetTitles Assert, "CrossTable", "TestHeaderRangeCoversTheColumnHeaders"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim rng As Range

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "no")
    Set tabl = BuildTable(OutputSheet(), 1)

    Set rng = tabl.HeaderRange()

    Assert.AreEqual CLng(1), CLng(rng.Rows.Count), _
                    "The header range is one row tall"
    Assert.AreEqual CLng(STANDARD_START_COL + 1), CLng(rng.Column), _
                    "The header range starts one column right of the row labels"
    Assert.AreEqual tabl.EndColumn, CLng(rng.Column + rng.Columns.Count - 1), _
                    "The header range reaches the last column of the table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHeaderRangeCoversTheColumnHeaders", Err.Number, Err.Description
End Sub

'@section Bivariate
'===============================================================================

'@sub-title Verify a bivariate build names one value column per category.
'@TestMethod("CrossTable")
Public Sub TestBivariateNamesOneColumnPerCategory()
    CustomTestSetTitles Assert, "CrossTable", "TestBivariateNamesOneColumnPerCategory"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.AreEqual CLng(2), tabl.NumberOfColumns, _
                    "Two seeded column categories give two data columns"
    Assert.IsTrue RangeExistsOnSheet(sh, "VALUES_COL_1_" & tabId), _
                  "The first value column should be named"
    Assert.IsTrue RangeExistsOnSheet(sh, "VALUES_COL_2_" & tabId), _
                  "The second value column should be named"
    Assert.IsTrue Not RangeExistsOnSheet(sh, "VALUES_COL_3_" & tabId), _
                  "No third value column should be named"
    Assert.AreEqual "X", sh.Cells(tabl.StartRow, STANDARD_START_COL + 1).Value, _
                    "The first column header should be the first seeded category"
    Assert.AreEqual "Total", sh.Cells(tabl.StartRow, STANDARD_START_COL + 3).Value, _
                    "A total column should follow the categories"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariateNamesOneColumnPerCategory", Err.Number, Err.Description
End Sub

'@sub-title Verify the bivariate total row and column meet at one named cell.
'@TestMethod("CrossTable")
Public Sub TestBivariateNamesTheTotalIntersection()
    CustomTestSetTitles Assert, "CrossTable", "TestBivariateNamesTheTotalIntersection"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "TOTAL_ROW_" & tabId), _
                  "A bivariate build should name the total row"
    Assert.IsTrue RangeExistsOnSheet(sh, "TOTAL_COL_" & tabId), _
                  "A bivariate build should name the total column"
    Assert.IsTrue RangeExistsOnSheet(sh, "TOTAL_TOTAL_" & tabId), _
                  "The total row and total column should meet at a named cell"
    Assert.AreEqual CLng(1), CLng(sh.Range("TOTAL_TOTAL_" & tabId).Cells.Count), _
                    "The total intersection is a single cell"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariateNamesTheTotalIntersection", Err.Number, Err.Description
End Sub

'@sub-title Verify a bivariate missing row and column are named.
'@TestMethod("CrossTable")
Public Sub TestBivariateWithMissingNamesBothAxes()
    CustomTestSetTitles Assert, "CrossTable", "TestBivariateWithMissingNamesBothAxes"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "all", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "MISSING_ROW_" & tabId), _
                  "Missing set to all should name the missing row"
    Assert.IsTrue RangeExistsOnSheet(sh, "MISSING_COL_" & tabId), _
                  "Missing set to all should name the missing column"
    Assert.IsTrue RangeExistsOnSheet(sh, "MISSING_MISSING_" & tabId), _
                  "The two missing axes should meet at a named cell"
    Assert.IsTrue RangeExistsOnSheet(sh, "TOTAL_MISSING_" & tabId), _
                  "The total row and missing column should meet at a named cell"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariateWithMissingNamesBothAxes", Err.Number, Err.Description
End Sub

'@sub-title Verify a bivariate percentage interleaves a column per category.
'@TestMethod("CrossTable")
Public Sub TestBivariateWithPercentageInterleavesColumns()
    CustomTestSetTitles Assert, "CrossTable", "TestBivariateWithPercentageInterleavesColumns"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "row")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "PERC_COL_1_" & tabId), _
                  "The first percentage column should be named"
    Assert.IsTrue RangeExistsOnSheet(sh, "PERC_COL_2_" & tabId), _
                  "The second percentage column should be named"
    Assert.AreEqual CLng(STANDARD_START_COL + 6), tabl.EndColumn, _
                    "Three headed columns each with a percentage give six columns"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariateWithPercentageInterleavesColumns", Err.Number, Err.Description
End Sub

'@sub-title Verify a bivariate table offsets its rows one below a univariate one.
'@TestMethod("CrossTable")
Public Sub TestBivariateOffsetsRowsByOne()
    CustomTestSetTitles Assert, "CrossTable", "TestBivariateOffsetsRowsByOne"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim startRw As Long

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    startRw = tabl.StartRow

    Assert.AreEqual "A", sh.Cells(startRw + 2, STANDARD_START_COL).Value, _
                    "Bivariate row categories start two rows below the start row"
    Assert.AreEqual "Total", sh.Cells(startRw + 5, STANDARD_START_COL).Value, _
                    "The total row follows the three categories"
    Assert.AreEqual startRw + 5, tabl.EndRow, _
                    "EndRow should be the total row"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariateOffsetsRowsByOne", Err.Number, Err.Description
End Sub

'@section Time series
'===============================================================================

'@sub-title Verify a time series build creates the section and the date controls.
'@TestMethod("CrossTable")
Public Sub TestTimeSeriesCreatesDateControls()
    CustomTestSetTitles Assert, "CrossTable", "TestTimeSeriesCreatesDateControls"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("yes")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "SECTION_" & tabId), _
                  "A first temporal table should name its section cell"
    Assert.IsTrue RangeExistsOnSheet(sh, "USER_START_DATE_" & tabId), _
                  "A temporal table should name the start date the user edits"
    Assert.IsTrue RangeExistsOnSheet(sh, "USER_END_DATE_" & tabId), _
                  "A temporal table should name the end date the user edits"
    Assert.IsTrue RangeExistsOnSheet(sh, "TIME_UNIT_" & tabId), _
                  "A temporal table should name the time unit cell"
    Assert.IsTrue RangeExistsOnSheet(sh, "INFO_ANA_PERIOD_" & tabId), _
                  "A temporal table should name the analysis period info cell"
    Assert.AreEqual "Start date", sh.Range("START_DATE_LABEL_" & tabId).Value, _
                    "The start date label should carry the translated text"
    Assert.IsTrue Not sh.Range("USER_START_DATE_" & tabId).Locked, _
                  "The start date the user edits should be unlocked"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesCreatesDateControls", Err.Number, Err.Description
End Sub

'@sub-title Verify the time unit cell starts on the week and the sheet builds no list.
'@details
'The five time units are one workbook-level dropdown that Linelist.Prepare
'adds, so a table writes its own starting value and nothing else. Each analysis
'sheet used to build five cells in column Z, under TIME_UNIT_LIST on a time
'series sheet and SPTIME_UNIT_LIST on a spatio-temporal one, and the reader in
'CustomLinelistFunctions named the first whatever sheet it ran on.
'@TestMethod("CrossTable")
Public Sub TestTimeSeriesStartsOnTheWeekAndBuildsNoList()
    CustomTestSetTitles Assert, "CrossTable", "TestTimeSeriesStartsOnTheWeekAndBuildsNoList"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("yes")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Assert.AreEqual "Week", sh.Range("TIME_UNIT_" & tabl.Specifications.TableId).Value, _
                    "The time unit cell should default to the week"
    Assert.IsFalse RangeExistsOnSheet(sh, "TIME_UNIT_LIST"), _
                   "The sheet should carry no time unit list of its own"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesStartsOnTheWeekAndBuildsNoList", Err.Number, Err.Description
End Sub

'@sub-title Verify a time series table reserves the fixed row grid.
'@TestMethod("CrossTable")
Public Sub TestTimeSeriesReservesTheFixedGrid()
    CustomTestSetTitles Assert, "CrossTable", "TestTimeSeriesReservesTheFixedGrid"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("yes")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Assert.AreEqual tabl.StartRow + NB_ROWS_TIME_SERIES, tabl.EndRow, _
                    "A temporal table reserves a fixed grid of rows"
    Assert.AreEqual "Period", sh.Cells(tabl.StartRow + 1, STANDARD_START_COL).Value, _
                    "The grid opens with the translated period label"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesReservesTheFixedGrid", Err.Number, Err.Description
End Sub

'@sub-title Verify the time period ranges left of the categories are named.
'@TestMethod("CrossTable")
Public Sub TestTimeSeriesNamesThePeriodRanges()
    CustomTestSetTitles Assert, "CrossTable", "TestTimeSeriesNamesThePeriodRanges"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("yes")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "END_TIME_PERIOD_" & tabId), _
                  "A temporal table should name the end of each period"
    Assert.IsTrue RangeExistsOnSheet(sh, "START_TIME_PERIOD_" & tabId), _
                  "A temporal table should name the start of each period"
    Assert.IsTrue RangeExistsOnSheet(sh, "FIRST_VALUE_START_TIME_" & tabId), _
                  "A temporal table should name the first period start"
    Assert.AreEqual CLng(1), CLng(sh.Range("FIRST_VALUE_START_TIME_" & tabId).Cells.Count), _
                    "The first period start is a single cell"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesNamesThePeriodRanges", Err.Number, Err.Description
End Sub

'@sub-title Verify the composite header text is built and written.
'@TestMethod("CrossTable")
Public Sub TestTimeSeriesBuildsTheCompositeHeader()
    CustomTestSetTitles Assert, "CrossTable", "TestTimeSeriesBuildsTheCompositeHeader"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("yes")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Assert.IsTrue InStr(1, tabl.TimeSeriesHeader, "Cases") > 0, _
                  "The composite header should carry the summary label"
    Assert.IsTrue InStr(1, tabl.TimeSeriesHeader, ChrW(9472)) > 0, _
                  "The composite header should separate its parts with a rule"
    Assert.AreEqual tabl.TimeSeriesHeader, _
                    CStr(sh.Range("STARTCOL_" & tabl.Specifications.TableId).Value), _
                    "The composite header should be written into the start column cell"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesBuildsTheCompositeHeader", Err.Number, Err.Description
End Sub

'@sub-title Verify a second temporal table lands beside the first one.
'@details
'The second table in a temporal section shares the row structure of the first
'and only adds columns, so it inherits the start row and starts at a column to
'the right of it. Building the first table is what makes that possible.
'@TestMethod("CrossTable")
Public Sub TestSecondTemporalTableSharesTheSectionRows()
    CustomTestSetTitles Assert, "CrossTable", "TestSecondTemporalTableSharesTheSectionRows"
    On Error GoTo TestFail

    Dim firstTable As CrossTable
    Dim secondTable As CrossTable
    Dim sh As Worksheet

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("yes")
    Set sh = OutputSheet()
    Set firstTable = BuildTable(sh, 1)
    Set secondTable = BuildTable(sh, 2)

    Assert.IsTrue Not secondTable.Specifications.IsNewSection(), _
                  "The second row of one section is no new section"
    Assert.AreEqual firstTable.StartRow, secondTable.StartRow, _
                    "The second temporal table should inherit the start row"
    Assert.AreEqual firstTable.EndRow, secondTable.EndRow, _
                    "Both temporal tables in a section end on the same row"
    Assert.IsTrue secondTable.EndColumn > firstTable.EndColumn, _
                  "The second temporal table should sit to the right of the first"
    Assert.IsTrue Not RangeExistsOnSheet(sh, "SECTION_" & secondTable.Specifications.TableId), _
                  "Only the first table of a section names the section cell"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSecondTemporalTableSharesTheSectionRows", Err.Number, Err.Description
End Sub

'@section Spatial
'===============================================================================

'@sub-title Verify a geo spatial build creates the admin dropdown and the lists.
'@TestMethod("CrossTable")
Public Sub TestSpatialGeoCreatesTheAdminControls()
    CustomTestSetTitles Assert, "CrossTable", "TestSpatialGeoCreatesTheAdminControls"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_SPATIAL, SpatialHeader(), SpatialRows(GEO_VARIABLE, "3")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.AreEqual "geo", tabl.Specifications.SpatialTableScopes(), _
                    "A row variable with an adm1_ twin is a geo table"
    Assert.IsTrue RangeExistsOnSheet(sh, "ADM_DROPDOWN_" & tabId), _
                  "A geo spatial build should name the admin dropdown"
    Assert.IsTrue RangeExistsOnSheet(sh, "DEVIDEPOP_" & tabId), _
                  "A geo spatial build should name the population divide cell"
    Assert.IsTrue RangeExistsOnSheet(sh, "POPFACT_" & tabId), _
                  "A geo spatial build should name the population factor"
    Assert.IsTrue RangeExistsOnSheet(sh, "ADM_UNIT_LIST"), _
                  "A spatial build should create the admin unit list"
    Assert.IsTrue RangeExistsOnSheet(sh, "POPULATION_FACTOR_LIST"), _
                  "A spatial build should create the population factor list"
    Assert.AreEqual CLng(100), CLng(sh.Range("POPFACT_" & tabId).Value), _
                    "The population factor should default to one hundred"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialGeoCreatesTheAdminControls", Err.Number, Err.Description
End Sub

'@sub-title Verify the spatial row count follows the geo count of the row.
'@TestMethod("CrossTable")
Public Sub TestSpatialRowsFollowTheGeoCount()
    CustomTestSetTitles Assert, "CrossTable", "TestSpatialRowsFollowTheGeoCount"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet

    BuildFixture TABLE_SPATIAL, SpatialHeader(), SpatialRows(GEO_VARIABLE, "3")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Assert.AreEqual CLng(3), tabl.Specifications.GeoCount(), _
                    "A geo max of three should answer three"
    Assert.AreEqual CLng(3), _
                    tabl.RowsCategoriesRange(onlyCategories:=True, includeHeaders:=False).Rows.Count, _
                    "Three geo units should reserve three category rows"
    Assert.AreEqual "Missing", sh.Cells(tabl.EndRow, STANDARD_START_COL).Value, _
                    "A missing row should close the geo rows"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialRowsFollowTheGeoCount", Err.Number, Err.Description
End Sub

'@sub-title Verify an empty geo count falls back to the default of five.
'@TestMethod("CrossTable")
Public Sub TestSpatialEmptyGeoCountUsesTheDefault()
    CustomTestSetTitles Assert, "CrossTable", "TestSpatialEmptyGeoCountUsesTheDefault"
    On Error GoTo TestFail

    Dim tabl As CrossTable

    BuildFixture TABLE_SPATIAL, SpatialHeader(), SpatialRows(GEO_VARIABLE, "")
    Set tabl = BuildTable(OutputSheet(), 1)

    Assert.AreEqual CLng(5), _
                    tabl.RowsCategoriesRange(onlyCategories:=True, includeHeaders:=False).Rows.Count, _
                    "An empty geo max should reserve the default five rows"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialEmptyGeoCountUsesTheDefault", Err.Number, Err.Description
End Sub

'@sub-title Verify a spatial build names the outer values used for population.
'@TestMethod("CrossTable")
Public Sub TestSpatialNamesTheOuterValues()
    CustomTestSetTitles Assert, "CrossTable", "TestSpatialNamesTheOuterValues"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_SPATIAL, SpatialHeader(), SpatialRows(GEO_VARIABLE, "3")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "OUTER_VALUES_" & tabId), _
                  "A spatial build should name the outer values"
    Assert.IsTrue sh.Range("OUTER_VALUES_" & tabId).Cells.Count > _
                  sh.Range("INTERIOR_VALUES_" & tabId).Cells.Count, _
                  "The outer values should reach past the interior values"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialNamesTheOuterValues", Err.Number, Err.Description
End Sub

'@sub-title Verify a health facility spatial build takes the other branch.
'@TestMethod("CrossTable")
Public Sub TestSpatialHealthFacilityCreatesItsDropdown()
    CustomTestSetTitles Assert, "CrossTable", "TestSpatialHealthFacilityCreatesItsDropdown"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_SPATIAL, SpatialHeader(), SpatialRows(HF_VARIABLE, "2")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.AreEqual "hf", tabl.Specifications.SpatialTableScopes(), _
                    "A row variable with an hf_ twin is a health facility table"
    Assert.IsTrue RangeExistsOnSheet(sh, "HF_DROPDOWN_" & tabId), _
                  "A health facility build should name its dropdown"
    Assert.AreEqual "Health facility", sh.Range("HF_DROPDOWN_" & tabId).Value, _
                    "The dropdown should carry the translated health facility label"
    Assert.IsTrue Not RangeExistsOnSheet(sh, "ADM_DROPDOWN_" & tabId), _
                  "A health facility build should name no admin dropdown"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialHealthFacilityCreatesItsDropdown", Err.Number, Err.Description
End Sub

'@section Spatio-temporal
'===============================================================================

'@sub-title Verify a spatio-temporal build creates the geo input cells.
'@details
'The wrap is asserted here because the geo form used to write it across the
'whole used range of the sheet on every validated place. The build owns it now,
'so the cells carry it the moment the section exists.
'@TestMethod("CrossTable")
Public Sub TestSpatioTemporalCreatesTheGeoInputs()
    CustomTestSetTitles Assert, "CrossTable", "TestSpatioTemporalCreatesTheGeoInputs"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRows(GEO_VARIABLE, "3", "geo")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "INPUTSPTGEO_LABEL_" & tabId), _
                  "A spatio-temporal build should name the geo input label"
    Assert.IsTrue RangeExistsOnSheet(sh, "INPUTSPTGEO_USER_" & tabId), _
                  "A spatio-temporal build should name the geo cells the user edits"
    Assert.AreEqual CLng(3), CLng(sh.Range("INPUTSPTGEO_USER_" & tabId).Rows.Count), _
                    "A geo count of three should give three input cells"
    Assert.IsTrue Not sh.Range("INPUTSPTGEO_USER_" & tabId).Locked, _
                  "The geo cells the user edits should be unlocked"
    Assert.IsTrue sh.Range("INPUTSPTGEO_USER_" & tabId).WrapText, _
                  "The geo cells the user edits should wrap from build time"
    Assert.AreEqual CLng(3), CLng(sh.Range("INPUTSPTGEO_NGEO_" & tabId).Value), _
                    "The geo count cell should carry the count"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalCreatesTheGeoInputs", Err.Number, Err.Description
End Sub

'@sub-title Verify a spatio-temporal build names one input cell per geo unit.
'@TestMethod("CrossTable")
Public Sub TestSpatioTemporalNamesOneInputPerGeoUnit()
    CustomTestSetTitles Assert, "CrossTable", "TestSpatioTemporalNamesOneInputPerGeoUnit"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim tabl As CrossTable
    Dim tabId As String

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRows(GEO_VARIABLE, "3", "geo")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "INPUTSPTGEO_1_" & tabId), _
                  "The first geo input cell should be named"
    Assert.IsTrue RangeExistsOnSheet(sh, "INPUTSPTGEO_3_" & tabId), _
                  "The third geo input cell should be named"
    Assert.IsTrue Not RangeExistsOnSheet(sh, "INPUTSPTGEO_4_" & tabId), _
                  "No fourth geo input cell should be named"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalNamesOneInputPerGeoUnit", Err.Number, Err.Description
End Sub

'@sub-title Verify a logged check files its message in the label.
'@details
'Checking.Add takes (key, label, scope). The call passed (message, scope), so the
'message landed in the key, the Byte scope was coerced into the label, and the
'scope the caller asked for was discarded in favour of the default note. The
'report therefore showed a bare number where the text belonged.
'
'The key names the class, the table and a counter. AnalysisOutput pours the
'entries of this class and of CrossTableFormula, for every table of the sheet,
'into one report, and Checking.Add raises on a duplicate key, so a bare counter
'made the second harvest raise on a key the first had taken.
'@TestMethod("CrossTable")
Public Sub TestALoggedCheckCarriesItsMessageAsTheLabel()
    CustomTestSetTitles Assert, "CrossTable", "TestALoggedCheckCarriesItsMessageAsTheLabel"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim tabl As CrossTable
    Dim checks As Checking
    Dim firstKey As String
    Dim label As String

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Assert.IsTrue tabl.HasCheckings, "Build should have logged at least one check"

    Set checks = tabl.CheckingValues
    firstKey = CStr(checks.ListOfKeys.Item(checks.ListOfKeys.LowerBound))
    label = checks.ValueOf(firstKey)

    Assert.AreEqual "CrossTable-" & tabl.Specifications.TableId & "-1", firstKey, _
                    "The key names the class, the table and the entry, and it reads [" & _
                    firstKey & "]"
    Assert.IsTrue Not IsNumeric(label), _
                  "The label should hold the message text the caller passed"
    Assert.IsTrue Len(label) > 1, _
                  "The label should carry readable text"

    ' The scope was the third casualty: with the message occupying the key and
    ' the scope byte occupying the label, the scope itself fell back to the
    ' default note, so a success was filed as a note.
    Assert.IsTrue InStr(1, checks.ValueOf(firstKey, checkingType), _
                        "Success", vbTextCompare) > 0, _
                  "The check should keep the success scope the caller asked for"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALoggedCheckCarriesItsMessageAsTheLabel", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a spatial table lacking both prefixes raises a stated error.
'@details
'AddHeader assigns its range inside the geo and facility arms and dereferences it
'underneath them. The third outcome of SpatialTableScopes reached that
'dereference with the range still Nothing, and raised 91 with no mention of the
'table or the cause.
'
'Only Err.Number is asserted, which is this module's convention for every error
'test in it. VBA carries the number across a class boundary and replaces the
'description with "Method 'X' of object 'Y' failed", so the guard's own wording
'is unreachable from here; the LLFormat and LLSheets suites pin that same
'behaviour. The description is captured anyway and travels in the assertion
'message, so a later surprise reports the text it actually got.
'@TestMethod("CrossTable")
Public Sub TestSpatialTableWithoutAPrefixNamesTheProblem()
    CustomTestSetTitles Assert, "CrossTable", "TestSpatialTableWithoutAPrefixNamesTheProblem"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim tabl As CrossTable
    Dim errNumber As Long
    Dim errDescription As String

    ' A choice variable has no adm1_ or hf_ twin in the dictionary fixture.
    BuildFixture TABLE_SPATIAL, SpatialHeader(), SpatialRows(ROW_CHOICE_VARIABLE, "3")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    On Error Resume Next
    tabl.AddHeader
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual ProjectError.ErrorUnexpectedState, errNumber, _
                    "A spatial table lacking both prefixes should raise " & _
                    "unexpected state - description was [" & errDescription & "]"

    ' The guard raises a message naming adm1_ and hf_, and that message is gone
    ' by the time it arrives here. This line prints what replaced it, so the
    ' behaviour is on the record in test-results.csv. The LLFormat and LLSheets
    ' suites assert the same shape for their own classes.
    Assert.LogSuccesses "AddHeader boundary left description [" & _
                        errDescription & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialTableWithoutAPrefixNamesTheProblem", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a facility table tags its inputs from the dictionary, not the setup column.
'@details
'THE SPATIAL TYPE CELL IS EMPTY ON PURPOSE HERE, AND THAT IS THE WHOLE TEST.
'The setup workbook validates "spatial type" on Tab_SpatioTemporal_Specs and
'deliberately leaves the column on Tab_SpatioTemporal_Analysis unvalidated;
'the propagation from specs to analysis clears the geo cell and re-sets its
'dropdown without ever writing the value across. So on a workbook filled the
'documented way this cell is blank, and reading it tagged every table
'geographic. The formula writer meanwhile probes the dictionary, agreed the
'table was a facility one, and referenced INPUTSPTHF_ names that the writer
'above had never created.
'
'Passing "hf" here instead would make this test pass against the old code and
'prove nothing, because the fault is that the cell is empty in the field.
'@TestMethod("CrossTable")
Public Sub TestFacilityInputsAreTaggedWithoutTheSpatialTypeColumn()
    CustomTestSetTitles Assert, "CrossTable", "TestFacilityInputsAreTaggedWithoutTheSpatialTypeColumn"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim tabl As CrossTable
    Dim tabId As String

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRows(HF_VARIABLE, "2", vbNullString)
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.AreEqual vbNullString, CStr(tabl.Specifications.Value("spatial type")), _
                    "The fixture leaves the spatial type cell empty, as the field does"
    Assert.AreEqual "hf", tabl.Specifications.SpatialTableScopes, _
                    "The dictionary probe still knows this is a facility table"

    Assert.IsTrue RangeExistsOnSheet(sh, "INPUTSPTHF_1_" & tabId), _
                  "The first facility input cell should carry the facility tag"
    Assert.IsTrue RangeExistsOnSheet(sh, "INPUTSPTHF_2_" & tabId), _
                  "The second facility input cell should carry the facility tag"
    Assert.IsTrue Not RangeExistsOnSheet(sh, "INPUTSPTGEO_1_" & tabId), _
                  "A facility table should not tag its inputs geographic"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFacilityInputsAreTaggedWithoutTheSpatialTypeColumn", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify an administrative table still tags its inputs geographic.
'@details
'The sibling of the test above. This one passed before the fix too, because the
'geographic tag was what an unread column fell back to; it is here so a later
'change that inverts the mapping trips something.
'@TestMethod("CrossTable")
Public Sub TestAdministrativeInputsKeepTheGeographicTag()
    CustomTestSetTitles Assert, "CrossTable", "TestAdministrativeInputsKeepTheGeographicTag"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim tabl As CrossTable
    Dim tabId As String

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRows(GEO_VARIABLE, "2", vbNullString)
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.AreEqual "geo", tabl.Specifications.SpatialTableScopes, _
                    "The dictionary probe knows this is an administrative table"
    Assert.IsTrue RangeExistsOnSheet(sh, "INPUTSPTGEO_1_" & tabId), _
                  "An administrative table should tag its inputs geographic"
    Assert.IsTrue Not RangeExistsOnSheet(sh, "INPUTSPTHF_1_" & tabId), _
                  "An administrative table should not carry the facility tag"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAdministrativeInputsKeepTheGeographicTag", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a spatio-temporal table takes its columns from the geo count.
'@TestMethod("CrossTable")
Public Sub TestSpatioTemporalColumnsFollowTheGeoCount()
    CustomTestSetTitles Assert, "CrossTable", "TestSpatioTemporalColumnsFollowTheGeoCount"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRows(GEO_VARIABLE, "3", "geo")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Assert.AreEqual CLng(3), tabl.NumberOfColumns, _
                    "The geo count decides how many columns a spatio-temporal table has"
    Assert.AreEqual tabl.StartRow + NB_ROWS_TIME_SERIES, tabl.EndRow, _
                    "A spatio-temporal table reserves the same fixed grid"
    Assert.IsFalse RangeExistsOnSheet(sh, "SPTIME_UNIT_LIST"), _
                   "A spatio-temporal build should carry no time unit list of its own"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalColumnsFollowTheGeoCount", Err.Number, Err.Description
End Sub

'@section Named ranges and checkings
'===============================================================================

'@sub-title Verify the build names its ranges on the worksheet.
'@details
'The export reads the names off the worksheet, so a name the build forgets to
'create is a name no exported analysis sheet can carry. This replaces the two
'tests that read the list the class used to keep by hand.
'@TestMethod("CrossTable")
Public Sub TestBuildNamesTheRangesOnTheSheet()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildNamesTheRangesOnTheSheet"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "INTERIOR_VALUES_" & tabId), _
                  "The build should name the interior values on the sheet"
    Assert.IsTrue RangeExistsOnSheet(sh, "ROW_CATEGORIES_" & tabId), _
                  "The build should name the row categories on the sheet"
    Assert.IsTrue RangeExistsOnSheet(sh, "LABEL_ROW_CATEGORIES_" & tabId), _
                  "The build should name the row category labels on the sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildNamesTheRangesOnTheSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Build files a check naming the table it built.
'@TestMethod("CrossTable")
Public Sub TestBuildFilesACheck()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildFilesACheck"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim checks As Checking

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set tabl = CrossTable.Create(CreateSpecs(1), OutputSheet(), lData)

    Assert.IsTrue Not tabl.HasCheckings, _
                  "A table reports no checks before it is built"

    tabl.Build
    Set checks = tabl.CheckingValues

    Assert.IsTrue tabl.HasCheckings, _
                  "A built table should report that it filed a check"
    Assert.IsTrue (Not checks Is Nothing), _
                  "CheckingValues should hand back the checks after a build"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildFilesACheck", Err.Number, Err.Description
End Sub

'@section Format
'===============================================================================

'@sub-title Verify Format runs over a built univariate table.
'@details
'Format touches about forty names through the design object. The assertion that
'matters here is that every name it reaches resolves, which a raise would show.
'@TestMethod("CrossTable")
Public Sub TestFormatRunsOverAUnivariateTable()
    CustomTestSetTitles Assert, "CrossTable", "TestFormatRunsOverAUnivariateTable"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim designFormat As LLFormat
    Dim errNumber As Long

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "yes", "yes")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    Set designFormat = LLFormat.Create(PrepareLLFormatFixture(FORMAT_SHEET))

    On Error Resume Next
    tabl.Format designFormat
    errNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(0), errNumber, _
                    "Formatting a built univariate table should raise nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFormatRunsOverAUnivariateTable", Err.Number, Err.Description
End Sub

'@sub-title Format one bivariate table and report what it did.
'@details
'Shared by the three bivariate format tests. Each one turns on one more feature,
'so the first of them that raises names the feature that carries the fault. The
'geometry goes into the results file through LogSuccesses either way:
'Debug.Print is lost in a headless run, and the column layout is the fact a
'failure here needs.
'@param testName String. The calling test, for the failure message.
'@param missing String. The missing-data flag for the fixture row.
'@param percentage String. The percentage flag for the fixture row.
Private Sub CheckBivariateFormat(ByVal testName As String, _
                                 ByVal missing As String, _
                                 ByVal percentage As String)
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim designFormat As LLFormat
    Dim errNumber As Long
    Dim errDescription As String
    Dim tabId As String

    On Error GoTo ArrangeFailed

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, missing, percentage)
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    Set designFormat = LLFormat.Create(PrepareLLFormatFixture(FORMAT_SHEET))
    tabId = tabl.Specifications.TableId

    Assert.LogSuccesses testName & " layout: startRow=" & tabl.StartRow & _
                        " endRow=" & tabl.EndRow & " endCol=" & tabl.EndColumn & _
                        " nbCols=" & tabl.NumberOfColumns & _
                        " interior=" & NamedAddress(sh, "INTERIOR_VALUES_" & tabId) & _
                        " totalCol=" & NamedAddress(sh, "TOTAL_COL_VALUES_" & tabId) & _
                        " totalPerc=" & NamedAddress(sh, "TOTAL_PERC_VALUES_" & tabId) & _
                        " missingCol=" & NamedAddress(sh, "MISSING_COL_VALUES_" & tabId) & _
                        " missingPerc=" & NamedAddress(sh, "MISSING_PERC_VALUES_" & tabId) & _
                        " labelCol1=" & NamedAddress(sh, "LABEL_COL_1_" & tabId)

    On Error Resume Next
    tabl.Format designFormat
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo ArrangeFailed

    Assert.AreEqual CLng(0), errNumber, _
                    "Formatting this bivariate table should raise nothing, and it " & _
                    "raised " & errNumber & ": " & errDescription

    Exit Sub
ArrangeFailed:
    CustomTestLogFailure Assert, testName, Err.Number, Err.Description
End Sub

'@sub-title Verify Format runs over a plain bivariate table.
'@TestMethod("CrossTable")
Public Sub TestFormatRunsOverAPlainBivariateTable()
    CustomTestSetTitles Assert, "CrossTable", "TestFormatRunsOverAPlainBivariateTable"
    CheckBivariateFormat "TestFormatRunsOverAPlainBivariateTable", "no", "no"
End Sub

'@sub-title Verify Format runs over a bivariate table with missing on both axes.
'@TestMethod("CrossTable")
Public Sub TestFormatRunsOverABivariateTableWithMissing()
    CustomTestSetTitles Assert, "CrossTable", "TestFormatRunsOverABivariateTableWithMissing"
    CheckBivariateFormat "TestFormatRunsOverABivariateTableWithMissing", "all", "no"
End Sub

'@sub-title Verify Format runs over a bivariate table with percentage columns.
'@TestMethod("CrossTable")
Public Sub TestFormatRunsOverABivariateTableWithPercentage()
    CustomTestSetTitles Assert, "CrossTable", "TestFormatRunsOverABivariateTableWithPercentage"
    CheckBivariateFormat "TestFormatRunsOverABivariateTableWithPercentage", "no", "row"
End Sub

'@sub-title Verify Format runs over a bivariate table with missing and percentage.
'@TestMethod("CrossTable")
Public Sub TestFormatRunsOverABivariateTableWithBoth()
    CustomTestSetTitles Assert, "CrossTable", "TestFormatRunsOverABivariateTableWithBoth"
    CheckBivariateFormat "TestFormatRunsOverABivariateTableWithBoth", "all", "row"
End Sub

'@sub-title Verify each build stage of a bivariate table with missing and percentage.
'@details
'Build runs AddHeader, AddRows, AddColumns and NameRanges in that order. This
'test calls them one at a time and logs the geometry after each, so a raise
'names the stage that carries it. One call to Build cannot say which of the
'four stopped.
'@TestMethod("CrossTable")
Public Sub TestBuildStagesOverABivariateTableWithBoth()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim stageName As String

    CustomTestSetTitles Assert, "CrossTable", "TestBuildStagesOverABivariateTableWithBoth"
    On Error GoTo StageFailed

    stageName = "arrange"
    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "all", "row")
    Set sh = OutputSheet()
    Set tabl = CrossTable.Create(CreateSpecs(1), sh, lData)

    stageName = "AddHeader"
    tabl.AddHeader
    Assert.LogSuccesses "stage AddHeader done: startRow=" & tabl.StartRow

    stageName = "AddRows"
    tabl.AddRows
    Assert.LogSuccesses "stage AddRows done: endRow=" & tabl.EndRow

    stageName = "AddColumns"
    tabl.AddColumns
    Assert.LogSuccesses "stage AddColumns done: endRow=" & tabl.EndRow & _
                        " endCol=" & tabl.EndColumn & _
                        " nbCols=" & tabl.NumberOfColumns

    stageName = "NameRanges"
    tabl.NameRanges
    Assert.LogSuccesses "stage NameRanges done: interior=" & _
                        NamedAddress(sh, "INTERIOR_VALUES_" & tabl.Specifications.TableId)

    Assert.IsTrue True, "Every build stage of this bivariate table should run"
    Exit Sub
StageFailed:
    CustomTestLogFailure Assert, "stage " & stageName, Err.Number, Err.Description
End Sub

'@sub-title Verify percentage columns are built from a zero-based category list.
'@details
'AddColumns appends Missing and Total to the column categories, then walks the
'result to interleave a percentage column after each one. That walk used to run
'index one to Length, which reads one past the end of a list that starts at
'zero, and the walk only runs when the table carries percentages, so the
'combination of a missing column and a percentage column was the one that
'raised. The choices class answers one-based lists, so this seeds a zero-based
'one on purpose and puts the bounds back before it returns.
'@TestMethod("CrossTable")
Public Sub TestBivariateColumnsAcceptAZeroBasedCategoryList()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim zeroBased As BetterArray
    Dim headerRow As Long
    Dim errNumber As Long
    Dim errDescription As String

    CustomTestSetTitles Assert, "CrossTable", "TestBivariateColumnsAcceptAZeroBasedCategoryList"
    On Error GoTo TestFail

    Set zeroBased = New BetterArray
    zeroBased.LowerBound = 0
    zeroBased.Push "X", "Y"
    lData.SetCategoriesVerbatim COL_CHOICE_VARIABLE, zeroBased

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "all", "row")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    headerRow = tabl.StartRow

    Assert.AreEqual CLng(2), CLng(tabl.NumberOfColumns), _
                    "A zero-based list of two categories is still two data columns"
    Assert.AreEqual CLng(STANDARD_START_COL + 8), tabl.EndColumn, _
                    "Two categories plus Missing and Total, each followed by a " & _
                    "percentage column, fill eight columns. The header row reads " & _
                    RowText(sh, headerRow, STANDARD_START_COL + 1, STANDARD_START_COL + 8)
    Assert.AreEqual trans.TranslatedValue("MSG_Total"), _
                    CellText(sh, headerRow, STANDARD_START_COL + 7), _
                    "The last category written is Total, and it is the entry the " & _
                    "one-based walk used to read past"

    lData.SetCategories COL_CHOICE_VARIABLE, BetterArrayFromList("X", "Y")
    Exit Sub
TestFail:
    ' Read the error before restoring the categories, for the reason given on
    ' the single-cell test below.
    errNumber = Err.Number
    errDescription = Err.Description
    lData.SetCategories COL_CHOICE_VARIABLE, BetterArrayFromList("X", "Y")
    CustomTestLogFailure Assert, "TestBivariateColumnsAcceptAZeroBasedCategoryList", _
                         errNumber, errDescription
End Sub

'@section Stage isolation — AddHeader
'===============================================================================
'@description
'Build runs AddHeader, AddRows, AddColumns and NameRanges in that order, and
'each stage is called on its own here so a raise names the stage it came from.
'AddHeader runs first in every one of these tests because it inserts rows above
'the table and settles the start row the later stages address from.

'@sub-title Verify AddHeader writes the title and the section and no rows.
'@TestMethod("CrossTable")
Public Sub TestAddHeaderWritesTheTitleAndSectionOnly()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    CustomTestSetTitles Assert, "CrossTable", "TestAddHeaderWritesTheTitleAndSectionOnly"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "TITLE_" & tabId), _
                  "AddHeader names the title cell"
    Assert.AreEqual "A bivariate table", CStr(sh.Range("TITLE_" & tabId).Value), _
                    "The title cell carries the title from the specification row"
    Assert.IsTrue RangeExistsOnSheet(sh, "SECTION_" & tabId), _
                  "A first table opens its section, so AddHeader names the section cell"
    Assert.AreEqual "S1", CStr(sh.Range("SECTION_" & tabId).Value), _
                    "The section cell carries the section from the specification row"
    Assert.AreEqual vbNullString, CellText(sh, tabl.StartRow + 1, STANDARD_START_COL), _
                    "AddHeader writes no row labels, since AddRows owns that cell"
    Assert.AreEqual CLng(0), tabl.EndRow, _
                    "AddHeader settles no end row, since AddRows owns it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddHeaderWritesTheTitleAndSectionOnly", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the global summary banner is written once for the whole block.
'@details
'Every global summary row shares one banner and one column pair, and the guard
'is the ROWGS_SET name rather than a flag on the class, so the second row has to
'find the first row's banner and leave it alone.
'@TestMethod("CrossTable")
Public Sub TestAddHeaderWritesTheGlobalSummaryBannerOnce()
    Dim firstTable As CrossTable
    Dim secondTable As CrossTable
    Dim sh As Worksheet
    Dim bannerAddress As String

    CustomTestSetTitles Assert, "CrossTable", "TestAddHeaderWritesTheGlobalSummaryBannerOnce"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), GlobalSummaryTwoRows()
    Set sh = OutputSheet()

    Set firstTable = NewTable(sh, 1)
    firstTable.AddHeader
    bannerAddress = NamedAddress(sh, "ROWGS_SET")

    Assert.IsTrue (bannerAddress <> vbNullString), _
                  "The first global summary row writes the shared banner"
    Assert.AreEqual trans.TranslatedValue("MSG_GlobalSummary"), _
                    CStr(sh.Range("ROWGS_SET").Value), _
                    "The banner carries the translated global summary label"

    Set secondTable = NewTable(sh, 2)
    secondTable.AddHeader

    Assert.AreEqual bannerAddress, NamedAddress(sh, "ROWGS_SET"), _
                    "The second row finds the banner already there and leaves it in place"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddHeaderWritesTheGlobalSummaryBannerOnce", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify AddHeader alone reserves the temporal control band.
'@details
'A time series table inserts seven rows above itself and fills them with the
'date and time-unit controls. Calling AddHeader on its own shows the band is
'this stage's work rather than a side effect of the later ones.
'@TestMethod("CrossTable")
Public Sub TestAddHeaderReservesTheTemporalBand()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    CustomTestSetTitles Assert, "CrossTable", "TestAddHeaderReservesTheTemporalBand"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "SECTION_" & tabId), _
                  "AddHeader opens the section of a first temporal table"
    Assert.IsTrue (sh.Range("SECTION_" & tabId).Row < tabl.StartRow), _
                  "The section cell sits above the table, in the band AddHeader inserted"
    Assert.AreEqual CLng(0), tabl.EndRow, _
                    "AddHeader reserves rows without settling the end row"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddHeaderReservesTheTemporalBand", _
                         Err.Number, Err.Description
End Sub

'@section Stage isolation — AddRows
'===============================================================================

'@sub-title Verify AddRows writes the univariate categories then the total.
'@TestMethod("CrossTable")
Public Sub TestAddRowsWritesUnivariateCategoriesThenTotal()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim startRw As Long

    CustomTestSetTitles Assert, "CrossTable", "TestAddRowsWritesUnivariateCategoriesThenTotal"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    startRw = tabl.StartRow
    tabl.AddRows

    Assert.AreEqual DictLabel(ROW_CHOICE_VARIABLE, "main label"), _
                    CellText(sh, startRw, STANDARD_START_COL), _
                    "The row variable label sits on the start row of a univariate table"
    Assert.AreEqual "A|B|C|" & trans.TranslatedValue("MSG_Total"), _
                    RowText(sh, startRw + 1, STANDARD_START_COL, STANDARD_START_COL) & "|" & _
                    RowText(sh, startRw + 2, STANDARD_START_COL, STANDARD_START_COL) & "|" & _
                    RowText(sh, startRw + 3, STANDARD_START_COL, STANDARD_START_COL) & "|" & _
                    RowText(sh, startRw + 4, STANDARD_START_COL, STANDARD_START_COL), _
                    "The three seeded categories are written in order and the total closes them"
    Assert.AreEqual CLng(startRw + 4), tabl.EndRow, _
                    "The end row is the last row AddRows wrote"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddRowsWritesUnivariateCategoriesThenTotal", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a bivariate row missing setting adds the missing row.
'@details
'The bivariate arm adds a missing row for both "row" and "all", and "all" also
'adds a missing column. This pins the "row" spelling, which adds the row alone.
'@TestMethod("CrossTable")
Public Sub TestAddRowsAddsTheMissingRowForRowMissing()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim startRw As Long

    CustomTestSetTitles Assert, "CrossTable", "TestAddRowsAddsTheMissingRowForRowMissing"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "row", "no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    startRw = tabl.StartRow
    tabl.AddRows

    Assert.AreEqual DictLabel(ROW_CHOICE_VARIABLE, "main label"), _
                    CellText(sh, startRw + 1, STANDARD_START_COL), _
                    "A bivariate table carries its row variable label one row lower " & _
                    "than a univariate one, because the column headers take two rows"
    Assert.AreEqual trans.TranslatedValue("MSG_NA"), _
                    CellText(sh, startRw + 5, STANDARD_START_COL), _
                    "The missing row follows the three categories"
    Assert.AreEqual trans.TranslatedValue("MSG_Total"), _
                    CellText(sh, startRw + 6, STANDARD_START_COL), _
                    "The total row closes the table under the missing row"
    Assert.AreEqual CLng(startRw + 6), tabl.EndRow, _
                    "The end row counts the label row, the categories, missing and total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddRowsAddsTheMissingRowForRowMissing", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify AddRows writes only the label for a global summary row.
'@details
'A global summary row is one cell in a shared block, so this arm writes the
'label and settles no end row. The block, not the row, owns the geometry.
'@TestMethod("CrossTable")
Public Sub TestAddRowsWritesOnlyTheLabelForGlobalSummary()
    Dim tabl As CrossTable
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestAddRowsWritesOnlyTheLabelForGlobalSummary"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), GlobalSummaryTwoRows()
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows

    Assert.AreEqual "First summary", CellText(sh, tabl.StartRow, STANDARD_START_COL), _
                    "The summary label is written on the start row"
    Assert.AreEqual CLng(0), tabl.EndRow, _
                    "A global summary row settles no end row of its own"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddRowsWritesOnlyTheLabelForGlobalSummary", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify AddRows reserves the fixed temporal grid.
'@details
'A temporal table always reserves the same number of rows whatever the date
'range turns out to be, because the dates are chosen at run time and the
'formulas address a grid of a known height.
'@TestMethod("CrossTable")
Public Sub TestAddRowsReservesTheTemporalGrid()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim startRw As Long

    CustomTestSetTitles Assert, "CrossTable", "TestAddRowsReservesTheTemporalGrid"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    startRw = tabl.StartRow
    tabl.AddRows

    Assert.AreEqual trans.TranslatedValue("MSG_Period"), _
                    CellText(sh, startRw + 1, STANDARD_START_COL), _
                    "The period label opens the temporal grid"
    Assert.AreEqual CLng(startRw + NB_ROWS_TIME_SERIES), tabl.EndRow, _
                    "The grid is a fixed height whatever the date range holds"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddRowsReservesTheTemporalGrid", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify spatial rows follow the geo count and close with missing.
'@TestMethod("CrossTable")
Public Sub TestAddRowsFollowsTheGeoCountForSpatial()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim startRw As Long
    Dim geoCount As Long

    CustomTestSetTitles Assert, "CrossTable", "TestAddRowsFollowsTheGeoCountForSpatial"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIAL, SpatialHeader(), SpatialRows(GEO_VARIABLE, "4")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    startRw = tabl.StartRow
    geoCount = tabl.Specifications.GeoCount()
    tabl.AddRows

    Assert.AreEqual CLng(4), geoCount, _
                    "The geo count is read from the specification row"
    Assert.AreEqual DictLabel(GEO_PREFIXED_VARIABLE, "sub section"), _
                    CellText(sh, startRw + 1, STANDARD_START_COL), _
                    "A geo spatial table labels its rows with the admin level sub section"
    Assert.AreEqual trans.TranslatedValue("MSG_NA"), _
                    CellText(sh, startRw + geoCount + 2, STANDARD_START_COL), _
                    "The missing row closes the reserved geo unit rows"
    Assert.AreEqual CLng(startRw + geoCount + 2), tabl.EndRow, _
                    "The end row counts the label, the geo units and the missing row"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddRowsFollowsTheGeoCountForSpatial", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a health facility spatial table uses the facility label.
'@details
'The spatial arm picks the label column by spatial type: an admin level reads
'sub section and a health facility reads main label. This is one of the four
'places the geo-versus-facility decision is made.
'@TestMethod("CrossTable")
Public Sub TestAddRowsUsesTheHealthFacilityLabel()
    Dim tabl As CrossTable
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestAddRowsUsesTheHealthFacilityLabel"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIAL, SpatialHeader(), SpatialRows(HF_VARIABLE, "3")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows

    Assert.AreEqual "hf", tabl.Specifications.SpatialTableScopes(), _
                    "A row naming a facility variable resolves to the facility scope"
    Assert.AreEqual DictLabel(HF_PREFIXED_VARIABLE, "main label"), _
                    CellText(sh, tabl.StartRow + 1, STANDARD_START_COL), _
                    "A facility spatial table labels its rows with the main label"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddRowsUsesTheHealthFacilityLabel", _
                         Err.Number, Err.Description
End Sub

'@section Stage isolation — AddColumns
'===============================================================================

'@sub-title Verify a global summary block gets its two shared columns.
'@TestMethod("CrossTable")
Public Sub TestAddColumnsWritesTheGlobalSummaryPair()
    Dim tabl As CrossTable
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestAddColumnsWritesTheGlobalSummaryPair"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), GlobalSummaryTwoRows()
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows
    tabl.AddColumns

    Assert.IsTrue RangeExistsOnSheet(sh, "COLGS_SET"), _
                  "The shared column pair is named so the next row can find it"
    Assert.AreEqual trans.TranslatedValue("MSG_AllData") & "|" & _
                    trans.TranslatedValue("MSG_FilteredData"), _
                    RowText(sh, tabl.StartRow - 1, STANDARD_START_COL + 1, _
                            STANDARD_START_COL + 2), _
                    "The pair is all data then filtered data"
    Assert.AreEqual CLng(2), CLng(tabl.NumberOfColumns), _
                    "A global summary block is two columns wide"
    Assert.AreEqual CLng(STANDARD_START_COL + 2), tabl.EndColumn, _
                    "The end column is the second of the pair"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddColumnsWritesTheGlobalSummaryPair", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a univariate table reports one data column.
'@TestMethod("CrossTable")
Public Sub TestAddColumnsCountsOneColumnForUnivariate()
    Dim tabl As CrossTable
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestAddColumnsCountsOneColumnForUnivariate"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows
    tabl.AddColumns

    Assert.AreEqual "Cases", CellText(sh, tabl.StartRow, STANDARD_START_COL + 1), _
                    "The summary label heads the single data column"
    Assert.AreEqual CLng(1), CLng(tabl.NumberOfColumns), _
                    "A univariate table has one data column"
    Assert.AreEqual CLng(STANDARD_START_COL + 1), tabl.EndColumn, _
                    "The end column is the data column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddColumnsCountsOneColumnForUnivariate", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a univariate percentage setting adds a second column.
'@TestMethod("CrossTable")
Public Sub TestAddColumnsAddsThePercentageColumnForUnivariate()
    Dim tabl As CrossTable
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestAddColumnsAddsThePercentageColumnForUnivariate"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "yes")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows
    tabl.AddColumns

    Assert.AreEqual "Cases|" & trans.TranslatedValue("MSG_Percent"), _
                    RowText(sh, tabl.StartRow, STANDARD_START_COL + 1, _
                            STANDARD_START_COL + 2), _
                    "The percentage column follows the summary column"
    Assert.AreEqual CLng(1), CLng(tabl.NumberOfColumns), _
                    "The percentage column is not a data column of its own"
    Assert.AreEqual CLng(STANDARD_START_COL + 2), tabl.EndColumn, _
                    "The end column counts the percentage column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddColumnsAddsThePercentageColumnForUnivariate", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the bivariate header row carries the categories then the total.
'@details
'This pins the order the column entries are written in, which every named
'column range is then cut from. The label row below repeats the summary label
'under each entry.
'@TestMethod("CrossTable")
Public Sub TestAddColumnsWritesTheBivariateHeaderRow()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim startRw As Long

    CustomTestSetTitles Assert, "CrossTable", "TestAddColumnsWritesTheBivariateHeaderRow"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows
    startRw = tabl.StartRow
    tabl.AddColumns

    Assert.AreEqual "X|Y|" & trans.TranslatedValue("MSG_Total"), _
                    RowText(sh, startRw, STANDARD_START_COL + 1, STANDARD_START_COL + 3), _
                    "The two seeded categories are written in order and the total follows"
    Assert.AreEqual "Cases|Cases|Cases", _
                    RowText(sh, startRw + 1, STANDARD_START_COL + 1, STANDARD_START_COL + 3), _
                    "The summary label is repeated under every column entry"
    Assert.AreEqual DictLabel(COL_CHOICE_VARIABLE, "main label"), _
                    CStr(sh.Range("STARTCOL_" & tabl.Specifications.TableId).Value), _
                    "The column variable label is written into the start column cell"
    Assert.AreEqual CLng(2), CLng(tabl.NumberOfColumns), _
                    "The total is not counted as a data column"
    Assert.AreEqual CLng(STANDARD_START_COL + 3), tabl.EndColumn, _
                    "The end column counts the total"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddColumnsWritesTheBivariateHeaderRow", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a bivariate table with no column variable falls back to one column.
'@details
'A bivariate row naming no column variable is a summary down one axis, and the
'fallback writes a single column and a placeholder under it. The placeholder is
'a space rather than an empty cell because the next table finds its own start
'column by walking left from the end of that row.
'
'This is the one scope where the fallback has to drop the total label itself.
'HasTotal ends in "and a column variable was named" for the temporal and the
'spatial scopes, so their label list is already empty here, and the bivariate
'scope answers True whatever the row says. A total row down a table grouped by
'nothing still means something, which is why the guard belongs in the column
'builder rather than in the flag.
'@TestMethod("CrossTable")
Public Sub TestAddColumnsFallsBackToOneSummaryColumn()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim startRw As Long

    CustomTestSetTitles Assert, "CrossTable", "TestAddColumnsFallsBackToOneSummaryColumn"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, vbNullString, "no", "no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows
    startRw = tabl.StartRow
    tabl.AddColumns

    Assert.IsTrue tabl.Specifications.HasTotal, _
                  "A bivariate row reports a total whether or not it names a column " & _
                  "variable, so the total label is in play when the fallback runs"
    Assert.AreEqual trans.TranslatedValue("MSG_Total"), _
                    CellText(sh, tabl.EndRow, STANDARD_START_COL), _
                    "The total row is still written, since totalling down the rows " & _
                    "of a table grouped by nothing is still meaningful"
    Assert.AreEqual "Cases", CellText(sh, startRw, STANDARD_START_COL + 1), _
                    "The summary label heads the single column, and the total label " & _
                    "is not written as a column header"
    Assert.AreEqual " ", CellText(sh, startRw + 1, STANDARD_START_COL + 1), _
                    "The placeholder under it is a space, so a walk left from the " & _
                    "end of the row stops here"
    Assert.AreEqual CLng(1), CLng(tabl.NumberOfColumns), _
                    "The fallback is one data column"
    Assert.AreEqual CLng(STANDARD_START_COL + 1), tabl.EndColumn, _
                    "The end column is that one column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddColumnsFallsBackToOneSummaryColumn", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a facility spatio-temporal table names the facility in its header.
'@details
'The composite header of a temporal table is what the navigation dropdown
'lists, and it is built from three labels. A facility spatio-temporal table
'takes the third from the main label of the prefixed facility variable.
'@TestMethod("CrossTable")
Public Sub TestAddColumnsBuildsTheFacilitySpatioTemporalHeader()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim header As String

    CustomTestSetTitles Assert, "CrossTable", "TestAddColumnsBuildsTheFacilitySpatioTemporalHeader"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRows(HF_VARIABLE, "3", "hf")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows
    tabl.AddColumns
    header = tabl.TimeSeriesHeader

    Assert.IsTrue (InStr(1, header, "Cases", vbBinaryCompare) > 0), _
                  "The composite header opens with the summary label. It reads " & header
    Assert.IsTrue (InStr(1, header, DictLabel(DATE_VARIABLE, "main label"), _
                         vbBinaryCompare) > 0), _
                  "The composite header names the time variable. It reads " & header
    Assert.IsTrue (InStr(1, header, DictLabel(HF_PREFIXED_VARIABLE, "main label"), _
                         vbBinaryCompare) > 0), _
                  "A facility table names the facility variable. It reads " & header

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddColumnsBuildsTheFacilitySpatioTemporalHeader", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a temporal table with no column variable names two labels only.
'@TestMethod("CrossTable")
Public Sub TestAddColumnsOmitsTheColumnLabelWhenThereIsNone()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim header As String

    CustomTestSetTitles Assert, "CrossTable", "TestAddColumnsOmitsTheColumnLabelWhenThereIsNone"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRowsNoColumn("no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows
    tabl.AddColumns
    header = tabl.TimeSeriesHeader

    Assert.AreEqual "Cases " & ChrW(9472) & " " & DictLabel(DATE_VARIABLE, "main label"), _
                    header, _
                    "With no column variable the header names the summary label and " & _
                    "the time variable and stops"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddColumnsOmitsTheColumnLabelWhenThereIsNone", _
                         Err.Number, Err.Description
End Sub

'@section Stage isolation — NameRanges
'===============================================================================

'@sub-title Verify the missing and total intersections are named.
'@details
'The four corner cells where a missing or total row crosses a missing or total
'column are named separately, because the formulas that fill them are different
'from the ones that fill the body.
'@TestMethod("CrossTable")
Public Sub TestNameRangesNamesTheMissingIntersections()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    CustomTestSetTitles Assert, "CrossTable", "TestNameRangesNamesTheMissingIntersections"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "all", "no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows
    tabl.AddColumns
    tabl.NameRanges
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "TOTAL_TOTAL_" & tabId), _
                  "The total row crossing the total column is named"
    Assert.IsTrue RangeExistsOnSheet(sh, "MISSING_TOTAL_" & tabId), _
                  "The missing row crossing the total column is named"
    Assert.IsTrue RangeExistsOnSheet(sh, "MISSING_MISSING_" & tabId), _
                  "The missing row crossing the missing column is named"
    Assert.IsTrue RangeExistsOnSheet(sh, "TOTAL_MISSING_" & tabId), _
                  "The total row crossing the missing column is named"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNameRangesNamesTheMissingIntersections", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the missing and total columns get their percentage twins.
'@TestMethod("CrossTable")
Public Sub TestNameRangesNamesThePercentageTwins()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim tabId As String

    CustomTestSetTitles Assert, "CrossTable", "TestNameRangesNamesThePercentageTwins"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "all", "row")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows
    tabl.AddColumns
    tabl.NameRanges
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "TOTAL_PERC_VALUES_" & tabId), _
                  "The total column carries a percentage twin"
    Assert.IsTrue RangeExistsOnSheet(sh, "MISSING_PERC_VALUES_" & tabId), _
                  "The missing column carries a percentage twin"
    Assert.AreEqual CLng(1), CLng(sh.Range("TOTAL_PERC_VALUES_" & tabId).Column - _
                                  sh.Range("TOTAL_COL_VALUES_" & tabId).Column), _
                    "A percentage twin sits one column right of the column it reports on"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNameRangesNamesThePercentageTwins", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the interior values range is named when it is a single cell.
'@details
'The interior range spans the first value column to the last, and it used to be
'assembled by splitting each column's address on a colon and pasting the two
'halves together. A one-cell range has no colon in its address, so the split
'yielded one element and the second of them raised.
'
'A univariate table reaches that shape with one category and no missing row: the
'value column is as tall as the category rows alone, since the total row is
'trimmed off. The interior range is not optional — Format and CrossTableFormula
'both read it, and twenty-odd names are already on the sheet by the time this
'line runs, so the table was left half-named.
'@TestMethod("CrossTable")
Public Sub TestNameRangesInteriorValuesWithSingleCellColumn()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim interior As Range
    Dim tabId As String
    Dim designFormat As LLFormat
    Dim errNumber As Long
    Dim errDescription As String

    CustomTestSetTitles Assert, "CrossTable", "TestNameRangesInteriorValuesWithSingleCellColumn"
    On Error GoTo TestFail

    lData.SetCategories ROW_CHOICE_VARIABLE, BetterArrayFromList("A")

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "VALUES_COL_1_" & tabId), _
                  "The single value column is named"
    Assert.AreEqual CLng(1), CLng(sh.Range("VALUES_COL_1_" & tabId).Rows.Count), _
                    "One category and no missing row leave a value column one cell tall"
    Assert.IsTrue RangeExistsOnSheet(sh, "INTERIOR_VALUES_" & tabId), _
                  "And the interior range is still named over it"

    Set interior = sh.Range("INTERIOR_VALUES_" & tabId)

    Assert.AreEqual CLng(1), CLng(interior.Rows.Count), _
                    "The interior range is one row tall"
    Assert.AreEqual CLng(1), CLng(interior.Columns.Count), _
                    "And one column wide"
    Assert.AreEqual sh.Range("VALUES_COL_1_" & tabId).Address(False, False), _
                    interior.Address(False, False), _
                    "So it is the value column itself"

    ' Format reads the interior range, and it is the reason the raise mattered:
    ' the table was left half-named and every later reader saw a missing name.
    Set designFormat = LLFormat.Create(PrepareLLFormatFixture(FORMAT_SHEET))
    On Error Resume Next
    tabl.Format designFormat
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(0), errNumber, _
                    "Formatting a table whose interior is one cell should raise " & _
                    "nothing, and it raised " & errNumber & ": " & errDescription

    lData.SetCategories ROW_CHOICE_VARIABLE, BetterArrayFromList("A", "B", "C")
    Exit Sub
TestFail:
    ' Read the error before restoring the categories. SetCategories closes its
    ' own On Error, and On Error GoTo 0 clears Err, so restoring first reports a
    ' blank failure.
    errNumber = Err.Number
    errDescription = Err.Description
    lData.SetCategories ROW_CHOICE_VARIABLE, BetterArrayFromList("A", "B", "C")
    CustomTestLogFailure Assert, "TestNameRangesInteriorValuesWithSingleCellColumn", _
                         errNumber, errDescription
End Sub

'@sub-title Verify the interior values range spans to the last percentage column.
'@details
'With percentages on, the interior reaches the percentage twin of the last data
'column rather than the data column itself, so the range is twice as wide as the
'column count. This is the other half of the same assembly and it travels with
'the single-cell case.
'@TestMethod("CrossTable")
Public Sub TestNameRangesInteriorValuesReachTheLastPercentageColumn()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim interior As Range
    Dim tabId As String

    CustomTestSetTitles Assert, "CrossTable", "TestNameRangesInteriorValuesReachTheLastPercentageColumn"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "row")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId
    Set interior = sh.Range("INTERIOR_VALUES_" & tabId)

    Assert.AreEqual sh.Range("VALUES_COL_1_" & tabId).Column, _
                    CLng(interior.Column), _
                    "The interior opens on the first value column"
    Assert.AreEqual sh.Range("PERC_COL_" & tabl.NumberOfColumns & "_" & tabId).Column, _
                    CLng(interior.Column + interior.Columns.Count - 1), _
                    "And closes on the percentage twin of the last data column"
    Assert.AreEqual CLng(3), CLng(interior.Rows.Count), _
                    "It is as tall as the three category rows"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNameRangesInteriorValuesReachTheLastPercentageColumn", _
                         Err.Number, Err.Description
End Sub

'@section Range access
'===============================================================================

'@sub-title Verify ColumnRange finds a column category and reports nothing otherwise.
'@TestMethod("CrossTable")
Public Sub TestColumnRangeFindsAColumnCategory()
    Dim tabl As CrossTable
    Dim rng As Range
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestColumnRangeFindsAColumnCategory"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Set rng = tabl.ColumnRange("X")

    Assert.IsTrue (Not rng Is Nothing), "ColumnRange should find a seeded category"
    Assert.AreEqual CLng(1), CLng(rng.Columns.Count), _
                    "A column range is one column wide"
    Assert.AreEqual CLng(STANDARD_START_COL + 1), CLng(rng.Column), _
                    "The first category sits one column right of the row labels"
    Assert.AreEqual tabl.EndRow, CLng(rng.Row + rng.Rows.Count - 1), _
                    "Left to itself a column range reaches the last row of the table"
    Assert.IsTrue (tabl.ColumnRange("not a category") Is Nothing), _
                  "ColumnRange should hand back nothing for a label it cannot find"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestColumnRangeFindsAColumnCategory", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify ColumnRange trims the total row when only categories are wanted.
'@TestMethod("CrossTable")
Public Sub TestColumnRangeTrimsTheTotalRow()
    Dim tabl As CrossTable
    Dim wholeColumn As Range
    Dim categoriesOnly As Range
    Dim withHeader As Range
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestColumnRangeTrimsTheTotalRow"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Set wholeColumn = tabl.ColumnRange("X")
    Set categoriesOnly = tabl.ColumnRange("X", onlyCategories:=True)
    Set withHeader = tabl.ColumnRange("X", onlyCategories:=True, includeHeaders:=True)

    Assert.AreEqual CLng(wholeColumn.Rows.Count - 1), CLng(categoriesOnly.Rows.Count), _
                    "Asking for categories only drops the total row"
    Assert.AreEqual CLng(categoriesOnly.Rows.Count + 1), CLng(withHeader.Rows.Count), _
                    "Asking for the header adds the label row above the categories"
    Assert.AreEqual CLng(categoriesOnly.Row - 1), CLng(withHeader.Row), _
                    "The header row is the row above the first category"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestColumnRangeTrimsTheTotalRow", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the row categories range skips the label row and the total row.
'@TestMethod("CrossTable")
Public Sub TestRowsCategoriesRangeTrimsBothEnds()
    Dim tabl As CrossTable
    Dim whole As Range
    Dim noHeader As Range
    Dim categoriesOnly As Range
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestRowsCategoriesRangeTrimsBothEnds"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)

    Set whole = tabl.RowsCategoriesRange()
    Set noHeader = tabl.RowsCategoriesRange(includeHeaders:=False)
    Set categoriesOnly = tabl.RowsCategoriesRange(includeHeaders:=False, _
                                                  onlyCategories:=True)

    Assert.AreEqual CLng(STANDARD_START_COL), CLng(whole.Column), _
                    "The row categories sit in the standard start column"
    Assert.AreEqual CLng(whole.Row + 1), CLng(noHeader.Row), _
                    "Dropping the header starts one row lower"
    Assert.AreEqual CLng(3), CLng(categoriesOnly.Rows.Count), _
                    "Three seeded categories are left once the label and total rows go"
    Assert.AreEqual "A|B|C", _
                    RowText(sh, categoriesOnly.Row, STANDARD_START_COL, STANDARD_START_COL) & "|" & _
                    RowText(sh, categoriesOnly.Row + 1, STANDARD_START_COL, STANDARD_START_COL) & "|" & _
                    RowText(sh, categoriesOnly.Row + 2, STANDARD_START_COL, STANDARD_START_COL), _
                    "The trimmed range holds the categories and nothing else"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRowsCategoriesRangeTrimsBothEnds", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify Previous answers nothing for the first table of a section.
'@details
'Previous hands back the table above this one in the same section, and the
'factory it calls rejects a specification of Nothing. A first table has no
'earlier specification, so this is answered before the factory is reached.
'@TestMethod("CrossTable")
Public Sub TestPreviousIsNothingForTheFirstTable()
    Dim tabl As CrossTable
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestPreviousIsNothingForTheFirstTable"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    Assert.IsTrue (tabl.Specifications.Previous Is Nothing), _
                  "The first row of the fixture has no earlier specification"
    Assert.IsTrue (tabl.Previous Is Nothing), _
                  "So the first table of a section has no previous table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreviousIsNothingForTheFirstTable", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify Previous carries the table above this one.
'@TestMethod("CrossTable")
Public Sub TestPreviousCarriesTheEarlierTable()
    Dim firstTable As CrossTable
    Dim secondTable As CrossTable
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestPreviousCarriesTheEarlierTable"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("no")
    Set sh = OutputSheet()
    Set firstTable = BuildTable(sh, 1)
    Set secondTable = NewTable(sh, 2)

    Assert.IsTrue (Not secondTable.Previous Is Nothing), _
                  "The second row of a section has a table above it"
    Assert.AreEqual firstTable.Specifications.TableId, _
                    secondTable.Previous.Specifications.TableId, _
                    "And that table is the one built from the first row"
    Assert.AreEqual firstTable.StartRow, secondTable.Previous.StartRow, _
                    "The two agree on the start row, which is how a section stays level"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreviousCarriesTheEarlierTable", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify the linelist data object is handed back as it was given.
'@details
'The parameter is typed Object so this class carries none of the linelist
'specification closure, and the members it needs are resolved at run time.
'@TestMethod("CrossTable")
Public Sub TestLinelistDataAnswersTheObjectItWasGiven()
    Dim tabl As CrossTable
    Dim sh As Worksheet

    CustomTestSetTitles Assert, "CrossTable", "TestLinelistDataAnswersTheObjectItWasGiven"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", "no")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    Assert.IsTrue (tabl.LinelistData Is lData), _
                  "The object handed to the factory is the object handed back"
    Assert.AreEqual "LinelistDataStub", TypeName(tabl.LinelistData), _
                    "And it is not narrowed to a linelist specification type on the way"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLinelistDataAnswersTheObjectItWasGiven", _
                         Err.Number, Err.Description
End Sub

'@sub-title Verify a temporal table reports no total when it names no column variable.
'@details
'A total column is a column, so HasTotal ends in "and a column variable was
'named" for the temporal and spatial scopes. The bivariate scope answers True
'with no such condition, which is why the fallback in the column builder has to
'drop the label itself rather than trust the flag. This pins the temporal half
'of that difference, so a later edit to HasTotal cannot quietly move the guard
'without a test naming it.
'@TestMethod("CrossTable")
Public Sub TestTemporalTableReportsNoTotalWithNoColumnVariable()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim startRw As Long

    CustomTestSetTitles Assert, "CrossTable", "TestTemporalTableReportsNoTotalWithNoColumnVariable"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRowsNoColumn("yes")
    Set sh = OutputSheet()
    Set tabl = NewTable(sh, 1)

    tabl.AddHeader
    tabl.AddRows
    startRw = tabl.StartRow
    tabl.AddColumns

    Assert.IsTrue tabl.Specifications.TotalRequested, _
                  "The row does ask for a total"
    Assert.IsTrue (Not tabl.Specifications.HasTotal), _
                  "But a temporal table with no column variable has no column for " & _
                  "the total to head, so it reports none"
    Assert.AreEqual "Cases", CellText(sh, startRw, STANDARD_START_COL + 1), _
                    "The summary label heads the single column"
    Assert.AreEqual CLng(1), CLng(tabl.NumberOfColumns), _
                    "The table is one data column wide"
    Assert.AreEqual CLng(STANDARD_START_COL + 1), tabl.EndColumn, _
                    "And the end column agrees with that count. The header row " & _
                    "reads " & RowText(sh, startRw, STANDARD_START_COL + 1, _
                                       STANDARD_START_COL + 3)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTemporalTableReportsNoTotalWithNoColumnVariable", _
                         Err.Number, Err.Description
End Sub

'@section Issue 338 - the total row of a temporal section
'===============================================================================
'@description
'A temporal section writes its Total and Missing labels whatever the tables of
'the section asked for, because the row block is a fixed grid and the category
'range is trimmed by count. So the answer to "Add total" is carried by the
'visibility of the row rather than by whether it was written, and these two
'tests are the pair that says so.

'@sub-title A temporal table that asked for no total hides the Total row.
'@details
'This is what the field reported in issue 338: "Add total" left empty and a
'Total row on the sheet anyway. The row is written by AddRows and hidden by
'Format, so the assertion has to come after both.
'@TestMethod("CrossTable")
Public Sub TestATemporalTableWithNoTotalHidesTheTotalRow()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim designFormat As LLFormat
    Dim tabId As String

    CustomTestSetTitles Assert, "CrossTable", "TestATemporalTableWithNoTotalHidesTheTotalRow"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId
    Set designFormat = LLFormat.Create(PrepareLLFormatFixture(FORMAT_SHEET))

    tabl.Format designFormat

    Assert.IsTrue (Not tabl.Specifications.TotalRequested), _
                  "The row leaves Add total empty, so no total was asked for"
    Assert.IsTrue RangeExistsOnSheet(sh, "TOTAL_ROW_" & tabId), _
                  "The Total row is still written, because the temporal row " & _
                  "block is a fixed grid trimmed by count"
    Assert.IsTrue sh.Range("TOTAL_ROW_" & tabId).EntireRow.Hidden, _
                  "And it is hidden, which is what issue 338 asked for"

    sh.Range("TOTAL_ROW_" & tabId).EntireRow.Hidden = False

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestATemporalTableWithNoTotalHidesTheTotalRow", _
                         Err.Number, Err.Description
End Sub

'@sub-title A temporal table that asked for a total shows the Total row.
'@details
'The other half of the same answer. Nothing hides a row a table has asked for,
'which is what keeps the result independent of the order a section is built in.
'@TestMethod("CrossTable")
Public Sub TestATemporalTableWithATotalShowsTheTotalRow()
    Dim tabl As CrossTable
    Dim sh As Worksheet
    Dim designFormat As LLFormat
    Dim tabId As String

    CustomTestSetTitles Assert, "CrossTable", "TestATemporalTableWithATotalShowsTheTotalRow"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), TimeSeriesRows("yes")
    Set sh = OutputSheet()
    Set tabl = BuildTable(sh, 1)
    tabId = tabl.Specifications.TableId
    Set designFormat = LLFormat.Create(PrepareLLFormatFixture(FORMAT_SHEET))

    tabl.Format designFormat

    Assert.IsTrue tabl.Specifications.TotalRequested, _
                  "The row asks for a total"
    Assert.IsTrue RangeExistsOnSheet(sh, "TOTAL_ROW_" & tabId), _
                  "So the Total row is written"
    Assert.IsTrue (Not sh.Range("TOTAL_ROW_" & tabId).EntireRow.Hidden), _
                  "And it is left visible"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestATemporalTableWithATotalShowsTheTotalRow", _
                         Err.Number, Err.Description
End Sub
