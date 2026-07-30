Attribute VB_Name = "TestCrossTableFormula"
Attribute VB_Description = "Tests for CrossTableFormula class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for CrossTableFormula class")

'@description
'Validates CrossTableFormula, which turns a built cross-table into Excel
'formulas. Six table scopes are covered, and what each test reads back is the
'formula text of a named range.
'
'THE FIXTURE CARRIES A LINELIST THE FORMULAS CAN POINT AT
'-------------------------------------------------------------------------------
'An analysis formula is a structured reference: COUNTIFS(ftable1[choi_v1], $C$10)
'names the filtered copy of a linelist table. So the fixture prepares the shared
'dictionary, which writes the "table name" column, then builds two ListObjects
'carrying that name and the same name with the "f" prefix, one column per
'variable the tests summarise. Excel accepts a formula against those and refuses
'one against a table that does not exist, which is exactly the judgement the
'class now relies on.
'
'THE FIXTURE IS A REAL LISTOBJECT
'-------------------------------------------------------------------------------
'TableSpecs reads the analysis scope from the name of the ListObject the
'specification row sits in, so every fixture wraps its header and its data rows
'in a ListObject carrying the name the setup workbook uses.
'
'DO NOT ASSERT A WHOLE FORMULA STRING
'-------------------------------------------------------------------------------
'A generated formula embeds cell addresses and table names, so an equality
'assertion on one churns on every fixture change and says little. These tests
'assert that a cell holds a formula, that the formula names the variable and the
'criteria it should, how many criteria it carries, and that the failure marker
'the class used to write is nowhere on the sheet.
'@depends CrossTableFormula, CrossTable, TableSpecs, AnalysisRanges, SpatialTables,
'  FormulaData, Formulas, LLdictionary, LLVariables, TranslationObject,
'  LinelistDataStub, Checking, BetterArray, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "CTFormulaFixture"
Private Const OUTPUT_SHEET As String = "CTFormulaOutput"
Private Const DICT_SHEET As String = "CTFormulaDict"
Private Const TRANS_SHEET As String = "CTFormulaTrans"
Private Const TRANS_TABLE As String = "T_CTFormulaTranslation"
Private Const TOKENS_SHEET As String = "CTFormulaTokens"
Private Const LINELIST_SHEET As String = "CTFormulaLinelist"
Private Const SPATIAL_SHEET As String = "spatial_tables__"

' The header row of every fixture table. Data rows start immediately below.
Private Const HEADER_ROW As Long = 5

' The blocks the two worksheets are reset over between tests. Both are bigger
' than anything this suite writes, and a bounded reset is what keeps the run
' inside the runner's cap.
Private Const OUTPUT_BLOCK As String = "A1:AN200"
Private Const FIXTURE_BLOCK As String = "A1:T30"

' The analysis ListObject names, spelled the way the setup workbook spells them.
Private Const TABLE_GLOBAL_SUMMARY As String = "Tab_Global_Summary"
Private Const TABLE_UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const TABLE_BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TABLE_TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const TABLE_SPATIAL As String = "Tab_Spatial_Analysis"
Private Const TABLE_SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"

' Variables of the shared dictionary fixture, and the three geo rows this suite
' appends to it. SpatialTableScopes probes hf_<var> and then adm1_<var>, and the
' spatial formulas read concat_adm1_<var>.
Private Const ROW_CHOICE_VARIABLE As String = "choi_v1"
Private Const COL_CHOICE_VARIABLE As String = "choi_ord_v1"
Private Const DATE_VARIABLE As String = "date_v1"
Private Const NUMBER_VARIABLE As String = "int_v1"
Private Const GEO_VARIABLE As String = "zone"
Private Const GEO_PREFIXED_VARIABLE As String = "adm1_zone"
Private Const GEO_CONCAT_VARIABLE As String = "concat_adm1_zone"
Private Const HF_VARIABLE As String = "center"
Private Const HF_PREFIXED_VARIABLE As String = "hf_center"

' The sheet the appended geo rows belong to, so they share one table name with
' the variables the rest of the suite summarises.
Private Const FIXTURE_SHEET_NAME As String = "vlist1D-sheet1"
Private Const FIXTURE_SHEET_TYPE As String = "vlist1D"

' Summary functions. The two count spellings take one formula shape and every
' other summary function takes the other, so both are exercised.
Private Const COUNT_CALL_FUNCTION As String = "N()"
Private Const COUNT_FUNCTION As String = "N"
Private Const SUM_FUNCTION As String = "SUM(int_v1)"
Private Const LINELIST_FUNCTION As String = "COUNT(choi_v1)"

' Eight summands, each expanding to a conditional sum over the linelist table,
' which puts the parsed formula past the 255 characters Range.FormulaArray
' accepts.
Private Const LONG_FUNCTION As String = "SUM(int_v1) + SUM(int_v1) + SUM(int_v1) + SUM(int_v1) + SUM(int_v1) + SUM(int_v1) + SUM(int_v1) + SUM(int_v1)"

' The number the single linelist record carries, so a summary function has
' something to add up and a wrongly entered formula answers wrongly.
Private Const RECORD_NUMBER As Long = 5

' The marker the class used to write into a cell whose formula it had declared
' broken. Nothing writes it now, and one test sweeps the whole output block.
Private Const FAILURE_MARKER As String = "formula parsing failed"

Private Assert As CustomTest
Private dict As LLdictionary
Private trans As TranslationObject
Private lData As LinelistDataStub
Private fData As FormulaData
Private linelistTable As String

'@section Fixture headers
'===============================================================================
'@description The six analysis header rows, spelled as the setup workbook has
'them. Measured for TestCrossTable and shared with it by copy, because a class
'keeps what it needs inside itself.

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
'@param summaryFunction String. The summary function of the row.
Private Function GlobalSummaryRows(ByVal summaryFunction As String) As Variant
    GlobalSummaryRows = Array(Array("A global summary", summaryFunction, "integer"))
End Function

'@sub-title One univariate row, as a row array.
'@details
'The row builder and the one-row wrapper are separate because one test needs
'five rows of different summary functions in a single fixture table, and
'indexing the result of a function call reads badly in VBA.
'@param rowVar String. The grouping variable.
'@param missing String. The "Add missing data" cell.
'@param summaryFunction String. The summary function.
'@param percentage String. The "Add percentage" cell.
Private Function UnivariateRow(ByVal rowVar As String, _
                               ByVal missing As String, _
                               ByVal summaryFunction As String, _
                               ByVal percentage As String) As Variant
    UnivariateRow = Array("S1", "A univariate table", rowVar, missing, _
                          summaryFunction, "Cases", "integer", percentage, "no", "no")
End Function

'@sub-title One univariate row.
'@param rowVar String. The grouping variable.
'@param missing String. The "Add missing data" cell.
'@param summaryFunction String. The summary function.
'@param percentage String. The "Add percentage" cell.
Private Function UnivariateRows(ByVal rowVar As String, _
                                ByVal missing As String, _
                                ByVal summaryFunction As String, _
                                ByVal percentage As String) As Variant
    UnivariateRows = Array(UnivariateRow(rowVar, missing, summaryFunction, percentage))
End Function

'@sub-title One bivariate row.
'@param missing String. The "Add missing data" cell: row, column, all or no.
'@param summaryFunction String. The summary function.
'@param percentage String. The "Add percentage" cell: row, column, total or no.
Private Function BivariateRows(ByVal missing As String, _
                               ByVal summaryFunction As String, _
                               ByVal percentage As String) As Variant
    BivariateRows = Array( _
        Array("S1", "A bivariate table", ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, _
              missing, summaryFunction, "Cases", "integer", percentage, "no", "no"))
End Function

'@sub-title One time series row.
'@param missing String. The "Add missing data" cell.
'@param summaryFunction String. The summary function.
'@param total String. The "Add total" cell.
Private Function TimeSeriesRows(ByVal missing As String, _
                                ByVal summaryFunction As String, _
                                ByVal total As String) As Variant
    TimeSeriesRows = Array( _
        Array("Series 1", "S1", DATE_VARIABLE, COL_CHOICE_VARIABLE, _
              "A time series table", missing, summaryFunction, "Cases", "integer", _
              "", total, "1"))
End Function

'@sub-title One spatial row.
'@param rowVar String. The unprefixed geo or facility variable.
'@param geoMax String. The "N geo max" cell.
'@param summaryFunction String. The summary function.
'@param percentage String. The "Add percentage" cell.
Private Function SpatialRows(ByVal rowVar As String, _
                             ByVal geoMax As String, _
                             ByVal summaryFunction As String, _
                             ByVal percentage As String) As Variant
    SpatialRows = Array( _
        Array("S1", "A spatial table", rowVar, geoMax, COL_CHOICE_VARIABLE, _
              "no", summaryFunction, "Cases", "integer", percentage, "no", "no"))
End Function

'@sub-title One spatio-temporal row.
'@param colVar String. The unprefixed geo or facility variable.
'@param geoMax String. The "N geo max" cell.
'@param summaryFunction String. The summary function.
Private Function SpatioTemporalRows(ByVal colVar As String, _
                                    ByVal geoMax As String, _
                                    ByVal summaryFunction As String) As Variant
    SpatioTemporalRows = Array( _
        Array("S1", DATE_VARIABLE, colVar, geoMax, "A spatio-temporal table", _
              "", summaryFunction, "Cases", "integer", "no"))
End Function

'@section Fixture helpers
'===============================================================================

'@sub-title Free an analysis table name wherever it is taken in the workbook.
'@details
'A ListObject name is unique across the workbook and the six analysis names are
'the only ones TableSpecs answers a scope for, so another suite's fixture sheet
'holding one blocks this one from taking it. Unlist turns the table back into an
'ordinary range and frees the name; the other suite rebuilds its own table on
'its next test.
'@param tableName String. The ListObject name to free.
Private Sub ReleaseTableName(ByVal tableName As String)
    Dim sh As Worksheet
    Dim idx As Long

    For Each sh In ThisWorkbook.Worksheets
        For idx = sh.ListObjects.Count To 1 Step -1
            If StrComp(sh.ListObjects(idx).Name, tableName, vbTextCompare) = 0 Then
                sh.ListObjects(idx).Unlist
            End If
        Next idx
    Next sh
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
    Dim idx As Long

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=False, visibility:=xlSheetHidden)

    For idx = sh.ListObjects.Count To 1 Step -1
        sh.ListObjects(idx).Unlist
    Next idx

    sh.Range(FIXTURE_BLOCK).Clear
    ReleaseTableName tableName

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
'@param dataRowIndex Long. The row to read.
Private Function FixtureDataRange(ByVal dataRowIndex As Long) As Range
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    Set FixtureDataRange = sh.ListObjects(1).ListRows(dataRowIndex).Range
End Function

'@sub-title Create a TableSpecs from a fixture data row index.
'@param dataRowIndex Long. The row to read.
Private Function CreateSpecs(ByVal dataRowIndex As Long) As TableSpecs
    Set CreateSpecs = TableSpecs.Create(FixtureHeaderRange(), _
                                        FixtureDataRange(dataRowIndex), _
                                        dict)
End Function

'@sub-title Delete every workbook-scoped name that points at one worksheet.
'@details
'Excel writes RefersTo as ='Sheet name'!$C$10 when the sheet name needs quoting
'and as =Sheetname!$C$10 when it does not. Both spellings are matched here.
'CrossTable creates its names with Cell.Name, which are workbook-scoped and
'outlive a Cells.Clear, and it asks whether a name exists to decide whether it
'has already built a piece of structure.
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
'The reset is bounded to OUTPUT_BLOCK and clearSheet is False. Running UnMerge,
'the two unhide writes and Clear over every cell of a sheet costs about seven
'seconds a test, which took a fully green run past the runner's cap once
'already.
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
Private Function BuildTranslator() As TranslationObject
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim translationRows As Variant
    Dim idx As Long

    Set sh = EnsureWorksheet(TRANS_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    For idx = sh.ListObjects.Count To 1 Step -1
        sh.ListObjects(idx).Unlist
    Next idx

    ReleaseTableName TRANS_TABLE

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

'@sub-title Build the two linelist tables the generated formulas reference.
'@details
'The analysis formulas name the linelist table of each variable and the
'filtered copy of it, which carries the "f" prefix. Both are built here with
'one column per variable this suite summarises, so Excel resolves the
'structured references and accepts the formulas.
Private Sub BuildLinelistTables()
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim columnNames As Variant
    Dim dataRow As Variant
    Dim idx As Long

    linelistTable = LLVariables.Create(dict).Value(colName:="table name", _
                                                   varName:=ROW_CHOICE_VARIABLE)

    Set sh = EnsureWorksheet(LINELIST_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    ReleaseTableName linelistTable
    ReleaseTableName "f" & linelistTable

    columnNames = Array(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, DATE_VARIABLE, _
                        NUMBER_VARIABLE, GEO_PREFIXED_VARIABLE, _
                        GEO_CONCAT_VARIABLE, HF_PREFIXED_VARIABLE)

    ' One record, and its values are the ones the tests group by: the first
    ' category of the row variable, the first category of the column variable and
    ' a number a summary function can add up. A formula that computes the wrong
    ' answer is the only way to see that a formula was entered the wrong way.
    dataRow = Array("A", "X", DateSerial(2026, 1, 15), RECORD_NUMBER, _
                    "Z1", "Z1", "C1")

    WriteMatrix sh.Cells(1, 1), RowsToMatrix(Array(columnNames))
    WriteMatrix sh.Cells(4, 1), RowsToMatrix(Array(columnNames))

    For idx = LBound(columnNames) To UBound(columnNames)
        sh.Cells(2, idx + 1).Value = dataRow(idx)
        sh.Cells(5, idx + 1).Value = dataRow(idx)
    Next idx

    Set lo = sh.ListObjects.Add(xlSrcRange, sh.Range(sh.Cells(1, 1), _
                                sh.Cells(2, UBound(columnNames) + 1)), , xlYes)
    lo.Name = linelistTable

    Set lo = sh.ListObjects.Add(xlSrcRange, sh.Range(sh.Cells(4, 1), _
                                sh.Cells(5, UBound(columnNames) + 1)), , xlYes)
    lo.Name = "f" & linelistTable
End Sub

'@sub-title Build one cross-table and its formula writer from a fixture row.
'@param dataRowIndex Long. The 1-based fixture data row.
'@param sh Worksheet. ByRef. Filled with the output worksheet.
'@param tabId String. ByRef. Filled with the table identifier.
'@param secId String. ByRef. Filled with the section identifier.
'@return CrossTableFormula. The writer, with AddFormulas already run.
Private Function WriteTable(ByVal dataRowIndex As Long, _
                            ByRef sh As Worksheet, _
                            ByRef tabId As String, _
                            ByRef secId As String) As CrossTableFormula
    Dim specs As TableSpecs
    Dim tabl As CrossTable
    Dim writer As CrossTableFormula

    Set sh = OutputSheet()
    Set specs = CreateSpecs(dataRowIndex)
    tabId = specs.TableId()
    secId = specs.TableSectionId()

    Set tabl = CrossTable.Create(specs, sh, lData)
    tabl.Build

    Set writer = CrossTableFormula.Create(tabl, fData)
    writer.AddFormulas
    Set WriteTable = writer
End Function

'@sub-title Build a cross-table and a writer without running AddFormulas.
'@param dataRowIndex Long. The 1-based fixture data row.
Private Function NewWriter(ByVal dataRowIndex As Long) As CrossTableFormula
    Dim tabl As CrossTable

    Set tabl = CrossTable.Create(CreateSpecs(dataRowIndex), OutputSheet(), lData)
    tabl.Build
    Set NewWriter = CrossTableFormula.Create(tabl, fData)
End Function

'@sub-title Create a writer over a cross-table that was never built.
'@details
'Valid reads the specification row and the token tables alone, so the layout
'does not have to exist. Building five tables of the same section one after
'another on a sheet reset between each would leave the second one looking for a
'marker the reset had removed.
'@param dataRowIndex Long. The 1-based fixture data row.
Private Function NewUnbuiltWriter(ByVal dataRowIndex As Long) As CrossTableFormula
    Dim tabl As CrossTable

    Set tabl = CrossTable.Create(CreateSpecs(dataRowIndex), _
                                 ThisWorkbook.Worksheets(OUTPUT_SHEET), lData)
    Set NewUnbuiltWriter = CrossTableFormula.Create(tabl, fData)
End Function

'@sub-title Resolve a named range on one worksheet, or answer Nothing.
'@param sh Worksheet. The worksheet to resolve against.
'@param rngName String. The name to look up.
Private Function NamedRange(ByVal sh As Worksheet, ByVal rngName As String) As Range
    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    Set NamedRange = rng
End Function

'@sub-title Read the formula of the first cell of a named range.
'@details
'An absent name answers an empty string. A test asserting on a formula then
'reports a missing formula, and the read itself raises nothing.
'@param sh Worksheet. The worksheet to read.
'@param rngName String. The name to look up.
Private Function NamedFormula(ByVal sh As Worksheet, ByVal rngName As String) As String
    Dim rng As Range

    Set rng = NamedRange(sh, rngName)
    If rng Is Nothing Then Exit Function
    NamedFormula = CStr(rng.Cells(1, 1).Formula)
End Function

'@sub-title Read the formula of one cell.
'@param sh Worksheet. The worksheet to read.
'@param rw Long. The row to read.
'@param col Long. The column to read.
Private Function CellFormula(ByVal sh As Worksheet, ByVal rw As Long, _
                             ByVal col As Long) As String
    CellFormula = CStr(sh.Cells(rw, col).Formula)
End Function

'@sub-title Whether a cell holds a formula.
'@param formulaText String. The text read back from the cell.
Private Function IsFormula(ByVal formulaText As String) As Boolean
    IsFormula = (Left$(formulaText, 1) = "=")
End Function

'@sub-title Count how many times one piece of text appears in another.
'@param haystack String. The text to search.
'@param needle String. The text to count.
Private Function OccurrenceCount(ByVal haystack As String, ByVal needle As String) As Long
    Dim position As Long
    Dim total As Long

    If LenB(needle) = 0 Then Exit Function

    position = InStr(1, haystack, needle, vbTextCompare)
    Do While position > 0
        total = total + 1
        position = InStr(position + Len(needle), haystack, needle, vbTextCompare)
    Loop

    OccurrenceCount = total
End Function

'@sub-title Read the list a time unit dropdown was bound to.
'@param sh Worksheet. The output worksheet.
'@param tabId String. The table identifier.
Private Function TimeUnitValidationOf(ByVal sh As Worksheet, ByVal tabId As String) As String
    Dim unitRng As Range

    Set unitRng = NamedRange(sh, "TIME_UNIT_" & tabId)
    If unitRng Is Nothing Then Exit Function

    On Error Resume Next
    TimeUnitValidationOf = unitRng.Validation.Formula1
    On Error GoTo 0
End Function

'@sub-title Whether the whole output block holds the old failure marker anywhere.
'@details
'One Find call does the sweep, because a walk over eight thousand cells costs
'far more. Every option is pinned: Range.Find inherits LookIn, LookAt and
'SearchOrder from the last search of the Excel session, including one a user ran
'from the Find dialog.
'@param sh Worksheet. The worksheet to sweep.
Private Function MarkerFound(ByVal sh As Worksheet) As Boolean
    Dim found As Range

    On Error Resume Next
    Set found = sh.Range(OUTPUT_BLOCK).Find(What:=FAILURE_MARKER, _
                                            LookIn:=xlValues, LookAt:=xlPart, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlNext, _
                                            MatchCase:=False)
    On Error GoTo 0
    MarkerFound = Not (found Is Nothing)
End Function

'@sub-title Join every message the writer filed into one string.
'@param writer CrossTableFormula. The writer to read.
Private Function CheckMessages(ByVal writer As CrossTableFormula) As String
    Dim checks As Checking
    Dim keyList As BetterArray
    Dim counter As Long
    Dim joined As String

    If Not writer.HasCheckings Then Exit Function

    Set checks = writer.CheckingValues
    If checks.Length = 0 Then Exit Function

    Set keyList = checks.ListOfKeys
    For counter = keyList.LowerBound To keyList.UpperBound
        joined = joined & "[" & checks.ValueOf(CStr(keyList.Item(counter)), checkingLabel) & "]"
    Next counter

    CheckMessages = joined
End Function

'@sub-title Count the entries the writer filed with the error scope.
'@param writer CrossTableFormula. The writer to read.
Private Function ErrorCheckCount(ByVal writer As CrossTableFormula) As Long
    Dim checks As Checking
    Dim keyList As BetterArray
    Dim counter As Long
    Dim total As Long

    If Not writer.HasCheckings Then Exit Function

    Set checks = writer.CheckingValues
    If checks.Length = 0 Then Exit Function

    Set keyList = checks.ListOfKeys
    For counter = keyList.LowerBound To keyList.UpperBound
        If InStr(1, checks.ValueOf(CStr(keyList.Item(counter)), checkingType), _
                 "Error", vbTextCompare) > 0 Then
            total = total + 1
        End If
    Next counter

    ErrorCheckCount = total
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Build the dictionary, the translator, the token tables and the linelist.
'@details
'The dictionary fixture is prepared once and Prepare is called on it, because
'Prepare is what writes the "table name" column every analysis formula reads.
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    Dim sh As Worksheet
    Dim appendRow As Long

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCrossTableFormula"

    PrepareDictionaryFixture DICT_SHEET

    ' adm1_zone gives the spatial scope a "geo" answer to find, hf_center gives
    ' it an "hf" answer, and concat_adm1_zone is the variable the geographic
    ' formulas summarise. All three carry the sheet of the variables the rest of
    ' the suite uses, so one linelist table serves every formula.
    Set sh = ThisWorkbook.Worksheets(DICT_SHEET)
    appendRow = 1 + DictionaryFixtureRowCount() + 1
    AppendGeoRow sh, appendRow, GEO_PREFIXED_VARIABLE, "Zone", "Zones"
    AppendGeoRow sh, appendRow + 1, GEO_CONCAT_VARIABLE, "Zone code", "Zones"
    AppendGeoRow sh, appendRow + 2, HF_PREFIXED_VARIABLE, "Health centre", vbNullString

    Set dict = LLdictionary.Create(sh, 1, 1)
    dict.Prepare

    Set trans = BuildTranslator()

    Set lData = New LinelistDataStub
    lData.SetTransObject trans
    lData.SetCategories ROW_CHOICE_VARIABLE, BetterArrayFromList("A", "B", "C")
    lData.SetCategories COL_CHOICE_VARIABLE, BetterArrayFromList("X", "Y")

    ' FormulaData resolves its two lookup tables by fixed names, so another
    ' suite's fixture sheet holding them blocks this one from taking them.
    ReleaseTableName "T_XlsFonctions"
    ReleaseTableName "T_ascii"
    Set fData = FormulaData.Create(PrepareFormulaFixtureSheet(TOKENS_SHEET))

    BuildLinelistTables

    ' SpatialTables.Create requires this worksheet in the same workbook, and the
    ' spatial arm registers every table it writes with it.
    EnsureWorksheet SPATIAL_SHEET, clearSheet:=True, visibility:=xlSheetHidden
End Sub

'@sub-title Write one dictionary row for an appended geo variable.
'@param sh Worksheet. The dictionary worksheet.
'@param rowNum Long. The row to write.
'@param varName String. The variable name.
'@param mainLabel String. The main label.
'@param subSection String. The sub section, which is where an administrative
'level carries its label.
Private Sub AppendGeoRow(ByVal sh As Worksheet, ByVal rowNum As Long, _
                         ByVal varName As String, ByVal mainLabel As String, _
                         ByVal subSection As String)
    sh.Cells(rowNum, 1).Value = varName
    sh.Cells(rowNum, 2).Value = mainLabel
    sh.Cells(rowNum, 7).Value = FIXTURE_SHEET_NAME
    sh.Cells(rowNum, 8).Value = FIXTURE_SHEET_TYPE
    sh.Cells(rowNum, 10).Value = subSection
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
    DeleteWorksheet TOKENS_SHEET
    DeleteWorksheet LINELIST_SHEET
    DeleteWorksheet SPATIAL_SHEET
    RestoreApp

    Set dict = Nothing
    Set trans = Nothing
    Set lData = Nothing
    Set fData = Nothing
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

'@section Fixture and factory
'===============================================================================

'@sub-title Verify the fixture summary functions parse in the analysis context.
'@details
'Every formula test below depends on Valid answering True, because AddFormulas
'exits at once when it answers False. This test localises a fixture fault: a
'summary function the parser rejects would otherwise show up as forty missing
'formulas spread over a dozen tests.
'@TestMethod("CrossTableFormula")
Public Sub TestFixtureSummaryFunctionsAreValid()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestFixtureSummaryFunctionsAreValid"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 Array(UnivariateRow(ROW_CHOICE_VARIABLE, "no", COUNT_CALL_FUNCTION, "no"), _
                       UnivariateRow(ROW_CHOICE_VARIABLE, "no", COUNT_FUNCTION, "no"), _
                       UnivariateRow(ROW_CHOICE_VARIABLE, "no", SUM_FUNCTION, "no"), _
                       UnivariateRow(ROW_CHOICE_VARIABLE, "no", LONG_FUNCTION, "no"), _
                       UnivariateRow(ROW_CHOICE_VARIABLE, "no", "InvalidFunc", "no"))

    OutputSheet

    Assert.IsTrue NewUnbuiltWriter(1).Valid, "N() should parse in the analysis context"
    Assert.IsTrue NewUnbuiltWriter(2).Valid, "N should parse in the analysis context"
    Assert.IsTrue NewUnbuiltWriter(3).Valid, _
                  "SUM over a variable should parse in the analysis context"
    Assert.IsTrue NewUnbuiltWriter(4).Valid, "The long summary function should parse"
    Assert.IsFalse NewUnbuiltWriter(5).Valid, "An unknown summary function should not parse"

    Assert.LogSuccesses "The linelist table of the fixture is [" & linelistTable & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFixtureSummaryFunctionsAreValid", Err.Number, Err.Description
End Sub

'@sub-title Verify Create rejects Nothing for the cross-table parameter.
'@TestMethod("CrossTableFormula")
Public Sub TestCreateRejectsNothingTable()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestCreateRejectsNothingTable"
    On Error GoTo TestFail

    Dim writer As CrossTableFormula
    Dim errNumber As Long

    On Error Resume Next
    Set writer = CrossTableFormula.Create(Nothing, fData)
    errNumber = Err.Number
    On Error GoTo 0

    Assert.IsTrue (writer Is Nothing), "Create with Nothing cross-table should fail"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A missing cross-table should raise InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingTable", Err.Number, Err.Description
End Sub

'@sub-title Verify Create rejects Nothing for the formula-data parameter.
'@TestMethod("CrossTableFormula")
Public Sub TestCreateRejectsNothingFormulaData()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestCreateRejectsNothingFormulaData"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", COUNT_CALL_FUNCTION, "no")

    Dim tabl As CrossTable
    Set tabl = CrossTable.Create(CreateSpecs(1), OutputSheet(), lData)
    tabl.Build

    Dim writer As CrossTableFormula
    Dim errNumber As Long

    On Error Resume Next
    Set writer = CrossTableFormula.Create(tabl, Nothing)
    errNumber = Err.Number
    On Error GoTo 0

    Assert.IsTrue (writer Is Nothing), "Create with Nothing formula data should fail"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "Missing formula data should raise InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingFormulaData", Err.Number, Err.Description
End Sub

'@sub-title Verify a sealed instance refuses a second setup write.
'@TestMethod("CrossTableFormula")
Public Sub TestFormulaSetupWriteAfterCreateRaises()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestFormulaSetupWriteAfterCreateRaises"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", COUNT_CALL_FUNCTION, "no")

    Dim writer As CrossTableFormula
    Dim errNumber As Long

    Set writer = NewWriter(1)
    Assert.IsTrue (Not writer Is Nothing), "Create with valid arguments should succeed"

    On Error Resume Next
    Set writer.formData = fData
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.SomethingWentWrong), errNumber, _
                    "A setup write after Create should raise"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFormulaSetupWriteAfterCreateRaises", Err.Number, Err.Description
End Sub

'@section Global summary
'===============================================================================

'@sub-title Verify both global summary columns hold a linelist formula.
'@details
'The first column counts the whole linelist and the second the filtered copy of
'it, which is the one carrying the "f" prefix. That prefix is the only
'difference between the two formulas.
'@TestMethod("CrossTableFormula")
Public Sub TestGlobalSummaryWritesBothColumns()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestGlobalSummaryWritesBothColumns"
    On Error GoTo TestFail

    BuildFixture TABLE_GLOBAL_SUMMARY, GlobalSummaryHeader(), _
                 GlobalSummaryRows(LINELIST_FUNCTION)

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim startRng As Range
    Dim allData As String
    Dim filtered As String

    Set writer = WriteTable(1, sh, tabId, secId)
    Set startRng = NamedRange(sh, "STARTROW_" & tabId)

    Assert.IsTrue (Not startRng Is Nothing), "The global summary row should be named"

    allData = CStr(startRng.Cells(1, 2).Formula)
    filtered = CStr(startRng.Cells(1, 3).Formula)

    Assert.IsTrue IsFormula(allData), _
                  "The all-data cell should hold a formula, and it holds [" & allData & "]"
    Assert.IsTrue IsFormula(filtered), _
                  "The filtered cell should hold a formula, and it holds [" & filtered & "]"
    Assert.IsTrue (InStr(1, allData, linelistTable & "[" & ROW_CHOICE_VARIABLE & "]") > 0), _
                  "The all-data formula should name the linelist table"
    Assert.IsTrue (InStr(1, filtered, "f" & linelistTable & "[") > 0), _
                  "The filtered formula should name the filtered linelist table"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A global summary table should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGlobalSummaryWritesBothColumns", Err.Number, Err.Description
End Sub

'@section Univariate
'===============================================================================

'@sub-title Verify the univariate value column is written and copied down.
'@TestMethod("CrossTableFormula")
Public Sub TestUnivariateWritesTheValueColumn()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestUnivariateWritesTheValueColumn"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim valRng As Range
    Dim firstFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    Set valRng = NamedRange(sh, "VALUES_COL_1_" & tabId)

    Assert.IsTrue (Not valRng Is Nothing), "The value column should be named"
    Assert.AreEqual CLng(3), valRng.Rows.Count, "Three categories give three value rows"

    firstFormula = CStr(valRng.Cells(1, 1).Formula)
    Assert.IsTrue IsFormula(firstFormula), _
                  "The first value cell should hold a formula, and it holds [" & _
                  firstFormula & "]"
    Assert.IsTrue (InStr(1, firstFormula, "COUNTIFS") > 0), _
                  "A count summary function should produce a COUNTIFS"
    Assert.IsTrue (InStr(1, firstFormula, "f" & linelistTable & "[" & _
                         ROW_CHOICE_VARIABLE & "]") > 0), _
                  "The formula should name the row variable of the filtered table"
    Assert.IsTrue IsFormula(CStr(valRng.Cells(3, 1).Formula)), _
                  "The last value cell should hold the copied formula"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A plain univariate table should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateWritesTheValueColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify the univariate total row counts the non-empty records.
'@TestMethod("CrossTableFormula")
Public Sub TestUnivariateWritesTheTotalRow()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestUnivariateWritesTheTotalRow"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim totalFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    totalFormula = NamedFormula(sh, "TOTAL_ROW_VALUES_" & tabId)

    Assert.IsTrue IsFormula(totalFormula), _
                  "The total row should hold a formula, and it holds [" & totalFormula & "]"
    Assert.IsTrue (InStr(1, totalFormula, "<>") > 0), _
                  "A table with no missing row totals the non-empty records"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateWritesTheTotalRow", Err.Number, Err.Description
End Sub

'@sub-title Verify the missing row and the total that adds both counts.
'@details
'With a missing row on a count table the total is the missing count plus the
'non-empty count, so the cell holds two COUNTIFS joined by a plus.
'@TestMethod("CrossTableFormula")
Public Sub TestUnivariateWithMissingAddsBothCounts()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestUnivariateWithMissingAddsBothCounts"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "yes", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim missingFormula As String
    Dim totalFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    missingFormula = NamedFormula(sh, "MISSING_ROW_VALUES_" & tabId)
    totalFormula = NamedFormula(sh, "TOTAL_ROW_VALUES_" & tabId)

    Assert.IsTrue IsFormula(missingFormula), _
                  "The missing row should hold a formula, and it holds [" & _
                  missingFormula & "]"
    Assert.IsTrue IsFormula(totalFormula), _
                  "The total row should hold a formula, and it holds [" & totalFormula & "]"
    Assert.AreEqual CLng(2), OccurrenceCount(totalFormula, "COUNTIFS"), _
                    "The total of a count table with a missing row adds two counts"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "The table should report no error, and it reported " & CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateWithMissingAddsBothCounts", Err.Number, Err.Description
End Sub

'@sub-title Verify the univariate percentage column and its total cell.
'@TestMethod("CrossTableFormula")
Public Sub TestUnivariateWithPercentageWritesTheTwin()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestUnivariateWithPercentageWritesTheTwin"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", COUNT_CALL_FUNCTION, "yes")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim percRng As Range
    Dim percFormula As String
    Dim totalRng As Range

    Set writer = WriteTable(1, sh, tabId, secId)
    Set percRng = NamedRange(sh, "PERC_COL_1_" & tabId)

    Assert.IsTrue (Not percRng Is Nothing), "The percentage column should be named"

    percFormula = CStr(percRng.Cells(1, 1).Formula)
    Assert.IsTrue (InStr(1, percFormula, "ISERR") > 0), _
                  "A percentage is guarded against a division error, and the cell holds [" & _
                  percFormula & "]"
    Assert.IsTrue IsFormula(CStr(percRng.Cells(3, 1).Formula)), _
                  "The percentage column should be copied down"

    Set totalRng = NamedRange(sh, "TOTAL_ROW_VALUES_" & tabId)
    Assert.IsTrue (InStr(1, CStr(totalRng.Cells(1, 2).Formula), "ISERR") > 0), _
                  "The total row carries its own percentage cell"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnivariateWithPercentageWritesTheTwin", Err.Number, Err.Description
End Sub

'@sub-title Verify a value column one cell tall still gets its formula.
'@details
'AutoFill needs a destination taller than one cell and raises otherwise. A
'univariate table with a single category has a value column of one cell, so the
'copy is skipped and the formula still has to be there.
'@TestMethod("CrossTableFormula")
Public Sub TestUnivariateWithOneCategoryWritesItsFormula()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestUnivariateWithOneCategoryWritesItsFormula"
    On Error GoTo TestFail

    lData.SetCategories ROW_CHOICE_VARIABLE, BetterArrayFromList("A")

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim valRng As Range
    Dim valueFormula As String
    Dim errCount As Long
    Dim messages As String

    Set writer = WriteTable(1, sh, tabId, secId)
    Set valRng = NamedRange(sh, "VALUES_COL_1_" & tabId)
    valueFormula = CStr(valRng.Cells(1, 1).Formula)
    errCount = ErrorCheckCount(writer)
    messages = CheckMessages(writer)

    lData.SetCategories ROW_CHOICE_VARIABLE, BetterArrayFromList("A", "B", "C")

    Assert.AreEqual CLng(1), valRng.Rows.Count, "One category gives a value column of one cell"
    Assert.IsTrue IsFormula(valueFormula), _
                  "The single value cell should hold a formula, and it holds [" & _
                  valueFormula & "]"
    Assert.AreEqual CLng(0), errCount, _
                    "A one-cell value column should report no error, and it reported " & messages

    Exit Sub
TestFail:
    lData.SetCategories ROW_CHOICE_VARIABLE, BetterArrayFromList("A", "B", "C")
    CustomTestLogFailure Assert, "TestUnivariateWithOneCategoryWritesItsFormula", Err.Number, Err.Description
End Sub

'@section Long formulas and the failure marker
'===============================================================================

'@sub-title Verify a formula over 255 characters reaches its cell.
'@details
'This is the fault the session exists for. Range.FormulaArray refuses a formula
'longer than 255 characters and Application.Evaluate refuses to even read one,
'so a long formula used to be declared broken and a text marker was written into
'the table in its place. The cell has to hold the formula, and it has to stay an
'array formula.
'@TestMethod("CrossTableFormula")
Public Sub TestALongFormulaReachesItsCell()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestALongFormulaReachesItsCell"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", LONG_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim valRng As Range
    Dim valueFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    Set valRng = NamedRange(sh, "VALUES_COL_1_" & tabId)
    valueFormula = CStr(valRng.Cells(1, 1).Formula)

    Assert.LogSuccesses "The long value formula is " & Len(valueFormula) & " characters"

    Assert.IsTrue (Len(valueFormula) > 255), _
                  "The formula under test has to be longer than 255 characters"
    Assert.IsTrue IsFormula(valueFormula), _
                  "The cell should hold a formula, and it holds [" & _
                  Left$(valueFormula, 120) & "]"
    Assert.LogSuccesses "The long value cell reports HasArray = " & _
                        valRng.Cells(1, 1).HasArray & " and the writer said " & _
                        CheckMessages(writer)
    Assert.IsTrue (OccurrenceCount(valueFormula, "f" & linelistTable & "[") >= 8), _
                  "Each of the eight summands names the filtered linelist table, and " & _
                  "the formula names it " & _
                  OccurrenceCount(valueFormula, "f" & linelistTable & "[") & " times"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A long formula should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALongFormulaReachesItsCell", Err.Number, Err.Description
End Sub

'@sub-title Verify a long formula lands in every cell of the table that needs one.
'@details
'The value cell and the total cell of the same table both carry the long formula,
'so a fault in the long path shows up twice. This test reads the total cell,
'because the value cell above reads the one the column loop writes.
'@TestMethod("CrossTableFormula")
Public Sub TestALongFormulaReachesTheTotalCell()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestALongFormulaReachesTheTotalCell"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", LONG_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim totalFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    totalFormula = NamedFormula(sh, "TOTAL_ROW_VALUES_" & tabId)

    Assert.LogSuccesses "The long total formula is " & Len(totalFormula) & " characters"

    Assert.IsTrue IsFormula(totalFormula), _
                  "The total cell should hold a formula, and it holds [" & _
                  Left$(totalFormula, 120) & "]"
    Assert.IsTrue (Len(totalFormula) > 255), _
                  "The total formula of this table is also over 255 characters"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A long total should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALongFormulaReachesTheTotalCell", Err.Number, Err.Description
End Sub

'@sub-title Verify a long formula computes the same answer as a short one.
'@details
'This is the test that says whether the way a long formula is entered matters.
'The fixture linelist holds one record whose number column is known, so the
'short summary function has one right answer and the long one, which is that
'summary function eight times over, has eight times it.
'
'Range.FormulaArray refuses a formula over 255 characters and Range.Replace
'refuses the swap on this host, so the long one goes in as an ordinary formula.
'If that entry changed the arithmetic, the two answers would disagree here.
'@TestMethod("CrossTableFormula")
Public Sub TestALongFormulaComputesTheSameAnswer()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestALongFormulaComputesTheSameAnswer"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim shortValue As Variant
    Dim longValue As Variant

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", SUM_FUNCTION, "no")
    Set writer = WriteTable(1, sh, tabId, secId)
    Application.Calculate
    shortValue = NamedRange(sh, "VALUES_COL_1_" & tabId).Cells(1, 1).Value

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", LONG_FUNCTION, "no")
    Set writer = WriteTable(1, sh, tabId, secId)
    Application.Calculate
    longValue = NamedRange(sh, "VALUES_COL_1_" & tabId).Cells(1, 1).Value

    Assert.LogSuccesses "The short summary answers [" & CStr(shortValue) & "] and the " & _
                        "long one answers [" & CStr(longValue) & "]"

    Assert.IsFalse IsError(shortValue), _
                   "The short summary function should compute a value"
    Assert.IsFalse IsError(longValue), _
                   "The long summary function should compute a value"
    Assert.AreEqual CDbl(RECORD_NUMBER), CDbl(shortValue), _
                    "One record of " & RECORD_NUMBER & " gives that as the summary"
    Assert.AreEqual CDbl(8 * RECORD_NUMBER), CDbl(longValue), _
                    "Eight summands of the same record give eight times the summary"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALongFormulaComputesTheSameAnswer", Err.Number, Err.Description
End Sub

'@sub-title Verify the old failure marker reaches no cell of a full table.
'@details
'The marker was written by the class itself, so nothing on the sheet can carry
'it now. A bivariate table with a missing row, a missing column, totals and
'percentages is the widest set of writes the class makes.
'@TestMethod("CrossTableFormula")
Public Sub TestNoFailureMarkerIsWritten()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestNoFailureMarkerIsWritten"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows("all", SUM_FUNCTION, "column")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula

    Set writer = WriteTable(1, sh, tabId, secId)

    Assert.IsFalse MarkerFound(sh), _
                   "No cell of the output block should carry the failure marker"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A fully featured bivariate table should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNoFailureMarkerIsWritten", Err.Number, Err.Description
End Sub

'@section Bivariate
'===============================================================================

'@sub-title Verify one formula per data column, each naming its own category.
'@TestMethod("CrossTableFormula")
Public Sub TestBivariateWritesOneFormulaPerDataColumn()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestBivariateWritesOneFormulaPerDataColumn"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows("no", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim firstFormula As String
    Dim secondFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    firstFormula = NamedFormula(sh, "VALUES_COL_1_" & tabId)
    secondFormula = NamedFormula(sh, "VALUES_COL_2_" & tabId)

    Assert.IsTrue IsFormula(firstFormula), _
                  "The first data column should hold a formula, and it holds [" & _
                  firstFormula & "]"
    Assert.IsTrue IsFormula(secondFormula), _
                  "The second data column should hold a formula, and it holds [" & _
                  secondFormula & "]"
    Assert.IsTrue (firstFormula <> secondFormula), _
                  "Each data column carries the address of its own column label"
    Assert.IsTrue (NamedRange(sh, "VALUES_COL_3_" & tabId) Is Nothing), _
                  "Two column categories give two data columns"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A plain bivariate table should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariateWritesOneFormulaPerDataColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify the total column carries the non-empty test.
'@details
'CrossTable.EndColumn counts the total and missing columns, so a loop walking to
'it wrote a value formula into the total column carrying the criteria of a data
'column. The dedicated block below the loop then overwrote it, which is why the
'sheet looked right. The loop stops at the data columns now, and this test pins
'what the total column holds: the non-empty test on the column variable.
'@TestMethod("CrossTableFormula")
Public Sub TestBivariateTotalColumnCarriesTheNonEmptyTest()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestBivariateTotalColumnCarriesTheNonEmptyTest"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows("no", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim totalFormula As String
    Dim totalHeader As Range

    Set writer = WriteTable(1, sh, tabId, secId)
    totalFormula = NamedFormula(sh, "TOTAL_COL_VALUES_" & tabId)

    Assert.IsTrue IsFormula(totalFormula), _
                  "The total column should hold a formula, and it holds [" & totalFormula & "]"
    Assert.IsTrue (InStr(1, totalFormula, "<>") > 0), _
                  "The total column tests the column variable for a value"

    Set totalHeader = NamedRange(sh, "TOTAL_LABEL_COL_" & tabId)
    If totalHeader Is Nothing Then Set totalHeader = NamedRange(sh, "TOTAL_COL_" & tabId)

    Assert.IsTrue (Not totalHeader Is Nothing), "The total column should be named"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariateTotalColumnCarriesTheNonEmptyTest", Err.Number, Err.Description
End Sub

'@sub-title Verify a bivariate table with missing on both axes writes both.
'@TestMethod("CrossTableFormula")
Public Sub TestBivariateWithMissingWritesBothAxes()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestBivariateWithMissingWritesBothAxes"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows("all", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula

    Set writer = WriteTable(1, sh, tabId, secId)

    Assert.IsTrue IsFormula(NamedFormula(sh, "MISSING_ROW_VALUES_" & tabId)), _
                  "The missing row should hold a formula"
    Assert.IsTrue IsFormula(NamedFormula(sh, "MISSING_COL_VALUES_" & tabId)), _
                  "The missing column should hold a formula"
    Assert.IsTrue IsFormula(NamedFormula(sh, "MISSING_MISSING_" & tabId)), _
                  "The cell where the two missing bands cross should hold a formula"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A bivariate table with both missing bands should report no error, " & _
                    "and it reported " & CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariateWithMissingWritesBothAxes", Err.Number, Err.Description
End Sub

'@sub-title Verify the four corner cells of a bivariate table.
'@TestMethod("CrossTableFormula")
Public Sub TestBivariateWritesTheFourCorners()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestBivariateWritesTheFourCorners"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows("all", SUM_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula

    Set writer = WriteTable(1, sh, tabId, secId)

    Assert.IsTrue IsFormula(NamedFormula(sh, "TOTAL_TOTAL_" & tabId)), _
                  "The grand total should hold a formula"
    Assert.IsTrue IsFormula(NamedFormula(sh, "MISSING_TOTAL_" & tabId)), _
                  "The missing row of the total column should hold a formula"
    Assert.IsTrue IsFormula(NamedFormula(sh, "TOTAL_MISSING_" & tabId)), _
                  "The total row of the missing column should hold a formula"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "The four corners should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariateWritesTheFourCorners", Err.Number, Err.Description
End Sub

'@sub-title Verify a count grand total sums the total column.
'@TestMethod("CrossTableFormula")
Public Sub TestBivariateCountGrandTotalSumsTheTotalColumn()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestBivariateCountGrandTotalSumsTheTotalColumn"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows("no", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim grandTotal As String

    Set writer = WriteTable(1, sh, tabId, secId)
    grandTotal = NamedFormula(sh, "TOTAL_TOTAL_" & tabId)

    Assert.IsTrue (InStr(1, grandTotal, "SUM(") > 0), _
                  "A count table sums its total column, and the cell holds [" & _
                  grandTotal & "]"
    Assert.IsTrue (InStr(1, grandTotal, "COUNTIFS") = 0), _
                  "The grand total of a count table needs no criteria of its own"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariateCountGrandTotalSumsTheTotalColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify every value column and the total column carry a percentage twin.
'@details
'The percentage twin of the total column used to be written by the value loop
'overrunning into it. The loop stops at the data columns now, so the total
'column's own block writes its twin.
'@TestMethod("CrossTableFormula")
Public Sub TestBivariatePercentageTwinsFollowEveryColumn()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestBivariatePercentageTwinsFollowEveryColumn"
    On Error GoTo TestFail

    BuildFixture TABLE_BIVARIATE, BivariateHeader(), _
                 BivariateRows("no", COUNT_CALL_FUNCTION, "column")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim totalValues As Range

    Set writer = WriteTable(1, sh, tabId, secId)

    Assert.IsTrue (InStr(1, NamedFormula(sh, "PERC_COL_1_" & tabId), "ISERR") > 0), _
                  "The first data column carries a percentage twin"
    Assert.IsTrue (InStr(1, NamedFormula(sh, "PERC_COL_2_" & tabId), "ISERR") > 0), _
                  "The second data column carries a percentage twin"

    Set totalValues = NamedRange(sh, "TOTAL_COL_VALUES_" & tabId)
    Assert.IsTrue (Not totalValues Is Nothing), "The total column should be named"
    Assert.IsTrue (InStr(1, CStr(totalValues.Cells(1, 2).Formula), "ISERR") > 0), _
                  "The total column carries a percentage twin of its own"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A bivariate percentage table should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBivariatePercentageTwinsFollowEveryColumn", Err.Number, Err.Description
End Sub

'@section Time series
'===============================================================================

'@sub-title Verify the period scaffolding of a new temporal section.
'@TestMethod("CrossTableFormula")
Public Sub TestTimeSeriesWritesThePeriodScaffolding()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestTimeSeriesWritesThePeriodScaffolding"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 TimeSeriesRows("no", COUNT_CALL_FUNCTION, "yes")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim periodFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    periodFormula = NamedFormula(sh, "ROW_CATEGORIES_" & tabId)

    Assert.IsTrue (InStr(1, periodFormula, "FormatDateFromLastDay") > 0), _
                  "The period label reads the worksheet function, and it holds [" & _
                  periodFormula & "]"
    Assert.IsTrue (InStr(1, periodFormula, "TIME_UNIT_" & tabId) > 0), _
                  "The period label reads the time unit control"
    Assert.IsTrue (InStr(1, NamedFormula(sh, "END_TIME_PERIOD_" & tabId), "FindLastDay") > 0), _
                  "The end of each period reads FindLastDay"
    Assert.IsTrue IsFormula(NamedFormula(sh, "START_TIME_PERIOD_" & tabId)), _
                  "The start of each period is one day after the previous end"
    Assert.IsTrue (InStr(1, NamedFormula(sh, "START_DATE_" & tabId), "MAX(") > 0), _
                  "The start date is the later of the typed date and the validated minimum"
    Assert.IsTrue (InStr(1, NamedFormula(sh, "END_DATE_" & tabId), "MIN(") > 0), _
                  "The end date is the earlier of the typed date and the validated maximum"
    Assert.IsTrue (InStr(1, NamedFormula(sh, "INFO_ANA_PERIOD_" & tabId), "PLAGE_VALUE") > 0), _
                  "The period information cell reads PLAGE_VALUE"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A temporal table should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesWritesThePeriodScaffolding", Err.Number, Err.Description
End Sub

'@sub-title Verify the date validation formulas read the cells the reader types into.
'@TestMethod("CrossTableFormula")
Public Sub TestTimeSeriesValidationReadsTheUserDates()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestTimeSeriesValidationReadsTheUserDates"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 TimeSeriesRows("no", COUNT_CALL_FUNCTION, "yes")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim maxFormula As String
    Dim minFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    maxFormula = NamedFormula(sh, "VALIDATION_MAX_DATE_" & tabId)
    minFormula = NamedFormula(sh, "VALIDATION_MIN_DATE_" & tabId)

    Assert.IsTrue (InStr(1, maxFormula, "USER_START_DATE_" & tabId) > 0), _
                  "The maximum validation reads the first typed date, and it holds [" & _
                  maxFormula & "]"
    Assert.IsTrue (InStr(1, maxFormula, "USER_END_DATE_" & tabId) > 0), _
                  "The maximum validation reads the last typed date"
    Assert.IsTrue (InStr(1, minFormula, "ValidMin") > 0), _
                  "The minimum validation reads ValidMin, and it holds [" & minFormula & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesValidationReadsTheUserDates", Err.Number, Err.Description
End Sub

'@sub-title Verify the time unit cell offers the shared choice list.
'@details
'The three dropdowns of this class used to be written inside blanket error
'handlers, so a table whose dropdown was never created looked exactly like one
'whose dropdown works.
'@TestMethod("CrossTableFormula")
Public Sub TestTimeSeriesCreatesTheTimeUnitDropdown()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestTimeSeriesCreatesTheTimeUnitDropdown"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 TimeSeriesRows("no", COUNT_CALL_FUNCTION, "yes")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim unitRng As Range
    Dim validationFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    Set unitRng = NamedRange(sh, "TIME_UNIT_" & tabId)

    Assert.IsTrue (Not unitRng Is Nothing), "The time unit cell should be named"

    validationFormula = TimeUnitValidationOf(sh, tabId)

    Assert.AreEqual "=TIME_UNIT_LIST", validationFormula, _
                    "A time series table reads the shared time unit list"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "Creating the dropdown should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesCreatesTheTimeUnitDropdown", Err.Number, Err.Description
End Sub

'@sub-title Verify every temporal value cell is bounded by its period.
'@TestMethod("CrossTableFormula")
Public Sub TestTimeSeriesBoundsEveryValueCellByThePeriod()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestTimeSeriesBoundsEveryValueCellByThePeriod"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 TimeSeriesRows("no", COUNT_CALL_FUNCTION, "yes")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim valueFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    valueFormula = NamedFormula(sh, "VALUES_COL_1_" & tabId)

    Assert.IsTrue IsFormula(valueFormula), _
                  "The first value cell should hold a formula, and it holds [" & _
                  valueFormula & "]"
    Assert.IsTrue (InStr(1, valueFormula, "IF(") = 2), _
                  "A temporal value cell blanks itself outside the selected period"
    Assert.AreEqual CLng(2), OccurrenceCount(valueFormula, "f" & linelistTable & "[" & _
                                        DATE_VARIABLE & "]"), _
                    "The two date bounds each name the time variable"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesBoundsEveryValueCellByThePeriod", Err.Number, Err.Description
End Sub

'@sub-title Verify the temporal missing row counts the records with no date.
'@details
'The missing row of a temporal table counts the records whose time variable is
'empty, so it carries no date bound. A record with no date sits in no period,
'and bounding that count would make it answer zero for every table.
'@TestMethod("CrossTableFormula")
Public Sub TestTimeSeriesMissingRowCountsTheRecordsWithNoDate()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestTimeSeriesMissingRowCountsTheRecordsWithNoDate"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 TimeSeriesRows("yes", COUNT_CALL_FUNCTION, "yes")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim missingFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    missingFormula = NamedFormula(sh, "MISSING_ROW_VALUES_" & tabId)

    Assert.IsTrue IsFormula(missingFormula), _
                  "The missing row should hold a formula, and it holds [" & _
                  missingFormula & "]"
    Assert.IsTrue (InStr(1, missingFormula, ">=") = 0), _
                  "The missing row carries no lower date bound"
    Assert.IsTrue IsFormula(NamedFormula(sh, "MISSING_COL_VALUES_" & tabId)), _
                  "The missing column of a temporal table is bounded and written"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesMissingRowCountsTheRecordsWithNoDate", Err.Number, Err.Description
End Sub

'@sub-title Verify the temporal grand total is written.
'@TestMethod("CrossTableFormula")
Public Sub TestTimeSeriesWritesTheGrandTotal()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestTimeSeriesWritesTheGrandTotal"
    On Error GoTo TestFail

    BuildFixture TABLE_TIMESERIES, TimeSeriesHeader(), _
                 TimeSeriesRows("no", SUM_FUNCTION, "yes")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula

    Set writer = WriteTable(1, sh, tabId, secId)

    Assert.IsTrue IsFormula(NamedFormula(sh, "TOTAL_COL_VALUES_" & tabId)), _
                  "The total column of a temporal table should hold a formula"
    Assert.IsTrue IsFormula(NamedFormula(sh, "TOTAL_TOTAL_" & tabId)), _
                  "The grand total of a temporal table should hold a formula"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A temporal total should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSeriesWritesTheGrandTotal", Err.Number, Err.Description
End Sub

'@section Spatio-temporal
'===============================================================================

'@sub-title Verify the column headers read the administrative input cells.
'@TestMethod("CrossTableFormula")
Public Sub TestSpatioTemporalLabelsReadTheGeoInputs()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestSpatioTemporalLabelsReadTheGeoInputs"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRows(GEO_VARIABLE, "2", COUNT_CALL_FUNCTION)

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim firstLabel As String
    Dim valueFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    firstLabel = NamedFormula(sh, "LABEL_COL_1_" & tabId)
    valueFormula = NamedFormula(sh, "VALUES_COL_1_" & tabId)

    Assert.IsTrue (InStr(1, firstLabel, "INPUTSPTGEO_1_" & secId) > 0), _
                  "The first column header reads the first administrative input cell, " & _
                  "and it holds [" & firstLabel & "]"
    Assert.IsTrue (InStr(1, NamedFormula(sh, "LABEL_COL_2_" & tabId), _
                         "INPUTSPTGEO_2_" & secId) > 0), _
                  "The second column header reads the second input cell"
    Assert.IsTrue (InStr(1, valueFormula, GEO_CONCAT_VARIABLE) > 0), _
                  "An administrative table is summarised over the concatenated column, " & _
                  "and the value cell holds [" & valueFormula & "]"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A spatio-temporal table should report no error, and it reported " & _
                    CheckMessages(writer)

    ' The spatio-temporal tables sit on a worksheet of their own and carry a time
    ' unit list of their own. The table writer built that name and this class
    ' asked for the plain one, so the dropdown was bound to a list the sheet does
    ' not carry and the failure was swallowed.
    Assert.AreEqual "=SPTIME_UNIT_LIST", TimeUnitValidationOf(sh, tabId), _
                    "A spatio-temporal table reads the spatio-temporal time unit list"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalLabelsReadTheGeoInputs", Err.Number, Err.Description
End Sub

'@sub-title Verify the label loop writes no header past the last data column.
'@details
'The label formulas are numbered by the column counter, and the counter runs to
'the number of data columns. A spatio-temporal table carries no total column, so
'the loop bound was already right here; this pins it, because the same counter
'now drives the value loops of the two scopes that do carry one.
'@TestMethod("CrossTableFormula")
Public Sub TestSpatioTemporalWritesNoHeaderPastTheLastColumn()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestSpatioTemporalWritesNoHeaderPastTheLastColumn"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRows(GEO_VARIABLE, "2", COUNT_CALL_FUNCTION)

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim lastLabel As Range
    Dim beyondFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    Set lastLabel = NamedRange(sh, "LABEL_COL_2_" & tabId)

    Assert.IsTrue (Not lastLabel Is Nothing), "The second column label should be named"

    beyondFormula = CStr(lastLabel.Cells(1, 2).Formula)

    Assert.IsTrue (InStr(1, beyondFormula, "INPUTSPTGEO_3_") = 0), _
                  "No input formula should reach the column after the last one, and " & _
                  "that cell holds [" & beyondFormula & "]"
    Assert.IsTrue (InStr(1, beyondFormula, "INPUTSPTHF_3_") = 0), _
                  "The same holds for the facility tag"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "No missing input cell should be reported, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalWritesNoHeaderPastTheLastColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify a facility table reads the facility input cells.
'@TestMethod("CrossTableFormula")
Public Sub TestSpatioTemporalFacilityLabelsUseTheFacilityTag()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestSpatioTemporalFacilityLabelsUseTheFacilityTag"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIOTEMPORAL, SpatioTemporalHeader(), _
                 SpatioTemporalRows(HF_VARIABLE, "2", COUNT_CALL_FUNCTION)

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim firstLabel As String
    Dim valueFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    firstLabel = NamedFormula(sh, "LABEL_COL_1_" & tabId)
    valueFormula = NamedFormula(sh, "VALUES_COL_1_" & tabId)

    Assert.IsTrue (InStr(1, firstLabel, "INPUTSPTHF_1_" & secId) > 0), _
                  "A facility table reads the facility input cells, and the header holds [" & _
                  firstLabel & "]"
    Assert.IsTrue (InStr(1, valueFormula, HF_PREFIXED_VARIABLE) > 0), _
                  "A facility table is summarised over the prefixed facility column, " & _
                  "and the value cell holds [" & valueFormula & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatioTemporalFacilityLabelsUseTheFacilityTag", Err.Number, Err.Description
End Sub

'@section Spatial
'===============================================================================

'@sub-title Verify a geographic spatial table writes its three lookups.
'@TestMethod("CrossTableFormula")
Public Sub TestSpatialGeoWritesItsLookups()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestSpatialGeoWritesItsLookups"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRows(GEO_VARIABLE, "3", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim categoryRng As Range
    Dim adminFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    Set categoryRng = NamedRange(sh, "ROW_CATEGORIES_" & tabId)

    Assert.IsTrue (Not categoryRng Is Nothing), "The row categories should be named"

    adminFormula = CStr(categoryRng.Cells(1, 1).Formula)

    Assert.IsTrue (InStr(1, adminFormula, "FindTopAdmin") > 0), _
                  "The category cells read FindTopAdmin, and the first holds [" & _
                  adminFormula & "]"
    Assert.IsTrue (InStr(1, adminFormula, GEO_CONCAT_VARIABLE) > 0), _
                  "The lookup names the concatenated administrative column"
    Assert.IsTrue (InStr(1, CStr(categoryRng.Cells(1, 0).Formula), "FindTopPop") > 0), _
                  "The column left of the categories reads FindTopPop"
    Assert.AreEqual CLng(1), CLng(categoryRng.Cells(1, -1).Value), _
                    "The order column starts at one"
    Assert.IsTrue IsFormula(CStr(categoryRng.Cells(2, -1).Formula)), _
                  "The order column counts up from the cell above"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A geographic spatial table should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialGeoWritesItsLookups", Err.Number, Err.Description
End Sub

'@sub-title Verify a geographic spatial table offers its two dropdowns.
'@TestMethod("CrossTableFormula")
Public Sub TestSpatialCreatesItsTwoDropdowns()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestSpatialCreatesItsTwoDropdowns"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRows(GEO_VARIABLE, "3", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim admFormula As String
    Dim popFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)

    On Error Resume Next
    admFormula = NamedRange(sh, "ADM_DROPDOWN_" & tabId).Validation.Formula1
    popFormula = NamedRange(sh, "DEVIDEPOP_" & tabId).Validation.Formula1
    On Error GoTo 0

    Assert.AreEqual "=ADM_UNIT_LIST", admFormula, _
                    "The administrative dropdown should read the shared unit list"
    Assert.AreEqual "=POPULATION_FACTOR_LIST", popFormula, _
                    "The divisor dropdown should read the shared factor list"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "Creating both dropdowns should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialCreatesItsTwoDropdowns", Err.Number, Err.Description
End Sub

'@sub-title Verify a facility spatial table writes the facility lookup.
'@TestMethod("CrossTableFormula")
Public Sub TestSpatialFacilityWritesTheFacilityLookup()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestSpatialFacilityWritesTheFacilityLookup"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRows(HF_VARIABLE, "3", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim facilityFormula As String

    Set writer = WriteTable(1, sh, tabId, secId)
    facilityFormula = NamedFormula(sh, "ROW_CATEGORIES_" & tabId)

    Assert.IsTrue (InStr(1, facilityFormula, "FindTopHF") > 0), _
                  "The category cells of a facility table read FindTopHF, and the " & _
                  "first holds [" & facilityFormula & "]"
    Assert.IsTrue (InStr(1, facilityFormula, HF_PREFIXED_VARIABLE) > 0), _
                  "The lookup names the prefixed facility column"
    Assert.IsTrue (NamedRange(sh, "ADM_DROPDOWN_" & tabId) Is Nothing), _
                  "A facility table builds no administrative dropdown"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A facility spatial table should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialFacilityWritesTheFacilityLookup", Err.Number, Err.Description
End Sub

'@sub-title Verify a spatial table showing one unit still writes its lookup.
'@details
'A geo count of one gives a row categories band of one cell, so the order
'column has no second row to count from and the value column has nothing to
'copy into. Both writes are skipped and the two formulas that belong there have
'to be there.
'@TestMethod("CrossTableFormula")
Public Sub TestSpatialWithOneUnitWritesItsLookup()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestSpatialWithOneUnitWritesItsLookup"
    On Error GoTo TestFail

    BuildFixture TABLE_SPATIAL, SpatialHeader(), _
                 SpatialRows(GEO_VARIABLE, "1", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim categoryRng As Range

    Set writer = WriteTable(1, sh, tabId, secId)
    Set categoryRng = NamedRange(sh, "ROW_CATEGORIES_" & tabId)

    Assert.AreEqual CLng(1), categoryRng.Rows.Count, "A geo count of one gives one row"
    Assert.IsTrue (InStr(1, CStr(categoryRng.Cells(1, 1).Formula), "FindTopAdmin") > 0), _
                  "The single category cell still reads FindTopAdmin"
    Assert.IsTrue IsFormula(NamedFormula(sh, "VALUES_COL_1_" & tabId)), _
                  "The single value cell still holds its formula"
    Assert.AreEqual CLng(0), ErrorCheckCount(writer), _
                    "A one-unit spatial table should report no error, and it reported " & _
                    CheckMessages(writer)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpatialWithOneUnitWritesItsLookup", Err.Number, Err.Description
End Sub

'@section Checkings
'===============================================================================

'@sub-title Verify a clean table reports nothing at all.
'@details
'HasCheckings is what AnalysisOutput tests before harvesting, so a table that
'went well has to answer False and hand back nothing.
'@TestMethod("CrossTableFormula")
Public Sub TestACleanTableReportsNothing()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestACleanTableReportsNothing"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "yes", COUNT_CALL_FUNCTION, "yes")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula

    Set writer = WriteTable(1, sh, tabId, secId)

    Assert.IsFalse writer.HasCheckings, _
                   "A table that went well should report nothing, and it reported " & _
                   CheckMessages(writer)
    Assert.IsTrue (writer.CheckingValues Is Nothing), _
                  "With nothing to report the entries should be Nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestACleanTableReportsNothing", Err.Number, Err.Description
End Sub

'@sub-title Verify a refused formula is reported and leaves the cell empty.
'@details
'The class writes the formula and lets Excel judge it. A formula Excel refuses
'has to leave the cell empty and file one entry naming the cell, which is the
'half of the fault that made the other half invisible: the report showed the
'analysis phase clean while the sheet carried a broken cell.
'
'The refusal is arranged by deleting the filtered linelist table the formulas
'reference, which is what a mis-generated linelist looks like.
'@TestMethod("CrossTableFormula")
Public Sub TestARefusedFormulaIsReportedAndTheCellIsEmpty()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestARefusedFormulaIsReportedAndTheCellIsEmpty"
    On Error GoTo TestFail

    BuildFixture TABLE_UNIVARIATE, UnivariateHeader(), _
                 UnivariateRows(ROW_CHOICE_VARIABLE, "no", COUNT_CALL_FUNCTION, "no")

    Dim sh As Worksheet
    Dim tabId As String
    Dim secId As String
    Dim writer As CrossTableFormula
    Dim valueText As String
    Dim errCount As Long
    Dim messages As String

    ReleaseTableName "f" & linelistTable

    Set writer = WriteTable(1, sh, tabId, secId)
    valueText = CStr(NamedRange(sh, "VALUES_COL_1_" & tabId).Cells(1, 1).Formula)
    errCount = ErrorCheckCount(writer)
    messages = CheckMessages(writer)

    BuildLinelistTables

    Assert.IsTrue (errCount > 0), _
                  "A formula Excel refuses should be reported, and the writer reported " & _
                  messages
    Assert.AreEqual vbNullString, valueText, _
                    "A refused formula should leave the cell empty, and it holds [" & _
                    valueText & "]"
    Assert.IsTrue (InStr(1, messages, "refused") > 0), _
                  "The message should say Excel refused the formula"
    Assert.IsFalse MarkerFound(sh), _
                   "A refused formula should write no marker into the table"

    Exit Sub
TestFail:
    BuildLinelistTables
    CustomTestLogFailure Assert, "TestARefusedFormulaIsReportedAndTheCellIsEmpty", Err.Number, Err.Description
End Sub
