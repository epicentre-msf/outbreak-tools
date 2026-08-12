Attribute VB_Name = "TestSpatialTables"
Attribute VB_Description = "Tests for SpatialTables class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for SpatialTables class")

'@description
'Drives SpatialTables, which writes the spatial ListObjects the generated
'linelist refreshes. The fixture is a real spatial specification row inside a
'Tab_Spatial_Analysis ListObject, a built CrossTable on an output worksheet, a
'linelist for the formulas to point at, and an empty "__spatial_tables" sheet.
'
'WHAT THE SUITE PINS
'-------------------------------------------------------------------------------
'The factory guards, the registry, the shape of a geographic table set and of a
'facility one, what is reported when a formula cannot be built, and the two
'contract tests: every name LLSpatial looks for is produced, and every row of the
'registry resolves to a real ListObject.
'@depends SpatialTables, CrossTable, CrossTableFormula, TableSpecs, Formulas,
'FormulaData, LLdictionary, LLVariables, TranslationObject, LinelistDataStub,
'CustomTest

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "SpTablesFixture"
Private Const OUTPUT_SHEET As String = "SpTablesOutput"
Private Const DICT_SHEET As String = "SpTablesDict"
Private Const TRANS_SHEET As String = "SpTablesTrans"
Private Const TRANS_TABLE As String = "T_SpTablesTranslation"
Private Const TOKENS_SHEET As String = "SpTablesTokens"
Private Const LINELIST_SHEET As String = "SpTablesLinelist"
Private Const SPATIAL_SHEET As String = "__spatial_tables"

' The specification header sits below the first rows the way the setup workbook
' lays it out, and the analysis scope is read from the ListObject name.
Private Const HEADER_ROW As Long = 5
Private Const TABLE_SPATIAL As String = "Tab_Spatial_Analysis"

' The reset of each worksheet is bounded to a block bigger than anything the
' suite writes. Clearing a whole worksheet costs seconds a test.
Private Const OUTPUT_BLOCK As String = "A1:AN200"
Private Const FIXTURE_BLOCK As String = "A1:T30"
Private Const SPATIAL_BLOCK As String = "A1:BZ40"

' The variables of the fixture linelist.
Private Const ROW_CHOICE_VARIABLE As String = "choi_v1"
Private Const COL_CHOICE_VARIABLE As String = "choi_ord_v1"
Private Const NUMBER_VARIABLE As String = "int_v1"
Private Const GEO_VARIABLE As String = "zone"
Private Const GEO_PREFIXED_VARIABLE As String = "adm1_zone"
Private Const HF_VARIABLE As String = "center"
Private Const HF_PREFIXED_VARIABLE As String = "hf_center"

' The sheet name and type every appended dictionary row carries.
Private Const FIXTURE_SHEET_NAME As String = "vlist1D-sheet1"
Private Const FIXTURE_SHEET_TYPE As String = "vlist1D"

Private Const COUNT_CALL_FUNCTION As String = "N()"

' A summary function the parser cannot read, which is how the suite reaches the
' paths that leave a value column empty.
Private Const BROKEN_FUNCTION As String = "NOSUCHFUNCTION(int_v1)"

Private Const REGISTRY_TABLE As String = "listofgeovars"

' Where Prepare puts the registry, and where the first spatial table starts.
Private Const REGISTRY_COL As Long = 3
Private Const FIRST_TABLE_COL As Long = 7

' An administrative table is four columns wide and the next one starts two
' columns past the last one it used.
Private Const SECOND_TABLE_COL As Long = 12

Private Assert As CustomTest
Private dict As LLdictionary
Private trans As TranslationObject
Private lData As LinelistDataStub
Private fData As FormulaData
Private linelistTable As String

'@section Fixture rows
'===============================================================================

'@sub-title Header of Tab_Spatial_Analysis, 12 columns.
Private Function SpatialHeader() As Variant
    SpatialHeader = Array( _
        "Section", "Table title", "Geo/HF variable (row)", "N geo max", _
        "Group by variable (column)", "Add missing data", "Summary function", _
        "Summary label", "Format", "Add percentage", "Add graph", _
        "Flip coordinates")
End Function

'@sub-title One spatial specification row.
'@param sectionName String. The section the table belongs to.
'@param rowVar String. The unprefixed geographic or facility variable.
'@param summaryFunction String. The summary function of the value column.
Private Function SpatialRow(ByVal sectionName As String, _
                            ByVal rowVar As String, _
                            ByVal summaryFunction As String) As Variant
    SpatialRow = Array(sectionName, "A spatial table", rowVar, "3", COL_CHOICE_VARIABLE, _
                       "no", summaryFunction, "Cases", "integer", "no", "no", "no")
End Function

'@sub-title A fixture holding one spatial row.
'@param rowVar String. The unprefixed geographic or facility variable.
'@param summaryFunction String. The summary function of the value column.
Private Function SpatialRows(ByVal rowVar As String, _
                             ByVal summaryFunction As String) As Variant
    SpatialRows = Array(SpatialRow("S1", rowVar, summaryFunction))
End Function

'@sub-title A fixture holding two spatial rows over two different variables.
'@details
'The two rows sit in two sections. The output worksheet is reset between two
'builds, and a second table of the same section looks for the section markers the
'reset has taken away.
Private Function TwoSpatialRows() As Variant
    TwoSpatialRows = Array(SpatialRow("S1", GEO_VARIABLE, COUNT_CALL_FUNCTION), _
                           SpatialRow("S2", HF_VARIABLE, COUNT_CALL_FUNCTION))
End Function

'@section Fixture helpers
'===============================================================================

'@sub-title Free a ListObject name wherever it is taken in the workbook.
'@details
'A ListObject name is unique across the workbook and the six analysis names are
'the only ones TableSpecs answers a scope for, so another suite fixture sheet
'holding one blocks this one from taking it. Unlist turns the table back into an
'ordinary range and frees the name.
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

'@sub-title Build the specification ListObject with a header row and data rows.
'@param dataRows Variant. An array of row arrays, each as wide as the header.
Private Sub BuildFixture(ByVal dataRows As Variant)
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim tableRng As Range
    Dim headerRow As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim idx As Long

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=False, visibility:=xlSheetHidden)

    For idx = sh.ListObjects.Count To 1 Step -1
        sh.ListObjects(idx).Unlist
    Next idx

    sh.Range(FIXTURE_BLOCK).Clear
    ReleaseTableName TABLE_SPATIAL

    headerRow = SpatialHeader()
    colCount = UBound(headerRow) - LBound(headerRow) + 1
    rowCount = UBound(dataRows) - LBound(dataRows) + 1

    WriteMatrix sh.Cells(HEADER_ROW, 1), RowsToMatrix(Array(headerRow))
    WriteMatrix sh.Cells(HEADER_ROW + 1, 1), RowsToMatrix(dataRows)

    Set tableRng = sh.Range(sh.Cells(HEADER_ROW, 1), _
                            sh.Cells(HEADER_ROW + rowCount, colCount))
    Set lo = sh.ListObjects.Add(xlSrcRange, tableRng, , xlYes)
    lo.Name = TABLE_SPATIAL
End Sub

'@sub-title Create a TableSpecs from a fixture data row index.
'@param dataRowIndex Long. The 1-based row to read.
Private Function CreateSpecs(ByVal dataRowIndex As Long) As TableSpecs
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    Set CreateSpecs = TableSpecs.Create(sh.ListObjects(1).HeaderRowRange, _
                                        sh.ListObjects(1).ListRows(dataRowIndex).Range, _
                                        dict)
End Function

'@sub-title Delete every workbook-scoped name that points at one worksheet.
'@details
'CrossTable creates its names with Cell.Name, which are workbook-scoped and
'outlive a Clear, and it asks whether a name exists to decide whether it has
'already built a piece of structure. Excel writes RefersTo as ='Sheet name'!$C$10
'when the sheet name needs quoting and as =Sheetname!$C$10 when it does not, and
'both spellings are matched here.
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

'@sub-title Return the analysis output worksheet with nothing on it.
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

'@sub-title Return the spatial worksheet with no tables, no names and no values.
'@details
'Every test starts from an empty spatial sheet, because the class asks whether
'the tables of a variable are already there and lays the next one out from the
'last used column of row 1.
Private Function SpatialSheet() As Worksheet
    Dim sh As Worksheet
    Dim idx As Long

    Set sh = EnsureWorksheet(SPATIAL_SHEET, clearSheet:=False, visibility:=xlSheetHidden)

    For idx = sh.ListObjects.Count To 1 Step -1
        sh.ListObjects(idx).Unlist
    Next idx

    RemoveSheetNames sh
    sh.Range(SPATIAL_BLOCK).Clear
    Set SpatialSheet = sh
End Function

'@sub-title Build one cross-table from a fixture row.
'@param dataRowIndex Long. The 1-based fixture data row.
'@return CrossTable. The built cross-table on a clean output worksheet.
Private Function BuiltTable(ByVal dataRowIndex As Long) As CrossTable
    Dim tabl As CrossTable

    Set tabl = CrossTable.Create(CreateSpecs(dataRowIndex), OutputSheet(), lData)
    tabl.Build
    Set BuiltTable = tabl
End Function

'@sub-title Build the summary function object of one fixture row.
'@param specs TableSpecs. The specification row.
Private Function FormulaOf(ByVal specs As TableSpecs) As Formulas
    Set FormulaOf = Formulas.Create(specs.Dictionary, fData, _
                                    specs.Value("summary function"))
End Function

'@sub-title Write the spatial tables of one fixture row.
'@param dataRowIndex Long. The 1-based fixture data row.
'@param varName String. ByRef. Filled with the registry name of the variable.
'@return SpatialTables. The builder, with Add already run.
Private Function WriteSpatial(ByVal dataRowIndex As Long, _
                              ByRef varName As String) As SpatialTables
    Dim tabl As CrossTable
    Dim spTab As SpatialTables

    Set tabl = BuiltTable(dataRowIndex)
    varName = tabl.Specifications.Value("row") & "_" & tabl.Specifications.TableId()

    Set spTab = SpatialTables.Create(tabl)
    spTab.Add FormulaOf(tabl.Specifications)
    Set WriteSpatial = spTab
End Function

'@sub-title Whether a ListObject of that name sits on the spatial sheet.
'@param loName String. The ListObject name to look for.
Private Function SpatialTableExists(ByVal loName As String) As Boolean
    Dim lo As ListObject

    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(SPATIAL_SHEET).ListObjects(loName)
    On Error GoTo 0

    SpatialTableExists = Not (lo Is Nothing)
End Function

'@sub-title The ListObject of that name on the spatial sheet, or Nothing.
'@param loName String. The ListObject name to look for.
Private Function SpatialTable(ByVal loName As String) As ListObject
    On Error Resume Next
    Set SpatialTable = ThisWorkbook.Worksheets(SPATIAL_SHEET).ListObjects(loName)
    On Error GoTo 0
End Function

'@sub-title Whether a ListObject carries a column of that name.
'@param lo ListObject. The table to look in.
'@param columnName String. The column name to look for.
Private Function HasColumn(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    Dim col As ListColumn

    If lo Is Nothing Then Exit Function

    On Error Resume Next
    Set col = lo.ListColumns(columnName)
    On Error GoTo 0

    HasColumn = Not (col Is Nothing)
End Function

'@sub-title The values of the registry column, joined for a message.
Private Function RegistryEntries() As String
    Dim lo As ListObject
    Dim idx As Long
    Dim joined As String

    Set lo = SpatialTable(REGISTRY_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function

    For idx = 1 To lo.ListRows.Count
        joined = joined & "[" & CStr(lo.ListRows(idx).Range.Cells(1, 1).Value) & "]"
    Next idx

    RegistryEntries = joined
End Function

'@sub-title How many registry rows hold a variable name.
Private Function RegistryCount() As Long
    Dim lo As ListObject
    Dim idx As Long
    Dim counter As Long

    Set lo = SpatialTable(REGISTRY_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function

    For idx = 1 To lo.ListRows.Count
        If CStr(lo.ListRows(idx).Range.Cells(1, 1).Value) <> vbNullString Then
            counter = counter + 1
        End If
    Next idx

    RegistryCount = counter
End Function

'@sub-title The messages a builder filed, joined for a failure message.
'@param spTab SpatialTables. The builder to read.
Private Function CheckMessages(ByVal spTab As SpatialTables) As String
    Dim checks As Checking
    Dim keyList As BetterArray
    Dim idx As Long
    Dim joined As String

    If Not spTab.HasCheckings Then Exit Function

    Set checks = spTab.CheckingValues
    Set keyList = checks.ListOfKeys

    For idx = keyList.LowerBound To keyList.UpperBound
        joined = joined & "[" & checks.ValueOf(CStr(keyList.Item(idx)), checkingLabel) & "]"
    Next idx

    CheckMessages = joined
End Function

'@sub-title The messages a formula writer filed, joined for a failure message.
'@param writer CrossTableFormula. The writer to read.
Private Function WriterMessages(ByVal writer As CrossTableFormula) As String
    Dim checks As Checking
    Dim keyList As BetterArray
    Dim idx As Long
    Dim joined As String

    If Not writer.HasCheckings Then Exit Function

    Set checks = writer.CheckingValues
    Set keyList = checks.ListOfKeys

    For idx = keyList.LowerBound To keyList.UpperBound
        joined = joined & "[" & CStr(keyList.Item(idx)) & "] "
    Next idx

    WriterMessages = joined
End Function

'@sub-title Whether a cell holds a formula.
'@param formulaText String. The text read off the cell.
Private Function IsFormula(ByVal formulaText As String) As Boolean
    IsFormula = (Left(formulaText, 1) = "=")
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Build the dictionary, the translator, the tokens and the linelist.
'@details
'The dictionary fixture is prepared once and Prepare is called on it, because
'Prepare writes the "table name" column every analysis formula reads. This
'routine is Public because the harness calls it by name through Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    Dim sh As Worksheet
    Dim appendRow As Long
    Dim counter As Long

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSpatialTables"

    PrepareDictionaryFixture DICT_SHEET

    ' A geographic spatial variable is summarised over one concatenated column
    ' per administrative level, so all four are in the dictionary and in the
    ' linelist. hf_center is the facility counterpart.
    Set sh = ThisWorkbook.Worksheets(DICT_SHEET)
    appendRow = 1 + DictionaryFixtureRowCount() + 1
    AppendGeoRow sh, appendRow, GEO_PREFIXED_VARIABLE, "Zone", "Zones"
    For counter = 1 To 4
        AppendGeoRow sh, appendRow + counter, _
                     "concat_adm" & counter & "_" & GEO_VARIABLE, _
                     "Zone code " & counter, "Zones"
    Next
    AppendGeoRow sh, appendRow + 5, HF_PREFIXED_VARIABLE, "Health centre", vbNullString

    Set dict = LLdictionary.Create(sh, 1, 1)
    dict.Prepare

    Set trans = BuildTranslator()

    Set lData = New LinelistDataStub
    lData.SetTransObject trans
    lData.SetCategories ROW_CHOICE_VARIABLE, BetterArrayFromList("A", "B", "C")
    lData.SetCategories COL_CHOICE_VARIABLE, BetterArrayFromList("X", "Y")

    ' FormulaData resolves its two lookup tables by fixed names, so another
    ' suite fixture sheet holding them blocks this one from taking them.
    ReleaseTableName "T_XlsFonctions"
    ReleaseTableName "T_ascii"
    Set fData = FormulaData.Create(PrepareFormulaFixtureSheet(TOKENS_SHEET))

    BuildLinelistTables

    EnsureWorksheet SPATIAL_SHEET, clearSheet:=True, visibility:=xlSheetHidden
End Sub

'@sub-title Write one dictionary row for an appended variable.
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
'An analysis formula is a structured reference into the linelist table of the
'variable and into the filtered copy of it, which carries the "f" prefix. Without
'both, Excel refuses every formula the class writes and the suite reads as a
'class fault.
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

    columnNames = Array(ROW_CHOICE_VARIABLE, COL_CHOICE_VARIABLE, NUMBER_VARIABLE, _
                        GEO_PREFIXED_VARIABLE, "concat_adm1_" & GEO_VARIABLE, _
                        "concat_adm2_" & GEO_VARIABLE, "concat_adm3_" & GEO_VARIABLE, _
                        "concat_adm4_" & GEO_VARIABLE, HF_PREFIXED_VARIABLE)

    dataRow = Array("A", "X", 5, "Z1", "Z1", "Z2", "Z3", "Z4", "C1")

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

'@section Factory and guards
'===============================================================================

'@sub-title Verify Create rejects a cross-table that was not given.
'@TestMethod("SpatialTables")
Public Sub TestCreateRejectsNothing()
    CustomTestSetTitles Assert, "SpatialTables", "TestCreateRejectsNothing"
    On Error GoTo TestFail

    Dim spTab As SpatialTables
    Dim errNumber As Long

    SpatialSheet

    On Error Resume Next
    Set spTab = SpatialTables.Create(Nothing)
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "Create with no cross-table should raise ObjectNotInitialized"
    Assert.IsTrue (spTab Is Nothing), "Nothing should come back from a rejected Create"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothing", Err.Number, Err.Description
End Sub

'@sub-title Verify Create names the worksheet it could not find.
'@details
'The guard used to test a variable that still held the analysis worksheet, and a
'failed Set leaves its target at what it held, so it could never fire. The
'missing worksheet then surfaced later as a bare subscript error attributed to
'this class with nothing pointing at a sheet.
'@TestMethod("SpatialTables")
Public Sub TestCreateRejectsMissingSpatialSheet()
    CustomTestSetTitles Assert, "SpatialTables", "TestCreateRejectsMissingSpatialSheet"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim spTab As SpatialTables
    Dim errNumber As Long
    Dim errText As String

    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    Set tabl = BuiltTable(1)
    DeleteWorksheet SPATIAL_SHEET

    On Error Resume Next
    Set spTab = SpatialTables.Create(tabl)
    errNumber = Err.Number
    errText = Err.Description
    On Error GoTo 0

    EnsureWorksheet SPATIAL_SHEET, clearSheet:=True, visibility:=xlSheetHidden

    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "A missing spatial worksheet should raise ElementNotFound, and it " & _
                    "raised (" & errNumber & ") " & errText
    Assert.IsTrue (spTab Is Nothing), "Nothing should come back from a rejected Create"

    Exit Sub
TestFail:
    EnsureWorksheet SPATIAL_SHEET, clearSheet:=True, visibility:=xlSheetHidden
    CustomTestLogFailure Assert, "TestCreateRejectsMissingSpatialSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify Create succeeds when the spatial worksheet is there.
'@TestMethod("SpatialTables")
Public Sub TestCreateSucceedsWithSpatialSheet()
    CustomTestSetTitles Assert, "SpatialTables", "TestCreateSucceedsWithSpatialSheet"
    On Error GoTo TestFail

    Dim spTab As SpatialTables

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)

    Set spTab = SpatialTables.Create(BuiltTable(1))

    Assert.IsTrue (Not spTab Is Nothing), "A workbook carrying the spatial sheet builds"
    Assert.IsFalse spTab.HasCheckings, "A builder that wrote nothing reports nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateSucceedsWithSpatialSheet", Err.Number, Err.Description
End Sub

'@sub-title Verify the cross-table cannot be swapped after creation.
'@TestMethod("SpatialTables")
Public Sub TestTableIsSetAtCreationOnly()
    CustomTestSetTitles Assert, "SpatialTables", "TestTableIsSetAtCreationOnly"
    On Error GoTo TestFail

    Dim spTab As SpatialTables
    Dim errNumber As Long

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    Set spTab = SpatialTables.Create(BuiltTable(1))

    On Error Resume Next
    Set spTab.Table = BuiltTable(1)
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.SomethingWentWrong), errNumber, _
                    "Assigning the table after creation should raise"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTableIsSetAtCreationOnly", Err.Number, Err.Description
End Sub

'@section Preparation
'===============================================================================

'@sub-title Verify the first Add puts the registry on the sheet.
'@TestMethod("SpatialTables")
Public Sub TestAddCreatesTheRegistry()
    CustomTestSetTitles Assert, "SpatialTables", "TestAddCreatesTheRegistry"
    On Error GoTo TestFail

    Dim varName As String
    Dim lo As ListObject

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    Set lo = SpatialTable(REGISTRY_TABLE)

    Assert.IsTrue (Not lo Is Nothing), "The registry should be on the spatial sheet"
    Assert.AreEqual CLng(REGISTRY_COL), lo.Range.Column, _
                    "The registry sits in the column LLSpatial reads it from"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddCreatesTheRegistry", Err.Number, Err.Description
End Sub

'@sub-title Verify a second variable reuses the registry already there.
'@TestMethod("SpatialTables")
Public Sub TestTheRegistryIsBuiltOnce()
    CustomTestSetTitles Assert, "SpatialTables", "TestTheRegistryIsBuiltOnce"
    On Error GoTo TestFail

    Dim firstName As String
    Dim secondName As String

    SpatialSheet
    BuildFixture TwoSpatialRows()
    WriteSpatial 1, firstName
    WriteSpatial 2, secondName

    Assert.AreEqual CLng(2), RegistryCount(), _
                    "Two variables give two registry rows, and the registry holds " & _
                    RegistryEntries()
    Assert.IsTrue (firstName <> secondName), "The two variables have different names"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheRegistryIsBuiltOnce", Err.Number, Err.Description
End Sub

'@section Geographic tables
'===============================================================================

'@sub-title Verify a geographic variable gets one table per administrative level.
'@TestMethod("SpatialTables")
Public Sub TestGeoCreatesFourListObjects()
    CustomTestSetTitles Assert, "SpatialTables", "TestGeoCreatesFourListObjects"
    On Error GoTo TestFail

    Dim varName As String
    Dim counter As Long

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    For counter = 1 To 4
        Assert.IsTrue SpatialTableExists("spatial_adm" & counter & "_" & varName), _
                      "spatial_adm" & counter & "_" & varName & " should be built"
    Next

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoCreatesFourListObjects", Err.Number, Err.Description
End Sub

'@sub-title Verify a geographic table carries the four columns LLSpatial reads.
'@TestMethod("SpatialTables")
Public Sub TestGeoWritesTheFourColumnHeaders()
    CustomTestSetTitles Assert, "SpatialTables", "TestGeoWritesTheFourColumnHeaders"
    On Error GoTo TestFail

    Dim varName As String
    Dim lo As ListObject
    Dim headerRng As Range

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    Set lo = SpatialTable("spatial_adm1_" & varName)
    Set headerRng = lo.HeaderRowRange

    Assert.AreEqual CLng(4), lo.ListColumns.Count, "A spatial table is four columns wide"
    Assert.AreEqual "tabl_adm1_" & varName, CStr(headerRng.Cells(1, 1).Value), _
                    "The first column is the lookup key"
    Assert.AreEqual "formula_adm1", CStr(headerRng.Cells(1, 2).Value), _
                    "The second column is the value LLSpatial sorts on"
    Assert.AreEqual "population_adm1", CStr(headerRng.Cells(1, 3).Value), _
                    "The third column is the population"
    Assert.AreEqual "attack_rate_adm1", CStr(headerRng.Cells(1, 4).Value), _
                    "The fourth column is the attack rate"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoWritesTheFourColumnHeaders", Err.Number, Err.Description
End Sub

'@sub-title Verify the value column holds the formula of the summary function.
'@TestMethod("SpatialTables")
Public Sub TestGeoWritesTheValueFormula()
    CustomTestSetTitles Assert, "SpatialTables", "TestGeoWritesTheValueFormula"
    On Error GoTo TestFail

    Dim varName As String
    Dim spTab As SpatialTables
    Dim lo As ListObject
    Dim valueFormula As String

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    Set spTab = WriteSpatial(1, varName)

    Set lo = SpatialTable("spatial_adm1_" & varName)
    valueFormula = CStr(lo.ListRows(1).Range.Cells(1, 2).Formula)

    Assert.IsTrue IsFormula(valueFormula), _
                  "The value cell should hold a formula, and it holds [" & _
                  valueFormula & "]"
    Assert.IsTrue (InStr(1, valueFormula, "concat_adm1_" & GEO_VARIABLE) > 0), _
                  "The formula reads the concatenated column of the level, and it " & _
                  "holds [" & valueFormula & "]"
    Assert.IsFalse spTab.HasCheckings, _
                   "A clean geographic set reports nothing, and it reported " & _
                   CheckMessages(spTab)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoWritesTheValueFormula", Err.Number, Err.Description
End Sub

'@sub-title Verify the population and attack rate cells are written.
'@TestMethod("SpatialTables")
Public Sub TestGeoWritesThePopulationAndAttackRate()
    CustomTestSetTitles Assert, "SpatialTables", "TestGeoWritesThePopulationAndAttackRate"
    On Error GoTo TestFail

    Dim varName As String
    Dim lo As ListObject
    Dim populationFormula As String
    Dim attackFormula As String

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    Set lo = SpatialTable("spatial_adm2_" & varName)
    populationFormula = CStr(lo.ListRows(1).Range.Cells(1, 3).Formula)
    attackFormula = CStr(lo.ListRows(1).Range.Cells(1, 4).Formula)

    Assert.IsTrue (InStr(1, populationFormula, "T_ADM2") > 0), _
                  "The population reads the geobase table of its level, and it holds [" & _
                  populationFormula & "]"
    Assert.IsTrue (InStr(1, populationFormula, "adm2_concat") > 0), _
                  "The population matches on the concatenated column of the geobase"
    Assert.IsTrue IsFormula(attackFormula), _
                  "The attack rate divides the value by the population, and it holds [" & _
                  attackFormula & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoWritesThePopulationAndAttackRate", Err.Number, Err.Description
End Sub

'@sub-title Verify the four tables are laid out side by side with a gap.
'@TestMethod("SpatialTables")
Public Sub TestGeoLeavesAGapBetweenTables()
    CustomTestSetTitles Assert, "SpatialTables", "TestGeoLeavesAGapBetweenTables"
    On Error GoTo TestFail

    Dim varName As String

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    Assert.AreEqual CLng(FIRST_TABLE_COL), _
                    SpatialTable("spatial_adm1_" & varName).Range.Column, _
                    "The first table starts in the column the layout gives it"
    Assert.AreEqual CLng(SECOND_TABLE_COL), _
                    SpatialTable("spatial_adm2_" & varName).Range.Column, _
                    "The next table starts two columns past the last one used"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoLeavesAGapBetweenTables", Err.Number, Err.Description
End Sub

'@section Facility tables
'===============================================================================

'@sub-title Verify a facility variable gets one table.
'@TestMethod("SpatialTables")
Public Sub TestFacilityCreatesOneListObject()
    CustomTestSetTitles Assert, "SpatialTables", "TestFacilityCreatesOneListObject"
    On Error GoTo TestFail

    Dim varName As String

    SpatialSheet
    BuildFixture SpatialRows(HF_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    Assert.IsTrue SpatialTableExists("spatial_hf_" & varName), _
                  "spatial_hf_" & varName & " should be built"
    Assert.IsFalse SpatialTableExists("spatial_adm1_" & varName), _
                   "A facility variable builds no administrative table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFacilityCreatesOneListObject", Err.Number, Err.Description
End Sub

'@sub-title Verify a facility table is the lookup key and the value.
'@details
'A health facility has no population, so the population and the attack rate
'belong to an administrative table alone. The facility table used to carry all
'four columns with the last two always blank, and LLSpatial.Update resized every
'spatial table to four columns, which is what kept them.
'@TestMethod("SpatialTables")
Public Sub TestAFacilityTableIsTwoColumns()
    CustomTestSetTitles Assert, "SpatialTables", "TestAFacilityTableIsTwoColumns"
    On Error GoTo TestFail

    Dim varName As String
    Dim lo As ListObject
    Dim headerRng As Range

    SpatialSheet
    BuildFixture SpatialRows(HF_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    Set lo = SpatialTable("spatial_hf_" & varName)
    Set headerRng = lo.HeaderRowRange

    Assert.AreEqual CLng(2), lo.ListColumns.Count, _
                    "A facility table is the lookup key and the value"
    Assert.AreEqual "tabl_hf_" & varName, CStr(headerRng.Cells(1, 1).Value), _
                    "The first column is the lookup key"
    Assert.AreEqual "formula_hf", CStr(headerRng.Cells(1, 2).Value), _
                    "The second column is the value LLSpatial reads"
    Assert.IsTrue IsFormula(CStr(lo.ListRows(1).Range.Cells(1, 2).Formula)), _
                  "The value column of a facility table holds its formula"
    Assert.IsFalse HasColumn(lo, "population_hf"), _
                   "A facility table carries no population column"
    Assert.IsFalse HasColumn(lo, "attack_rate_hf"), _
                   "A facility table carries no attack rate column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFacilityTableIsTwoColumns", Err.Number, Err.Description
End Sub

'@sub-title Verify a facility variable already built is left alone.
'@details
'The guard used to ask for four administrative tables whatever the spatial type,
'so a facility variable answered "not built" every time: it was registered twice
'and the second pass raised 1004 on a ListObject name the first had taken.
'@TestMethod("SpatialTables")
Public Sub TestFacilityAddIsIdempotent()
    CustomTestSetTitles Assert, "SpatialTables", "TestFacilityAddIsIdempotent"
    On Error GoTo TestFail

    Dim varName As String
    Dim secondName As String
    Dim errNumber As Long

    SpatialSheet
    BuildFixture SpatialRows(HF_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    On Error Resume Next
    WriteSpatial 1, secondName
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(0), errNumber, _
                    "A second Add over the same facility variable should raise nothing"
    Assert.AreEqual CLng(1), RegistryCount(), _
                    "The variable should be registered once, and the registry holds " & _
                    RegistryEntries()

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFacilityAddIsIdempotent", Err.Number, Err.Description
End Sub

'@sub-title Verify a geographic variable already built is left alone.
'@TestMethod("SpatialTables")
Public Sub TestGeoAddIsIdempotent()
    CustomTestSetTitles Assert, "SpatialTables", "TestGeoAddIsIdempotent"
    On Error GoTo TestFail

    Dim varName As String
    Dim secondName As String
    Dim errNumber As Long

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    On Error Resume Next
    WriteSpatial 1, secondName
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(0), errNumber, _
                    "A second Add over the same geographic variable should raise nothing"
    Assert.AreEqual CLng(1), RegistryCount(), _
                    "The variable should be registered once, and the registry holds " & _
                    RegistryEntries()

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoAddIsIdempotent", Err.Number, Err.Description
End Sub

'@section The registry
'===============================================================================

'@sub-title Verify the variable is registered under the name the tables carry.
'@TestMethod("SpatialTables")
Public Sub TestAddRegistersTheVariable()
    CustomTestSetTitles Assert, "SpatialTables", "TestAddRegistersTheVariable"
    On Error GoTo TestFail

    Dim varName As String

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    Assert.AreEqual CLng(1), RegistryCount(), "One variable gives one registry row"
    Assert.IsTrue (InStr(1, RegistryEntries(), varName) > 0), _
                  "The registry should name " & varName & ", and it holds " & _
                  RegistryEntries()

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddRegistersTheVariable", Err.Number, Err.Description
End Sub

'@sub-title Verify a blank row in the registry costs no entry.
'@details
'The row a variable was written to used to be counted with CountA over the whole
'registry. CountA answers how many cells hold something, and the write needed a
'row number, so one blank row anywhere made the next variable overwrite an
'existing entry. LLSpatial
'walks that list, so a lost entry is a spatial table that never refreshes.
'@TestMethod("SpatialTables")
Public Sub TestABlankRegistryRowCostsNoEntry()
    CustomTestSetTitles Assert, "SpatialTables", "TestABlankRegistryRowCostsNoEntry"
    On Error GoTo TestFail

    Dim firstName As String
    Dim secondName As String
    Dim lo As ListObject

    Dim newRow As ListRow

    SpatialSheet
    BuildFixture TwoSpatialRows()
    WriteSpatial 1, firstName

    ' Three entries with the middle one lost. A count of the cells holding
    ' something answers three, which points at the row of the last entry.
    Set lo = SpatialTable(REGISTRY_TABLE)
    Set newRow = lo.ListRows.Add
    newRow.Range.Cells(1, 1).Value = "kept_one"
    Set newRow = lo.ListRows.Add
    newRow.Range.Cells(1, 1).Value = "kept_two"
    lo.ListRows(2).Range.Cells(1, 1).ClearContents

    WriteSpatial 2, secondName

    Assert.IsTrue (InStr(1, RegistryEntries(), "kept_two") > 0), _
                  "The last entry should survive the next variable, and the registry " & _
                  "holds " & RegistryEntries()
    Assert.IsTrue (InStr(1, RegistryEntries(), secondName) > 0), _
                  "The second variable should be registered too, and the registry " & _
                  "holds " & RegistryEntries()
    Assert.AreEqual CLng(3), RegistryCount(), _
                    "Three names should be readable, and the registry holds " & _
                    RegistryEntries()

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestABlankRegistryRowCostsNoEntry", Err.Number, Err.Description
End Sub

'@sub-title Verify a variable whose tables failed is left out of the registry.
'@details
'The variable used to be registered before the tables were built, so a failure
'in the middle left a name in the registry with nothing behind it and LLSpatial
'walked to a table that was not there. The failure is arranged by taking one of
'the four ListObject names first: a name already in use raises 1004.
'@TestMethod("SpatialTables")
Public Sub TestAFailedTableSetRegistersNothing()
    CustomTestSetTitles Assert, "SpatialTables", "TestAFailedTableSetRegistersNothing"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim spTab As SpatialTables
    Dim sh As Worksheet
    Dim varName As String
    Dim blockingName As String
    Dim errNumber As Long

    Set sh = SpatialSheet()
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)

    Set tabl = BuiltTable(1)
    varName = tabl.Specifications.Value("row") & "_" & tabl.Specifications.TableId()
    blockingName = "spatial_adm2_" & varName

    sh.Cells(1, 40).Value = "taken"
    sh.ListObjects.Add(xlSrcRange, sh.Range(sh.Cells(1, 40), sh.Cells(2, 40)), , xlYes).Name = blockingName

    Set spTab = SpatialTables.Create(tabl)

    On Error Resume Next
    spTab.Add FormulaOf(tabl.Specifications)
    errNumber = Err.Number
    On Error GoTo 0

    Assert.IsTrue (errNumber <> 0), _
                  "A ListObject name already taken should raise out of Add"
    Assert.AreEqual CLng(0), RegistryCount(), _
                    "A variable whose tables failed should not be registered, and the " & _
                    "registry holds " & RegistryEntries()

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFailedTableSetRegistersNothing", Err.Number, Err.Description
End Sub

'@section Checkings
'===============================================================================

'@sub-title Verify a value column that could not be built is reported.
'@details
'The value column used to be left empty with nothing said, and the attack rate
'two cells over divided by it, so the delivered table showed 0 or a division
'error and the generation report read clean.
'@TestMethod("SpatialTables")
Public Sub TestAnUnbuildableFormulaIsReported()
    CustomTestSetTitles Assert, "SpatialTables", "TestAnUnbuildableFormulaIsReported"
    On Error GoTo TestFail

    Dim varName As String
    Dim spTab As SpatialTables
    Dim lo As ListObject
    Dim messages As String

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, BROKEN_FUNCTION)
    Set spTab = WriteSpatial(1, varName)
    messages = CheckMessages(spTab)

    Set lo = SpatialTable("spatial_adm1_" & varName)

    Assert.IsTrue spTab.HasCheckings, _
                  "A summary function that cannot be read should be reported"
    Assert.IsTrue (InStr(1, messages, "No formula could be built") > 0), _
                  "Each empty value cell should be named, and the report holds " & messages
    Assert.AreEqual vbNullString, CStr(lo.ListRows(1).Range.Cells(1, 2).Formula), _
                    "The value cell of a formula that could not be built stays empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnUnbuildableFormulaIsReported", Err.Number, Err.Description
End Sub

'@sub-title Verify a clean table set reports nothing at all.
'@TestMethod("SpatialTables")
Public Sub TestACleanAddReportsNothing()
    CustomTestSetTitles Assert, "SpatialTables", "TestACleanAddReportsNothing"
    On Error GoTo TestFail

    Dim varName As String
    Dim spTab As SpatialTables

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    Set spTab = WriteSpatial(1, varName)

    Assert.IsFalse spTab.HasCheckings, _
                   "A table set that went well reports nothing, and it reported " & _
                   CheckMessages(spTab)
    Assert.IsTrue (spTab.CheckingValues Is Nothing), _
                  "With nothing to report the entries should be Nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestACleanAddReportsNothing", Err.Number, Err.Description
End Sub

'@sub-title Verify what the spatial builder reports reaches the formula writer.
'@details
'CrossTableFormula creates the spatial builder and is the only class that can
'carry its entries up to the generation report. Nothing harvested from it before,
'so everything filed here was filed into a report nobody read.
'@TestMethod("SpatialTables")
Public Sub TestTheSpatialReportReachesTheFormulaWriter()
    CustomTestSetTitles Assert, "SpatialTables", "TestTheSpatialReportReachesTheFormulaWriter"
    On Error GoTo TestFail

    Dim tabl As CrossTable
    Dim writer As CrossTableFormula
    Dim keys As String

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)

    Set tabl = BuiltTable(1)

    ' The formulas point at the filtered linelist table, so taking it away is
    ' what a mis-generated linelist looks like and Excel refuses every formula.
    ReleaseTableName "f" & linelistTable

    Set writer = CrossTableFormula.Create(tabl, fData)
    writer.AddFormulas
    keys = WriterMessages(writer)

    BuildLinelistTables

    Assert.IsTrue writer.HasCheckings, _
                  "A table whose formulas were refused should report something"
    Assert.IsTrue (InStr(1, keys, "SpatialTables-") > 0), _
                  "The spatial entries should reach the writer report, and it carries " & _
                  "the keys " & keys

    Exit Sub
TestFail:
    BuildLinelistTables
    CustomTestLogFailure Assert, "TestTheSpatialReportReachesTheFormulaWriter", Err.Number, Err.Description
End Sub

'@section The contract with LLSpatial
'===============================================================================

'@sub-title Verify every name LLSpatial looks for is produced.
'@details
'LLSpatial reads listofgeovars, the ListObject of each level and the two columns
'it sorts on. The two classes never reference each other, so this test is the
'contract.
'@TestMethod("SpatialTables")
Public Sub TestTheNamesLLSpatialReadsAreProduced()
    CustomTestSetTitles Assert, "SpatialTables", "TestTheNamesLLSpatialReadsAreProduced"
    On Error GoTo TestFail

    Dim varName As String
    Dim counter As Long
    Dim lo As ListObject
    Dim missingName As String

    SpatialSheet
    BuildFixture SpatialRows(GEO_VARIABLE, COUNT_CALL_FUNCTION)
    WriteSpatial 1, varName

    For counter = 1 To 4
        Set lo = SpatialTable("spatial_adm" & counter & "_" & varName)

        If lo Is Nothing Then
            missingName = missingName & "[spatial_adm" & counter & "_" & varName & "]"
        Else
            If Not HasColumn(lo, "formula_adm" & counter) Then _
                missingName = missingName & "[formula_adm" & counter & "]"
            If Not HasColumn(lo, "attack_rate_adm" & counter) Then _
                missingName = missingName & "[attack_rate_adm" & counter & "]"
        End If
    Next

    Assert.AreEqual vbNullString, missingName, _
                    "Every name LLSpatial asks for should be there, and these were " & _
                    "missing: " & missingName
    Assert.IsTrue SpatialTableExists(REGISTRY_TABLE), _
                  "LLSpatial walks " & REGISTRY_TABLE & " to find the variables"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheNamesLLSpatialReadsAreProduced", Err.Number, Err.Description
End Sub

'@sub-title Verify every registry row resolves to a real ListObject.
'@details
'This is the invariant LLSpatial depends on without saying so: it reads a name
'out of the registry and asks the worksheet for the table of that name.
'@TestMethod("SpatialTables")
Public Sub TestEveryRegistryRowResolvesToATable()
    CustomTestSetTitles Assert, "SpatialTables", "TestEveryRegistryRowResolvesToATable"
    On Error GoTo TestFail

    Dim firstName As String
    Dim secondName As String
    Dim lo As ListObject
    Dim idx As Long
    Dim entry As String
    Dim unresolved As String

    SpatialSheet
    BuildFixture TwoSpatialRows()
    WriteSpatial 1, firstName
    WriteSpatial 2, secondName

    Set lo = SpatialTable(REGISTRY_TABLE)

    For idx = 1 To lo.ListRows.Count
        entry = CStr(lo.ListRows(idx).Range.Cells(1, 1).Value)
        If entry <> vbNullString Then
            If Not (SpatialTableExists("spatial_adm1_" & entry) Or _
                    SpatialTableExists("spatial_hf_" & entry)) Then
                unresolved = unresolved & "[" & entry & "]"
            End If
        End If
    Next

    Assert.AreEqual CLng(2), RegistryCount(), _
                    "Two variables give two registry rows, and the registry holds " & _
                    RegistryEntries()
    Assert.AreEqual vbNullString, unresolved, _
                    "Every registry row should resolve to a table, and these did not: " & _
                    unresolved

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEveryRegistryRowResolvesToATable", Err.Number, Err.Description
End Sub
