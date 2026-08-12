Attribute VB_Name = "TestLLSpatial"
Attribute VB_Description = "Tests for LLSpatial class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLSpatial class")

'@description
'Validates the LLSpatial class, which fills the ListObjects of the
'"__spatial_tables" worksheet from the filtered companion sheets of a linelist,
'orders them, and reads ranked values back out.
'
'THE FIXTURE COMES IN TWO SHAPES
'-------------------------------------------------------------------------------
'BuildSpatialFixture builds one hidden worksheet in this workbook carrying the
'listofgeovars table and the scratch column, which is what the factory, the
'Exists tests and the empty-table reads need.
'
'BuildSpatialWorkbook builds a whole workbook: the spatial worksheet, an HList
'worksheet whose columns carry their control values as hidden names, and the
'filtered companion sheet holding the concatenated geo values. Update walks the
'worksheets of the workbook that holds the spatial sheet, so it needs a workbook
'of its own. Everything below line 168 of the class is reached through this one.
'
'BuildAnalysisFixture is the third, smaller shape: one sheet carrying the named
'ranges of one spatial analysis table, which is what the formula rewriting
'members read.
'
'BuildSectionFixture is the fourth: one sheet carrying the named ranges of one
'spatio-temporal section, which is what PreviousSectionLevel and
'MigrateSection read.
'
'THE SPATIAL TABLES ARE BUILT BY HAND
'-------------------------------------------------------------------------------
'SpatialTables builds them in production. The fixture writes them itself, so
'these tests pin what LLSpatial does with a table of a given shape without the
'two classes having to agree first.
'
'THE COMPANION DATA CARRIES A DUPLICATE AND A BLANK
'-------------------------------------------------------------------------------
'P1 appears twice, P2 and P3 once each, and the last row is empty on every
'level. So a filled table holds three rows, P1 ranks first on a count, and a
'blank row reaching the table is visible as a fourth row.
'
'ONE TEST SWITCHES THE CALCULATED-COLUMN AUTOFILL OF THE HOST OFF
'-------------------------------------------------------------------------------
'The host writes the same formula into the same cells as the class does, so a
'filled row says nothing about which of the two filled it.
'TestUpdateWritesTheFormulasOfNewRows switches
'Application.AutoCorrect.AutoFillFormulasInLists off before it builds its
'fixture and asserts the switch took. ModuleInitialize reads what the host held
'and ModuleCleanup puts it back.
'@depends LLSpatial, CustomTest, TestHelpersLite, HiddenNames, BetterArray

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SPATIAL_SHEET As String = "__spatial_tables"
Private Const SPATIAL_WRONG As String = "WrongSheetName"
Private Const REGISTRY_TABLE As String = "listofgeovars"
Private Const HLIST_SHEET As String = "HLIST_ONE"
Private Const FILTERED_SHEET As String = "FILT_ONE"
Private Const START_NAME As String = "TAB_ONE_START"
Private Const ANALYSIS_SHEET As String = "SPT_ANALYSIS_FIX"

Private Assert As CustomTest

' What the host held for its list autofill before this module touched it, and
' whether it answered at all. The formula test switches the autofill off and
' ModuleCleanup puts it back whatever happens in between.
Private hostFillWas As Boolean
Private hostFillAnswered As Boolean

'@section Module lifecycle
'===============================================================================

'@sub-title Initialise the test module before any tests run.
'@details
'Suppresses screen updates via BusyApp, ensures the test output sheet
'exists, creates the CustomTest assertion object targeting that sheet,
'and sets the module name for result grouping. The host setting the formula
'test switches is read here, so the restore has something to put back however
'that test ends.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLSpatial"
    hostFillWas = HostFillsListFormulas(hostFillAnswered)
End Sub

'@sub-title Tear down the module after all tests complete.
'@details
'Prints accumulated test results to the output sheet, restores the
'application state via RestoreApp, releases the assertion object, and
'deletes all temporary worksheets created during the test run. The host list
'autofill goes back to what it held before the module ran.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If hostFillAnswered Then SetHostFillsListFormulas hostFillWas

    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
    DeleteWorksheets SPATIAL_SHEET, SPATIAL_WRONG, ANALYSIS_SHEET
End Sub

'@sub-title Whether the host fills the formulas of a table column by itself.
'@details
'Application.AutoCorrect.AutoFillFormulasInLists is the calculated-column
'autofill of the host. The formula test needs it off, so that whichever rows
'carry a formula afterwards carry it because LLSpatial wrote it.
'
'The answered flag is what tells a missing property apart from a property
'holding False. Without it, a host with no such member would read as "the
'autofill is already off" and the test would go back to proving nothing.
'@param answered Boolean. Set True when the host answered the read.
'@return Boolean. What the host holds.
Private Function HostFillsListFormulas(ByRef answered As Boolean) As Boolean
    answered = False

    On Error Resume Next
    Err.Clear
    HostFillsListFormulas = Application.AutoCorrect.AutoFillFormulasInLists
    answered = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'@sub-title Switch the calculated-column autofill of the host on or off.
'@param switchOn Boolean. True to let the host fill table formulas.
Private Sub SetHostFillsListFormulas(ByVal switchOn As Boolean)
    On Error Resume Next
    Application.AutoCorrect.AutoFillFormulasInLists = switchOn
    On Error GoTo 0
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

'@section Fixture helpers — one worksheet
'===============================================================================

'@sub-title Build a spatial worksheet in this workbook.
'@details
'Creates a hidden worksheet named "__spatial_tables" carrying the
'"listofgeovars" ListObject and the RNG_PastingCol scratch cell. When addVars
'is True the table holds "cases_sp1" and "deaths_sp1", so Exists lookups have
'something to find. When withRegistry is False the table is left out, which is
'the shape of a linelist that has geo variables and no spatial analysis.
'@param addVars Optional Boolean. True to populate sample variable rows. Defaults to True.
'@param withRegistry Optional Boolean. True to build the listofgeovars table. Defaults to True.
'@return Worksheet. The prepared spatial fixture sheet.
Private Function BuildSpatialFixture(Optional ByVal addVars As Boolean = True, _
                                     Optional ByVal withRegistry As Boolean = True) As Worksheet
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(SPATIAL_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    If withRegistry Then
        BuildRegistry sh

        If addVars Then
            RegisterSpatialVar sh, "cases_sp1"
            RegisterSpatialVar sh, "deaths_sp1"
        End If
    End If

    sh.Cells(1, 5).Value = "scratch"
    sh.Cells(1, 5).Name = "RNG_PastingCol"

    Set BuildSpatialFixture = sh
End Function

'@sub-title Write the listofgeovars table on a spatial worksheet.
'@details
'Column 3 and the header "listofvars" are what SpatialTables.Prepare writes.
'The table starts with one empty data row, the way a ListObject built over two
'cells does.
'@param sh Worksheet. The spatial worksheet.
Private Sub BuildRegistry(ByVal sh As Worksheet)
    Dim rng As Range

    sh.Cells(1, 3).Value = "listofvars"
    Set rng = sh.Range(sh.Cells(1, 3), sh.Cells(2, 3))
    sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = REGISTRY_TABLE
End Sub

'@sub-title Register one spatial variable name.
'@details
'Fills the empty starting row first, then adds a row, which is what
'SpatialTables.AddVarNameToList does.
'@param sh Worksheet. The spatial worksheet.
'@param varName String. The registered name, "<variable>_<tableId>".
Private Sub RegisterSpatialVar(ByVal sh As Worksheet, ByVal varName As String)
    Dim Lo As ListObject
    Dim newRow As ListRow

    Set Lo = sh.ListObjects(REGISTRY_TABLE)

    If Lo.ListRows.Count = 0 Then
        Set newRow = Lo.ListRows.Add
    ElseIf Lo.ListRows(Lo.ListRows.Count).Range.Cells(1, 1).Value = vbNullString Then
        Set newRow = Lo.ListRows(Lo.ListRows.Count)
    Else
        Set newRow = Lo.ListRows.Add
    End If

    newRow.Range.Cells(1, 1).Value = varName
End Sub

'@section Fixture helpers — a whole workbook
'===============================================================================

'@sub-title Build a workbook holding a spatial worksheet.
'@param withRegistry Boolean. True to build the listofgeovars table.
'@return Workbook. The workbook, with the spatial sheet as its first worksheet.
Private Function BuildSpatialWorkbook(ByVal withRegistry As Boolean) As Workbook
    Dim wb As Workbook
    Dim sh As Worksheet

    Set wb = NewWorkbook()
    Set sh = wb.Worksheets(1)
    sh.Name = SPATIAL_SHEET

    If withRegistry Then BuildRegistry sh

    sh.Cells(1, 5).Value = "scratch"
    sh.Cells(1, 5).Name = "RNG_PastingCol"

    Set BuildSpatialWorkbook = wb
End Function

'@sub-title Add the filtered companion sheet of an HList worksheet.
'@details
'One ListObject carrying the four concat columns of the "cases" variable. P1
'appears twice and the last row is empty on every level, so a correct update
'writes three rows and no blank one.
'@param wb Workbook. The workbook to add it to.
'@return Worksheet. The filtered companion sheet.
Private Function AddFilteredSheet(ByVal wb As Workbook) As Worksheet
    Dim sh As Worksheet

    Set sh = wb.Worksheets.Add
    sh.Name = FILTERED_SHEET

    WriteRow sh.Cells(1, 1), "concat_adm1_cases", "concat_adm2_cases", _
                             "concat_adm3_cases", "concat_adm4_cases"
    WriteRow sh.Cells(2, 1), "P1", "P1D1", "P1D1S1", "P1D1S1V1"
    WriteRow sh.Cells(3, 1), "P2", "P2D1", "P2D1S1", "P2D1S1V1"
    WriteRow sh.Cells(4, 1), "P1", "P1D1", "P1D1S1", "P1D1S1V1"
    WriteRow sh.Cells(5, 1), "P3", "P3D1", "P3D1S1", "P3D1S1V1"

    sh.ListObjects.Add(xlSrcRange, sh.Range(sh.Cells(1, 1), sh.Cells(6, 4)), , xlYes).Name = _
        "filtered_cases"

    Set AddFilteredSheet = sh
End Function

'@sub-title Add an HList worksheet the way a linelist carries one.
'@details
'The sheet gets a sheet_type of HList, a table_name, the name of its filtered
'companion, and one column per variable name given. Row 8 holds the variable
'name and the _START cell, and each column carries its control value under the
'key VarWriter writes: "<variable> -- control". The columns of one geo variable
'take twelve places on the sheet, so the next one starts twelve columns along.
'@param wb Workbook. The workbook to add it to.
'@param columnNames Variant. The header value of each geo column, "adm1_<variable>".
'@return Worksheet. The HList worksheet.
Private Function AddHListSheet(ByVal wb As Workbook, ByVal columnNames As Variant) As Worksheet
    Dim sh As Worksheet
    Dim store As HiddenNames
    Dim counter As Long
    Dim colIndex As Long

    Set sh = wb.Worksheets.Add
    sh.Name = HLIST_SHEET

    Set store = HiddenNames.Create(sh)
    store.EnsureName "sheet_type", "HList", HiddenNameTypeString
    store.EnsureName "table_name", "TAB_ONE", HiddenNameTypeString
    store.EnsureName "filtered_sheet", FILTERED_SHEET, HiddenNameTypeString

    colIndex = 2

    For counter = LBound(columnNames) To UBound(columnNames)
        sh.Cells(8, colIndex).Value = columnNames(counter)
        store.EnsureName CStr(columnNames(counter)) & " -- control", "geo1", _
                         HiddenNameTypeString

        If counter = LBound(columnNames) Then sh.Cells(8, colIndex).Name = START_NAME

        colIndex = colIndex + 12
    Next

    Set AddHListSheet = sh
End Function

'@sub-title The column a new spatial table starts in.
'@details
'The tables sit left to right with a gap of one column between them, and the
'scratch column at 5 is the left edge of the first one.
'@param sh Worksheet. The spatial worksheet.
'@return Long. The first free column.
Private Function NextTableColumn(ByVal sh As Worksheet) As Long
    Dim Lo As ListObject
    Dim lastCol As Long
    Dim endCol As Long

    lastCol = 5

    For Each Lo In sh.ListObjects
        endCol = Lo.Range.Column + Lo.Range.Columns.Count - 1
        If endCol > lastCol Then lastCol = endCol
    Next

    NextTableColumn = lastCol + 2
End Function

'@sub-title The address of one concat column of the filtered sheet.
'@param filtSh Worksheet. The filtered companion sheet.
'@param colIndex Long. The column number.
'@return String. A sheet qualified absolute address.
Private Function FilteredColumnAddress(ByVal filtSh As Worksheet, _
                                       ByVal colIndex As Long) As String
    FilteredColumnAddress = "'" & filtSh.Name & "'!" & _
        filtSh.Range(filtSh.Cells(2, colIndex), filtSh.Cells(50, colIndex)).Address
End Function

'@sub-title Build the four administrative tables of one spatial variable.
'@details
'Each table is a header row and one data row over four columns, with the value
'column written as an array formula, which is what SpatialTables does. An array
'formula is the least likely of all to be copied down by the calculated-column
'autofill of the host, so a table filled without the class copying the formulas
'itself holds a number in its first row alone.
'
'A Nothing companion sheet leaves the value column empty, which is what the
'tests that write their own numbers into it want.
'@param sh Worksheet. The spatial worksheet.
'@param varName String. The registered name, "<variable>_<tableId>".
'@param filtSh Worksheet. The filtered companion the value formula counts over, or Nothing.
Private Sub AddSpatialTables(ByVal sh As Worksheet, ByVal varName As String, _
                             ByVal filtSh As Worksheet)
    Dim counter As Long
    Dim colIndex As Long
    Dim adminName As String
    Dim cellRng As Range

    For counter = 1 To 4
        adminName = "adm" & counter
        colIndex = NextTableColumn(sh)
        Set cellRng = sh.Cells(1, colIndex)

        cellRng.Value = "tabl_" & adminName & "_" & varName
        cellRng.Cells(1, 2).Value = "formula_" & adminName
        cellRng.Cells(1, 3).Value = "population_" & adminName
        cellRng.Cells(1, 4).Value = "attack_rate_" & adminName

        If Not filtSh Is Nothing Then
            cellRng.Cells(2, 2).FormulaArray = "=COUNTIF(" & _
                FilteredColumnAddress(filtSh, counter) & "," & _
                cellRng.Cells(2, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
        End If

        cellRng.Cells(2, 3).Formula = "=1000"
        cellRng.Cells(2, 4).Formula = "=" & _
            cellRng.Cells(2, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "/" & _
            cellRng.Cells(2, 3).Address(RowAbsolute:=False, ColumnAbsolute:=False)

        sh.ListObjects.Add(xlSrcRange, _
                           sh.Range(cellRng.Cells(1, 1), cellRng.Cells(2, 4)), _
                           , xlYes).Name = "spatial_" & adminName & "_" & varName
    Next
End Sub

'@sub-title Build the single facility table of one spatial variable.
'@details
'A facility table is two columns: the lookup key and the value. A health
'facility has no population.
'@param sh Worksheet. The spatial worksheet.
'@param varName String. The registered name, "<variable>_<tableId>".
Private Sub AddHFTable(ByVal sh As Worksheet, ByVal varName As String)
    Dim colIndex As Long
    Dim cellRng As Range

    colIndex = NextTableColumn(sh)
    Set cellRng = sh.Cells(1, colIndex)

    cellRng.Value = "tabl_hf_" & varName
    cellRng.Cells(1, 2).Value = "formula_hf"
    cellRng.Cells(2, 1).Value = "Clinic A"
    cellRng.Cells(2, 2).Value = 7

    sh.ListObjects.Add(xlSrcRange, _
                       sh.Range(cellRng.Cells(1, 1), cellRng.Cells(2, 2)), _
                       , xlYes).Name = "spatial_hf_" & varName
End Sub

'@sub-title Write keys and values into a spatial table by hand.
'@details
'Grows the table to the number of keys given and writes the key column and the
'value column, so a test can order a table without a companion sheet behind it.
'@param sh Worksheet. The spatial worksheet.
'@param loName String. The ListObject name.
'@param keys Variant. The key column values.
'@param values Variant. The value column values, in the same order.
Private Sub FillTableByHand(ByVal sh As Worksheet, ByVal loName As String, _
                            ByVal keys As Variant, ByVal values As Variant)
    Dim Lo As ListObject
    Dim headerRng As Range
    Dim counter As Long
    Dim rowIndex As Long

    Set Lo = sh.ListObjects(loName)
    Set headerRng = Lo.HeaderRowRange

    Lo.Resize sh.Range(headerRng.Cells(1, 1), _
                       headerRng.Cells(UBound(keys) - LBound(keys) + 2, _
                                       Lo.ListColumns.Count))

    For counter = LBound(keys) To UBound(keys)
        rowIndex = counter - LBound(keys) + 1
        Lo.ListColumns(1).DataBodyRange.Cells(rowIndex, 1).Value = keys(counter)
        Lo.ListColumns(2).DataBodyRange.Cells(rowIndex, 1).Value = values(counter)
    Next
End Sub

'@sub-title Read one cell of the key column of a spatial table.
'@param sh Worksheet. The spatial worksheet.
'@param loName String. The ListObject name.
'@param rowIndex Long. The data row to read, 1 for the first.
'@return String. The value held there.
Private Function KeyAt(ByVal sh As Worksheet, ByVal loName As String, _
                       ByVal rowIndex As Long) As String
    KeyAt = CStr(sh.ListObjects(loName).ListColumns(1).DataBodyRange.Cells(rowIndex, 1).Value)
End Function

'@sub-title How many data rows a spatial table holds.
'@param sh Worksheet. The spatial worksheet.
'@param loName String. The ListObject name.
'@return Long. The row count.
Private Function RowCountOf(ByVal sh As Worksheet, ByVal loName As String) As Long
    RowCountOf = sh.ListObjects(loName).ListRows.Count
End Function

'@section Factory validation tests
'===============================================================================

'@sub-title Verify Create refuses a Nothing worksheet with ObjectNotInitialized.
'@details
'Acts by calling LLSpatial.Create with Nothing under On Error Resume Next and
'captures the error number. Asserts the number is ObjectNotInitialized, so the
'test fails when the factory raises anything else. It used to assert that the
'result was Nothing, which passes whatever the failure was.
'@TestMethod("LLSpatial")
Public Sub TestCreateRejectsNothing()
    CustomTestSetTitles Assert, "LLSpatial", "TestCreateRejectsNothing"
    On Error GoTo TestFail

    Dim sp As LLSpatial
    Dim errNumber As Long
    Dim errDescription As String

    On Error Resume Next
    Set sp = LLSpatial.Create(Nothing)
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "Create with a Nothing sheet should raise ObjectNotInitialized - " & _
                    "description was [" & errDescription & "]"
    Assert.IsTrue (sp Is Nothing), _
                  "Create with Nothing sheet should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothing", Err.Number, Err.Description
End Sub

'@sub-title Verify Create refuses a wrongly named worksheet with InvalidArgument.
'@details
'Arranges a hidden worksheet named "WrongSheetName". Acts by calling
'LLSpatial.Create with it under On Error Resume Next and captures the error
'number. Asserts the number is InvalidArgument.
'@TestMethod("LLSpatial")
Public Sub TestCreateRejectsWrongSheetName()
    CustomTestSetTitles Assert, "LLSpatial", "TestCreateRejectsWrongSheetName"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial
    Dim errNumber As Long
    Dim errDescription As String

    Set sh = EnsureWorksheet(SPATIAL_WRONG, clearSheet:=True, visibility:=xlSheetHidden)

    On Error Resume Next
    Set sp = LLSpatial.Create(sh)
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "Create with a wrongly named sheet should raise InvalidArgument - " & _
                    "description was [" & errDescription & "]"
    Assert.IsTrue (sp Is Nothing), _
                  "Create with wrong sheet name should hand back nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsWrongSheetName", Err.Number, Err.Description
End Sub

'@sub-title Verify Create accepts the spatial worksheet.
'@details
'Arranges a spatial fixture with the correct sheet name. Acts by calling
'LLSpatial.Create with it. Asserts the result is not Nothing. The factory
'checks the name of the sheet, and the listofgeovars table is left to the
'members that read it, so this test says nothing about the table.
'@TestMethod("LLSpatial")
Public Sub TestCreateSucceedsWithCorrectSheet()
    CustomTestSetTitles Assert, "LLSpatial", "TestCreateSucceedsWithCorrectSheet"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture()

    Dim sp As LLSpatial
    Set sp = LLSpatial.Create(sh)

    Assert.IsNotNothing sp, _
                        "Create with correctly named sheet should succeed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateSucceedsWithCorrectSheet", Err.Number, Err.Description
End Sub

'@section Exists tests
'===============================================================================

'@sub-title Verify Exists returns True for a variable present in listofgeovars.
'@details
'Arranges a spatial fixture with "cases_sp1" and "deaths_sp1" in the
'listofgeovars table. Acts by creating an LLSpatial instance and calling
'Exists("cases"). Asserts that the result is True, confirming a registered
'name built on that variable is found.
'@TestMethod("LLSpatial")
Public Sub TestExistsReturnsTrueForKnownVar()
    CustomTestSetTitles Assert, "LLSpatial", "TestExistsReturnsTrueForKnownVar"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture(addVars:=True)

    Dim sp As LLSpatial
    Set sp = LLSpatial.Create(sh)

    Assert.IsTrue sp.Exists("cases"), _
                  "Exists should return True for a variable matching 'cases'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestExistsReturnsTrueForKnownVar", Err.Number, Err.Description
End Sub

'@sub-title Verify Exists returns False for a variable not in listofgeovars.
'@details
'Arranges a spatial fixture with known variables. Acts by creating an
'LLSpatial instance and calling Exists("nonexistent_var"). Asserts that the
'result is False.
'@TestMethod("LLSpatial")
Public Sub TestExistsReturnsFalseForUnknownVar()
    CustomTestSetTitles Assert, "LLSpatial", "TestExistsReturnsFalseForUnknownVar"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture(addVars:=True)

    Dim sp As LLSpatial
    Set sp = LLSpatial.Create(sh)

    Assert.IsFalse sp.Exists("nonexistent_var"), _
                   "Exists should return False for an unknown variable"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestExistsReturnsFalseForUnknownVar", Err.Number, Err.Description
End Sub

'@sub-title Verify Exists reads the rows of the table and not its header.
'@details
'Arranges a spatial fixture whose listofgeovars table holds no variable. Acts
'by asking for "list", which is part of the header value "listofvars". Asserts
'the answer is False. A search over the whole table range answered True here.
'@TestMethod("LLSpatial")
Public Sub TestExistsIgnoresTheTableHeader()
    CustomTestSetTitles Assert, "LLSpatial", "TestExistsIgnoresTheTableHeader"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture(addVars:=False)

    Dim sp As LLSpatial
    Set sp = LLSpatial.Create(sh)

    Assert.IsFalse sp.Exists("list"), _
                   "The header of the variables table should not answer a lookup"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestExistsIgnoresTheTableHeader", Err.Number, Err.Description
End Sub

'@section Update tests
'===============================================================================

'@sub-title Verify Update fills the admin table with the distinct values.
'@details
'Arranges a workbook with a spatial sheet, an HList sheet whose one column is
'controlled by geo1, and a filtered companion holding P1, P2, P1, P3 and a
'blank. Acts by calling Update. Asserts the table holds three rows and P1 is
'first, which needs the control value read under the key VarWriter writes and
'the table ordered on a key inside itself. Update matched nothing at all until
'that key was corrected.
'@TestMethod("LLSpatial")
Public Sub TestUpdateFillsTheAdminTable()
    CustomTestSetTitles Assert, "LLSpatial", "TestUpdateFillsTheAdminTable"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim filtSh As Worksheet
    Dim sp As LLSpatial

    Set wb = BuildSpatialWorkbook(withRegistry:=True)
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    Set filtSh = AddFilteredSheet(wb)
    AddHListSheet wb, Array("adm1_cases")
    RegisterSpatialVar sh, "cases_sp1"
    AddSpatialTables sh, "cases_sp1", filtSh

    Set sp = LLSpatial.Create(sh)
    sp.Update

    Assert.AreEqual CLng(3), RowCountOf(sh, "spatial_adm1_cases_sp1"), _
                    "The admin 1 table should hold the three distinct values"
    Assert.AreEqual "P1", KeyAt(sh, "spatial_adm1_cases_sp1", 1), _
                    "The value counted twice should be ranked first"
    Assert.AreEqual CLng(3), RowCountOf(sh, "spatial_adm4_cases_sp1"), _
                    "The admin 4 table should be filled from the fourth concat column"
    Assert.AreEqual "P1D1S1V1", KeyAt(sh, "spatial_adm4_cases_sp1", 1), _
                    "The admin 4 table should carry admin 4 values"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateFillsTheAdminTable", Err.Number, Err.Description
End Sub

'@sub-title Verify Update puts no blank row in a spatial table.
'@details
'Arranges the standard workbook, whose filtered companion carries an empty last
'row. Acts by calling Update. Asserts every key of the table holds something.
'GeoConcat answers an empty string for a record whose geo levels are not all
'filled, RemoveDuplicates collapsed those to one blank, and that blank became a
'data row that could take rank 1.
'@TestMethod("LLSpatial")
Public Sub TestUpdateDropsTheBlankValues()
    CustomTestSetTitles Assert, "LLSpatial", "TestUpdateDropsTheBlankValues"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim filtSh As Worksheet
    Dim sp As LLSpatial
    Dim counter As Long
    Dim blanks As Long

    Set wb = BuildSpatialWorkbook(withRegistry:=True)
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    Set filtSh = AddFilteredSheet(wb)
    AddHListSheet wb, Array("adm1_cases")
    RegisterSpatialVar sh, "cases_sp1"
    AddSpatialTables sh, "cases_sp1", filtSh

    Set sp = LLSpatial.Create(sh)
    sp.Update

    For counter = 1 To RowCountOf(sh, "spatial_adm1_cases_sp1")
        If LenB(KeyAt(sh, "spatial_adm1_cases_sp1", counter)) = 0 Then blanks = blanks + 1
    Next

    Assert.AreEqual CLng(0), blanks, _
                    "A blank concat value should not become a row of the table"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateDropsTheBlankValues", Err.Number, Err.Description
End Sub

'@sub-title Verify Update writes the formulas of every new row.
'@details
'Arranges the standard workbook and calls Update, which grows the table from
'one row to three. Asserts rows 2 and 3 carry the value formula and the attack
'rate. The class used to leave those rows to the calculated-column autofill of
'the host, which is an option a user can switch off and which an array formula
'seldom triggers, so only the first row held a number.
'
'THE HOST AUTOFILL IS SWITCHED OFF FIRST, AND THAT IS THE POINT OF THE TEST
'-------------------------------------------------------------------------------
'The host and the class write the same formula into the same cells, so a row
'carrying a formula proves nothing about which of the two put it there. With
'Application.AutoCorrect.AutoFillFormulasInLists off, the host cannot have
'written it. The switch is asserted to have taken, so a host with no such member
'fails here rather than quietly turning this back into a test of the host.
'@TestMethod("LLSpatial")
Public Sub TestUpdateWritesTheFormulasOfNewRows()
    CustomTestSetTitles Assert, "LLSpatial", "TestUpdateWritesTheFormulasOfNewRows"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim filtSh As Worksheet
    Dim sp As LLSpatial
    Dim Lo As ListObject
    Dim hostFills As Boolean
    Dim hostAnswered As Boolean

    'The switch comes before the fixture, so the value column is never made a
    'calculated column in the first place.
    SetHostFillsListFormulas False
    hostFills = HostFillsListFormulas(hostAnswered)

    Set wb = BuildSpatialWorkbook(withRegistry:=True)
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    Set filtSh = AddFilteredSheet(wb)
    AddHListSheet wb, Array("adm1_cases")
    RegisterSpatialVar sh, "cases_sp1"
    AddSpatialTables sh, "cases_sp1", filtSh

    Set sp = LLSpatial.Create(sh)
    sp.Update

    If hostFillAnswered Then SetHostFillsListFormulas hostFillWas

    Set Lo = sh.ListObjects("spatial_adm1_cases_sp1")

    Assert.IsTrue (hostAnswered And Not hostFills), _
                  "The host list autofill should read as off, so the rows below the " & _
                  "first carry a formula because the class wrote it"
    Assert.IsTrue (LenB(Lo.ListColumns("formula_adm1").DataBodyRange.Cells(2, 1).Formula) > 0), _
                  "The second row should carry the value formula"
    Assert.IsTrue (LenB(Lo.ListColumns("formula_adm1").DataBodyRange.Cells(3, 1).Formula) > 0), _
                  "The third row should carry the value formula"
    Assert.AreEqual CLng(1), _
                    CLng(Lo.ListColumns("formula_adm1").DataBodyRange.Cells(3, 1).Value), _
                    "The third row should count the records of its own admin unit"
    Assert.IsTrue (LenB(Lo.ListColumns("attack_rate_adm1").DataBodyRange.Cells(3, 1).Formula) > 0), _
                  "The third row should carry the attack rate"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    If hostFillAnswered Then SetHostFillsListFormulas hostFillWas
    CustomTestLogFailure Assert, "TestUpdateWritesTheFormulasOfNewRows", Err.Number, Err.Description
End Sub

'@sub-title Verify a variable with no concat column leaves its tables alone.
'@details
'Arranges a workbook with two geo variables on the HList sheet, where the
'companion sheet carries the concat columns of the first alone. Acts by calling
'Update. Asserts the tables of the second are untouched. The data range was
'declared once for the whole walk and never reset, so a failed column lookup
'left the range of the variable before it in hand and its values were written
'into the tables of this one.
'@TestMethod("LLSpatial")
Public Sub TestUpdateLeavesAVariableWithNoConcatColumnAlone()
    CustomTestSetTitles Assert, "LLSpatial", "TestUpdateLeavesAVariableWithNoConcatColumnAlone"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim filtSh As Worksheet
    Dim sp As LLSpatial

    Set wb = BuildSpatialWorkbook(withRegistry:=True)
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    Set filtSh = AddFilteredSheet(wb)
    AddHListSheet wb, Array("adm1_cases", "adm1_deaths")
    RegisterSpatialVar sh, "cases_sp1"
    RegisterSpatialVar sh, "deaths_sp1"
    AddSpatialTables sh, "cases_sp1", filtSh
    AddSpatialTables sh, "deaths_sp1", Nothing

    Set sp = LLSpatial.Create(sh)
    sp.Update

    Assert.AreEqual CLng(3), RowCountOf(sh, "spatial_adm1_cases_sp1"), _
                    "The variable with a concat column should be filled"
    Assert.AreEqual CLng(1), RowCountOf(sh, "spatial_adm1_deaths_sp1"), _
                    "The variable with no concat column should keep its single row"
    Assert.AreEqual vbNullString, KeyAt(sh, "spatial_adm1_deaths_sp1", 1), _
                    "The variable with no concat column should take no values"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateLeavesAVariableWithNoConcatColumnAlone", Err.Number, Err.Description
End Sub

'@sub-title Verify a variable whose name contains another one is left alone.
'@details
'Arranges a workbook where "cases_sp1" and "new_cases_sp1" are both registered
'and the HList sheet carries the "cases" column alone. Acts by calling Update.
'Asserts the tables of "new_cases" are untouched. A substring lookup answered
'with both names, so updating "cases" rewrote the four tables of "new_cases"
'from the "concat_adm1_cases" column and the values looked plausible.
'@TestMethod("LLSpatial")
Public Sub TestUpdateLeavesAVariableSharingASubstringAlone()
    CustomTestSetTitles Assert, "LLSpatial", "TestUpdateLeavesAVariableSharingASubstringAlone"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim filtSh As Worksheet
    Dim sp As LLSpatial

    Set wb = BuildSpatialWorkbook(withRegistry:=True)
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    Set filtSh = AddFilteredSheet(wb)
    AddHListSheet wb, Array("adm1_cases")
    RegisterSpatialVar sh, "cases_sp1"
    RegisterSpatialVar sh, "new_cases_sp1"
    AddSpatialTables sh, "cases_sp1", filtSh
    AddSpatialTables sh, "new_cases_sp1", Nothing

    Set sp = LLSpatial.Create(sh)
    sp.Update

    Assert.AreEqual CLng(1), RowCountOf(sh, "spatial_adm1_new_cases_sp1"), _
                    "The tables of new_cases should not move when cases is updated"
    Assert.AreEqual vbNullString, KeyAt(sh, "spatial_adm1_new_cases_sp1", 1), _
                    "The tables of new_cases should take no value from cases"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateLeavesAVariableSharingASubstringAlone", Err.Number, Err.Description
End Sub

'@sub-title Verify the number format of the key column survives an update.
'@details
'Arranges the standard workbook, formats the key column as text, and calls
'Update twice. Asserts the format is still text. Clear took the number formats,
'the borders and the validation of the column with the values.
'@TestMethod("LLSpatial")
Public Sub TestUpdateKeepsTheNumberFormatOfTheKeyColumn()
    CustomTestSetTitles Assert, "LLSpatial", "TestUpdateKeepsTheNumberFormatOfTheKeyColumn"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim filtSh As Worksheet
    Dim sp As LLSpatial
    Dim Lo As ListObject

    Set wb = BuildSpatialWorkbook(withRegistry:=True)
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    Set filtSh = AddFilteredSheet(wb)
    AddHListSheet wb, Array("adm1_cases")
    RegisterSpatialVar sh, "cases_sp1"
    AddSpatialTables sh, "cases_sp1", filtSh

    Set Lo = sh.ListObjects("spatial_adm1_cases_sp1")
    Lo.ListColumns(1).Range.NumberFormat = "@"

    Set sp = LLSpatial.Create(sh)
    sp.Update
    sp.Update

    Assert.AreEqual "@", _
                    CStr(sh.ListObjects("spatial_adm1_cases_sp1"). _
                         ListColumns(1).DataBodyRange.Cells(1, 1).NumberFormat), _
                    "The key column should keep its number format across an update"
    Assert.AreEqual CLng(3), RowCountOf(sh, "spatial_adm1_cases_sp1"), _
                    "A second update should leave the same three rows"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateKeepsTheNumberFormatOfTheKeyColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify Update is a clean no-op with no variables table.
'@details
'Arranges a workbook whose spatial sheet carries no listofgeovars table, which
'is what a linelist with geo variables and no spatial analysis holds. Acts by
'calling Update under On Error Resume Next. Asserts nothing was raised. Every
'member that reads the table used to reach it unguarded, so the Calculate button
'raised on the first HList sheet and the four analysis sheets were never
'recalculated.
'@TestMethod("LLSpatial")
Public Sub TestUpdateWithNoVariablesTableDoesNotRaise()
    CustomTestSetTitles Assert, "LLSpatial", "TestUpdateWithNoVariablesTableDoesNotRaise"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim sp As LLSpatial
    Dim errNumber As Long
    Dim errDescription As String

    Set wb = BuildSpatialWorkbook(withRegistry:=False)
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    AddFilteredSheet wb
    AddHListSheet wb, Array("adm1_cases")

    Set sp = LLSpatial.Create(sh)

    On Error Resume Next
    sp.Update
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(0), errNumber, _
                    "Update on a sheet with no variables table should raise nothing - " & _
                    "description was [" & errDescription & "]"
    Assert.IsFalse sp.Exists("cases"), _
                   "Exists should answer False when the variables table is absent"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUpdateWithNoVariablesTableDoesNotRaise", Err.Number, Err.Description
End Sub

'@section Sort tests
'===============================================================================

'@sub-title Verify Sort orders a table on its value column.
'@details
'Arranges a table holding three rows written out of order. Acts by calling
'Sort. Asserts the key column comes back highest value first. Excel wants the
'sort key inside the range it sorts and raises 1004 when it sits outside, and
'this class ordered the first column alone with a key in the second, under On
'Error Resume Next, so nothing was ever ordered.
'@TestMethod("LLSpatial")
Public Sub TestSortOrdersTheTableOnItsValueColumn()
    CustomTestSetTitles Assert, "LLSpatial", "TestSortOrdersTheTableOnItsValueColumn"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set wb = BuildSpatialWorkbook(withRegistry:=True)
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    RegisterSpatialVar sh, "cases_sp1"
    AddSpatialTables sh, "cases_sp1", Nothing
    FillTableByHand sh, "spatial_adm1_cases_sp1", _
                    Array("P1", "P2", "P3"), Array(1, 5, 3)

    Set sp = LLSpatial.Create(sh)
    sp.Sort "sp1"

    Assert.AreEqual "P2", KeyAt(sh, "spatial_adm1_cases_sp1", 1), _
                    "The highest value should be ranked first"
    Assert.AreEqual "P3", KeyAt(sh, "spatial_adm1_cases_sp1", 2), _
                    "The middle value should be ranked second"
    Assert.AreEqual "P1", KeyAt(sh, "spatial_adm1_cases_sp1", 3), _
                    "The lowest value should be ranked last"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSortOrdersTheTableOnItsValueColumn", Err.Number, Err.Description
End Sub

'@sub-title Verify Sort picks the table whose identifier was asked for.
'@details
'Arranges "cases_sp10" registered ahead of "cases_sp1", both with tables
'holding three rows out of order. Acts by calling Sort with "sp1". Asserts the
'sp1 table is ordered and the sp10 table is untouched. A substring lookup found
'"cases_sp10" first, so the day a linelist carried ten spatial cross-tables the
'wrong table moved.
'@TestMethod("LLSpatial")
Public Sub TestSortPicksTheExactTableId()
    CustomTestSetTitles Assert, "LLSpatial", "TestSortPicksTheExactTableId"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set wb = BuildSpatialWorkbook(withRegistry:=True)
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    RegisterSpatialVar sh, "cases_sp10"
    RegisterSpatialVar sh, "cases_sp1"
    AddSpatialTables sh, "cases_sp10", Nothing
    AddSpatialTables sh, "cases_sp1", Nothing
    FillTableByHand sh, "spatial_adm1_cases_sp10", _
                    Array("Q1", "Q2", "Q3"), Array(1, 5, 3)
    FillTableByHand sh, "spatial_adm1_cases_sp1", _
                    Array("P1", "P2", "P3"), Array(1, 5, 3)

    Set sp = LLSpatial.Create(sh)
    sp.Sort "sp1"

    Assert.AreEqual "P2", KeyAt(sh, "spatial_adm1_cases_sp1", 1), _
                    "The table asked for should be ordered"
    Assert.AreEqual "Q1", KeyAt(sh, "spatial_adm1_cases_sp10", 1), _
                    "The table of another identifier should not move"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSortPicksTheExactTableId", Err.Number, Err.Description
End Sub

'@section Formula rewriting fixture
'===============================================================================

'@sub-title Build a small spatial analysis sheet for the formula rewriting.
'@details
'Names the block B2:C3 as OUTER_VALUES, the column B2:B3 as ROW_CATEGORIES,
'and one cell each for PREVIOUS_ADM, POPFACT, POPPREVFACT and POPFACTLABEL,
'all suffixed with the given identifier and all scoped to the sheet, so
'deleting the sheet takes the names with it. The formulas are written by each
'test, so a test says which cells are plain and which are array ones. C2 is
'the one data cell right of the row categories, which is where the population
'division applies.
'@param tabId String. The table identifier the names carry.
'@return Worksheet. The prepared analysis fixture sheet.
Private Function BuildAnalysisFixture(ByVal tabId As String) As Worksheet
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(ANALYSIS_SHEET, clearSheet:=True, _
                             visibility:=xlSheetHidden)

    sh.Names.Add Name:="OUTER_VALUES_" & tabId, RefersTo:=sh.Range("B2:C3")
    sh.Names.Add Name:="ROW_CATEGORIES_" & tabId, RefersTo:=sh.Range("B2:B3")
    sh.Names.Add Name:="PREVIOUS_ADM_" & tabId, RefersTo:=sh.Cells(1, 5)
    sh.Names.Add Name:="POPFACT_" & tabId, RefersTo:=sh.Cells(2, 5)
    sh.Names.Add Name:="POPPREVFACT_" & tabId, RefersTo:=sh.Cells(3, 5)
    sh.Names.Add Name:="POPFACTLABEL_" & tabId, RefersTo:=sh.Cells(4, 5)

    Set BuildAnalysisFixture = sh
End Function

'@section Formula rewriting tests
'===============================================================================

'@sub-title Verify RewriteFormulas keeps a plain formula plain.
'@details
'Arranges one cell holding a plain formula that reads a concat column. Acts by
'rewriting the admin token. Asserts the formula carries the new token and the
'cell still holds a plain formula. The old readers pushed every write through
'FormulaArray, which turned the cell into an array formula.
'@TestMethod("LLSpatial")
Public Sub TestRewriteFormulasKeepsAPlainFormulaPlain()
    CustomTestSetTitles Assert, "LLSpatial", "TestRewriteFormulasKeepsAPlainFormulaPlain"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set sh = BuildAnalysisFixture("tsta")
    sh.Range("B2").Formula = "=SUM(concat_adm1_cases)"

    Set sp = LLSpatial.Create(BuildSpatialFixture())
    sp.RewriteFormulas sh.Range("B2"), "concat_adm1", "concat_adm3"

    Assert.IsTrue (InStr(1, sh.Range("B2").Formula, "concat_adm3_cases") > 0), _
                  "The formula should read the new concat column"
    Assert.IsTrue (Not sh.Range("B2").HasArray), _
                  "A plain formula should stay plain after the rewrite"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRewriteFormulasKeepsAPlainFormulaPlain", Err.Number, Err.Description
End Sub

'@sub-title Verify RewriteFormulas keeps an array formula an array one.
'@details
'Arranges one cell holding an array formula that reads a concat column. Acts
'by rewriting the admin token. Asserts the formula carries the new token and
'the cell still holds an array formula.
'@TestMethod("LLSpatial")
Public Sub TestRewriteFormulasKeepsAnArrayFormula()
    CustomTestSetTitles Assert, "LLSpatial", "TestRewriteFormulasKeepsAnArrayFormula"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set sh = BuildAnalysisFixture("tstb")
    sh.Range("B2").FormulaArray = "=SUM(concat_adm1_cases)"

    Set sp = LLSpatial.Create(BuildSpatialFixture())
    sp.RewriteFormulas sh.Range("B2"), "concat_adm1", "concat_adm3"

    Assert.IsTrue (InStr(1, sh.Range("B2").FormulaArray, "concat_adm3_cases") > 0), _
                  "The array formula should read the new concat column"
    Assert.IsTrue sh.Range("B2").HasArray, _
                  "An array formula should stay an array one after the rewrite"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRewriteFormulasKeepsAnArrayFormula", Err.Number, Err.Description
End Sub

'@sub-title Verify RewriteFormulas leaves a formula without the token alone.
'@details
'Arranges one cell holding a formula free of the token. Acts by rewriting.
'Asserts the formula is byte for byte what it was.
'@TestMethod("LLSpatial")
Public Sub TestRewriteFormulasLeavesOtherFormulasAlone()
    CustomTestSetTitles Assert, "LLSpatial", "TestRewriteFormulasLeavesOtherFormulasAlone"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set sh = BuildAnalysisFixture("tstc")
    sh.Range("B2").Formula = "=SUM(1,2)"

    Set sp = LLSpatial.Create(BuildSpatialFixture())
    sp.RewriteFormulas sh.Range("B2"), "concat_adm1", "concat_adm3"

    Assert.AreEqual "=SUM(1,2)", sh.Range("B2").Formula, _
                    "A formula without the token should be untouched"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRewriteFormulasLeavesOtherFormulasAlone", Err.Number, Err.Description
End Sub

'@sub-title Verify RewriteFormulas writes a plain formula over 255 characters.
'@details
'Arranges one cell holding a plain formula longer than 255 characters that
'reads a concat column. Acts by rewriting the admin token. Asserts the write
'lands and the formula carries the new token. FormulaArray refuses any
'formula over 255 characters with error 1004, and the old readers pushed
'every write through it, so this length raised on every admin level change.
'@TestMethod("LLSpatial")
Public Sub TestRewriteFormulasWritesALongPlainFormula()
    CustomTestSetTitles Assert, "LLSpatial", "TestRewriteFormulasWritesALongPlainFormula"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial
    Dim longFormula As String

    Set sh = BuildAnalysisFixture("tstd")
    longFormula = "=SUM(concat_adm1_cases)&""" & String(280, "a") & """"
    sh.Range("B2").Formula = longFormula

    Set sp = LLSpatial.Create(BuildSpatialFixture())
    sp.RewriteFormulas sh.Range("B2"), "concat_adm1", "concat_adm2"

    Assert.IsTrue (Len(sh.Range("B2").Formula) > 255), _
                  "The rewritten formula should keep its full length"
    Assert.IsTrue (InStr(1, sh.Range("B2").Formula, "concat_adm2_cases") > 0), _
                  "The long formula should read the new concat column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRewriteFormulasWritesALongPlainFormula", Err.Number, Err.Description
End Sub

'@sub-title Verify ChangeAdminLevel moves a table and records the new level.
'@details
'Arranges an analysis fixture whose PREVIOUS_ADM cell holds "adm1", with one
'plain formula and one array formula reading the adm1 concat column. Acts by
'changing the level to adm4. Asserts both formulas read adm4, each kept its
'array state, and the PREVIOUS_ADM cell holds the new code.
'@TestMethod("LLSpatial")
Public Sub TestChangeAdminLevelMovesTheTable()
    CustomTestSetTitles Assert, "LLSpatial", "TestChangeAdminLevelMovesTheTable"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set sh = BuildAnalysisFixture("tste")
    sh.Range("PREVIOUS_ADM_tste").Value = "adm1"
    sh.Range("B2").Formula = "=SUM(concat_adm1_cases)"
    sh.Range("C2").FormulaArray = "=SUM(concat_adm1_cases)"

    Set sp = LLSpatial.Create(BuildSpatialFixture())
    sp.ChangeAdminLevel sh, "tste", "adm4"

    Assert.IsTrue (InStr(1, sh.Range("B2").Formula, "concat_adm4_cases") > 0), _
                  "The plain formula should read the new concat column"
    Assert.IsTrue (Not sh.Range("B2").HasArray), _
                  "The plain formula should stay plain"
    Assert.IsTrue (InStr(1, sh.Range("C2").FormulaArray, "concat_adm4_cases") > 0), _
                  "The array formula should read the new concat column"
    Assert.IsTrue sh.Range("C2").HasArray, _
                  "The array formula should stay an array one"
    Assert.AreEqual "adm4", CStr(sh.Range("PREVIOUS_ADM_tste").Value), _
                    "The new level should be recorded for the next change"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestChangeAdminLevelMovesTheTable", Err.Number, Err.Description
End Sub

'@section Spatio-temporal section migration tests
'===============================================================================

'@sub-title Build one sheet carrying the named ranges of a spatio-temporal section.
'@details
'The header row B1:D1 carries the SPT_FORMULA_COLUMN_ name. Column B is a
'named formula column: its header cell carries a LABEL name and B2 a one-row
'VALUES block, so the migration grows it to B2:B4 and rows B3 and B4 stand in
'for the Total and Missing rows. Column C matches the selector with an unnamed
'header cell. Column D carries a named header whose formula stays clear of the
'selector. F1 is the level selector and G1 the cell recording the level. The
'data formulas are written by each test.
'@param tabId String. The table identifier the names carry.
'@param selName String. The name given to the selector cell.
'@return Worksheet. The prepared section fixture sheet.
Private Function BuildSectionFixture(ByVal tabId As String, _
                                     ByVal selName As String) As Worksheet
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(ANALYSIS_SHEET, clearSheet:=True, _
                             visibility:=xlSheetHidden)

    sh.Names.Add Name:="SPT_FORMULA_COLUMN_" & tabId, RefersTo:=sh.Range("B1:D1")
    sh.Names.Add Name:="SPT_LABEL_1_" & tabId, RefersTo:=sh.Range("B1")
    sh.Names.Add Name:="SPT_VALUES_1_" & tabId, RefersTo:=sh.Range("B2")
    sh.Names.Add Name:="SPT_LABEL_3_" & tabId, RefersTo:=sh.Range("D1")
    sh.Names.Add Name:="SPT_VALUES_3_" & tabId, RefersTo:=sh.Range("D2")
    sh.Names.Add Name:=selName, RefersTo:=sh.Range("F1")

    sh.Range("B1").Formula = "=" & selName
    sh.Range("C1").Formula = "=" & selName
    sh.Range("D1").Formula = "=1"

    Set BuildSectionFixture = sh
End Function

'@sub-title Verify PreviousSectionLevel answers the recorded level.
'@details
'Arranges a section fixture whose recording cell holds 3. Acts by reading the
'level. Asserts the answer is 3.
'@TestMethod("LLSpatial")
Public Sub TestPreviousSectionLevelAnswersTheRecordedLevel()
    CustomTestSetTitles Assert, "LLSpatial", "TestPreviousSectionLevelAnswersTheRecordedLevel"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set sh = BuildSectionFixture("tstg", "SPT_SEL_tstg")
    sh.Range("G1").Value = 3

    Set sp = LLSpatial.Create(BuildSpatialFixture())

    Assert.AreEqual CLng(3), sp.PreviousSectionLevel(sh, "SPT_SEL_tstg", "tstg"), _
                    "The recorded level should be answered as it stands"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreviousSectionLevelAnswersTheRecordedLevel", Err.Number, Err.Description
End Sub

'@sub-title Verify PreviousSectionLevel refuses a cell holding text.
'@details
'Arranges a section fixture whose recording cell holds text. Acts by reading
'the level under On Error Resume Next and captures the error number. Asserts
'the number is InvalidArgument. The old reader ran CLng on the raw value, so a
'text cell raised a bare type mismatch into the caller's silence.
'@TestMethod("LLSpatial")
Public Sub TestPreviousSectionLevelRefusesText()
    CustomTestSetTitles Assert, "LLSpatial", "TestPreviousSectionLevelRefusesText"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial
    Dim errNumber As Long
    Dim errDescription As String

    Set sh = BuildSectionFixture("tsth", "SPT_SEL_tsth")
    sh.Range("G1").Value = "cleared"

    Set sp = LLSpatial.Create(BuildSpatialFixture())

    On Error Resume Next
    sp.PreviousSectionLevel sh, "SPT_SEL_tsth", "tsth"
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A text level should raise InvalidArgument - " & _
                    "description was [" & errDescription & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreviousSectionLevelRefusesText", Err.Number, Err.Description
End Sub

'@sub-title Verify PreviousSectionLevel refuses a level outside 1 to 4.
'@details
'Arranges a section fixture and reads the level twice, once over 0 and once
'over 5, each under On Error Resume Next. Asserts both reads raise
'InvalidArgument. A blank cell reads as 0, matches no concat column, and a
'rewrite on it would record the new level while migrating nothing.
'@TestMethod("LLSpatial")
Public Sub TestPreviousSectionLevelRefusesALevelOutsideTheRange()
    CustomTestSetTitles Assert, "LLSpatial", "TestPreviousSectionLevelRefusesALevelOutsideTheRange"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial
    Dim errBelow As Long
    Dim errAbove As Long

    Set sh = BuildSectionFixture("tsti", "SPT_SEL_tsti")
    Set sp = LLSpatial.Create(BuildSpatialFixture())

    sh.Range("G1").Value = 0
    On Error Resume Next
    sp.PreviousSectionLevel sh, "SPT_SEL_tsti", "tsti"
    errBelow = Err.Number
    Err.Clear
    On Error GoTo TestFail

    sh.Range("G1").Value = 5
    On Error Resume Next
    sp.PreviousSectionLevel sh, "SPT_SEL_tsti", "tsti"
    errAbove = Err.Number
    Err.Clear
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errBelow, _
                    "A level of 0 should raise InvalidArgument"
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errAbove, _
                    "A level of 5 should raise InvalidArgument"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreviousSectionLevelRefusesALevelOutsideTheRange", Err.Number, Err.Description
End Sub

'@sub-title Verify MigrateSection moves the named column and its two footer rows.
'@details
'Arranges the section fixture with formulas on the adm1 concat column in the
'VALUES cell, the two rows under it, and a third row past the growth. Acts by
'migrating from level 1 to level 2. Asserts the VALUES cell and both footer
'rows read adm2, the row past the growth still reads adm1, and the recording
'cell holds the new level.
'@TestMethod("LLSpatial")
Public Sub TestMigrateSectionMovesTheColumnAndItsFooterRows()
    CustomTestSetTitles Assert, "LLSpatial", "TestMigrateSectionMovesTheColumnAndItsFooterRows"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set sh = BuildSectionFixture("tstj", "SPT_SEL_tstj")
    sh.Range("G1").Value = 1
    sh.Range("B2").Formula = "=SUM(concat_adm1_cases)"
    sh.Range("B3").Formula = "=SUM(concat_adm1_total)"
    sh.Range("B4").Formula = "=SUM(concat_adm1_missing)"
    sh.Range("B5").Formula = "=SUM(concat_adm1_below)"

    Set sp = LLSpatial.Create(BuildSpatialFixture())
    sp.MigrateSection sh, "SPT_SEL_tstj", "tstj", 1, 2

    Assert.IsTrue (InStr(1, sh.Range("B2").Formula, "concat_adm2_cases") > 0), _
                  "The VALUES cell should read the new concat column"
    Assert.IsTrue (InStr(1, sh.Range("B3").Formula, "concat_adm2_total") > 0), _
                  "The first footer row should move with the block"
    Assert.IsTrue (InStr(1, sh.Range("B4").Formula, "concat_adm2_missing") > 0), _
                  "The second footer row should move with the block"
    Assert.IsTrue (InStr(1, sh.Range("B5").Formula, "concat_adm1_below") > 0), _
                  "The growth should stop two rows under the block"
    Assert.AreEqual CLng(2), CLng(sh.Range("G1").Value), _
                    "The new level should be recorded beside the selector"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestMigrateSectionMovesTheColumnAndItsFooterRows", Err.Number, Err.Description
End Sub

'@sub-title Verify MigrateSection leaves a column with an unnamed header alone.
'@details
'Arranges the fixture with formulas under the named column and under the
'column whose header cell has no name. Acts by migrating. Asserts the named
'column moved and the unnamed one is untouched. A header cell has a name only
'when something named it, and a stale name from the previous column used to
'send the rewrite into the previous column's block.
'@TestMethod("LLSpatial")
Public Sub TestMigrateSectionLeavesAnUnnamedColumnAlone()
    CustomTestSetTitles Assert, "LLSpatial", "TestMigrateSectionLeavesAnUnnamedColumnAlone"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set sh = BuildSectionFixture("tstk", "SPT_SEL_tstk")
    sh.Range("G1").Value = 1
    sh.Range("B2").Formula = "=SUM(concat_adm1_cases)"
    sh.Range("C2").Formula = "=SUM(concat_adm1_cases)"

    Set sp = LLSpatial.Create(BuildSpatialFixture())
    sp.MigrateSection sh, "SPT_SEL_tstk", "tstk", 1, 4

    Assert.IsTrue (InStr(1, sh.Range("B2").Formula, "concat_adm4_cases") > 0), _
                  "The named column should read the new concat column"
    Assert.IsTrue (InStr(1, sh.Range("C2").Formula, "concat_adm1_cases") > 0), _
                  "The unnamed column should be untouched"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestMigrateSectionLeavesAnUnnamedColumnAlone", Err.Number, Err.Description
End Sub

'@sub-title Verify MigrateSection leaves the column of another selector alone.
'@details
'Arranges the fixture with a formula under the named column whose header
'formula stays clear of the selector. Acts by migrating. Asserts that column
'is untouched: the header match is what decides which columns belong to the
'selector that fired.
'@TestMethod("LLSpatial")
Public Sub TestMigrateSectionLeavesAnotherSelectorAlone()
    CustomTestSetTitles Assert, "LLSpatial", "TestMigrateSectionLeavesAnotherSelectorAlone"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set sh = BuildSectionFixture("tstl", "SPT_SEL_tstl")
    sh.Range("G1").Value = 1
    sh.Range("D2").Formula = "=SUM(concat_adm1_cases)"

    Set sp = LLSpatial.Create(BuildSpatialFixture())
    sp.MigrateSection sh, "SPT_SEL_tstl", "tstl", 1, 2

    Assert.IsTrue (InStr(1, sh.Range("D2").Formula, "concat_adm1_cases") > 0), _
                  "A column clear of the selector should be untouched"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestMigrateSectionLeavesAnotherSelectorAlone", Err.Number, Err.Description
End Sub

'@sub-title Verify ApplyPopulationFactor wraps the formula and revertBack unwraps it.
'@details
'Arranges a data cell right of the row categories, holding a plain formula on
'the adm1 concat column, a factor of 50 and no previous factor. Acts by
'applying the division, then reverting it. Asserts the applied formula holds
'the factor and the population cell, the previous factor cell tracks each
'step, and the revert restores a formula free of both.
'@TestMethod("LLSpatial")
Public Sub TestApplyPopulationFactorDividesThenReverts()
    CustomTestSetTitles Assert, "LLSpatial", "TestApplyPopulationFactorDividesThenReverts"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set sh = BuildAnalysisFixture("tstf")
    sh.Range("C2").Formula = "=SUM(concat_adm1_cases)"
    sh.Range("POPFACT_tstf").Value = 50
    sh.Range("POPPREVFACT_tstf").Value = 0

    Set sp = LLSpatial.Create(BuildSpatialFixture())
    sp.ApplyPopulationFactor sh, "tstf", "adm1"

    Assert.IsTrue (InStr(1, sh.Range("C2").Formula, "50*") > 0), _
                  "The applied formula should carry the factor"
    Assert.IsTrue (InStr(1, sh.Range("C2").Formula, "/$A$2") > 0), _
                  "The applied formula should divide by the population cell"
    Assert.AreEqual 50, CLng(sh.Range("POPPREVFACT_tstf").Value), _
                    "The applied factor should be recorded"

    sp.ApplyPopulationFactor sh, "tstf", "adm1", revertBack:=True

    Assert.IsTrue (InStr(1, sh.Range("C2").Formula, "50*") = 0), _
                  "The reverted formula should hold no factor"
    Assert.IsTrue (InStr(1, sh.Range("C2").Formula, "/$A$2") = 0), _
                  "The reverted formula should hold no population cell"
    Assert.AreEqual 0, CLng(sh.Range("POPPREVFACT_tstf").Value), _
                    "The revert should clear the recorded factor"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestApplyPopulationFactorDividesThenReverts", Err.Number, Err.Description
End Sub

'@sub-title Verify FormatPopulationFactor shows and hides the factor cells.
'@details
'Acts by hiding the factor cells, then showing them. Asserts the factor cell
'is white and locked while hidden, and black and editable while shown, with
'the label cell following.
'@TestMethod("LLSpatial")
Public Sub TestFormatPopulationFactorTogglesTheCells()
    CustomTestSetTitles Assert, "LLSpatial", "TestFormatPopulationFactorTogglesTheCells"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim sp As LLSpatial

    Set sh = BuildAnalysisFixture("tstg")
    Set sp = LLSpatial.Create(BuildSpatialFixture())

    sp.FormatPopulationFactor sh, "tstg", factorVisible:=False

    Assert.IsTrue (sh.Range("POPFACT_tstg").Font.Color = vbWhite), _
                  "A hidden factor cell should be white"
    Assert.IsTrue sh.Range("POPFACT_tstg").Locked, _
                  "A hidden factor cell should be locked"
    Assert.IsTrue sh.Range("POPFACTLABEL_tstg").Locked, _
                  "A hidden factor label should be locked"

    sp.FormatPopulationFactor sh, "tstg", factorVisible:=True

    Assert.IsTrue (sh.Range("POPFACT_tstg").Font.Color = vbBlack), _
                  "A shown factor cell should be black"
    Assert.IsTrue (Not sh.Range("POPFACT_tstg").Locked), _
                  "A shown factor cell should be editable"
    Assert.IsTrue (Not sh.Range("POPFACTLABEL_tstg").Locked), _
                  "A shown factor label should be editable"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFormatPopulationFactorTogglesTheCells", Err.Number, Err.Description
End Sub

'@section TopGeoValue / TopHFValue tests
'===============================================================================

'@sub-title Verify TopGeoValue returns empty when the spatial ListObject does not exist.
'@details
'Arranges a spatial fixture with listofgeovars but no admin-level spatial
'ListObjects. Acts by creating an LLSpatial instance and calling
'TopGeoValue("adm1", 1, "cases", "sp1"). Asserts that the result is
'vbNullString.
'@TestMethod("LLSpatial")
Public Sub TestTopGeoValueReturnsEmptyForMissingTable()
    CustomTestSetTitles Assert, "LLSpatial", "TestTopGeoValueReturnsEmptyForMissingTable"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture()

    Dim sp As LLSpatial
    Set sp = LLSpatial.Create(sh)

    Dim result As String
    result = sp.TopGeoValue("adm1", 1, "cases", "sp1")

    Assert.AreEqual vbNullString, result, _
                    "TopGeoValue should return empty when spatial table does not exist"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTopGeoValueReturnsEmptyForMissingTable", Err.Number, Err.Description
End Sub

'@sub-title Verify TopHFValue returns empty when the spatial ListObject does not exist.
'@details
'Arranges a spatial fixture with listofgeovars but no health facility spatial
'ListObjects. Acts by creating an LLSpatial instance and calling
'TopHFValue(1, "cases", "sp1"). Asserts that the result is vbNullString.
'@TestMethod("LLSpatial")
Public Sub TestTopHFValueReturnsEmptyForMissingTable()
    CustomTestSetTitles Assert, "LLSpatial", "TestTopHFValueReturnsEmptyForMissingTable"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture()

    Dim sp As LLSpatial
    Set sp = LLSpatial.Create(sh)

    Dim result As String
    result = sp.TopHFValue(1, "cases", "sp1")

    Assert.AreEqual vbNullString, result, _
                    "TopHFValue should return empty when spatial table does not exist"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTopHFValueReturnsEmptyForMissingTable", Err.Number, Err.Description
End Sub

'@sub-title Verify the rank properties refuse a rank below one.
'@details
'Arranges a workbook holding a filled administrative table and a facility
'table. Acts by asking both rank properties for rank 0. Asserts both answer an
'empty string. Only the upper bound was tested, so rank 0 read the header cell
'and =TopAdmin("adm1", 0, ...) in a cell answered with the internal table name.
'@TestMethod("LLSpatial")
Public Sub TestRankPropertiesRefuseARankBelowOne()
    CustomTestSetTitles Assert, "LLSpatial", "TestRankPropertiesRefuseARankBelowOne"
    On Error GoTo TestFail

    Dim wb As Workbook
    Dim sh As Worksheet
    Dim filtSh As Worksheet
    Dim sp As LLSpatial

    Set wb = BuildSpatialWorkbook(withRegistry:=True)
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    Set filtSh = AddFilteredSheet(wb)
    AddHListSheet wb, Array("adm1_cases")
    RegisterSpatialVar sh, "cases_sp1"
    AddSpatialTables sh, "cases_sp1", filtSh
    AddHFTable sh, "cases_sp1"

    Set sp = LLSpatial.Create(sh)
    sp.Update

    Assert.AreEqual vbNullString, sp.TopGeoValue("adm1", 0, "cases", "sp1"), _
                    "Rank 0 should answer nothing rather than the header cell"
    Assert.AreEqual vbNullString, sp.TopHFValue(0, "cases", "sp1"), _
                    "Rank 0 of a facility table should answer nothing"
    Assert.AreEqual "P1", sp.TopGeoValue("adm1", 1, "cases", "sp1"), _
                    "Rank 1 should answer the top ranked admin unit"
    Assert.AreEqual "Clinic A", sp.TopHFValue(1, "cases", "sp1"), _
                    "Rank 1 of a facility table should answer its facility"

    DeleteWorkbook wb

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRankPropertiesRefuseARankBelowOne", Err.Number, Err.Description
End Sub

'@sub-title Verify the rank properties say when the table identifier is missing.
'@details
'Arranges a spatial fixture and asks for a rank with no table identifier. A
'spatial table name always carries the identifier of its analysis table, so an
'empty one builds a name nothing can match. Asserts the answer names the
'reason, which reaches the cell of the worksheet function. It used to answer an
'empty string, which reads on the sheet like a table with no data in it.
'@TestMethod("LLSpatial")
Public Sub TestRankPropertiesSayWhenTheTableIdIsMissing()
    CustomTestSetTitles Assert, "LLSpatial", "TestRankPropertiesSayWhenTheTableIdIsMissing"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = BuildSpatialFixture()

    Dim sp As LLSpatial
    Set sp = LLSpatial.Create(sh)

    Assert.AreEqual "#missing table id", sp.TopGeoValue("adm1", 1, "cases"), _
                    "A rank asked for with no table identifier should say so"
    Assert.AreEqual "#missing table id", sp.TopHFValue(1, "cases"), _
                    "A facility rank asked for with no table identifier should say so"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRankPropertiesSayWhenTheTableIdIsMissing", Err.Number, Err.Description
End Sub
