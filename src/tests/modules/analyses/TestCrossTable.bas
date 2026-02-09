Attribute VB_Name = "TestCrossTable"
Attribute VB_Description = "Tests for CrossTable class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for CrossTable class")

' CrossTable tests focus on factory validation and basic Build behavior.
' Full integration tests with named ranges and formatting are exercised
' through TestAnalysisOutput. These tests verify:
' - Factory rejects Nothing parameters
' - Build creates named ranges for simple table types
' - Position properties are set after Build

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "CrossTableFixture"
Private Const OUTPUT_SHEET As String = "CrossTableOutput"

' Column layout for the fixture (must match TableSpecs header order)
Private Const COL_SECTION As Long = 1
Private Const COL_ROW As Long = 2
Private Const COL_COLUMN As Long = 3
Private Const COL_TOTAL As Long = 4
Private Const COL_PERCENTAGE As Long = 5
Private Const COL_MISSING As Long = 6
Private Const COL_GRAPH As Long = 7
Private Const COL_LABEL As Long = 8
Private Const COL_FUNCTION As Long = 9
Private Const COL_NGEO As Long = 10
Private Const NUM_COLUMNS As Long = 10

Private Assert As ICustomTest
Private lDataStub As TableSpecsLinelistStub
Private transStub As LinelistSpecsTranslationStub

'@section Helpers
'===============================================================================

Private Sub BuildFixture(ByVal tableTypeName As String, dataRows As Variant)
    Dim sh As Worksheet
    Dim headerArray As Variant
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant

    Set sh = EnsureWorksheet(FIXTURE_SHEET, clearSheet:=True, visibility:=xlSheetHidden)

    sh.Cells(1, 1).Value = tableTypeName
    headerArray = Array( _
        Array("section", "row", "column", "total", "percentage", _
              "missing", "graph", "label", "function", "n geo"))
    headerMatrix = RowsToMatrix(headerArray)
    WriteMatrix sh.Cells(3, 1), headerMatrix

    If Not IsEmpty(dataRows) Then
        dataMatrix = RowsToMatrix(dataRows)
        WriteMatrix sh.Cells(4, 1), dataMatrix
    End If
End Sub

Private Function FixtureHeaderRange() As Range
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    Set FixtureHeaderRange = sh.Range(sh.Cells(3, 1), sh.Cells(3, NUM_COLUMNS))
End Function

Private Function FixtureDataRange(ByVal dataRowIndex As Long) As Range
    Dim sh As Worksheet
    Dim rowNum As Long
    Set sh = ThisWorkbook.Worksheets(FIXTURE_SHEET)
    rowNum = 3 + dataRowIndex
    Set FixtureDataRange = sh.Range(sh.Cells(rowNum, 1), sh.Cells(rowNum, NUM_COLUMNS))
End Function

Private Function CreateSpecs(ByVal dataRowIndex As Long) As ITableSpecs
    Set CreateSpecs = TableSpecs.Create( _
        FixtureHeaderRange(), _
        FixtureDataRange(dataRowIndex), _
        lDataStub)
End Function

Private Function OutputSheet() As Worksheet
    Set OutputSheet = EnsureWorksheet(OUTPUT_SHEET, clearSheet:=True, visibility:=xlSheetHidden)
End Function

Private Function RangeExistsOnSheet(ByVal sh As Worksheet, ByVal rngName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExistsOnSheet = (Not rng Is Nothing)
End Function

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCrossTable"
    Set lDataStub = New TableSpecsLinelistStub
    Set transStub = New LinelistSpecsTranslationStub
    transStub.Initialise "TestTrans"
    transStub.SetTranslation "MSG_NA", "Missing"
    transStub.SetTranslation "MSG_Total", "Total"
    transStub.SetTranslation "MSG_Percent", "%"
    transStub.SetTranslation "MSG_AllData", "All Data"
    transStub.SetTranslation "MSG_FilteredData", "Filtered Data"
    transStub.SetTranslation "MSG_GlobalSummary", "Global Summary"
    transStub.SetTranslation "MSG_Period", "Period"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet FIXTURE_SHEET
    DeleteWorksheet OUTPUT_SHEET
    RestoreApp
    Set lDataStub = Nothing
    Set transStub = Nothing
    Set Assert = Nothing
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

'@section Factory validation tests
'===============================================================================

'@TestMethod("CrossTable")
Public Sub TestCreateRejectsNothingSpecs()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateRejectsNothingSpecs"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = OutputSheet()

    On Error Resume Next
    Dim ct As ICrossTable
    Set ct = CrossTable.Create(Nothing, sh, transStub)
    On Error GoTo 0

    Assert.IsTrue (ct Is Nothing), _
                  "Create with Nothing specs should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingSpecs", Err.Number, Err.Description
End Sub

'@TestMethod("CrossTable")
Public Sub TestCreateRejectsNothingWorksheet()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateRejectsNothingWorksheet"
    On Error GoTo TestFail

    ' Build a valid specs fixture
    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    On Error Resume Next
    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, Nothing, transStub)
    On Error GoTo 0

    Assert.IsTrue (ct Is Nothing), _
                  "Create with Nothing worksheet should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingWorksheet", Err.Number, Err.Description
End Sub

'@TestMethod("CrossTable")
Public Sub TestCreateRejectsNothingTranslation()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateRejectsNothingTranslation"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    On Error Resume Next
    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, Nothing)
    On Error GoTo 0

    Assert.IsTrue (ct Is Nothing), _
                  "Create with Nothing translations should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingTranslation", Err.Number, Err.Description
End Sub

'@TestMethod("CrossTable")
Public Sub TestCreateReturnsValidObject()
    CustomTestSetTitles Assert, "CrossTable", "TestCreateReturnsValidObject"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)

    Assert.IsTrue (Not ct Is Nothing), _
                  "Create with valid params should succeed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsValidObject", Err.Number, Err.Description
End Sub

'@section Property tests
'===============================================================================

'@TestMethod("CrossTable")
Public Sub TestSpecificationsProperty()
    CustomTestSetTitles Assert, "CrossTable", "TestSpecificationsProperty"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)

    Assert.AreEqual specs.TableId, ct.Specifications.TableId, _
                    "Specifications should return the specs passed at creation"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSpecificationsProperty", Err.Number, Err.Description
End Sub

'@TestMethod("CrossTable")
Public Sub TestWkshProperty()
    CustomTestSetTitles Assert, "CrossTable", "TestWkshProperty"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)

    Assert.AreEqual sh.Name, ct.Wksh.Name, _
                    "Wksh should return the output worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWkshProperty", Err.Number, Err.Description
End Sub

'@section Build tests — GlobalSummary
'===============================================================================

'@TestMethod("CrossTable")
Public Sub TestBuildGlobalSummaryCreatesRowGsSet()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildGlobalSummaryCreatesRowGsSet"
    On Error GoTo TestFail

    BuildFixture "Add or remove rows of Global Summary", Array( _
        Array("S1", "", "", "", "", "", "", "Total Cases", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)
    ct.Build

    Assert.IsTrue RangeExistsOnSheet(sh, "ROWGS_SET"), _
                  "Build should create ROWGS_SET named range for global summary"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildGlobalSummaryCreatesRowGsSet", Err.Number, Err.Description
End Sub

'@TestMethod("CrossTable")
Public Sub TestBuildGlobalSummaryCreatesColGsSet()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildGlobalSummaryCreatesColGsSet"
    On Error GoTo TestFail

    BuildFixture "Add or remove rows of Global Summary", Array( _
        Array("S1", "", "", "", "", "", "", "Total Cases", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)
    ct.Build

    Assert.IsTrue RangeExistsOnSheet(sh, "COLGS_SET"), _
                  "Build should create COLGS_SET column header for global summary"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildGlobalSummaryCreatesColGsSet", Err.Number, Err.Description
End Sub

'@section Build tests — Univariate
'===============================================================================

'@TestMethod("CrossTable")
Public Sub TestBuildUnivariateCreatesNamedRanges()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildUnivariateCreatesNamedRanges"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)
    ct.Build

    Dim tabId As String
    tabId = specs.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "ENDTABLE_" & tabId), _
                  "Build should create ENDTABLE_ named range"
    Assert.IsTrue RangeExistsOnSheet(sh, "ROW_CATEGORIES_" & tabId), _
                  "Build should create ROW_CATEGORIES_ named range"
    Assert.IsTrue RangeExistsOnSheet(sh, "VALUES_COL_1_" & tabId), _
                  "Build should create VALUES_COL_1_ named range"
    Assert.IsTrue RangeExistsOnSheet(sh, "INTERIOR_VALUES_" & tabId), _
                  "Build should create INTERIOR_VALUES_ named range"
    Assert.IsTrue RangeExistsOnSheet(sh, "TITLE_" & tabId), _
                  "Build should create TITLE_ named range"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildUnivariateCreatesNamedRanges", Err.Number, Err.Description
End Sub

'@TestMethod("CrossTable")
Public Sub TestBuildUnivariateEndRowSet()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildUnivariateEndRowSet"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)
    ct.Build

    Assert.IsTrue ct.EndRow > 0, _
                  "EndRow should be set after Build"
    Assert.IsTrue ct.EndColumn > 0, _
                  "EndColumn should be set after Build"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildUnivariateEndRowSet", Err.Number, Err.Description
End Sub

'@TestMethod("CrossTable")
Public Sub TestBuildUnivariateNumberOfColumns()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildUnivariateNumberOfColumns"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)
    ct.Build

    Assert.AreEqual 1&, ct.NumberOfColumns, _
                    "Univariate table should have 1 data column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildUnivariateNumberOfColumns", Err.Number, Err.Description
End Sub

'@section Build tests — Univariate with section
'===============================================================================

'@TestMethod("CrossTable")
Public Sub TestBuildUnivariateNewSectionCreatesSection()
    CustomTestSetTitles Assert, "CrossTable", "TestBuildUnivariateNewSectionCreatesSection"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)
    ct.Build

    Dim tabId As String
    tabId = specs.TableId

    Assert.IsTrue RangeExistsOnSheet(sh, "SECTION_" & tabId), _
                  "Build should create SECTION_ named range for new section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildUnivariateNewSectionCreatesSection", Err.Number, Err.Description
End Sub

'@section NamedRangesList tests
'===============================================================================

'@TestMethod("CrossTable")
Public Sub TestNamedRangesListPopulatedAfterBuild()
    CustomTestSetTitles Assert, "CrossTable", "TestNamedRangesListPopulatedAfterBuild"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim sh As Worksheet
    Set sh = OutputSheet()

    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)
    ct.Build

    Dim rangeNames As BetterArray
    Set rangeNames = ct.NamedRangesList

    Assert.IsTrue rangeNames.Length > 0, _
                  "NamedRangesList should contain entries after Build"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNamedRangesListPopulatedAfterBuild", Err.Number, Err.Description
End Sub
