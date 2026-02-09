Attribute VB_Name = "TestCrossTableFormula"
Attribute VB_Description = "Tests for CrossTableFormula class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for CrossTableFormula class")

' CrossTableFormula tests focus on factory validation and the Valid property.
' Full formula writing tests require a complete linelist workbook with
' dictionary tables and formula data sheets, and are exercised through
' integration tests in TestAnalysisOutput.

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "CTFormulaFixture"
Private Const OUTPUT_SHEET As String = "CTFormulaOutput"

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

Private Function BuildCrossTable(ByVal specs As ITableSpecs) As ICrossTable
    Dim sh As Worksheet
    Set sh = OutputSheet()
    Dim ct As ICrossTable
    Set ct = CrossTable.Create(specs, sh, transStub)
    ct.Build
    Set BuildCrossTable = ct
End Function

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCrossTableFormula"
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

'@TestMethod("CrossTableFormula")
Public Sub TestCreateRejectsNothingTable()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestCreateRejectsNothingTable"
    On Error GoTo TestFail

    ' FormulaData requires a worksheet with formula lookup data.
    ' For factory validation we just need to pass Nothing for tabl.
    Dim sh As Worksheet
    Set sh = OutputSheet()
    Dim fData As IFormulaData
    Set fData = FormulaData.Create(sh)

    On Error Resume Next
    Dim ctf As ICrossTableFormula
    Set ctf = CrossTableFormula.Create(Nothing, fData)
    On Error GoTo 0

    Assert.IsTrue (ctf Is Nothing), _
                  "Create with Nothing cross-table should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingTable", Err.Number, Err.Description
End Sub

'@TestMethod("CrossTableFormula")
Public Sub TestCreateRejectsNothingFormulaData()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestCreateRejectsNothingFormulaData"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim ct As ICrossTable
    Set ct = BuildCrossTable(specs)

    On Error Resume Next
    Dim ctf As ICrossTableFormula
    Set ctf = CrossTableFormula.Create(ct, Nothing)
    On Error GoTo 0

    Assert.IsTrue (ctf Is Nothing), _
                  "Create with Nothing formula data should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingFormulaData", Err.Number, Err.Description
End Sub

'@TestMethod("CrossTableFormula")
Public Sub TestCreateReturnsValidObject()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestCreateReturnsValidObject"
    On Error GoTo TestFail

    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "N", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim ct As ICrossTable
    Set ct = BuildCrossTable(specs)

    Dim sh As Worksheet
    Set sh = OutputSheet()
    Dim fData As IFormulaData
    Set fData = FormulaData.Create(sh)

    Dim ctf As ICrossTableFormula
    Set ctf = CrossTableFormula.Create(ct, fData)

    Assert.IsTrue (Not ctf Is Nothing), _
                  "Create with valid params should succeed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsValidObject", Err.Number, Err.Description
End Sub

'@section Valid property tests
'===============================================================================

'@TestMethod("CrossTableFormula")
Public Sub TestValidReturnsFalseForInvalidFormula()
    CustomTestSetTitles Assert, "CrossTableFormula", "TestValidReturnsFalseForInvalidFormula"
    On Error GoTo TestFail

    ' "InvalidFunc" is not a recognized summary function, so Valid should be False
    BuildFixture "univariate analysis", Array( _
        Array("S1", "row_var", "", "yes", "no", "no", "no", "Count", "InvalidFunc", ""))
    Dim specs As ITableSpecs
    Set specs = CreateSpecs(1)

    Dim ct As ICrossTable
    Set ct = BuildCrossTable(specs)

    Dim sh As Worksheet
    Set sh = OutputSheet()
    Dim fData As IFormulaData
    Set fData = FormulaData.Create(sh)

    Dim ctf As ICrossTableFormula
    Set ctf = CrossTableFormula.Create(ct, fData)

    Assert.IsFalse ctf.Valid, _
                   "Valid should be False for unrecognized summary function"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValidReturnsFalseForInvalidFormula", Err.Number, Err.Description
End Sub
