Attribute VB_Name = "TestTableSpecsNavigator"
Option Explicit
Option Private Module

'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Unit tests for the TableSpecsNavigator helper")

Private Const NAV_SHEET As String = "TableSpecsNavigator"

Private Assert As Object
Private Builder As TableSpecsNavigatorBuilderStub
Private LinelistStub As TableSpecsLinelistStub
Private HeaderRange As Range
Private CurrentRange As Range
Private ColumnMap As ITableSpecsColumnMap
Private Navigator As ITableSpecsNavigator

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet NAV_SHEET
    Set Navigator = Nothing
    Set ColumnMap = Nothing
    Set CurrentRange = Nothing
    Set HeaderRange = Nothing
    Set LinelistStub = Nothing
    Set Builder = Nothing
    Set Assert = Nothing
End Sub

'@section Test lifecycle
'===============================================================================

'@TestInitialize
Private Sub TestInitialize()
    SeedWorksheet
    Set Builder = New TableSpecsNavigatorBuilderStub
    Set LinelistStub = New TableSpecsLinelistStub
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Navigator = Nothing
    Set ColumnMap = Nothing
    Set CurrentRange = Nothing
    Set HeaderRange = Nothing
    Set Builder = Nothing
    Set LinelistStub = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("TableSpecsNavigator")
Private Sub TestIsNewSectionDetectsContinuation()
    On Error GoTo Fail

    PrepareNavigator rowNumber:=3, tableType:=TABLE_TYPE_UNIVARIATE
    Builder.SetRowValidity 2, True
    Builder.SetRowValidity 3, True

    Assert.IsFalse Navigator.IsNewSection, "Row three should continue the existing section"

    PrepareNavigator rowNumber:=4, tableType:=TABLE_TYPE_UNIVARIATE
    Assert.IsTrue Navigator.IsNewSection, "Row four should start a new section"

    PrepareNavigator rowNumber:=4, tableType:=TABLE_TYPE_GLOBAL_SUMMARY
    Assert.IsFalse Navigator.IsNewSection, "Global summary tables should never mark new sections"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestIsNewSectionDetectsContinuation"
End Sub

'@TestMethod("TableSpecsNavigator")
Private Sub TestPreviousSkipsInvalidRows()
    Dim previousSpec As ITablesSpecs
    Dim stub As TableSpecsNavigatorSpecStub

    On Error GoTo Fail

    Builder.SetRowValidity 2, True
    Builder.SetRowValidity 3, False
    PrepareNavigator rowNumber:=4, tableType:=TABLE_TYPE_UNIVARIATE

    On Error Resume Next
        Set previousSpec = Navigator.PreviousSpec
    On Error GoTo Fail

    Assert.IsNotNothing previousSpec, "Navigator should return a previous specification"
    Assert.AreEqual "TableSpecsNavigatorSpecStub", TypeName(previousSpec), "Returned spec should match stub type"

    Set stub = previousSpec
    Assert.AreEqual 2&, stub.RowNumber, "Navigator should skip invalid rows when searching backwards"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestPreviousSkipsInvalidRows"
End Sub

'@TestMethod("TableSpecsNavigator")
Private Sub TestPreviousRaisesOnNewSection()
    On Error GoTo Fail

    Builder.SetRowValidity 2, True
    Builder.SetRowValidity 3, True
    PrepareNavigator rowNumber:=4, tableType:=TABLE_TYPE_UNIVARIATE

    On Error Resume Next
        Navigator.PreviousSpec
        Dim capturedError As Long
        capturedError = Err.Number
        Err.Clear
    On Error GoTo Fail

    Assert.AreEqual CLng(ProjectError.ElementNotFound), capturedError, _
        "Expected navigator to raise ElementNotFound for new section"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestPreviousRaisesOnNewSection"
End Sub

'@TestMethod("TableSpecsNavigator")
Private Sub TestNextSpecReturnsNextValid()
    Dim nextSpec As ITablesSpecs
    Dim stub As TableSpecsNavigatorSpecStub
    Dim anchorRange As Range

    On Error GoTo Fail

    Builder.SetRowValidity 3, False
    Builder.SetRowValidity 4, True

    PrepareNavigator rowNumber:=2, tableType:=TABLE_TYPE_UNIVARIATE
    Set anchorRange = WorksheetForTests.Range("A5:C5")

    Set nextSpec = Navigator.NextSpec(anchorRange)
    Assert.IsNotNothing nextSpec, "Navigator should return the next valid spec"
    Assert.AreEqual "TableSpecsNavigatorSpecStub", TypeName(nextSpec)

    Set stub = nextSpec
    Assert.AreEqual 4&, stub.RowNumber, "Navigator should skip invalid rows when searching forwards"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestNextSpecReturnsNextValid"
End Sub

'@section Helpers
'===============================================================================

Private Sub SeedWorksheet()
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(NAV_SHEET)

    sh.Cells.Clear
    sh.Range("A1").Value = "section"
    sh.Range("B1").Value = "row"
    sh.Range("C1").Value = "column"

    sh.Range("A2").Value = "A"
    sh.Range("B2").Value = "row2"
    sh.Range("C2").Value = "col2"

    sh.Range("A3").Value = "A"
    sh.Range("B3").Value = "row3"
    sh.Range("C3").Value = "col3"

    sh.Range("A4").Value = "B"
    sh.Range("B4").Value = "row4"
    sh.Range("C4").Value = "col4"

    sh.Range("A5").Value = "B"
    sh.Range("B5").Value = "row5"
    sh.Range("C5").Value = "col5"
End Sub

Private Sub PrepareNavigator(ByVal rowNumber As Long, ByVal tableType As Byte)
    Dim sh As Worksheet

    Set sh = WorksheetForTests
    Set HeaderRange = sh.Range("A1:C1")
    Set CurrentRange = sh.Range("A" & CStr(rowNumber) & ":C" & CStr(rowNumber))
    Set ColumnMap = TableSpecsColumnMap.Create(HeaderRange, CurrentRange)
    Set Navigator = TableSpecsNavigator.Create(HeaderRange, CurrentRange, LinelistStub, ColumnMap, Builder, tableType)
End Sub

Private Property Get WorksheetForTests() As Worksheet
    Set WorksheetForTests = ThisWorkbook.Worksheets(NAV_SHEET)
End Property

