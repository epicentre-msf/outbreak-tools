Attribute VB_Name = "TestDiseaseComparisonReport"
Attribute VB_Description = "Tests covering the writer that renders a disease comparison on __compRep"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests covering the writer that renders a disease comparison on __compRep")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const REPORT_SHEET_NAME As String = "__compRep"
Private Const MISSING_SHEET_NAME As String = "__compRepMissing"
Private Const FIRST_SHEET_NAME As String = "CompReportFirst"
Private Const SECOND_SHEET_NAME As String = "CompReportSecond"
Private Const FIRST_TABLE_NAME As String = "T_CompReportFirst"
Private Const SECOND_TABLE_NAME As String = "T_CompReportSecond"

'CheckingOutput writes the parent title of every row in column B and the three
'visible pieces of an entry in columns C, D and E.
Private Const HIDDEN_TITLE_COLUMN_INDEX As Long = 2
Private Const FIRST_VISIBLE_COLUMN_INDEX As Long = 3

Private Assert As CustomTest
Private ReportSheet As Worksheet
Private FirstSheet As Worksheet
Private SecondSheet As Worksheet
Private FirstTable As ListObject
Private SecondTable As ListObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseComparisonReport"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        DeleteWorksheets REPORT_SHEET_NAME, FIRST_SHEET_NAME, SECOND_SHEET_NAME
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    'A fresh report sheet each time: ClearWorksheet drops the sheet names
    'CheckingOutput keeps, so every test starts from an untouched sheet.
    Set ReportSheet = EnsureWorksheet(REPORT_SHEET_NAME)
    PrepareFirstTable
    PrepareSecondTable
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'The assertions of a test reach the results sheet only once flushed.
    Assert.Flush
    If Not ReportSheet Is Nothing Then ClearWorksheet ReportSheet
    If Not FirstSheet Is Nothing Then ClearWorksheet FirstSheet
    If Not SecondSheet Is Nothing Then ClearWorksheet SecondSheet
    Set FirstTable = Nothing
    Set SecondTable = Nothing
End Sub

'@section Tests
'===============================================================================

'The fixture pair:
'  var_a  shared, same choices spelled in another order and case
'  var_b  only in the first disease
'  var_c  only in the second disease
'  var_d  shared, choices differ ("no" only in 1, "maybe" only in 2)
'  var_e  shared, Choice cell empty on both sides

'@TestMethod("DiseaseComparisonReport")
Public Sub TestPrintReportWritesTheFiveSections()
    CustomTestSetTitles Assert, "DiseaseComparisonReport", "TestPrintReportWritesTheFiveSections"

    Dim writer As DiseaseComparisonReport

    On Error GoTo Fail

    Set writer = DiseaseComparisonReport.Create(ThisWorkbook)
    writer.PrintComparison FirstTable, SecondTable, "Cholera", "Measles"

    Assert.IsTrue TitleOnSheet("Variables only in Cholera"), "The report opens with the variables only in the first disease"
    Assert.IsTrue TitleOnSheet("Variables only in Measles"), "The report carries the variables only in the second disease"
    Assert.IsTrue TitleOnSheet("Shared variables whose choices differ"), "The report carries the differing choices"
    Assert.IsTrue TitleOnSheet("Equivalent variables"), "The report carries the equivalent variables"
    Assert.IsTrue TitleOnSheet("Comparison statistics"), "The report carries the statistics"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestPrintReportWritesTheFiveSections", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparisonReport")
Public Sub TestPrintReportWritesTheEntriesOfEachSection()
    CustomTestSetTitles Assert, "DiseaseComparisonReport", "TestPrintReportWritesTheEntriesOfEachSection"

    Dim writer As DiseaseComparisonReport
    Dim entryRow As Long

    On Error GoTo Fail

    Set writer = DiseaseComparisonReport.Create(ThisWorkbook)
    writer.PrintComparison FirstTable, SecondTable, "Cholera", "Measles"

    entryRow = RowOfEntry("var_b (symptoms)")
    Assert.IsTrue entryRow > 0, "The variable only in the first disease reaches the sheet"
    Assert.IsTrue InStr(1, DetailAt(entryRow), "Missing from Measles", vbTextCompare) > 0, _
                  "Its detail names the disease the variable is missing from"

    entryRow = RowOfEntry("var_c (history)")
    Assert.IsTrue entryRow > 0, "The variable only in the second disease reaches the sheet"

    entryRow = RowOfEntry("var_d (outcome)")
    Assert.IsTrue entryRow > 0, "The variable whose choices differ reaches the sheet"

    entryRow = RowOfEntry("var_e (notes)")
    Assert.IsTrue entryRow > 0, "The equivalent variable reaches the sheet"

    entryRow = RowOfEntry("Both diseases")
    Assert.IsTrue entryRow > 0, "The statistics reach the sheet"
    Assert.AreEqual "3 shared variable(s).", DetailAt(entryRow), "The shared count is the one the comparison answers"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestPrintReportWritesTheEntriesOfEachSection", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparisonReport")
Public Sub TestSecondRunLeavesOnlyTheSecondReport()
    CustomTestSetTitles Assert, "DiseaseComparisonReport", "TestSecondRunLeavesOnlyTheSecondReport"

    Dim writer As DiseaseComparisonReport
    Dim firstRunLastRow As Long
    Dim secondRunLastRow As Long

    On Error GoTo Fail

    Set writer = DiseaseComparisonReport.Create(ThisWorkbook)

    writer.PrintComparison FirstTable, SecondTable, "Cholera", "Measles"
    firstRunLastRow = LastWrittenRow()

    writer.PrintComparison FirstTable, SecondTable, "Ebola", "Malaria"
    secondRunLastRow = LastWrittenRow()

    Assert.IsTrue TitleOnSheet("Variables only in Ebola"), "The sheet carries the second report"
    Assert.IsFalse TitleOnSheet("Variables only in Cholera"), "The first report is gone from the sheet"
    Assert.AreEqual firstRunLastRow, secondRunLastRow, "The second report starts where the first one did"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSecondRunLeavesOnlyTheSecondReport", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparisonReport")
Public Sub TestPrintReportShowsAHiddenReportSheet()
    CustomTestSetTitles Assert, "DiseaseComparisonReport", "TestPrintReportShowsAHiddenReportSheet"

    Dim writer As DiseaseComparisonReport

    On Error GoTo Fail

    'The deploy step hides __compRep; the button has to bring it back.
    ReportSheet.Visible = xlSheetVeryHidden

    Set writer = DiseaseComparisonReport.Create(ThisWorkbook)
    writer.PrintComparison FirstTable, SecondTable, "Cholera", "Measles"

    Assert.AreEqual CLng(xlSheetVisible), CLng(ReportSheet.Visible), "The report sheet is shown before the render starts"
    Assert.IsTrue TitleOnSheet("Comparison statistics"), "The render reaches a sheet the deploy step had hidden"
    Assert.AreEqual REPORT_SHEET_NAME, ThisWorkbook.ActiveSheet.Name, "The report sheet ends up active"

    Exit Sub

Fail:
    'The sheet is put back on show whatever happened, so the next test finds it.
    On Error Resume Next
        ReportSheet.Visible = xlSheetVisible
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestPrintReportShowsAHiddenReportSheet", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparisonReport")
Public Sub TestSheetNameDefaultsToTheCompareReport()
    CustomTestSetTitles Assert, "DiseaseComparisonReport", "TestSheetNameDefaultsToTheCompareReport"

    Dim writer As DiseaseComparisonReport

    On Error GoTo Fail

    Set writer = DiseaseComparisonReport.Create(ThisWorkbook)
    Assert.AreEqual REPORT_SHEET_NAME, writer.SheetName, "The writer defaults to the compare report sheet"

    Set writer = DiseaseComparisonReport.Create(ThisWorkbook, "   ")
    Assert.AreEqual REPORT_SHEET_NAME, writer.SheetName, "A blank name falls back to the compare report sheet"

    Assert.AreEqual REPORT_SHEET_NAME, writer.ReportSheet.Name, "The resolved worksheet is the compare report sheet"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSheetNameDefaultsToTheCompareReport", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparisonReport")
Public Sub TestCreateRefusesAMissingWorkbook()
    CustomTestSetTitles Assert, "DiseaseComparisonReport", "TestCreateRefusesAMissingWorkbook"

    Dim writer As DiseaseComparisonReport
    Dim raisedNumber As Long

    On Error GoTo Fail

    On Error Resume Next
        Set writer = DiseaseComparisonReport.Create(Nothing)
        raisedNumber = Err.Number
    On Error GoTo Fail

    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), raisedNumber, "A Nothing workbook raises ObjectNotInitialized"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateRefusesAMissingWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparisonReport")
Public Sub TestPrintReportRefusesAMissingComparison()
    CustomTestSetTitles Assert, "DiseaseComparisonReport", "TestPrintReportRefusesAMissingComparison"

    Dim writer As DiseaseComparisonReport
    Dim raisedNumber As Long

    On Error GoTo Fail

    Set writer = DiseaseComparisonReport.Create(ThisWorkbook)

    On Error Resume Next
        writer.PrintReport Nothing
        raisedNumber = Err.Number
    On Error GoTo Fail

    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), raisedNumber, "A Nothing comparison raises ObjectNotInitialized"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestPrintReportRefusesAMissingComparison", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparisonReport")
Public Sub TestPrintReportRefusesAMissingSheet()
    CustomTestSetTitles Assert, "DiseaseComparisonReport", "TestPrintReportRefusesAMissingSheet"

    Dim writer As DiseaseComparisonReport
    Dim comparison As DiseaseComparison
    Dim raisedNumber As Long

    On Error GoTo Fail

    Set writer = DiseaseComparisonReport.Create(ThisWorkbook, MISSING_SHEET_NAME)
    Set comparison = DiseaseComparison.Create(FirstTable, SecondTable)

    On Error Resume Next
        writer.PrintReport comparison
        raisedNumber = Err.Number
    On Error GoTo Fail

    Assert.AreEqual CLng(ProjectError.ElementNotFound), raisedNumber, "A workbook without the report sheet raises ElementNotFound"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestPrintReportRefusesAMissingSheet", Err.Number, Err.Description
End Sub

'@section Sheet readers
'===============================================================================

'@sub-title Say whether column B carries a bundle title.
Private Function TitleOnSheet(ByVal titleText As String) As Boolean
    Dim rowIndex As Long
    Dim lastRow As Long

    lastRow = ReportSheet.Cells(ReportSheet.rows.Count, HIDDEN_TITLE_COLUMN_INDEX).End(xlUp).row

    For rowIndex = 1 To lastRow
        If StrComp(CellTextAt(rowIndex, HIDDEN_TITLE_COLUMN_INDEX), titleText, vbTextCompare) = 0 Then
            TitleOnSheet = True
            Exit Function
        End If
    Next rowIndex
End Function

'@sub-title Row carrying an entry whose "Where?" piece is the given text, 0 when there is none.
Private Function RowOfEntry(ByVal whereText As String) As Long
    Dim rowIndex As Long
    Dim lastRow As Long

    lastRow = LastWrittenRow()

    For rowIndex = 1 To lastRow
        If StrComp(CellTextAt(rowIndex, FIRST_VISIBLE_COLUMN_INDEX + 1), whereText, vbTextCompare) = 0 Then
            RowOfEntry = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

'@sub-title The "Details" piece of an entry row.
Private Function DetailAt(ByVal rowIndex As Long) As String
    DetailAt = CellTextAt(rowIndex, FIRST_VISIBLE_COLUMN_INDEX + 2)
End Function

'@sub-title The last row of the report sheet carrying visible output.
Private Function LastWrittenRow() As Long
    LastWrittenRow = ReportSheet.Cells(ReportSheet.rows.Count, FIRST_VISIBLE_COLUMN_INDEX).End(xlUp).row
End Function

'@sub-title Read one cell of the report sheet as trimmed text.
Private Function CellTextAt(ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    Dim cellValue As Variant

    cellValue = ReportSheet.Cells(rowIndex, columnIndex).value
    If IsError(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function

    CellTextAt = Trim$(CStr(cellValue))
End Function

'@section Fixtures
'===============================================================================

Private Sub PrepareFirstTable()
    Dim headerMatrix As Variant
    Dim bodyMatrix As Variant
    Dim tableRange As Range

    Set FirstSheet = EnsureWorksheet(FIRST_SHEET_NAME)
    ClearWorksheet FirstSheet

    headerMatrix = RowsToMatrix(Array(Array("Variable Order", "Variable Name", "Variable Section", "Main Label", "Choice", "Choice Values", "Status")))
    bodyMatrix = RowsToMatrix(Array( _
        Array(1, "var_a", "demographics", "LabelA", "choiceA", "0-4 | 5-14 | 15+", "core"), _
        Array(2, "var_b", "symptoms", "LabelB", "choiceB", "yes | no", "core"), _
        Array(3, "var_d", "outcome", "LabelD", "choiceD", "yes | no", "core"), _
        Array(4, "var_e", "notes", "LabelE", "", "", "optional") _
    ))

    WriteMatrix FirstSheet.Range("A1"), headerMatrix
    WriteMatrix FirstSheet.Range("A2"), bodyMatrix

    Set tableRange = FirstSheet.Range("A1").Resize(UBound(bodyMatrix, 1) + 1, 7)
    Set FirstTable = FirstSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                XlListObjectHasHeaders:=xlYes)
    FirstTable.Name = FIRST_TABLE_NAME
End Sub

Private Sub PrepareSecondTable()
    Dim headerMatrix As Variant
    Dim bodyMatrix As Variant
    Dim tableRange As Range

    Set SecondSheet = EnsureWorksheet(SECOND_SHEET_NAME)
    ClearWorksheet SecondSheet

    headerMatrix = RowsToMatrix(Array(Array("Variable Order", "Variable Name", "Variable Section", "Main Label", "Choice", "Choice Values", "Status")))
    bodyMatrix = RowsToMatrix(Array( _
        Array(1, "VAR_A", "demographics", "LabelA", "choiceA", "15+ | 5-14 | 0-4", "core"), _
        Array(2, "var_c", "history", "LabelC", "choiceC", "low | high", "optional"), _
        Array(3, "var_d", "outcome", "LabelD", "choiceD", "yes | maybe", "core"), _
        Array(4, "var_e", "notes", "LabelE", "", "stale | values", "optional") _
    ))

    WriteMatrix SecondSheet.Range("A1"), headerMatrix
    WriteMatrix SecondSheet.Range("A2"), bodyMatrix

    Set tableRange = SecondSheet.Range("A1").Resize(UBound(bodyMatrix, 1) + 1, 7)
    Set SecondTable = SecondSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                  XlListObjectHasHeaders:=xlYes)
    SecondTable.Name = SECOND_TABLE_NAME
End Sub
