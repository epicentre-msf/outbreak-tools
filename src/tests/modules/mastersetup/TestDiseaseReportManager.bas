Attribute VB_Name = "TestDiseaseReportManager"
Attribute VB_Description = "Tests validating DiseaseReportManager reporting helpers"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests validating DiseaseReportManager reporting helpers")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const REPORT_SHEET_NAME As String = "DiseaseReportFixture"
Private Const MAIN_REPORT_TABLE As String = "T_ReportMain"
Private Const SECONDARY_REPORT_TABLE As String = "T_ReportSecondary"

Private Assert As ICustomTest
Private ReportManager As IDiseaseReportManager
Private ReportSheet As Worksheet
Private MainTable As ListObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseReportManager"
    Set ReportManager = New DiseaseReportManager
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        DeleteWorksheet REPORT_SHEET_NAME
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
    Set ReportManager = Nothing
    Set ReportSheet = Nothing
    Set MainTable = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    PrepareReportSheet
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not ReportSheet Is Nothing Then
        ClearWorksheet ReportSheet
    End If
    Set MainTable = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseReportManager")
Public Sub TestHasPendingReportReturnsTrue()
    CustomTestSetTitles Assert, "DiseaseReportManager", "TestHasPendingReportReturnsTrue"

    Dim requiresReport As Boolean

    On Error GoTo Fail

    requiresReport = ReportManager.HasPendingReport(MainTable, "Ebola")
    Assert.IsTrue requiresReport, "Ebola entry should require a report"

    requiresReport = ReportManager.HasPendingReport(MainTable, "Influenza")
    Assert.IsFalse requiresReport, "Influenza entry should not require a report"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestHasPendingReportReturnsTrue", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseReportManager")
Public Sub TestClearReportStatusRemovesRows()
    CustomTestSetTitles Assert, "DiseaseReportManager", "TestClearReportStatusRemovesRows"

    Dim diseaseName As String
    Dim remaining As Boolean

    On Error GoTo Fail

    diseaseName = "Ebola"
    ReportManager.ClearReportStatus ReportSheet, diseaseName

    remaining = ContainsDisease(MainTable, diseaseName)
    Assert.IsFalse remaining, "Main table should no longer contain the disease row"

    remaining = ContainsDisease(ReportSheet.ListObjects(SECONDARY_REPORT_TABLE), diseaseName)
    Assert.IsFalse remaining, "Secondary table should no longer contain the disease row"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestClearReportStatusRemovesRows", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

Private Sub PrepareReportSheet()
    Dim headerMatrix As Variant
    Dim mainData As Variant
    Dim secondaryData As Variant
    Dim tableRange As Range

    Set ReportSheet = EnsureWorksheet(REPORT_SHEET_NAME)
    ClearWorksheet ReportSheet

    headerMatrix = RowsToMatrix(Array(Array("Disease", "NeedReport")))
    mainData = RowsToMatrix(Array( _
        Array("Ebola", "yes"), _
        Array("Influenza", "no") _
    ))

    secondaryData = RowsToMatrix(Array(Array("Ebola", "note", "urgent")))

    WriteMatrix ReportSheet.Range("A1"), headerMatrix
    WriteMatrix ReportSheet.Range("A2"), mainData
    Set tableRange = ReportSheet.Range("A1").Resize(UBound(mainData, 1) + 1, 2)
    Set MainTable = ReportSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                XlListObjectHasHeaders:=xlYes)
    MainTable.Name = MAIN_REPORT_TABLE

    WriteMatrix ReportSheet.Range("E1"), RowsToMatrix(Array(Array("Disease", "Label", "Status")))
    WriteMatrix ReportSheet.Range("E2"), secondaryData
    Set tableRange = ReportSheet.Range("E1").Resize(UBound(secondaryData, 1) + 1, 3)
    With ReportSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                     XlListObjectHasHeaders:=xlYes)
        .Name = SECONDARY_REPORT_TABLE
    End With
End Sub

'@section Helpers
'===============================================================================

Private Function ContainsDisease(ByVal reportTable As ListObject, ByVal diseaseName As String) As Boolean
    Dim rowIndex As Long
    Dim totalRows As Long

    If reportTable Is Nothing Then Exit Function
    totalRows = reportTable.ListRows.Count
    If totalRows = 0 Then Exit Function

    For rowIndex = 1 To totalRows
        If StrComp(reportTable.DataBodyRange.Cells(rowIndex, 1).Value, diseaseName, vbTextCompare) = 0 Then
            ContainsDisease = True
            Exit Function
        End If
    Next rowIndex
End Function
