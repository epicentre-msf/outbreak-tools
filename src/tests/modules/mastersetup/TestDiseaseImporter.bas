Attribute VB_Name = "TestDiseaseImporter"
Attribute VB_Description = "Tests covering DiseaseImporter merge behaviour"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests covering DiseaseImporter merge behaviour")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TARGET_SHEET_NAME As String = "DiseaseImportTarget"
Private Const SOURCE_SHEET_NAME As String = "DiseaseImportSource"
Private Const TARGET_TABLE_NAME As String = "T_TargetDisease"
Private Const SOURCE_TABLE_NAME As String = "T_SourceDisease"

Private Assert As ICustomTest
Private Importer As IDiseaseImporter
Private TargetSheet As Worksheet
Private SourceSheet As Worksheet
Private TargetTable As ListObject
Private SourceTable As ListObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseImporter"
    Set Importer = New DiseaseImporter
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        DeleteWorksheets TARGET_SHEET_NAME, SOURCE_SHEET_NAME
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
    Set Importer = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    PrepareTargetTable
    PrepareSourceTable
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not TargetSheet Is Nothing Then ClearWorksheet TargetSheet
    If Not SourceSheet Is Nothing Then ClearWorksheet SourceSheet
    Set TargetTable = Nothing
    Set SourceTable = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseImporter")
Public Sub TestMergeUpdatesExistingAndAppendsNew()
    CustomTestSetTitles Assert, "DiseaseImporter", "TestMergeUpdatesExistingAndAppendsNew"

    Dim summary As IDiseaseImportSummary
    Dim data As Variant
    Dim missing As BetterArray
    Dim updated As BetterArray
    Dim appended As BetterArray

    On Error GoTo Fail

    Set summary = Importer.MergeDisease(TargetTable, SourceTable, True, DiseaseImportPriority_Foreign)

    data = TargetTable.DataBodyRange.Value
    Assert.AreEqual "LabelAUpdated", data(1, 2), "Existing variable should be updated from import"
    Assert.AreEqual "choiceA2", data(1, 5), "Choice column should be updated"
    Assert.AreEqual "LabelB", data(2, 2), "Unimported variable should keep original values"
    Assert.AreEqual "var_c", data(3, 1), "New variable should be appended"

    Set missing = summary.MissingVariables
    Assert.AreEqual 1, missing.Length, "Exactly one variable should be missing"
    Assert.AreEqual "var_b", missing.Item(missing.LowerBound), "var_b should be flagged as missing"

    Set updated = summary.UpdatedVariables
    Assert.AreEqual 1, updated.Length, "One variable should be updated"
    Assert.AreEqual "var_a", updated.Item(updated.LowerBound), "var_a should be flagged as updated"

    Set appended = summary.AppendedVariables
    Assert.AreEqual 1, appended.Length, "One variable should be appended"
    Assert.AreEqual "var_c", appended.Item(appended.LowerBound), "var_c should be flagged as appended"

    Assert.IsTrue summary.RequiresReport, "Missing or appended variables should flag reports"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMergeUpdatesExistingAndAppendsNew", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseImporter")
Public Sub TestMergeDiseaseLogsOperations()
    CustomTestSetTitles Assert, "DiseaseImporter", "TestMergeDiseaseLogsOperations"

    Dim summary As IDiseaseImportSummary
    Dim logger As IDiseaseLogger
    Dim entries As BetterArray

    On Error GoTo Fail

    Set logger = New DiseaseLogger
    Set summary = Importer.MergeDisease(TargetTable, SourceTable, True, DiseaseImportPriority_Foreign, logger)

    Assert.IsTrue logger.HasEntries, "Merge should record logging information"

    Set entries = logger.Entries
    Assert.IsTrue entries.Length >= 3, "Logger should include entries for updated, appended, and missing variables"

    Assert.IsTrue summary.RequiresReport, "Logging test should still reflect summary state"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMergeDiseaseLogsOperations", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseImporter")
Public Sub TestReplaceTableCopiesSourceWhenMergeDisabled()
    CustomTestSetTitles Assert, "DiseaseImporter", "TestReplaceTableCopiesSourceWhenMergeDisabled"

    Dim summary As IDiseaseImportSummary
    Dim data As Variant

    On Error GoTo Fail

    Set summary = Importer.MergeDisease(TargetTable, SourceTable, False, DiseaseImportPriority_Foreign)

    data = TargetTable.DataBodyRange.Value
    Assert.AreEqual "var_a", data(1, 1), "First row variable should match source"
    Assert.AreEqual "LabelC", data(2, 2), "Second row label should match source"
    Assert.AreEqual 2, TargetTable.ListRows.Count, "Target table should match source row count"

    Assert.IsTrue summary.AppendedVariables.Length > 0, "Summary should contain appended variables after replace"
    Assert.IsTrue summary.RequiresReport, "Replacing should flag a report"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestReplaceTableCopiesSourceWhenMergeDisabled", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

Private Sub PrepareTargetTable()
    Dim headerMatrix As Variant
    Dim bodyMatrix As Variant
    Dim tableRange As Range

    Set TargetSheet = EnsureWorksheet(TARGET_SHEET_NAME)
    ClearWorksheet TargetSheet

    headerMatrix = RowsToMatrix(Array(Array("Variable", "Label", "Type", "Format", "Choice", "Active")))
    bodyMatrix = RowsToMatrix(Array( _
        Array("var_a", "LabelA", "string", "formatA", "choiceA", "yes"), _
        Array("var_b", "LabelB", "number", "formatB", "choiceB", "yes") _
    ))

    WriteMatrix TargetSheet.Range("A1"), headerMatrix
    WriteMatrix TargetSheet.Range("A2"), bodyMatrix

    Set tableRange = TargetSheet.Range("A1").Resize(UBound(bodyMatrix, 1) + 1, 6)
    Set TargetTable = TargetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                  XlListObjectHasHeaders:=xlYes)
    TargetTable.Name = TARGET_TABLE_NAME
End Sub

Private Sub PrepareSourceTable()
    Dim headerMatrix As Variant
    Dim bodyMatrix As Variant
    Dim tableRange As Range

    Set SourceSheet = EnsureWorksheet(SOURCE_SHEET_NAME)
    ClearWorksheet SourceSheet

    headerMatrix = RowsToMatrix(Array(Array("Variable", "Label", "Type", "Format", "Choice", "Active")))
    bodyMatrix = RowsToMatrix(Array( _
        Array("var_a", "LabelAUpdated", "string", "formatA2", "choiceA2", "no"), _
        Array("var_c", "LabelC", "string", "formatC", "choiceC", "yes") _
    ))

    WriteMatrix SourceSheet.Range("A1"), headerMatrix
    WriteMatrix SourceSheet.Range("A2"), bodyMatrix

    Set tableRange = SourceSheet.Range("A1").Resize(UBound(bodyMatrix, 1) + 1, 6)
    Set SourceTable = SourceSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                  XlListObjectHasHeaders:=xlYes)
    SourceTable.Name = SOURCE_TABLE_NAME
End Sub
