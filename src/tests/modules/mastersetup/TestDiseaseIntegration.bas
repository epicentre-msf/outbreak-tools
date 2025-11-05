Attribute VB_Name = "TestDiseaseIntegration"
Attribute VB_Description = "Integration tests covering disease add/export/import/remove workflows"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Integration tests covering disease add/export/import/remove workflows")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const ANCHOR_SHEET As String = "Variables"
Private Const DROPDOWN_SHEET As String = "IntegrationDropdown"
Private Const TRANSLATION_SHEET As String = "IntegrationTranslations"
Private Const IMPORT_SHEET As String = "IntegrationImport"

Private Assert As ICustomTest
Private Builder As IDiseaseSheetBuilder
Private Importer As IDiseaseImporter
Private Exporter As IDiseaseExporter
Private ExportManager As IDiseaseExportWorkbook
Private AppGuard As IDiseaseApplicationState
Private Dropdowns As IDropdownLists
Private RibbonTranslations As ITranslationObject
Private TranslationTable As ListObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseIntegration"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        CleanupEnvironment
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
    Set Builder = Nothing
    Set Importer = Nothing
    Set Exporter = Nothing
    Set ExportManager = Nothing
    Set AppGuard = Nothing
    Set Dropdowns = Nothing
    Set RibbonTranslations = Nothing
    Set TranslationTable = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    CleanupEnvironment
    PrepareEnvironment
End Sub

'@TestCleanup
Private Sub TestCleanup()
    CleanupEnvironment
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseIntegration")
Public Sub TestAddExportImportRemove()
    CustomTestSetTitles Assert, "DiseaseIntegration", "TestAddExportImportRemove"

    Dim diseaseSheet As Worksheet
    Dim diseaseTable As ListObject
    Dim exportBook As Workbook
    Dim logger As IDiseaseLogger
    Dim summary As IDiseaseImportSummary
    Dim entries As BetterArray
    Dim importTable As ListObject
    Dim manager As IDiseaseWorksheetManager

    On Error GoTo Fail

    Set diseaseSheet = Builder.Build("Alpha", 1)
    Set diseaseTable = diseaseSheet.ListObjects(1)

    PopulateDiseaseTable diseaseTable

    Set exportBook = Exporter.BuildDiseaseWorkbook(diseaseSheet, TranslationTable, RibbonTranslations, _
                                                   "Alpha", diseaseSheet.Cells(2, 2).Value, diseaseSheet.Cells(2, 3).Value)

    Assert.AreEqual "Alpha", exportBook.Worksheets("Metadata").Cells(3, 2).Value, "Metadata should reference disease name"
    Assert.AreEqual "LabelA", exportBook.Worksheets("Dictionary").Cells(2, 4).Value, "Dictionary should capture existing variables"

    exportBook.Close SaveChanges:=False

    Set importTable = PrepareImportTable()
    Set logger = New DiseaseLogger

    Set summary = Importer.MergeDisease(diseaseTable, importTable, True, DiseaseImportPriority_Foreign, logger)

    Assert.AreEqual "LabelAUpdated", diseaseTable.DataBodyRange.Cells(1, 2).Value, "Merge should update existing variable label"
    Assert.AreEqual "var_d", diseaseTable.DataBodyRange.Cells(3, 1).Value, "Merge should append new variables"
    Assert.IsTrue summary.RequiresReport, "Summary should indicate report requirement"

    Assert.IsTrue logger.HasEntries, "Logger should capture merge operations"
    Set entries = logger.Entries
    Assert.IsTrue entries.Length >= 3, "Logger should contain multiple entries for merge operations"

    Set manager = New DiseaseWorksheetManager
    Assert.IsTrue manager.RemoveWorksheet(ThisWorkbook, "Alpha"), "Worksheet manager should remove disease sheet"
    Assert.IsFalse WorksheetExists("Alpha"), "Disease worksheet should be removed"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddExportImportRemove", Err.Number, Err.Description
End Sub

'@section Helpers
'===============================================================================

Private Sub PrepareEnvironment()
    Dim dropdownSheet As Worksheet
    Dim translationSheet As Worksheet
    Dim data As Variant

    EnsureWorksheet ANCHOR_SHEET

    Set dropdownSheet = EnsureWorksheet(DROPDOWN_SHEET)
    ClearWorksheet dropdownSheet

    dropdownSheet.Range("A1").Value = "Languages"
    dropdownSheet.Range("A2:A4").Value = Application.WorksheetFunction.Transpose(Array("ENG", "FRA", "ESP"))
    dropdownSheet.Range("B1").Value = "Status"
    dropdownSheet.Range("B2:B4").Value = Application.WorksheetFunction.Transpose(Array("core", "optional", "hidden"))
    dropdownSheet.Range("C1").Value = "VarNames"
    dropdownSheet.Range("C2:C5").Value = Application.WorksheetFunction.Transpose(Array("var_a", "var_b", "var_c", "var_d"))
    dropdownSheet.Range("D1").Value = "Choices"
    dropdownSheet.Range("D2:D5").Value = Application.WorksheetFunction.Transpose(Array("choice_age", "choice_fever", "choice_other", "choice_new"))

    AddName "__languages", dropdownSheet.Range("A2:A4")
    AddName "__var_status", dropdownSheet.Range("B2:B4")
    AddName "PARAMVARNAME", dropdownSheet.Range("C2:C5")
    AddName "PARAMCHOICESLIST", dropdownSheet.Range("D2:D5")

    Set Dropdowns = New TestDropdownStub

    Set translationSheet = EnsureWorksheet(TRANSLATION_SHEET)
    ClearWorksheet translationSheet

    data = Array( _
        Array("tag", "ENG"), _
        Array("selectValue", "Select a value"), _
        Array("infoSelectLang", "Select language"), _
        Array("varOrder", "Variable Order"), _
        Array("varSection", "Variable Section"), _
        Array("varName", "Variable Name"), _
        Array("varLabel", "Main Label"), _
        Array("varChoice", "Control"), _
        Array("choiceVal", "Choices"), _
        Array("varStatus", "Status"), _
        Array("errLang", "Please select a language") _
    )

    translationSheet.Range("A1").Resize(UBound(data) + 1, 2).Value = data
    translationSheet.ListObjects.Add SourceType:=xlSrcRange, _
                                      Source:=translationSheet.Range("A1").Resize(UBound(data) + 1, 2), _
                                      XlListObjectHasHeaders:=xlYes

    Set TranslationTable = translationSheet.ListObjects(1)
    Set RibbonTranslations = TranslationObject.Create(TranslationTable, "ENG")

    Set Builder = DiseaseSheetBuilder.Create(ThisWorkbook, Dropdowns, RibbonTranslations)
    Set Importer = New DiseaseImporter
    Set ExportManager = New DiseaseExportWorkbook
    Set AppGuard = New DiseaseApplicationState
    Set Exporter = DiseaseExporter.Create(ExportManager, AppGuard)
End Sub

Private Sub CleanupEnvironment()
    DeleteWorksheetSafe "Alpha"
    DeleteWorksheetSafe DROPDOWN_SHEET
    DeleteWorksheetSafe TRANSLATION_SHEET
    DeleteWorksheetSafe IMPORT_SHEET
    DeleteNameSafe "__languages"
    DeleteNameSafe "__var_status"
    DeleteNameSafe "PARAMVARNAME"
    DeleteNameSafe "PARAMCHOICESLIST"
    DeleteNameSafe "disLang_1"
End Sub

Private Sub PopulateDiseaseTable(ByVal table As ListObject)
    table.ListRows(1).Range.Value = Array("var_a", "LabelA", "string", "formatA", "choice_age", "0-4 | 5-14", "core")
    table.ListRows(2).Range.Value = Array("var_b", "LabelB", "number", "formatB", "choice_fever", "yes | no", "core")
End Sub

Private Function PrepareImportTable() As ListObject
    Dim importSheet As Worksheet
    Dim header As Variant
    Dim rows As Variant
    Dim tableRange As Range

    Set importSheet = EnsureWorksheet(IMPORT_SHEET)
    ClearWorksheet importSheet

    header = RowsToMatrix(Array(Array("Variable", "Label", "Type", "Format", "Choice", "Choices", "Status")))
    rows = RowsToMatrix(Array( _
        Array("var_a", "LabelAUpdated", "string", "formatAUpdated", "choice_age", "0-4 | 5-14", "core"), _
        Array("var_c", "LabelC", "string", "formatC", "choice_other", "low | high", "optional"), _
        Array("var_d", "LabelD", "string", "formatD", "choice_new", "alpha | beta", "core") _
    ))

    WriteMatrix importSheet.Range("A1"), header
    WriteMatrix importSheet.Range("A2"), rows

    Set tableRange = importSheet.Range("A1").Resize(UBound(rows, 1) + 1, UBound(rows, 2))
    Set PrepareImportTable = importSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                         XlListObjectHasHeaders:=xlYes)
End Function

Private Sub DeleteWorksheetSafe(ByVal sheetName As String)
    On Error Resume Next
        ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
End Sub

Private Sub DeleteNameSafe(ByVal nameValue As String)
    On Error Resume Next
        ThisWorkbook.Names(nameValue).Delete
    On Error GoTo 0
End Sub

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    WorksheetExists = Not FindWorksheet(ThisWorkbook, sheetName) Is Nothing
End Function

Private Function FindWorksheet(ByVal workbook As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet

    For Each sheet In workbook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = sheet
            Exit Function
        End If
    Next sheet
End Function

Private Sub AddName(ByVal nameValue As String, ByVal targetRange As Range)
    DeleteNameSafe nameValue
    ThisWorkbook.Names.Add Name:=nameValue, RefersTo:=targetRange
End Sub
