Attribute VB_Name = "TestDiseaseExporter"
Attribute VB_Description = "Tests validating DiseaseExporter builds array-based disease and migration workbooks"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests validating DiseaseExporter builds array-based disease and migration workbooks")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const DISEASE_SHEET_PREFIX As String = "DiseaseTest_"
Private Const TRANSLATION_SHEET As String = "TranslationFixture"
Private Const RIBBON_SHEET As String = "RibbonFixture"

Private Assert As ICustomTest
Private Exporter As IDiseaseExporter
Private Manager As IDiseaseExportWorkbook
Private Guard As IDiseaseApplicationState
Private TempFolder As String

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseExporter"

    Set Manager = New DiseaseExportWorkbook
    Set Guard = New DiseaseApplicationState
    Set Exporter = DiseaseExporter.Create(Manager, Guard)

    TempFolder = ThisWorkbook.Path & Application.PathSeparator & "temp"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        Guard.Restore
        Manager.Release
        DeleteFixtureSheets
    On Error GoTo 0

    RestoreApp
    Set Exporter = Nothing
    Set Manager = Nothing
    Set Guard = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Guard.Restore
    Manager.Release
    DeleteFixtureSheets
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Guard.Restore
    Manager.Release
    DeleteFixtureSheets
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseExporter")
Public Sub TestBuildDiseaseWorkbookCopiesDictionaryAndChoices()
    CustomTestSetTitles Assert, "DiseaseExporter", "TestBuildDiseaseWorkbookCopiesDictionaryAndChoices"

    Dim diseaseSheet As Worksheet
    Dim translationTable As ListObject
    Dim ribbonTranslations As ITranslationObject
    Dim workbook As Workbook
    Dim dictionaryValues As Variant
    Dim choicesValues As Variant

    On Error GoTo Fail

    Set diseaseSheet = PrepareDiseaseWorksheet("Alpha", "ENG", "ALPHA_CODE")
    Set translationTable = PrepareTranslationTable()
    Set ribbonTranslations = PrepareRibbonTranslations()

    Set workbook = Exporter.BuildDiseaseWorkbook(diseaseSheet, translationTable, ribbonTranslations, _
                                                diseaseSheet.Name, diseaseSheet.Cells(2, 2).Value, "ALPHA_CODE")

    dictionaryValues = workbook.Worksheets("Dictionary").Range("A2").Resize(2, 6).Value
    choicesValues = workbook.Worksheets("Choices").Range("A2").Resize(4, 4).Value

    Assert.AreEqual 1, dictionaryValues(1, 1), "First variable order should be copied"
    Assert.AreEqual "core", dictionaryValues(1, 6), "Status column should be copied"
    Assert.AreEqual "symptoms", dictionaryValues(2, 2), "Section should be copied"

    Assert.AreEqual "choice_age", choicesValues(1, 1), "Control name should populate choices sheet"
    Assert.AreEqual "0-4", choicesValues(1, 2), "Choice value should populate choices sheet"
    Assert.AreEqual 2, choicesValues(2, 4), "Ordering should follow original order"
    Assert.AreEqual "choice_fever", choicesValues(3, 1), "Multiple controls should be captured"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBuildDiseaseWorkbookCopiesDictionaryAndChoices", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseExporter")
Public Sub TestBuildMigrationWorkbookAggregatesDiseases()
    CustomTestSetTitles Assert, "DiseaseExporter", "TestBuildMigrationWorkbookAggregatesDiseases"

    Dim diseaseNames As BetterArray
    Dim workbook As Workbook
    Dim diseasesSheet As Worksheet
    Dim translationsSheet As Worksheet
    Dim diseaseMeta As Variant
    Dim translationValues As Variant

    On Error GoTo Fail

    Dim sourceWorkbook As Workbook
    Set sourceWorkbook = PrepareMigrationSourceWorkbook()

    Set diseaseNames = New BetterArray
    diseaseNames.Push "Beta"
    diseaseNames.Push "Gamma"

    Set workbook = Exporter.BuildMigrationWorkbook(sourceWorkbook, diseaseNames)
    Set diseasesSheet = workbook.Worksheets("Diseases")
    Set translationsSheet = workbook.Worksheets("Translations")

    diseaseMeta = diseasesSheet.Range("A1").Resize(2, 6).Value
    translationValues = translationsSheet.Range("A1").Resize(3, 2).Value

    Assert.AreEqual "Disease", diseaseMeta(1, 1), "Metadata headers should be created"
    Assert.AreEqual "Beta", diseaseMeta(2, 1), "First disease should be copied"
    Assert.AreEqual "Gamma", diseaseMeta(2, 4), "Second disease metadata should be appended"
    Assert.AreEqual "ENG", diseaseMeta(2, 2), "Language should be copied with metadata"
    Assert.AreEqual "FRA", diseaseMeta(2, 5), "Second disease language should be copied"

    Assert.AreEqual "tag", translationValues(1, 1), "Translation header should be copied"
    Assert.AreEqual "hello", translationValues(2, 1), "Translation rows should be copied"

    sourceWorkbook.Close SaveChanges:=False
    Exit Sub

Fail:
    On Error Resume Next
        sourceWorkbook.Close SaveChanges:=False
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestBuildMigrationWorkbookAggregatesDiseases", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

Private Function PrepareDiseaseWorksheet(ByVal diseaseName As String, _
                                         ByVal languageTag As String, _
                                         ByVal diseaseCode As String) As Worksheet

    Dim sheet As Worksheet
    Dim header As Variant
    Dim dataRows As Variant
    Dim startRange As Range
    Dim listRange As Range

    DeleteWorksheet diseaseName
    Set sheet = EnsureWorksheet(diseaseName)
    ClearWorksheet sheet

    sheet.Cells(2, 2).Value = languageTag
    sheet.Cells(2, 3).Value = diseaseCode

    header = RowsToMatrix(Array(Array("Order", "Section", "Name", "Label", "Control", "Choices", "Status")))
    dataRows = RowsToMatrix(Array( _
        Array(1, "demographics", "age", "Age", "choice_age", "0-4 | 5-14 | 15+", "core"), _
        Array(2, "symptoms", "fever", "Fever", "choice_fever", "yes | no", "core") _
    ))

    Set startRange = sheet.Range("B4")
    WriteMatrix startRange, header
    WriteMatrix startRange.Offset(1), dataRows

    Set listRange = sheet.Range("B4").Resize(UBound(dataRows, 1) + 1, UBound(dataRows, 2))
    sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=listRange, XlListObjectHasHeaders:=xlYes

    Set PrepareDiseaseWorksheet = sheet
End Function

Private Function PrepareTranslationTable() As ListObject
    Dim sheet As Worksheet
    Dim header As Variant
    Dim dataRows As Variant
    Dim listRange As Range

    DeleteWorksheet TRANSLATION_SHEET
    Set sheet = EnsureWorksheet(TRANSLATION_SHEET)
    ClearWorksheet sheet

    header = RowsToMatrix(Array(Array("tag", "ENG")))
    dataRows = RowsToMatrix(Array( _
        Array("hello", "Hello"), _
        Array("world", "World") _
    ))

    WriteMatrix sheet.Range("A1"), header
    WriteMatrix sheet.Range("A2"), dataRows

    Set listRange = sheet.Range("A1").Resize(UBound(dataRows, 1) + 1, UBound(dataRows, 2))
    sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=listRange, XlListObjectHasHeaders:=xlYes

    Set PrepareTranslationTable = sheet.ListObjects(1)
End Function

Private Function PrepareRibbonTranslations() As ITranslationObject
    Dim sheet As Worksheet
    Dim header As Variant
    Dim dataRows As Variant
    Dim listRange As Range

    DeleteWorksheet RIBBON_SHEET
    Set sheet = EnsureWorksheet(RIBBON_SHEET)
    ClearWorksheet sheet

    header = RowsToMatrix(Array(Array("tag", "ENG")))
    dataRows = RowsToMatrix(Array( _
        Array("list name", "List Name"), _
        Array("label", "Label"), _
        Array("short label", "Short Label"), _
        Array("ordering list", "Ordering") _
    ))

    WriteMatrix sheet.Range("A1"), header
    WriteMatrix sheet.Range("A2"), dataRows

    Set listRange = sheet.Range("A1").Resize(UBound(dataRows, 1) + 1, UBound(dataRows, 2))
    sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=listRange, XlListObjectHasHeaders:=xlYes

    Set PrepareRibbonTranslations = TranslationObject.Create(sheet.ListObjects(1), "ENG")
End Function

Private Sub DeleteFixtureSheets()
    DeleteWorksheet TRANSLATION_SHEET
    DeleteWorksheet RIBBON_SHEET
    DeleteWorksheet "Alpha"
End Sub

Private Function PrepareMigrationSourceWorkbook() As Workbook
    Dim wb As Workbook
    Dim translations As Worksheet
    Dim choices As Worksheet
    Dim variables As Worksheet
    Dim diseaseBeta As Worksheet
    Dim diseaseGamma As Worksheet
    Dim values As Variant
    Dim rangeObj As Range

    Set wb = Workbooks.Add

    Set translations = wb.Worksheets(1)
    translations.Name = "Translations"
    Set choices = wb.Worksheets(2)
    choices.Name = "Choices"
    Set variables = wb.Worksheets(3)
    variables.Name = "Variables"
    Set diseaseBeta = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    diseaseBeta.Name = "Beta"
    Set diseaseGamma = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    diseaseGamma.Name = "Gamma"

    values = RowsToMatrix(Array( _
        Array("tag", "ENG"), _
        Array("hello", "hello"), _
        Array("world", "world") _
    ))
    translations.Range("A1").Resize(UBound(values, 1), UBound(values, 2)).Value = values
    Set rangeObj = translations.Range("A1").Resize(UBound(values, 1), UBound(values, 2))
    translations.ListObjects.Add SourceType:=xlSrcRange, Source:=rangeObj, XlListObjectHasHeaders:=xlYes

    values = RowsToMatrix(Array( _
        Array("List", "Label"), _
        Array("choice_age", "0-4"), _
        Array("choice_age", "5-14") _
    ))
    choices.Range("A1").Resize(UBound(values, 1), UBound(values, 2)).Value = values
    Set rangeObj = choices.Range("A1").Resize(UBound(values, 1), UBound(values, 2))
    choices.ListObjects.Add SourceType:=xlSrcRange, Source:=rangeObj, XlListObjectHasHeaders:=xlYes

    values = RowsToMatrix(Array( _
        Array("Tag", "Label"), _
        Array("age", "Age"), _
        Array("fever", "Fever") _
    ))
    variables.Range("A1").Resize(UBound(values, 1), UBound(values, 2)).Value = values
    Set rangeObj = variables.Range("A1").Resize(UBound(values, 1), UBound(values, 2))
    variables.ListObjects.Add SourceType:=xlSrcRange, Source:=rangeObj, XlListObjectHasHeaders:=xlYes

    PopulateDiseaseSheet diseaseBeta, RowsToMatrix(Array( _
        Array("Order", "Section", "Name", "Label", "Control", "Choices", "Status"), _
        Array(1, "demographics", "age", "Age", "choice_age", "0-4 | 5-14", "core"), _
        Array(2, "symptoms", "fever", "Fever", "choice_fever", "yes | no", "core") _
    )), "ENG", "BETA_CODE"

    PopulateDiseaseSheet diseaseGamma, RowsToMatrix(Array( _
        Array("Order", "Section", "Name", "Label", "Control", "Choices", "Status"), _
        Array(1, "history", "travel", "Recent travel", "choice_travel", "yes | no", "core") _
    )), "FRA", "GAMMA_CODE"

    Set PrepareMigrationSourceWorkbook = wb
End Function

Private Sub PopulateDiseaseSheet(ByVal sheet As Worksheet, _
                                 ByVal data As Variant, _
                                 ByVal languageTag As String, _
                                 ByVal diseaseCode As String)

    Dim rowCount As Long
    Dim columnCount As Long
    Dim rangeObj As Range

    sheet.Cells.Clear
    sheet.Cells(2, 2).Value = languageTag
    sheet.Cells(2, 3).Value = diseaseCode

    rowCount = UBound(data, 1)
    columnCount = UBound(data, 2)

    sheet.Range("B4").Resize(rowCount, columnCount).Value = data
    Set rangeObj = sheet.Range("B4").Resize(rowCount, columnCount)
    sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=rangeObj, XlListObjectHasHeaders:=xlYes
End Sub
