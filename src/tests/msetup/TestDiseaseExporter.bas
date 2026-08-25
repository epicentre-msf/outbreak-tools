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
Private Const MASTER_CHOICES_SHEET As String = "Choices"

Private Assert As CustomTest
Private Exporter As DiseaseExporter
Private Manager As DiseaseExportWorkbook
Private Guard As ApplicationState
Private TempFolder As String

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseExporter"

    Set Manager = DiseaseExportWorkbook.Create()
    Set Guard = ApplicationState.Create(Application)
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
        Manager.ReleaseWorkbook
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
    Manager.ReleaseWorkbook
    DeleteFixtureSheets
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'The assertions of a test reach the results sheet only once flushed.
    Assert.Flush
    Guard.Restore
    Manager.ReleaseWorkbook
    DeleteFixtureSheets
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseExporter")
Public Sub TestBuildDiseaseWorkbookCopiesDictionaryAndChoices()
    CustomTestSetTitles Assert, "DiseaseExporter", "TestBuildDiseaseWorkbookCopiesDictionaryAndChoices"

    Dim diseaseWksh As Worksheet
    Dim translationTable As ListObject
    Dim targetBook As Workbook
    Dim dictionaryHeaders As Variant
    Dim dictionaryValues As Variant
    Dim choicesHeaders As Variant
    Dim choicesValues As Variant

    On Error GoTo Fail

    Set diseaseWksh = PrepareDiseaseWorksheet("Alpha", "ENG", "ALPHA_CODE")
    Set translationTable = PrepareTranslationTable()

    Set targetBook = Exporter.BuildDiseaseWorkbook(diseaseWksh, translationTable, _
                                                diseaseWksh.Name, diseaseWksh.Cells(2, 2).Value, "ALPHA_CODE")

    dictionaryHeaders = targetBook.Worksheets("Dictionary").Range("A1").Resize(1, 6).Value
    dictionaryValues = targetBook.Worksheets("Dictionary").Range("A2").Resize(2, 6).Value
    choicesHeaders = targetBook.Worksheets("Choices").Range("A1").Resize(1, 6).Value
    choicesValues = targetBook.Worksheets("Choices").Range("A2").Resize(6, 6).Value

    Assert.AreEqual 1, dictionaryValues(1, 1), "First variable order should be copied"
    Assert.AreEqual "core", dictionaryValues(1, 6), "Status column should be copied"
    Assert.AreEqual "symptoms", dictionaryValues(2, 2), "Section should be copied"
    Assert.AreEqual "Variable Order", dictionaryHeaders(1, 1), "The order column keeps its name"
    Assert.AreEqual "Main Section", dictionaryHeaders(1, 2), "The section column carries the setup header"
    Assert.AreEqual "Variable Name", dictionaryHeaders(1, 3), "The name column carries the setup header"

    Assert.AreEqual "Tab_Dictionary", targetBook.Worksheets("Dictionary").ListObjects(1).Name, _
                    "The dictionary block is the table a setup names Tab_Dictionary"
    Assert.AreEqual "Tab_Choices", targetBook.Worksheets("Choices").ListObjects(1).Name, _
                    "The choices block is the table a setup names Tab_Choices"
    Assert.AreEqual "Tab_Translations", targetBook.Worksheets("Translations").ListObjects(1).Name, _
                    "The translations block is the table a setup names Tab_Translations"

    Assert.AreEqual "list name", choicesHeaders(1, 1), "The choices sheet carries the setup headers"
    Assert.AreEqual "ordering list", choicesHeaders(1, 2), "The ordering list is the second setup column"
    Assert.AreEqual "non translated label", choicesHeaders(1, 3), "The non translated label is the third setup column"
    Assert.AreEqual "translated label", choicesHeaders(1, 4), "The translated label is the fourth setup column"
    Assert.AreEqual "label", choicesHeaders(1, 5), "The label is the fifth setup column"
    Assert.AreEqual "short label", choicesHeaders(1, 6), "The short label is the sixth setup column"

    Assert.AreEqual "choice_age", choicesValues(1, 1), "Control name should populate choices sheet"
    Assert.AreEqual "0-4", choicesValues(1, 3), "Choice value should populate the non translated label"
    Assert.AreEqual "0-4", choicesValues(1, 4), "Choice value should populate the translated label"
    Assert.AreEqual "0-4", choicesValues(1, 5), "The label formula should read the translated label"
    Assert.AreEqual "0-4", choicesValues(1, 6), "Choice value should populate the short label"
    Assert.AreEqual 2, choicesValues(2, 2), "Ordering should follow original order"
    Assert.AreEqual "choice_fever", choicesValues(4, 1), "Multiple controls should be captured"
    Assert.IsTrue IsEmpty(choicesValues(6, 1)), "Two lists of three and two values fill five rows"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBuildDiseaseWorkbookCopiesDictionaryAndChoices", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseExporter")
Public Sub TestChoicesExportCoversEveryChoiceOnce()
    CustomTestSetTitles Assert, "DiseaseExporter", "TestChoicesExportCoversEveryChoiceOnce"

    Dim diseaseWksh As Worksheet
    Dim translationTable As ListObject
    Dim targetBook As Workbook
    Dim choicesValues As Variant

    On Error GoTo Fail

    Set diseaseWksh = PrepareDiseaseWorksheetRows("Alpha", "ENG", "ALPHA_CODE", Array( _
        Array(1, "age", "demographics", "Age", "choice_age", "0-4 | 5-14", "core"), _
        Array(2, "fever", "symptoms", "Fever", "choice_fever", "yes | no", "core"), _
        Array(3, "age_grp", "demographics", "Age group", "choice_age", "0-4 | 5-14", "optional") _
    ))
    Set translationTable = PrepareTranslationTable()

    Set targetBook = Exporter.BuildDiseaseWorkbook(diseaseWksh, translationTable, _
                                                diseaseWksh.Name, diseaseWksh.Cells(2, 2).Value, "ALPHA_CODE")

    choicesValues = targetBook.Worksheets("Choices").Range("A2").Resize(5, 6).Value

    Assert.AreEqual "choice_age", choicesValues(1, 1), "The first choice of the disease should be exported"
    Assert.AreEqual "choice_fever", choicesValues(3, 1), "Every distinct choice of the disease should be exported"
    Assert.AreEqual "yes", choicesValues(3, 3), "The choice values should travel with their list"
    Assert.AreEqual 1, choicesValues(3, 2), "The ordering restarts on every list"
    Assert.IsTrue IsEmpty(choicesValues(5, 1)), "A choice used twice should be exported once"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestChoicesExportCoversEveryChoiceOnce", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseExporter")
Public Sub TestBuildMigrationWorkbookAggregatesDiseases()
    CustomTestSetTitles Assert, "DiseaseExporter", "TestBuildMigrationWorkbookAggregatesDiseases"

    Dim diseaseNames As BetterArray
    Dim targetBook As Workbook
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

    Set targetBook = Exporter.BuildMigrationWorkbook(sourceWorkbook, diseaseNames)
    Set diseasesSheet = targetBook.Worksheets("Diseases")
    Set translationsSheet = targetBook.Worksheets("Translations")

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
    'A failure before the source workbook exists leaves it Nothing;
    'the guard keeps the handler from raising over the real error.
    On Error Resume Next
        If Not sourceWorkbook Is Nothing Then sourceWorkbook.Close SaveChanges:=False
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestBuildMigrationWorkbookAggregatesDiseases", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseExporter")
Public Sub TestChoicesExportTakesMasterListsAndTranslates()
    CustomTestSetTitles Assert, "DiseaseExporter", "TestChoicesExportTakesMasterListsAndTranslates"

    Dim diseaseWksh As Worksheet
    Dim translationTable As ListObject
    Dim logger As DiseaseLogger
    Dim targetBook As Workbook
    Dim choicesSheet As Worksheet
    Dim choicesValues As Variant

    On Error GoTo Fail

    'The master lists: choice_age carries three values, the disease reads two of
    'them; choice_missing is on the disease sheet only. The values are picked
    'so that Excel stores none of them as a date, and the translations stay in
    'plain ASCII because the VBE reads this file in the ANSI code page.
    PrepareMasterChoicesSheet Array( _
        Array("choice_age", "", "0-4", "0 to 4"), _
        Array("choice_age", "", "15+", ""), _
        Array("choice_age", "", "65+", "65 and over"), _
        Array("choice_fever", "", "yes", "Y"), _
        Array("choice_fever", "", "no", "N"), _
        Array("choice_sex", "", "M", "M") _
    )
    Set diseaseWksh = PrepareDiseaseWorksheetRows("Alpha", "FRA", "ALPHA_CODE", Array( _
        Array(1, "age", "demographics", "Age", "choice_age", "0-4 | 15+", "core"), _
        Array(2, "fever", "symptoms", "Fever", "choice_fever", "yes | no", "core"), _
        Array(3, "other", "symptoms", "Other", "choice_missing", "a | b", "optional") _
    ))
    Set translationTable = PrepareTranslationTable("FRA", Array( _
        Array("0-4", "de 0 a 4"), _
        Array("yes", "oui"), _
        Array("Y", "O") _
    ))
    Set logger = DiseaseLogger.Create()

    Set targetBook = Exporter.BuildDiseaseWorkbook(diseaseWksh, translationTable, _
                                                diseaseWksh.Name, "FRA", "ALPHA_CODE", logger)

    Set choicesSheet = targetBook.Worksheets("Choices")
    choicesValues = choicesSheet.Range("A2").Resize(8, 6).Value

    'choice_age: the three master rows, in master order.
    Assert.AreEqual "choice_age", choicesValues(1, 1), "The master list travels under its name"
    Assert.AreEqual "0-4", choicesValues(1, 3), "The non translated label keeps the master value"
    Assert.AreEqual "de 0 a 4", choicesValues(1, 4), "The translated label is in the language of the disease"
    Assert.AreEqual "de 0 a 4", choicesValues(1, 5), "The label reads the translated label"
    Assert.AreEqual "0 to 4", choicesValues(1, 6), "A short label with no translation keeps the master value"
    Assert.AreEqual "15+", choicesValues(2, 6), "An empty master short label takes the label"
    Assert.AreEqual "65+", choicesValues(3, 3), "A master value the disease sheet does not reach is exported"
    Assert.AreEqual 3, choicesValues(3, 2), "The ordering follows the master order"

    'choice_fever: the label formula and the short label translation.
    Assert.AreEqual "choice_fever", choicesValues(4, 1), "The second list follows the first"
    Assert.AreEqual "oui", choicesValues(4, 5), "The label formula reads the translated value"
    Assert.AreEqual "O", choicesValues(4, 6), "The short label is translated too"
    Assert.IsTrue Left$(choicesSheet.Cells(5, 5).Formula, 1) = "=", "The label column carries a formula"

    'choice_missing: the disease values, and one warning.
    Assert.AreEqual "choice_missing", choicesValues(6, 1), "A list the master lacks is still exported"
    Assert.AreEqual "a", choicesValues(6, 3), "A missing list takes the disease values"
    Assert.AreEqual "b", choicesValues(7, 4), "The disease values fill the translated label too"
    Assert.IsTrue IsEmpty(choicesValues(8, 1)), "choice_sex is on the master sheet only and stays out"
    Assert.AreEqual 1, logger.Entries.Length, "One warning for the list the master lacks"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestChoicesExportTakesMasterListsAndTranslates", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

'The master Choices sheet: headers on row 4, column 1, the way the master
'file carries them, with the label column as the formula the workbook owns.
Private Sub PrepareMasterChoicesSheet(ByVal rows As Variant)
    Dim sheet As Worksheet
    Dim data As Variant
    Dim rowCount As Long

    DeleteWorksheet MASTER_CHOICES_SHEET
    Set sheet = EnsureWorksheet(MASTER_CHOICES_SHEET)
    ClearWorksheet sheet

    data = RowsToMatrix(rows)
    rowCount = UBound(data, 1)

    WriteMatrix sheet.Range("A4"), RowsToMatrix(Array(Array("list name", "translated label", "label", "short label")))
    WriteMatrix sheet.Range("A5"), data
    sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=sheet.Range("A4").Resize(rowCount + 1, 4), _
                          XlListObjectHasHeaders:=xlYes
End Sub

Private Function PrepareDiseaseWorksheet(ByVal diseaseName As String, _
                                         ByVal languageTag As String, _
                                         ByVal diseaseCode As String) As Worksheet

    Set PrepareDiseaseWorksheet = PrepareDiseaseWorksheetRows(diseaseName, languageTag, diseaseCode, Array( _
        Array(1, "age", "demographics", "Age", "choice_age", "0-4 | 5-14 | 15+", "core"), _
        Array(2, "fever", "symptoms", "Fever", "choice_fever", "yes | no", "core") _
    ))
End Function

Private Function PrepareDiseaseWorksheetRows(ByVal diseaseName As String, _
                                             ByVal languageTag As String, _
                                             ByVal diseaseCode As String, _
                                             ByVal rows As Variant) As Worksheet

    Dim sheet As Worksheet
    Dim header As Variant
    Dim dataRows As Variant
    Dim startRange As Range
    Dim listRange As Range

    DeleteWorksheet diseaseName
    Set sheet = EnsureWorksheet(diseaseName)
    ClearWorksheet sheet

    sheet.Cells(2, 2).Value = languageTag
    SeedDiseaseTags sheet, languageTag, diseaseCode

    header = RowsToMatrix(Array(Array("Variable Order", "Variable Name", "Variable Section", "Main Label", "Choice", "Choice Values", "Status")))
    dataRows = RowsToMatrix(rows)

    Set startRange = sheet.Range("B4")
    WriteMatrix startRange, header
    WriteMatrix startRange.Offset(1), dataRows

    Set listRange = sheet.Range("B4").Resize(UBound(dataRows, 1) + 1, UBound(dataRows, 2))
    sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=listRange, XlListObjectHasHeaders:=xlYes

    Set PrepareDiseaseWorksheetRows = sheet
End Function

'A translations table of one language column; the default rows carry no
'choice value, so the choices of the other tests travel untranslated.
Private Function PrepareTranslationTable(Optional ByVal languageTag As String = "ENG", _
                                         Optional ByVal rows As Variant) As ListObject
    Dim sheet As Worksheet
    Dim header As Variant
    Dim dataRows As Variant
    Dim listRange As Range

    DeleteWorksheet TRANSLATION_SHEET
    Set sheet = EnsureWorksheet(TRANSLATION_SHEET)
    ClearWorksheet sheet

    If IsMissing(rows) Then
        rows = Array( _
            Array("hello", "Hello"), _
            Array("world", "World") _
        )
    End If

    header = RowsToMatrix(Array(Array("tag", languageTag)))
    dataRows = RowsToMatrix(rows)

    WriteMatrix sheet.Range("A1"), header
    WriteMatrix sheet.Range("A2"), dataRows

    Set listRange = sheet.Range("A1").Resize(UBound(dataRows, 1) + 1, UBound(dataRows, 2))
    sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=listRange, XlListObjectHasHeaders:=xlYes

    Set PrepareTranslationTable = sheet.ListObjects(1)
End Function

'The tags of a disease sheet live in its hidden names; the fixtures seed
'them the way DiseaseSheet writes them.
Private Sub SeedDiseaseTags(ByVal sheet As Worksheet, ByVal languageTag As String, ByVal diseaseCode As String)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(sheet)
    store.EnsureName "sheetTag", "disease", HiddenNameTypeString
    store.EnsureName "__Var_DISLANG", languageTag, HiddenNameTypeString
    store.EnsureName "__Var_DISCODE", diseaseCode, HiddenNameTypeString
End Sub

Private Sub DeleteFixtureSheets()
    DeleteWorksheet TRANSLATION_SHEET
    DeleteWorksheet MASTER_CHOICES_SHEET
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
        Array("Variable Order", "Variable Name", "Variable Section", "Main Label", "Choice", "Choice Values", "Status"), _
        Array(1, "age", "demographics", "Age", "choice_age", "0-4 | 5-14", "core"), _
        Array(2, "fever", "symptoms", "Fever", "choice_fever", "yes | no", "core") _
    )), "ENG", "BETA_CODE"

    PopulateDiseaseSheet diseaseGamma, RowsToMatrix(Array( _
        Array("Variable Order", "Variable Name", "Variable Section", "Main Label", "Choice", "Choice Values", "Status"), _
        Array(1, "travel", "history", "Recent travel", "choice_travel", "yes | no", "core") _
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
    SeedDiseaseTags sheet, languageTag, diseaseCode

    rowCount = UBound(data, 1)
    columnCount = UBound(data, 2)

    sheet.Range("B4").Resize(rowCount, columnCount).Value = data
    Set rangeObj = sheet.Range("B4").Resize(rowCount, columnCount)
    sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=rangeObj, XlListObjectHasHeaders:=xlYes
End Sub
