Attribute VB_Name = "TestMasterSetupMigration"
Attribute VB_Description = "Tests proving the whole master setup moves to a fresh file through one migration workbook"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests proving the whole master setup moves to a fresh file through one migration workbook")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const VARIABLES_SHEET As String = "Variables"
Private Const CHOICES_SHEET As String = "Choices"
Private Const TRANSLATIONS_SHEET As String = "Translations"
Private Const DROPDOWN_SHEET As String = "__dropdowns"
Private Const STAGING_SHEET As String = "__mig_import"
Private Const DISEASE_ALPHA As String = "Alpha"
Private Const DISEASE_BETA As String = "Beta"
Private Const LANGUAGES_LIST As String = "__data_languages"
Private Const STATUS_LIST As String = "__var_status"
Private Const CHOICES_LIST As String = "__lst_choices"
Private Const PROHIBITED_LIST As String = "__prohibited_diseases_list"
Private Const DISEASES_LIST As String = "__diseases_list"
Private Const VARIABLE_NAME_RANGE As String = "__Col__Variables"
Private Const MARKER_NAME_PREFIX As String = "DISSHEET"
Private Const CHOICES_HEADER_ROW As Long = 4
'A block is seven columns wide and the next one opens two columns after it.
Private Const BLOCK_STRIDE As Long = 9

Private Assert As CustomTest
Private Builder As DiseaseSheet
Private Dropdowns As DropdownLists
Private VariablesManager As MasterSetupVariables
Private VariablesTable As ListObject
Private MasterChoices As LLChoices
Private TranslationTable As ListObject
Private Service As MasterSetupImportService
Private Migration As MasterSetupMigration
Private MigrationBook As Workbook

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestMasterSetupMigration"
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
    Set Dropdowns = Nothing
    Set VariablesManager = Nothing
    Set VariablesTable = Nothing
    Set MasterChoices = Nothing
    Set TranslationTable = Nothing
    Set Service = Nothing
    Set Migration = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    CleanupEnvironment
    PrepareEnvironment
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'The assertions of a test reach the results sheet only once flushed.
    Assert.Flush
    CleanupEnvironment
End Sub

'@section Tests
'===============================================================================

'@TestMethod("MasterSetupMigration")
Public Sub TestExportLandsFiveSheetsAndTheBlocks()
    CustomTestSetTitles Assert, "MasterSetupMigration", "TestExportLandsFiveSheetsAndTheBlocks"

    Dim diseasesSheet As Worksheet
    Dim store As HiddenNames

    On Error GoTo Fail

    BuildTwoDiseases

    Set MigrationBook = Migration.BuildMigrationWorkbook()

    Assert.AreEqual 5, MigrationBook.Worksheets.Count, "The migration file carries five sheets"
    Assert.IsTrue SheetExists(MigrationBook, "Metadata"), "The Metadata sheet is written"
    Assert.IsTrue SheetExists(MigrationBook, "Diseases"), "The Diseases sheet is written"
    Assert.IsTrue SheetExists(MigrationBook, "Variables"), "The Variables sheet is written"
    Assert.IsTrue SheetExists(MigrationBook, "Choices"), "The Choices sheet is written"
    Assert.IsTrue SheetExists(MigrationBook, "Translations"), "The Translations sheet is written"

    Set diseasesSheet = MigrationBook.Worksheets("Diseases")
    Assert.AreEqual "Disease", diseasesSheet.Cells(1, 1).Value, "The first block opens at column 1"
    Assert.AreEqual DISEASE_ALPHA, diseasesSheet.Cells(2, 1).Value, "The first block names its disease"
    Assert.AreEqual "ENG", diseasesSheet.Cells(2, 2).Value, "The first block carries its language"
    Assert.AreEqual "Variable Order", diseasesSheet.Cells(3, 1).Value, "The table headers sit on row 3"
    Assert.AreEqual "var_a", diseasesSheet.Cells(4, 2).Value, "The first line follows the headers"
    Assert.AreEqual "Disease", diseasesSheet.Cells(1, 1 + BLOCK_STRIDE).Value, "The second block opens at column 10"
    Assert.AreEqual DISEASE_BETA, diseasesSheet.Cells(2, 1 + BLOCK_STRIDE).Value, "The second block names its disease"
    Assert.AreEqual "FRA", diseasesSheet.Cells(2, 2 + BLOCK_STRIDE).Value, "The second block carries its language"
    Assert.IsTrue IsEmpty(diseasesSheet.Cells(1, BLOCK_STRIDE - 1).Value), "One empty column separates the blocks"
    Assert.AreEqual "Consolas", diseasesSheet.Cells(3, 1).Font.Name, "The Diseases sheet takes the export format"

    Assert.AreEqual "Tab_Variables", MigrationBook.Worksheets("Variables").ListObjects(1).Name, "The Variables block is the Tab_Variables table"
    Assert.AreEqual "Variable Order", MigrationBook.Worksheets("Variables").Cells(1, 1).Value, "The Variables table keeps the master headers"
    Assert.AreEqual 2, MigrationBook.Worksheets("Variables").ListObjects(1).ListRows.Count, "Every Variables line travels"
    Assert.AreEqual "Tab_Choices", MigrationBook.Worksheets("Choices").ListObjects(1).Name, "The Choices block is the Tab_Choices table"
    Assert.AreEqual "list name", MigrationBook.Worksheets("Choices").Cells(1, 1).Value, "The Choices headers are written from row 1"
    Assert.AreEqual 4, MigrationBook.Worksheets("Choices").ListObjects(1).ListRows.Count, "Every Choices line travels"
    Assert.AreEqual "Tab_Translations", MigrationBook.Worksheets("Translations").ListObjects(1).Name, "The Translations block is the Tab_Translations table"
    Assert.AreEqual "ENG", MigrationBook.Worksheets("Translations").Cells(1, 1).Value, "The translations keep their language headers"

    Assert.AreEqual "2", MetadataValue("disease_count"), "The Metadata sheet counts the diseases"
    Assert.AreEqual "ENG;FRA", MetadataValue("languages"), "The Metadata sheet lists the languages"
    Set store = HiddenNames.Create(MigrationBook)
    Assert.AreEqual "1", store.ValueAsString("__Var_MIGRATION"), "The workbook is tagged as a migration file"
    Assert.AreEqual 2, store.ValueAsLong("__Var_DISCOUNT"), "The workbook tag carries the disease count"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportLandsFiveSheetsAndTheBlocks", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupMigration")
Public Sub TestRoundTripRestoresTheMasterSetup()
    CustomTestSetTitles Assert, "MasterSetupMigration", "TestRoundTripRestoresTheMasterSetup"

    Dim logger As DiseaseLogger
    Dim rebuiltTable As ListObject

    On Error GoTo Fail

    BuildTwoDiseases
    Set MigrationBook = Migration.BuildMigrationWorkbook()

    'The target becomes a fresh file: no disease, no line, and one language less.
    WipeTarget removeLanguage:=True
    Set logger = DiseaseLogger.Create()

    Migration.Import MigrationBook, logger

    Assert.AreEqual 2, Migration.ImportedDiseaseCount, "Both disease blocks come back"
    Assert.IsTrue WorksheetExists(DISEASE_ALPHA), "The first disease worksheet is rebuilt"
    Assert.IsTrue WorksheetExists(DISEASE_BETA), "The second disease worksheet is rebuilt"
    Assert.AreEqual "ENG", ThisWorkbook.Worksheets(DISEASE_ALPHA).Cells(2, 2).Value, "The first disease keeps its language"
    Assert.AreEqual "FRA", ThisWorkbook.Worksheets(DISEASE_BETA).Cells(2, 2).Value, "The second disease keeps the language the file added"

    Set rebuiltTable = ThisWorkbook.Worksheets(DISEASE_ALPHA).ListObjects(1)
    Assert.AreEqual 2, rebuiltTable.ListRows.Count, "Every line of the block lands"
    Assert.AreEqual "var_a", rebuiltTable.DataBodyRange.Cells(1, 2).Value, "The first variable lands on the first line"
    Assert.AreEqual "choice_fever", rebuiltTable.DataBodyRange.Cells(2, 5).Value, "The choice travels with the variable"
    Assert.IsTrue Left$(rebuiltTable.DataBodyRange.Cells(1, 4).Formula, 1) = "=", "The label column carries its formula again"
    Assert.IsTrue Left$(rebuiltTable.DataBodyRange.Cells(2, 6).Formula, 1) = "=", "The choice values column carries its formula again"

    Assert.IsTrue VariablesManager.HasVariable("var_a"), "The Variables table takes its first line back"
    Assert.IsTrue VariablesManager.HasVariable("var_b"), "The Variables table takes its second line back"
    Assert.AreEqual "symptoms", VariablesManager.SectionFor("var_b"), "The section travels with the variable"

    Assert.IsTrue MasterChoices.ChoiceExists("choice_age"), "The Choices sheet takes its first list back"
    Assert.IsTrue MasterChoices.ChoiceExists("choice_fever"), "The Choices sheet takes its second list back"
    Assert.AreEqual "5 to 14", ChoiceLabel("choice_age", 2), "The labels travel as they are"
    Assert.IsTrue ChoicesTableReaches("choice_fever"), "The choices table stretches over the imported rows"

    Assert.IsTrue TableHasColumn(TranslationTable, "FRA"), "The language the target lacked is added"
    Assert.AreEqual 1, Migration.AddedLanguages.Length, "One language counts as added"
    Assert.AreEqual "FRA", Migration.AddedLanguages.Item(1), "The added language is named"
    Assert.AreEqual "Age FR", TranslationValue("Age", "FRA"), "The translation rows travel with the language"
    Assert.AreEqual "Fievre", TranslationValue("Fever", "FRA"), "Every translation row travels"

    Assert.IsTrue Dropdowns.Values(LANGUAGES_LIST).Includes("FRA"), "The languages list carries the added language"
    Assert.IsTrue Dropdowns.Values(CHOICES_LIST).Includes("choice_fever"), "The choices list carries the imported lists"
    Assert.IsFalse WorksheetExists(STAGING_SHEET), "The staging sheet goes away"
    Assert.IsTrue logger.HasEntries, "The steps are logged"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRoundTripRestoresTheMasterSetup", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupMigration")
Public Sub TestImportRefusesATargetThatCarriesData()
    CustomTestSetTitles Assert, "MasterSetupMigration", "TestImportRefusesATargetThatCarriesData"

    Dim manager As DiseaseWorksheetManager
    Dim errNumber As Long
    Dim errText As String

    On Error GoTo Fail

    BuildTwoDiseases
    Set MigrationBook = Migration.BuildMigrationWorkbook()
    Set manager = DiseaseWorksheetManager.Create()

    'A disease worksheet is still there.
    On Error Resume Next
        Migration.Import MigrationBook
        errNumber = Err.Number
        errText = Err.Description
    On Error GoTo Fail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, "A target carrying a disease is refused"
    Assert.IsTrue InStr(Migration.LastRefusal, DISEASE_ALPHA) > 0, _
                  "The refusal names the disease it found. Actual: " & Migration.LastRefusal

    'The diseases are gone; the Variables lines are still there.
    manager.RemoveWorksheet ThisWorkbook, DISEASE_ALPHA
    manager.RemoveWorksheet ThisWorkbook, DISEASE_BETA
    errNumber = 0
    errText = vbNullString
    On Error Resume Next
        Migration.Import MigrationBook
        errNumber = Err.Number
        errText = Err.Description
    On Error GoTo Fail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, "A target carrying Variables lines is refused"
    Assert.IsTrue InStr(Migration.LastRefusal, "Variables") > 0, _
                  "The refusal names the Variables table. Actual: " & Migration.LastRefusal
    Assert.IsFalse WorksheetExists(DISEASE_ALPHA), "A refused import writes no disease worksheet"

    'The Variables lines are gone; the Choices lines are still there.
    VariablesTable.DataBodyRange.ClearContents
    errNumber = 0
    errText = vbNullString
    On Error Resume Next
        Migration.Import MigrationBook
        errNumber = Err.Number
        errText = Err.Description
    On Error GoTo Fail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, "A target carrying Choices lines is refused"
    Assert.IsTrue InStr(Migration.LastRefusal, "Choices") > 0, _
                  "The refusal names the Choices sheet. Actual: " & Migration.LastRefusal
    Assert.IsFalse WorksheetExists(DISEASE_ALPHA), "A refused import still writes no disease worksheet"
    Assert.IsFalse VariablesManager.HasVariable("var_a"), "A refused import writes no Variables line"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportRefusesATargetThatCarriesData", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupMigration")
Public Sub TestBlockWalkSurvivesAnEmptyBlock()
    CustomTestSetTitles Assert, "MasterSetupMigration", "TestBlockWalkSurvivesAnEmptyBlock"

    Dim diseasesSheet As Worksheet
    Dim lastRow As Long
    Dim blockRange As Range

    On Error GoTo Fail

    BuildTwoDiseases
    Set MigrationBook = Migration.BuildMigrationWorkbook()

    'The second block moves one stride further, leaving an empty block
    'between the two.
    Set diseasesSheet = MigrationBook.Worksheets("Diseases")
    lastRow = diseasesSheet.UsedRange.Rows.Count
    Set blockRange = diseasesSheet.Cells(1, 1 + BLOCK_STRIDE).Resize(lastRow, BLOCK_STRIDE - 2)
    blockRange.Cut Destination:=diseasesSheet.Cells(1, 1 + 2 * BLOCK_STRIDE)

    Assert.IsTrue IsEmpty(diseasesSheet.Cells(1, 1 + BLOCK_STRIDE).Value), "The middle block is empty"
    Assert.AreEqual DISEASE_BETA, diseasesSheet.Cells(2, 1 + 2 * BLOCK_STRIDE).Value, "The second block sits at column 19"

    WipeTarget removeLanguage:=False

    Migration.Import MigrationBook

    Assert.AreEqual 2, Migration.ImportedDiseaseCount, "The walk finds both blocks across the gap"
    Assert.IsTrue WorksheetExists(DISEASE_ALPHA), "The first disease comes back"
    Assert.IsTrue WorksheetExists(DISEASE_BETA), "The disease past the gap comes back"
    Assert.AreEqual "var_c", ThisWorkbook.Worksheets(DISEASE_BETA).ListObjects(1).DataBodyRange.Cells(1, 2).Value, _
                    "The block past the gap lands whole"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestBlockWalkSurvivesAnEmptyBlock", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

Private Sub PrepareEnvironment()
    Dim variablesSheet As Worksheet
    Dim dropdownSheet As Worksheet
    Dim translationSheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim data As Variant

    'Variables sheet: the master table, in the eight columns the master carries.
    Set variablesSheet = EnsureWorksheet(VARIABLES_SHEET)
    ClearWorksheet variablesSheet
    data = RowsToMatrix(Array( _
        Array("Variable Order", "Variable Section", "Variable Name", "Label", "Default Choice", "Choices Values", "Default Status", "Comments"), _
        Array(1, "demographics", "var_a", "Age", "choice_age", "0 to 4 | 5 to 14", "core", ""), _
        Array(2, "symptoms", "var_b", "Fever", "choice_fever", "yes | no", "core", "") _
    ))
    WriteMatrix variablesSheet.Range("A1"), data
    Set VariablesTable = variablesSheet.ListObjects.Add(xlSrcRange, variablesSheet.Range("A1").Resize(3, 8), _
                                                        XlListObjectHasHeaders:=xlYes)
    VariablesTable.Name = "TST_MigrationVariables"

    Set VariablesManager = MasterSetupVariables.Create(VariablesTable)
    RegisterVariableName VariablesTable

    Set dropdownSheet = EnsureWorksheet(DROPDOWN_SHEET)
    ClearWorksheet dropdownSheet

    Set Dropdowns = DropdownLists.Create(dropdownSheet)
    AddDropdownList Dropdowns, LANGUAGES_LIST, Array("ENG", "FRA")
    AddDropdownList Dropdowns, STATUS_LIST, Array("core", "optional", "hidden")
    AddDropdownList Dropdowns, CHOICES_LIST, Array("choice_age", "choice_fever")
    AddDropdownList Dropdowns, PROHIBITED_LIST, Array("Variables", "Translations", "Choices")
    AddDropdownList Dropdowns, DISEASES_LIST, Array("", "")

    'Translations sheet: the table sits one column right of its helper tag
    'column, the way SetupTranslationsTable wants it; the headers are the
    'languages.
    Set translationSheet = EnsureWorksheet(TRANSLATIONS_SHEET)
    ClearWorksheet translationSheet
    translationSheet.Range("A1").Value = "__TagInternal__"
    data = RowsToMatrix(Array(Array("ENG", "FRA"), Array("Age", "Age FR"), Array("Fever", "Fievre")))
    WriteMatrix translationSheet.Range("B1"), data
    Set TranslationTable = translationSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                            Source:=translationSheet.Range("B1").Resize(3, 2), _
                                                            XlListObjectHasHeaders:=xlYes)
    TranslationTable.Name = "Tab_Translations"

    'Choices sheet: headers on row 4, the way the master file carries them.
    'The values are picked so that Excel stores none of them as a date.
    Set choicesSheet = EnsureWorksheet(CHOICES_SHEET)
    ClearWorksheet choicesSheet
    data = RowsToMatrix(Array( _
        Array("list name", "ordering list", "translated label", "label", "short label"), _
        Array("choice_age", 1, "", "0 to 4", "0 to 4"), _
        Array("choice_age", 2, "", "5 to 14", "5 to 14"), _
        Array("choice_fever", 1, "", "yes", "yes"), _
        Array("choice_fever", 2, "", "no", "no") _
    ))
    WriteMatrix choicesSheet.Cells(CHOICES_HEADER_ROW, 1), data
    choicesSheet.ListObjects.Add SourceType:=xlSrcRange, _
                                 Source:=choicesSheet.Cells(CHOICES_HEADER_ROW, 1).Resize(5, 5), _
                                 XlListObjectHasHeaders:=xlYes
    Set MasterChoices = LLChoices.Create(choicesSheet, CHOICES_HEADER_ROW, 1)

    Set Builder = DiseaseSheet.Create(ThisWorkbook, Dropdowns, VariablesManager)
    Set Service = MasterSetupImportService.Create(ThisWorkbook, Builder, Dropdowns, VariablesManager, MasterChoices)
    Set Migration = MasterSetupMigration.Create(ThisWorkbook, Dropdowns, VariablesManager, MasterChoices, _
                                                TranslationTable, Builder, Service)
End Sub

Private Sub CleanupEnvironment()
    On Error Resume Next
        If Not Migration Is Nothing Then Migration.ReleaseWorkbook
        If Not MigrationBook Is Nothing Then MigrationBook.Close SaveChanges:=False
    On Error GoTo 0
    Set MigrationBook = Nothing

    DeleteWorksheetSafe DISEASE_ALPHA
    DeleteWorksheetSafe DISEASE_BETA
    DeleteWorksheetSafe DROPDOWN_SHEET
    DeleteWorksheetSafe TRANSLATIONS_SHEET
    DeleteWorksheetSafe CHOICES_SHEET
    DeleteWorksheetSafe STAGING_SHEET
    ClearWorksheetSafe VARIABLES_SHEET

    DeleteNameSafe VARIABLE_NAME_RANGE
    DeleteNameSafe MARKER_NAME_PREFIX & "001"
    DeleteNameSafe MARKER_NAME_PREFIX & "002"
    DeleteNameSafe MARKER_NAME_PREFIX & "003"
    DeleteNameSafe MARKER_NAME_PREFIX & "004"
End Sub

'@description Two disease sheets, one per language, with their lines.
Private Sub BuildTwoDiseases()
    Dim diseaseWksh As Worksheet

    Set diseaseWksh = Builder.Build(DISEASE_ALPHA, "ENG")
    FillDiseaseTable diseaseWksh.ListObjects(1), Array( _
        Array(1, "var_a", "demographics", "Age", "choice_age", "0 to 4 | 5 to 14", "core"), _
        Array(2, "var_b", "symptoms", "Fever", "choice_fever", "yes | no", "core") _
    )

    Set diseaseWksh = Builder.Build(DISEASE_BETA, "FRA")
    FillDiseaseTable diseaseWksh.ListObjects(1), Array( _
        Array(1, "var_c", "history", "Travel", "choice_fever", "yes | no", "optional") _
    )
End Sub

'@description Make the target a fresh file: no disease, no Variables line, no Choices line.
'@param removeLanguage Boolean. When True the FRA column leaves the
'   translations table and the languages list, so the import has one to add.
Private Sub WipeTarget(ByVal removeLanguage As Boolean)
    Dim manager As DiseaseWorksheetManager
    Dim choicesSheet As Worksheet
    Dim choicesTable As ListObject
    Dim lastRow As Long

    Set manager = DiseaseWorksheetManager.Create()
    manager.RemoveWorksheet ThisWorkbook, DISEASE_ALPHA
    manager.RemoveWorksheet ThisWorkbook, DISEASE_BETA

    VariablesTable.DataBodyRange.ClearContents

    Set choicesSheet = ThisWorkbook.Worksheets(CHOICES_SHEET)
    Set choicesTable = choicesSheet.ListObjects(1)
    lastRow = choicesSheet.Cells(choicesSheet.Rows.Count, 1).End(xlUp).Row
    If lastRow > CHOICES_HEADER_ROW Then
        choicesSheet.Range(choicesSheet.Cells(CHOICES_HEADER_ROW + 1, 1), choicesSheet.Cells(lastRow, 5)).ClearContents
    End If
    choicesTable.Resize choicesSheet.Cells(CHOICES_HEADER_ROW, 1).Resize(2, 5)

    If removeLanguage Then
        TranslationTable.ListColumns("FRA").Delete
        AddDropdownList Dropdowns, LANGUAGES_LIST, Array("ENG"), updateExisting:=True
    End If
End Sub

'@description Write the lines of a disease table over the line the builder left.
Private Sub FillDiseaseTable(ByVal table As ListObject, ByVal rows As Variant)
    Dim rowIndex As Long

    For rowIndex = LBound(rows) To UBound(rows)
        If table.ListRows.Count < rowIndex - LBound(rows) + 1 Then table.ListRows.Add
        table.ListRows(rowIndex - LBound(rows) + 1).Range.Value = rows(rowIndex)
    Next rowIndex
End Sub

'@description One value of the Metadata sheet of the migration workbook, read by its label.
Private Function MetadataValue(ByVal metaLabel As String) As String
    Dim metadataSheet As Worksheet
    Dim rowIndex As Long

    Set metadataSheet = MigrationBook.Worksheets("Metadata")
    For rowIndex = 1 To 10
        If StrComp(CStr(metadataSheet.Cells(rowIndex, 1).Value), metaLabel, vbTextCompare) = 0 Then
            MetadataValue = CStr(metadataSheet.Cells(rowIndex, 2).Value)
            Exit Function
        End If
    Next rowIndex

    MetadataValue = "<missing>"
End Function

'@description The label written for the nth row of a list on the master Choices sheet.
Private Function ChoiceLabel(ByVal listName As String, ByVal position As Long) As String
    Dim choicesSheet As Worksheet
    Dim rowIndex As Long
    Dim seen As Long
    Dim lastRow As Long

    Set choicesSheet = ThisWorkbook.Worksheets(CHOICES_SHEET)
    lastRow = choicesSheet.Cells(choicesSheet.Rows.Count, 1).End(xlUp).Row

    For rowIndex = CHOICES_HEADER_ROW + 1 To lastRow
        If StrComp(CStr(choicesSheet.Cells(rowIndex, 1).Value), listName, vbTextCompare) = 0 Then
            seen = seen + 1
            If seen = position Then
                'The fourth column is the label of the fixture.
                ChoiceLabel = CStr(choicesSheet.Cells(rowIndex, 4).Value)
                Exit Function
            End If
        End If
    Next rowIndex

    ChoiceLabel = "<missing>"
End Function

'@description Whether the choices ListObject covers every row naming the list.
Private Function ChoicesTableReaches(ByVal listName As String) As Boolean
    Dim choicesSheet As Worksheet
    Dim table As ListObject
    Dim lastRow As Long
    Dim rowIndex As Long

    Set choicesSheet = ThisWorkbook.Worksheets(CHOICES_SHEET)
    Set table = choicesSheet.ListObjects(1)
    lastRow = choicesSheet.Cells(choicesSheet.Rows.Count, 1).End(xlUp).Row

    For rowIndex = lastRow To CHOICES_HEADER_ROW + 1 Step -1
        If StrComp(CStr(choicesSheet.Cells(rowIndex, 1).Value), listName, vbTextCompare) = 0 Then
            ChoicesTableReaches = Not Intersect(table.Range, choicesSheet.Cells(rowIndex, 1)) Is Nothing
            Exit Function
        End If
    Next rowIndex
End Function

'@description The value of one translation row in one language column.
Private Function TranslationValue(ByVal label As String, ByVal languageTag As String) As String
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim block As Variant

    block = TranslationTable.Range.Value
    For columnIndex = 1 To UBound(block, 2)
        If StrComp(CStr(block(1, columnIndex)), languageTag, vbTextCompare) <> 0 Then GoTo NextColumn
        For rowIndex = 2 To UBound(block, 1)
            If StrComp(CStr(block(rowIndex, 1)), label, vbBinaryCompare) = 0 Then
                TranslationValue = CStr(block(rowIndex, columnIndex))
                Exit Function
            End If
        Next rowIndex
NextColumn:
    Next columnIndex

    TranslationValue = "<missing>"
End Function

Private Function TableHasColumn(ByVal table As ListObject, ByVal headerName As String) As Boolean
    Dim column As ListColumn

    For Each column In table.ListColumns
        If StrComp(column.Name, headerName, vbTextCompare) = 0 Then
            TableHasColumn = True
            Exit Function
        End If
    Next column
End Function

Private Function SheetExists(ByVal targetBook As Workbook, ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet

    For Each sheet In targetBook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next sheet
End Function

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    WorksheetExists = SheetExists(ThisWorkbook, sheetName)
End Function

Private Sub DeleteWorksheetSafe(ByVal sheetName As String)
    On Error Resume Next
        ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
End Sub

Private Sub ClearWorksheetSafe(ByVal sheetName As String)
    Dim sh As Worksheet

    On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If Not sh Is Nothing Then
        ClearWorksheet sh
    End If
End Sub

Private Sub DeleteNameSafe(ByVal nameValue As String)
    On Error Resume Next
        ThisWorkbook.Names(nameValue).Delete
    On Error GoTo 0
End Sub

Private Sub RegisterVariableName(ByVal lo As ListObject)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(ThisWorkbook)
    store.SetListObjectHeader VARIABLE_NAME_RANGE, lo, "Variable Name"
End Sub

Private Sub AddDropdownList(ByVal target As DropdownLists, ByVal listName As String, ByVal values As Variant, _
                            Optional ByVal updateExisting As Boolean = False)
    Dim listValues As BetterArray

    Set listValues = BuildBetterArray(values)
    If listValues Is Nothing Then Exit Sub

    If updateExisting And target.Exists(listName) Then
        target.Update listValues, listName
    Else
        target.Add listValues, listName
    End If
End Sub

Private Function BuildBetterArray(ByVal values As Variant) As BetterArray
    Dim arr As BetterArray
    Dim idx As Long

    If Not IsArray(values) Then Exit Function

    Set arr = New BetterArray
    arr.LowerBound = 1
    For idx = LBound(values) To UBound(values)
        arr.Push CStr(values(idx))
    Next idx

    Set BuildBetterArray = arr
End Function
