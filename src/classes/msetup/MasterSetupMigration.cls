VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MasterSetupMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Writes the whole master setup into one migration file and reads it back into a fresh one"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Writes the whole master setup into one migration file and reads it back into a fresh one")
'@depends DiseaseExporter, DiseaseExportWorkbook, DiseaseSheet, DiseaseWorksheetManager, DiseaseLogger, MasterSetupImportService, MasterSetupPreparation, MasterSetupVariables, SetupTranslationsTable, LLChoices, DropdownLists, HiddenNames, BetterArray
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'MasterSetupMigration.cls without interface -- single implementation, no test
'fake, no compile cycle.
'
'The migration is the whole configuration of a master setup file, as it
'stands, for the next version of that file. One class owns the format at
'both ends, so the writer and the reader agree on it.
'
'THE FILE
'-------------------------------------------------------------------------------
'One flat workbook, master_setup_migration_<yyyy-mm-dd>.xlsx, five sheets:
'- Metadata: a label and value block (export_date, source_file,
'  disease_count, languages), plus the workbook hidden names __Var_MIGRATION
'  and __Var_DISCOUNT, which is how the import knows a migration file.
'- Diseases: one block per disease worksheet, the first at column 1, the
'  next one 7 + 2 columns further. Row 1 reads Disease | Language | Tag,
'  row 2 carries the values, row 3 the seven headers of the disease table
'  and the lines follow. Values only: the two formula columns are the
'  master's own and come back on import.
'- Variables: the master table whole, wrapped in Tab_Variables.
'- Choices: the master Choices block whole, wrapped in Tab_Choices.
'- Translations: the master table whole, wrapped in Tab_Translations.
'
'THE IMPORT GOES INTO A FRESH FILE
'-------------------------------------------------------------------------------
'A target holding a disease worksheet, a Variables line or a Choices line is
'refused before anything is written. The order of the import follows what
'each step feeds: the languages and the translations, the Variables table,
'the Choices block, the dropdowns, then every disease block through
'MasterSetupImportService.ImportDiseaseTable with the formulas restored.

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "MasterSetupMigration"

Private Const METADATA_SHEET As String = "Metadata"
Private Const DISEASES_SHEET As String = "Diseases"
Private Const VARIABLES_SHEET As String = "Variables"
Private Const CHOICES_SHEET As String = "Choices"
Private Const TRANSLATIONS_SHEET As String = "Translations"

Private Const TABLE_VARIABLES As String = "Tab_Variables"
Private Const TABLE_CHOICES As String = "Tab_Choices"
Private Const TABLE_TRANSLATIONS As String = "Tab_Translations"

Private Const META_EXPORT_DATE As String = "export_date"
Private Const META_SOURCE_FILE As String = "source_file"
Private Const META_DISEASE_COUNT As String = "disease_count"
Private Const META_LANGUAGES As String = "languages"
Private Const LANGUAGE_SEPARATOR As String = ";"

'The workbook tags of a migration file.
Private Const NAME_MIGRATION As String = "__Var_MIGRATION"
Private Const NAME_DISCOUNT As String = "__Var_DISCOUNT"
Private Const MIGRATION_TAG_VALUE As String = "1"

'The tags of a disease worksheet, as DiseaseSheet writes them.
Private Const NAME_DISLANG As String = "__Var_DISLANG"
Private Const NAME_DISCODE As String = "__Var_DISCODE"
Private Const SHEET_TAG_NAME As String = "sheetTag"
Private Const DISEASE_TAG_VALUE As String = "disease"

'The block layout of the Diseases sheet.
Private Const BLOCK_HEADER_ROW As Long = 1
Private Const BLOCK_META_ROW As Long = 2
Private Const BLOCK_TABLE_ROW As Long = 3
Private Const BLOCK_HEADER_DISEASE As String = "Disease"
Private Const BLOCK_HEADER_LANGUAGE As String = "Language"
Private Const BLOCK_HEADER_TAG As String = "Tag"
Private Const DISEASE_TABLE_WIDTH As Long = 7
Private Const BLOCK_GAP As Long = 2
'The variable name is the second column of a disease table.
Private Const BLOCK_NAME_OFFSET As Long = 1

Private Const HEADER_VARIABLE_NAME As String = "Variable Name"
Private Const STAGING_SHEET As String = "__mig_import"
Private Const FILE_PREFIX As String = "master_setup_migration_"
Private Const FILE_EXTENSION As String = ".xlsx"
Private Const LOG_OPERATION As String = "Migration"

'@section State
'===============================================================================

Private Type TMigration
    hostBook As Workbook
    dropdowns As DropdownLists
    variables As MasterSetupVariables
    choices As LLChoices
    translationsTable As ListObject
    builder As DiseaseSheet
    importService As MasterSetupImportService
    manager As DiseaseExportWorkbook
    importedDiseases As BetterArray
    addedLanguages As BetterArray
End Type

Private this As TMigration

'@section Factory
'===============================================================================

'@sub-title Build a migration over the master setup workbook and its managers.
'@param hostBook Workbook. The master setup workbook: the source of an
'   export and the target of an import.
'@param dropdowns DropdownLists. The master dropdowns.
'@param variables MasterSetupVariables. The manager of the Variables table.
'@param choices LLChoices. The master Choices block, headers on row 4.
'@param translationsTable ListObject. The master translations table.
'@param builder DiseaseSheet. Builds a disease worksheet on import.
'@param importService MasterSetupImportService. Lands a block on a disease
'   worksheet.
'@return MasterSetupMigration ready to export or import.
Public Function Create(ByVal hostBook As Workbook, _
                       ByVal dropdowns As DropdownLists, _
                       ByVal variables As MasterSetupVariables, _
                       ByVal choices As LLChoices, _
                       ByVal translationsTable As ListObject, _
                       ByVal builder As DiseaseSheet, _
                       ByVal importService As MasterSetupImportService) As MasterSetupMigration

    Dim migration As MasterSetupMigration
    Set migration = New MasterSetupMigration
    migration.Initialise hostBook, dropdowns, variables, choices, translationsTable, builder, importService
    Set Create = migration
End Function

'@section Configuration
'===============================================================================

'@sub-title Attach the workbook and the managers the migration reads and writes through.
Public Sub Initialise(ByVal hostBook As Workbook, _
                      ByVal dropdowns As DropdownLists, _
                      ByVal variables As MasterSetupVariables, _
                      ByVal choices As LLChoices, _
                      ByVal translationsTable As ListObject, _
                      ByVal builder As DiseaseSheet, _
                      ByVal importService As MasterSetupImportService)

    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook cannot be Nothing"
    End If
    If dropdowns Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Dropdown lists cannot be Nothing"
    End If
    If variables Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Variables manager cannot be Nothing"
    End If
    If choices Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Choices manager cannot be Nothing"
    End If
    If translationsTable Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Translations table cannot be Nothing"
    End If
    If builder Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Disease sheet builder cannot be Nothing"
    End If
    If importService Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Import service cannot be Nothing"
    End If

    Set this.hostBook = hostBook
    Set this.dropdowns = dropdowns
    Set this.variables = variables
    Set this.choices = choices
    Set this.translationsTable = translationsTable
    Set this.builder = builder
    Set this.importService = importService
    Set this.manager = DiseaseExportWorkbook.Create()

    ResetTallies
End Sub

'@section Export
'===============================================================================

'@sub-title Build the migration workbook and answer it open.
'@details The five sheets are written in order and take the export format
'DiseaseExporter gives a disease export. The workbook stays open and belongs
'to this class until ExportMigration saves it or the class is released.
'@return Workbook. The migration workbook, open.
Public Function BuildMigrationWorkbook() As Workbook
    Dim targetBook As Workbook
    Dim diseaseSheets As Collection
    Dim metadataSheet As Worksheet
    Dim diseasesSheet As Worksheet
    Dim variablesSheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim translationsSheet As Worksheet

    Set diseaseSheets = CollectDiseaseSheets()

    If this.manager.HasWorkbook Then this.manager.ReleaseWorkbook
    Set targetBook = this.manager.EnsureWorkbook(Application)

    Set metadataSheet = EnsureWorksheet(targetBook, METADATA_SHEET, targetBook.Worksheets(1))
    Set diseasesSheet = EnsureWorksheet(targetBook, DISEASES_SHEET)
    Set variablesSheet = EnsureWorksheet(targetBook, VARIABLES_SHEET)
    Set choicesSheet = EnsureWorksheet(targetBook, CHOICES_SHEET)
    Set translationsSheet = EnsureWorksheet(targetBook, TRANSLATIONS_SHEET)

    WriteMetadata metadataSheet, diseaseSheets.Count
    WriteWorkbookTags targetBook, diseaseSheets.Count
    WriteDiseaseBlocks diseasesSheet, diseaseSheets
    WriteBlock variablesSheet, this.variables.Table.Range, TABLE_VARIABLES
    WriteBlock choicesSheet, MasterChoicesRange(), TABLE_CHOICES
    WriteBlock translationsSheet, this.translationsTable.Range, TABLE_TRANSLATIONS

    DiseaseExporter.FormatMetadataSheet metadataSheet
    DiseaseExporter.FormatWrittenSheet diseasesSheet
    DiseaseExporter.FormatWrittenSheet variablesSheet
    DiseaseExporter.FormatWrittenSheet choicesSheet
    DiseaseExporter.FormatWrittenSheet translationsSheet

    Set BuildMigrationWorkbook = targetBook
End Function

'@sub-title Write the migration workbook to disk and answer the saved path.
'@param exportFolder String. The folder receiving the file.
'@return String. The path of the saved file.
Public Function ExportMigration(ByVal exportFolder As String) As String
    Dim filePath As String

    If LenB(exportFolder) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Export folder cannot be empty"
    End If

    BuildMigrationWorkbook

    filePath = exportFolder & Application.PathSeparator & _
               FILE_PREFIX & Format$(Now, "yyyy-mm-dd") & FILE_EXTENSION
    this.manager.SaveAs filePath
    this.manager.CloseWorkbook False

    ExportMigration = filePath
End Function

'@sub-title Close the migration workbook this class still holds, without saving.
Public Sub ReleaseWorkbook()
    If this.manager Is Nothing Then Exit Sub
    If this.manager.HasWorkbook Then this.manager.ReleaseWorkbook
End Sub

'@section Import
'===============================================================================

'@sub-title Read a migration file into the host workbook, which has to be empty.
'@details The file is checked first (its tag and its four sheets), then
'the target: a disease worksheet, a Variables line or a Choices line
'refuses the import with ProjectError.InvalidArgument naming what was
'found, and nothing is written. The steps then run in the order each one
'feeds the next; the staging sheet of the disease blocks goes away on
'every path.
'@param sourceBook Workbook. The migration file, open.
'@param logger Optional DiseaseLogger. Receives one line per step.
Public Sub Import(ByVal sourceBook As Workbook, Optional ByVal logger As DiseaseLogger = Nothing)
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    ValidateSource sourceBook
    RefuseFilledTarget
    ResetTallies

    On Error GoTo Handler

    ImportTranslations sourceBook, logger
    ImportVariables sourceBook, logger
    ImportChoices sourceBook, logger
    RefreshDropdowns
    ImportDiseases sourceBook, logger

Cleanup:
    'Shielded: the staging sheet goes away whatever happened above, and a
    'raise from the removal would come straight back here.
    On Error Resume Next
    RemoveStagingSheet
    On Error GoTo 0

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If
    Exit Sub

Handler:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume Cleanup
End Sub

'@section Results
'===============================================================================

'@sub-title Names of the disease worksheets the last import built.
Public Property Get ImportedDiseases() As BetterArray
    Set ImportedDiseases = this.importedDiseases.Clone()
End Property

'@sub-title How many disease blocks the last import landed.
Public Property Get ImportedDiseaseCount() As Long
    ImportedDiseaseCount = this.importedDiseases.Length
End Property

'@sub-title The language columns the last import added to the translations table.
Public Property Get AddedLanguages() As BetterArray
    Set AddedLanguages = this.addedLanguages.Clone()
End Property

'@section Export Writers
'===============================================================================

'@sub-title The disease worksheets of the host workbook, in sheet order.
Private Function CollectDiseaseSheets() As Collection
    Dim sheets As New Collection
    Dim sheet As Worksheet

    For Each sheet In this.hostBook.Worksheets
        If IsDiseaseSheet(sheet) Then sheets.Add sheet
    Next sheet

    Set CollectDiseaseSheets = sheets
End Function

'@sub-title Write the label and value block of the Metadata sheet.
Private Sub WriteMetadata(ByVal metadataSheet As Worksheet, ByVal diseaseCount As Long)
    Dim metadata(1 To 5, 1 To 2) As Variant

    metadata(1, 1) = "Label"
    metadata(1, 2) = "Value"
    metadata(2, 1) = META_EXPORT_DATE
    metadata(2, 2) = Format$(Now, "yyyy-mm-dd HH:nn")
    metadata(3, 1) = META_SOURCE_FILE
    metadata(3, 2) = this.hostBook.Name
    metadata(4, 1) = META_DISEASE_COUNT
    metadata(4, 2) = diseaseCount
    metadata(5, 1) = META_LANGUAGES
    metadata(5, 2) = JoinedLanguages(HeaderNames(this.translationsTable))

    metadataSheet.Range("A1").Resize(UBound(metadata, 1), UBound(metadata, 2)).Value = metadata
End Sub

'@sub-title Tag the workbook as a migration file, with its disease count.
Private Sub WriteWorkbookTags(ByVal targetBook As Workbook, ByVal diseaseCount As Long)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(targetBook)
    store.EnsureName NAME_MIGRATION, MIGRATION_TAG_VALUE, HiddenNameTypeString
    store.EnsureName NAME_DISCOUNT, diseaseCount, HiddenNameTypeLong
End Sub

'@sub-title Write every disease block, the first at column 1, then one stride further each.
'@details The writer keeps the next column itself. The UsedRange of an
'empty sheet is $A$1, so asking the sheet where it ends puts the first
'block at column 3.
Private Sub WriteDiseaseBlocks(ByVal diseasesSheet As Worksheet, ByVal diseaseSheets As Collection)
    Dim diseaseWksh As Worksheet
    Dim nextColumn As Long

    nextColumn = 1
    For Each diseaseWksh In diseaseSheets
        WriteDiseaseBlock diseasesSheet, diseaseWksh, nextColumn
        nextColumn = nextColumn + DISEASE_TABLE_WIDTH + BLOCK_GAP
    Next diseaseWksh
End Sub

'@sub-title Write one disease block: the three tags, then the table as values.
Private Sub WriteDiseaseBlock(ByVal diseasesSheet As Worksheet, _
                              ByVal diseaseWksh As Worksheet, _
                              ByVal startColumn As Long)
    Dim store As HiddenNames
    Dim header(1 To 1, 1 To 3) As Variant
    Dim meta(1 To 1, 1 To 3) As Variant
    Dim values As Variant

    If diseaseWksh.ListObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, "The disease worksheet '" & diseaseWksh.Name & "' carries no table"
    End If

    Set store = HiddenNames.Create(diseaseWksh)

    header(1, 1) = BLOCK_HEADER_DISEASE
    header(1, 2) = BLOCK_HEADER_LANGUAGE
    header(1, 3) = BLOCK_HEADER_TAG
    meta(1, 1) = diseaseWksh.Name
    meta(1, 2) = store.ValueAsString(NAME_DISLANG)
    meta(1, 3) = store.ValueAsString(NAME_DISCODE)

    diseasesSheet.Cells(BLOCK_HEADER_ROW, startColumn).Resize(1, 3).Value = header
    diseasesSheet.Cells(BLOCK_META_ROW, startColumn).Resize(1, 3).Value = meta

    values = RangeBlock(diseaseWksh.ListObjects(1).Range)
    diseasesSheet.Cells(BLOCK_TABLE_ROW, startColumn).Resize(UBound(values, 1), UBound(values, 2)).Value = values
End Sub

'@sub-title Copy a block of values to A1 of a sheet and wrap it in a plain table.
Private Sub WriteBlock(ByVal targetSheet As Worksheet, ByVal sourceRange As Range, ByVal tableName As String)
    Dim values As Variant
    Dim blockRange As Range
    Dim table As ListObject

    If sourceRange Is Nothing Then Exit Sub

    values = RangeBlock(sourceRange)
    Set blockRange = targetSheet.Range("A1").Resize(UBound(values, 1), UBound(values, 2))
    blockRange.Value = values

    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=blockRange, XlListObjectHasHeaders:=xlYes)
    table.Name = tableName
    table.TableStyle = ""
End Sub

'@sub-title The master Choices block with its headers, as the master carries it.
Private Function MasterChoicesRange() As Range
    Set MasterChoicesRange = this.choices.ChoiceData.DataRange("__all__", includeHeaders:=True)
End Function

'@section Import Checks
'===============================================================================

'@sub-title Refuse a source that is no migration file or lacks one of its sheets.
Private Sub ValidateSource(ByVal sourceBook As Workbook)
    If sourceBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source workbook cannot be Nothing"
    End If

    If StrComp(TagValue(sourceBook, NAME_MIGRATION), MIGRATION_TAG_VALUE, vbTextCompare) <> 0 Then
        ThrowError ProjectError.InvalidArgument, "The workbook '" & sourceBook.Name & _
                   "' is no migration file: it carries no '" & NAME_MIGRATION & "' tag"
    End If

    RequireWorksheet sourceBook, DISEASES_SHEET
    RequireWorksheet sourceBook, VARIABLES_SHEET
    RequireWorksheet sourceBook, CHOICES_SHEET
    RequireWorksheet sourceBook, TRANSLATIONS_SHEET
End Sub

Private Sub RequireWorksheet(ByVal sourceBook As Workbook, ByVal sheetName As String)
    If FindWorksheet(sourceBook, sheetName) Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "The migration file carries no '" & sheetName & "' worksheet"
    End If
End Sub

'@sub-title Refuse a target that already carries data, naming what was found.
Private Sub RefuseFilledTarget()
    Dim diseaseName As String
    Dim lineCount As Long

    diseaseName = FirstDiseaseSheetName()
    If LenB(diseaseName) > 0 Then
        ThrowError ProjectError.InvalidArgument, "The file already carries the disease '" & diseaseName & "'"
    End If

    lineCount = VariablesLineCount()
    If lineCount > 0 Then
        ThrowError ProjectError.InvalidArgument, "The Variables table already carries " & lineCount & " line(s)"
    End If

    lineCount = ChoicesLineCount()
    If lineCount > 0 Then
        ThrowError ProjectError.InvalidArgument, "The Choices sheet already carries " & lineCount & " line(s)"
    End If
End Sub

'@sub-title The first disease worksheet of the host workbook; empty with none.
Private Function FirstDiseaseSheetName() As String
    Dim sheet As Worksheet

    For Each sheet In this.hostBook.Worksheets
        If IsDiseaseSheet(sheet) Then
            FirstDiseaseSheetName = sheet.Name
            Exit Function
        End If
    Next sheet
End Function

'@sub-title The lines of the Variables table carrying a variable name.
Private Function VariablesLineCount() As Long
    Dim nameRange As Range

    Set nameRange = this.variables.DataRange(HEADER_VARIABLE_NAME)
    If nameRange Is Nothing Then Exit Function

    VariablesLineCount = Application.WorksheetFunction.CountA(nameRange)
End Function

'@sub-title The lines under the headers of the master Choices block.
Private Function ChoicesLineCount() As Long
    Dim choicesSheet As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim nameColumn As Long

    Set choicesSheet = this.choices.Wksh
    firstRow = this.choices.StartRow + 1
    nameColumn = this.choices.StartColumn
    lastRow = choicesSheet.Cells(choicesSheet.Rows.Count, nameColumn).End(xlUp).Row
    If lastRow < firstRow Then Exit Function

    ChoicesLineCount = Application.WorksheetFunction.CountA( _
                       choicesSheet.Range(choicesSheet.Cells(firstRow, nameColumn), _
                                          choicesSheet.Cells(lastRow, nameColumn)))
End Function

'@section Import Steps
'===============================================================================

'@sub-title The languages first, through SetupTranslationsTable, then the rows.
'@details A language column the file carries and the target lacks is added
'through EnsureLanguages. The rows are then matched on their label: a row
'the target carries takes the language values of the file, a row it lacks
'is appended whole.
Private Sub ImportTranslations(ByVal sourceBook As Workbook, ByVal logger As DiseaseLogger)
    Dim sourceTable As ListObject
    Dim sourceBlock As Variant
    Dim manager As SetupTranslationsTable
    Dim before As BetterArray
    Dim after As BetterArray
    Dim index As Long

    Set sourceTable = FindTable(sourceBook, TABLE_TRANSLATIONS, TRANSLATIONS_SHEET)
    If sourceTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "The migration file carries no '" & TABLE_TRANSLATIONS & "' table"
    End If

    sourceBlock = RangeBlock(sourceTable.Range)

    Set before = HeaderNames(this.translationsTable)

    Set manager = SetupTranslationsTable.Create(this.translationsTable)
    manager.SetDisplayPrompts False
    manager.EnsureLanguages JoinedLanguages(BlockHeaders(sourceBlock))

    Set after = HeaderNames(this.translationsTable)
    For index = after.LowerBound To after.UpperBound
        If Not ContainsKey(before, CStr(after.Item(index))) Then
            this.addedLanguages.Push CStr(after.Item(index))
            Info logger, "Added the language '" & after.Item(index) & "' to the translations table"
        End If
    Next index

    WriteTranslationRows sourceBlock, logger
End Sub

'@sub-title Write the translation rows of the file onto the target table.
Private Sub WriteTranslationRows(ByRef sourceBlock As Variant, ByVal logger As DiseaseLogger)
    Dim targetBlock As Variant
    Dim rowIndex As Long
    Dim rowKeys As Collection
    Dim columnIndexes As Collection
    Dim label As String
    Dim targetRow As Long
    Dim newRow As ListRow
    Dim written As Long
    Dim appended As Long

    targetBlock = RangeBlock(this.translationsTable.Range)
    Set rowKeys = LabelIndex(targetBlock)
    Set columnIndexes = HeaderIndex(targetBlock)

    For rowIndex = 2 To UBound(sourceBlock, 1)
        label = TextOf(sourceBlock(rowIndex, 1))
        If LenB(Trim$(label)) = 0 Then GoTo NextRow

        targetRow = LookupRow(rowKeys, label)
        If targetRow = 0 Then
            Set newRow = this.translationsTable.ListRows.Add
            targetRow = newRow.Range.Row - this.translationsTable.Range.Row + 1
            this.translationsTable.Range.Cells(targetRow, 1).Value = label
            rowKeys.Add targetRow, LabelKey(label)
            appended = appended + 1
        End If

        WriteLanguageCells sourceBlock, rowIndex, targetRow, columnIndexes
        written = written + 1
NextRow:
    Next rowIndex

    Info logger, written & " translation row(s) read from the file, " & appended & " appended"
End Sub

'@sub-title Write the language cells of one source row on one target row.
Private Sub WriteLanguageCells(ByRef sourceBlock As Variant, _
                               ByVal sourceRow As Long, _
                               ByVal targetRow As Long, _
                               ByVal columnIndexes As Collection)
    Dim columnIndex As Long
    Dim targetColumn As Long
    Dim cellValue As String

    For columnIndex = 2 To UBound(sourceBlock, 2)
        cellValue = TextOf(sourceBlock(sourceRow, columnIndex))
        If LenB(cellValue) = 0 Then GoTo NextColumn

        targetColumn = LookupColumn(columnIndexes, TextOf(sourceBlock(1, columnIndex)))
        If targetColumn = 0 Then GoTo NextColumn

        this.translationsTable.Range.Cells(targetRow, targetColumn).Value = cellValue
NextColumn:
    Next columnIndex
End Sub

'@sub-title The Variables table, through the manager, which matches by header.
Private Sub ImportVariables(ByVal sourceBook As Workbook, ByVal logger As DiseaseLogger)
    Dim sourceSheet As Worksheet

    Set sourceSheet = FindWorksheet(sourceBook, VARIABLES_SHEET)
    this.variables.ImportFromWorksheet sourceSheet

    Info logger, "The Variables table carries " & VariablesLineCount() & " line(s) after the import"
End Sub

'@sub-title The Choices block, onto the headers the target already has.
Private Sub ImportChoices(ByVal sourceBook As Workbook, ByVal logger As DiseaseLogger)
    Dim sourceSheet As Worksheet

    Set sourceSheet = FindWorksheet(sourceBook, CHOICES_SHEET)
    this.choices.Import sourceSheet, 1, 1, clearSheet:=False
    StretchChoicesTable

    Info logger, "The Choices sheet carries " & ChoicesLineCount() & " line(s) after the import"
End Sub

'@sub-title Stretch the choices ListObject over the rows the import wrote under it.
'@details LLChoices writes the sheet and knows nothing of the table; the
'master event handlers read the sheet through the table.
Private Sub StretchChoicesTable()
    Dim choicesSheet As Worksheet
    Dim table As ListObject
    Dim lastRow As Long
    Dim firstColumn As Long
    Dim lastColumn As Long

    Set choicesSheet = this.choices.Wksh
    If choicesSheet.ListObjects.Count = 0 Then Exit Sub
    Set table = choicesSheet.ListObjects(1)

    firstColumn = table.Range.Column
    lastColumn = firstColumn + table.Range.Columns.Count - 1
    lastRow = choicesSheet.Cells(choicesSheet.Rows.Count, this.choices.StartColumn).End(xlUp).Row

    If lastRow <= table.Range.Row + table.Range.Rows.Count - 1 Then Exit Sub

    table.Resize choicesSheet.Range(choicesSheet.Cells(table.Range.Row, firstColumn), _
                                    choicesSheet.Cells(lastRow, lastColumn))
End Sub

'@sub-title Rebuild the dropdowns, so the languages, the lists and the diseases exist.
'@details The preparation reads the languages off the translations table
'headers and the lists off the Choices sheet, which is why it runs after
'those two steps and before any disease sheet is built.
Private Sub RefreshDropdowns()
    Dim preparation As MasterSetupPreparation

    Set preparation = MasterSetupPreparation.Create(this.hostBook)
    preparation.EnsureDropdowns
End Sub

'@sub-title Every Diseases block, in file order.
'@details Row 1 is walked for the word Disease, so the blocks can sit at
'any stride and an empty gap between two blocks is passed over.
Private Sub ImportDiseases(ByVal sourceBook As Workbook, ByVal logger As DiseaseLogger)
    Dim diseasesSheet As Worksheet
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim diseaseName As String
    Dim languageTag As String

    Set diseasesSheet = FindWorksheet(sourceBook, DISEASES_SHEET)
    lastColumn = LastUsedColumn(diseasesSheet)

    For columnIndex = 1 To lastColumn
        If StrComp(Trim$(TextOf(diseasesSheet.Cells(BLOCK_HEADER_ROW, columnIndex).Value)), _
                   BLOCK_HEADER_DISEASE, vbTextCompare) <> 0 Then GoTo NextColumn

        diseaseName = Trim$(TextOf(diseasesSheet.Cells(BLOCK_META_ROW, columnIndex).Value))
        languageTag = Trim$(TextOf(diseasesSheet.Cells(BLOCK_META_ROW, columnIndex + 1).Value))
        If LenB(diseaseName) = 0 Then GoTo NextColumn

        ImportOneBlock diseasesSheet, columnIndex, diseaseName, languageTag, logger
NextColumn:
    Next columnIndex
End Sub

'@sub-title Stage one block on a hidden sheet and land it through the service.
Private Sub ImportOneBlock(ByVal diseasesSheet As Worksheet, _
                           ByVal startColumn As Long, _
                           ByVal diseaseName As String, _
                           ByVal languageTag As String, _
                           ByVal logger As DiseaseLogger)
    Dim stagingTable As ListObject

    Set stagingTable = BuildStagingTable(diseasesSheet, startColumn)

    this.importService.ImportDiseaseTable stagingTable, diseaseName, languageTag, _
                                          restoreFormulas:=True, logger:=logger
    this.importedDiseases.Push diseaseName
    Info logger, "Imported the disease '" & diseaseName & "' in '" & languageTag & "'"

    RemoveStagingSheet
End Sub

'@sub-title Copy one block into the staging sheet and answer its ListObject.
Private Function BuildStagingTable(ByVal diseasesSheet As Worksheet, ByVal startColumn As Long) As ListObject
    Dim stagingSheet As Worksheet
    Dim rowCount As Long
    Dim blockRange As Range
    Dim tableRange As Range

    rowCount = CountBlockRows(diseasesSheet, startColumn)

    RemoveStagingSheet
    Set stagingSheet = this.hostBook.Worksheets.Add(After:=this.hostBook.Worksheets(this.hostBook.Worksheets.Count))
    stagingSheet.Name = STAGING_SHEET
    stagingSheet.Visible = xlSheetHidden

    'Header plus lines, copied as values in one assignment.
    Set blockRange = diseasesSheet.Cells(BLOCK_TABLE_ROW, startColumn).Resize(rowCount + 1, DISEASE_TABLE_WIDTH)
    Set tableRange = stagingSheet.Range("A1").Resize(rowCount + 1, DISEASE_TABLE_WIDTH)
    tableRange.Value = blockRange.Value

    Set BuildStagingTable = stagingSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                         XlListObjectHasHeaders:=xlYes)
End Function

'@sub-title Count the lines of a block, reading its variable name column.
Private Function CountBlockRows(ByVal diseasesSheet As Worksheet, ByVal startColumn As Long) As Long
    Dim rowIndex As Long
    Dim nameColumn As Long

    nameColumn = startColumn + BLOCK_NAME_OFFSET
    rowIndex = BLOCK_TABLE_ROW + 1

    Do While LenB(Trim$(TextOf(diseasesSheet.Cells(rowIndex, nameColumn).Value))) > 0
        CountBlockRows = CountBlockRows + 1
        rowIndex = rowIndex + 1
    Loop
End Function

'@sub-title Drop the staging sheet when the host workbook carries one.
Private Sub RemoveStagingSheet()
    Dim manager As DiseaseWorksheetManager

    Set manager = DiseaseWorksheetManager.Create()
    manager.RemoveWorksheet this.hostBook, STAGING_SHEET
End Sub

'@section Lookups
'===============================================================================

'@sub-title A disease worksheet is known by its hidden sheetTag.
Private Function IsDiseaseSheet(ByVal targetSheet As Worksheet) As Boolean
    Dim store As HiddenNames

    If targetSheet Is Nothing Then Exit Function

    'A sheet with no hidden names store answers no tag.
    On Error Resume Next
        Set store = HiddenNames.Create(targetSheet)
    On Error GoTo 0
    If store Is Nothing Then Exit Function

    IsDiseaseSheet = (StrComp(Trim$(store.ValueAsString(SHEET_TAG_NAME)), DISEASE_TAG_VALUE, vbTextCompare) = 0)
End Function

'@sub-title A workbook tag; a file with no hidden names store answers empty.
Private Function TagValue(ByVal sourceBook As Workbook, ByVal nameId As String) As String
    Dim store As HiddenNames

    'A foreign workbook can carry no store at all, and the store raises on it.
    On Error Resume Next
        Set store = HiddenNames.Create(sourceBook)
        If Not store Is Nothing Then TagValue = store.ValueAsString(nameId)
    On Error GoTo 0
End Function

'@sub-title Find a ListObject by name anywhere in the workbook, then by sheet.
Private Function FindTable(ByVal sourceBook As Workbook, _
                           ByVal tableName As String, _
                           ByVal sheetName As String) As ListObject
    Dim sheet As Worksheet
    Dim table As ListObject

    For Each sheet In sourceBook.Worksheets
        For Each table In sheet.ListObjects
            If StrComp(table.Name, tableName, vbTextCompare) = 0 Then
                Set FindTable = table
                Exit Function
            End If
        Next table
    Next sheet

    Set sheet = FindWorksheet(sourceBook, sheetName)
    If sheet Is Nothing Then Exit Function
    If sheet.ListObjects.Count = 0 Then Exit Function

    Set FindTable = sheet.ListObjects(1)
End Function

'@sub-title Locate a worksheet by name without relying on implicit errors.
Private Function FindWorksheet(ByVal targetBook As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet

    For Each sheet In targetBook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = sheet
            Exit Function
        End If
    Next sheet
End Function

'@sub-title Ensure a worksheet exists, renaming the given template when there is one.
Private Function EnsureWorksheet(ByVal targetBook As Workbook, _
                                 ByVal sheetName As String, _
                                 Optional ByVal template As Worksheet = Nothing) As Worksheet
    Dim sheet As Worksheet

    Set sheet = FindWorksheet(targetBook, sheetName)

    If sheet Is Nothing Then
        If Not template Is Nothing Then
            template.Name = sheetName
            Set sheet = template
        Else
            Set sheet = targetBook.Worksheets.Add(After:=targetBook.Worksheets(targetBook.Worksheets.Count))
            sheet.Name = sheetName
        End If
    End If

    sheet.Cells.Clear
    Set EnsureWorksheet = sheet
End Function

'@sub-title The last used column of a sheet; 0 for a sheet holding nothing.
Private Function LastUsedColumn(ByVal sheet As Worksheet) As Long
    Dim used As Range

    Set used = sheet.UsedRange
    If used Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(used) = 0 Then Exit Function

    LastUsedColumn = used.Columns(used.Columns.Count).Column
End Function

'@sub-title The header names of a ListObject, in column order.
Private Function HeaderNames(ByVal table As ListObject) As BetterArray
    Dim names As BetterArray
    Dim column As ListColumn

    Set names = New BetterArray
    names.LowerBound = 1

    For Each column In table.ListColumns
        names.Push column.Name
    Next column

    Set HeaderNames = names
End Function

'@sub-title The header row of a block read with its headers.
Private Function BlockHeaders(ByRef block As Variant) As BetterArray
    Dim names As BetterArray
    Dim columnIndex As Long

    Set names = New BetterArray
    names.LowerBound = 1

    For columnIndex = 1 To UBound(block, 2)
        names.Push Trim$(TextOf(block(1, columnIndex)))
    Next columnIndex

    Set BlockHeaders = names
End Function

'@sub-title The languages joined with the separator the Metadata sheet carries.
Private Function JoinedLanguages(ByVal names As BetterArray) As String
    Dim index As Long
    Dim languageName As String

    For index = names.LowerBound To names.UpperBound
        languageName = Trim$(CStr(names.Item(index)))
        If LenB(languageName) = 0 Then GoTo NextName
        If LenB(JoinedLanguages) > 0 Then JoinedLanguages = JoinedLanguages & LANGUAGE_SEPARATOR
        JoinedLanguages = JoinedLanguages & languageName
NextName:
    Next index
End Function

'@sub-title Index of the labels of a translations block: label key to row.
'@details The first row of the block is the header row; a label found twice
'keeps its first row.
Private Function LabelIndex(ByRef block As Variant) As Collection
    Dim keys As New Collection
    Dim rowIndex As Long
    Dim label As String

    For rowIndex = 2 To UBound(block, 1)
        label = TextOf(block(rowIndex, 1))
        If LenB(Trim$(label)) = 0 Then GoTo NextRow
        If LookupRow(keys, label) > 0 Then GoTo NextRow
        keys.Add rowIndex, LabelKey(label)
NextRow:
    Next rowIndex

    Set LabelIndex = keys
End Function

'@sub-title Index of the headers of a block: header to column, without regard to case.
Private Function HeaderIndex(ByRef block As Variant) As Collection
    Dim keys As New Collection
    Dim columnIndex As Long
    Dim header As String

    For columnIndex = 1 To UBound(block, 2)
        header = Trim$(TextOf(block(1, columnIndex)))
        If LenB(header) = 0 Then GoTo NextColumn
        If LookupColumn(keys, header) > 0 Then GoTo NextColumn
        keys.Add columnIndex, LCase$(header)
NextColumn:
    Next columnIndex

    Set HeaderIndex = keys
End Function

'@sub-title The row of a label in the index, 0 when it is not there.
Private Function LookupRow(ByVal keys As Collection, ByVal label As String) As Long
    'A Collection raises on a key it does not hold, and that is the answer.
    On Error Resume Next
        LookupRow = keys.Item(LabelKey(label))
    On Error GoTo 0
End Function

'@sub-title The column of a header in the index, 0 when it is not there.
Private Function LookupColumn(ByVal keys As Collection, ByVal header As String) As Long
    On Error Resume Next
        LookupColumn = keys.Item(LCase$(Trim$(header)))
    On Error GoTo 0
End Function

'@sub-title A Collection key that carries the case of a label.
'@details A Collection matches its keys without regard to case, and a
'translation label carries its case: "Hello" and "hello" are two labels.
'The key spells the capitalisation out beside the label, the way
'SetupTranslationsTable keys its own index.
Private Function LabelKey(ByVal label As String) As String
    Dim position As Long
    Dim character As String
    Dim signature As String

    For position = 1 To Len(label)
        character = Mid$(label, position, 1)
        If (UCase$(character) = character) And (LCase$(character) <> character) Then
            signature = signature & "1"
        Else
            signature = signature & "0"
        End If
    Next position

    LabelKey = label & "|" & signature
End Function

'@sub-title Whether a name has already been seen, without regard to case.
Private Function ContainsKey(ByVal keys As BetterArray, ByVal key As String) As Boolean
    Dim idx As Long

    If keys Is Nothing Then Exit Function
    If keys.Length = 0 Then Exit Function

    For idx = keys.LowerBound To keys.UpperBound
        If StrComp(CStr(keys.Item(idx)), key, vbTextCompare) = 0 Then
            ContainsKey = True
            Exit Function
        End If
    Next idx
End Function

'@sub-title Read a range as a two-dimensional block, one cell included.
Private Function RangeBlock(ByVal target As Range) As Variant
    Dim oneCell(1 To 1, 1 To 1) As Variant

    If target.Cells.Count = 1 Then
        oneCell(1, 1) = target.Value
        RangeBlock = oneCell
    Else
        RangeBlock = target.Value
    End If
End Function

'@sub-title Text of a cell value; an error value or an empty cell reads as empty.
Private Function TextOf(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    TextOf = CStr(cellValue)
End Function

'@sub-title Start the tallies of a new import.
Private Sub ResetTallies()
    Set this.importedDiseases = New BetterArray
    this.importedDiseases.LowerBound = 1
    Set this.addedLanguages = New BetterArray
    this.addedLanguages.LowerBound = 1
End Sub

'@section Logging
'===============================================================================

Private Sub Info(ByVal logger As DiseaseLogger, ByVal message As String)
    If logger Is Nothing Then Exit Sub
    logger.Record LOG_OPERATION, DiseaseLogInfo, message
End Sub

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised error raising aligned with ProjectError codes.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, CLASS_NAME, errMessage
End Sub
