VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Creates workbook exports for disease dictionaries and migration bundles"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Creates workbook exports for disease dictionaries and migration bundles")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'DiseaseExporter.cls without interface -- single implementation, no test fake,
'no compile cycle.

'@section Constants
'===============================================================================

Private Const DEFAULT_CHOICE_SEPARATOR As String = " | "
Private Const DISEASES_SHEET As String = "Diseases"
Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const CHOICES_SHEET As String = "Choices"
Private Const TRANSLATIONS_SHEET As String = "Translations"
Private Const METADATA_SHEET As String = "Metadata"
Private Const VARIABLES_SHEET As String = "Variables"
'The exported tables take the names a setup gives its own, since
'SetupImport matches the source tables to the host ones by name.
Private Const TABLE_DICTIONARY As String = "Tab_Dictionary"
Private Const TABLE_CHOICES As String = "Tab_Choices"
Private Const TABLE_TRANSLATIONS As String = "Tab_Translations"
Private Const HEADER_LIST_NAME As String = "list name"
Private Const HEADER_ORDER As String = "ordering list"
Private Const HEADER_RAW_LABEL As String = "non translated label"
Private Const HEADER_TRANSLATED_LABEL As String = "translated label"
Private Const HEADER_LABEL As String = "label"
Private Const HEADER_SHORT_LABEL As String = "short label"
Private Const HEADER_VARIABLE_ORDER As String = "Variable Order"
Private Const HEADER_MAIN_SECTION As String = "Main Section"
Private Const HEADER_VARIABLE_NAME As String = "Variable Name"
Private Const HEADER_MAIN_LABEL As String = "Main Label"
Private Const HEADER_CONTROL As String = "Control"
Private Const HEADER_STATUS As String = "Status"
Private Const MIGRATION_FILE_PREFIX As String = "flat_disease_file_"
Private Const DISEASE_FILE_PREFIX As String = "setup_dictionary_"
Private Const DISEASE_FILE_EXTENSION As String = ".xlsx"
Private Const LOG_OPERATION As String = "Export"

'The master Choices sheet carries its headers on row 4, column 1, the way
'MasterSetupHelpers reads it.
Private Const MASTER_CHOICES_START_ROW As Long = 4
Private Const MASTER_CHOICES_START_COLUMN As Long = 1

'The language and the code of a disease sheet live in its hidden names
'(owner rule, 2026-08-25). No cell carries them.
Private Const NAME_DISLANG As String = "__Var_DISLANG"
Private Const NAME_DISCODE As String = "__Var_DISCODE"

'Column order of a disease table, as DiseaseSheet builds it:
'Order, Name, Section, Label, Choice, Choice Values, Status.
Private Enum DiseaseColumns
    DiseaseOrderColumn = 1
    DiseaseNameColumn = 2
    DiseaseSectionColumn = 3
    DiseaseLabelColumn = 4
    DiseaseControlColumn = 5
    DiseaseChoicesColumn = 6
    DiseaseStatusColumn = 7
End Enum

'Column order of the exported Choices sheet. It is the layout of a setup
'Choices table, which is what SetupImport reads.
Private Enum ChoiceColumns
    ChoiceListColumn = 1
    ChoiceOrderColumn = 2
    ChoiceRawLabelColumn = 3
    ChoiceTranslatedColumn = 4
    ChoiceLabelColumn = 5
    ChoiceShortLabelColumn = 6
End Enum

Private Const CHOICE_COLUMN_COUNT As Long = 6
Private Const DICTIONARY_COLUMN_COUNT As Long = 6

'@section State
'===============================================================================

Private Type TDiseaseExporter
    Manager As DiseaseExportWorkbook
    Guard As ApplicationState
End Type

Private this As TDiseaseExporter

'@section Factory
'===============================================================================

'@sub-title Instantiate a DiseaseExporter with its dependencies configured.
'@param workbookManager DiseaseExportWorkbook providing workbook lifecycle handling.
'@param appGuard ApplicationState maintaining Application toggles.
'@return DiseaseExporter ready to generate exports.
Public Function Create(ByVal workbookManager As DiseaseExportWorkbook, _
                       ByVal appGuard As ApplicationState) As DiseaseExporter

    Dim exporter As DiseaseExporter
    Set exporter = New DiseaseExporter
    exporter.Initialise workbookManager, appGuard
    Set Create = exporter
End Function

'@section Configuration
'===============================================================================

'@sub-title Attach the workbook manager and application guard.
Public Sub Initialise(ByVal workbookManager As DiseaseExportWorkbook, _
                      ByVal appGuard As ApplicationState)

    If workbookManager Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Export workbook manager cannot be Nothing"
    End If

    If appGuard Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Application guard cannot be Nothing"
    End If

    Set this.Manager = workbookManager
    Set this.Guard = appGuard
End Sub

'@section Disease Export
'===============================================================================

'@sub-title Populate a workbook with dictionary, choices, translations, and metadata.
'@details
'The Choices sheet is the master Choices sheet cut down to the lists the
'disease uses, in the six-column layout of a setup. The two editable label
'columns are then translated into the language of the disease, and the
'label column carries the formula a setup keeps there: the translated label
'when present, the plain label otherwise. The dictionary, choices and
'translations blocks are each wrapped in a ListObject named the way a setup
'names its own, so SetupImport finds them by name.
'@param logger Optional DiseaseLogger. Receives one warning per list the
'   master Choices sheet does not carry, and one when the language has no
'   column in the translations table.
Public Function BuildDiseaseWorkbook(ByVal diseaseWksh As Worksheet, _
                                     ByVal translationTable As ListObject, _
                                     ByVal diseaseName As String, _
                                     ByVal languageTag As String, _
                                     Optional ByVal diseaseCode As String = vbNullString, _
                                     Optional ByVal logger As DiseaseLogger = Nothing) As Workbook

    Dim targetBook As Workbook
    Dim metadataSheet As Worksheet
    Dim dictionarySheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim translationsSheet As Worksheet
    Dim diseaseTable As ListObject
    Dim guardStarted As Boolean

    ValidateDiseaseInputs diseaseWksh, diseaseName, languageTag

    BeginGuard guardStarted

    EnsureFreshWorkbook
    Set targetBook = this.Manager.EnsureWorkbook(Application)

    PrepareSheets targetBook, metadataSheet, dictionarySheet, choicesSheet, translationsSheet

    Set diseaseTable = ResolveDiseaseTable(diseaseWksh)

    PopulateMetadata metadataSheet, diseaseName, languageTag, diseaseCode
    PopulateDictionary dictionarySheet, diseaseTable
    PopulateChoices choicesSheet, diseaseTable, ResolveMasterChoices(diseaseWksh), logger
    TranslateChoices choicesSheet, translationTable, languageTag, logger
    PopulateTranslations translationsSheet, translationTable, TABLE_TRANSLATIONS

    ApplySheetCosmetics metadataSheet, dictionarySheet, choicesSheet, translationsSheet

    Set BuildDiseaseWorkbook = targetBook

    RestoreGuard guardStarted
End Function

'@sub-title Export the disease workbook to disk and return the saved file path.
Public Function ExportDisease(ByVal exportFolder As String, _
                              ByVal diseaseWksh As Worksheet, _
                              ByVal translationTable As ListObject, _
                              ByVal diseaseName As String, _
                              ByVal languageTag As String, _
                              Optional ByVal diseaseCode As String = vbNullString, _
                              Optional ByVal logger As DiseaseLogger = Nothing) As String

    Dim targetBook As Workbook
    Dim filePath As String

    ValidateFolder exportFolder

    Set targetBook = BuildDiseaseWorkbook(diseaseWksh, translationTable, _
                                        diseaseName, languageTag, diseaseCode, logger)

    filePath = BuildDiseaseFilePath(exportFolder, diseaseName)
    this.Manager.SaveAs filePath
    this.Manager.CloseWorkbook False

    ExportDisease = filePath
End Function

'@section Migration Export
'===============================================================================

'@sub-title Build a workbook aggregating all diseases for migration.
Public Function BuildMigrationWorkbook(ByVal sourceWorkbook As Workbook, _
                                       ByVal diseaseNames As BetterArray) As Workbook

    Dim targetBook As Workbook
    Dim guardStarted As Boolean
    Dim translationsTable As ListObject
    Dim choicesTable As ListObject
    Dim variablesTable As ListObject
    Dim diseaseWksh As Worksheet
    Dim diseaseList As BetterArray

    ValidateMigrationInputs sourceWorkbook, diseaseNames

    BeginGuard guardStarted

    EnsureFreshWorkbook
    Set targetBook = this.Manager.EnsureWorkbook(Application)
    PrepareMigrationSheets targetBook

    Set translationsTable = ResolveListObject(sourceWorkbook, TRANSLATIONS_SHEET)
    Set choicesTable = ResolveListObject(sourceWorkbook, CHOICES_SHEET)
    Set variablesTable = ResolveListObject(sourceWorkbook, VARIABLES_SHEET)

    PopulateTranslations targetBook.Worksheets(TRANSLATIONS_SHEET), translationsTable
    PopulateTranslations targetBook.Worksheets(CHOICES_SHEET), choicesTable
    PopulateTranslations targetBook.Worksheets(VARIABLES_SHEET), variablesTable

    Set diseaseList = diseaseNames.Clone()
    For Each diseaseWksh In ResolveSheets(sourceWorkbook, diseaseList)
        PopulateMigrationDisease targetBook.Worksheets(DISEASES_SHEET), diseaseWksh
    Next diseaseWksh

    ApplyMigrationFormatting targetBook

    Set BuildMigrationWorkbook = targetBook

    RestoreGuard guardStarted
End Function

'@sub-title Persist the migration workbook to disk and return the saved file path.
Public Function ExportForMigration(ByVal exportFolder As String, _
                                   ByVal sourceWorkbook As Workbook, _
                                   ByVal diseaseNames As BetterArray) As String

    Dim targetBook As Workbook
    Dim filePath As String

    ValidateFolder exportFolder

    Set targetBook = BuildMigrationWorkbook(sourceWorkbook, diseaseNames)

    filePath = BuildMigrationFilePath(exportFolder)
    this.Manager.SaveAs filePath
    this.Manager.CloseWorkbook False

    ExportForMigration = filePath
End Function

'@section Internal Helpers
'===============================================================================

'@sub-title Ensure the workbook manager has a pristine workbook.
Private Sub EnsureFreshWorkbook()
    If this.Manager.HasWorkbook Then this.Manager.ReleaseWorkbook
End Sub

'@sub-title Begin the guarded Application state only once per call.
Private Sub BeginGuard(ByRef guardStarted As Boolean)
    If this.Guard Is Nothing Then Exit Sub
    this.Guard.ApplyBusyState suppressEvents:=True
    guardStarted = True
End Sub

'@sub-title Restore guarded Application state when required.
Private Sub RestoreGuard(ByVal guardStarted As Boolean)
    If Not guardStarted Then Exit Sub
    If Not this.Guard Is Nothing Then
        this.Guard.Restore
    End If
End Sub

'@sub-title Validate disease export inputs.
Private Sub ValidateDiseaseInputs(ByVal diseaseWksh As Worksheet, _
                                  ByVal diseaseName As String, _
                                  ByVal languageTag As String)

    If diseaseWksh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Disease worksheet cannot be Nothing"
    End If

    If LenB(diseaseName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Disease name cannot be empty"
    End If

    If LenB(languageTag) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Language tag cannot be empty"
    End If
End Sub

'@sub-title Validate migration export inputs.
Private Sub ValidateMigrationInputs(ByVal sourceWorkbook As Workbook, _
                                    ByVal diseaseNames As BetterArray)

    If sourceWorkbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source workbook cannot be Nothing"
    End If

    If diseaseNames Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Disease name list cannot be Nothing"
    End If

    If diseaseNames.Length = 0 Then
        ThrowError ProjectError.InvalidArgument, "Disease name list cannot be empty"
    End If
End Sub

'@sub-title Validate that a folder path was supplied.
Private Sub ValidateFolder(ByVal exportFolder As String)
    If LenB(exportFolder) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Export folder cannot be empty"
    End If
End Sub

'@sub-title Prepare disease export worksheets.
Private Sub PrepareSheets(ByVal targetBook As Workbook, _
                          ByRef metadataSheet As Worksheet, _
                          ByRef dictionarySheet As Worksheet, _
                          ByRef choicesSheet As Worksheet, _
                          ByRef translationsSheet As Worksheet)

    Set metadataSheet = EnsureWorksheet(targetBook, METADATA_SHEET, targetBook.Worksheets(1))
    Set dictionarySheet = EnsureWorksheet(targetBook, DICTIONARY_SHEET)
    Set choicesSheet = EnsureWorksheet(targetBook, CHOICES_SHEET)
    Set translationsSheet = EnsureWorksheet(targetBook, TRANSLATIONS_SHEET)

    ClearSheet metadataSheet
    ClearSheet dictionarySheet
    ClearSheet choicesSheet
    ClearSheet translationsSheet
End Sub

'@sub-title Prepare migration workbook worksheets.
Private Sub PrepareMigrationSheets(ByVal targetBook As Workbook)
    ClearSheet EnsureWorksheet(targetBook, DISEASES_SHEET, targetBook.Worksheets(1))
    ClearSheet EnsureWorksheet(targetBook, TRANSLATIONS_SHEET)
    ClearSheet EnsureWorksheet(targetBook, CHOICES_SHEET)
    ClearSheet EnsureWorksheet(targetBook, VARIABLES_SHEET)
End Sub

'@sub-title Retrieve the disease table from the worksheet.
Private Function ResolveDiseaseTable(ByVal diseaseWksh As Worksheet) As ListObject
    If diseaseWksh.ListObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, "Disease worksheet must contain a ListObject"
    End If
    Set ResolveDiseaseTable = diseaseWksh.ListObjects(1)
End Function

'@sub-title Retrieve a ListObject from a named worksheet.
Private Function ResolveListObject(ByVal targetBook As Workbook, ByVal sheetName As String) As ListObject
    Dim sheet As Worksheet
    Set sheet = targetBook.Worksheets(sheetName)
    If sheet.ListObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, "Worksheet '" & sheetName & "' is missing its ListObject"
    End If
    Set ResolveListObject = sheet.ListObjects(1)
End Function

'@sub-title Wrap the master Choices sheet of the workbook holding the disease.
'@details Answers Nothing when the workbook carries no Choices sheet; the
'export then falls back on the values the disease sheet carries.
Private Function ResolveMasterChoices(ByVal diseaseWksh As Worksheet) As LLChoices
    Dim choicesWksh As Worksheet

    Set choicesWksh = FindWorksheet(diseaseWksh.Parent, CHOICES_SHEET)
    If choicesWksh Is Nothing Then Exit Function

    Set ResolveMasterChoices = LLChoices.Create(choicesWksh, MASTER_CHOICES_START_ROW, MASTER_CHOICES_START_COLUMN)
End Function

'@sub-title Enumerate worksheets from provided disease names.
Private Function ResolveSheets(ByVal targetBook As Workbook, ByVal diseaseNames As BetterArray) As Collection
    Dim results As New Collection
    Dim index As Long
    Dim diseaseName As String
    Dim sheet As Worksheet

    For index = diseaseNames.LowerBound To diseaseNames.UpperBound
        diseaseName = CStr(diseaseNames.Item(index))
        Set sheet = FindWorksheet(targetBook, diseaseName)
        If sheet Is Nothing Then
            ThrowError ProjectError.ElementNotFound, "Disease worksheet '" & diseaseName & "' does not exist"
        End If
        results.Add sheet
    Next index

    Set ResolveSheets = results
End Function

'@sub-title Populate metadata worksheet with export details.
Private Sub PopulateMetadata(ByVal metadataSheet As Worksheet, _
                             ByVal diseaseName As String, _
                             ByVal languageTag As String, _
                             ByVal diseaseCode As String)

    Dim metadata(1 To 5, 1 To 2) As Variant

    metadata(1, 1) = "Label"
    metadata(1, 2) = "Value"
    metadata(2, 1) = "export_date"
    metadata(2, 2) = Format$(Now, "yyyy-mm-dd HH:nn")
    metadata(3, 1) = "export_disease"
    metadata(3, 2) = diseaseName
    metadata(4, 1) = "dictionary_language"
    metadata(4, 2) = languageTag
    metadata(5, 1) = "dictionary_code"
    metadata(5, 2) = diseaseCode

    metadataSheet.Range("A1").Resize(UBound(metadata, 1), UBound(metadata, 2)).Value = metadata
End Sub

'@sub-title Populate the dictionary sheet in the Tab_Dictionary table.
'@details
'The headers are the ones of a setup dictionary, so CustomTable.Import
'matches them by name. Variable Order stays as the first column: the
'setup dictionary has no such column and the import skips it.
Private Sub PopulateDictionary(ByVal dictionarySheet As Worksheet, ByVal diseaseTable As ListObject)

    Dim sourceData As Variant
    Dim prepared As Variant
    Dim rowCount As Long

    WriteHeaderRow dictionarySheet.Range("A1"), _
        Array(HEADER_VARIABLE_ORDER, HEADER_MAIN_SECTION, HEADER_VARIABLE_NAME, _
              HEADER_MAIN_LABEL, HEADER_CONTROL, HEADER_STATUS)

    If Not diseaseTable.DataBodyRange Is Nothing Then
        sourceData = diseaseTable.DataBodyRange.Value
        rowCount = UBound(sourceData, 1)
        prepared = BuildDictionaryMatrix(sourceData, rowCount)

        If Not IsEmpty(prepared) Then
            dictionarySheet.Range("A2").Resize(rowCount, DICTIONARY_COLUMN_COUNT).Value = prepared
        End If
    End If

    WrapAsTable dictionarySheet, rowCount, DICTIONARY_COLUMN_COUNT, TABLE_DICTIONARY
End Sub

'@sub-title Build a dictionary matrix selecting required columns.
Private Function BuildDictionaryMatrix(ByRef sourceData As Variant, ByVal rowCount As Long) As Variant
    Dim matrix() As Variant
    Dim rowIndex As Long

    ReDim matrix(1 To rowCount, 1 To DICTIONARY_COLUMN_COUNT)

    For rowIndex = 1 To rowCount
        matrix(rowIndex, 1) = sourceData(rowIndex, DiseaseOrderColumn)
        matrix(rowIndex, 2) = sourceData(rowIndex, DiseaseSectionColumn)
        matrix(rowIndex, 3) = sourceData(rowIndex, DiseaseNameColumn)
        matrix(rowIndex, 4) = sourceData(rowIndex, DiseaseLabelColumn)
        matrix(rowIndex, 5) = sourceData(rowIndex, DiseaseControlColumn)
        matrix(rowIndex, 6) = sourceData(rowIndex, DiseaseStatusColumn)
    Next rowIndex

    BuildDictionaryMatrix = matrix
End Function

'@section Choices Export
'===============================================================================

'@sub-title Populate the choices sheet with the lists the disease uses.
'@details
'Six columns, in the order of a setup Choices table: list name, ordering
'list, non translated label, translated label, label, short label. The rows
'come from the master Choices sheet, read once as a block and kept to the
'lists the disease names. A list the master sheet does not carry takes the
'values the disease sheet holds for it, and is logged. The label column is
'written as the formula a setup keeps there. The block is wrapped in the
'Tab_Choices table.
Private Sub PopulateChoices(ByVal choicesSheet As Worksheet, _
                            ByVal diseaseTable As ListObject, _
                            ByVal masterChoices As LLChoices, _
                            ByVal logger As DiseaseLogger)

    Dim choiceRows As Collection
    Dim choiceMatrix As Variant
    Dim rowCount As Long

    WriteHeaderRow choicesSheet.Range("A1"), _
        Array(HEADER_LIST_NAME, HEADER_ORDER, HEADER_RAW_LABEL, _
              HEADER_TRANSLATED_LABEL, HEADER_LABEL, HEADER_SHORT_LABEL)

    If Not diseaseTable.DataBodyRange Is Nothing Then
        Set choiceRows = BuildChoiceRows(diseaseTable.DataBodyRange.Value, masterChoices, logger)
        choiceMatrix = CollectionToMatrix(choiceRows, CHOICE_COLUMN_COUNT)

        If Not IsEmpty(choiceMatrix) Then
            rowCount = UBound(choiceMatrix, 1)
            choicesSheet.Range("A2").Resize(rowCount, CHOICE_COLUMN_COUNT).Value = choiceMatrix
            WriteLabelFormulas choicesSheet, rowCount
        End If
    End If

    WrapAsTable choicesSheet, rowCount, CHOICE_COLUMN_COUNT, TABLE_CHOICES
End Sub

'@sub-title Build the choice rows: the master rows of every list the disease names.
Private Function BuildChoiceRows(ByRef sourceData As Variant, _
                                 ByVal masterChoices As LLChoices, _
                                 ByVal logger As DiseaseLogger) As Collection
    Dim rows As New Collection
    Dim listNames As BetterArray
    Dim masterBlock As Variant
    Dim nameColumn As Long
    Dim labelColumn As Long
    Dim shortColumn As Long
    Dim hasMaster As Boolean
    Dim listIndex As Long
    Dim listName As String
    Dim rowsAdded As Long

    Set listNames = DiseaseListNames(sourceData)

    'One read of the master block serves every list of the disease.
    If Not masterChoices Is Nothing Then
        hasMaster = ReadMasterBlock(masterChoices, masterBlock, nameColumn, labelColumn, shortColumn)
    End If

    For listIndex = listNames.LowerBound To listNames.UpperBound
        listName = CStr(listNames.Item(listIndex))
        rowsAdded = 0
        If hasMaster Then
            rowsAdded = AppendMasterRows(rows, masterBlock, nameColumn, labelColumn, shortColumn, listName)
        End If
        If rowsAdded = 0 Then
            AppendDiseaseRows rows, sourceData, listName
            Warn logger, "The master Choices sheet carries no list named '" & listName & _
                         "'; the export takes the values of the disease sheet"
        End If
    Next listIndex

    Set BuildChoiceRows = rows
End Function

'@sub-title Distinct list names of the disease table, in first-use order.
Private Function DiseaseListNames(ByRef sourceData As Variant) As BetterArray
    Dim names As BetterArray
    Dim rowIndex As Long
    Dim controlName As String

    Set names = New BetterArray
    names.LowerBound = 1

    For rowIndex = 1 To UBound(sourceData, 1)
        controlName = Trim$(CStr(sourceData(rowIndex, DiseaseControlColumn)))
        If LenB(controlName) = 0 Then GoTo NextRow
        If Not ContainsKey(names, controlName) Then names.Push controlName
NextRow:
    Next rowIndex

    Set DiseaseListNames = names
End Function

'@sub-title Read the master Choices block with its headers, once.
'@return Boolean. True when the block carries a list name and a label column.
Private Function ReadMasterBlock(ByVal masterChoices As LLChoices, _
                                 ByRef block As Variant, _
                                 ByRef nameColumn As Long, _
                                 ByRef labelColumn As Long, _
                                 ByRef shortColumn As Long) As Boolean
    Dim data As DataSheet
    Dim columnIndex As Long
    Dim headerText As String

    Set data = masterChoices.ChoiceData
    If Not data.ColumnExists(HEADER_LIST_NAME, matchCase:=False) Then Exit Function
    If Not data.ColumnExists(HEADER_LABEL, matchCase:=False) Then Exit Function

    block = RangeBlock(data.DataRange("__all__", includeHeaders:=True))

    For columnIndex = 1 To UBound(block, 2)
        headerText = LCase$(Trim$(TextOf(block(1, columnIndex))))
        Select Case headerText
            Case HEADER_LIST_NAME
                nameColumn = columnIndex
            Case HEADER_LABEL
                labelColumn = columnIndex
            Case HEADER_SHORT_LABEL
                shortColumn = columnIndex
        End Select
    Next columnIndex

    ReadMasterBlock = (nameColumn > 0) And (labelColumn > 0)
End Function

'@sub-title Append the master rows of one list; answers how many were added.
Private Function AppendMasterRows(ByVal rows As Collection, _
                                  ByRef block As Variant, _
                                  ByVal nameColumn As Long, _
                                  ByVal labelColumn As Long, _
                                  ByVal shortColumn As Long, _
                                  ByVal listName As String) As Long
    Dim rowIndex As Long
    Dim position As Long
    Dim labelValue As String
    Dim shortValue As String

    For rowIndex = 2 To UBound(block, 1)
        If StrComp(Trim$(TextOf(block(rowIndex, nameColumn))), listName, vbTextCompare) <> 0 Then GoTo NextRow

        labelValue = Trim$(TextOf(block(rowIndex, labelColumn)))
        If LenB(labelValue) = 0 Then GoTo NextRow

        shortValue = vbNullString
        If shortColumn > 0 Then shortValue = Trim$(TextOf(block(rowIndex, shortColumn)))
        If LenB(shortValue) = 0 Then shortValue = labelValue

        position = position + 1
        rows.Add ChoiceRow(listName, position, labelValue, shortValue)
NextRow:
    Next rowIndex

    AppendMasterRows = position
End Function

'@sub-title Append the values the disease sheet carries for a list the master lacks.
Private Sub AppendDiseaseRows(ByVal rows As Collection, _
                              ByRef sourceData As Variant, _
                              ByVal listName As String)
    Dim rowIndex As Long
    Dim splitValues As Variant
    Dim itemIndex As Long
    Dim choiceValue As String
    Dim position As Long

    For rowIndex = 1 To UBound(sourceData, 1)
        If StrComp(Trim$(CStr(sourceData(rowIndex, DiseaseControlColumn))), listName, vbTextCompare) <> 0 Then GoTo NextRow

        splitValues = SplitChoices(sourceData(rowIndex, DiseaseChoicesColumn))
        If IsEmpty(splitValues) Then Exit Sub

        For itemIndex = LBound(splitValues) To UBound(splitValues)
            choiceValue = Trim$(CStr(splitValues(itemIndex)))
            If LenB(choiceValue) = 0 Then GoTo NextChoice
            position = position + 1
            rows.Add ChoiceRow(listName, position, choiceValue, choiceValue)
NextChoice:
        Next itemIndex

        'The first row naming the list carries its values; the others repeat them.
        Exit Sub
NextRow:
    Next rowIndex
End Sub

'@sub-title Build one export row; the label column is filled by its formula later.
Private Function ChoiceRow(ByVal listName As String, _
                           ByVal position As Long, _
                           ByVal labelValue As String, _
                           ByVal shortValue As String) As Variant
    Dim rowValues(1 To CHOICE_COLUMN_COUNT) As Variant

    rowValues(ChoiceListColumn) = listName
    rowValues(ChoiceOrderColumn) = position
    rowValues(ChoiceRawLabelColumn) = labelValue
    rowValues(ChoiceTranslatedColumn) = labelValue
    rowValues(ChoiceLabelColumn) = vbNullString
    rowValues(ChoiceShortLabelColumn) = shortValue

    ChoiceRow = rowValues
End Function

'@sub-title Write the label formula on every exported choice row.
'@details The translated label when present, the non translated label
'otherwise, which is the formula the setup Choices sheet keeps.
Private Sub WriteLabelFormulas(ByVal choicesSheet As Worksheet, ByVal rowCount As Long)
    Dim labelRange As Range
    Dim translatedAddress As String
    Dim rawAddress As String

    Set labelRange = choicesSheet.Cells(2, ChoiceLabelColumn).Resize(rowCount, 1)
    translatedAddress = choicesSheet.Cells(2, ChoiceTranslatedColumn).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    rawAddress = choicesSheet.Cells(2, ChoiceRawLabelColumn).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    labelRange.Formula = "=IF(" & translatedAddress & "<>"""", " & translatedAddress & ", " & rawAddress & ")"

    'A workbook under manual calculation reads the formula results right away.
    labelRange.Calculate
End Sub

'@sub-title Translate the exported choices into the language of the disease.
'@details LLChoices.Translate writes the translated label and the short
'label columns; the non translated label keeps the master value and the
'label formula follows. A language the translations table has no column for
'leaves the values as they are and is logged.
Private Sub TranslateChoices(ByVal choicesSheet As Worksheet, _
                             ByVal translationTable As ListObject, _
                             ByVal languageTag As String, _
                             ByVal logger As DiseaseLogger)
    Dim exported As LLChoices
    Dim trads As TranslationObject

    If LenB(TextOf(choicesSheet.Cells(2, ChoiceListColumn).Value)) = 0 Then Exit Sub
    If translationTable Is Nothing Then Exit Sub

    If Not TableHasColumn(translationTable, languageTag) Then
        Warn logger, "The translations table has no column for the language '" & languageTag & _
                     "'; the choices keep the master values"
        Exit Sub
    End If

    Set trads = TranslationObject.Create(translationTable, languageTag)
    Set exported = LLChoices.Create(choicesSheet, 1, 1)
    exported.Translate trads
End Sub

'@sub-title Whether a ListObject carries a header of the given name.
Private Function TableHasColumn(ByVal table As ListObject, ByVal headerName As String) As Boolean
    Dim headerCell As Range

    If table.HeaderRowRange Is Nothing Then Exit Function

    For Each headerCell In table.HeaderRowRange.Cells
        If StrComp(Trim$(TextOf(headerCell.Value)), Trim$(headerName), vbTextCompare) = 0 Then
            TableHasColumn = True
            Exit Function
        End If
    Next headerCell
End Function

'@sub-title Record a warning on the logger when one is given.
Private Sub Warn(ByVal logger As DiseaseLogger, ByVal message As String)
    If logger Is Nothing Then Exit Sub
    logger.Record LOG_OPERATION, DiseaseLogWarning, message
End Sub

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

'@sub-title Text of a cell value; an error value reads as empty.
Private Function TextOf(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    TextOf = CStr(cellValue)
End Function

'@section Shared Helpers
'===============================================================================

'@sub-title Determine whether a name has already been seen, without regard to case.
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

'@sub-title Split choice strings safely.
Private Function SplitChoices(ByVal rawValue As Variant) As Variant
    Dim textValue As String
    textValue = Trim$(CStr(rawValue))
    If LenB(textValue) = 0 Then Exit Function
    SplitChoices = Split(textValue, DEFAULT_CHOICE_SEPARATOR)
End Function

'@sub-title Copy the values of a ListObject to a worksheet in one assignment.
'@param tableName Optional String. When given, the copied block is wrapped
'   in a ListObject of that name.
Private Sub PopulateTranslations(ByVal targetSheet As Worksheet, _
                                 ByVal sourceTable As ListObject, _
                                 Optional ByVal tableName As String = vbNullString)
    Dim sourceRange As Range
    Dim values As Variant

    If sourceTable Is Nothing Then Exit Sub

    Set sourceRange = sourceTable.Range
    If sourceRange Is Nothing Then Exit Sub

    values = sourceRange.Value
    targetSheet.Range("A1").Resize(UBound(values, 1), UBound(values, 2)).Value = values

    If LenB(tableName) = 0 Then Exit Sub
    WrapAsTable targetSheet, UBound(values, 1) - 1, UBound(values, 2), tableName
End Sub

'@sub-title Wrap a block anchored on A1 in a plain ListObject of the given name.
'@param dataRowCount Long. The rows under the header row; zero wraps the
'   header row alone.
Private Sub WrapAsTable(ByVal targetSheet As Worksheet, _
                        ByVal dataRowCount As Long, _
                        ByVal columnCount As Long, _
                        ByVal tableName As String)
    Dim blockRange As Range
    Dim table As ListObject

    Set blockRange = targetSheet.Range("A1").Resize(dataRowCount + 1, columnCount)
    Set table = targetSheet.ListObjects.Add(xlSrcRange, blockRange, , xlYes)
    table.Name = tableName
    table.TableStyle = ""
End Sub

'@sub-title Populate a migration worksheet with disease data.
Private Sub PopulateMigrationDisease(ByVal targetSheet As Worksheet, ByVal diseaseWksh As Worksheet)
    Dim diseaseTable As ListObject
    Dim sourceRange As Range
    Dim values As Variant
    Dim lastColumn As Long
    Dim store As HiddenNames

    Set diseaseTable = ResolveDiseaseTable(diseaseWksh)

    Set sourceRange = diseaseTable.Range
    If sourceRange Is Nothing Then Exit Sub

    values = sourceRange.Value
    lastColumn = LastUsedColumn(targetSheet) + 2
    If lastColumn < 1 Then lastColumn = 1

    Set store = HiddenNames.Create(diseaseWksh)

    targetSheet.Cells(1, lastColumn).Value = "Disease"
    targetSheet.Cells(1, lastColumn + 1).Value = "Language"
    targetSheet.Cells(1, lastColumn + 2).Value = "Tag"

    targetSheet.Cells(2, lastColumn).Value = diseaseWksh.Name
    targetSheet.Cells(2, lastColumn + 1).Value = store.ValueAsString(NAME_DISLANG)
    targetSheet.Cells(2, lastColumn + 2).Value = store.ValueAsString(NAME_DISCODE)

    targetSheet.Range(targetSheet.Cells(3, lastColumn), _
                      targetSheet.Cells(2 + UBound(values, 1), lastColumn + UBound(values, 2) - 1)).Value = values
End Sub

'@sub-title Apply light formatting to export worksheets.
Private Sub ApplySheetCosmetics(ByVal metadataSheet As Worksheet, _
                                ByVal dictionarySheet As Worksheet, _
                                ByVal choicesSheet As Worksheet, _
                                ByVal translationsSheet As Worksheet)

    AutoFitSheet metadataSheet
    AutoFitSheet dictionarySheet
    AutoFitSheet choicesSheet
    AutoFitSheet translationsSheet
End Sub

'@sub-title Apply formatting to migration workbook sheets.
Private Sub ApplyMigrationFormatting(ByVal targetBook As Workbook)
    AutoFitSheet targetBook.Worksheets(DISEASES_SHEET)
    AutoFitSheet targetBook.Worksheets(TRANSLATIONS_SHEET)
    AutoFitSheet targetBook.Worksheets(CHOICES_SHEET)
    AutoFitSheet targetBook.Worksheets(VARIABLES_SHEET)
End Sub

'@sub-title Auto fit worksheet columns and style header row.
Private Sub AutoFitSheet(ByVal sheet As Worksheet)
    With sheet
        .Cells.WrapText = False
        .Rows(1).Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

'@sub-title Ensure a worksheet exists and optionally rename the provided template.
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

    Set EnsureWorksheet = sheet
End Function

'@sub-title Clear worksheet contents while preserving structure.
Private Sub ClearSheet(ByVal sheet As Worksheet)
    If sheet Is Nothing Then Exit Sub
    sheet.Cells.Clear
End Sub

'@sub-title Write one header row through a BetterArray.
Private Sub WriteHeaderRow(ByVal anchorCell As Range, ByVal values As Variant)
    Dim headers As BetterArray

    Set headers = New BetterArray
    headers.LowerBound = 1
    headers.Items = values
    headers.ToExcelRange anchorCell, TransposeValues:=True
End Sub

'@sub-title Convert a collection of row arrays into a matrix.
Private Function CollectionToMatrix(ByVal rows As Collection, ByVal columnCount As Long) As Variant
    Dim matrix() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim rowValues() As Variant

    If rows Is Nothing Then Exit Function
    If rows.Count = 0 Then Exit Function

    ReDim matrix(1 To rows.Count, 1 To columnCount)

    For rowIndex = 1 To rows.Count
        rowValues = rows(rowIndex)
        For colIndex = 1 To columnCount
            matrix(rowIndex, colIndex) = rowValues(colIndex)
        Next colIndex
    Next rowIndex

    CollectionToMatrix = matrix
End Function

'@sub-title Determine the last used column within the target sheet.
Private Function LastUsedColumn(ByVal sheet As Worksheet) As Long
    Dim usedRange As Range

    Set usedRange = sheet.UsedRange
    If usedRange Is Nothing Then
        LastUsedColumn = 0
    Else
        LastUsedColumn = usedRange.Columns(usedRange.Columns.Count).Column
    End If
End Function

'@sub-title Build path for disease export.
Private Function BuildDiseaseFilePath(ByVal exportFolder As String, ByVal diseaseName As String) As String
    BuildDiseaseFilePath = exportFolder & Application.PathSeparator & _
                           DISEASE_FILE_PREFIX & SanitizeFileName(diseaseName) & DISEASE_FILE_EXTENSION
End Function

'@sub-title Build path for migration export.
Private Function BuildMigrationFilePath(ByVal exportFolder As String) As String
    BuildMigrationFilePath = exportFolder & Application.PathSeparator & _
                             MIGRATION_FILE_PREFIX & Format$(Now, "yyyy-mm-dd") & DISEASE_FILE_EXTENSION
End Function

'@sub-title Remove characters invalid in file names.
Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim result As String
    Dim index As Long
    Dim invalidChars As Variant

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    result = fileName

    For index = LBound(invalidChars) To UBound(invalidChars)
        result = Replace$(result, CStr(invalidChars(index)), "_")
    Next index

    SanitizeFileName = result
End Function

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised ProjectError-based exception helper.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseExporter", errMessage
End Sub
