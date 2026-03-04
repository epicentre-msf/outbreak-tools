VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Creates workbook exports for disease dictionaries and migration bundles"

'@Folder("Msetup")
'@ModuleDescription("Creates workbook exports for disease dictionaries and migration bundles")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDiseaseExporter

'@section Constants
'===============================================================================

Private Const DEFAULT_CHOICE_SEPARATOR As String = " | "
Private Const DISEASES_SHEET As String = "Diseases"
Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const CHOICES_SHEET As String = "Choices"
Private Const TRANSLATIONS_SHEET As String = "Translations"
Private Const METADATA_SHEET As String = "Metadata"
Private Const VARIABLES_SHEET As String = "Variables"
Private Const HEADER_LIST_NAME As String = "list name"
Private Const HEADER_LABEL As String = "label"
Private Const HEADER_SHORT_LABEL As String = "short label"
Private Const HEADER_ORDER As String = "ordering list"
Private Const HEADER_VARIABLE_ORDER As String = "Variable Order"
Private Const HEADER_VARIABLE_SECTION As String = "Variable Section"
Private Const HEADER_VARIABLE_NAME As String = "Variable Name"
Private Const HEADER_MAIN_LABEL As String = "Main Label"
Private Const HEADER_CONTROL As String = "Control"
Private Const HEADER_STATUS As String = "Status"
Private Const MIGRATION_FILE_PREFIX As String = "flat_disease_file_"
Private Const DISEASE_FILE_PREFIX As String = "setup_dictionary_"
Private Const DISEASE_FILE_EXTENSION As String = ".xlsx"

Private Enum DiseaseColumns
    DiseaseOrderColumn = 1
    DiseaseSectionColumn = 2
    DiseaseNameColumn = 3
    DiseaseLabelColumn = 4
    DiseaseControlColumn = 5
    DiseaseChoicesColumn = 6
    DiseaseStatusColumn = 7
End Enum

'@section State
'===============================================================================

Private Type TDiseaseExporter
    Manager As IDiseaseExportWorkbook
    Guard As IDiseaseApplicationState
End Type

Private this As TDiseaseExporter

'@section Factory
'===============================================================================

'@sub-title Instantiate a DiseaseExporter with its dependencies configured.
'@param workbookManager IDiseaseExportWorkbook providing workbook lifecycle handling.
'@param appGuard IDiseaseApplicationState maintaining Application toggles.
'@return IDiseaseExporter ready to generate exports.
Public Function Create(ByVal workbookManager As IDiseaseExportWorkbook, _
                       ByVal appGuard As IDiseaseApplicationState) As IDiseaseExporter

    With New DiseaseExporter
        .Initialise workbookManager, appGuard
        Set Create = .Self
    End With
End Function

'@sub-title Return the current instance as its interface.
Public Property Get Self() As IDiseaseExporter
    Set Self = Me
End Property

'@section Configuration
'===============================================================================

'@sub-title Attach the workbook manager and application guard.
Public Sub Initialise(ByVal workbookManager As IDiseaseExportWorkbook, _
                      ByVal appGuard As IDiseaseApplicationState)

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
Public Function BuildDiseaseWorkbook(ByVal diseaseSheet As Worksheet, _
                                     ByVal translationTable As ListObject, _
                                     ByVal ribbonTranslations As ITranslationObject, _
                                     ByVal diseaseName As String, _
                                     ByVal languageTag As String, _
                                     Optional ByVal diseaseCode As String = vbNullString) As Workbook

    Dim workbook As Workbook
    Dim metadataSheet As Worksheet
    Dim dictionarySheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim translationsSheet As Worksheet
    Dim diseaseTable As ListObject
    Dim headerCache As IDiseaseTranslationCache
    Dim guardStarted As Boolean

    ValidateDiseaseInputs diseaseSheet, diseaseName, languageTag

    BeginGuard guardStarted

    EnsureFreshWorkbook
    Set workbook = this.Manager.EnsureWorkbook(Application)

    PrepareSheets workbook, metadataSheet, dictionarySheet, choicesSheet, translationsSheet

    Set diseaseTable = ResolveDiseaseTable(diseaseSheet)

    Set headerCache = DiseaseTranslationCache.Create(ribbonTranslations)

    PopulateMetadata metadataSheet, diseaseName, languageTag, diseaseCode
    PopulateDictionary dictionarySheet, diseaseTable
    PopulateChoices choicesSheet, diseaseTable, headerCache
    PopulateTranslations translationsSheet, translationTable

    ApplySheetCosmetics metadataSheet, dictionarySheet, choicesSheet, translationsSheet

    Set BuildDiseaseWorkbook = workbook

    RestoreGuard guardStarted
End Function

'@sub-title Export the disease workbook to disk and return the saved file path.
Public Function ExportDisease(ByVal exportFolder As String, _
                              ByVal diseaseSheet As Worksheet, _
                              ByVal translationTable As ListObject, _
                              ByVal ribbonTranslations As ITranslationObject, _
                              ByVal diseaseName As String, _
                              ByVal languageTag As String, _
                              Optional ByVal diseaseCode As String = vbNullString) As String

    Dim workbook As Workbook
    Dim filePath As String

    ValidateFolder exportFolder

    Set workbook = BuildDiseaseWorkbook(diseaseSheet, translationTable, ribbonTranslations, _
                                        diseaseName, languageTag, diseaseCode)

    filePath = BuildDiseaseFilePath(exportFolder, diseaseName)
    this.Manager.SaveAs filePath
    this.Manager.Close False

    ExportDisease = filePath
End Function

'@section Migration Export
'===============================================================================

'@sub-title Build a workbook aggregating all diseases for migration.
Public Function BuildMigrationWorkbook(ByVal sourceWorkbook As Workbook, _
                                       ByVal diseaseNames As BetterArray) As Workbook

    Dim workbook As Workbook
    Dim guardStarted As Boolean
    Dim translationsTable As ListObject
    Dim choicesTable As ListObject
    Dim variablesTable As ListObject
    Dim diseaseSheet As Worksheet
    Dim diseaseList As BetterArray

    ValidateMigrationInputs sourceWorkbook, diseaseNames

    BeginGuard guardStarted

    EnsureFreshWorkbook
    Set workbook = this.Manager.EnsureWorkbook(Application)
    PrepareMigrationSheets workbook

    Set translationsTable = ResolveListObject(sourceWorkbook, TRANSLATIONS_SHEET)
    Set choicesTable = ResolveListObject(sourceWorkbook, CHOICES_SHEET)
    Set variablesTable = ResolveListObject(sourceWorkbook, VARIABLES_SHEET)

    PopulateTranslations workbook.Worksheets(TRANSLATIONS_SHEET), translationsTable
    PopulateTranslations workbook.Worksheets(CHOICES_SHEET), choicesTable
    PopulateTranslations workbook.Worksheets(VARIABLES_SHEET), variablesTable

    Set diseaseList = diseaseNames.Clone()
    For Each diseaseSheet In ResolveSheets(sourceWorkbook, diseaseList)
        PopulateMigrationDisease workbook.Worksheets(DISEASES_SHEET), diseaseSheet
    Next diseaseSheet

    ApplyMigrationFormatting workbook

    Set BuildMigrationWorkbook = workbook

    RestoreGuard guardStarted
End Function

'@sub-title Persist the migration workbook to disk and return the saved file path.
Public Function ExportForMigration(ByVal exportFolder As String, _
                                   ByVal sourceWorkbook As Workbook, _
                                   ByVal diseaseNames As BetterArray) As String

    Dim workbook As Workbook
    Dim filePath As String

    ValidateFolder exportFolder

    Set workbook = BuildMigrationWorkbook(sourceWorkbook, diseaseNames)

    filePath = BuildMigrationFilePath(exportFolder)
    this.Manager.SaveAs filePath
    this.Manager.Close False

    ExportForMigration = filePath
End Function

'@section Internal Helpers
'===============================================================================

'@sub-title Ensure the workbook manager has a pristine workbook.
Private Sub EnsureFreshWorkbook()
    If this.Manager.HasWorkbook Then this.Manager.Release
End Sub

'@sub-title Begin the guarded Application state only once per call.
Private Sub BeginGuard(ByRef guardStarted As Boolean)
    If this.Guard Is Nothing Then Exit Sub
    this.Guard.BeginGuard Application
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
Private Sub ValidateDiseaseInputs(ByVal diseaseSheet As Worksheet, _
                                  ByVal diseaseName As String, _
                                  ByVal languageTag As String)

    If diseaseSheet Is Nothing Then
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
Private Sub PrepareSheets(ByVal workbook As Workbook, _
                          ByRef metadataSheet As Worksheet, _
                          ByRef dictionarySheet As Worksheet, _
                          ByRef choicesSheet As Worksheet, _
                          ByRef translationsSheet As Worksheet)

    Set metadataSheet = EnsureWorksheet(workbook, METADATA_SHEET, workbook.Worksheets(1))
    Set dictionarySheet = EnsureWorksheet(workbook, DICTIONARY_SHEET)
    Set choicesSheet = EnsureWorksheet(workbook, CHOICES_SHEET)
    Set translationsSheet = EnsureWorksheet(workbook, TRANSLATIONS_SHEET)

    ClearSheet metadataSheet
    ClearSheet dictionarySheet
    ClearSheet choicesSheet
    ClearSheet translationsSheet
End Sub

'@sub-title Prepare migration workbook worksheets.
Private Sub PrepareMigrationSheets(ByVal workbook As Workbook)
    ClearSheet EnsureWorksheet(workbook, DISEASES_SHEET, workbook.Worksheets(1))
    ClearSheet EnsureWorksheet(workbook, TRANSLATIONS_SHEET)
    ClearSheet EnsureWorksheet(workbook, CHOICES_SHEET)
    ClearSheet EnsureWorksheet(workbook, VARIABLES_SHEET)
End Sub

'@sub-title Retrieve the disease table from the worksheet.
Private Function ResolveDiseaseTable(ByVal diseaseSheet As Worksheet) As ListObject
    If diseaseSheet.ListObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, "Disease worksheet must contain a ListObject"
    End If
    Set ResolveDiseaseTable = diseaseSheet.ListObjects(1)
End Function

'@sub-title Retrieve a ListObject from a named worksheet.
Private Function ResolveListObject(ByVal workbook As Workbook, ByVal sheetName As String) As ListObject
    Dim sheet As Worksheet
    Set sheet = workbook.Worksheets(sheetName)
    If sheet.ListObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, "Worksheet '" & sheetName & "' is missing its ListObject"
    End If
    Set ResolveListObject = sheet.ListObjects(1)
End Function

'@sub-title Enumerate worksheets from provided disease names.
Private Function ResolveSheets(ByVal workbook As Workbook, ByVal diseaseNames As BetterArray) As Collection
    Dim results As New Collection
    Dim index As Long
    Dim diseaseName As String
    Dim sheet As Worksheet

    For index = diseaseNames.LowerBound To diseaseNames.UpperBound
        diseaseName = CStr(diseaseNames.Item(index))
        Set sheet = FindWorksheet(workbook, diseaseName)
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

'@sub-title Populate dictionary sheet using array-based assignment.
Private Sub PopulateDictionary(ByVal dictionarySheet As Worksheet, ByVal diseaseTable As ListObject)

    Dim sourceData As Variant
    Dim prepared As Variant
    Dim rowCount As Long

    dictionarySheet.Range("A1").Resize(1, 6).Value = ArrayToRow( _
        Array(HEADER_VARIABLE_ORDER, HEADER_VARIABLE_SECTION, HEADER_VARIABLE_NAME, _
              HEADER_MAIN_LABEL, HEADER_CONTROL, HEADER_STATUS))

    If diseaseTable.DataBodyRange Is Nothing Then Exit Sub

    sourceData = diseaseTable.DataBodyRange.Value
    rowCount = UBound(sourceData, 1)
    prepared = BuildDictionaryMatrix(sourceData, rowCount)

    If Not IsEmpty(prepared) Then
        dictionarySheet.Range("A2").Resize(rowCount, 6).Value = prepared
    End If
End Sub

'@sub-title Build a dictionary matrix selecting required columns.
Private Function BuildDictionaryMatrix(ByRef sourceData As Variant, ByVal rowCount As Long) As Variant
    Dim matrix() As Variant
    Dim rowIndex As Long

    ReDim matrix(1 To rowCount, 1 To 6)

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

'@sub-title Populate choices sheet translating headers when available.
Private Sub PopulateChoices(ByVal choicesSheet As Worksheet, _
                            ByVal diseaseTable As ListObject, _
                            ByVal headerCache As IDiseaseTranslationCache)

    Dim choiceRows As Collection
    Dim choiceMatrix As Variant

    choicesSheet.Range("A1").Resize(1, 4).Value = ArrayToRow(Array( _
        TranslateHeader(headerCache, HEADER_LIST_NAME, HEADER_LIST_NAME), _
        TranslateHeader(headerCache, HEADER_LABEL, HEADER_LABEL), _
        TranslateHeader(headerCache, HEADER_SHORT_LABEL, HEADER_SHORT_LABEL), _
        TranslateHeader(headerCache, HEADER_ORDER, HEADER_ORDER) _
    ))

    If diseaseTable.DataBodyRange Is Nothing Then Exit Sub

    Set choiceRows = BuildChoiceRows(diseaseTable.DataBodyRange.Value)
    choiceMatrix = CollectionToMatrix(choiceRows, 4)

    If Not IsEmpty(choiceMatrix) Then
        choicesSheet.Range("A2").Resize(UBound(choiceMatrix, 1), UBound(choiceMatrix, 2)).Value = choiceMatrix
    End If
End Sub

'@sub-title Build a collection containing distinct choice rows.
Private Function BuildChoiceRows(ByRef sourceData As Variant) As Collection
    Dim rows As New Collection
    Dim dedupe As BetterArray
    Dim rowIndex As Long
    Dim splitValues As Variant
    Dim itemIndex As Long
    Dim controlName As String
    Dim choiceValue As String
    Dim key As String
    Dim choiceRow(1 To 4) As Variant

    Set dedupe = New BetterArray

    For rowIndex = 1 To UBound(sourceData, 1)
        controlName = Trim$(CStr(sourceData(rowIndex, DiseaseControlColumn)))
        If LenB(controlName) = 0 Then GoTo NextRow

        splitValues = SplitChoices(sourceData(rowIndex, DiseaseChoicesColumn))
        If IsEmpty(splitValues) Then GoTo NextRow

        For itemIndex = LBound(splitValues) To UBound(splitValues)
            choiceValue = Trim$(CStr(splitValues(itemIndex)))
            If LenB(choiceValue) = 0 Then GoTo NextChoice

            key = KeyOf(controlName, choiceValue)
            If Not ContainsKey(dedupe, key) Then
                dedupe.Push key
                choiceRow(1) = controlName
                choiceRow(2) = choiceValue
                choiceRow(3) = choiceValue
                choiceRow(4) = itemIndex - LBound(splitValues) + 1
                rows.Add choiceRow
            End If

NextChoice:
        Next itemIndex

NextRow:
    Next rowIndex

    Set BuildChoiceRows = rows
End Function

'@sub-title Determine whether a key has already been processed.
Private Function ContainsKey(ByVal keys As BetterArray, ByVal key As String) As Boolean
    Dim idx As Long
    Dim normalised As String

    If keys Is Nothing Then Exit Function
    If keys.Length = 0 Then Exit Function

    normalised = LCase$(key)

    For idx = keys.LowerBound To keys.UpperBound
        If StrComp(CStr(keys.Item(idx)), normalised, vbBinaryCompare) = 0 Then
            ContainsKey = True
            Exit Function
        End If
    Next idx
End Function

'@sub-title Locate a worksheet by name without relying on implicit errors.
Private Function FindWorksheet(ByVal workbook As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet

    For Each sheet In workbook.Worksheets
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

'@sub-title Populate a worksheet with a ListObject using a single array assignment.
Private Sub PopulateTranslations(ByVal targetSheet As Worksheet, ByVal sourceTable As ListObject)
    Dim sourceRange As Range
    Dim values As Variant

    If sourceTable Is Nothing Then Exit Sub

    Set sourceRange = sourceTable.Range
    If sourceRange Is Nothing Then Exit Sub

    values = sourceRange.Value
    targetSheet.Range("A1").Resize(UBound(values, 1), UBound(values, 2)).Value = values
End Sub

'@sub-title Populate a migration worksheet with disease data.
Private Sub PopulateMigrationDisease(ByVal targetSheet As Worksheet, ByVal diseaseSheet As Worksheet)
    Dim diseaseTable As ListObject
    Dim sourceRange As Range
    Dim values As Variant
    Dim columnOffset As Long
    Dim lastColumn As Long

    Set diseaseTable = ResolveDiseaseTable(diseaseSheet)

    Set sourceRange = diseaseTable.Range
    If sourceRange Is Nothing Then Exit Sub

    values = sourceRange.Value
    lastColumn = LastUsedColumn(targetSheet) + 2
    If lastColumn < 1 Then lastColumn = 1

    targetSheet.Cells(1, lastColumn).Value = "Disease"
    targetSheet.Cells(1, lastColumn + 1).Value = "Language"
    targetSheet.Cells(1, lastColumn + 2).Value = "Tag"

    targetSheet.Cells(2, lastColumn).Value = diseaseSheet.Name
    targetSheet.Cells(2, lastColumn + 1).Value = diseaseSheet.Cells(2, 2).Value
    targetSheet.Cells(2, lastColumn + 2).Value = diseaseSheet.Cells(2, 3).Value

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
Private Sub ApplyMigrationFormatting(ByVal workbook As Workbook)
    AutoFitSheet workbook.Worksheets(DISEASES_SHEET)
    AutoFitSheet workbook.Worksheets(TRANSLATIONS_SHEET)
    AutoFitSheet workbook.Worksheets(CHOICES_SHEET)
    AutoFitSheet workbook.Worksheets(VARIABLES_SHEET)
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
Private Function EnsureWorksheet(ByVal workbook As Workbook, _
                                 ByVal sheetName As String, _
                                 Optional ByVal template As Worksheet = Nothing) As Worksheet

    Dim sheet As Worksheet

    Set sheet = FindWorksheet(workbook, sheetName)

    If sheet Is Nothing Then
        If Not template Is Nothing Then
            template.Name = sheetName
            Set sheet = template
        Else
            Set sheet = workbook.Worksheets.Add(After:=workbook.Worksheets(workbook.Worksheets.Count))
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

'@sub-title Convert one-dimensional array into a 1xN matrix for assignment.
Private Function ArrayToRow(ByVal values As Variant) As Variant
    Dim columnCount As Long
    Dim row(1 To 1, 1 To 1) As Variant
    Dim index As Long

    columnCount = UBound(values) - LBound(values) + 1
    ReDim row(1 To 1, 1 To columnCount)

    For index = LBound(values) To UBound(values)
        row(1, index - LBound(values) + 1) = values(index)
    Next index

    ArrayToRow = row
End Function

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

'@sub-title Translate headers when the translation object is provided.
Private Function TranslateHeader(ByVal cache As IDiseaseTranslationCache, _
                                 ByVal key As String, _
                                 ByVal fallback As String) As String
    If cache Is Nothing Then
        TranslateHeader = fallback
    Else
        TranslateHeader = cache.Resolve(key, fallback)
    End If
End Function

'@sub-title Build a unique key for deduplication.
Private Function KeyOf(ByVal controlName As String, ByVal choiceValue As String) As String
    KeyOf = LCase$(controlName & "|" & choiceValue)
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

'@section Interface
'===============================================================================

Private Sub IDiseaseExporter_Initialise(ByVal workbookManager As IDiseaseExportWorkbook, ByVal appGuard As IDiseaseApplicationState)
    Initialise workbookManager, appGuard
End Sub

Private Function IDiseaseExporter_BuildDiseaseWorkbook(ByVal diseaseSheet As Worksheet, ByVal translationTable As ListObject, ByVal ribbonTranslations As ITranslationObject, ByVal diseaseName As String, ByVal languageTag As String, Optional ByVal diseaseCode As String = vbNullString) As Workbook
    Set IDiseaseExporter_BuildDiseaseWorkbook = BuildDiseaseWorkbook(diseaseSheet, translationTable, ribbonTranslations, diseaseName, languageTag, diseaseCode)
End Function

Private Function IDiseaseExporter_ExportDisease(ByVal exportFolder As String, ByVal diseaseSheet As Worksheet, ByVal translationTable As ListObject, ByVal ribbonTranslations As ITranslationObject, ByVal diseaseName As String, ByVal languageTag As String, Optional ByVal diseaseCode As String = vbNullString) As String
    IDiseaseExporter_ExportDisease = ExportDisease(exportFolder, diseaseSheet, translationTable, ribbonTranslations, diseaseName, languageTag, diseaseCode)
End Function

Private Function IDiseaseExporter_BuildMigrationWorkbook(ByVal sourceWorkbook As Workbook, ByVal diseaseNames As BetterArray) As Workbook
    Set IDiseaseExporter_BuildMigrationWorkbook = BuildMigrationWorkbook(sourceWorkbook, diseaseNames)
End Function

Private Function IDiseaseExporter_ExportForMigration(ByVal exportFolder As String, ByVal sourceWorkbook As Workbook, ByVal diseaseNames As BetterArray) As String
    IDiseaseExporter_ExportForMigration = ExportForMigration(exportFolder, sourceWorkbook, diseaseNames)
End Function
