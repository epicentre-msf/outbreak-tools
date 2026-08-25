VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MasterSetupImportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Folds an exported disease workbook or a migration block back into the master setup"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Folds an exported disease workbook or a migration block back into the master setup")
'@depends DiseaseSheet, DiseaseImporter, DiseaseImportSummary, DiseaseLogger, DiseaseWorksheetManager, MasterSetupVariables, LLChoices, DropdownLists, HiddenNames, BetterArray
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'MasterSetupImportService.cls without interface -- single implementation, no
'test fake, no compile cycle.
'
'Two doors, both landing a disease table on a disease worksheet of the
'master setup workbook:
'- ImportSetupExport reads a workbook DiseaseExporter wrote for a setup. Its
'  Metadata sheet names the disease, the language and the code, its
'  Tab_Dictionary carries the variables and its Tab_Choices carries the
'  lists. A variable the Variables table lacks is added to it, a list the
'  Choices sheet lacks is added to it, and the disease worksheet is rebuilt
'  when it is gone or merged when it is there, with the file winning.
'- ImportDiseaseTable lands one table that already has the shape of a
'  disease table. MasterSetupExports stages every block of a migration file
'  that way.
'
'THE DISEASE NAME
'-------------------------------------------------------------------------------
'The name comes from the workbook tag the export writes, then from the
'Metadata sheet, then from the caller. A file carrying none and given none
'is refused; the ribbon door asks the user and checks the answer with
'DiseaseNameIsFree before passing it in.
'
'CHOICE LABELS AND THE LANGUAGE
'-------------------------------------------------------------------------------
'The master Choices sheet holds its labels in English. An exported file
'carries them in the language of the disease, so a list added from a file
'in another language takes its name and its rows with every label left
'empty, for the user to fill in English. A tag starting with "en" counts as
'English.

'@section Constants
'===============================================================================

Private Const METADATA_SHEET As String = "Metadata"
Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const CHOICES_SHEET As String = "Choices"
Private Const TABLE_DICTIONARY As String = "Tab_Dictionary"
Private Const TABLE_CHOICES As String = "Tab_Choices"
Private Const META_DISEASE As String = "export_disease"
Private Const META_LANGUAGE As String = "dictionary_language"
Private Const META_CODE As String = "dictionary_code"
'The tags DiseaseExporter writes at workbook level; read before the sheet.
Private Const NAME_DISNAME As String = "__Var_DISNAME"
Private Const NAME_DISLANG As String = "__Var_DISLANG"
Private Const NAME_DISCODE As String = "__Var_DISCODE"
Private Const PROHIBITED_DROPDOWN As String = "__prohibited_diseases_list"
'The Metadata sheet carries five lines; reading a few more costs nothing.
Private Const META_ROWS_READ As Long = 20

'Headers of the exported dictionary, the ones a setup dictionary carries.
Private Const HEADER_VARIABLE_NAME As String = "Variable Name"
Private Const HEADER_MAIN_SECTION As String = "Main Section"
Private Const HEADER_MAIN_LABEL As String = "Main Label"
Private Const HEADER_CONTROL As String = "Control"
Private Const HEADER_STATUS As String = "Status"
Private Const HEADER_VARIABLE_ORDER As String = "Variable Order"

'Headers of the exported choices, the ones a setup Choices table carries.
Private Const HEADER_LIST_NAME As String = "list name"
Private Const HEADER_LABEL As String = "label"
Private Const HEADER_TRANSLATED_LABEL As String = "translated label"
Private Const HEADER_RAW_LABEL As String = "non translated label"
Private Const HEADER_SHORT_LABEL As String = "short label"

'Headers of a disease table, as DiseaseSheet writes them.
Private Const DISEASE_HEADER_ORDER As String = "Variable Order"
Private Const DISEASE_HEADER_NAME As String = "Variable Name"
Private Const DISEASE_HEADER_SECTION As String = "Variable Section"
Private Const DISEASE_HEADER_LABEL As String = "Main Label"
Private Const DISEASE_HEADER_CHOICE As String = "Choice"
Private Const DISEASE_HEADER_VALUES As String = "Choice Values"
Private Const DISEASE_HEADER_STATUS As String = "Status"

Private Const IMPORT_STAGING_SHEET As String = "__dis_import"
Private Const LANGUAGES_DROPDOWN As String = "__data_languages"
Private Const DISEASES_DROPDOWN As String = "__diseases_list"
Private Const SHEET_TAG_NAME As String = "sheetTag"
Private Const DISEASE_TAG_VALUE As String = "disease"
Private Const ENGLISH_TAG_PREFIX As String = "en"
Private Const LOG_OPERATION As String = "ImportSetup"
Private Const DISEASE_TABLE_WIDTH As Long = 7

'Column order of a disease table, as DiseaseSheet builds it.
Private Enum DiseaseColumns
    DiseaseOrderColumn = 1
    DiseaseNameColumn = 2
    DiseaseSectionColumn = 3
    DiseaseLabelColumn = 4
    DiseaseControlColumn = 5
    DiseaseChoicesColumn = 6
    DiseaseStatusColumn = 7
End Enum

'@section State
'===============================================================================

Private Type TImportService
    targetBook As Workbook
    builder As DiseaseSheet
    dropdowns As DropdownLists
    variables As MasterSetupVariables
    choices As LLChoices
    importer As DiseaseImporter
    addedVariables As BetterArray
    addedChoices As BetterArray
    diseaseName As String
    languageTag As String
    diseaseCode As String
End Type

Private this As TImportService

'@section Factory
'===============================================================================

'@sub-title Build a service over the master setup workbook and its managers.
'@param targetBook Workbook. The master setup workbook receiving the import.
'@param builder DiseaseSheet. Builds a disease sheet that is gone and puts
'   the line formulas back after a merge.
'@param dropdowns DropdownLists. The master dropdowns: the languages list
'   and the diseases list.
'@param variables MasterSetupVariables. The manager of the Variables table.
'@param choices Optional LLChoices. The master Choices sheet; with none the
'   lists of a file are left out and logged.
'@param importer Optional DiseaseImporter. Defaults to a fresh importer.
'@return MasterSetupImportService ready to import.
Public Function Create(ByVal targetBook As Workbook, _
                       ByVal builder As DiseaseSheet, _
                       ByVal dropdowns As DropdownLists, _
                       ByVal variables As MasterSetupVariables, _
                       Optional ByVal choices As LLChoices = Nothing, _
                       Optional ByVal importer As DiseaseImporter = Nothing) As MasterSetupImportService

    Dim service As MasterSetupImportService
    Set service = New MasterSetupImportService
    service.Initialise targetBook, builder, dropdowns, variables, choices, importer
    Set Create = service
End Function

'@section Configuration
'===============================================================================

'@sub-title Attach the workbook and the managers the imports write through.
Public Sub Initialise(ByVal targetBook As Workbook, _
                      ByVal builder As DiseaseSheet, _
                      ByVal dropdowns As DropdownLists, _
                      ByVal variables As MasterSetupVariables, _
                      Optional ByVal choices As LLChoices = Nothing, _
                      Optional ByVal importer As DiseaseImporter = Nothing)

    If targetBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target workbook cannot be Nothing"
    End If
    If builder Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Disease sheet builder cannot be Nothing"
    End If
    If dropdowns Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Dropdown lists cannot be Nothing"
    End If
    If variables Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Variables manager cannot be Nothing"
    End If

    Set this.targetBook = targetBook
    Set this.builder = builder
    Set this.dropdowns = dropdowns
    Set this.variables = variables
    Set this.choices = choices

    If importer Is Nothing Then
        Set this.importer = DiseaseImporter.Create()
    Else
        Set this.importer = importer
    End If

    ResetTallies
End Sub

'@section Setup Export Import
'===============================================================================

'@sub-title Fold a workbook exported for a setup back into the master setup.
'@details The Metadata sheet is read first, then the Variables table takes
'the variables it lacks, the Choices sheet takes the lists it lacks, and
'the dictionary is staged as a disease table on a hidden sheet and landed
'on the disease worksheet. The staging sheet goes away on every path.
'@param sourceBook Workbook. The exported workbook, open.
'@param logger Optional DiseaseLogger. Receives one line per variable and
'   list added, and one warning per thing left out.
'@param diseaseName Optional String. The name to use when the file carries
'   none; it also wins over the one the file carries when given.
'@return DiseaseImportSummary of the merge, or Nothing when no disease
'   worksheet could take the block.
Public Function ImportSetupExport(ByVal sourceBook As Workbook, _
                                  Optional ByVal logger As DiseaseLogger = Nothing, _
                                  Optional ByVal diseaseName As String = vbNullString) As DiseaseImportSummary

    Dim dictionaryTable As ListObject
    Dim choicesTable As ListObject
    Dim stagingTable As ListObject
    Dim manager As DiseaseWorksheetManager
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If sourceBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source workbook cannot be Nothing"
    End If

    ResetTallies
    ReadMetadata sourceBook, diseaseName

    Set dictionaryTable = FindTable(sourceBook, TABLE_DICTIONARY, DICTIONARY_SHEET)
    If dictionaryTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "The workbook carries no '" & TABLE_DICTIONARY & "' table"
    End If
    Set choicesTable = FindTable(sourceBook, TABLE_CHOICES, CHOICES_SHEET)

    Set manager = DiseaseWorksheetManager.Create()

    On Error GoTo Handler

    UpdateVariables dictionaryTable, logger
    UpdateChoices choicesTable, logger
    RefreshAddedVariables

    manager.RemoveWorksheet this.targetBook, IMPORT_STAGING_SHEET
    Set stagingTable = BuildStagingTable(dictionaryTable)

    Set ImportSetupExport = ImportDiseaseTable(stagingTable, this.diseaseName, this.languageTag, _
                                               restoreFormulas:=True, logger:=logger)

Cleanup:
    'Shielded: the staging sheet goes away whatever happened above, and a
    'raise from the removal would come straight back here.
    On Error Resume Next
    manager.RemoveWorksheet this.targetBook, IMPORT_STAGING_SHEET
    On Error GoTo 0

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If
    Exit Function

Handler:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume Cleanup
End Function

'@section Disease Table Import
'===============================================================================

'@sub-title Land a table shaped as a disease table on the disease worksheet.
'@details A worksheet of that name that is gone is built again, on the
'language given when the languages list carries it and on the default
'language otherwise; the block then replaces its table whole. A worksheet
'that is there is merged, with the imported values winning on the lines
'both carry. A worksheet of that name that is no disease worksheet is left
'alone and logged.
'@param sourceTable ListObject. Seven columns in the disease table order.
'@param diseaseName String. Name of the disease worksheet.
'@param languageTag String. Language the sheet takes when it is built.
'@param restoreFormulas Optional Boolean. When True the label and the choice
'   values formulas go back on every line once the values have landed,
'   which is what a setup export wants: the master setup answers those two
'   columns itself. A migration block keeps the values it carries.
'@param logger Optional DiseaseLogger.
'@return DiseaseImportSummary of the merge, or Nothing when the block was
'   left out.
Public Function ImportDiseaseTable(ByVal sourceTable As ListObject, _
                                   ByVal diseaseName As String, _
                                   ByVal languageTag As String, _
                                   Optional ByVal restoreFormulas As Boolean = False, _
                                   Optional ByVal logger As DiseaseLogger = Nothing) As DiseaseImportSummary

    Dim targetSheet As Worksheet
    Dim freshSheet As Boolean

    If sourceTable Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source table cannot be Nothing"
    End If

    diseaseName = Trim$(diseaseName)
    If LenB(diseaseName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Disease name cannot be empty"
    End If

    Set targetSheet = FindWorksheet(this.targetBook, diseaseName)
    freshSheet = targetSheet Is Nothing

    If freshSheet Then
        Set targetSheet = BuildDiseaseSheet(diseaseName, languageTag, logger)
    ElseIf Not IsDiseaseSheet(targetSheet) Then
        Warn logger, "The worksheet '" & diseaseName & "' is no disease worksheet; the block is left out"
        Exit Function
    End If

    If targetSheet.ListObjects.Count = 0 Then
        Warn logger, "The worksheet '" & diseaseName & "' carries no table; the block is left out"
        Exit Function
    End If

    'A fresh sheet takes the block whole; an existing one merges it, with the
    'imported values winning.
    Set ImportDiseaseTable = this.importer.MergeDisease(targetSheet.ListObjects(1), sourceTable, _
                                                        mergeValues:=Not freshSheet, _
                                                        priority:=DiseaseImportPriority_Foreign, _
                                                        logger:=logger)

    If restoreFormulas Then this.builder.WriteLineFormulas targetSheet.ListObjects(1)
End Function

'@section Name Resolution
'===============================================================================

'@sub-title The disease name a file carries: its workbook tag, then its Metadata sheet.
'@return String, empty when the file names no disease.
Public Function ReadDiseaseName(ByVal sourceBook As Workbook) As String
    If sourceBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source workbook cannot be Nothing"
    End If

    ReadDiseaseName = Trim$(TagValue(sourceBook, NAME_DISNAME))
    If LenB(ReadDiseaseName) > 0 Then Exit Function

    ReadDiseaseName = MetadataValue(sourceBook, META_DISEASE)
End Function

'@sub-title Whether a name can take a new disease worksheet in the target workbook.
'@details Empty names, names of a worksheet already there and names of the
'prohibited list are refused. A name already there would merge the file
'into a sheet the user did not name.
Public Function DiseaseNameIsFree(ByVal candidate As String) As Boolean
    Dim prohibited As BetterArray

    candidate = Trim$(candidate)
    If LenB(candidate) = 0 Then Exit Function
    If Not FindWorksheet(this.targetBook, candidate) Is Nothing Then Exit Function

    Set prohibited = this.dropdowns.Values(PROHIBITED_DROPDOWN)
    If Not prohibited Is Nothing Then
        If prohibited.Includes(candidate) Then Exit Function
    End If

    DiseaseNameIsFree = True
End Function

'@section Results
'===============================================================================

'@sub-title Names of the variables the last import added to the Variables table.
Public Property Get AddedVariables() As BetterArray
    Set AddedVariables = this.addedVariables.Clone()
End Property

'@sub-title Names of the lists the last import added to the Choices sheet.
Public Property Get AddedChoices() As BetterArray
    Set AddedChoices = this.addedChoices.Clone()
End Property

'@sub-title Disease name read from the Metadata sheet of the last file.
Public Property Get DiseaseName() As String
    DiseaseName = this.diseaseName
End Property

'@sub-title Language read from the Metadata sheet of the last file.
Public Property Get LanguageTag() As String
    LanguageTag = this.languageTag
End Property

'@sub-title Disease code read from the Metadata sheet of the last file.
Public Property Get DiseaseCode() As String
    DiseaseCode = this.diseaseCode
End Property

'@section Metadata
'===============================================================================

'@sub-title Settle the disease name, the language and the code of the file.
'@details The workbook tags come first, the Metadata sheet fills what they
'leave empty, and a name given by the caller wins. A file with no name from
'any of the three is refused.
Private Sub ReadMetadata(ByVal sourceBook As Workbook, ByVal givenName As String)
    this.diseaseName = Trim$(TagValue(sourceBook, NAME_DISNAME))
    this.languageTag = Trim$(TagValue(sourceBook, NAME_DISLANG))
    this.diseaseCode = Trim$(TagValue(sourceBook, NAME_DISCODE))

    If LenB(this.diseaseName) = 0 Then this.diseaseName = MetadataValue(sourceBook, META_DISEASE)
    If LenB(this.languageTag) = 0 Then this.languageTag = MetadataValue(sourceBook, META_LANGUAGE)
    If LenB(this.diseaseCode) = 0 Then this.diseaseCode = MetadataValue(sourceBook, META_CODE)

    If LenB(Trim$(givenName)) > 0 Then this.diseaseName = Trim$(givenName)

    If LenB(this.diseaseName) = 0 Then
        ThrowError ProjectError.ElementNotFound, "The workbook carries no disease name and none was given"
    End If
End Sub

'@sub-title A workbook tag; a file with no hidden names store answers empty.
Private Function TagValue(ByVal sourceBook As Workbook, ByVal nameId As String) As String
    Dim store As HiddenNames

    'A foreign workbook can carry no store at all, and the store raises on it.
    On Error Resume Next
        Set store = HiddenNames.Create(sourceBook)
        If Not store Is Nothing Then TagValue = store.ValueAsString(nameId)
    On Error GoTo 0
End Function

'@sub-title One value of the Metadata sheet, read by its label; empty without the sheet.
'@details The Metadata sheet is a two-column block of labels and values,
'read by label so the order of the lines carries no weight.
Private Function MetadataValue(ByVal sourceBook As Workbook, ByVal metaLabel As String) As String
    Dim metadataSheet As Worksheet
    Dim block As Variant
    Dim rowIndex As Long

    Set metadataSheet = FindWorksheet(sourceBook, METADATA_SHEET)
    If metadataSheet Is Nothing Then Exit Function

    block = metadataSheet.Range("A1").Resize(META_ROWS_READ, 2).Value

    For rowIndex = 1 To META_ROWS_READ
        If StrComp(Trim$(TextOf(block(rowIndex, 1))), metaLabel, vbTextCompare) = 0 Then
            MetadataValue = Trim$(TextOf(block(rowIndex, 2)))
            Exit Function
        End If
    Next rowIndex
End Function

'@section Variables
'===============================================================================

'@sub-title Add to the Variables table every variable of the dictionary it lacks.
'@details One read of the dictionary block serves every line. A name the
'table carries, or one the file names twice, is passed over.
Private Sub UpdateVariables(ByVal dictionaryTable As ListObject, ByVal logger As DiseaseLogger)
    Dim block As Variant
    Dim nameColumn As Long
    Dim sectionColumn As Long
    Dim labelColumn As Long
    Dim controlColumn As Long
    Dim statusColumn As Long
    Dim rowIndex As Long
    Dim variableName As String

    block = TableBlock(dictionaryTable)

    nameColumn = HeaderColumn(block, HEADER_VARIABLE_NAME)
    If nameColumn = 0 Then
        ThrowError ProjectError.ElementNotFound, "The dictionary carries no '" & HEADER_VARIABLE_NAME & "' column"
    End If
    sectionColumn = HeaderColumn(block, HEADER_MAIN_SECTION)
    labelColumn = HeaderColumn(block, HEADER_MAIN_LABEL)
    controlColumn = HeaderColumn(block, HEADER_CONTROL)
    statusColumn = HeaderColumn(block, HEADER_STATUS)

    For rowIndex = 2 To UBound(block, 1)
        variableName = Trim$(TextOf(block(rowIndex, nameColumn)))
        If LenB(variableName) = 0 Then GoTo NextRow
        If this.variables.HasVariable(variableName) Then GoTo NextRow
        If ContainsKey(this.addedVariables, variableName) Then GoTo NextRow

        this.variables.AddVariable variableName, _
                                   CellAt(block, rowIndex, sectionColumn), _
                                   CellAt(block, rowIndex, labelColumn), _
                                   CellAt(block, rowIndex, controlColumn), _
                                   CellAt(block, rowIndex, statusColumn)
        this.addedVariables.Push variableName
        Info logger, "Added the variable '" & variableName & "' to the Variables table"
NextRow:
    Next rowIndex
End Sub

'@sub-title Fill the joined choices values of the variables just added.
'@details Only a default choice the Choices sheet carried before this import
'is joined: the labels of a list added by the same import sit where the
'user types them, and the label column follows once the sheet is edited.
Private Sub RefreshAddedVariables()
    Dim index As Long
    Dim variableName As String
    Dim defaultChoice As String

    If this.choices Is Nothing Then Exit Sub

    For index = this.addedVariables.LowerBound To this.addedVariables.UpperBound
        variableName = CStr(this.addedVariables.Item(index))
        defaultChoice = this.variables.DefaultChoiceFor(variableName)
        If LenB(defaultChoice) = 0 Then GoTo NextVariable
        If ContainsKey(this.addedChoices, defaultChoice) Then GoTo NextVariable

        this.variables.RefreshChoices variableName, this.choices
NextVariable:
    Next index
End Sub

'@section Choices
'===============================================================================

'@sub-title Add to the Choices sheet every list of the file it lacks.
'@details The labels travel only when the file is in English; otherwise the
'list lands with its name and its rows and every label empty. The choices
'ListObject of the master sheet, when there is one, is stretched over the
'rows added so the event handlers keep reading them.
Private Sub UpdateChoices(ByVal choicesTable As ListObject, ByVal logger As DiseaseLogger)
    Dim block As Variant
    Dim listNameColumn As Long
    Dim labelColumn As Long
    Dim translatedColumn As Long
    Dim rawColumn As Long
    Dim shortColumn As Long
    Dim listNames As BetterArray
    Dim listIndex As Long
    Dim listName As String
    Dim labels As BetterArray
    Dim shortLabels As BetterArray
    Dim keepLabels As Boolean

    If choicesTable Is Nothing Then Exit Sub

    If this.choices Is Nothing Then
        Warn logger, "The master setup carries no Choices sheet; the lists of the file are left out"
        Exit Sub
    End If

    block = TableBlock(choicesTable)
    listNameColumn = HeaderColumn(block, HEADER_LIST_NAME)
    If listNameColumn = 0 Then
        Warn logger, "The choices table carries no '" & HEADER_LIST_NAME & "' column; the lists are left out"
        Exit Sub
    End If
    labelColumn = HeaderColumn(block, HEADER_LABEL)
    translatedColumn = HeaderColumn(block, HEADER_TRANSLATED_LABEL)
    rawColumn = HeaderColumn(block, HEADER_RAW_LABEL)
    shortColumn = HeaderColumn(block, HEADER_SHORT_LABEL)
    keepLabels = IsEnglish(this.languageTag)

    Set listNames = DistinctValues(block, listNameColumn)

    For listIndex = listNames.LowerBound To listNames.UpperBound
        listName = CStr(listNames.Item(listIndex))
        If this.choices.ChoiceExists(listName) Then GoTo NextList

        Set labels = New BetterArray
        labels.LowerBound = 1
        Set shortLabels = New BetterArray
        shortLabels.LowerBound = 1

        CollectListLabels block, listNameColumn, listName, labelColumn, translatedColumn, rawColumn, _
                          shortColumn, keepLabels, labels, shortLabels
        If labels.Length = 0 Then GoTo NextList

        this.choices.AddChoice listName, labels, shortLabels

        If this.choices.ChoiceExists(listName) Then
            this.addedChoices.Push listName
            If keepLabels Then
                Info logger, "Added the list '" & listName & "' to the Choices sheet"
            Else
                Info logger, "Added the list '" & listName & "' to the Choices sheet with empty labels: " & _
                             "the file is in '" & this.languageTag & "' and the Choices sheet is in English"
            End If
        Else
            Warn logger, "The list '" & listName & "' could not be added to the Choices sheet; check its headers"
        End If
NextList:
    Next listIndex

    If this.addedChoices.Length > 0 Then StretchChoicesTable
End Sub

'@sub-title Gather the labels of one list, in the order the file carries them.
'@details The label of a row is the label column when filled, then the
'translated label, then the non translated label. The short label falls
'back on the label. With keepLabels False both come out empty.
Private Sub CollectListLabels(ByRef block As Variant, _
                              ByVal listNameColumn As Long, _
                              ByVal listName As String, _
                              ByVal labelColumn As Long, _
                              ByVal translatedColumn As Long, _
                              ByVal rawColumn As Long, _
                              ByVal shortColumn As Long, _
                              ByVal keepLabels As Boolean, _
                              ByVal labels As BetterArray, _
                              ByVal shortLabels As BetterArray)
    Dim rowIndex As Long
    Dim labelValue As String
    Dim shortValue As String

    For rowIndex = 2 To UBound(block, 1)
        If StrComp(Trim$(TextOf(block(rowIndex, listNameColumn))), listName, vbTextCompare) <> 0 Then GoTo NextRow

        labelValue = CellAt(block, rowIndex, labelColumn)
        If LenB(labelValue) = 0 Then labelValue = CellAt(block, rowIndex, translatedColumn)
        If LenB(labelValue) = 0 Then labelValue = CellAt(block, rowIndex, rawColumn)
        If LenB(labelValue) = 0 Then GoTo NextRow

        shortValue = CellAt(block, rowIndex, shortColumn)
        If LenB(shortValue) = 0 Then shortValue = labelValue

        If keepLabels Then
            labels.Push labelValue
            shortLabels.Push shortValue
        Else
            labels.Push vbNullString
            shortLabels.Push vbNullString
        End If
NextRow:
    Next rowIndex
End Sub

'@sub-title Stretch the choices ListObject over the rows AddChoice wrote under it.
'@details LLChoices writes below the last filled line and knows nothing of
'the table; the master event handlers read the sheet through the table.
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

'@section Staging
'===============================================================================

'@sub-title Write the dictionary as a disease table on a hidden staging sheet.
'@details Seven columns in the disease table order. The order column takes
'the dictionary order when it is a number and the position of the line
'otherwise; the choice values column stays empty, its formula comes back
'once the block has landed. A dictionary with no named line stages a
'header-only table.
Private Function BuildStagingTable(ByVal dictionaryTable As ListObject) As ListObject
    Dim block As Variant
    Dim matrix() As Variant
    Dim nameColumn As Long
    Dim orderColumn As Long
    Dim sectionColumn As Long
    Dim labelColumn As Long
    Dim controlColumn As Long
    Dim statusColumn As Long
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim written As Long
    Dim variableName As String
    Dim stagingSheet As Worksheet
    Dim tableRange As Range

    block = TableBlock(dictionaryTable)
    nameColumn = HeaderColumn(block, HEADER_VARIABLE_NAME)
    If nameColumn = 0 Then Exit Function

    orderColumn = HeaderColumn(block, HEADER_VARIABLE_ORDER)
    sectionColumn = HeaderColumn(block, HEADER_MAIN_SECTION)
    labelColumn = HeaderColumn(block, HEADER_MAIN_LABEL)
    controlColumn = HeaderColumn(block, HEADER_CONTROL)
    statusColumn = HeaderColumn(block, HEADER_STATUS)

    rowCount = CountNamedRows(block, nameColumn)
    ReDim matrix(1 To rowCount + 1, 1 To DISEASE_TABLE_WIDTH)

    matrix(1, DiseaseOrderColumn) = DISEASE_HEADER_ORDER
    matrix(1, DiseaseNameColumn) = DISEASE_HEADER_NAME
    matrix(1, DiseaseSectionColumn) = DISEASE_HEADER_SECTION
    matrix(1, DiseaseLabelColumn) = DISEASE_HEADER_LABEL
    matrix(1, DiseaseControlColumn) = DISEASE_HEADER_CHOICE
    matrix(1, DiseaseChoicesColumn) = DISEASE_HEADER_VALUES
    matrix(1, DiseaseStatusColumn) = DISEASE_HEADER_STATUS

    For rowIndex = 2 To UBound(block, 1)
        variableName = Trim$(TextOf(block(rowIndex, nameColumn)))
        If LenB(variableName) = 0 Then GoTo NextRow

        written = written + 1
        matrix(written + 1, DiseaseOrderColumn) = OrderValue(block, rowIndex, orderColumn, written)
        matrix(written + 1, DiseaseNameColumn) = variableName
        matrix(written + 1, DiseaseSectionColumn) = CellAt(block, rowIndex, sectionColumn)
        matrix(written + 1, DiseaseLabelColumn) = CellAt(block, rowIndex, labelColumn)
        matrix(written + 1, DiseaseControlColumn) = CellAt(block, rowIndex, controlColumn)
        matrix(written + 1, DiseaseChoicesColumn) = vbNullString
        matrix(written + 1, DiseaseStatusColumn) = CellAt(block, rowIndex, statusColumn)
NextRow:
    Next rowIndex

    Set stagingSheet = this.targetBook.Worksheets.Add(After:=this.targetBook.Worksheets(this.targetBook.Worksheets.Count))
    stagingSheet.Name = IMPORT_STAGING_SHEET
    stagingSheet.Visible = xlSheetHidden

    Set tableRange = stagingSheet.Range("A1").Resize(rowCount + 1, DISEASE_TABLE_WIDTH)
    tableRange.Value = matrix

    Set BuildStagingTable = stagingSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                         XlListObjectHasHeaders:=xlYes)
End Function

'@sub-title The dictionary order of a line when it is a number, its position otherwise.
Private Function OrderValue(ByRef block As Variant, ByVal rowIndex As Long, _
                            ByVal orderColumn As Long, ByVal position As Long) As Variant
    OrderValue = position
    If orderColumn = 0 Then Exit Function
    If IsError(block(rowIndex, orderColumn)) Then Exit Function
    If IsNumeric(block(rowIndex, orderColumn)) And Not IsEmpty(block(rowIndex, orderColumn)) Then
        OrderValue = block(rowIndex, orderColumn)
    End If
End Function

'@sub-title Count the lines of a block carrying a variable name.
Private Function CountNamedRows(ByRef block As Variant, ByVal nameColumn As Long) As Long
    Dim rowIndex As Long

    For rowIndex = 2 To UBound(block, 1)
        If LenB(Trim$(TextOf(block(rowIndex, nameColumn)))) > 0 Then
            CountNamedRows = CountNamedRows + 1
        End If
    Next rowIndex
End Function

'@section Disease Worksheet
'===============================================================================

'@sub-title Build a disease sheet for an imported block.
Private Function BuildDiseaseSheet(ByVal diseaseName As String, _
                                   ByVal languageTag As String, _
                                   ByVal logger As DiseaseLogger) As Worksheet
    Dim languages As BetterArray

    'A disease removed earlier can leave its name in the diseases list, and a
    'stale name refuses the rebuild as a duplicate.
    PurgeStaleDiseaseName diseaseName

    'A language the target file does not carry falls back on the default.
    Set languages = this.dropdowns.Values(LANGUAGES_DROPDOWN)
    If Not languages Is Nothing Then
        If languages.Length > 0 And LenB(languageTag) > 0 Then
            If Not languages.Includes(languageTag) Then
                Warn logger, "The language '" & languageTag & "' is not in the languages list; " & _
                             "the worksheet '" & diseaseName & "' takes the default language"
                languageTag = vbNullString
            End If
        End If
    End If

    Set BuildDiseaseSheet = this.builder.Build(diseaseName, languageTag)
End Function

'@sub-title Drop a diseases-list entry whose worksheet is gone.
Private Sub PurgeStaleDiseaseName(ByVal diseaseName As String)
    Dim names As BetterArray
    Dim kept As BetterArray
    Dim idx As Long
    Dim candidate As String

    If Not FindWorksheet(this.targetBook, diseaseName) Is Nothing Then Exit Sub

    Set names = this.dropdowns.Values(DISEASES_DROPDOWN)
    If names Is Nothing Then Exit Sub
    If Not names.Includes(diseaseName) Then Exit Sub

    Set kept = New BetterArray
    kept.LowerBound = 1

    For idx = names.LowerBound To names.UpperBound
        candidate = CStr(names.Item(idx))
        If StrComp(candidate, diseaseName, vbTextCompare) <> 0 Then
            kept.Push candidate
        End If
    Next idx

    If kept.Length = 0 Then kept.Push vbNullString

    this.dropdowns.Update kept, DISEASES_DROPDOWN
End Sub

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

'@section Lookups
'===============================================================================

'@sub-title Find a ListObject by name anywhere in the workbook, then by sheet.
'@details The exported tables carry the names a setup gives its own. A file
'edited by hand can have lost the table name and kept the sheet, so the
'first table of the named sheet is the fallback.
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

'@sub-title Read a ListObject with its header row as a two-dimensional block.
Private Function TableBlock(ByVal table As ListObject) As Variant
    Dim oneCell(1 To 1, 1 To 1) As Variant

    If table.Range.Cells.Count = 1 Then
        oneCell(1, 1) = table.Range.Value
        TableBlock = oneCell
    Else
        TableBlock = table.Range.Value
    End If
End Function

'@sub-title Column of a header in a block read with its header row, 0 when absent.
Private Function HeaderColumn(ByRef block As Variant, ByVal headerName As String) As Long
    Dim columnIndex As Long

    For columnIndex = 1 To UBound(block, 2)
        If StrComp(Trim$(TextOf(block(1, columnIndex))), headerName, vbTextCompare) = 0 Then
            HeaderColumn = columnIndex
            Exit Function
        End If
    Next columnIndex
End Function

'@sub-title Text of a block cell; a column of 0 reads as empty.
Private Function CellAt(ByRef block As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    If columnIndex = 0 Then Exit Function
    CellAt = Trim$(TextOf(block(rowIndex, columnIndex)))
End Function

'@sub-title Distinct values of one column of a block, in first-use order.
Private Function DistinctValues(ByRef block As Variant, ByVal columnIndex As Long) As BetterArray
    Dim values As BetterArray
    Dim rowIndex As Long
    Dim candidate As String

    Set values = New BetterArray
    values.LowerBound = 1

    For rowIndex = 2 To UBound(block, 1)
        candidate = Trim$(TextOf(block(rowIndex, columnIndex)))
        If LenB(candidate) = 0 Then GoTo NextRow
        If Not ContainsKey(values, candidate) Then values.Push candidate
NextRow:
    Next rowIndex

    Set DistinctValues = values
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

'@sub-title Text of a cell value; an error value or an empty cell reads as empty.
Private Function TextOf(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    TextOf = CStr(cellValue)
End Function

'@sub-title A language tag starting with "en" is English.
Private Function IsEnglish(ByVal languageTag As String) As Boolean
    IsEnglish = (Left$(LCase$(Trim$(languageTag)), Len(ENGLISH_TAG_PREFIX)) = ENGLISH_TAG_PREFIX)
End Function

'@sub-title Start the tallies of a new import.
Private Sub ResetTallies()
    Set this.addedVariables = New BetterArray
    this.addedVariables.LowerBound = 1
    Set this.addedChoices = New BetterArray
    this.addedChoices.LowerBound = 1
    this.diseaseName = vbNullString
    this.languageTag = vbNullString
    this.diseaseCode = vbNullString
End Sub

'@section Logging
'===============================================================================

Private Sub Info(ByVal logger As DiseaseLogger, ByVal message As String)
    If logger Is Nothing Then Exit Sub
    logger.Record LOG_OPERATION, DiseaseLogInfo, message
End Sub

Private Sub Warn(ByVal logger As DiseaseLogger, ByVal message As String)
    If logger Is Nothing Then Exit Sub
    logger.Record LOG_OPERATION, DiseaseLogWarning, message
End Sub

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised error raising aligned with ProjectError codes.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "MasterSetupImportService", errMessage
End Sub
