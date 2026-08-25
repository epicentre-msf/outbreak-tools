Attribute VB_Name = "MasterSetupExports"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("Disease exports, the setup import and the migration round trip for the master setup workbook.")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, ExcelMemberMayReturnNothing, UseMeaningfulName

'Four doors to the outside world, all built on the msetup classes:
'- ExportToSetup writes one disease as a dictionary workbook a setup imports.
'- ImportFromSetup reads such a workbook back: the disease worksheet is
'  rebuilt or merged, and the Variables and Choices sheets take what they
'  lack.
'- ExportForMigration writes every disease of this file into one flat
'  workbook, beside the Translations, Choices and Variables sheets.
'- ImportFlatFile reads such a flat workbook back and recreates or merges
'  every disease worksheet it carries.
'
'The landing of a block on a disease worksheet is MasterSetupImportService's
'work; this module walks the files and builds the service over the master
'managers.

Private Const DISEASES_SHEET As String = "Diseases"
Private Const IMPORT_STAGING_SHEET As String = "__dis_import"
Private Const NAME_DISLANG As String = "__Var_DISLANG"
Private Const NAME_DISCODE As String = "__Var_DISCODE"
Private Const DISEASE_TABLE_WIDTH As Long = 7
Private Const BLOCK_HEADER_ROW As Long = 1
Private Const BLOCK_META_ROW As Long = 2
Private Const BLOCK_TABLE_ROW As Long = 3
Private Const IMPORT_FILTERS As String = "*.xlsx;*.xlsb;*.xlsm"
'How many times the door asks for a disease name before it gives up.
Private Const MAX_NAME_ATTEMPTS As Long = 4

'@section Exports
'===============================================================================

'@sub-title Export the active disease worksheet as a setup dictionary workbook.
Public Sub ExportToSetup()
    Dim targetSheet As Worksheet
    Dim io As OSFiles
    Dim exporter As DiseaseExporter
    Dim translationTable As ListObject
    Dim store As HiddenNames
    Dim filePath As String

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    If Not MasterSetupHelpers.IsMasterDiseaseSheet(targetSheet) Then
        MsgBox "Select a disease worksheet before exporting it.", vbExclamation + vbOKOnly, "Export"
        Exit Sub
    End If

    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder() Then Exit Sub

    Set translationTable = ResolveTranslationsTable()

    'The language and the code of the sheet live in its hidden names.
    Set store = HiddenNames.Create(targetSheet)

    Set exporter = BuildExporter()
    filePath = exporter.ExportDisease(io.Folder(), targetSheet, translationTable, _
                                      targetSheet.Name, _
                                      store.ValueAsString(NAME_DISLANG), _
                                      store.ValueAsString(NAME_DISCODE))

    MsgBox "Done! The disease file is saved at:" & vbNewLine & filePath, vbInformation + vbOKOnly, "Export"
End Sub

'@sub-title Export every disease of this workbook into one flat migration file.
Public Sub ExportForMigration()
    Dim io As OSFiles
    Dim exporter As DiseaseExporter
    Dim diseaseNames As BetterArray
    Dim filePath As String

    Set diseaseNames = CollectDiseaseNames()
    If diseaseNames.Length = 0 Then
        MsgBox "This workbook carries no disease worksheet to export.", vbExclamation + vbOKOnly, "Export"
        Exit Sub
    End If

    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder() Then Exit Sub

    Set exporter = BuildExporter()
    filePath = exporter.ExportForMigration(io.Folder(), ThisWorkbook, diseaseNames)

    MsgBox "Done! The migration file is saved at:" & vbNewLine & filePath, vbInformation + vbOKOnly, "Export"
End Sub

'@section Setup import
'===============================================================================

'@sub-title Read a disease workbook exported for a setup back into this file.
'@details The workbook is the one ExportToSetup writes. Its disease
'worksheet is rebuilt when it is gone and merged when it is there, the
'Variables table takes the variables it lacks and the Choices sheet takes
'the lists it lacks. The lists of a file in another language than English
'land with their labels empty. A file that names no disease is asked for
'one; a name already in the workbook is refused, and four bad answers
'abort the import with an error.
Public Sub ImportFromSetup()
    Dim io As OSFiles
    Dim sourceBook As Workbook
    Dim service As MasterSetupImportService
    Dim diseaseName As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set io = OSFiles.Create()
    io.LoadFile IMPORT_FILTERS
    If Not io.HasValidFile() Then Exit Sub

    Set sourceBook = Application.Workbooks.Open(fileName:=io.File(), ReadOnly:=True)

    On Error GoTo Handler

    Set service = BuildImportService(ThisWorkbook)

    diseaseName = service.ReadDiseaseName(sourceBook)
    If LenB(diseaseName) = 0 Then diseaseName = AskDiseaseName(service)

    service.ImportSetupExport sourceBook, diseaseName:=diseaseName

    MsgBox "Done! The disease '" & service.DiseaseName & "' is imported." & vbNewLine & _
           service.AddedVariables.Length & " variable(s) added to the Variables sheet, " & _
           service.AddedChoices.Length & " list(s) added to the Choices sheet.", _
           vbInformation + vbOKOnly, "Import"

CloseSource:
    'Shielded close: the source workbook must not stay open on any path, and
    'a failure of the import itself still reaches the caller.
    On Error Resume Next
    sourceBook.Close saveChanges:=False
    On Error GoTo 0

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If
    Exit Sub

Handler:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CloseSource
End Sub

'@sub-title Ask the user for the disease name of a file that carries none.
'@details The answer has to be free: no worksheet of that name, and none of
'the prohibited list. The question is put MAX_NAME_ATTEMPTS times, then the
'import is aborted with an error.
Private Function AskDiseaseName(ByVal service As MasterSetupImportService) As String
    Dim attempt As Long
    Dim candidate As String
    Dim prompt As String

    prompt = "The file names no disease. Enter the name of the disease worksheet to create:"

    For attempt = 1 To MAX_NAME_ATTEMPTS
        candidate = MasterSetupHelpers.CleanMasterSheetName(InputBox(prompt, "Import"))
        If service.DiseaseNameIsFree(candidate) Then
            AskDiseaseName = candidate
            Exit Function
        End If
        prompt = "'" & candidate & "' is empty, reserved or already a worksheet of this file. " & _
                 "Enter another name (" & attempt & " of " & MAX_NAME_ATTEMPTS & " tries used):"
    Next attempt

    Err.Raise ProjectError.InvalidArgument, "MasterSetupExports.ImportFromSetup", _
              "No valid disease name after " & MAX_NAME_ATTEMPTS & " tries; the import is aborted."
End Function

'@sub-title Fold an open setup export into the target workbook.
'@details Public and fully parameterised so a suite can drive the import
'without a file picker. ImportFromSetup wraps it with the picker and the
'name question.
'@param sourceBook Workbook. The exported workbook, open.
'@param targetBook Workbook. The master setup receiving it.
'@param logger Optional DiseaseLogger.
'@param diseaseName Optional String. The name to use when the file names none.
'@return MasterSetupImportService carrying the disease name, the variables
'   and the lists added.
Public Function ImportSetupWorkbook(ByVal sourceBook As Workbook, _
                                    ByVal targetBook As Workbook, _
                                    Optional ByVal logger As DiseaseLogger = Nothing, _
                                    Optional ByVal diseaseName As String = vbNullString) As MasterSetupImportService
    Dim service As MasterSetupImportService

    Set service = BuildImportService(targetBook)
    service.ImportSetupExport sourceBook, logger, diseaseName

    Set ImportSetupWorkbook = service
End Function

'@section Migration import
'===============================================================================

'@sub-title Read a flat migration file and recreate or merge its diseases.
Public Sub ImportFlatFile()
    Dim io As OSFiles
    Dim sourceBook As Workbook
    Dim diseasesSheet As Worksheet
    Dim importedCount As Long

    Set io = OSFiles.Create()
    io.LoadFile IMPORT_FILTERS
    If Not io.HasValidFile() Then Exit Sub

    Set sourceBook = Application.Workbooks.Open(fileName:=io.File(), ReadOnly:=True)

    On Error GoTo CloseSource

    Set diseasesSheet = FindWorksheet(sourceBook, DISEASES_SHEET)
    If diseasesSheet Is Nothing Then
        MsgBox "The selected file carries no '" & DISEASES_SHEET & "' worksheet.", _
               vbExclamation + vbOKOnly, "Import"
        GoTo CloseSource
    End If

    importedCount = ImportMigrationDiseases(sourceBook, ThisWorkbook)

    MsgBox "Done! " & importedCount & " disease worksheet(s) imported.", _
           vbInformation + vbOKOnly, "Import"

CloseSource:
    'Shielded close: the source workbook must not stay open on any path.
    On Error Resume Next
    sourceBook.Close saveChanges:=False
    On Error GoTo 0
End Sub

'@sub-title Walk the Diseases sheet blocks of a flat file and import each one.
'@details Public and fully parameterised so a suite can drive the round trip
'without a file picker. ImportFlatFile wraps it with the pickers. The
'header row is walked for the word Disease, so the blocks can sit at any
'stride.
'@param sourceBook Workbook carrying the flat Diseases sheet.
'@param targetBook Workbook receiving the disease worksheets.
'@return Long number of disease blocks imported.
Public Function ImportMigrationDiseases(ByVal sourceBook As Workbook, ByVal targetBook As Workbook) As Long
    Dim diseasesSheet As Worksheet
    Dim service As MasterSetupImportService
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim diseaseName As String
    Dim languageTag As String

    Set diseasesSheet = FindWorksheet(sourceBook, DISEASES_SHEET)
    If diseasesSheet Is Nothing Then Exit Function

    Set service = BuildImportService(targetBook)

    lastColumn = diseasesSheet.UsedRange.Columns(diseasesSheet.UsedRange.Columns.Count).Column

    For columnIndex = 1 To lastColumn
        If MasterSetupHelpers.SafeValue(diseasesSheet.Cells(BLOCK_HEADER_ROW, columnIndex).Value) = "Disease" Then
            diseaseName = MasterSetupHelpers.CleanMasterSheetName( _
                          MasterSetupHelpers.SafeValue(diseasesSheet.Cells(BLOCK_META_ROW, columnIndex).Value))
            languageTag = MasterSetupHelpers.SafeValue(diseasesSheet.Cells(BLOCK_META_ROW, columnIndex + 1).Value)

            If LenB(diseaseName) > 0 Then
                ImportOneDisease service, diseasesSheet, columnIndex, diseaseName, languageTag, targetBook
                ImportMigrationDiseases = ImportMigrationDiseases + 1
            End If
        End If
    Next columnIndex
End Function

'@sub-title Stage one column block and land it through the service.
Private Sub ImportOneDisease(ByVal service As MasterSetupImportService, _
                             ByVal diseasesSheet As Worksheet, ByVal startColumn As Long, _
                             ByVal diseaseName As String, ByVal languageTag As String, _
                             ByVal targetBook As Workbook)

    Dim stagingTable As ListObject
    Dim manager As DiseaseWorksheetManager

    Set stagingTable = BuildStagingTable(diseasesSheet, startColumn, targetBook)
    If stagingTable Is Nothing Then Exit Sub

    'A migration block keeps the values it carries; the formulas are M7's.
    service.ImportDiseaseTable stagingTable, diseaseName, languageTag

    Set manager = DiseaseWorksheetManager.Create()
    manager.RemoveWorksheet targetBook, IMPORT_STAGING_SHEET
End Sub

'@sub-title Copy one block into a staging sheet and answer its ListObject.
Private Function BuildStagingTable(ByVal diseasesSheet As Worksheet, ByVal startColumn As Long, _
                                   ByVal targetBook As Workbook) As ListObject
    Dim stagingSheet As Worksheet
    Dim manager As DiseaseWorksheetManager
    Dim rowCount As Long
    Dim blockRange As Range
    Dim tableRange As Range

    rowCount = CountBlockRows(diseasesSheet, startColumn)
    If rowCount = 0 Then Exit Function

    Set manager = DiseaseWorksheetManager.Create()
    manager.RemoveWorksheet targetBook, IMPORT_STAGING_SHEET

    Set stagingSheet = targetBook.Worksheets.Add(After:=targetBook.Worksheets(targetBook.Worksheets.Count))
    stagingSheet.Name = IMPORT_STAGING_SHEET
    stagingSheet.Visible = xlSheetHidden

    'Header plus data, copied as values in one assignment.
    Set blockRange = diseasesSheet.Cells(BLOCK_TABLE_ROW, startColumn).Resize(rowCount + 1, DISEASE_TABLE_WIDTH)
    Set tableRange = stagingSheet.Range("A1").Resize(rowCount + 1, DISEASE_TABLE_WIDTH)
    tableRange.Value = blockRange.Value

    Set BuildStagingTable = stagingSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                         XlListObjectHasHeaders:=xlYes)
End Function

'@sub-title Count the data rows of a block, reading the name column.
Private Function CountBlockRows(ByVal diseasesSheet As Worksheet, ByVal startColumn As Long) As Long
    Dim rowIndex As Long
    Dim nameColumn As Long

    'The variable name is the second column of a disease table.
    nameColumn = startColumn + 1
    rowIndex = BLOCK_TABLE_ROW + 1

    Do While LenB(MasterSetupHelpers.SafeValue(diseasesSheet.Cells(rowIndex, nameColumn).Value)) > 0
        CountBlockRows = CountBlockRows + 1
        rowIndex = rowIndex + 1
    Loop
End Function

'@section Shared helpers
'===============================================================================

'@sub-title Build a ready exporter over a fresh workbook manager and guard.
Private Function BuildExporter() As DiseaseExporter
    Set BuildExporter = DiseaseExporter.Create(DiseaseExportWorkbook.Create(), _
                                               ApplicationState.Create(Application))
End Function

'@sub-title Build the import service over the master managers of a workbook.
'@details The Choices sheet is optional: a workbook without one still takes
'its diseases, and the service logs the lists it leaves out.
Private Function BuildImportService(ByVal targetBook As Workbook) As MasterSetupImportService
    Dim dropdowns As DropdownLists
    Dim variables As MasterSetupVariables
    Dim choices As LLChoices
    Dim choicesSheet As Worksheet
    Dim builder As DiseaseSheet

    Set dropdowns = MasterSetupHelpers.ResolveMasterDropdowns( _
                    MasterSetupHelpers.ResolveMasterDropdownsSheet(targetBook))
    Set variables = MasterSetupHelpers.ResolveMasterSetupVariables( _
                    MasterSetupHelpers.ResolveMasterVariablesSheet(targetBook))

    Set choicesSheet = MasterSetupHelpers.ResolveMasterChoicesSheet(targetBook)
    If Not choicesSheet Is Nothing Then
        Set choices = MasterSetupHelpers.ResolveMasterChoices(choicesSheet)
    End If

    Set builder = DiseaseSheet.Create(targetBook, dropdowns, variables)

    Set BuildImportService = MasterSetupImportService.Create(targetBook, builder, dropdowns, variables, choices)
End Function

'@sub-title Collect the names of every disease worksheet of this workbook.
Private Function CollectDiseaseNames() As BetterArray
    Dim sh As Worksheet

    Set CollectDiseaseNames = New BetterArray
    CollectDiseaseNames.LowerBound = 1

    For Each sh In ThisWorkbook.Worksheets
        If MasterSetupHelpers.IsMasterDiseaseSheet(sh) Then
            CollectDiseaseNames.Push sh.Name
        End If
    Next sh
End Function

Private Function ResolveTranslationsTable() As ListObject
    Dim translationsSheet As Worksheet

    Set translationsSheet = MasterSetupHelpers.ResolveMasterTranslationsSheet()
    If translationsSheet Is Nothing Then Exit Function
    If translationsSheet.ListObjects.Count = 0 Then Exit Function

    Set ResolveTranslationsTable = translationsSheet.ListObjects(1)
End Function

Private Function FindWorksheet(ByVal targetBook As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet

    For Each sheet In targetBook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = sheet
            Exit Function
        End If
    Next sheet
End Function
