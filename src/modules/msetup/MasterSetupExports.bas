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
'- ExportForMigration writes the whole master setup into one migration
'  file: every disease, the Variables table, the Choices block and the
'  translations.
'- ImportFlatFile reads such a migration file into this workbook, which
'  has to be empty of diseases, variables and choices.
'
'The landing of a block on a disease worksheet is MasterSetupImportService's
'work, and the migration format at both ends is MasterSetupMigration's; this
'module walks the pickers and builds the classes over the master managers.

Private Const NAME_DISLANG As String = "__Var_DISLANG"
Private Const NAME_DISCODE As String = "__Var_DISCODE"
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

    Set translationTable = ResolveTranslationsTable(ThisWorkbook)

    'The language and the code of the sheet live in its hidden names.
    Set store = HiddenNames.Create(targetSheet)

    Set exporter = BuildExporter()
    filePath = exporter.ExportDisease(io.Folder(), targetSheet, translationTable, _
                                      targetSheet.Name, _
                                      store.ValueAsString(NAME_DISLANG), _
                                      store.ValueAsString(NAME_DISCODE))

    MsgBox "Done! The disease file is saved at:" & vbNewLine & filePath, vbInformation + vbOKOnly, "Export"
End Sub

'@sub-title Export the whole master setup into one migration file.
Public Sub ExportForMigration()
    Dim io As OSFiles
    Dim migration As MasterSetupMigration
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

    Set migration = BuildMigration(ThisWorkbook)
    filePath = migration.ExportMigration(io.Folder())

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

'@sub-title Read a migration file into this workbook, which has to be empty.
'@details The file is the one ExportForMigration writes. The class refuses
'a workbook that already carries a disease, a Variables line or a Choices
'line, and writes nothing in that case. The "Done" message names the
'diseases landed and the languages added.
Public Sub ImportFlatFile()
    Dim io As OSFiles
    Dim sourceBook As Workbook
    Dim migration As MasterSetupMigration
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set io = OSFiles.Create()
    io.LoadFile IMPORT_FILTERS
    If Not io.HasValidFile() Then Exit Sub

    Set sourceBook = Application.Workbooks.Open(fileName:=io.File(), ReadOnly:=True)

    On Error GoTo Handler

    Set migration = BuildMigration(ThisWorkbook)
    migration.Import sourceBook

    MsgBox "Done! " & migration.ImportedDiseaseCount & " disease worksheet(s) imported, " & _
           migration.AddedLanguages.Length & " language(s) added to the translations table.", _
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

'@sub-title Build the migration over the master managers of a workbook.
'@details The migration carries the whole file, so every master sheet has
'to be there: a workbook missing its Choices or Translations sheet is
'refused with ElementNotFound.
Private Function BuildMigration(ByVal targetBook As Workbook) As MasterSetupMigration
    Dim dropdowns As DropdownLists
    Dim variables As MasterSetupVariables
    Dim choices As LLChoices
    Dim choicesSheet As Worksheet
    Dim translationTable As ListObject
    Dim builder As DiseaseSheet
    Dim service As MasterSetupImportService

    Set dropdowns = MasterSetupHelpers.ResolveMasterDropdowns( _
                    MasterSetupHelpers.ResolveMasterDropdownsSheet(targetBook))
    Set variables = MasterSetupHelpers.ResolveMasterSetupVariables( _
                    MasterSetupHelpers.ResolveMasterVariablesSheet(targetBook))

    Set choicesSheet = MasterSetupHelpers.ResolveMasterChoicesSheet(targetBook)
    If choicesSheet Is Nothing Then
        Err.Raise ProjectError.ElementNotFound, "MasterSetupExports.BuildMigration", _
                  "The workbook carries no Choices worksheet; the migration needs one."
    End If
    Set choices = MasterSetupHelpers.ResolveMasterChoices(choicesSheet)

    Set translationTable = ResolveTranslationsTable(targetBook)
    If translationTable Is Nothing Then
        Err.Raise ProjectError.ElementNotFound, "MasterSetupExports.BuildMigration", _
                  "The workbook carries no translations table; the migration needs one."
    End If

    Set builder = DiseaseSheet.Create(targetBook, dropdowns, variables)
    Set service = MasterSetupImportService.Create(targetBook, builder, dropdowns, variables, choices)

    Set BuildMigration = MasterSetupMigration.Create(targetBook, dropdowns, variables, choices, _
                                                     translationTable, builder, service)
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

'@sub-title The translations table of a master setup workbook, or Nothing.
Private Function ResolveTranslationsTable(ByVal targetBook As Workbook) As ListObject
    Dim translationsSheet As Worksheet

    Set translationsSheet = MasterSetupHelpers.ResolveMasterTranslationsSheet(targetBook)
    If translationsSheet Is Nothing Then Exit Function
    If translationsSheet.ListObjects.Count = 0 Then Exit Function

    Set ResolveTranslationsTable = translationsSheet.ListObjects(1)
End Function
