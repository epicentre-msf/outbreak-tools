Attribute VB_Name = "MasterSetupExports"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("Disease exports, the setup import and the migration round trip for the master setup workbook.")
'@depends MasterSetupEventsManager, EventMasterSetup, MasterSetupLog, MasterSetupHelpers, MasterSetupImportService, MasterSetupMigration, MasterSetupVariables, DiseaseExporter, DiseaseExportWorkbook, DiseaseComparisonReport, DiseaseSheet, DiseaseLogger, LLChoices, DropdownLists, ApplicationState, OSFiles, HiddenNames, BetterArray
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, ExcelMemberMayReturnNothing, UseMeaningfulName

'Five doors to the outside world, all built on the msetup classes:
'- ExportToSetup writes one disease as a dictionary workbook a setup imports.
'- ImportFromSetup reads such a workbook back: the disease worksheet is
'  rebuilt or merged, and the Variables and Choices sheets take what they
'  lack.
'- ExportForMigration writes the whole master setup into one migration
'  file: every disease, the Variables table, the Choices block and the
'  translations.
'- ImportFlatFile reads such a migration file into this workbook, which
'  has to be empty of diseases, variables and choices.
'- CompareDiseaseSheets asks for two diseases and writes their differences
'  on the __compRep worksheet.
'
'The landing of a block on a disease worksheet is MasterSetupImportService's
'work, and the migration format at both ends is MasterSetupMigration's; this
'module walks the pickers and builds the classes over the master managers.
'
'Every door writes its outcome in the user log EventMasterSetup holds on
'__log: one success line naming the file or the diseases, one warning line
'on a refusal or a cancelled picker. The failure line is written by the
'ribbon callback that opened the door, because the raise lands there. The
'DiseaseLogger the exporter and the migration fill is read back here and
'folded into the success line as one count.

Private Const NAME_DISLANG As String = "__Var_DISLANG"
Private Const NAME_DISCODE As String = "__Var_DISCODE"
Private Const IMPORT_FILTERS As String = "*.xlsx;*.xlsb;*.xlsm"
'How many times the door asks for a disease name before it gives up.
Private Const MAX_NAME_ATTEMPTS As Long = 4
Private Const COMPARE_TITLE As String = "Compare"

'@section Exports
'===============================================================================

'@sub-title Export the active disease worksheet as a setup dictionary workbook.
Public Sub ExportToSetup()
    Dim targetSheet As Worksheet
    Dim io As OSFiles
    Dim exporter As DiseaseExporter
    Dim translationTable As ListObject
    Dim store As HiddenNames
    Dim logger As DiseaseLogger
    Dim filePath As String

    Set targetSheet = ActiveSheet
    If targetSheet Is Nothing Then Exit Sub

    If Not MasterSetupHelpers.IsMasterDiseaseSheet(targetSheet) Then
        MsgBox "Select a disease worksheet before exporting it.", vbExclamation + vbOKOnly, "Export"
        LogWarningLine "export-setup", "'" & targetSheet.Name & "' is no disease worksheet", "ExportToSetup"
        Exit Sub
    End If

    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder() Then
        LogWarningLine "export-setup", "folder picker cancelled", "ExportToSetup"
        Exit Sub
    End If

    Set translationTable = ResolveTranslationsTable(ThisWorkbook)

    'The language and the code of the sheet live in its hidden names.
    Set store = HiddenNames.Create(targetSheet)

    Set logger = DiseaseLogger.Create()
    Set exporter = BuildExporter()
    filePath = exporter.ExportDisease(io.Folder(), targetSheet, translationTable, _
                                      targetSheet.Name, _
                                      store.ValueAsString(NAME_DISLANG), _
                                      store.ValueAsString(NAME_DISCODE), _
                                      logger)

    LogSuccessLine "export-setup", targetSheet.Name & " to " & filePath & LoggerSummary(logger), "ExportToSetup"
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
        LogWarningLine "export-migration", "no disease worksheet to export", "ExportForMigration"
        Exit Sub
    End If

    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder() Then
        LogWarningLine "export-migration", "folder picker cancelled", "ExportForMigration"
        Exit Sub
    End If

    Set migration = BuildMigration(ThisWorkbook)
    filePath = migration.ExportMigration(io.Folder())

    LogSuccessLine "export-migration", diseaseNames.Length & " disease(s) to " & filePath, "ExportForMigration"
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
    Dim logger As DiseaseLogger
    Dim diseaseName As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set io = OSFiles.Create()
    io.LoadFile IMPORT_FILTERS
    If Not io.HasValidFile() Then
        LogWarningLine "import-setup", "file picker cancelled", "ImportFromSetup"
        Exit Sub
    End If

    Set sourceBook = Application.Workbooks.Open(fileName:=io.File(), ReadOnly:=True)

    On Error GoTo Handler

    Set service = BuildImportService(ThisWorkbook)

    diseaseName = service.ReadDiseaseName(sourceBook)
    If LenB(diseaseName) = 0 Then diseaseName = AskDiseaseName(service)

    Set logger = DiseaseLogger.Create()
    service.ImportSetupExport sourceBook, logger, diseaseName

    LogSuccessLine "import-setup", service.DiseaseName & " from " & io.File() & ", " & _
                   service.AddedVariables.Length & " variable(s) and " & _
                   service.AddedChoices.Length & " list(s) added" & LoggerSummary(logger), _
                   "ImportFromSetup"
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
    Dim logger As DiseaseLogger
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set io = OSFiles.Create()
    io.LoadFile IMPORT_FILTERS
    If Not io.HasValidFile() Then
        LogWarningLine "import-migration", "file picker cancelled", "ImportFlatFile"
        Exit Sub
    End If

    Set sourceBook = Application.Workbooks.Open(fileName:=io.File(), ReadOnly:=True)

    On Error GoTo Handler

    Set logger = DiseaseLogger.Create()
    Set migration = BuildMigration(ThisWorkbook)
    migration.Import sourceBook, logger

    LogSuccessLine "import-migration", migration.ImportedDiseaseCount & " disease(s) from " & io.File() & _
                   ", " & migration.AddedLanguages.Length & " language(s) added" & LoggerSummary(logger), _
                   "ImportFlatFile"
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

'@section Disease comparison
'===============================================================================

'@sub-title Compare two disease worksheets and write the report on __compRep.
'@details The user picks disease 1 from a numbered list of every disease
'worksheet, then disease 2 from the same list with disease 1 taken out, so
'one disease is never compared with itself. A workbook with fewer than two
'disease worksheets is refused with one message. The comparison runs in a
'busy scope and ends on the report sheet.
Public Sub CompareDiseaseSheets()
    Dim scope As ApplicationState
    Dim diseaseNames As BetterArray
    Dim remainingNames As BetterArray
    Dim firstName As String
    Dim secondName As String
    Dim report As DiseaseComparisonReport
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set diseaseNames = CollectDiseaseNames()
    If diseaseNames.Length < 2 Then
        MsgBox "This workbook needs two disease worksheets to compare.", vbExclamation + vbOKOnly, COMPARE_TITLE
        LogWarningLine "compare-diseases", "fewer than two disease worksheets", "CompareDiseaseSheets"
        Exit Sub
    End If

    firstName = PromptDiseaseName(diseaseNames, "Select the first disease to compare:")
    If LenB(firstName) = 0 Then
        LogWarningLine "compare-diseases", "no first disease picked", "CompareDiseaseSheets"
        Exit Sub
    End If

    Set remainingNames = NamesWithout(diseaseNames, firstName)
    secondName = PromptDiseaseName(remainingNames, "Select the disease to compare with '" & firstName & "':")
    If LenB(secondName) = 0 Then
        LogWarningLine "compare-diseases", "no second disease picked", "CompareDiseaseSheets"
        Exit Sub
    End If

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set report = DiseaseComparisonReport.Create(ThisWorkbook)
    report.PrintComparison ResolveDiseaseTable(ThisWorkbook.Worksheets(firstName)), _
                           ResolveDiseaseTable(ThisWorkbook.Worksheets(secondName)), _
                           firstName, secondName
    LogSuccessLine "compare-diseases", firstName & " with " & secondName, "CompareDiseaseSheets"

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
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

'@sub-title Ask the user to pick one disease out of a numbered list.
'@details The list is written into the prompt, one name per line, and the
'answer is the number of the line. A cancel answers an empty string. A
'number that is not whole or falls outside the list is refused with one
'message and answers an empty string too.
'@param diseaseNames BetterArray. The names to choose from, based at 1.
'@param prompt String. The question shown above the list.
'@return String. The chosen name, or an empty string.
Private Function PromptDiseaseName(ByVal diseaseNames As BetterArray, ByVal prompt As String) As String
    Dim idx As Long
    Dim promptText As String
    Dim response As Variant
    Dim numericResponse As Double
    Dim selection As Long

    If diseaseNames Is Nothing Then Exit Function
    If diseaseNames.Length = 0 Then Exit Function

    promptText = prompt & vbLf
    For idx = diseaseNames.LowerBound To diseaseNames.UpperBound
        promptText = promptText & CStr(idx - diseaseNames.LowerBound + 1) & ". " & _
                     CStr(diseaseNames.Item(idx)) & vbLf
    Next idx

    response = Application.InputBox(promptText, COMPARE_TITLE, Type:=1)
    'A cancel answers False.
    If VarType(response) = vbBoolean Then Exit Function

    numericResponse = CDbl(response)
    If numericResponse <> Int(numericResponse) Then GoTo InvalidSelection
    If numericResponse < 1 Or numericResponse > diseaseNames.Length Then GoTo InvalidSelection

    selection = CLng(numericResponse)
    PromptDiseaseName = Trim$(CStr(diseaseNames.Item(diseaseNames.LowerBound + selection - 1)))
    Exit Function

InvalidSelection:
    MsgBox "Invalid selection.", vbExclamation + vbOKOnly, COMPARE_TITLE
End Function

'@sub-title A copy of a name list with one name taken out.
Private Function NamesWithout(ByVal diseaseNames As BetterArray, ByVal excludedName As String) As BetterArray
    Dim idx As Long
    Dim candidate As String

    Set NamesWithout = New BetterArray
    NamesWithout.LowerBound = 1

    For idx = diseaseNames.LowerBound To diseaseNames.UpperBound
        candidate = CStr(diseaseNames.Item(idx))
        If StrComp(candidate, excludedName, vbTextCompare) <> 0 Then
            NamesWithout.Push candidate
        End If
    Next idx
End Function

'@sub-title The table of a disease worksheet.
'@details A disease worksheet carries one ListObject; a sheet without one
'is refused with ElementNotFound.
Private Function ResolveDiseaseTable(ByVal targetSheet As Worksheet) As ListObject
    If targetSheet.ListObjects.Count = 0 Then
        Err.Raise ProjectError.ElementNotFound, "MasterSetupExports.CompareDiseaseSheets", _
                  "The disease worksheet '" & targetSheet.Name & "' carries no table."
    End If
    Set ResolveDiseaseTable = targetSheet.ListObjects(1)
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

'@section User log
'===============================================================================

'@sub-title The user log the event service holds.
'@details A workbook whose log cannot be built answers Nothing and every
'log line of this module stays quiet.
Private Function UserLogOf() As MasterSetupLog
    Dim service As EventMasterSetup

    Set service = MasterSetupEventsManager.MasterSetupService()
    If service Is Nothing Then Exit Function

    Set UserLogOf = service.UserLog()
End Function

'@sub-title Write the success line of a finished door.
'@details The write is guarded so a log fault never takes down the walk
'it records.
Private Sub LogSuccessLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString, _
                           Optional ByVal source As String = vbNullString)
    Dim logStore As MasterSetupLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogSuccess action, detail, source
    On Error GoTo 0
End Sub

'@sub-title Write the warning line of a refused or cancelled door.
Private Sub LogWarningLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString, _
                           Optional ByVal source As String = vbNullString)
    Dim logStore As MasterSetupLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogWarning action, detail, source
    On Error GoTo 0
End Sub

'@sub-title The warnings and errors of a DiseaseLogger, as one short count.
'@details Every entry of the logger is a BetterArray whose third item is
'the severity. Answers an empty string when the logger holds only info
'lines, so the success line stays short on a clean run.
'@param logger DiseaseLogger. The logger the class filled.
'@return String. ", N warning(s), M error(s)" or an empty string.
Public Function LoggerSummary(ByVal logger As DiseaseLogger) As String
    Dim entries As BetterArray
    Dim entry As BetterArray
    Dim index As Long
    Dim severity As Long
    Dim warningCount As Long
    Dim errorCount As Long

    If logger Is Nothing Then Exit Function
    If Not logger.HasEntries Then Exit Function

    Set entries = logger.Entries
    For index = entries.LowerBound To entries.UpperBound
        Set entry = entries.Item(index)
        severity = CLng(entry.Item(entry.LowerBound + 2))
        If severity = DiseaseLogWarning Then warningCount = warningCount + 1
        If severity = DiseaseLogError Then errorCount = errorCount + 1
    Next index

    If warningCount = 0 And errorCount = 0 Then Exit Function
    LoggerSummary = ", " & warningCount & " warning(s), " & errorCount & " error(s)"
End Function

'@sub-title The translations table of a master setup workbook, or Nothing.
Private Function ResolveTranslationsTable(ByVal targetBook As Workbook) As ListObject
    Dim translationsSheet As Worksheet

    Set translationsSheet = MasterSetupHelpers.ResolveMasterTranslationsSheet(targetBook)
    If translationsSheet Is Nothing Then Exit Function
    If translationsSheet.ListObjects.Count = 0 Then Exit Function

    Set ResolveTranslationsTable = translationsSheet.ListObjects(1)
End Function
