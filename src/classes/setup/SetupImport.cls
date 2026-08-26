VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetupImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Improved setup import service orchestrating workbook synchronisation"

'@PredeclaredId
'@Folder("Setup")
'@ModuleDescription("Improved setup import service orchestrating workbook synchronisation")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends ProjectError, ApplicationState, BetterArray, CustomTable, DataSheet, HiddenNames, LLChoices, LLdictionary, LLExport, LLFormat, OSFiles, Passwords, Analysis, SetupTranslationsTable

Option Explicit

'@class SetupImport
'@description
'Moves the content of a setup workbook in and out. Check validates the setup
'sheets, Import reads the Dictionary, Choices, Exports, Analysis and
'Translations sheets from a file, ImportFromWorkbook reads them from an open
'workbook, Export writes them out, and Clean empties them. Path and
'SetExportFolder say where files are read from and written to, LastExportFile
'names the file the last export produced, ProgressObject takes the label the
'form shows progress on, and DisplayPrompts decides whether the class puts a
'message in front of the user.
'
'HOW A SHEET IS MOVED
'-------------------------------------------------------------------------------
'A sheet travels either as a whole table through CustomTable, or through the
'domain manager that owns it: LLdictionary, LLChoices, LLExport, Analysis and
'SetupTranslationsTable. Excel state is held by an ApplicationState for the
'length of the run, and the worksheets are unprotected and protected again
'through Passwords.
'
'THREE RULES HOLD ACROSS THE CLASS
'-------------------------------------------------------------------------------
'1. A cleanup block reached by falling off the end of the work runs under its
'   own handler, with the failure handler disarmed first.
'2. Every failure handler copies Err.Number, Err.Source and Err.Description
'   into locals as its first three statements. Cleanup calls run On Error
'   Resume Next and clear Err, so an error read after them belongs to the
'   cleanup.
'3. Sheet names are compared without regard to case. Excel matches worksheet
'   names that way itself, so "choices" reaches the same code as "Choices".


'@section State and constants
'===============================================================================
Private Const CLASS_NAME As String = "SetupImport"
Private Const EXPORT_SHEET_NAME As String = "Exports"
Private Const DICTIONARY_SHEET_NAME As String = "Dictionary"
Private Const ANALYSIS_SHEET_NAME As String = "Analysis"
Private Const CHOICES_SHEET_NAME As String = "Choices"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const FORMATTER_SHEET_NAME As String = "__formatter"
'Default host anchor positions for setup tables using domain objects
Private Const DICTIONARY_HOST_START_ROW As Long = 5
Private Const DICTIONARY_HOST_START_COLUMN As Long = 1
Private Const CHOICES_HOST_START_ROW As Long = 4
Private Const CHOICES_HOST_START_COLUMN As Long = 1
Private Const EXPORT_HOST_START_ROW As Long = 4
Private Const EXPORT_HOST_START_COLUMN As Long = 1
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"
'Incoming workbooks default to top-left anchored data
Private Const SOURCE_START_ROW As Long = 1
Private Const SOURCE_START_COLUMN As Long = 1
'A translations sheet keeps one spare column on the left of its table for the
'row tags, so the labels start on the second column.
Private Const TRANSLATION_SOURCE_START_ROW As Long = 1
Private Const TRANSLATION_SOURCE_START_COLUMN As Long = 2
'Separator wrapping every entry of a sheet-name key string. It sits on both
'sides of each name so a short name cannot match inside a longer one.
Private Const LIST_SEPARATOR As String = vbNullChar
'A sheet list is a handful of names, so 64 covers every real call with room to
'spare and stops a broken list building a huge key string.
Private Const MAX_SHEET_LIST_LENGTH As Long = 64

Private Type TSetupImportState
    importPath As String                    'Path to the source setup workbook
    progressDisplay As Object               'UI element displaying progress messages
    importWorkbook As Workbook              'Lazy loaded reference to the imported workbook
    appScope As ApplicationState            'Application state guard ensuring Excel settings are restored
    busyDepth As Long                       'Number of public calls currently holding the busy state
    hasCaption As Boolean
    exportPath As String
    exportWorkbook As Workbook
    exportBaseSheets As Collection          'Sheets the new export workbook was born with
    displayPrompt As Boolean
    lastExportFile As String
End Type

Private this As TSetupImportState

'@label:initialize
'@sub-title Set default state on construction.
Private Sub Class_Initialize()
    this.displayPrompt = True
    this.lastExportFile = vbNullString
End Sub


'@section Factory helpers
'===============================================================================
'@label:create
'@sub-title Create a ready-to-use service instance
'@details
'Creates a new SetupImport instance and assigns the path and the display
'through the property setters, which validate both.
'
'The validation is left to the setters on purpose. Create runs on the
'predeclared default instance, so a validator called here writes its answer
'(this.hasCaption) onto that default instance and the new object keeps a stale
'value until a setter fills it in.
'@param importPath String. Path to the source setup workbook.
'@param progressDisplay Optional Object. UI element receiving textual updates (must expose Caption or Value).
'@return SetupImport. Configured service instance.
'@export
Public Function Create(ByVal importPath As String, Optional ByVal progressDisplay As Object) As SetupImport
    Dim instance As SetupImport

    Set instance = New SetupImport
    instance.Path = importPath
    Set instance.ProgressObject = progressDisplay

    Set Create = instance
End Function


'@section Public configuration
'===============================================================================
'@label:path
'@prop-title Retrieve configured import path.
'@return String. The configured import path.
'@export
Public Property Get Path() As String
    Path = this.importPath
End Property

'@label:path-set
'@prop-title Store the path to the setup workbook.
'@param value String. The path to assign.
Public Property Let Path(ByVal value As String)
    ValidateImportPath value
    this.importPath = Trim$(value)
End Property

'@label:progress-object
'@prop-title Retrieve the progress display object.
'@return Object. The progress display object.
'@export
Public Property Get ProgressObject() As Object
    Set ProgressObject = this.progressDisplay
End Property

'@label:progress-object-set
'@prop-title Store the progress display object reference.
'@param value Object. The progress display to store.
'@throws ProjectError.InvalidArgument When the object exposes neither Caption nor Value.
Public Property Set ProgressObject(ByVal value As Object)
    ValidateProgressDisplay value
    Set this.progressDisplay = value
End Property

'@label:display-prompts-set
'@prop-title Toggle UI prompts displayed during operations.
'@details
'When False, suppresses folder-picker dialogs and other interactive prompts.
'Primarily used by automated tests.
'@param state Boolean. True to enable prompts.
Public Property Let DisplayPrompts(ByVal state As Boolean)
    this.displayPrompt = state
End Property

'@label:display-prompts
'@prop-title Retrieve the display prompts toggle.
'@return Boolean. True when prompts are enabled.
'@export
Public Property Get DisplayPrompts() As Boolean
    DisplayPrompts = this.displayPrompt
End Property

'@label:set-export-folder
'@sub-title Provide an explicit folder for exports, bypassing the folder picker.
'@details
'The folder is consumed by one export. Export clears it when it closes the
'export workbook, so a second export on the same instance needs a second call
'to this method.
'@param folderPath String. The folder path to use for exports.
'@export
Public Sub SetExportFolder(ByVal folderPath As String)
    this.exportPath = Trim$(folderPath)
    If LenB(this.exportPath) = 0 Then
        this.exportPath = vbNullString
    End If
End Sub

'@label:last-export-file
'@prop-title Path of the last export produced by the service.
'@details
'Returns the full file path of the last export workbook saved by the
'service. Returns an empty string when no export has occurred, and when the
'caller supplied its own workbook to Export.
'@return String. The last export file path, or empty.
'@export
Public Property Get LastExportFile() As String
    LastExportFile = this.lastExportFile
End Property


'@section Application state coordination
'===============================================================================
'@label:application-scope
'@sub-title Guard Excel with the reusable ApplicationState scope.
Private Function ApplicationScope() As ApplicationState
    If this.appScope Is Nothing Then
        Set this.appScope = ApplicationState.Create(Application)
    End If

    Set ApplicationScope = this.appScope
End Function

'@label:enter-busy-state
'@sub-title Apply the busy state when heavy work starts.
'@details
'One scope covers a whole operation, including any inner call that enters the
'busy state again. The depth counter is what makes that work: only the
'outermost call restores, so an inner return can no longer put Calculation back
'to automatic in the middle of the work.
Private Sub EnterBusyState()
    Dim scope As ApplicationState

    this.busyDepth = this.busyDepth + 1
    If this.busyDepth > 1 Then Exit Sub

    Set scope = ApplicationScope()
    If Not scope.IsBusy Then
        scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, blockSecurity:=True
    End If
End Sub

'@label:leave-busy-state
'@sub-title Restore Excel configuration when the outermost call completes.
Private Sub LeaveBusyState(Optional ByVal silent As Boolean = False)
    If this.busyDepth > 0 Then this.busyDepth = this.busyDepth - 1
    If this.busyDepth > 0 Then Exit Sub

    RestoreApplicationState silent
End Sub

'@label:restore-application-state
'@sub-title Restore Excel and release the scope.
'@details
'The scope is released once it has restored. ApplicationState captures its
'snapshot once per object, so a released scope is what makes the next operation
'record the settings as they are then.
Private Sub RestoreApplicationState(ByVal silent As Boolean)
    If this.appScope Is Nothing Then Exit Sub

    'The snapshot decides, not the busy flag. A failed ApplyBusyState clears
    'IsBusy after it has already written screen updating and alerts, so a scope
    'that reads idle can still be holding the screen off. Restore whenever
    'there is something to put back; it is safe to call on an idle scope.
    If Not this.appScope.HasSnapshot Then
        Set this.appScope = Nothing
        Exit Sub
    End If

    On Error GoTo RestoreFailed
        this.appScope.Restore silent:=silent
    On Error GoTo 0
    Set this.appScope = Nothing
    Exit Sub

RestoreFailed:
    Set this.appScope = Nothing
    If silent Then
        Err.Clear
    Else
        Err.Raise Err.Number, CLASS_NAME, "Failed to restore application state: " & Err.Description
    End If
End Sub


'@section Core workflow
'===============================================================================
'@label:check
'@sub-title Validate the requested import operation before execution.
'@details
'Ensures at least one import option is selected, verifies the import
'path is configured and the file exists, then attempts to open and
'immediately close the source workbook to confirm it is readable.
'@param importDictionary Boolean. Flag requesting dictionary import.
'@param importChoices Boolean. Flag requesting choices import.
'@param importExports Boolean. Flag requesting exports import.
'@param importAnalysis Boolean. Flag requesting analysis import.
'@param importTranslations Boolean. Flag requesting translations import.
'@param cleanSetup Optional Boolean. When True, performs a clean-only run. Defaults to False.
'@throws ProjectError.InvalidArgument When no import option is selected.
'@throws ProjectError.ElementNotFound When no file sits at the configured path.
'@throws ProjectError.SomethingWentWrong When the source workbook cannot be opened.
'@export
Public Sub Check(ByVal importDictionary As Boolean, _
                 ByVal importChoices As Boolean, _
                 ByVal importExports As Boolean, _
                 ByVal importAnalysis As Boolean, _
                 ByVal importTranslations As Boolean, _
                 Optional ByVal cleanSetup As Boolean = False)

    Dim sourcePath As String
    Dim opened As Workbook
    Dim failureDetails As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    EnsureImportSelection importDictionary, importChoices, importExports, importAnalysis, importTranslations, cleanSetup

    If cleanSetup Then Exit Sub

    sourcePath = Path
    EnsureImportPathConfigured sourcePath
    EnsureImportFileExists sourcePath

    EnterBusyState
    On Error GoTo CheckFailed

        WriteInfo "Reading the setup file..."

        On Error Resume Next
            Set opened = Workbooks.Open(sourcePath)
            failureDetails = Err.Description
            Err.Clear
        On Error GoTo CheckFailed

        If opened Is Nothing Then
            ReportAndThrow ProjectError.SomethingWentWrong, _
                            BuildOpenFailureMessage(sourcePath, failureDetails)
        End If

    'Cleanup runs with the failure handler disarmed. CloseWorkbook swallows its
    'own errors, so only LeaveBusyState can raise here, and that raise belongs
    'to the caller.
    On Error GoTo 0
    CloseWorkbook opened
    LeaveBusyState
    Exit Sub

CheckCleanup:
    'Reached through Resume, which leaves the handler. The two On Error lines
    'here are load-bearing: Resume leaves the handler and leaves CheckFailed
    'ARMED, so a raise below would jump back into it and Resume here again, for
    'ever. Every cleanup label in this class disarms before it does its work.
    On Error Resume Next
        CloseWorkbook opened
    On Error GoTo 0
    LeaveBusyState True
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CheckFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CheckCleanup
End Sub

'@label:clean
'@sub-title Clean target worksheets ahead of a fresh import.
'@details
'Iterates over the requested sheets, unprotects each one, clears all
'ListObject data and worksheet comments, then re-protects. A sheet the host
'workbook does not carry is skipped, which is what Import does with the same
'list.
'@param pass Passwords. Password handler for worksheet protection.
'@param sheetsList BetterArray. Worksheets participating in the clean-up.
'@throws ProjectError.ObjectNotInitialized When no password handler is supplied.
'@throws ProjectError.InvalidArgument When the sheets list is missing or empty.
'@export
Public Sub Clean(ByVal pass As Passwords, ByVal sheetsList As BetterArray)
    Dim hostWorkbook As Workbook
    Dim idx As Long
    Dim sheetName As String
    Dim targetSheet As Worksheet
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set hostWorkbook = ThisWorkbook
    ValidatePasswords pass
    EnsureSheetsList sheetsList

    EnterBusyState
    On Error GoTo CleanFailed

        For idx = sheetsList.LowerBound To sheetsList.UpperBound
            sheetName = CStr(sheetsList.Item(idx))
            Set targetSheet = TryWorksheet(hostWorkbook, sheetName)

            If Not targetSheet Is Nothing Then
                WriteInfo "Clearing " & sheetName & "..."
                pass.UnProtect sheetName
                sheetUnlocked = True

                CleanWorksheetTables targetSheet
                ClearWorksheetComments targetSheet

                ProtectWorksheet pass, sheetName
                sheetUnlocked = False
            End If
        Next idx

    On Error GoTo 0
    LeaveBusyState
    Exit Sub

CleanCleanup:
    'Reached through Resume, which leaves the handler. A sheet that refused the
    'clean is also a sheet that can refuse the re-protect, and a second error
    'inside an active handler is not trappable.
    On Error Resume Next
        If sheetUnlocked Then
            ProtectWorksheet pass, sheetName
            sheetUnlocked = False
        End If
    On Error GoTo 0
    LeaveBusyState True
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CleanCleanup
End Sub

'@label:import
'@sub-title Perform the setup import from the configured workbook into the host workbook.
'@details
'Opens the source workbook, iterates over each requested sheet, matches
'ListObjects between source and host, and delegates import to
'CustomTable.Import. Handles sheet protection/unprotection and export
'column alignment via PrepareImport.
'
'PrepareImport renames two Choices columns for the length of the run and
'PostImport puts the names back. PostImport runs exactly once per call,
'whichever way the run ends: the success path guards it with its own handler
'and the failure path checks postImportDone. A second run would leave two
'columns named "Translated Label" and no "Formula Label".
'@param pass Passwords. Password handler for worksheet protection.
'@param sheetsList BetterArray. Worksheets participating in the import.
'@throws ProjectError.ObjectNotInitialized When no password handler is supplied.
'@throws ProjectError.InvalidArgument When the sheets list is missing or empty.
'@export
Public Sub Import(ByVal pass As Passwords, ByVal sheetsList As BetterArray)
    Dim importBook As Workbook
    Dim hostWorkbook As Workbook
    Dim sheetName As String
    Dim idx As Long
    Dim importSheet As Worksheet
    Dim hostSheet As Worksheet
    Dim importList As ListObject
    Dim hostList As ListObject
    Dim importTable As CustomTable
    Dim hostTable As CustomTable
    Dim sheetUnlocked As Boolean
    Dim lastSheetName As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim dictionaryFormat As BetterArray
    Dim formatList As BetterArray
    Dim postImportDone As Boolean
    Dim totalSheets As Long
    Dim doneSheets As Long

    ValidatePasswords pass
    EnsureSheetsList sheetsList

    EnterBusyState
    On Error GoTo ImportFailed

        PrepareImport pass, sheetsList, renameChoices:=True
        Set importBook = EnsureImportWorkbook()
        Set hostWorkbook = ThisWorkbook
        'One list for the whole run. Only the dictionary reads it.
        Set dictionaryFormat = DictionaryFormatHeaders()
        totalSheets = sheetsList.Length

        For idx = sheetsList.LowerBound To sheetsList.UpperBound
            sheetName = CStr(sheetsList.Item(idx))

            Set formatList = Nothing
            If SameSheetName(sheetName, DICTIONARY_SHEET_NAME) Then Set formatList = dictionaryFormat

            Set importSheet = TryWorksheet(importBook, sheetName)
            Set hostSheet = TryWorksheet(hostWorkbook, sheetName)

            If Not (importSheet Is Nothing) Then
                If Not (hostSheet Is Nothing) Then
                    WriteInfo "Importing " & sheetName & "..."
                    pass.UnProtect sheetName
                    sheetUnlocked = True
                    lastSheetName = sheetName

                    For Each importList In importSheet.ListObjects
                        Set hostList = TryListObject(hostSheet, importList.Name)
                        If Not hostList Is Nothing Then
                            Set importTable = CustomTable.Create(importList)
                            Set hostTable = CustomTable.Create(hostList)
                            hostTable.Import importTable, _
                                              keepSourceHeaders:=SameSheetName(sheetName, TRANSLATIONS_SHEET_NAME), _
                                              formatHeaders:=formatList, _
                                              insertShift:=SameSheetName(sheetName, ANALYSIS_SHEET_NAME), _
                                              strictColumnSearch:=True
                        End If
                    Next importList

                    ProtectWorksheet pass, sheetName
                    sheetUnlocked = False
                End If
            End If

            doneSheets = doneSheets + 1
            WriteProgress ProgressPercent(doneSheets, totalSheets)
        Next idx

    'Cleanup gets its own handler. With the failure handler still armed a raise
    'in here would re-enter it and run PostImport a second time.
    On Error GoTo ImportCleanupFailed
        postImportDone = True
        PostImport pass, sheetsList
        CloseImportWorkbook
    On Error GoTo 0
    LeaveBusyState
    Exit Sub

ImportCleanup:
    'Both handlers reach this label through Resume, which LEAVES the handler.
    'A second error raised while a handler is still active is not trappable and
    'ends the run where it stands, so the recovery below has to sit outside it.
    'That is what used to keep the borrowed "Formula Label" name on the Choices
    'table: re-protecting the sheet that had just refused the write raised, and
    'PostImport never ran.
    On Error Resume Next
        If sheetUnlocked Then
            ProtectWorksheet pass, lastSheetName
            sheetUnlocked = False
        End If
        If Not postImportDone Then
            postImportDone = True
            PostImport pass, sheetsList
        End If
        CloseImportWorkbook
    On Error GoTo 0
    LeaveBusyState True
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

ImportCleanupFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume ImportCleanup

ImportFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume ImportCleanup
End Sub


'@section Workbook-driven import
'===============================================================================
'@label:import-from-workbook
'@sub-title Perform the setup import leveraging dedicated worksheet classes.
'@details
'Uses domain-specific classes (LLdictionary, LLChoices, LLExport, Analysis,
'SetupTranslationsTable) to import each worksheet from the source workbook.
'Falls back to the default sheets list when none is provided.
'
'The sheet list is resolved before anything else, and PrepareImport is handed
'the resolved list from inside the protected block. PrepareImport exits on a
'missing list, so handing it the raw argument used to skip the export-row sync
'on every call that named no sheets - and the run still imported Exports into a
'host table that was too short.
'@param pass Passwords. Password handler for worksheet protection.
'@param sheetsList Optional BetterArray. Worksheets to import. Defaults to all five setup sheets.
'@throws ProjectError.ObjectNotInitialized When no password handler is supplied.
'@export
Public Sub ImportFromWorkbook(ByVal pass As Passwords, Optional ByVal sheetsList As BetterArray = Nothing)
    Dim importBook As Workbook
    Dim hostWorkbook As Workbook
    Dim resolvedSheets As BetterArray
    Dim sheetKeys As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    ValidatePasswords pass
    Set resolvedSheets = ResolveImportSheets(sheetsList)
    sheetKeys = SheetListKeys(resolvedSheets)

    EnterBusyState
    On Error GoTo ImportWorkbookFailed

        'The exports produced by an export for migration start on row 1, column 1.
        PrepareImport pass, resolvedSheets, True, SOURCE_START_ROW, SOURCE_START_COLUMN

        Set importBook = EnsureImportWorkbook()
        Set hostWorkbook = ThisWorkbook

        If SheetInKeys(sheetKeys, DICTIONARY_SHEET_NAME) Then
            ImportDictionaryUsingClass pass, hostWorkbook, importBook
        End If
        WriteProgress 20

        If SheetInKeys(sheetKeys, EXPORT_SHEET_NAME) Then
            ImportExportUsingClass pass, hostWorkbook, importBook
        End If
        WriteProgress 40

        If SheetInKeys(sheetKeys, CHOICES_SHEET_NAME) Then
            ImportChoicesUsingClass pass, hostWorkbook, importBook
        End If
        WriteProgress 60

        If SheetInKeys(sheetKeys, ANALYSIS_SHEET_NAME) Then
            ImportAnalysisUsingClass pass, hostWorkbook, importBook
        End If
        WriteProgress 80

        If SheetInKeys(sheetKeys, TRANSLATIONS_SHEET_NAME) Then
            ImportTranslationsUsingClass pass, hostWorkbook, importBook
        End If
        WriteProgress 100

    On Error GoTo 0
    CloseImportWorkbook
    LeaveBusyState
    Exit Sub

ImportWorkbookCleanup:
    'Reached through Resume, so the close below runs outside the handler. A
    'second error inside an active handler is not trappable.
    On Error Resume Next
        CloseImportWorkbook
    On Error GoTo 0
    LeaveBusyState True
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

ImportWorkbookFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume ImportWorkbookCleanup
End Sub

'@label:import-dictionary
'@sub-title Import dictionary worksheet using LLdictionary.
Private Sub ImportDictionaryUsingClass(ByVal pass As Passwords, _
                                       ByVal hostWorkbook As Workbook, _
                                       ByVal importWorkbook As Workbook)

    Dim hostSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim dictionaryManager As LLdictionary
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set hostSheet = TryWorksheet(hostWorkbook, DICTIONARY_SHEET_NAME)
    If hostSheet Is Nothing Then Exit Sub
    Set sourceSheet = TryWorksheet(importWorkbook, DICTIONARY_SHEET_NAME)
    If sourceSheet Is Nothing Then Exit Sub

    WriteInfo "Importing the dictionary..."

    On Error GoTo ImportFailed
        pass.UnProtect DICTIONARY_SHEET_NAME
        sheetUnlocked = True

        Set dictionaryManager = LLdictionary.Create(hostSheet, DICTIONARY_HOST_START_ROW, DICTIONARY_HOST_START_COLUMN)
        dictionaryManager.Import sourceSheet, SOURCE_START_ROW, SOURCE_START_COLUMN

ImportCleanup:
        On Error Resume Next
            If sheetUnlocked Then
                ProtectWorksheet pass, DICTIONARY_SHEET_NAME
                sheetUnlocked = False
            End If
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
        Exit Sub

ImportFailed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume ImportCleanup
End Sub

'@label:import-choices
'@sub-title Import choices worksheet using LLChoices.
Private Sub ImportChoicesUsingClass(ByVal pass As Passwords, _
                                    ByVal hostWorkbook As Workbook, _
                                    ByVal importWorkbook As Workbook)

    Dim hostSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim choicesManager As LLChoices
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set hostSheet = TryWorksheet(hostWorkbook, CHOICES_SHEET_NAME)
    If hostSheet Is Nothing Then Exit Sub
    Set sourceSheet = TryWorksheet(importWorkbook, CHOICES_SHEET_NAME)
    If sourceSheet Is Nothing Then Exit Sub

    WriteInfo "Importing the choices..."

    On Error GoTo ImportFailed
        pass.UnProtect CHOICES_SHEET_NAME
        sheetUnlocked = True

        Set choicesManager = LLChoices.Create(hostSheet, CHOICES_HOST_START_ROW, CHOICES_HOST_START_COLUMN)
        choicesManager.Import sourceSheet, SOURCE_START_ROW, SOURCE_START_COLUMN

ImportCleanup:
        On Error Resume Next
            If sheetUnlocked Then
                ProtectWorksheet pass, CHOICES_SHEET_NAME
                sheetUnlocked = False
            End If
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
        Exit Sub

ImportFailed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume ImportCleanup
End Sub

'@label:import-exports
'@sub-title Import export specifications via LLExport.
Private Sub ImportExportUsingClass(ByVal pass As Passwords, _
                                   ByVal hostWorkbook As Workbook, _
                                   ByVal importWorkbook As Workbook)

    Dim hostSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim exportsManager As LLExport
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set hostSheet = TryWorksheet(hostWorkbook, EXPORT_SHEET_NAME)
    If hostSheet Is Nothing Then Exit Sub
    Set sourceSheet = TryWorksheet(importWorkbook, EXPORT_SHEET_NAME)
    If sourceSheet Is Nothing Then Exit Sub

    WriteInfo "Importing the exports..."

    On Error GoTo ImportFailed
        pass.UnProtect EXPORT_SHEET_NAME
        sheetUnlocked = True

        Set exportsManager = LLExport.Create(hostSheet, EXPORT_HOST_START_ROW, EXPORT_HOST_START_COLUMN)
        exportsManager.ImportSpecs sourceSheet, SOURCE_START_ROW, SOURCE_START_COLUMN

ImportCleanup:
        On Error Resume Next
            If sheetUnlocked Then
                ProtectWorksheet pass, EXPORT_SHEET_NAME
                sheetUnlocked = False
            End If
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
        Exit Sub

ImportFailed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume ImportCleanup
End Sub

'@label:import-analysis
'@sub-title Import analysis worksheet content via Analysis class.
Private Sub ImportAnalysisUsingClass(ByVal pass As Passwords, _
                                     ByVal hostWorkbook As Workbook, _
                                     ByVal importWorkbook As Workbook)

    Dim hostSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim analysisManager As Analysis
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set hostSheet = TryWorksheet(hostWorkbook, ANALYSIS_SHEET_NAME)
    If hostSheet Is Nothing Then Exit Sub
    Set sourceSheet = TryWorksheet(importWorkbook, ANALYSIS_SHEET_NAME)
    If sourceSheet Is Nothing Then Exit Sub

    WriteInfo "Importing the analysis..."

    On Error GoTo ImportFailed
        pass.UnProtect ANALYSIS_SHEET_NAME
        sheetUnlocked = True

        Set analysisManager = Analysis.Create(hostSheet)
        analysisManager.Import sourceSheet

ImportCleanup:
        On Error Resume Next
            If sheetUnlocked Then
                ProtectWorksheet pass, ANALYSIS_SHEET_NAME
                sheetUnlocked = False
            End If
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
        Exit Sub

ImportFailed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume ImportCleanup
End Sub

'@label:import-translations
'@sub-title Import translations worksheet using CustomTable and DataSheet.
Private Sub ImportTranslationsUsingClass(ByVal pass As Passwords, _
                                         ByVal hostWorkbook As Workbook, _
                                         ByVal importWorkbook As Workbook)

    Dim hostSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim hostTable As ListObject
    Dim hostTableAdapter As CustomTable
    Dim sourceTableAdapter As DataSheet
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set hostSheet = TryWorksheet(hostWorkbook, TRANSLATIONS_SHEET_NAME)
    If hostSheet Is Nothing Then Exit Sub
    Set sourceSheet = TryWorksheet(importWorkbook, TRANSLATIONS_SHEET_NAME)
    If sourceSheet Is Nothing Then Exit Sub

    Set hostTable = ResolveTranslationsList(hostSheet)
    If hostTable Is Nothing Then Exit Sub

    WriteInfo "Importing the translations..."

    On Error GoTo ImportFailed
        pass.UnProtect TRANSLATIONS_SHEET_NAME
        sheetUnlocked = True

        Set hostTableAdapter = CustomTable.Create(hostTable)
        Set sourceTableAdapter = DataSheet.Create(sourceSheet, TRANSLATION_SOURCE_START_ROW, TRANSLATION_SOURCE_START_COLUMN)
        hostTableAdapter.Import sourceTableAdapter, keepSourceHeaders:=True

        'The rows are new to the table: the next update reviews the labels
        'no range of the setup produces with the user.
        SetupTranslationsTable.Create(hostTable).MarkImported

ImportCleanup:
        On Error Resume Next
            If sheetUnlocked Then
                ProtectWorksheet pass, TRANSLATIONS_SHEET_NAME
                sheetUnlocked = False
            End If
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
        Exit Sub

ImportFailed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume ImportCleanup
End Sub


'@label:resolve-translations-list
'@sub-title Resolve the translations ListObject from a worksheet.
Private Function ResolveTranslationsList(ByVal sheetRef As Worksheet, _
                                         Optional ByVal preferred As ListObject = Nothing) As ListObject
    Dim candidate As ListObject

    If sheetRef Is Nothing Then Exit Function

    If Not preferred Is Nothing Then
        Set candidate = TryListObject(sheetRef, preferred.Name)
    End If

    If candidate Is Nothing Then
        Set candidate = TryListObject(sheetRef, TRANSLATIONS_TABLE_NAME)
    End If

    If candidate Is Nothing Then
        On Error Resume Next
            If sheetRef.ListObjects.Count > 0 Then
                Set candidate = sheetRef.ListObjects(1)
            End If
        On Error GoTo 0
    End If

    Set ResolveTranslationsList = candidate
End Function


'@section Workbook driven Export
'===============================================================================
'@label:export-to-workbook
'@sub-title Export all setup worksheets to a new workbook.
'@details
'Creates or reuses an export workbook, then exports each setup worksheet
'(Dictionary, Choices, Exports, Analysis, Translations) using domain-specific
'managers, along with hidden names and formatting data.
'
'The caller owns any workbook it supplies through outwb. On that path the
'service writes the sheets and stops: it saves nothing, leaves LastExportFile
'empty, leaves the workbook open, and removes no sheet from it. A workbook the
'service creates itself is saved to the configured folder, closed, and stripped
'of the blank sheets Workbooks.Add gave it.
'@param outwb Optional Workbook. Pre-existing target workbook. When Nothing, creates a new one.
'@export
Public Sub Export(Optional ByVal outwb As Workbook)

    Dim dictionaryManager As LLdictionary
    Dim translationManager As SetupTranslationsTable
    Dim translationsTable As ListObject
    Dim analysisManager As Analysis
    Dim exportsManager As LLExport
    Dim choicesManager As LLChoices
    Dim formatManager As LLFormat
    Dim targetWorkbook As Workbook
    Dim hostWorkbook As Workbook
    Dim hostSheet As Worksheet
    Dim formatterSheet As Worksheet
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim store As HiddenNames

    this.lastExportFile = vbNullString

    If outwb Is Nothing Then
        Set targetWorkbook = EnsureExportWorkbook()
    Else
        Set targetWorkbook = outwb
    End If

    If targetWorkbook Is Nothing Then Exit Sub
    Set hostWorkbook = ThisWorkbook
    If hostWorkbook Is Nothing Then Exit Sub

    EnterBusyState
    On Error GoTo ExportFailed

        'Export the dictionary
        Set hostSheet = TryWorksheet(hostWorkbook, DICTIONARY_SHEET_NAME)
        If Not hostSheet Is Nothing Then
            WriteInfo "Exporting the dictionary..."
            Set dictionaryManager = LLdictionary.Create(hostSheet, DICTIONARY_HOST_START_ROW, DICTIONARY_HOST_START_COLUMN)
            dictionaryManager.Export targetWorkbook, Hide:=xlSheetVisible
        End If
        WriteProgress 15

        'Export the choices
        Set hostSheet = TryWorksheet(hostWorkbook, CHOICES_SHEET_NAME)
        If Not hostSheet Is Nothing Then
            WriteInfo "Exporting the choices..."
            Set choicesManager = LLChoices.Create(hostSheet, CHOICES_HOST_START_ROW, CHOICES_HOST_START_COLUMN)
            choicesManager.Export targetWorkbook, Hide:=xlSheetVisible
        End If
        WriteProgress 30

        'Export the exports
        Set hostSheet = TryWorksheet(hostWorkbook, EXPORT_SHEET_NAME)
        If Not hostSheet Is Nothing Then
            WriteInfo "Exporting the exports..."
            Set exportsManager = LLExport.Create(hostSheet, EXPORT_HOST_START_ROW, EXPORT_HOST_START_COLUMN)
            exportsManager.ExportSpecs targetWorkbook, Hide:=xlSheetVisible
        End If
        WriteProgress 45

        'Export the analysis
        Set hostSheet = TryWorksheet(hostWorkbook, ANALYSIS_SHEET_NAME)
        If Not hostSheet Is Nothing Then
            WriteInfo "Exporting the analysis..."
            Set analysisManager = Analysis.Create(hostSheet)
            analysisManager.Export targetWorkbook, Hide:=xlSheetVisible
        End If
        WriteProgress 60

        'Export the translations
        Set hostSheet = TryWorksheet(hostWorkbook, TRANSLATIONS_SHEET_NAME)
        If Not hostSheet Is Nothing Then
            Set translationsTable = ResolveTranslationsList(hostSheet)

            If Not translationsTable Is Nothing Then
                WriteInfo "Exporting the translations..."
                Set translationManager = SetupTranslationsTable.Create(translationsTable)
                translationManager.Export targetWorkbook, Hide:=xlSheetVisible
            End If
        End If
        WriteProgress 75

        'Export hidden names in the workbook
        Set store = HiddenNames.Create(hostWorkbook)
        store.ExportNamesToWorkbook targetWorkbook
        WriteProgress 85

        'Export the LLFormat
        Set hostSheet = TryWorksheet(hostWorkbook, FORMATTER_SHEET_NAME)
        If Not hostSheet Is Nothing Then
            WriteInfo "Exporting the formats..."
            Set formatManager = LLFormat.Create(hostSheet)
            formatManager.Export targetWorkbook

            'Set the exported worksheet to hidden, still reachable from the sheet menu
            Set formatterSheet = TryWorksheet(targetWorkbook, FORMATTER_SHEET_NAME)
            If Not formatterSheet Is Nothing Then
                formatterSheet.Visible = xlSheetHidden
            End If
        End If

        RemoveExportBaseSheets targetWorkbook
        WriteProgress 100

    'Cleanup gets its own handler: EnsureSavedExport can raise, and the export
    'workbook has to close whether it saved or not.
    On Error GoTo ExportCleanupFailed
        EnsureSavedExport
        CloseExportWorkbook
    On Error GoTo 0
    LeaveBusyState
    Exit Sub

ExportCleanup:
    'Both handlers reach this label through Resume, which leaves the handler. A
    'second error inside an active handler is not trappable, so the close has to
    'sit outside it or a half-written export workbook stays open.
    On Error Resume Next
        CloseExportWorkbook
    On Error GoTo 0
    LeaveBusyState True
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

ExportCleanupFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume ExportCleanup

ExportFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume ExportCleanup
End Sub


'@section Preparation helpers
'===============================================================================
'@label:prepare-import
'@sub-title Ensure the exports and dictionary sheets are aligned prior to import.
Private Sub PrepareImport(ByVal pass As Passwords, _
                          Optional ByVal sheetsList As BetterArray, _
                          Optional ByVal forceSync As Boolean = True, _
                          Optional ByVal impStartRow As Long = EXPORT_HOST_START_ROW, _
                          Optional ByVal impStartColumn As Long = EXPORT_HOST_START_COLUMN, _
                          Optional ByVal renameChoices As Boolean = False)

    Dim importBook As Workbook
    Dim hostWorkbook As Workbook
    Dim currentExportsSheet As Worksheet
    Dim importExportsSheet As Worksheet
    Dim importDictionarySheet As Worksheet
    Dim hostDictionarySheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim choicesTable As ListObject
    Dim currentExportObj As LLExport
    Dim importExportObj As LLExport
    Dim currentDictionary As LLdictionary
    Dim canSyncExports As Boolean
    Dim targetExports As Long
    Dim previousExports As Long
    Dim updatedExports As Long
    Dim dictUnlocked As Boolean
    Dim exportsUnlocked As Boolean
    Dim choiUnlocked As Boolean
    Dim currentChoiceData As CustomTable
    Dim sheetKeys As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If sheetsList Is Nothing Then Exit Sub

    sheetKeys = SheetListKeys(sheetsList)

    If (Not SheetInKeys(sheetKeys, EXPORT_SHEET_NAME)) And _
       (Not forceSync) And _
       (Not SheetInKeys(sheetKeys, CHOICES_SHEET_NAME)) Then Exit Sub

    Set importBook = EnsureImportWorkbook()
    Set hostWorkbook = ThisWorkbook

    On Error GoTo PrepareFailed

    If SheetInKeys(sheetKeys, EXPORT_SHEET_NAME) Or forceSync Then
        Set importExportsSheet = TryWorksheet(importBook, EXPORT_SHEET_NAME)
        Set currentExportsSheet = TryWorksheet(hostWorkbook, EXPORT_SHEET_NAME)
        Set importDictionarySheet = TryWorksheet(importBook, DICTIONARY_SHEET_NAME)
        Set hostDictionarySheet = TryWorksheet(hostWorkbook, DICTIONARY_SHEET_NAME)

        'VBA evaluates both sides of And, so the four sheets are tested one at
        'a time.
        canSyncExports = Not (importExportsSheet Is Nothing)
        If canSyncExports Then canSyncExports = Not (currentExportsSheet Is Nothing)
        If canSyncExports Then canSyncExports = Not (importDictionarySheet Is Nothing)
        If canSyncExports Then canSyncExports = Not (hostDictionarySheet Is Nothing)

        If canSyncExports Then
            Set currentExportObj = LLExport.Create(currentExportsSheet, EXPORT_HOST_START_ROW, EXPORT_HOST_START_COLUMN)
            Set importExportObj = LLExport.Create(importExportsSheet, impStartRow, impStartColumn)

            'Read both counts once. LLExport.NumberOfExports rebuilds the data
            'range and counts the rows on every read.
            targetExports = importExportObj.NumberOfExports
            updatedExports = currentExportObj.NumberOfExports

            If updatedExports < targetExports Then
                pass.UnProtect DICTIONARY_SHEET_NAME
                dictUnlocked = True
                pass.UnProtect EXPORT_SHEET_NAME
                exportsUnlocked = True

                Set currentDictionary = LLdictionary.Create(hostDictionarySheet, _
                                                            DICTIONARY_HOST_START_ROW, _
                                                            DICTIONARY_HOST_START_COLUMN)

                Do While updatedExports < targetExports
                    previousExports = updatedExports
                    currentExportObj.AddRows currentDictionary
                    updatedExports = currentExportObj.NumberOfExports

                    'A pass that adds no row would spin here for ever, with
                    'events and screen updating off. A host export table left
                    'short drops columns on the import that follows, so this is
                    'reported rather than accepted.
                    If updatedExports <= previousExports Then
                        ThrowError ProjectError.SomethingWentWrong, _
                                    "The exports table stopped growing at " & CStr(updatedExports) & _
                                    " rows and the setup file needs " & CStr(targetExports) & "."
                    End If
                Loop
            End If
        End If
    End If

    If SheetInKeys(sheetKeys, CHOICES_SHEET_NAME) And renameChoices Then
        If Not (TryWorksheet(importBook, CHOICES_SHEET_NAME) Is Nothing) Then
            Set choicesSheet = TryWorksheet(hostWorkbook, CHOICES_SHEET_NAME)

            If Not choicesSheet Is Nothing Then
                Set choicesTable = FirstListObject(choicesSheet)

                If Not choicesTable Is Nothing Then
                    pass.UnProtect CHOICES_SHEET_NAME
                    choiUnlocked = True
                    Set currentChoiceData = CustomTable.Create(choicesTable)
                    currentChoiceData.RenameColumn "Label", "Formula Label", matchCase:=False
                    currentChoiceData.RenameColumn "Translated Label", "Label", matchCase:=False
                End If
            End If
        End If
    End If

PrepareCleanup:
    On Error Resume Next
        If dictUnlocked Then
            ProtectWorksheet pass, DICTIONARY_SHEET_NAME
            dictUnlocked = False
        End If
        If exportsUnlocked Then
            ProtectWorksheet pass, EXPORT_SHEET_NAME
            exportsUnlocked = False
        End If
        If choiUnlocked Then
            ProtectWorksheet pass, CHOICES_SHEET_NAME
            choiUnlocked = False
        End If

    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

PrepareFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume PrepareCleanup
End Sub


'@label:post-import
'@sub-title Restore choices column names after import completes.
Private Sub PostImport(ByVal pass As Passwords, _
                        Optional ByVal sheetsList As BetterArray)

    Dim importBook As Workbook
    Dim hostWorkbook As Workbook
    Dim choicesSheet As Worksheet
    Dim choicesTable As ListObject
    Dim choiUnlocked As Boolean
    Dim currentChoiceData As CustomTable
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If sheetsList Is Nothing Then Exit Sub
    If Not SheetInKeys(SheetListKeys(sheetsList), CHOICES_SHEET_NAME) Then Exit Sub

    On Error GoTo PostFailed

    Set importBook = EnsureImportWorkbook()
    Set hostWorkbook = ThisWorkbook

    If Not (TryWorksheet(importBook, CHOICES_SHEET_NAME) Is Nothing) Then
        Set choicesSheet = TryWorksheet(hostWorkbook, CHOICES_SHEET_NAME)

        If Not choicesSheet Is Nothing Then
            Set choicesTable = FirstListObject(choicesSheet)

            If Not choicesTable Is Nothing Then
                pass.UnProtect CHOICES_SHEET_NAME
                choiUnlocked = True
                Set currentChoiceData = CustomTable.Create(choicesTable)
                currentChoiceData.RenameColumn "Label", "Translated Label", matchCase:=False
                currentChoiceData.RenameColumn "Formula Label", "Label", matchCase:=False
            End If
        End If
    End If

PostCleanUp:
    On Error Resume Next
        If choiUnlocked Then
            ProtectWorksheet pass, CHOICES_SHEET_NAME
            choiUnlocked = False
        End If
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

PostFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume PostCleanUp
End Sub

'@section Validation helpers
'===============================================================================
'@label:resolve-import-sheets
'@sub-title Resolve the requested sheets set, defaulting to core setup worksheets.
Private Function ResolveImportSheets(ByVal sheetsList As BetterArray) As BetterArray
    If sheetsList Is Nothing Then
        Set ResolveImportSheets = DefaultImportSheets()
    Else
        Set ResolveImportSheets = EnsureSheetsList(sheetsList)
    End If
End Function

'@label:default-import-sheets
'@sub-title Default sheet list covering the core setup worksheets.
Private Function DefaultImportSheets() As BetterArray
    Dim defaults As BetterArray

    Set defaults = New BetterArray
    defaults.LowerBound = 1
    defaults.Push DICTIONARY_SHEET_NAME, CHOICES_SHEET_NAME, EXPORT_SHEET_NAME, ANALYSIS_SHEET_NAME, TRANSLATIONS_SHEET_NAME

    Set DefaultImportSheets = defaults
End Function

'@label:sheet-list-keys
'@sub-title Build one lower-cased key string out of a sheets list.
'@details
'The list is walked once per call and every later question is answered with
'InStr. The separator sits on both sides of each name, so a short name cannot
'match inside a longer one.
Private Function SheetListKeys(ByVal sheetsList As BetterArray) As String
    Dim idx As Long
    Dim keys As String
    Dim written As Long

    If sheetsList Is Nothing Then Exit Function
    If sheetsList.Length = 0 Then Exit Function

    keys = LIST_SEPARATOR
    For idx = sheetsList.LowerBound To sheetsList.UpperBound
        keys = keys & LCase$(Trim$(CStr(sheetsList.Item(idx)))) & LIST_SEPARATOR
        written = written + 1
        If written >= MAX_SHEET_LIST_LENGTH Then Exit For
    Next idx

    SheetListKeys = keys
End Function

'@label:sheet-in-keys
'@sub-title Answer whether a sheet name sits in a key string.
Private Function SheetInKeys(ByVal sheetKeys As String, ByVal sheetName As String) As Boolean
    If LenB(sheetKeys) = 0 Then Exit Function

    SheetInKeys = (InStr(1, sheetKeys, _
                         LIST_SEPARATOR & LCase$(Trim$(sheetName)) & LIST_SEPARATOR, _
                         vbBinaryCompare) > 0)
End Function

'@label:same-sheet-name
'@sub-title Compare two worksheet names the way Excel does.
Private Function SameSheetName(ByVal leftName As String, ByVal rightName As String) As Boolean
    SameSheetName = (StrComp(Trim$(leftName), Trim$(rightName), vbTextCompare) = 0)
End Function

'@label:ensure-import-selection
'@sub-title Ensure at least one import option is selected.
Private Sub EnsureImportSelection(ByVal importDictionary As Boolean, _
                                  ByVal importChoices As Boolean, _
                                  ByVal importExports As Boolean, _
                                  ByVal importAnalysis As Boolean, _
                                  ByVal importTranslations As Boolean, _
                                  ByVal cleanSetup As Boolean)

    Const OPTIONS_WARNING As String = "Please select at least one import option (Dictionary, Choices, Exports, Analysis or Translations)."

    If Not importDictionary And Not importChoices And Not importExports _
       And Not importAnalysis And Not importTranslations Then
        If cleanSetup Then
            ReportAndThrow ProjectError.InvalidArgument, _
                            "Clean operation requires at least one worksheet to process."
        Else
            ReportAndThrow ProjectError.InvalidArgument, OPTIONS_WARNING
        End If
    End If
End Sub

'@label:ensure-import-path-configured
'@sub-title Ensure the import path has been configured.
Private Sub EnsureImportPathConfigured(ByVal value As String)
    If LenB(Trim$(value)) = 0 Then
        ReportAndThrow ProjectError.InvalidArgument, "No setup file path is configured for this import."
    End If
End Sub

'@label:validate-passwords
'@sub-title Validate availability of requested passwords object.
Private Sub ValidatePasswords(ByVal pass As Passwords)
    If pass Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Passwords handler is required for worksheet protection."
    End If
End Sub

'@label:ensure-sheets-list
'@sub-title Ensure the sheets list is initialised.
Private Function EnsureSheetsList(ByVal sheetsList As BetterArray) As BetterArray
    If sheetsList Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "Sheets list cannot be Nothing."
    End If

    If sheetsList.Length = 0 Then
        ThrowError ProjectError.InvalidArgument, "Sheets list cannot be empty."
    End If

    Set EnsureSheetsList = sheetsList
End Function

'@label:validate-import-path
'@sub-title Confirm the import path string is valid.
Private Sub ValidateImportPath(ByVal pathValue As String)
    If LenB(Trim$(pathValue)) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Import path cannot be blank."
    End If
End Sub

'@label:validate-progress-display
'@sub-title Ensure the progress display object exposes a Caption or Value property.
'@details
'An object that carries neither property is refused. The refusal used to sit
'below a plain Exit Sub, on a label nothing could jump to, so any object at all
'was accepted and the writers below then wrote to nothing.
Private Sub ValidateProgressDisplay(ByVal progressDisplay As Object)
    If progressDisplay Is Nothing Then Exit Sub
    If TestProperty(progressDisplay, this.hasCaption, "Caption") Then Exit Sub
    If TestProperty(progressDisplay, this.hasCaption, "Value") Then Exit Sub

    ThrowError ProjectError.InvalidArgument, _
                "Progress display object must expose a Caption or Value property."
End Sub

'@label:test-property
'@sub-title Test whether an object supports a named property.
'@details
'The probe reads the property. Writing to it was how this used to work, so
'asking a label whether it had a caption erased the caption. A name the probe
'does not know answers False.
Private Function TestProperty(ByVal probeObject As Object, _
                              ByRef hasProp As Boolean, _
                              ByVal propName As String) As Boolean

    Dim probe As Variant

    hasProp = False

    Select Case propName
        Case "Caption"
            On Error Resume Next
                probe = probeObject.Caption
                hasProp = (Err.Number = 0)
                Err.Clear
            On Error GoTo 0

        Case "Value"
            On Error Resume Next
                probe = probeObject.Value
                hasProp = (Err.Number = 0)
                Err.Clear
            On Error GoTo 0
    End Select

    TestProperty = hasProp
End Function

'@label:ensure-import-file-exists
'@sub-title Ensure the import workbook file exists before opening.
Private Sub EnsureImportFileExists(ByVal filePath As String)
    Dim resolved As String

    On Error Resume Next
        resolved = Dir(filePath)
    On Error GoTo 0

    If LenB(resolved) = 0 Then
        ReportAndThrow ProjectError.ElementNotFound, _
                        "No setup file was found at '" & filePath & "'."
    End If
End Sub


'@section Workbook helpers
'===============================================================================
'@label:ensure-import-workbook
'@sub-title Obtain the imported workbook, opening it when needed.
Private Function EnsureImportWorkbook() As Workbook
    Dim failureDetails As String

    EnsureImportPathConfigured this.importPath

    If this.importWorkbook Is Nothing Then
        On Error Resume Next
            Set this.importWorkbook = Workbooks.Open(this.importPath)
            failureDetails = Err.Description
            Err.Clear
        On Error GoTo 0

        If this.importWorkbook Is Nothing Then
            ReportAndThrow ProjectError.SomethingWentWrong, _
                            BuildOpenFailureMessage(this.importPath, failureDetails)
        End If
    End If

    Set EnsureImportWorkbook = this.importWorkbook
End Function

'@label:ensure-export-workbook
'@sub-title Obtain or create the export workbook, prompting for a folder when needed.
'@details
'The sheets the new workbook is born with are captured here, by reference, so
'the export can drop exactly those at the end. Deleting Worksheets(1) instead
'assumed one blank sheet and assumed every manager appends after the last one.
Private Function EnsureExportWorkbook() As Workbook

    Dim io As OSFiles

    If (this.exportWorkbook Is Nothing) Then
        If LenB(this.exportPath) = 0 Then
            If this.displayPrompt Then
                Set io = OSFiles.Create()

                io.LoadFolder

                If io.HasValidFolder() Then
                    this.exportPath = io.Folder()
                Else
                    Set EnsureExportWorkbook = Nothing
                    Exit Function
                End If
            Else
                Set EnsureExportWorkbook = Nothing
                Exit Function
            End If
        End If

        If LenB(this.exportPath) = 0 Then
            Set EnsureExportWorkbook = Nothing
            Exit Function
        End If

        Set this.exportWorkbook = Workbooks.Add
        Set this.exportBaseSheets = CaptureWorksheets(this.exportWorkbook)
    End If

    Set EnsureExportWorkbook = this.exportWorkbook
End Function

'@label:capture-worksheets
'@sub-title Hold every worksheet of a workbook in a Collection.
Private Function CaptureWorksheets(ByVal workbookRef As Workbook) As Collection
    Dim captured As Collection
    Dim sheetRef As Worksheet

    Set captured = New Collection
    If workbookRef Is Nothing Then
        Set CaptureWorksheets = captured
        Exit Function
    End If

    For Each sheetRef In workbookRef.Worksheets
        captured.Add sheetRef
    Next sheetRef

    Set CaptureWorksheets = captured
End Function

'@label:remove-export-base-sheets
'@sub-title Drop the blank sheets the export workbook was created with.
'@details
'Only a workbook the service created carries captured sheets, so a workbook the
'caller supplied keeps every sheet it had. The delete stops while one sheet is
'left, because a workbook cannot hold zero worksheets.
Private Sub RemoveExportBaseSheets(ByVal targetWorkbook As Workbook)
    Dim sheetRef As Worksheet
    Dim idx As Long
    Dim remaining As Long

    If this.exportBaseSheets Is Nothing Then Exit Sub
    If targetWorkbook Is Nothing Then Exit Sub

    remaining = targetWorkbook.Worksheets.Count

    For idx = 1 To this.exportBaseSheets.Count
        If remaining <= 1 Then Exit For

        Set sheetRef = Nothing
        On Error Resume Next
            Set sheetRef = this.exportBaseSheets.Item(idx)
        On Error GoTo 0

        If Not sheetRef Is Nothing Then
            On Error Resume Next
                sheetRef.Delete
                If Err.Number = 0 Then remaining = remaining - 1
                Err.Clear
            On Error GoTo 0
        End If
    Next idx

    Set this.exportBaseSheets = Nothing
End Sub

'@label:ensure-saved-export
'@sub-title Save the export workbook to the configured path.
'@details
'The format is stated. Without it Excel writes in whatever default format the
'machine is set to, so a machine set to .xlsb produced a binary payload under
'an .xlsx name.
Private Sub EnsureSavedExport()

    Dim targetWorkbook As Workbook
    Dim filePath As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set targetWorkbook = this.exportWorkbook
    If targetWorkbook Is Nothing Then Exit Sub
    If LenB(this.exportPath) = 0 Then Exit Sub

    filePath = BuildExportFilePath()

    'A locked file, a missing folder or a name Excel refuses raises here, and
    'the bare message never says which path was tried. The handler adds it and
    're-raises; Export's own cleanup label closes the workbook.
    On Error GoTo SaveFailed
    targetWorkbook.SaveAs filePath, _
                          FileFormat:=xlOpenXMLWorkbook, _
                          CreateBackup:=False, _
                          ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    On Error GoTo 0
    this.lastExportFile = filePath
    Exit Sub

SaveFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    Err.Raise errNumber, errSource, _
              errDescription & " (export file: " & filePath & ")"
End Sub

'@label:dictionary-format-headers
'@sub-title Header names the dictionary import carries its formatting for.
Private Function DictionaryFormatHeaders() As BetterArray

    Dim headers As BetterArray

    Set headers = New BetterArray
    headers.Push "Formatting Condition", _
                 "Formatting Values", _
                 "Variable Name", _
                 "Control", "Main Label", "Lock Cells", "Dev Comments"

    Set DictionaryFormatHeaders = headers
End Function

'@label:build-export-file-path
'@sub-title Build the stamped export file path from the configured folder.
'@details
'Whatever follows the last dot of the host name is dropped, so a host saved as
'.xlsm or .XLSB no longer carries its extension inside the new name. The stamp
'carries the time as well as the date, so two exports in one day are two files.
'Two exports inside the same second still land on one stamp, so a name already
'taken takes a counter behind the stamp rather than a raise from SaveAs.
Private Function BuildExportFilePath() As String
    Dim baseName As String
    Dim dotPosition As Long
    Dim stemPath As String
    Dim candidate As String
    Dim suffix As Long

    baseName = ThisWorkbook.Name
    dotPosition = InStrRev(baseName, ".")
    If dotPosition > 1 Then baseName = Left$(baseName, dotPosition - 1)

    stemPath = this.exportPath & Application.PathSeparator & _
               baseName & "_export_" & Format$(Now(), "yyyymmdd_hhnnss")

    candidate = stemPath & ".xlsx"
    Do While ExportNameIsTaken(candidate)
        suffix = suffix + 1
        candidate = stemPath & "_" & CStr(suffix) & ".xlsx"
    Loop

    BuildExportFilePath = candidate
End Function

'@label:export-name-is-taken
'@sub-title Answer whether an export path is already spoken for.
'@details
'Two exports inside the same second build the same stamp, and Excel refuses
'SaveAs against a name one of the open workbooks already carries, whatever
'folder that workbook sits in. Both the disk and the open workbooks are read
'here so the caller can step to the next free name instead of raising.
'@param filePath String the candidate path.
'@return Boolean True when a file or an open workbook holds that name.
Private Function ExportNameIsTaken(ByVal filePath As String) As Boolean
    Dim fileName As String
    Dim openBook As Workbook
    Dim onDisk As String

    'Shielded: Dir$ raises on a folder that is not there, and a missing folder
    'is the SaveAs call's own failure to report, not this function's.
    On Error Resume Next
    onDisk = Dir$(filePath)
    Err.Clear
    On Error GoTo 0

    If LenB(onDisk) <> 0 Then
        ExportNameIsTaken = True
        Exit Function
    End If

    fileName = Mid$(filePath, InStrRev(filePath, Application.PathSeparator) + 1)

    For Each openBook In Application.Workbooks
        If StrComp(openBook.Name, fileName, vbTextCompare) = 0 Then
            ExportNameIsTaken = True
            Exit Function
        End If
    Next openBook
End Function

'@label:close-import-workbook
'@sub-title Close and release the imported workbook.
Private Sub CloseImportWorkbook()
    CloseWorkbook this.importWorkbook
    Set this.importWorkbook = Nothing
End Sub

'@label:close-export-workbook
'@sub-title Close and release the export workbook.
'@details
'The folder is cleared with the workbook, so a second export on the same
'instance needs a fresh SetExportFolder call.
Private Sub CloseExportWorkbook()
    CloseWorkbook this.exportWorkbook
    Set this.exportWorkbook = Nothing
    Set this.exportBaseSheets = Nothing
    this.exportPath = vbNullString
End Sub


'@label:close-workbook
'@sub-title Close workbook safely without raising errors.
Private Sub CloseWorkbook(ByVal workbookRef As Workbook)
    On Error Resume Next
        If Not workbookRef Is Nothing Then
            workbookRef.Close savechanges:=False
        End If
    On Error GoTo 0
End Sub

'@label:try-worksheet
'@sub-title Resolve a worksheet, answering Nothing when it is absent.
Private Function TryWorksheet(ByVal workbookRef As Workbook, ByVal sheetName As String) As Worksheet
    If workbookRef Is Nothing Then Exit Function

    On Error Resume Next
        Set TryWorksheet = workbookRef.Worksheets(sheetName)
        Err.Clear
    On Error GoTo 0
End Function

'@label:try-list-object
'@sub-title Resolve a ListObject, answering Nothing when it is absent.
Private Function TryListObject(ByVal sheetRef As Worksheet, ByVal listName As String) As ListObject
    If sheetRef Is Nothing Then Exit Function

    On Error Resume Next
        Set TryListObject = sheetRef.ListObjects(listName)
        Err.Clear
    On Error GoTo 0
End Function

'@label:first-list-object
'@sub-title Resolve the first ListObject of a worksheet.
Private Function FirstListObject(ByVal sheetRef As Worksheet) As ListObject
    If sheetRef Is Nothing Then Exit Function

    On Error Resume Next
        If sheetRef.ListObjects.Count > 0 Then
            Set FirstListObject = sheetRef.ListObjects(1)
        End If
        Err.Clear
    On Error GoTo 0
End Function


'@section Worksheet helpers
'===============================================================================
'@label:clean-worksheet-tables
'@sub-title Clean all ListObjects within a worksheet.
Private Sub CleanWorksheetTables(ByVal targetSheet As Worksheet)
    Dim listRef As ListObject
    Dim table As CustomTable

    For Each listRef In targetSheet.ListObjects
        Set table = CustomTable.Create(listRef)
        table.Clean
    Next listRef
End Sub

'@label:clear-worksheet-comments
'@sub-title Remove the classic comments from a worksheet.
'@details
'Range.ClearComments covers classic comments. Threaded comments are reached
'through Worksheet.CommentsThreaded, which Mac Excel does not carry.
Private Sub ClearWorksheetComments(ByVal targetSheet As Worksheet)
    Dim usedCells As Range

    On Error Resume Next
        Set usedCells = targetSheet.UsedRange
    On Error GoTo 0

    If usedCells Is Nothing Then Exit Sub

    On Error Resume Next
        usedCells.ClearComments
        Err.Clear
    On Error GoTo 0
End Sub

'@label:protect-worksheet
'@sub-title Protect worksheet with default rules.
'@details
'Every protect call in this class comes through here. Passwords.Protect
'defaults both flags to True, so a bare call left the Choices sheet allowing
'shapes and row deletion at the end of every import.
Private Sub ProtectWorksheet(ByVal pass As Passwords, ByVal sheetName As String)
    pass.Protect sheetName, _
                      (SameSheetName(sheetName, TRANSLATIONS_SHEET_NAME) Or SameSheetName(sheetName, ANALYSIS_SHEET_NAME)), _
                      SameSheetName(sheetName, TRANSLATIONS_SHEET_NAME)
End Sub


'@section Messaging helpers
'===============================================================================
'@label:write-info
'@sub-title Display informational text through the progress object.
Private Sub WriteInfo(ByVal message As String)
    If Not this.hasCaption Then Exit Sub
    If this.progressDisplay Is Nothing Then Exit Sub

    On Error Resume Next
        this.progressDisplay.Caption = message
        this.progressDisplay.Value = message
        Err.Clear
    On Error GoTo 0
End Sub

'@label:write-progress
'@sub-title Update a textual progress bar representation.
'@details
'The percentage is clamped on both sides in plain VBA. A worksheet function
'call for a two-value minimum is a round trip into Excel, and it left the low
'side open, so a negative percentage reached String$ and raised error 5.
Private Sub WriteProgress(ByVal percentage As Long)
    Const TOTAL_BARS As Long = 30

    Dim completedBars As Long
    Dim clamped As Long
    Dim bar As String

    If Not this.hasCaption Then Exit Sub
    If this.progressDisplay Is Nothing Then Exit Sub

    clamped = percentage
    If clamped < 0 Then clamped = 0
    If clamped > 100 Then clamped = 100

    completedBars = (TOTAL_BARS * clamped) \ 100
    bar = "[" & String$(completedBars, "|") & Space$(TOTAL_BARS - completedBars) & "] " & CStr(clamped) & "%"

    On Error Resume Next
        this.progressDisplay.Caption = bar
        this.progressDisplay.Value = bar
        Err.Clear
    On Error GoTo 0
End Sub

'@label:progress-percent
'@sub-title Turn a done-out-of-total pair into a percentage.
Private Function ProgressPercent(ByVal completed As Long, ByVal total As Long) As Long
    If total <= 0 Then Exit Function

    ProgressPercent = CLng((completed * 100) / total)
End Function

'@label:build-open-failure-message
'@sub-title Combine failure details when opening a workbook fails.
Private Function BuildOpenFailureMessage(ByVal sourcePath As String, ByVal failureDetails As String) As String
    Dim message As String

    message = "Unable to open the setup workbook at '" & sourcePath & "'."
    If LenB(Trim$(failureDetails)) > 0 Then
        message = message & " (" & failureDetails & ")"
    End If

    BuildOpenFailureMessage = message
End Function


'@section Error handling
'===============================================================================
'@label:report-and-throw
'@sub-title Report message then raise a project error.
'@param errNumber Long. A ProjectError code.
'@param message String. Text shown to the user and carried by the error.
Private Sub ReportAndThrow(ByVal errNumber As Long, ByVal message As String)
    WriteInfo message
    ThrowError errNumber, message
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception.
'@details
'The code is typed Long, which is the repo rule: an Enum in a parameter list
'misbehaves on Mac Excel, and every call site already passes a ProjectError
'member that widens to Long on its own.
'@param errNumber Long. A ProjectError code.
'@param message String. Text carried by the error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise errNumber, CLASS_NAME, message
End Sub

'@label:terminate
'@sub-title Cleanup resources on destruction.
Private Sub Class_Terminate()
    On Error Resume Next
        CloseImportWorkbook
        CloseExportWorkbook
    On Error GoTo 0

    this.busyDepth = 0
    RestoreApplicationState True

    Set this.progressDisplay = Nothing
    Set this.appScope = Nothing
End Sub
