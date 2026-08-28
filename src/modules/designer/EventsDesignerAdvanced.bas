Attribute VB_Name = "EventsDesignerAdvanced"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Non-core ribbon callbacks for the designer workbook.")
'@depends DesignerPreparation, DesignerEntry, EventsDesignerCore, RibbonDev, ApplicationState, OSFiles, HiddenNames, BetterArray, DropdownLists, Checking, GenerationLog, SetupTranslationsTable, BuildSteps, GenerationHost, ProgressBar, TemporaryRepos
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'Non-core ribbon logics are callbacks whose absence will not fire a
'warning at workbook opening on the designer. They only execute in
'response to explicit user actions (onAction), never at ribbon load
'time (getLabel, getPressed, getVisible).
'
'The DesignerEntry every callback works through is the shared one,
'EventsDesignerCore.EntryManager(). It carries the held translator with it,
'so a press reads no translation table.

Private Const SHEET_MAIN As String = "Main"
Private Const SHEET_DROPDOWNS As String = "__dropdowns"
Private Const PROMPT_TITLE As String = "Designer"

Private Const SHEET_TRANSLATIONS As String = "Translations"

'Dropdown name used by DesignerPreparation for setup languages
Private Const DROP_SETUP_LANGUAGES As String = "__setup_languages"

'Leading shape of the internal tag columns of a setup translations table.
'The fallback header read drops these before the dropdown update.
Private Const INTERNAL_TAG_LEAD As String = "__"

'The status texts GenerateOne writes at its milestones, for a status
'target that shows them (the multi driver's row result cell).
Private Const STATUS_TRANSFER As String = "transfer"
Private Const STATUS_LINELIST As String = "linelist"
Private Const STATUS_DROPDOWNS As String = "dropdowns"
Private Const STATUS_ANALYSES As String = "analyses"
Private Const STATUS_SAVE As String = "save"

'The lead of a step outcome that failed, and the mark before the kept
'path at its end. Both are the shape BuildSteps answers.
Private Const OUTCOME_ERROR_LEAD As String = "ERROR "
Private Const KEPT_MARK As String = " | kept "

'The separator of the two totals BuildSteps.BuildCounts answers
Private Const COUNTS_SEP As String = "|"

'The bundles this module files on its own: the kept file of a failed
'build, a file-name clash found before the build, and a build instance
'that would not quit.
Private Const TITLE_BUILD_ABORTED As String = "build aborted"
Private Const TITLE_PRE_FLIGHT As String = "build pre-flight"
Private Const TITLE_BUILD_INSTANCE As String = "build instance"

'The module every step name is qualified with when the host runs it
Private Const STEP_MODULE_LEAD As String = "BuildSteps."

'The step names, as GenerateOne asks for them
Private Const STEP_BEGIN As String = "BuildBegin"
Private Const STEP_BEGIN_ENTRIES As String = "BuildBeginEntries"
Private Const STEP_LINELIST As String = "BuildLinelist"
Private Const STEP_SHEET_COUNT As String = "BuildSheetCount"
Private Const STEP_SHEET As String = "BuildSheet"
Private Const STEP_DROPDOWNS As String = "BuildDropdowns"
Private Const STEP_ANALYSES As String = "BuildAnalyses"
Private Const STEP_SAVE As String = "BuildSave"
Private Const STEP_CHECKINGS As String = "BuildCheckings"
Private Const STEP_COUNTS As String = "BuildCounts"
Private Const STEP_ABORT As String = "BuildAbort"

'The separators of the entries text BuildBeginEntries takes: one entry
'per vertical tab, the tag before the first equals sign.
Private Const ENTRY_SEP As String = vbVerticalTab
Private Const ENTRY_EQUALS As String = "="

'The entry tags a build reads, in the order they are handed over as text
'to a build in another instance
Private Const ENTRY_TAGS As String = "setuppath,geopath,lldir,llname,llpassword,debugpassword," & _
                                     "setuplang,lllang,epiweekstart,design,temppath"

'The name of the unfinished workbook a failed build keeps, the one
'Linelist.DiscardBuild writes. The pre-flight checks it is open nowhere.
Private Const KEPT_FILE_NAME As String = "__temp.xlsb"

'The bar over the single build: two names on the Main sheet, owner hand
'work on the designer. A missing name means no bar.
Private Const RNG_PROGRESS_BAR As String = "RNG_ProgressBar"
Private Const RNG_PROGRESS_STATUS As String = "RNG_ProgressStatus"

'The steps of a build besides the data entry sheets: begin, linelist,
'dropdowns, analyses, save. The bar's maximum starts here and grows by
'the sheet count once BuildSheetCount answers.
Private Const FIXED_STEPS As Long = 5

'The log of the current or last generation run. Both drivers open it
'through StartRunLog and every flush goes through CollectIntoLog, so
'one run holds one record. The reference outlives Finish on purpose:
'the record is what a later text re-export reads while the designer
'stays open.
Private heldLog As GenerationLog

'What the run has built so far. GenerateOne adds the totals of each build,
'so a multi run counts every row of the batch, and FinishRunLog hands them
'to the closing bundle. StartRunLog puts them back to zero.
Private builtSheets As Long
Private builtVariables As Long

'The path of the file the last failed build kept, or empty. GenerateOne
'sets it from the step outcome; clickGenerate's cleanup names it in the
'report.
Private keptFilePath As String


'@section Run log services
'===============================================================================

'@Description("Open the log of a new generation run on the designer __check sheet.")
'@details
'Both drivers call this once at the start of a run. The single build
'passes the setup path and the linelist name for the run header; the
'multi driver opens the log bare, since every row names itself in its
'own header bundle.
'@param setupPath String. The setup file path of the run. Empty skips the header entry.
'@param linelistName String. The output linelist name. Empty skips the header entry.
'@param designerBook Workbook. The designer holding the __check sheet. Nothing reads ThisWorkbook.
'@return GenerationLog. The opened log.
Public Function StartRunLog(Optional ByVal setupPath As String = vbNullString, _
                            Optional ByVal linelistName As String = vbNullString, _
                            Optional ByVal designerBook As Workbook = Nothing) As GenerationLog
    'A designer press logs on itself. The headless build runs this code from
    'the driver workbook, which carries no __check sheet, so it hands over
    'the designer copy and the report lands beside the linelist.
    If designerBook Is Nothing Then Set designerBook = ThisWorkbook
    Set heldLog = GenerationLog.Create(designerBook)
    heldLog.Start setupPath, linelistName
    builtSheets = 0
    builtVariables = 0
    Set StartRunLog = heldLog
End Function

'@Description("The log of the current or last run. Nothing before the first run.")
Public Function RunLog() As GenerationLog
    Set RunLog = heldLog
End Function

'@Description("Take one bundle into the run log. Without an open run the bundle is dropped.")
'@details
'recordOnly keeps a bundle out of the __check worksheet and puts it in
'the run's record alone, which is what the text file is written from.
'The per-variable milestones travel that way: they are a few hundred
'entries and every worksheet row costs an EntireColumn.AutoFit.
'@param checks Checking. The bundle to take.
'@param recordOnly Optional Boolean. True keeps the bundle off the worksheet.
Public Sub CollectIntoLog(ByVal checks As Checking, _
                          Optional ByVal recordOnly As Boolean = False)
    If heldLog Is Nothing Then Exit Sub
    heldLog.Collect checks, recordOnly
End Sub

'@Description("Open one section of the run log. Without an open run the call is ignored.")
'@details
'A section is one whole build. The multi driver opens one per row, so
'every bundle of that row's build lands under one heading with the parts
'of the build as subsections. The single build opens none and keeps one
'heading per part.
'@param sectionTitle String. The title of the section.
Public Sub OpenLogSection(ByVal sectionTitle As String)
    If heldLog Is Nothing Then Exit Sub
    heldLog.OpenSection sectionTitle
End Sub

'@Description("Write the open section of the run log and close it.")
'@details
'Safe to call with no section open, so a driver may close after every row
'and after the loop.
Public Sub CloseLogSection()
    If heldLog Is Nothing Then Exit Sub
    heldLog.CloseSection
End Sub

'@Description("Close the run log with the outcome text. Without an open run the call is ignored.")
'@details
'The counts the run added up ride into the closing bundle beside the
'outcome, so the last lines of the report say how much was built.
'@param outcome String. The outcome text of the run.
Public Sub FinishRunLog(ByVal outcome As String)
    If heldLog Is Nothing Then Exit Sub
    heldLog.Finish outcome, builtSheets, builtVariables
End Sub

'@Description("Bring the generation report to the front. Answers False when there is none to show.")
'@details
'The report of the run that just ended, or of the last run of a previous
'session: the sheet outlives the log object, so this builds a log over
'ThisWorkbook when no run has been opened yet and asks it for the same
'sheet. That is what lets the ribbon button work on a designer that was
'just opened.
'@return Boolean. True when a report was found and shown.
Public Function ShowRunLog() As Boolean
    Dim reportLog As GenerationLog

    Set reportLog = heldLog
    If reportLog Is Nothing Then
        On Error Resume Next
        Set reportLog = GenerationLog.Create(ThisWorkbook)
        On Error GoTo 0
    End If

    If reportLog Is Nothing Then Exit Function

    ShowRunLog = reportLog.ShowReport()
End Function

'@Description("Callback for the ribbon button that opens the generation report.")
'@EntryPoint
Public Sub clickOpenLog(ByRef ribbonControl As IRibbonControl)
    Dim entry As DesignerEntry

    On Error GoTo Cleanup

    If ShowRunLog() Then Exit Sub

    'Nothing has been generated yet in this workbook, so there is no report
    'to open. Saying so beats leaving the press with no answer at all.
    Set entry = EventsDesignerCore.EntryManager()
    MsgBox entry.TranslateMessage("MSG_NoRunLog"), _
           vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Debug.Print "clickOpenLog: "; Err.Number; Err.Description
End Sub


'@section Dev group callbacks
'===============================================================================

'@Description("Initialise the designer workbook: import translations, hide sheets, seed flags.")
'@EntryPoint
Public Sub clickDesignerInitialize(ByRef ribbonControl As IRibbonControl)
    Dim prep As DesignerPreparation
    Dim appScope As ApplicationState

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set prep = DesignerPreparation.Create(ThisWorkbook)
    prep.Prepare RibbonDev.EnsureDevelopment()

    'Preparation re-imports the translation tables, so the held pair is stale.
    EventsDesignerCore.ResetDesignerCaches

    appScope.Restore
    MsgBox "Done!", vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickDesignerInitialize: "; errNumber; errDesc
        MsgBox "Unable to initialise designer: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub


'@section Manage group callbacks
'===============================================================================

'@Description("Clear all entry input ranges on the Main sheet.")
'@EntryPoint
Public Sub clickClearEnt(ByRef ribbonControl As IRibbonControl)
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = EventsDesignerCore.EntryManager()
    entry.Clear

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickClearEnt: "; errNumber; errDesc
        MsgBox "Unable to clear entries: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub


'@section File and folder loading callbacks
'===============================================================================

'@Description("Load a setup file (dictionary): store path, extract languages, update dropdown.")
'@EntryPoint
Public Sub clickLoadFileDic()
    Dim io As OSFiles
    Dim entry As DesignerEntry
    Dim prep As DesignerPreparation
    Dim appScope As ApplicationState
    Dim setupBook As Workbook
    Dim tradSheet As Worksheet

    'Show the file dialog before entering busy state (dialog needs UI)
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsb;*.xlsx"

    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = EventsDesignerCore.EntryManager()

    'Open the selected setup workbook read-only
    Set setupBook = Workbooks.Open(io.File(), ReadOnly:=True)

    'Validate that the setup has a Translations worksheet
    On Error Resume Next
    Set tradSheet = setupBook.Worksheets(SHEET_TRANSLATIONS)
    On Error GoTo Cleanup

    If tradSheet Is Nothing Then
        setupBook.Close saveChanges:=False
        Set setupBook = Nothing
        entry.AddInfo entry.TranslateMessage("MSG_OpeAnnule"), "edition"
        GoTo Cleanup
    End If

    'Write the setup path to the Main sheet
    entry.AddInfo io.File(), "setuppath"
    entry.AddInfo entry.TranslateMessage("MSG_ChemFich"), "edition"

    'A new setup file brings its own __formatter sheet, so the designer's
    'copy stops being the live one. The styles import button sets the flag
    'again when the user wants the designer's formatter for this setup.
    Set prep = DesignerPreparation.Create(ThisWorkbook)
    prep.FormatterImported = False

    'Extract languages from the setup Translations worksheet HiddenNames
    'and update the setup languages dropdown for the designer
    ExtractAndUpdateLanguages tradSheet, entry

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    'Close the setup workbook if still open
    If Not setupBook Is Nothing Then
        setupBook.Close saveChanges:=False
    End If
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickLoadFileDic: "; errNumber; errDesc
        MsgBox "Unable to load setup file: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Load a geobase file path into the Main sheet.")
'@EntryPoint
Public Sub clickLoadGeoFile()
    Dim io As OSFiles
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState

    'Show the file dialog before entering busy state
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"

    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = EventsDesignerCore.EntryManager()
    entry.AddInfo io.File(), "geopath"

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickLoadGeoFile: "; errNumber; errDesc
        MsgBox "Unable to load geobase: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Select a folder for linelist output directory.")
'@EntryPoint
Public Sub clickLinelistDir()
    Dim io As OSFiles
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState

    'Show the folder dialog before entering busy state
    Set io = OSFiles.Create()
    io.LoadFolder

    If Not io.HasValidFolder() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = EventsDesignerCore.EntryManager()
    entry.AddInfo io.Folder(), "lldir"

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickLinelistDir: "; errNumber; errDesc
        MsgBox "Unable to set linelist directory: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Load a template file for linelist creation.")
'@EntryPoint
Public Sub clickLoadTemplate()
    Dim io As OSFiles
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState

    'Show the file dialog before entering busy state
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsb"

    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = EventsDesignerCore.EntryManager()
    entry.AddInfo io.File(), "temppath"
    entry.AddInfo entry.TranslateMessage("MSG_ChemFich"), "edition"

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickLoadTemplate: "; errNumber; errDesc
        MsgBox "Unable to load template: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub


'@section Generation callbacks
'===============================================================================

'@Description("Import setup, prepare specifications, build output linelist workbook, and save.")
'@details
'The press reads the path first, through GenerationHost.InPlace. In
'place (Mac always, Windows with the chkBuildInPlace box checked) the
'build runs in this Excel with the screen off from the press to the
'restore, so the user sees the designer once, at the end, with the report
'in front. On the instance path (Windows, box unchecked) the build runs in
'a hidden Excel and the bar on Main is the only thing that moves here:
'the screen stays on for the whole run and the designer never leaves the
'front. A build that fails keeps its unfinished workbook on disk on both
'paths: the report names the kept file and the message box carries the
'fault. Nothing asks a question.
'@EntryPoint
Public Sub clickGenerate()
    Dim host As GenerationHost

    On Error GoTo Cleanup
    Set host = GenerationHost.Create(ThisWorkbook)

    If host.InPlace Then
        Set host = Nothing
        GenerateInPlace
    Else
        GenerateInInstance host
    End If
    Exit Sub

Cleanup:
    Debug.Print "clickGenerate: "; Err.Number; Err.Description
    MsgBox "Generation failed: " & Err.Description, _
           vbExclamation + vbOKOnly, PROMPT_TITLE
End Sub

'@Description("The single build in this Excel, with the screen off from the press to the restore.")
'@details
'The A3 driver: one busy scope over the whole run, the steps in order
'through GenerateOne, the restore, the report once. A failed build files
'the kept file in the report before the log closes.
Private Sub GenerateInPlace()
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlNorthWestArrow

    Set entry = EventsDesignerCore.EntryManager()

    If Not PrepareRun(entry) Then
        appScope.Restore
        'The report names which entry is not ready, so it is the thing the
        'user has to read here.
        ShowRunLog
        MsgBox entry.TranslateMessage("MSG_NotReady"), _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    entry.AddInfo entry.TranslateMessage("MSG_ReadSetup"), "edition"

    'The whole build: specifications, linelist, sheets, dropdowns, analyses,
    'save. The phase checkings flush to the log after each step.
    GenerateOne entry

    CloseRunAsBuilt entry

    appScope.Restore

    'The report, once, with the screen already back on.
    ShowRunLog
    MsgBox entry.TranslateMessage("MSG_LLCreated"), vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    'The kept file, named in the report before the log closes over the
    'failure. The build has already saved and closed it.
    FileKeptPathWarning
    FinishRunLog "Failed: " & errDesc
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    'A run that failed is the run whose report is worth most: it names the
    'phase that raised and the file the user may open.
    ShowRunLog
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickGenerate: "; errNumber; errDesc
        MsgBox "Generation failed: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("The single build in the hidden instance, with the bar on Main moving here.")
'@details
'This Excel is never busy: events go off and the cursor changes, and the
'screen stays on, so every bar write paints on its own. The copy of the
'designer is written in the scratch folder of the output folder, beside
'the kept file of a failed build; the repository built here to name that
'folder is a throwaway whose drop empties the folder, so the kept name is
'marked on it first. The pre-flight applies the file-name rules once the
'copy is open, since the copy is one more open name in the instance. The
'instance is released on every exit, and in the cleanup the host is
'dropped before anything else runs: a quit instance leaves the process
'table when the last reference to it goes.
'@param host GenerationHost. The host, created and on the instance path.
Private Sub GenerateInInstance(ByVal host As GenerationHost)
    Dim entry As DesignerEntry
    Dim bar As ProgressBar
    Dim scratch As TemporaryRepos
    Dim clashText As String
    Dim previousEvents As Boolean
    Dim previousCursor As Long
    Dim sideHeld As Boolean

    On Error GoTo Cleanup
    Set entry = EventsDesignerCore.EntryManager()

    previousEvents = Application.EnableEvents
    previousCursor = Application.Cursor
    Application.EnableEvents = False
    Application.Cursor = xlWait
    sideHeld = True

    If Not PrepareRun(entry) Then
        RestoreVisibleSide previousEvents, previousCursor
        ShowRunLog
        MsgBox entry.TranslateMessage("MSG_NotReady"), _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    entry.AddInfo entry.TranslateMessage("MSG_ReadSetup"), "edition"
    Set bar = ResolveMainProgressBar()
    LeadBar bar, entry.TranslateMessage("MSG_ReadSetup")

    Set scratch = TemporaryRepos.Create(entry.ValueOf("lldir"))
    scratch.EnsureReady
    scratch.KeepFile KEPT_FILE_NAME

    host.Acquire
    host.OpenDesignerCopy scratch.RootPath

    clashText = CheckBuildFileNames(host, entry, scratch.RootPath & KEPT_FILE_NAME)
    If LenB(clashText) > 0 Then
        entry.AddInfo entry.TranslateMessage("MSG_NotReady"), "edition"
        FinishRunLog entry.TranslateMessage("MSG_NotReady")
        ReleaseBuildHost host
        Set host = Nothing
        If Not bar Is Nothing Then bar.Reset entry.TranslateMessage("MSG_NotReady")
        RestoreVisibleSide previousEvents, previousCursor
        ShowRunLog
        MsgBox entry.TranslateMessage("MSG_NotReady") & vbNewLine & vbNewLine & clashText, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    GenerateOne entry, Nothing, host, bar

    CloseRunAsBuilt entry
    If Not bar Is Nothing Then bar.Complete entry.TranslateMessage("MSG_LLCreated")

    ReleaseBuildHost host
    Set host = Nothing

    RestoreVisibleSide previousEvents, previousCursor

    ShowRunLog
    MsgBox entry.TranslateMessage("MSG_LLCreated"), vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    FileKeptPathWarning
    ReleaseBuildHost host
    Set host = Nothing
    If Not bar Is Nothing Then bar.Reset errDesc
    FinishRunLog "Failed: " & errDesc
    If sideHeld Then RestoreVisibleSide previousEvents, previousCursor
    ShowRunLog
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickGenerate: "; errNumber; errDesc
        MsgBox "Generation failed: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Open the run log and run the entry checks; a failed check closes the log with the not-ready text.")
'@details
'The header carries the setup path and the linelist name the user typed.
'The entry checks are the log's first bundle after the header, and a
'checking with any entry closes the run before anything is built.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@return Boolean. True when the build may start.
Private Function PrepareRun(ByVal entry As DesignerEntry) As Boolean
    StartRunLog entry.ValueOf("setuppath"), entry.ValueOf("llname")

    If ValidateEntries(entry) Then
        PrepareRun = True
        Exit Function
    End If

    entry.AddInfo entry.TranslateMessage("MSG_NotReady"), "edition"
    FinishRunLog entry.TranslateMessage("MSG_NotReady")
End Function

'@Description("Close the run log over a build that saved, and write its text record beside the linelist.")
'@param entry DesignerEntry. The entry manager over the Main worksheet.
Private Sub CloseRunAsBuilt(ByVal entry As DesignerEntry)
    FinishRunLog entry.TranslateMessage("MSG_LLCreated")
    heldLog.ExportText entry.ValueOf("lldir"), entry.ValueOf("llname")
    entry.AddInfo entry.TranslateMessage("MSG_LLCreated"), "edition"
End Sub

'@Description("Put the events and the cursor of this Excel back.")
'@param previousEvents Boolean. The events setting before the run.
'@param previousCursor Long. The cursor before the run.
Private Sub RestoreVisibleSide(ByVal previousEvents As Boolean, ByVal previousCursor As Long)
    On Error Resume Next
    Application.EnableEvents = previousEvents
    Application.Cursor = previousCursor
    On Error GoTo 0
End Sub

'@Description("Apply the file-name rules to the files of a build; a clash is filed in the log and answered.")
'@details
'Excel opens one file of a name at a time in one instance, and refuses a
'save onto a file open anywhere. The setup, the template, the geobase,
'the linelist file and the kept file are checked against the instance
'the build runs in, where the designer copy is one more open name; the
'two files the build writes are checked against this Excel too, for the
'user who opened the last delivered linelist or the last failed build to
'look at it. Every refusal lands in the run log as one entry with the
'error scope, so the report names it.
'@param host GenerationHost. The host, with its copy open.
'@param entry DesignerEntry. The entry manager holding the build's entries.
'@param keptPath String. The full path the failed build would keep its file at.
'@return String. Empty when nothing clashes; otherwise one refusal per line.
Public Function CheckBuildFileNames(ByVal host As GenerationHost, _
                                    ByVal entry As DesignerEntry, _
                                    ByVal keptPath As String) As String
    Dim names As BetterArray
    Dim outputs As BetterArray
    Dim clashText As String
    Dim linelistPath As String
    Dim refusals() As String
    Dim index As Long
    Dim note As Checking

    linelistPath = entry.ValueOf("lldir") & Application.PathSeparator & _
                   entry.ValueOf("llname") & ".xlsb"

    Set names = New BetterArray
    names.LowerBound = 1
    names.Push entry.ValueOf("setuppath")
    names.Push entry.ValueOf("temppath")
    names.Push entry.ValueOf("geopath")
    names.Push linelistPath
    names.Push keptPath

    Set outputs = New BetterArray
    outputs.LowerBound = 1
    outputs.Push linelistPath
    outputs.Push keptPath

    clashText = host.CheckOpenNames(names)
    AppendClash clashText, host.CheckOpenNames(outputs, Application)

    If LenB(clashText) = 0 Then Exit Function

    Set note = Checking.Create(TITLE_PRE_FLIGHT)
    refusals = Split(clashText, vbLf)
    For index = LBound(refusals) To UBound(refusals)
        note.Add "file " & CStr(index + 1), refusals(index), checkingError
    Next index
    CollectIntoLog note

    CheckBuildFileNames = clashText
End Function

'@Description("Add the refusals of one check to the text so far, one per line.")
'@param clashText String. ByRef. The refusals so far.
'@param moreText String. The refusals to add. Empty adds nothing.
Private Sub AppendClash(ByRef clashText As String, ByVal moreText As String)
    If LenB(moreText) = 0 Then Exit Sub
    If LenB(clashText) > 0 Then clashText = clashText & vbLf
    clashText = clashText & moreText
End Sub

'@Description("Release the build instance of a host and file a release that failed in the log.")
'@details
'A Quit that fails leaves the instance behind, and the outcome names its
'window handle; the report is where the user learns that. A host that is
'Nothing, or was never acquired, releases nothing and files nothing.
'@param host GenerationHost. The host to release.
Public Sub ReleaseBuildHost(ByVal host As GenerationHost)
    Dim outcome As String
    Dim note As Checking

    If host Is Nothing Then Exit Sub

    outcome = host.ReleaseInstance()
    If Not IsErrorOutcome(outcome) Then Exit Sub

    Set note = Checking.Create(TITLE_BUILD_INSTANCE)
    note.Add "release", outcome, checkingWarning
    CollectIntoLog note
End Sub

'@Description("Run the entry checks and flush the faults as a log bundle.")
'@details
'Runs DesignerEntry.Validate over the Main entries and takes a checking
'that holds any fault into the run log as one bundle. The log has to be
'opened by the caller. clickGenerate runs this once over the typed
'entries; the multi driver runs it per row after the row's values land
'on Main.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@param faults Checking. Answers what the checks filed, for a caller that
'                shows the fault where the user is looking. The multi
'                driver writes the names into the row's result cell.
'@return Boolean. True when every entry passes and the build may start.
Public Function ValidateEntries(ByVal entry As DesignerEntry, _
                                Optional ByRef faults As Checking = Nothing) As Boolean
    Dim entryChecks As Checking

    Set entryChecks = entry.Validate()
    Set faults = entryChecks

    If entryChecks.Length = 0 Then
        ValidateEntries = True
        Exit Function
    End If

    CollectIntoLog entryChecks
End Function

'@Description("Run one whole build over the Main entries and return the written path.")
'@details
'The driver over the steps of BuildSteps: begin (specifications and
'transfer), linelist, one step per data entry sheet, dropdowns, analyses,
'save. With no host, or a host on the in-place path, the steps are called
'in this instance; with a host on the instance path every step runs in
'the build instance through host.Run, and the entries cross as text with
'the first step, since a workbook cannot. clickGenerate runs it once over
'the entries the user typed; the multi driver writes one T_Multi row onto
'Main and runs it per row. After every step the checkings the step filed
'are pulled into the run log, on both paths, so the report of a build
'that dies still carries the phases that finished. The status target
'takes the milestone texts as plain writes, which is how a multi row's
'result cell reads "sheet 3 of 15" while the row builds; the bar, when
'there is one, is led with the step under way and stepped once after it.
'Both arrive as Nothing when unused.
'
'The caller owns everything around the build: the entry checks
'(ValidateEntries), the log lifecycle (StartRunLog and FinishRunLog), the
'busy state, the host and every dialog. A step that answers an error has
'already kept the unfinished workbook as __temp.xlsb and closed it; the
'outcome is raised to the caller with the kept path in the description,
'and LastKeptPath answers the path on its own. The totals of the build are
'added to the run's counts on both exits.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@param statusTarget Range. One cell taking the milestone texts. Nothing means no writes.
'@param host GenerationHost. The host the steps run through. Nothing runs them here.
'@param bar ProgressBar. The bar over the steps. Nothing means no bar.
'@return String. The full path of the written linelist file.
Public Function GenerateOne(ByVal entry As DesignerEntry, _
                            Optional ByVal statusTarget As Range = Nothing, _
                            Optional ByVal host As GenerationHost = Nothing, _
                            Optional ByVal bar As ProgressBar = Nothing) As String
    Dim designerBook As Workbook
    Dim sheetCount As Long
    Dim counter As Long
    Dim sheetText As String

    keptFilePath = vbNullString

    'The designer is the workbook the entry sits on. A ribbon press builds
    'the entry over ThisWorkbook.Worksheets("Main"), so a designer press
    'reads the same workbook it always did. A caller from outside the
    'designer -- the headless build -- hands over an entry on the Main
    'sheet of the designer copy, and the steps then read that copy.
    Set designerBook = entry.HostSheet.Parent

    If Not bar Is Nothing Then bar.Maximum = FIXED_STEPS

    LeadBar bar, STATUS_TRANSFER
    RunStep StepBegin(host, designerBook, entry), host
    WriteStatus statusTarget, STATUS_TRANSFER
    StepBar bar

    LeadBar bar, STATUS_LINELIST
    RunStep StepOutcome(host, STEP_LINELIST), host
    WriteStatus statusTarget, STATUS_LINELIST
    StepBar bar

    sheetCount = CLng(RunStep(StepOutcome(host, STEP_SHEET_COUNT), host))
    If Not bar Is Nothing Then bar.Maximum = FIXED_STEPS + sheetCount

    For counter = 1 To sheetCount
        'The status leads the sheet it names, so the cell reads the sheet
        'under construction while the build runs.
        sheetText = "sheet " & CStr(counter) & " of " & CStr(sheetCount)
        WriteStatus statusTarget, sheetText
        LeadBar bar, sheetText
        RunStep StepOutcome(host, STEP_SHEET, counter), host
        StepBar bar
    Next counter

    WriteStatus statusTarget, STATUS_DROPDOWNS
    LeadBar bar, STATUS_DROPDOWNS
    RunStep StepOutcome(host, STEP_DROPDOWNS), host
    StepBar bar

    WriteStatus statusTarget, STATUS_ANALYSES
    LeadBar bar, STATUS_ANALYSES
    RunStep StepOutcome(host, STEP_ANALYSES), host
    StepBar bar

    WriteStatus statusTarget, STATUS_SAVE
    LeadBar bar, STATUS_SAVE
    RunStep StepOutcome(host, STEP_SAVE), host
    TakeBuildCounts host
    StepBar bar

    'The path SaveLL wrote, read from the same entries it read
    GenerateOne = entry.ValueOf("lldir") & Application.PathSeparator & _
                  entry.ValueOf("llname") & ".xlsb"
End Function

'@Description("The path of the file the last failed build kept. Empty when nothing was kept.")
Public Function LastKeptPath() As String
    LastKeptPath = keptFilePath
End Function

'@Description("Stop the build under way and keep its file, through the host when there is one.")
'@details
'BuildAbort through the same route the steps took. A step that failed has
'already aborted the build, and the answer is then a bare OK; a build the
'driver stops on its own side, and a build whose instance went away, get
'their kept file named here. The path answered lands in LastKeptPath too.
'@param host GenerationHost. The host the steps ran through. Nothing aborts here.
'@return String. The kept path, or empty when nothing was kept.
Public Function AbortBuild(Optional ByVal host As GenerationHost = Nothing) As String
    Dim kept As String

    kept = KeptPathOf(StepOutcome(host, STEP_ABORT))
    If LenB(kept) > 0 Then keptFilePath = kept

    AbortBuild = kept
End Function

'@Description("Whether a build runs in another instance through its host.")
'@param host GenerationHost. The host, or Nothing.
'@return Boolean. True when the steps have to go through host.Run.
Private Function UsesInstance(ByVal host As GenerationHost) As Boolean
    If host Is Nothing Then Exit Function
    UsesInstance = Not host.InPlace
End Function

'@Description("Run the first step: the designer workbook here, the entries as text in the instance.")
'@param host GenerationHost. The host, or Nothing.
'@param designerBook Workbook. The designer the entries sit on.
'@param entry DesignerEntry. The entry manager, for the entries text.
'@return String. What the step answered.
Private Function StepBegin(ByVal host As GenerationHost, _
                           ByVal designerBook As Workbook, _
                           ByVal entry As DesignerEntry) As String
    If UsesInstance(host) Then
        StepBegin = host.Run(STEP_MODULE_LEAD & STEP_BEGIN_ENTRIES, EntriesTextOf(entry))
    Else
        StepBegin = BuildSteps.BuildBegin(designerBook)
    End If
End Function

'@Description("Run one step by name: here, or through the host.")
'@details
'The in-place calls are spelt out one by one, because a step reached
'through Application.Run in a workbook whose project does not compile
'(the headless designer copy) answers a compile fault, and the steps then
'have to run in the project of the driver.
'@param host GenerationHost. The host, or Nothing.
'@param stepName String. A STEP_* name.
'@param argument Variant. The one argument of the step, when it takes one.
'@return String. What the step answered.
Private Function StepOutcome(ByVal host As GenerationHost, _
                             ByVal stepName As String, _
                             Optional ByVal argument As Variant) As String
    If UsesInstance(host) Then
        If IsMissing(argument) Then
            StepOutcome = host.Run(STEP_MODULE_LEAD & stepName)
        Else
            StepOutcome = host.Run(STEP_MODULE_LEAD & stepName, argument)
        End If
        Exit Function
    End If

    Select Case stepName
        Case STEP_LINELIST
            StepOutcome = BuildSteps.BuildLinelist()
        Case STEP_SHEET_COUNT
            StepOutcome = BuildSteps.BuildSheetCount()
        Case STEP_SHEET
            StepOutcome = BuildSteps.BuildSheet(CLng(argument))
        Case STEP_DROPDOWNS
            StepOutcome = BuildSteps.BuildDropdowns()
        Case STEP_ANALYSES
            StepOutcome = BuildSteps.BuildAnalyses()
        Case STEP_SAVE
            StepOutcome = BuildSteps.BuildSave()
        Case STEP_CHECKINGS
            StepOutcome = BuildSteps.BuildCheckings()
        Case STEP_COUNTS
            StepOutcome = BuildSteps.BuildCounts()
        Case STEP_ABORT
            StepOutcome = BuildSteps.BuildAbort()
        Case Else
            Err.Raise ProjectError.InvalidArgument, "EventsDesignerAdvanced.StepOutcome", _
                      "There is no build step named " & stepName
    End Select
End Function

'@Description("The entries of a build as text, one tag=value per line, for a build in another instance.")
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@return String. The entries text BuildSteps.BuildBeginEntries reads.
Private Function EntriesTextOf(ByVal entry As DesignerEntry) As String
    Dim tags() As String
    Dim index As Long
    Dim textOut As String

    tags = Split(ENTRY_TAGS, ",")
    For index = LBound(tags) To UBound(tags)
        If LenB(textOut) > 0 Then textOut = textOut & ENTRY_SEP
        textOut = textOut & tags(index) & ENTRY_EQUALS & entry.ValueOf(tags(index))
    Next index

    EntriesTextOf = textOut
End Function

'@Description("Take one step outcome: pull its checkings, raise on a failure.")
'@details
'The checkings of the step are pulled first, whatever the step answered,
'since a failed step files the phase it reached before it answers. An
'outcome that failed has already aborted the build; the totals are taken,
'the kept path is read off the outcome and the fault is raised with the
'number the step reported and the rest of the outcome as description.
'@param outcome String. What the step answered.
'@param host GenerationHost. The host the step ran through, or Nothing.
'@return String. The outcome, for the steps that answer a value.
Private Function RunStep(ByVal outcome As String, ByVal host As GenerationHost) As String
    Dim errNumber As Long
    Dim errDesc As String

    PullCheckings host

    If Not IsErrorOutcome(outcome) Then
        RunStep = outcome
        Exit Function
    End If

    TakeBuildCounts host
    keptFilePath = KeptPathOf(outcome)

    errNumber = ErrorNumberOf(outcome)
    errDesc = outcome
    If InStr(outcome, ": ") > 0 Then errDesc = Mid$(outcome, InStr(outcome, ": ") + 2)

    Err.Raise errNumber, "EventsDesignerAdvanced.GenerateOne", errDesc
End Function

'@Description("Pull the checkings queued since the last step into the run log.")
'@details
'A pull that answers an error outcome (the crossing itself failed) is
'dropped: the step outcome is what carries the fault of the build.
'@param host GenerationHost. The host the step ran through, or Nothing.
Private Sub PullCheckings(ByVal host As GenerationHost)
    Dim bundlesText As String

    If heldLog Is Nothing Then Exit Sub

    bundlesText = StepOutcome(host, STEP_CHECKINGS)
    If IsErrorOutcome(bundlesText) Then Exit Sub

    heldLog.CollectText bundlesText
End Sub

'@Description("Whether an outcome carries the error lead.")
'@param outcome String. What a step answered.
'@return Boolean.
Private Function IsErrorOutcome(ByVal outcome As String) As Boolean
    IsErrorOutcome = (Left$(outcome, Len(OUTCOME_ERROR_LEAD)) = OUTCOME_ERROR_LEAD)
End Function

'@Description("The error number a failed outcome carries. A missing number reads as an unexpected state.")
'@param outcome String. An outcome starting with the error lead.
'@return Long. The number.
Private Function ErrorNumberOf(ByVal outcome As String) As Long
    Dim number As Long

    number = CLng(Val(Mid$(outcome, Len(OUTCOME_ERROR_LEAD) + 1)))
    If number = 0 Then number = ProjectError.ErrorUnexpectedState

    ErrorNumberOf = number
End Function

'@Description("The kept path at the end of an outcome, or empty.")
'@param outcome String. What the step answered.
'@return String. The path after the kept mark.
Private Function KeptPathOf(ByVal outcome As String) As String
    Dim markAt As Long

    markAt = InStr(outcome, KEPT_MARK)
    If markAt = 0 Then Exit Function
    KeptPathOf = Trim$(Mid$(outcome, markAt + Len(KEPT_MARK)))
End Function

'@Description("Add the totals of the current build to the run's counts.")
'@details
'Read once per build: on the failure exit before the raise, and after the
'save on the success path.
'@param host GenerationHost. The host the steps ran through, or Nothing.
Private Sub TakeBuildCounts(ByVal host As GenerationHost)
    Dim countsText As String
    Dim parts() As String

    countsText = StepOutcome(host, STEP_COUNTS)
    If IsErrorOutcome(countsText) Then Exit Sub

    parts = Split(countsText, COUNTS_SEP)
    If UBound(parts) < 1 Then Exit Sub

    builtSheets = builtSheets + CLng(Val(parts(0)))
    builtVariables = builtVariables + CLng(Val(parts(1)))
End Sub

'@Description("Write one milestone text into the status target, when there is one.")
'@details
'A plain write. The in-place build runs with the screen off, so the text
'shows when the screen comes back; a multi row's result cell then reads
'the last milestone the row reached before its outcome overwrites it.
'@param statusTarget Range. One cell, or Nothing.
'@param statusText String. The milestone text.
Private Sub WriteStatus(ByVal statusTarget As Range, ByVal statusText As String)
    If statusTarget Is Nothing Then Exit Sub
    statusTarget.Value = statusText
End Sub

'@Description("Show the step under way on the bar, when there is one.")
'@param bar ProgressBar. The bar, or Nothing.
'@param statusText String. The step under way.
Private Sub LeadBar(ByVal bar As ProgressBar, ByVal statusText As String)
    If bar Is Nothing Then Exit Sub
    bar.Update bar.Value, statusText
End Sub

'@Description("Move the bar one step on, when there is one.")
'@param bar ProgressBar. The bar, or Nothing.
Private Sub StepBar(ByVal bar As ProgressBar)
    If bar Is Nothing Then Exit Sub
    bar.StepBy
End Sub

'@Description("Build the bar over the single build when the Main sheet carries its named range.")
'@details
'The bar range and the status cell are owner hand work on the designer.
'A missing name means no bar and no raise; the generation runs the same.
'A name that resolves outside the Main sheet is treated as missing, so the
'bar of the GenerateMultiple sheet keeps its own range.
'@return ProgressBar. The bar, or Nothing when the range is missing.
Private Function ResolveMainProgressBar() As ProgressBar
    Dim mainSheet As Worksheet
    Dim barRange As Range
    Dim statusRange As Range
    Dim bar As ProgressBar

    On Error Resume Next
    Set mainSheet = ThisWorkbook.Worksheets(SHEET_MAIN)
    Set barRange = mainSheet.Range(RNG_PROGRESS_BAR)
    Set statusRange = mainSheet.Range(RNG_PROGRESS_STATUS)
    On Error GoTo 0

    If barRange Is Nothing Then Exit Function
    If Not barRange.Worksheet Is mainSheet Then Exit Function

    'A malformed hand-made range stops the bar alone; the generation runs on
    On Error Resume Next
    Set bar = ProgressBar.Create(barRange, FIXED_STEPS)
    If Not statusRange Is Nothing Then bar.AttachStatusCell statusRange
    On Error GoTo 0

    Set ResolveMainProgressBar = bar
End Function

'@Description("File a warning naming the kept file of the failed build, when there is one.")
Private Sub FileKeptPathWarning()
    Dim aborted As Checking

    If LenB(keptFilePath) = 0 Then Exit Sub

    Set aborted = Checking.Create(TITLE_BUILD_ABORTED)
    aborted.Add "kept file", "The unfinished linelist was kept at " & keptFilePath, _
                checkingWarning
    CollectIntoLog aborted
End Sub


'@section Internal helpers
'===============================================================================

'@Description("Read the language names of a setup Translations sheet.")
'@details
'The one shared language extraction: the Multi group reads a setup's
'languages per row through this routine too. The persisted HiddenNames
'list of the sheet wins; the fallback reads the header row of the first
'ListObject and drops the internal tag columns (__TagInternal__ is
'machinery of the setup table, and it used to land in the dropdown).
'@param tradSheet Worksheet. The Translations worksheet of a setup workbook.
'@return BetterArray. Language names (1-based). Empty when the sheet carries none.
Public Function SetupLanguages(ByVal tradSheet As Worksheet) As BetterArray
    Dim setupStore As HiddenNames
    Dim languagesTag As String
    Dim langString As String
    Dim languages() As String
    Dim langValues As BetterArray
    Dim headerValues As BetterArray
    Dim headerText As String
    Dim lo As ListObject
    Dim idx As Long

    Set langValues = New BetterArray
    langValues.LowerBound = 1
    Set SetupLanguages = langValues

    'The HiddenName key belongs to SetupTranslationsTable, the class that
    'writes the language list on the setup's Translations sheet.
    languagesTag = SetupTranslationsTable.LanguagesNameId

    'Read the persisted language list from the setup's Translations worksheet
    Set setupStore = HiddenNames.Create(tradSheet)

    If setupStore.HasName(languagesTag) Then
        langString = setupStore.ValueAsString(languagesTag)
        If LenB(langString) > 0 Then
            'Split the semicolon-separated string into language names
            languages = Split(langString, ";")
            For idx = LBound(languages) To UBound(languages)
                If LenB(Trim$(languages(idx))) > 0 Then
                    langValues.Push Trim$(languages(idx))
                End If
            Next idx
            If langValues.Length > 0 Then Exit Function
        End If
    End If

    'Fallback: read the header row of the first ListObject on the sheet
    If tradSheet.ListObjects.Count = 0 Then Exit Function

    Set lo = tradSheet.ListObjects(1)
    If lo.HeaderRowRange Is Nothing Then Exit Function

    Set headerValues = New BetterArray
    headerValues.LowerBound = 1
    headerValues.FromExcelRange lo.HeaderRowRange, _
                                DetectLastRow:=False, DetectLastColumn:=False

    'The languages are the header row minus the internal tag columns
    For idx = headerValues.LowerBound To headerValues.UpperBound
        headerText = Trim$(CStr(headerValues.Item(idx)))
        If LenB(headerText) > 0 Then
            If Left$(headerText, Len(INTERNAL_TAG_LEAD)) <> INTERNAL_TAG_LEAD Then
                langValues.Push headerText
            End If
        End If
    Next idx
End Function

'@Description("Update the setup languages dropdown from a setup Translations sheet and auto-select the first language.")
Private Sub ExtractAndUpdateLanguages(ByVal tradSheet As Worksheet, ByVal entry As DesignerEntry)
    Dim langValues As BetterArray
    Dim drop As DropdownLists

    Set langValues = SetupLanguages(tradSheet)
    If langValues.Length = 0 Then Exit Sub

    'Update the setup languages dropdown directly
    Set drop = DropdownLists.Create(ThisWorkbook.Worksheets(SHEET_DROPDOWNS))
    drop.Update langValues, DROP_SETUP_LANGUAGES

    'Auto-select the first setup language (owner decision). The write goes
    'through the entry so the range resolution lives in one place, and
    'Validate is the net under a value the user never touches.
    entry.AddInfo langValues.Item(langValues.LowerBound), "setuplang"
End Sub
