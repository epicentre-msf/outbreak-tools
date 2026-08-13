Attribute VB_Name = "EventsDesignerAdvanced"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Non-core ribbon callbacks for the designer workbook.")
'@depends DesignerPreparation, DesignerEntry, EventsDesignerCore, RibbonDev, LLGeo, ApplicationState, OSFiles, HiddenNames, BetterArray, DropdownLists, LinelistSpecs, Linelist, LLDataEntry, LLSheets, AnalysisOutput, Checking, GenerationLog, InitTransfer, SetupTranslationsTable, ProgressBar
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'Non-core ribbon logics are callbacks whose absence will not fire a
'warning at workbook opening on the designer. They only execute in
'response to explicit user actions (onAction), never at ribbon load
'time (getLabel, getPressed, getVisible).
'
'The DesignerEntry every callback works through is the shared one,
'EventsDesignerCore.EntryManager(). It carries the held translator with it,
'so a press reads no translation table.

Private Const SHEET_GEO As String = "Geo"
Private Const SHEET_MAIN As String = "Main"
Private Const SHEET_DROPDOWNS As String = "__dropdowns"
Private Const PROMPT_TITLE As String = "Designer"

Private Const SHEET_TRANSLATIONS As String = "Translations"

'Dropdown name used by DesignerPreparation for setup languages
Private Const DROP_SETUP_LANGUAGES As String = "__setup_languages"

'Leading shape of the internal tag columns of a setup translations table.
'The fallback header read drops these before the dropdown update.
Private Const INTERNAL_TAG_LEAD As String = "__"

'Progress over the milestones of one build. The bar range is owner hand
'work on the Main sheet of the mock; a missing name means no bar and no
'raise, so this code lands before the hand work and wakes up with it.
'The status cell is the edition cell the entry already writes.
Private Const RNG_PROGRESS_BAR As String = "RNG_ProgressBar"
Private Const RNG_EDITION As String = "RNG_Edition"

'The fixed milestones of one build: entry checks, transfer, linelist
'prepare, the two dropdown flushes, analyses, save. Each data entry
'sheet built adds one step, and Complete is the finalise message.
Private Const FIXED_STEPS As Long = 7

'The log of the current or last generation run. Both drivers open it
'through StartRunLog and every flush goes through CollectIntoLog, so
'one run holds one record. The reference outlives Finish on purpose:
'the record is what a later text re-export reads while the designer
'stays open.
Private heldLog As GenerationLog


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
'@return GenerationLog. The opened log.
Public Function StartRunLog(Optional ByVal setupPath As String = vbNullString, _
                            Optional ByVal linelistName As String = vbNullString) As GenerationLog
    Set heldLog = GenerationLog.Create(ThisWorkbook)
    heldLog.Start setupPath, linelistName
    Set StartRunLog = heldLog
End Function

'@Description("The log of the current or last run. Nothing before the first run.")
Public Function RunLog() As GenerationLog
    Set RunLog = heldLog
End Function

'@Description("Take one bundle into the run log. Without an open run the bundle is dropped.")
'@param checks Checking. The bundle to take.
Public Sub CollectIntoLog(ByVal checks As Checking)
    If heldLog Is Nothing Then Exit Sub
    heldLog.Collect checks
End Sub

'@Description("Close the run log with the outcome text. Without an open run the call is ignored.")
'@param outcome String. The outcome text of the run.
Public Sub FinishRunLog(ByVal outcome As String)
    If heldLog Is Nothing Then Exit Sub
    heldLog.Finish outcome
End Sub


'@section Dev group callbacks
'===============================================================================

'@Description("Initialise the designer workbook: import translations, hide sheets, seed flags.")
'@EntryPoint
Public Sub clickDesignerInitialize(ByRef control As IRibbonControl)
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

'@Description("Clear all geobase data from the Geo worksheet.")
'@EntryPoint
Public Sub clickDelGeo(ByRef control As IRibbonControl)
    Dim geoSheet As Worksheet
    Dim geo As LLGeo
    Dim appScope As ApplicationState

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set geoSheet = ThisWorkbook.Worksheets(SHEET_GEO)
    Set geo = LLGeo.Create(geoSheet)
    geo.Clear

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
        Debug.Print "clickDelGeo: "; errNumber; errDesc
        MsgBox "Unable to clear geobase: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Clear all entry input ranges on the Main sheet.")
'@EntryPoint
Public Sub clickClearEnt(ByRef control As IRibbonControl)
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
'@EntryPoint
Public Sub clickGenerate()
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState
    Dim ll As Linelist
    Dim bar As ProgressBar
    Dim savedPath As String

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlNorthWestArrow

    Set entry = EventsDesignerCore.EntryManager()

    'Open the run log on the designer __check sheet. The header carries
    'the setup path and the linelist name the user typed.
    StartRunLog entry.ValueOf("setuppath"), entry.ValueOf("llname")

    'The entry checks are the log's first bundle after the header. Every
    'fault carries the error scope, so a checking with any entry aborts
    'the run with the report sheet shown.
    If Not ValidateEntries(entry) Then
        entry.AddInfo entry.TranslateMessage("MSG_NotReady"), "edition"
        FinishRunLog entry.TranslateMessage("MSG_NotReady")
        appScope.Restore
        MsgBox entry.TranslateMessage("MSG_NotReady"), _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    entry.AddInfo entry.TranslateMessage("MSG_ReadSetup"), "edition"

    'The bar over the milestones. During the build it owns the edition
    'cell; the first tick is the entry checks that just passed. The
    'maximum starts at the fixed steps and grows by the sheet count once
    'the build knows it.
    Set bar = ResolveMainProgressBar(entry)
    If Not bar Is Nothing Then
        bar.Update 1, entry.TranslateMessage("MSG_ReadSetup"), forceRepaint:=True
    End If

    'The whole build: specifications, linelist, sheets, dropdowns, analyses,
    'save. The phase checkings flush to the log as each phase completes.
    savedPath = GenerateOne(entry, ll, bar)

    'Close the log and write the text record beside the generated linelist
    FinishRunLog entry.TranslateMessage("MSG_LLCreated")
    heldLog.ExportText entry.ValueOf("lldir"), entry.ValueOf("llname")

    entry.AddInfo entry.TranslateMessage("MSG_LLCreated"), "edition"
    If Not bar Is Nothing Then bar.Complete entry.TranslateMessage("MSG_LLCreated")

    appScope.Restore
    MsgBox entry.TranslateMessage("MSG_LLCreated"), vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    'A half-drawn bar never survives the run; the error rides the
    'edition cell through the bar's status write.
    If Not bar Is Nothing Then bar.Reset errDesc
    'Close the log over whatever was written before the error
    FinishRunLog "Failed: " & errDesc
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickGenerate: "; errNumber; errDesc

        'When the linelist object exists, offer the user to view the
        'incomplete workbook or close it; otherwise show a simple error
        If Not ll Is Nothing Then
            ll.ErrorManage errDesc
        Else
            MsgBox "Generation failed: " & errDesc, _
                   vbExclamation + vbOKOnly, PROMPT_TITLE
        End If
    End If
End Sub


'@Description("Run the entry checks and flush the faults as a log bundle.")
'@details
'Runs DesignerEntry.Validate over the Main entries and takes a checking
'that holds any fault into the run log as one bundle. The log has to be
'opened by the caller. clickGenerate runs this once over the typed
'entries; the multi driver runs it per row after the row's values land
'on Main.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@return Boolean. True when every entry passes and the build may start.
Public Function ValidateEntries(ByVal entry As DesignerEntry) As Boolean
    Dim entryChecks As Checking

    Set entryChecks = entry.Validate()

    If entryChecks.Length = 0 Then
        ValidateEntries = True
        Exit Function
    End If

    CollectIntoLog entryChecks
End Function

'@Description("Run one whole build over the Main entries and return the written path.")
'@details
'The single-build core: specifications, linelist workbook, data entry
'sheets, dropdowns, analyses, save. clickGenerate runs it once over the
'entries the user typed; the multi driver writes one T_Multi row onto
'Main and runs it per row. The setup workbook is opened once, inside
'LinelistSpecs.Prepare.
'
'The phase checkings flush to the run log as each phase completes, so
'the report of a build that dies still carries the phases that
'finished. The caller owns everything around the build: the entry
'checks (ValidateEntries), the log lifecycle (StartRunLog and
'FinishRunLog), the busy state and every dialog. A build fault raises
'to the caller; builtLinelist is set as soon as the linelist exists, so
'the caller's handler holds it for ErrorManage or DiscardBuild.
'
'The build ticks at its milestones: transfer, linelist prepare, one
'tick per data entry sheet, the two dropdown flushes, analyses, save.
'The bar hangs off those ticks and repaints itself under the caller's
'busy state; the status target takes the same texts as plain writes,
'which is how a multi row's result cell reads "sheet 3 of 15" while
'the row builds. Both arrive as Nothing when unused.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@param builtLinelist Linelist. Answers the linelist of the build, set before any build step runs.
'@param bar ProgressBar. The bar over the milestones. Nothing means no bar.
'@param statusTarget Range. One cell taking the milestone texts. Nothing means no writes.
'@return String. The full path of the written linelist file.
Public Function GenerateOne(ByVal entry As DesignerEntry, _
                            ByRef builtLinelist As Linelist, _
                            Optional ByVal bar As ProgressBar = Nothing, _
                            Optional ByVal statusTarget As Range = Nothing) As String
    Dim specs As LinelistSpecs
    Dim ll As Linelist
    Dim setupPath As String
    Dim sheetLists As BetterArray
    Dim counter As Long
    Dim anaOut As AnalysisOutput

    setupPath = entry.ValueOf("setuppath")

    'Prepare creates the output workbook and hands it to InitTransfer,
    'which fills it from the setup file and from this designer.
    Set specs = LinelistSpecs.Create(ThisWorkbook)
    specs.Prepare setupPath

    'Flush Phase 1: specification checkings (dictionary, choices, exports,
    'etc.), then the transfer record. A setup whose translations table is
    'missing is filed there: the linelist then keeps the designer's own
    'translation rows. The class never calls the module, so the pull of
    'the transfer record lives here with the driver.
    If Not heldLog Is Nothing Then heldLog.Harvest specs
    If InitTransfer.HasCheckings() Then CollectIntoLog InitTransfer.CheckingValues()

    TickProgress bar, statusTarget, "transfer"

    'After the preparation step of the specifications, internal specifications
    'object shift focus from the designer to the linelist workbook as they
    'are now exported.

    'Build the output linelist workbook (sheets, temp sheets, admin, code transfer)
    Set ll = Linelist.Create(specs)
    Set builtLinelist = ll
    ll.Prepare

    'Flush Phase 1b: code transfer checkings. A component the output workbook
    'already carried was replaced by the designer's copy, and this is where the
    'report names it.
    If ll.HasCheckings Then CollectIntoLog ll.CheckingValues

    TickProgress bar, statusTarget, "linelist"

    'Build data entry worksheets (sections, variables, formatting). The sheet
    'name list is the one Linelist.Prepare already walked the dictionary for.
    Set sheetLists = ll.SheetNames

    'The maximum was provisional until here: the fixed steps plus one
    'step per sheet the loop below will build.
    If Not bar Is Nothing Then bar.Maximum = FIXED_STEPS + sheetLists.Length

    If sheetLists.Length > 0 Then
        Dim listBld As LLDataEntry
        Dim llSheetInfo As LLSheets

        'The shared LLSheets the linelist holds. This loop and TransferAllCode
        'each created their own over the same dictionary, so every row
        'resolution was computed twice.
        Set llSheetInfo = ll.SheetInfoManager

        For counter = sheetLists.LowerBound To sheetLists.UpperBound
            'The tick leads the sheet it names, so the bar reads the
            'sheet under construction while the build runs.
            TickProgress bar, statusTarget, _
                         CStr(sheetLists.Item(counter)), _
                         "sheet " & CStr(counter - sheetLists.LowerBound + 1) & _
                         " of " & CStr(sheetLists.Length) & " - " & _
                         CStr(sheetLists.Item(counter))
            Set listBld = BuildOneSheet(llSheetInfo, ll, sheetLists.Item(counter))

            'Flush Phase 2: the sheet's build checkings, one bundle per
            'sheet, so the report grows with the build.
            If Not listBld Is Nothing Then
                If listBld.HasCheckings Then CollectIntoLog listBld.CheckingValues
            End If
        Next
    End If

    'Flush Phase 2b: shared dropdown checkings, one bundle per store
    Dim dropStd As DropdownLists
    Set dropStd = ll.Dropdown(1)
    If dropStd.HasCheckings Then CollectIntoLog dropStd.CheckingValues
    TickProgress bar, statusTarget, "dropdowns"

    Dim dropCust As DropdownLists
    Set dropCust = ll.Dropdown(2)
    If dropCust.HasCheckings Then CollectIntoLog dropCust.CheckingValues
    TickProgress bar, statusTarget, "dropdowns"

    'Build the analyses
    TickProgress bar, statusTarget, "analyses"
    Set anaOut = AnalysisOutput.Create(specs.AnalysisObject.Wksh(), ll)
    ' All four analysis sheets. The call used to stop after the time series
    ' tables, so the generated linelist carried no time series chart, no
    ' navigation dropdown on that sheet, and two empty sheets where the spatial
    ' and spatio-temporal analyses belong.
    anaOut.WriteAnalysis AnalysisBuildStageAll

    'Flush Phase 3: analysis checkings
    If anaOut.HasCheckings Then CollectIntoLog anaOut.CheckingValues

    'Save the linelist as .xlsb with password protection
    TickProgress bar, statusTarget, "save"
    ll.SaveLL

    'The path SaveLL wrote, read from the same values it read
    GenerateOne = specs.Value("lldir") & Application.PathSeparator & _
                  specs.Value("llname") & ".xlsb"
End Function

'@Description("Move the progress displays one milestone forward.")
'@details
'One tick, two observers: the bar steps with its own repaint, and the
'status target takes the text as a plain write. When the tick has a bar
'the bar's repaint shows the target's write too; a target alone repaints
'here, since the busy state keeps every write invisible without it.
'@param bar ProgressBar. The bar over the milestones. Nothing means no bar.
'@param statusTarget Range. One cell taking the milestone text. Nothing means no write.
'@param statusText String. The milestone text.
'@param targetText String. Text for the status target when it differs from the bar's. Defaults to statusText.
Private Sub TickProgress(ByVal bar As ProgressBar, _
                         ByVal statusTarget As Range, _
                         ByVal statusText As String, _
                         Optional ByVal targetText As String = vbNullString)
    If Not statusTarget Is Nothing Then
        If LenB(targetText) = 0 Then targetText = statusText
        statusTarget.Value = targetText
    End If

    If Not bar Is Nothing Then
        bar.StepBy 1, statusText, forceRepaint:=True
    ElseIf Not statusTarget Is Nothing Then
        Application.ScreenUpdating = True
        DoEvents
        Application.ScreenUpdating = False
    End If
End Sub

'@Description("Build the bar over the milestones when the Main sheet carries its range.")
'@details
'The bar range is owner hand work on the mock. A missing name means no
'bar and no raise; the generation runs the same. A name that resolves
'outside the Main sheet is treated as missing, so the multi bar on the
'GenerateMultiple sheet keeps its own range. The edition cell rides
'along as the status cell, which is how the milestone texts land where
'the entry writes its start and end messages.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@return ProgressBar. The bar, or Nothing when the range is missing.
Private Function ResolveMainProgressBar(ByVal entry As DesignerEntry) As ProgressBar
    Dim mainSheet As Worksheet
    Dim barRange As Range
    Dim statusRange As Range
    Dim bar As ProgressBar

    Set mainSheet = entry.HostSheet

    On Error Resume Next
    Set barRange = mainSheet.Range(RNG_PROGRESS_BAR)
    Set statusRange = mainSheet.Range(RNG_EDITION)
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


'@section Internal helpers
'===============================================================================

'@Description("Build a data entry worksheet from the dictionary and return the builder.")
Private Function BuildOneSheet(ByVal llshs As LLSheets, ByVal ll As Linelist, ByVal sheetName As String) As LLDataEntry
    Dim sheetType As String
    Dim layer As Byte
    Dim listBld As LLDataEntry

    sheetType = llshs.SheetInfo(sheetName)

    If sheetType = "vlist1D" Then
        layer = LLDataEntryLayerVList
    ElseIf sheetType = "hlist2D" Then
        layer = LLDataEntryLayerHList
    Else
        Exit Function
    End If

    'The builder takes the LLSheets this loop already holds. It used to build
    'its own, and so did each of the three members inside it, so one sheet cost
    'five searches of the dictionary for the same row.
    Set listBld = LLDataEntry.Create(layer, sheetName, ll, llshs)
    listBld.Build

    Set BuildOneSheet = listBld
End Function

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
