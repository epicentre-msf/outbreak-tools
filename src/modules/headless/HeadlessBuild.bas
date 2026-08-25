Attribute VB_Name = "HeadlessBuild"
Attribute VB_Description = "Drive a setup fill and a linelist build with no designer and no dialog"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("Headless")
'@ModuleDescription("Drive a setup fill and a linelist build with no designer and no dialog")

Option Explicit

'@description
'THE HEADLESS WORKFLOW, STEP ONE: FILLING A SETUP.
'
'Two things this project can only do by hand today are a setup fill and a
'linelist build. Both run inside Excel, both are driven from a ribbon, and
'both need a designer workbook whose baked code is whatever was last imported
'into it. This module drives them from code instead, so a run proves the
'source in src/ rather than a binary nobody can diff.
'
'IT IS WRITTEN TO BE CALLED FROM OUTSIDE
'-------------------------------------------------------------------------------
'Every routine takes its paths as parameters, answers a STRING outcome, and
'raises nothing. That is deliberate: the caller is a test module today and an
'R package later, and neither can read a dialog. An outcome reads "OK" or
'"ERROR <number>: <description>", and a caller decides what to do with it.
'
'WHY THE SETUP IMPORT IS INJECTED RATHER THAN CALLED
'-------------------------------------------------------------------------------
'SetupImport resolves its destination as ThisWorkbook, in seven places. Called
'from here it would import the source INTO THE DRIVER WORKBOOK. So the caller
'of SetupImport has to live inside the workbook being filled, and
'ImportSetupFromWorkbook puts it there: it reads
'scripts/headless/vba/OBTSetupImportHeadless.bas off disk, imports it into the
'target's VBProject, runs its entry point through Application.Run, and removes
'it again. The injection needs "Trust access to the VBA project object model",
'which the headless test runner already requires for its own imports.
'
'THE HEADLESS WORKFLOW, STEP TWO: BUILDING A LINELIST.
'-------------------------------------------------------------------------------
'BuildLinelistFromSetup runs the generation with no designer machinery: no
'ribbon press, no DesignerEntry, no dialog and no progress bar. It is the body
'of EventsDesignerAdvanced.GenerateOne, driven from here, and it differs from
'that routine in ONE way that matters: GenerateOne builds its specifications
'over ThisWorkbook, so it can only ever generate from the workbook it is
'running inside. This one takes the designer-shaped workbook as a parameter,
'which is what lets a test module -- or an R session -- drive a build from
'outside.
'
'WHY A DESIGNER FILE IS STILL COPIED
'-------------------------------------------------------------------------------
'"No designer" means no designer MACHINERY, not no designer WORKSHEETS. The
'build reads eight of them: Main carries the entries, __formatter the design,
'__pass the passwords, __formula the tokens, LinelistTranslation the five
'translation tables, Geo the geobase, DesignerTranslation the labels and
'__check the run log. Those are data, and no sensible amount of code builds
'them from nothing.
'
'So a designer file is copied and its CODE is thrown away: every class, module
'and form is re-imported from src/ and from the merged forms folder before the
'transfer runs. The copy is what the build reads and the copy is what
'CodeTransfer exports from, so the linelist ships the source in the repository
'rather than whatever was last pasted into a binary. The copy is saved beside
'the linelist, because it is the evidence of what the run actually carried.
'
'The copy's VBProject never has to COMPILE. CodeTransfer exports components
'from it, and an export reads text. That is why importing all of src/ into a
'workbook holding a stale designer is safe here and would not be safe in a
'workbook that had to run.
'THE HEADLESS WORKFLOW, STEP THREE: BUILDING SEVERAL AT ONCE.
'-------------------------------------------------------------------------------
'BuildMultipleFromTable runs the designer's OWN multiple generation loop,
'EventsDesignerMulti.GenerateMultipleRows, over the T_Multi table of a designer
'copy. The loop is the one a ribbon press runs, so a run here measures the code
'a user presses instead of a second copy of it written for the harness.
'@depends BetterArray, Checking, ApplicationState, LinelistSpecs, Linelist
'@depends LLDataEntry, LLSheets, AnalysisOutput, DropdownLists, GenerationLog
'@depends InitTransfer, EventsDesignerAdvanced, EventsDesignerMulti
'@depends DesignerEntry, DesignerPreparation

'The module injected into the target setup, and the entry point it carries.
Private Const INJECTED_MODULE As String = "OBTSetupImportHeadless"
Private Const INJECTED_ENTRY As String = "OBTHeadlessImportSetup"

Private Const OUTCOME_OK As String = "OK"

'What Err.Raise reports as the source. The house shape: every module carries
'its own ThrowError, because the helper is Private in each of them.
Private Const MODULE_NAME As String = "HeadlessBuild"

'The designer worksheet the entries are written on, and the three entries a
'build cannot be pointed anywhere without.
Private Const SHEET_MAIN As String = "Main"
Private Const SHEET_SETUP_TRANSLATIONS As String = "Translations"
Private Const RNG_LL_DIR As String = "RNG_LLDir"
Private Const RNG_LL_NAME As String = "RNG_LLName"
Private Const RNG_LL_TEMPLATE As String = "RNG_LLTemp"

'The rest of the entries, each written when the option carrying it is given.
Private Const RNG_PATH_DICO As String = "RNG_PathDico"
Private Const RNG_PATH_GEO As String = "RNG_PathGeo"
Private Const RNG_LANG_SETUP As String = "RNG_LangSetup"
Private Const RNG_LL_FORM As String = "RNG_LLForm"
Private Const RNG_LL_PWD_OPEN As String = "RNG_LLPwdOpen"

'The two Main entries a multi row falls back to when it names neither.
'Both are read off the copied designer, so a run says nothing about a
'design or an epiweek and still builds what the designer would have built.
Private Const RNG_DESIGN_LL As String = "RNG_DesignLL"
Private Const RNG_EPIWEEK As String = "RNG_DefaultEpiWeek"

'The format worksheet, and the one name on it that says which design of the
'format table a run is written in. DESIGNTYPE DECIDES THE DESIGN:
'LinelistSpecs.EnsureFormat and InitTransfer.DesignerDesignName both read it.
'The Main entry beside it is the dropdown a person fills, and the designer's
'own event handlers are what copy that cell into DESIGNTYPE. A headless run
'throws those handlers away, so this module writes both names itself.
Private Const SHEET_FORMATTER As String = "__formatter"
Private Const RANGE_DESIGN_TYPE As String = "DESIGNTYPE"

'The multiple generation table, and the twelve columns the driver reads.
'The names are repeated from EventsDesignerMulti because the constants are
'Private there. They are the headers of a shipped worksheet, so a change to
'one is a change to the designer binaries as well.
Private Const SHEET_GENERATE_MULTIPLE As String = "GenerateMultiple"
Private Const TABLE_MULTI As String = "T_Multi"
Private Const COL_MULTI_SETUPS As String = "setups"
Private Const COL_MULTI_GEOBASES As String = "geobases"
Private Const COL_MULTI_OUTPUT_FOLDERS As String = "output folders"
Private Const COL_MULTI_OUTPUT_FILES As String = "output files"
Private Const COL_MULTI_PASSWORD As String = "output file password"
Private Const COL_MULTI_DEBUG_PASSWORD As String = "output file debugging password"
Private Const COL_MULTI_LANG_DICTIONARY As String = "language of the dictionary"
Private Const COL_MULTI_LANG_INTERFACE As String = "language of the interface"
Private Const COL_MULTI_EPIWEEK_START As String = "epiweek start"
Private Const COL_MULTI_DESIGN As String = "design"
Private Const COL_MULTI_RESULT As String = "result"

'What the multi driver writes into the result column of a row that built.
Private Const RESULT_BUILT As String = "OK"

'What separates the rows of a multi run in the spec string. The fields
'inside one row are separated by the same pipe the options use, so a
'second separator is needed for the rows themselves. A path can hold a
'space and a dash; two tildes together belong to nothing on either host.
Private Const ROW_SEPARATOR As String = "~~"

'The source folders holding the components Linelist.TransferAllCode moves, and
'nothing else. Nine of the folders under src/ carry all of them; msetup,
'mastersetup, setup, designer, dev, rubberduck, stale and formulas carry none.
'
'`sections` is on the list for ONE class. SectionMap is the only thing the
'transfer takes out of it, and EventsLinelistButtons types SectionMap, so a
'linelist built without the folder loses its project compile.
'
'`linelistform` is NOT on the list, and that is deliberate. Its ten FormLogic
'modules are the code BEHIND the forms: every one of them uses Me and declares
'control event handlers, neither of which compiles in a standard module.
'merge-form-code.R writes each module into the code module of the form it
'belongs to, so the .frm files imported below already carry that code where it
'is legal. Importing the .bas files as well put four of them in the delivered
'linelist as standard modules, and that alone cost the file its compile.
'
'The narrow list is not an optimisation. Excel for Mac is SANDBOXED: reading a
'folder it has no security-scoped grant for pops a dialog, and a dialog in a
'headless run is a hang. Walking the whole tree asked for the mastersetup
'disease classes -- which no linelist has ever carried -- one prompt at a time.
'
'A folder that goes stale here is LOUD rather than silent, and that is what
'StripProject buys: the copy's own components are removed before any import, so
'a component the transfer asks for that no folder here supplied is ABSENT, and
'CodeTransfer raises ElementNotFound naming it. Without the strip, the copy's
'stale component of that name would be exported and nothing would be said.
Private Const CLASSES_FOLDER As String = "src/classes"
Private Const MODULES_FOLDER As String = "src/modules"
Private Const TRANSFER_CLASS_FOLDERS As String = _
    "analyses|dataio|dictionary|general|geo|graphs|linelist|sections|showhide"
Private Const TRANSFER_MODULE_FOLDERS As String = "linelist"

'What VBComponents calls a worksheet or workbook component. Naming the VBIDE
'constant would put the library in the compile, and an identifier from an
'unreferenced library costs the WHOLE project its compile.
Private Const COMPONENT_DOCUMENT As Long = 100

'The scratch folder TemporaryRepos makes under the output folder, and the one
'CodeTransfer writes every exported component into. The name is repeated from
'TemporaryRepos.DEFAULT_FOLDER_NAME because that constant is Private there, and
'this module needs it BEFORE any class has run to ask for the sandbox grant.
Private Const SCRATCH_FOLDER As String = "OBTApp_"

'The designer worksheet holding the five translation tables, and the one whose
'column headers name the language codes a linelist can be written in.
Private Const SHEET_LL_TRANSLATION As String = "LinelistTranslation"
Private Const TABLE_LL_MESSAGES As String = "T_TradLLMsg"

'What the run last did, read back through the properties below. A caller that
'has the outcome string still has no way to say how much was built, and "OK"
'over an empty linelist is exactly the answer a test has to be able to refuse.
Private runNarrative As String
Private lastSheets As Long
Private lastVariables As Long
Private lastComponents As Long
Private lastLinelist As String
Private lastLog As String

'How many rows of a multi run built and how many failed. A single build
'leaves both at zero, and the summary carries them on every run so one
'reader serves both entry points.
Private lastBuilt As Long
Private lastFailed As Long

'What the last grant call did, when it did something worth saying. Read back
'through LastAccessNote so a report can tell a broken call from a refusal.
'
'The name differs from that Function on purpose. VBA ignores case, so a private
'accessNote and a public AccessNote are ONE name and the project stops
'compiling with "Ambiguous name detected". Every pair in this module is spelt
'apart for that reason.
Private accessNote As String

'Where the narrative is mirrored, line by line, as the run goes. The in-memory
'narrative above is read back at the END of a run, and a run that takes Excel
'down with it never gets there: the caller reads a dead connection and has no
'idea which phase was on screen. This file is opened, written and closed on
'every line, so whatever reached it survives anything.
Private tracePath As String

'The options of the run being prepared, read out of the options string.
Private optTemplate As String
Private optGeo As String
Private optSetupLang As String
Private optLinelistLang As String
Private optPassword As String
Private optDesign As String
Private optStyle As String


'@section File access
'===============================================================================

'@Description("Ask Excel for persistent access to every path of a run, in one dialog.")
'@details
'Excel for Mac is sandboxed: VBA file reads and writes on a path the sandbox
'does not know pop ONE DIALOG PER FILE, and a headless run over 84 source
'files is 84 dialogs. GrantAccessToMultipleFiles is the API Microsoft ships
'for exactly this: called BEFORE any file work, it shows one consolidated
'dialog for whatever is not yet granted, and the grant PERSISTS across
'sessions -- an already-granted path asks nothing at all. A folder path
'grants its whole tree, so the callers here pass roots rather than files.
'
'The call is late-bound through Object because the member exists only in the
'Mac type library, and one identifier from a missing library costs the whole
'project its compile. On Windows there is no sandbox and the silent failure
'is the right answer.
'@param paths Variant. An array of POSIX paths, folders or files.
'@return Boolean. True when every path is granted; False on anything else.
Public Function EnsureFileAccess(ByVal paths As Variant) As Boolean
    Dim host As Object

    accessNote = vbNullString

    On Error Resume Next
        Set host = Application
        EnsureFileAccess = host.GrantAccessToMultipleFiles(paths)

        'The error is KEPT, and that is the point of this routine now.
        '
        'It used to Err.Clear here and answer False, and False was reported as
        '"not confirmed" - English that reads as "the operator said no". On
        '2026-08-14 a run showed the call raising 438, Object doesn't support
        'this property or method: the member is absent on that Excel and the
        'project had never made a single grant. Every run for months printed
        '"not confirmed" and meant "this call is broken".
        'A swallowed error reported as an ambiguous phrase hides a dead call.
        If Err.Number <> 0 Then
            accessNote = "the call FAILED, error " & CStr(Err.Number) & _
                             ": " & Err.Description
            Err.Clear
        End If
    On Error GoTo 0
End Function

'@Description("Word the grant result for a report, keeping a broken call visible.")
'@details
'Three outcomes, and they used to be two. A call that RAISED and a call that
'answered False both printed "not confirmed", so the reader could not tell a
'missing API from an operator saying no.
'@param granted Boolean. What EnsureFileAccess answered.
'@return String. The phrase for the report line.
Private Function AccessOutcome(ByVal granted As Boolean) As String
    If granted Then
        AccessOutcome = "granted"
    ElseIf LenB(LastAccessNote()) > 0 Then
        AccessOutcome = LastAccessNote() & " - NO GRANT WAS MADE, expect prompts"
    Else
        AccessOutcome = "the call ran and answered no - expect prompts"
    End If
End Function

'@Description("The platform this build ran on, on one tag.")
'@details
'A build report read on another machine, or months later, says nothing about
'where it was produced, and the two platforms do not behave the same: a geobase
'export Windows accepts has refused on a Mac. The name and the bitness come from
'the compile constants, so they are the build that is running and cannot be
'misread. The Excel version comes from the application.
'@return String. Something of the shape "mac-64 excel-16.90".
Public Function PlatformTag() As String
    Dim osName As String
    Dim bits As String

    #If Mac Then
        osName = "mac"
    #Else
        osName = "win"
    #End If

    #If Win64 Then
        bits = "64"
    #Else
        bits = "32"
    #End If

    PlatformTag = osName & "-" & bits & " excel-" & Application.Version
End Function

'@Description("Say what the last EnsureFileAccess call actually did.")
'@details
'Empty when the call ran. A caller writes this into its report so a broken
'grant reads as broken instead of reading as a refusal.
'@return String. The failure, or empty.
Public Function LastAccessNote() As String
    LastAccessNote = accessNote
End Function


'@section Filling a setup
'===============================================================================

'@Description("Fill one setup workbook from another, headless.")
'@details
'Opens the target, injects the import module, runs it against the source, and
'saves. The target is opened read/write and SAVED, so the caller passes a COPY
'unless it means to change the original.
'@param targetPath String. The setup workbook to fill, as a full path.
'@param sourcePath String. The setup workbook to read from.
'@param modulePath String. Full path of OBTSetupImportHeadless.bas.
'@return String. OUTCOME_OK, or "ERROR <number>: <description>".
Public Function ImportSetupFromWorkbook(ByVal targetPath As String, _
                                        ByVal sourcePath As String, _
                                        ByVal modulePath As String) As String
    Dim target As Workbook
    Dim outcome As String
    Dim prevAlerts As Boolean

    On Error GoTo Failed

    'Before the first Dir$: an ungranted path prompts from there on.
    EnsureFileAccess Array(targetPath, sourcePath, modulePath)

    If Len(Dir$(targetPath)) = 0 Then
        ImportSetupFromWorkbook = "ERROR 0: no setup to fill at " & targetPath
        Exit Function
    End If

    If Len(Dir$(modulePath)) = 0 Then
        ImportSetupFromWorkbook = "ERROR 0: the injected module is missing - " & modulePath
        Exit Function
    End If

    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Set target = Application.Workbooks.Open(fileName:=targetPath, ReadOnly:=False)

    InjectModule target, modulePath
    outcome = CStr(Application.Run("'" & target.Name & "'!" & INJECTED_ENTRY, sourcePath))

    'The module goes back out before the save, so the filled setup carries no
    'trace of the machinery that filled it.
    RemoveModule target, INJECTED_MODULE

    target.Save
    target.Close SaveChanges:=False
    Set target = Nothing

    Application.DisplayAlerts = prevAlerts
    ImportSetupFromWorkbook = outcome
    Exit Function

Failed:
    ImportSetupFromWorkbook = "ERROR " & CStr(Err.Number) & ": " & Err.Description

    On Error Resume Next
        If Not target Is Nothing Then target.Close SaveChanges:=False
        Application.DisplayAlerts = prevAlerts
    On Error GoTo 0
End Function


'@section Building a linelist
'===============================================================================

'@Description("Build one linelist from a filled setup, with no designer machinery.")
'@details
'The whole run, in order: copy the designer file, re-import every component
'from source into the copy, write the entries onto its Main worksheet, and run
'the generation. The linelist lands at
'<outputFolder>/<outputName>.xlsb, the run log beside it as
'<outputName>-generation.txt, and the designer the run actually used as
'<outputName>-designer.xlsb.
'
'THE OPTIONS STRING
'-------------------------------------------------------------------------------
'Pipe-separated key=value pairs, every one of them optional:
'
'  temppath=<path>   the ribbon template. EMPTY MEANS THE BUTTONS BUILD -- the
'                    linelist then carries action buttons on its sheets and an
'                    Admin sheet, and no ribbon. This is the switch, and it is
'                    the one option worth being deliberate about.
'  geopath=<path>    the geobase to import. Empty is a valid choice and the
'                    common one: most setups carry no geography.
'  setuplang=<name>  the language of the setup file
'  lllang=<name>     the language the linelist is written in
'  llpassword=<text> the open password of the saved file. Empty saves a file
'                    that opens on a double-click, which is what a demo wants.
'  stylepath=<path>  a linelist style workbook to import onto the designer's
'                    format sheet before the build, the file the ribbon's
'                    "import styles" button takes. GIVEN, THE DESIGNER'S OWN
'                    FORMATTER IS SHIPPED; empty, the SETUP's formatter is,
'                    which is the design a setup carries of its own.
'  design=<name>     the column of the format table the run is written in --
'                    "design 1", "design 2", "user defined". Empty keeps
'                    whatever the designer's DESIGNTYPE already names. A name
'                    no column matches falls back to the format table's own
'                    default, the same answer a designer gives.
'
'A key this routine does not know is reported rather than ignored, because a
'misspelled key that reads as "not given" is a build silently pointed
'somewhere else.
'@param designerPath String. The designer workbook to copy, as a full path.
'@param setupPath String. The filled setup to generate from.
'@param sourceRoot String. The repository root holding src/classes and src/modules.
'@param formsFolder String. The folder of merged .frm files.
'@param outputFolder String. Where the three files land.
'@param outputName String. The base name of the linelist.
'@param options String. Pipe-separated key=value pairs. See above.
'@param grantRoot String. The ONE folder every path above sits under, for the
'                 sandbox grant. Empty falls back to granting the five paths
'                 separately, which is what a caller outside the headless
'                 launcher (a test module) still does.
'@return String. OUTCOME_OK, or "ERROR <number>: <description>".
Public Function BuildLinelistFromSetup(ByVal designerPath As String, _
                                       ByVal setupPath As String, _
                                       ByVal sourceRoot As String, _
                                       ByVal formsFolder As String, _
                                       ByVal outputFolder As String, _
                                       ByVal outputName As String, _
                                       Optional ByVal options As String = vbNullString, _
                                       Optional ByVal grantRoot As String = vbNullString) As String
    Dim designerBook As Workbook
    Dim workingPath As String
    Dim appScope As ApplicationState
    Dim missingPath As String
    Dim accessGranted As Boolean

    ResetRunState

    'The scratch folder is created here rather than left to TemporaryRepos,
    'which does not make it until the build is well past the last moment a
    'dialog could be answered. CodeTransfer exports every component to
    '<lldir>/OBTApp_ and imports it back, so a run touches two files per
    'component in there.
    EnsureFolder JoinPath(outputFolder, SCRATCH_FOLDER)

    'The trace, started fresh here so a run is never read against the lines of
    'the one before it. Everything AddToReport records from now on is on disk a
    'line at a time.
    tracePath = JoinPath(outputFolder, outputName & "-trace.txt")
    DeleteFile tracePath
    AddToReport "build started, output " & outputFolder & " as " & outputName

    'ONE grant when the launcher gave a root, because a folder grant persists
    'across Excel sessions and covers files created inside it afterwards --
    'proven by probe, written up in .obt/gotchas/macos-sandbox-grant.md. The
    'headless launcher stages designer, setup, source, forms and output under a
    'single root precisely so this is one path.
    '
    'The five-path fallback is for a caller that has no such root, which today
    'means a test module driving this directly. It grants what it can name, and
    'the files each run creates fresh underneath are what still prompt.
    If LenB(Trim$(grantRoot)) > 0 Then
        accessGranted = EnsureFileAccess(Array(grantRoot))
        AddToReport "file access: one root, " & AccessOutcome(accessGranted) & _
                    " - " & grantRoot
    Else
        accessGranted = EnsureFileAccess(Array(sourceRoot, formsFolder, outputFolder, _
                                               designerPath, setupPath))
        AddToReport "file access: five paths, " & AccessOutcome(accessGranted) & _
                    " - no grant root was given"
    End If

    missingPath = FirstMissingPath(designerPath, setupPath, sourceRoot, formsFolder)
    If LenB(missingPath) > 0 Then
        BuildLinelistFromSetup = "ERROR 0: " & missingPath
        Exit Function
    End If

    If LenB(Trim$(outputName)) = 0 Then
        BuildLinelistFromSetup = "ERROR 0: the linelist needs a name"
        Exit Function
    End If

    On Error GoTo Failed

    ReadOptions options

    'The designer is copied before anything is opened, so the file the project
    'ships is never the file this run rewrites.
    workingPath = JoinPath(outputFolder, outputName & "-designer.xlsb")
    DeleteFile workingPath
    FileCopy designerPath, workingPath
    AddToReport "designer copied to " & workingPath

    'Events off before the open: a designer carries a Workbook_Open handler and
    'this run has no screen for whatever it would put there.
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    'The setup language, before any entry is written. Loading a setup by hand
    'fills it on Main; the headless path has no load step, so it is read off the
    'setup itself, the way clickLoadFileDic reads it. An empty language is a
    'translation column nothing can resolve.
    ResolveLanguages setupPath
    AddToReport "languages resolved, setup " & optSetupLang

    Set designerBook = Application.Workbooks.Open(fileName:=workingPath, ReadOnly:=False)
    AddToReport "designer copy opened"

    RefreshSourceCode designerBook, sourceRoot, formsFolder

    'The linelist language needs the designer open: it is settled against the
    'translation table's own column headers rather than against the setup.
    ResolveInterfaceLanguage designerBook
    AddToReport "interface language resolved: " & optLinelistLang

    PrepareDesignerGeo designerBook
    AddToReport "designer geo prepared"

    'The style first, the design second. Importing a style workbook writes that
    'workbook's own DESIGNTYPE onto the format sheet, so a design named by the
    'caller has to be applied after the import or the file would overrule it.
    ImportLinelistStyle designerBook
    ApplyDesignChoice designerBook

    WriteDesignerEntries designerBook, setupPath, outputFolder, outputName
    AddToReport "designer entries written"

    RunGeneration designerBook, setupPath, outputFolder, outputName
    AddToReport "generation finished, saving the designer copy"

    'The designer is saved with the source it carried, so a run that produced a
    'surprising linelist can be read back rather than guessed at.
    designerBook.Save
    designerBook.Close SaveChanges:=False
    Set designerBook = Nothing

    appScope.Restore
    BuildLinelistFromSetup = OUTCOME_OK
    Exit Function

Failed:
    BuildLinelistFromSetup = "ERROR " & CStr(Err.Number) & " (" & Err.Source & "): " & _
                             Err.Description
    'The NUMBER goes in the narrative too. A description does not always survive
    'a class boundary and the number does, so a line carrying only the
    'description can read "failed: ()" and name nothing at all.
    AddToReport "failed: " & CStr(Err.Number) & " (" & Err.Source & ") " & _
                Err.Description

    'The designer copy is saved even now: its __check sheet and Main entries
    'are the record of how far the run got, and a copy closed unsaved answers
    'nothing.
    On Error Resume Next
        If Not designerBook Is Nothing Then
            designerBook.Save
            designerBook.Close SaveChanges:=False
        End If

        'The linelist under construction. LinelistSpecs.Prepare saves the
        'template copy as __temp.xlsb and closes it itself when IT fails; a
        'failure in any phase after Prepare leaves that workbook open and
        'unsaved, and the next save or quit then hangs headless on the
        'save-changes prompt. Found by walking the collection rather than by
        'indexing on the name: the index raises 9 when the workbook is not
        'there, and under Break-on-All-Errors trapping that raise stops the
        'run with a dialog nobody is there to answer.
        Dim orphanBook As Workbook
        Dim openBook As Workbook
        For Each openBook In Application.Workbooks
            If openBook.Name = "__temp.xlsb" Then
                Set orphanBook = openBook
                Exit For
            End If
        Next
        If Not orphanBook Is Nothing Then orphanBook.Close SaveChanges:=False

        If Not appScope Is Nothing Then appScope.Restore
        Err.Clear
    On Error GoTo 0
End Function


'@section The multiple generation, headless
'===============================================================================

'@Description("Build one linelist per row of the designer's T_Multi table, with no dialog.")
'@details
'THE MULTI DRIVER ITSELF RUNS HERE, and that is the whole point of this
'entry point. EventsDesignerMulti.GenerateMultipleRows is the loop a
'designer press runs: it walks T_Multi, writes each row onto Main through
'the shared DesignerEntry, runs the entry checks, calls the single-build
'core, writes the row's outcome into the result column, and keeps going
'when a row fails. This hands that same loop the T_Multi and the Main of
'a DESIGNER COPY, so the run measures the code a user presses.
'
'Two things made the loop reachable from outside the designer.
'GenerateOne reads its designer off the entry's own host sheet, so an
'entry built over the copy's Main builds the copy's specifications.
'StartRunLog takes the designer that carries the __check sheet, because
'the driver workbook this code runs inside carries none.
'
'ONE LOG FOR THE RUN, ONE REPORT FILE
'-------------------------------------------------------------------------------
'Every row flushes its own bundles into the one run log, headed by the
'row ID, which is what a designer multi press produces. The text file is
'written once, as <outputName>-generation.txt, and it holds every row.
'
'WHAT EACH ROW OWNS AND WHAT THE RUN OWNS
'-------------------------------------------------------------------------------
'A row carries its setup, its geobase, its output name, its two
'passwords, its two languages, its epiweek and its design. The run
'carries the designer, the source, the forms, the output folder and the
'ribbon template: T_Multi has no template column, so the template entry
'is written on Main once and every row builds with it.
'@param designerPath String. The designer workbook to copy, as a full path.
'@param rowsSpec String. The rows to build. See FillMultiTable for the shape.
'@param sourceRoot String. The repository root holding src/classes and src/modules.
'@param formsFolder String. The folder of merged .frm files.
'@param outputFolder String. Where every linelist and the run report land.
'@param outputName String. The name of the RUN: the designer copy, the trace and the report.
'@param options String. Pipe-separated key=value pairs, the same keys the single build reads.
'@param grantRoot String. The ONE folder every path above sits under, for the sandbox grant.
'@return String. OUTCOME_OK, or "ERROR <number>: <description>".
Public Function BuildMultipleFromTable(ByVal designerPath As String, _
                                       ByVal rowsSpec As String, _
                                       ByVal sourceRoot As String, _
                                       ByVal formsFolder As String, _
                                       ByVal outputFolder As String, _
                                       ByVal outputName As String, _
                                       Optional ByVal options As String = vbNullString, _
                                       Optional ByVal grantRoot As String = vbNullString) As String
    Dim designerBook As Workbook
    Dim workingPath As String
    Dim appScope As ApplicationState
    Dim missingPath As String
    Dim accessGranted As Boolean
    Dim multiTable As ListObject
    Dim entry As DesignerEntry
    Dim rowCount As Long

    ResetRunState

    EnsureFolder JoinPath(outputFolder, SCRATCH_FOLDER)

    tracePath = JoinPath(outputFolder, outputName & "-trace.txt")
    DeleteFile tracePath
    AddToReport "multi build started, output " & outputFolder & " as " & outputName

    If LenB(Trim$(grantRoot)) > 0 Then
        accessGranted = EnsureFileAccess(Array(grantRoot))
        AddToReport "file access: one root, " & AccessOutcome(accessGranted) & _
                    " - " & grantRoot
    Else
        accessGranted = EnsureFileAccess(Array(sourceRoot, formsFolder, outputFolder, _
                                               designerPath))
        AddToReport "file access: four paths, " & AccessOutcome(accessGranted) & _
                    " - no grant root was given"
    End If

    missingPath = FirstMissingMultiPath(designerPath, rowsSpec, sourceRoot, formsFolder)
    If LenB(missingPath) > 0 Then
        BuildMultipleFromTable = "ERROR 0: " & missingPath
        Exit Function
    End If

    If LenB(Trim$(outputName)) = 0 Then
        BuildMultipleFromTable = "ERROR 0: the run needs a name"
        Exit Function
    End If

    On Error GoTo Failed

    ReadOptions options

    workingPath = JoinPath(outputFolder, outputName & "-designer.xlsb")
    DeleteFile workingPath
    FileCopy designerPath, workingPath
    AddToReport "designer copied to " & workingPath

    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    'The setup language of the run, read off the FIRST row's setup. A row
    'that names its own language overrides it below.
    ResolveLanguages RowSpecValue(FirstRowSpec(rowsSpec), "setup")
    AddToReport "languages resolved, setup " & optSetupLang

    Set designerBook = Application.Workbooks.Open(fileName:=workingPath, ReadOnly:=False)
    AddToReport "designer copy opened"

    RefreshSourceCode designerBook, sourceRoot, formsFolder

    ResolveInterfaceLanguage designerBook
    AddToReport "interface language resolved: " & optLinelistLang

    PrepareDesignerGeo designerBook
    AddToReport "designer geo prepared"

    'The template rides on Main for the whole run. It is written even when
    'the option is empty, because an empty template is the buttons build
    'and the copied designer's own value would otherwise decide it.
    WriteEntry designerBook.Worksheets(SHEET_MAIN), RNG_LL_TEMPLATE, optTemplate
    AddToReport "template entry written: " & _
                IIf(LenB(optTemplate) = 0, "(none, buttons build)", optTemplate)

    Set multiTable = EventsDesignerMulti.ResolveMultiTable(designerBook)
    If multiTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "The designer copy carries no " & TABLE_MULTI & " table on the " & _
                   SHEET_GENERATE_MULTIPLE & " worksheet, so there is no multi run to make."
    End If

    rowCount = FillMultiTable(multiTable, rowsSpec, designerBook, outputFolder)
    AddToReport "T_Multi filled with " & CStr(rowCount) & " row(s)"

    'The entry over the COPY's Main. Everything the loop writes per row
    'lands here, and GenerateOne reads its designer back off this sheet.
    Set entry = DesignerEntry.Create(designerBook.Worksheets(SHEET_MAIN))

    'One log for the run, on the copy's __check sheet. The multi driver
    'opens it bare: every row files its own header bundle.
    EventsDesignerAdvanced.StartRunLog vbNullString, vbNullString, designerBook
    AddToReport "run log opened on the designer copy"

    EventsDesignerMulti.GenerateMultipleRows multiTable, entry, Nothing, _
                                             lastBuilt, lastFailed
    AddToReport "loop finished: " & CStr(lastBuilt) & " built, " & _
                CStr(lastFailed) & " failed"

    EventsDesignerAdvanced.FinishRunLog CStr(lastBuilt) & " linelist(s) built, " & _
                                        CStr(lastFailed) & " failed"

    'The one text report of the run, named after the run.
    On Error Resume Next
        lastLog = EventsDesignerAdvanced.RunLog().ExportText(outputFolder, outputName)
        Err.Clear
    On Error GoTo Failed

    ReadRowOutcomes multiTable

    designerBook.Save
    designerBook.Close SaveChanges:=False
    Set designerBook = Nothing

    appScope.Restore

    'A run where every row failed answers its own fault. The rows each
    'carry their reason in the result column and in the report, and a
    'caller that read OK over nothing built would deliver an empty folder.
    If lastBuilt = 0 Then
        BuildMultipleFromTable = "ERROR 0: no row built. " & CStr(lastFailed) & _
                                 " row(s) failed; read the generation report."
        Exit Function
    End If

    BuildMultipleFromTable = OUTCOME_OK
    Exit Function

Failed:
    BuildMultipleFromTable = "ERROR " & CStr(Err.Number) & " (" & Err.Source & "): " & _
                             Err.Description
    AddToReport "failed: " & CStr(Err.Number) & " (" & Err.Source & ") " & _
                Err.Description

    On Error Resume Next
        If Not designerBook Is Nothing Then
            designerBook.Save
            designerBook.Close SaveChanges:=False
        End If

        'Same non-raising lookup as BuildLinelistFromSetup's handler: indexing
        'Workbooks by a missing name raises 9, and a raise inside this handler
        'can stop a headless run with a dialog.
        Dim orphanBook As Workbook
        Dim openBook As Workbook
        For Each openBook In Application.Workbooks
            If openBook.Name = "__temp.xlsb" Then
                Set orphanBook = openBook
                Exit For
            End If
        Next
        If Not orphanBook Is Nothing Then orphanBook.Close SaveChanges:=False

        If Not appScope Is Nothing Then appScope.Restore
        Err.Clear
    On Error GoTo 0
End Function

'@Description("Write the rows of a run into the designer copy's T_Multi table.")
'@details
'THE SHAPE OF rowsSpec. Rows are separated by "~~" and the fields of one
'row by "|", each field a key=value pair:
'
'    setup=<path>|outname=<stem>|geo=<path>|setuplang=<column>|
'    lllang=<code>|password=<text>|debugpassword=<text>|
'    epiweek=<value>|design=<name>
'
'Only `setup` and `outname` have to be there. Every other field falls
'back to the run's own option, and then to what the copied designer
'already holds, so a caller names the language once for the whole run
'and a row that wants its own says so.
'
'The output folder of every row is the run's output folder. Excel writes
'inside the granted root and nowhere else, and the launcher copies the
'finished files out afterwards.
'
'The table's own rows go first. A designer ships T_Multi with rows in it,
'and a leftover row carrying a setup path would build a linelist nobody
'asked for.
'@param multiTable ListObject. The T_Multi table on the designer copy.
'@param rowsSpec String. The rows to build.
'@param designerBook Workbook. The designer copy, read for the fallback entries.
'@param outputFolder String. Where every row writes its linelist.
'@return Long. The number of rows written.
Private Function FillMultiTable(ByVal multiTable As ListObject, _
                                ByVal rowsSpec As String, _
                                ByVal designerBook As Workbook, _
                                ByVal outputFolder As String) As Long
    Dim rowSpecs() As String
    Dim counter As Long
    Dim rowIdx As Long
    Dim oneSpec As String
    Dim defaultDesign As String
    Dim defaultEpiweek As String

    defaultDesign = ReadMainEntry(designerBook, RNG_DESIGN_LL)
    defaultEpiweek = ReadMainEntry(designerBook, RNG_EPIWEEK)

    rowSpecs = Split(rowsSpec, ROW_SEPARATOR)

    'Every row the designer shipped goes, so the run builds what the
    'caller asked for and nothing else.
    Do While multiTable.ListRows.Count > 0
        multiTable.ListRows(1).Delete
    Loop

    For counter = LBound(rowSpecs) To UBound(rowSpecs)
        oneSpec = Trim$(rowSpecs(counter))
        If LenB(oneSpec) > 0 Then
            multiTable.ListRows.Add
            rowIdx = multiTable.ListRows.Count

            SetRowCell multiTable, rowIdx, COL_MULTI_SETUPS, _
                       RowSpecValue(oneSpec, "setup")
            SetRowCell multiTable, rowIdx, COL_MULTI_OUTPUT_FILES, _
                       RowSpecValue(oneSpec, "outname")
            SetRowCell multiTable, rowIdx, COL_MULTI_OUTPUT_FOLDERS, outputFolder
            SetRowCell multiTable, rowIdx, COL_MULTI_GEOBASES, _
                       FallbackValue(RowSpecValue(oneSpec, "geo"), optGeo)
            SetRowCell multiTable, rowIdx, COL_MULTI_LANG_DICTIONARY, _
                       FallbackValue(RowSpecValue(oneSpec, "setuplang"), optSetupLang)
            SetRowCell multiTable, rowIdx, COL_MULTI_LANG_INTERFACE, _
                       FallbackValue(RowSpecValue(oneSpec, "lllang"), optLinelistLang)
            SetRowCell multiTable, rowIdx, COL_MULTI_PASSWORD, _
                       FallbackValue(RowSpecValue(oneSpec, "password"), optPassword)
            SetRowCell multiTable, rowIdx, COL_MULTI_DEBUG_PASSWORD, _
                       RowSpecValue(oneSpec, "debugpassword")
            SetRowCell multiTable, rowIdx, COL_MULTI_EPIWEEK_START, _
                       FallbackValue(RowSpecValue(oneSpec, "epiweek"), defaultEpiweek)
            SetRowCell multiTable, rowIdx, COL_MULTI_DESIGN, _
                       FallbackValue(RowSpecValue(oneSpec, "design"), defaultDesign)
            SetRowCell multiTable, rowIdx, COL_MULTI_RESULT, vbNullString

            AddToReport "row " & CStr(rowIdx) & ": " & _
                        RowSpecValue(oneSpec, "outname") & " from " & _
                        BaseName(RowSpecValue(oneSpec, "setup"))
        End If
    Next

    'The IDs the log headers name each row by. The multi module owns this
    'rule, so the fill asks it rather than numbering the rows itself.
    EventsDesignerMulti.EnsureRowIds multiTable

    FillMultiTable = multiTable.ListRows.Count
End Function

'@Description("Take every row's outcome into the narrative of the run.")
'@details
'The result column is what the loop wrote per row, and it is the only
'account of a row that failed while the run as a whole answered OK. A
'caller outside the process reads the narrative and never opens the
'designer copy.
'@param multiTable ListObject. The T_Multi table after the run.
Private Sub ReadRowOutcomes(ByVal multiTable As ListObject)
    Dim rowIdx As Long
    Dim outName As String
    Dim outcomeText As String

    If multiTable.DataBodyRange Is Nothing Then Exit Sub

    For rowIdx = 1 To multiTable.ListRows.Count
        outName = ReadRowCell(multiTable, rowIdx, COL_MULTI_OUTPUT_FILES)
        outcomeText = ReadRowCell(multiTable, rowIdx, COL_MULTI_RESULT)

        AddToReport "row " & CStr(rowIdx) & " -> " & outcomeText & _
                    "  [" & outName & "]"

        'The first written path is what the summary carries, so a caller
        'holding one path has a real file to look at.
        If LenB(lastLinelist) = 0 Then
            If StrComp(outcomeText, RESULT_BUILT, vbTextCompare) = 0 Then
                lastLinelist = outName
            End If
        End If
    Next
End Sub

'@Description("Read one value out of a row spec. An absent key answers empty.")
'@param oneSpec String. One row of the spec.
'@param keyName String. The field to read.
'@return String. The value.
Private Function RowSpecValue(ByVal oneSpec As String, _
                              ByVal keyName As String) As String
    Dim fields() As String
    Dim counter As Long
    Dim oneField As String
    Dim signAt As Long

    If LenB(oneSpec) = 0 Then Exit Function

    fields = Split(oneSpec, "|")

    For counter = LBound(fields) To UBound(fields)
        oneField = Trim$(fields(counter))
        signAt = InStr(1, oneField, "=")

        If signAt > 0 Then
            If StrComp(Trim$(Left$(oneField, signAt - 1)), keyName, vbTextCompare) = 0 Then
                RowSpecValue = Trim$(Mid$(oneField, signAt + 1))
                Exit Function
            End If
        End If
    Next
End Function

'@Description("The first row of a spec, for the values the whole run reads once.")
'@param rowsSpec String. The rows of the run.
'@return String. The first row, empty when there is none.
Private Function FirstRowSpec(ByVal rowsSpec As String) As String
    Dim rowSpecs() As String

    If LenB(Trim$(rowsSpec)) = 0 Then Exit Function

    rowSpecs = Split(rowsSpec, ROW_SEPARATOR)
    FirstRowSpec = Trim$(rowSpecs(LBound(rowSpecs)))
End Function

'@Description("The first value when it carries something, the second otherwise.")
'@param wanted String. What the row asked for.
'@param fallback String. What the run holds.
'@return String. The value to write.
Private Function FallbackValue(ByVal wanted As String, _
                               ByVal fallback As String) As String
    If LenB(Trim$(wanted)) > 0 Then
        FallbackValue = wanted
    Else
        FallbackValue = fallback
    End If
End Function

'@Description("Write one cell of a T_Multi row by column header. A missing column is skipped.")
'@param multiTable ListObject. The T_Multi table.
'@param rowIdx Long. The ListRows position.
'@param colName String. The column header.
'@param cellValue String. What to write.
Private Sub SetRowCell(ByVal multiTable As ListObject, _
                       ByVal rowIdx As Long, _
                       ByVal colName As String, _
                       ByVal cellValue As String)
    Dim target As Range

    Set target = RowCellOf(multiTable, rowIdx, colName)
    If target Is Nothing Then
        AddToReport "T_Multi carries no column named " & colName
        Exit Sub
    End If

    target.Value = cellValue
End Sub

'@Description("Read one cell of a T_Multi row by column header.")
'@param multiTable ListObject. The T_Multi table.
'@param rowIdx Long. The ListRows position.
'@param colName String. The column header.
'@return String. The trimmed cell text, empty when the column is missing.
Private Function ReadRowCell(ByVal multiTable As ListObject, _
                             ByVal rowIdx As Long, _
                             ByVal colName As String) As String
    Dim target As Range

    Set target = RowCellOf(multiTable, rowIdx, colName)
    If target Is Nothing Then Exit Function
    ReadRowCell = Trim$(CStr(target.Value))
End Function

'@Description("The one cell of a T_Multi row by column header.")
'@param multiTable ListObject. The T_Multi table.
'@param rowIdx Long. The ListRows position.
'@param colName String. The column header.
'@return Range. The cell, or Nothing when the column is missing.
Private Function RowCellOf(ByVal multiTable As ListObject, _
                           ByVal rowIdx As Long, _
                           ByVal colName As String) As Range
    Dim col As ListColumn

    On Error Resume Next
        Set col = multiTable.ListColumns(colName)
        Err.Clear
    On Error GoTo 0

    If col Is Nothing Then Exit Function
    Set RowCellOf = multiTable.ListRows(rowIdx).Range.Cells(1, col.Index)
End Function

'@Description("The first path of a multi run that is not on disk.")
'@details
'Every row's setup is checked here rather than inside the loop. A run
'that would fail its third row on a path typo is worth stopping before
'Excel spends four minutes on the first two.
'@param designerPath String. The designer workbook.
'@param rowsSpec String. The rows of the run.
'@param sourceRoot String. The repository root.
'@param formsFolder String. The merged forms folder.
'@return String. The complaint, empty when everything is there.
Private Function FirstMissingMultiPath(ByVal designerPath As String, _
                                       ByVal rowsSpec As String, _
                                       ByVal sourceRoot As String, _
                                       ByVal formsFolder As String) As String
    Dim rowSpecs() As String
    Dim counter As Long
    Dim oneSpec As String
    Dim setupPath As String

    If Len(Dir$(designerPath)) = 0 Then
        FirstMissingMultiPath = "no designer workbook at " & designerPath
        Exit Function
    End If

    If LenB(Trim$(rowsSpec)) = 0 Then
        FirstMissingMultiPath = "the run carries no rows, so there is nothing to build"
        Exit Function
    End If

    If Not IsFolder(JoinPath(sourceRoot, CLASSES_FOLDER)) Then
        FirstMissingMultiPath = "no " & CLASSES_FOLDER & " folder under " & sourceRoot
        Exit Function
    End If

    If Not IsFolder(formsFolder) Then
        FirstMissingMultiPath = "no merged forms folder at " & formsFolder & _
                                " - run scripts/headless/merge-form-code.R first"
        Exit Function
    End If

    rowSpecs = Split(rowsSpec, ROW_SEPARATOR)

    For counter = LBound(rowSpecs) To UBound(rowSpecs)
        oneSpec = Trim$(rowSpecs(counter))
        If LenB(oneSpec) > 0 Then
            setupPath = RowSpecValue(oneSpec, "setup")

            If LenB(setupPath) = 0 Then
                FirstMissingMultiPath = "row " & CStr(counter + 1) & " names no setup"
                Exit Function
            End If

            If Len(Dir$(setupPath)) = 0 Then
                FirstMissingMultiPath = "row " & CStr(counter + 1) & _
                                        ": no setup workbook at " & setupPath
                Exit Function
            End If

            If LenB(RowSpecValue(oneSpec, "outname")) = 0 Then
                FirstMissingMultiPath = "row " & CStr(counter + 1) & _
                                        " names no output file"
                Exit Function
            End If
        End If
    Next
End Function


'@section What the last run did
'===============================================================================
'@description
'A caller holding "OK" still knows nothing about the size of what was built,
'and an empty linelist reports OK as readily as a full one. These six are what
'a test asserts on and what an R session reads back.

'@Description("The narrative of the last run, one line per step.")
'@return String. The report, empty before the first run.
Public Property Get LastReport() As String
    LastReport = runNarrative
End Property

'@Description("How many data entry sheets the last run built.")
'@return Long. The sheet count.
Public Property Get LastSheetCount() As Long
    LastSheetCount = lastSheets
End Property

'@Description("How many variables the last run wrote across every sheet.")
'@return Long. The variable count.
Public Property Get LastVariableCount() As Long
    LastVariableCount = lastVariables
End Property

'@Description("How many VBA components were re-imported into the designer copy.")
'@return Long. The component count.
Public Property Get LastComponentCount() As Long
    LastComponentCount = lastComponents
End Property

'@Description("The full path of the linelist the last run wrote.")
'@return String. The path, empty when nothing was saved.
Public Property Get LastLinelistPath() As String
    LastLinelistPath = lastLinelist
End Property

'@Description("The full path of the generation log the last run wrote.")
'@return String. The path, empty when no log was written.
Public Property Get LastLogPath() As String
    LastLogPath = lastLog
End Property

'@Description("Everything the last run recorded, in one call.")
'@details
'The six above are Property Get, and a Property Get cannot be reached through
'Application.Run. So a caller from OUTSIDE the process -- osascript, and the R
'session behind it -- can read the outcome string of the build and nothing
'else: it gets "OK" with no idea whether the linelist holds three sheets or
'none. This answers all of it in one Function call, which Application.Run can
'reach.
'
'The narrative carries newlines of its own, so it goes last, below a marker,
'and a reader takes everything after that marker verbatim.
'@return String. Key=value lines, a marker, then the narrative. Safe to call
'before any run has happened.
Public Function LastBuildSummary() As String
    LastBuildSummary = "linelist=" & lastLinelist & vbLf & _
                       "log=" & lastLog & vbLf & _
                       "platform=" & PlatformTag() & vbLf & _
                       "sheets=" & CStr(lastSheets) & vbLf & _
                       "variables=" & CStr(lastVariables) & vbLf & _
                       "components=" & CStr(lastComponents) & vbLf & _
                       "built=" & CStr(lastBuilt) & vbLf & _
                       "failed=" & CStr(lastFailed) & vbLf & _
                       "--report--" & vbLf & runNarrative
End Function


'@section The generation itself
'===============================================================================

'@Description("Run the build phases over a prepared designer workbook.")
'@details
'The order of EventsDesignerAdvanced.GenerateOne: specifications, linelist,
'one pass per data entry sheet, the two dropdown stores, the analyses, the
'save. The run log takes each phase's checkings as that phase completes, so a
'build that dies still leaves a report of the phases that finished.
'
'The log is opened over the designer's __check worksheet and is optional: a
'designer without that sheet builds the same linelist and reports it in the
'narrative rather than raising.
'@param designerBook Workbook. The prepared designer copy.
'@param setupPath String. The setup to generate from.
'@param outputFolder String. Where the log is written.
'@param outputName String. The base name of the log file.
'@throws Whatever any build phase raises.
Private Sub RunGeneration(ByVal designerBook As Workbook, _
                          ByVal setupPath As String, _
                          ByVal outputFolder As String, _
                          ByVal outputName As String)
    Dim specs As LinelistSpecs
    Dim ll As Linelist
    Dim runLog As GenerationLog
    Dim sheetList As BetterArray
    Dim sheetInfo As LLSheets
    Dim anaOut As AnalysisOutput
    Dim store As DropdownLists
    Dim counter As Long

    Set runLog = OpenRunLog(designerBook, setupPath, outputName)

    Set specs = LinelistSpecs.Create(designerBook)
    AddToReport "preparing the specifications from " & setupPath
    specs.Prepare setupPath
    AddToReport "specifications prepared, template build: " & CStr(specs.HasTemplate())

    If Not runLog Is Nothing Then runLog.Harvest specs
    If InitTransfer.HasCheckings() Then CollectInto runLog, InitTransfer.CheckingValues()
    AddToReport "specification checkings harvested"

    Set ll = Linelist.Create(specs)
    AddToReport "preparing the linelist workbook"
    ll.Prepare
    AddToReport "linelist prepared and code transferred"

    If ll.HasCheckings Then CollectInto runLog, ll.CheckingValues

    Set sheetList = ll.SheetNames
    Set sheetInfo = ll.SheetInfoManager

    For counter = sheetList.LowerBound To sheetList.UpperBound
        AddToReport "building data entry sheet " & CStr(sheetList.Item(counter))
        BuildOneDataSheet ll, sheetInfo, CStr(sheetList.Item(counter)), runLog
    Next

    AddToReport "built " & CStr(lastSheets) & " data entry sheet(s), " & _
                CStr(lastVariables) & " variable(s)"

    AddToReport "writing the dropdown stores"
    Set store = ll.Dropdown(1)
    If store.HasCheckings Then CollectInto runLog, store.CheckingValues

    Set store = ll.Dropdown(2)
    If store.HasCheckings Then CollectInto runLog, store.CheckingValues

    AddToReport "writing the analyses"
    Set anaOut = AnalysisOutput.Create(specs.AnalysisObject.Wksh(), ll)

    'The handler keeps the analyses' own entries when the stage raises. Without
    'it a stage that died took them with it, and the generation log said nothing
    'about which scope or which table refused -- see the same note in
    'EventsDesignerAdvanced.GenerateOne, where a Windows type mismatch showed
    'exactly that.
    On Error GoTo AnalysesFailed
    anaOut.WriteAnalysis AnalysisBuildStageAll
    On Error GoTo 0

    If anaOut.HasCheckings Then CollectInto runLog, anaOut.CheckingValues
    AddToReport "analyses written"

    'The path is read before SaveLL, because the save closes the workbook and
    'drops both references to it.
    lastLinelist = JoinPath(specs.Value("lldir"), specs.Value("llname") & ".xlsb")
    AddToReport "saving the linelist to " & lastLinelist
    ll.SaveLL
    AddToReport "saved: " & lastLinelist

    If Not runLog Is Nothing Then
        runLog.Finish "Headless linelist built", lastSheets, lastVariables
        lastLog = runLog.ExportText(outputFolder, outputName)
    End If
    Exit Sub

AnalysesFailed:
    Dim anaErrNumber As Long
    Dim anaErrDesc As String

    anaErrNumber = Err.Number
    anaErrDesc = Err.Description

    'Silently, and the log is written out here as well: the caller turns this
    'raise into the run's outcome and never reaches the export below, so an
    'unexported log is a log nobody can read.
    On Error Resume Next
    If anaOut.HasCheckings Then CollectInto runLog, anaOut.CheckingValues
    If Not runLog Is Nothing Then
        runLog.Finish "Failed writing the analyses: " & anaErrDesc, _
                      lastSheets, lastVariables
        lastLog = runLog.ExportText(outputFolder, outputName)
    End If
    On Error GoTo 0

    AddToReport "failed writing the analyses: " & anaErrDesc
    Err.Raise anaErrNumber, "HeadlessBuild.RunGeneration", anaErrDesc
End Sub

'@Description("Build one data entry sheet and take what it filed into the log.")
'@details
'A dictionary sheet of neither layout is not this routine's to judge: LLSheets
'answers its type and anything else is skipped, which is what GenerateOne's own
'BuildOneSheet does.
'@param ll Linelist. The linelist under construction.
'@param sheetInfo LLSheets. The shared sheet information manager.
'@param sheetName String. The sheet to build.
'@param runLog GenerationLog. The open log, or Nothing.
Private Sub BuildOneDataSheet(ByVal ll As Linelist, _
                              ByVal sheetInfo As LLSheets, _
                              ByVal sheetName As String, _
                              ByVal runLog As GenerationLog)
    Dim sheetType As String
    Dim layer As Byte
    Dim listBld As LLDataEntry

    sheetType = sheetInfo.SheetInfo(sheetName)

    If sheetType = "vlist1D" Then
        layer = LLDataEntryLayerVList
    ElseIf sheetType = "hlist2D" Then
        layer = LLDataEntryLayerHList
    Else
        Exit Sub
    End If

    Set listBld = LLDataEntry.Create(layer, sheetName, ll, sheetInfo)
    listBld.Build

    If listBld.HasCheckings Then CollectInto runLog, listBld.CheckingValues
    If listBld.HasMilestones Then CollectInto runLog, listBld.MilestoneValues, True

    lastSheets = lastSheets + 1
    lastVariables = lastVariables + listBld.VariablesWritten
End Sub

'@Description("Open the run log over the designer's __check worksheet.")
'@details
'A designer without that worksheet still builds a linelist, so the failure is
'recorded in the narrative and the build carries on with no log.
'@param designerBook Workbook. The designer copy.
'@param setupPath String. What the log header names as the source.
'@param outputName String. What the log header names as the output.
'@return GenerationLog. The open log, or Nothing.
Private Function OpenRunLog(ByVal designerBook As Workbook, _
                            ByVal setupPath As String, _
                            ByVal outputName As String) As GenerationLog
    Dim builtLog As GenerationLog

    On Error Resume Next
        Set builtLog = GenerationLog.Create(designerBook)
        If Not builtLog Is Nothing Then builtLog.Start setupPath, outputName
    On Error GoTo 0

    If builtLog Is Nothing Then
        AddToReport "no run log: the designer carries no __check worksheet"
    End If

    Set OpenRunLog = builtLog
End Function

'@Description("Hand one bundle of checkings to the log when there is a log.")
'@param runLog GenerationLog. The open log, or Nothing.
'@param checks Checking. The bundle.
'@param recordOnly Optional Boolean. True keeps the bundle out of the worksheet.
Private Sub CollectInto(ByVal runLog As GenerationLog, _
                        ByVal checks As Checking, _
                        Optional ByVal recordOnly As Boolean = False)
    If runLog Is Nothing Then Exit Sub
    If checks Is Nothing Then Exit Sub
    runLog.Collect checks, recordOnly
End Sub


'@section Preparing the designer copy
'===============================================================================

'@Description("Replace every component of a workbook's project with the source on disk.")
'@details
'Each .cls under <sourceRoot>/src/classes, each .bas under
'<sourceRoot>/src/modules and each .frm in the forms folder is imported,
'removing first whatever component of that name the workbook already carries.
'
'THE WHOLE TREE, NOT THE TRANSFER LIST
'-------------------------------------------------------------------------------
'Linelist.TransferAllCode names 55 components. Copying that list here would put
'a second copy of it in the project, and the copy would go stale the first time
'a component was added to the transfer -- silently, because the build would
'then export the designer's own stale copy of it and report nothing. Importing
'the tree cannot drift.
'
'The removal before each import is what keeps this safe. VBComponents.Import
'over a name the project already holds imports under a SECOND name and keeps
'both, which is how a workbook ends up carrying Passwords and Passwords1.
'@param target Workbook. The workbook whose project is rewritten.
'@param sourceRoot String. The repository root.
'@param formsFolder String. The folder of merged .frm files.
Private Sub RefreshSourceCode(ByVal target As Workbook, _
                              ByVal sourceRoot As String, _
                              ByVal formsFolder As String)
    lastComponents = 0

    AddToReport "stripping the designer copy's own project"
    StripProject target
    AddToReport "project stripped, importing the source"

    ImportNamedFolders target, JoinPath(sourceRoot, CLASSES_FOLDER), _
                       TRANSFER_CLASS_FOLDERS, "*.cls"
    ImportNamedFolders target, JoinPath(sourceRoot, MODULES_FOLDER), _
                       TRANSFER_MODULE_FOLDERS, "*.bas"
    ImportFolder target, formsFolder, "*.frm"

    AddToReport "re-imported " & CStr(lastComponents) & " component(s) from source"
End Sub

'@Description("Remove every component of a workbook's project that is not a document.")
'@details
'Worksheet and workbook components cannot be removed and are left alone;
'everything else goes, so what the copy carries afterwards is what this run put
'there and nothing else.
'
'The names are collected before the first removal. Removing from a collection
'while a For Each walks it skips entries, and the entries it skips here are
'components the transfer would then export stale.
'@param target Workbook. The workbook whose project is emptied.
Private Sub StripProject(ByVal target As Workbook)
    Dim proj As Object
    Dim comp As Object
    Dim doomed As BetterArray
    Dim counter As Long

    Set proj = ProjectOf(target)
    Set doomed = New BetterArray
    doomed.LowerBound = 1

    For Each comp In proj.VBComponents
        If comp.Type <> COMPONENT_DOCUMENT Then doomed.Push comp.Name
    Next

    For counter = doomed.LowerBound To doomed.UpperBound
        RemoveModule target, CStr(doomed.Item(counter))
    Next

    AddToReport "stripped " & CStr(doomed.Length) & " component(s) from the designer copy"
End Sub

'@Description("Import one file pattern from each of the folders named under a root.")
'@details
'Only the folders named are opened. A sandboxed Excel prompts for every folder
'it has no grant for, so a walk of the whole tree is a walk through a stack of
'dialogs -- and the folders it would add carry nothing a linelist can use.
'
'Dir$ holds ONE search at a time, which is why the file names of a folder are
'read to the end before any of them is imported.
'@param target Workbook. The workbook to import into.
'@param rootPath String. The folder holding the component folders.
'@param folderList String. Pipe-separated folder names.
'@param pattern String. The file pattern, "*.cls" or "*.bas".
Private Sub ImportNamedFolders(ByVal target As Workbook, _
                               ByVal rootPath As String, _
                               ByVal folderList As String, _
                               ByVal pattern As String)
    Dim folderNames() As String
    Dim counter As Long
    Dim childPath As String

    folderNames = Split(folderList, "|")

    For counter = LBound(folderNames) To UBound(folderNames)
        childPath = JoinPath(rootPath, folderNames(counter))

        If IsFolder(childPath) Then
            ImportFolder target, childPath, pattern
        Else
            AddToReport "source folder missing: " & childPath
        End If
    Next
End Sub

'@Description("Import every matching file of one folder, replacing what is there.")
'@param target Workbook. The workbook to import into.
'@param folderPath String. The folder to read.
'@param pattern String. The file pattern.
Private Sub ImportFolder(ByVal target As Workbook, _
                         ByVal folderPath As String, _
                         ByVal pattern As String)
    Dim fileNames As BetterArray
    Dim counter As Long
    Dim filePath As String
    Dim proj As Object

    Set fileNames = MatchingNames(folderPath, pattern)
    If fileNames.Length = 0 Then Exit Sub

    Set proj = ProjectOf(target)

    For counter = fileNames.LowerBound To fileNames.UpperBound
        filePath = JoinPath(folderPath, CStr(fileNames.Item(counter)))
        RemoveModule target, BaseName(filePath)
        proj.VBComponents.Import filePath
        lastComponents = lastComponents + 1
    Next
End Sub

'@Description("Seed the hidden names the designer's Geo worksheet must carry.")
'@details
'LLGeo.CheckRequirements asks the Geo worksheet for RNG_GeoName,
'RNG_GeoUpdated, RNG_GeoLangCode and RNG_MetaLang, for the five level labels at
'workbook scope, and for the RNG_PastingGeoCol cell. A designer that has never
'been through DesignerPreparation.Prepare with the current code carries none of
'them, and the consequences are quiet and total: LLGeo.Create fails, the
'dictionary files "Geo object should be of type LLGeo, geolines not append",
'AppendGeoLines never runs, and every geo variable stays ONE column instead of
'expanding into the twelve it owns -- four admin levels, four p-codes and four
'concatenations. The delivered linelist then has no geography at all and says so
'only in the run log.
'
'This is the seeding half of Prepare and nothing else. Prepare's first step
'opens a file dialog to import translations, so a headless run cannot go through
'it; EnsureGeoFlags is Public for exactly this call. Every write is an
'EnsureName, so a designer already carrying the names is left as it is.
'
'A failure is reported rather than raised. A build with no geography is a real
'linelist and a caller may want it; a build that dies here delivers nothing.
'@param designerBook Workbook. The designer copy.
Private Sub PrepareDesignerGeo(ByVal designerBook As Workbook)
    Dim prep As DesignerPreparation

    On Error Resume Next
        Set prep = DesignerPreparation.Create(designerBook)
        If Not prep Is Nothing Then prep.EnsureGeoFlags

        If Err.Number <> 0 Then
            'Number first. This failure is SWALLOWED, so this line is the only
            'trace of it, and a description that arrived empty across a class
            'boundary would leave the line naming nothing.
            AddToReport "the designer's Geo worksheet could not be seeded " & _
                        "(error " & CStr(Err.Number) & ": " & Err.Description & _
                        "), so the build carries no geography"
            Err.Clear
        Else
            AddToReport "geo hidden names seeded on the designer copy"
        End If
    On Error GoTo 0
End Sub

'@Description("Import a linelist style workbook onto the designer copy's format sheet.")
'@details
'What the ribbon's clickImpStyle does, driven from code. LLFormat.Import
'merges the style workbook's design columns into the designer's own
'T_Formatter and copies the file's DESIGNTYPE across.
'
'THE FLAG IS WHAT MAKES THE IMPORT COUNT. InitTransfer ships the DESIGNER's
'formatter only when DesignerPreparation.FormatterImported reads True, and the
'SETUP's formatter otherwise. Import the columns and leave the flag alone, and
'the linelist comes out formatted from the setup while the imported style sits
'in the designer copy, silent, and reads as though it had been used.
'
'A run that names no style exits here, and that is the common run: the setup's
'own formatter is the design a build gets when nobody asks for another.
'@param designerBook Workbook. The designer copy.
'@throws ProjectError.ElementNotFound When the style workbook cannot be read.
Private Sub ImportLinelistStyle(ByVal designerBook As Workbook)
    Dim styleBook As Workbook
    Dim formatManager As LLFormat
    Dim prep As DesignerPreparation

    If LenB(optStyle) = 0 Then Exit Sub

    If Len(Dir$(optStyle)) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "The style workbook does not exist: " & optStyle
    End If

    Set styleBook = Application.Workbooks.Open(fileName:=optStyle, ReadOnly:=True)

    'Closed on the way out of a failure too. A style workbook left open is a
    'workbook the quit at the end of the run stops to ask about.
    On Error GoTo CloseAndRaise
    Set formatManager = LLFormat.Create(designerBook.Worksheets(SHEET_FORMATTER))
    formatManager.Import styleBook.Worksheets(1)
    On Error GoTo 0

    styleBook.Close SaveChanges:=False
    Set styleBook = Nothing

    Set prep = DesignerPreparation.Create(designerBook)
    prep.FormatterImported = True

    AddToReport "linelist style imported from " & optStyle & _
                ", the designer's formatter is the live one"
    Exit Sub

CloseAndRaise:
    Dim failedNumber As Long
    Dim failedText As String
    Dim failedSource As String

    failedNumber = Err.Number
    failedText = Err.Description
    failedSource = Err.Source

    On Error Resume Next
        If Not styleBook Is Nothing Then styleBook.Close SaveChanges:=False
        Err.Clear
    On Error GoTo 0

    Err.Raise failedNumber, failedSource, failedText
End Sub

'@Description("Name the design the run is written in, on the designer copy.")
'@details
'DESIGNTYPE on the format sheet is what decides the design, in both branches:
'LinelistSpecs.EnsureFormat builds its LLFormat with it, and
'InitTransfer.DesignerDesignName reads it to pick the column of whichever
'formatter is shipped. RNG_DesignLL on Main is the dropdown a person fills, and
'in the designer the workbook's own event handlers copy that cell into
'DESIGNTYPE. This run threw those handlers away with the rest of the designer's
'code, so the name is written to both places here.
'
'A design the format table has no column for still builds. LLFormat falls back
'to its own default for an unknown name, which is the answer a designer gives
'as well, so the run reports the value it wrote and a reader can see when the
'fallback took over.
'@param designerBook Workbook. The designer copy.
Private Sub ApplyDesignChoice(ByVal designerBook As Workbook)
    Dim formatSheet As Worksheet
    Dim designRange As Range

    If LenB(optDesign) = 0 Then Exit Sub

    On Error Resume Next
        Set formatSheet = designerBook.Worksheets(SHEET_FORMATTER)
        Err.Clear
    On Error GoTo 0

    If formatSheet Is Nothing Then
        AddToReport "design not applied, the designer copy carries no '" & _
                    SHEET_FORMATTER & "' worksheet: " & optDesign
        Exit Sub
    End If

    On Error Resume Next
        Set designRange = formatSheet.Range(RANGE_DESIGN_TYPE)
        Err.Clear
    On Error GoTo 0

    If designRange Is Nothing Then
        AddToReport "design not applied, no '" & RANGE_DESIGN_TYPE & _
                    "' range on " & SHEET_FORMATTER & ": " & optDesign
        Exit Sub
    End If

    designRange.Cells(1, 1).Value = optDesign
    AddToReport "design applied: " & optDesign
End Sub

'@Description("Write the build entries onto the designer's Main worksheet.")
'@details
'Three entries have to land, because nothing else says where the linelist goes
'or which of the two builds runs: the output folder, the output name and the
'template path. A designer whose Main carries none of those names is not a
'designer this routine can drive, and it says so rather than building
'something into a place nobody asked for.
'
'The template entry is written even when the option is EMPTY. An empty value is
'the buttons build, and leaving whatever the copied designer held would hand
'the run someone else's choice.
'@param designerBook Workbook. The designer copy.
'@param setupPath String. The setup the build reads.
'@param outputFolder String. Where the linelist goes.
'@param outputName String. What it is called.
'@throws ProjectError.ElementNotFound When a required entry name is missing.
Private Sub WriteDesignerEntries(ByVal designerBook As Workbook, _
                                 ByVal setupPath As String, _
                                 ByVal outputFolder As String, _
                                 ByVal outputName As String)
    Dim mainSheet As Worksheet

    On Error GoTo NoMainSheet
    Set mainSheet = designerBook.Worksheets(SHEET_MAIN)
    On Error GoTo 0

    RequireEntry mainSheet, RNG_LL_DIR, outputFolder
    RequireEntry mainSheet, RNG_LL_NAME, outputName
    RequireEntry mainSheet, RNG_LL_TEMPLATE, optTemplate

    WriteEntry mainSheet, RNG_PATH_DICO, setupPath
    WriteEntry mainSheet, RNG_PATH_GEO, optGeo
    WriteEntry mainSheet, RNG_LL_PWD_OPEN, optPassword

    'Each language is left as the copied designer holds it when nothing better
    'was settled on, because the designer's own entry is a valid answer and an
    'empty one is not. ResolveInterfaceLanguage has already checked the linelist
    'one against the translation table's columns by the time this runs.
    If LenB(optSetupLang) > 0 Then WriteEntry mainSheet, RNG_LANG_SETUP, optSetupLang
    If LenB(optLinelistLang) > 0 Then WriteEntry mainSheet, RNG_LL_FORM, optLinelistLang

    'The dropdown beside the design ApplyDesignChoice already wrote onto the
    'format sheet. The build reads DESIGNTYPE for the design it applies, and
    'this cell keeps the copy's own record straight: a designer whose Main says
    'one design while its formatter holds another misleads the next reader.
    If LenB(optDesign) > 0 Then WriteEntry mainSheet, RNG_DESIGN_LL, optDesign

    AddToReport "entries written on Main, template: " & _
                IIf(LenB(optTemplate) = 0, "(none, buttons build)", optTemplate)
    Exit Sub

NoMainSheet:
    ThrowError ProjectError.ElementNotFound, _
               "The designer copy carries no '" & SHEET_MAIN & "' worksheet."
End Sub

'@Description("Write one entry that the build cannot run without.")
'@param mainSheet Worksheet. The designer's Main worksheet.
'@param rangeName String. The named range carrying the entry.
'@param entryValue String. What to write.
'@throws ProjectError.ElementNotFound When the name does not resolve.
Private Sub RequireEntry(ByVal mainSheet As Worksheet, _
                         ByVal rangeName As String, _
                         ByVal entryValue As String)
    If WriteEntry(mainSheet, rangeName, entryValue) Then Exit Sub

    ThrowError ProjectError.ElementNotFound, _
               "The designer's '" & SHEET_MAIN & "' worksheet carries no '" & _
               rangeName & "' range, so the build cannot be pointed at an output."
End Sub

'@Description("Write one entry when its named range resolves.")
'@param mainSheet Worksheet. The designer's Main worksheet.
'@param rangeName String. The named range carrying the entry.
'@param entryValue String. What to write.
'@return Boolean. True when the entry landed.
Private Function WriteEntry(ByVal mainSheet As Worksheet, _
                            ByVal rangeName As String, _
                            ByVal entryValue As String) As Boolean
    Dim target As Range

    On Error Resume Next
        Set target = mainSheet.Range(rangeName)
    On Error GoTo 0

    If target Is Nothing Then
        AddToReport "entry not written, no such range: " & rangeName
        Exit Function
    End If

    target.Cells(1, 1).Value = entryValue
    WriteEntry = True
End Function


'@section Options and run state
'===============================================================================

'@Description("Read the pipe-separated options of one run.")
'@details
'Split on an empty string answers a ZERO-LENGTH array, so the empty options
'string is answered before any split runs. A key this routine does not know is
'reported: a misspelled key reads exactly like an option nobody passed, and
'that difference is a build pointed somewhere else.
'@param options String. Pipe-separated key=value pairs.
Private Sub ReadOptions(ByVal options As String)
    Dim pairs() As String
    Dim counter As Long
    Dim onePair As String
    Dim signAt As Long
    Dim keyName As String
    Dim keyValue As String

    optTemplate = vbNullString
    optGeo = vbNullString
    optSetupLang = vbNullString
    optLinelistLang = vbNullString
    optPassword = vbNullString
    optDesign = vbNullString
    optStyle = vbNullString

    If LenB(Trim$(options)) = 0 Then Exit Sub

    pairs = Split(options, "|")

    For counter = LBound(pairs) To UBound(pairs)
        onePair = Trim$(pairs(counter))
        If LenB(onePair) > 0 Then
            'The FIRST separator only: a path is free to carry an equals sign
            'and the value is everything after the key.
            signAt = InStr(1, onePair, "=")

            If signAt = 0 Then
                AddToReport "option ignored, no value: " & onePair
            Else
                keyName = LCase$(Trim$(Left$(onePair, signAt - 1)))
                keyValue = Trim$(Mid$(onePair, signAt + 1))

                Select Case keyName
                    Case "temppath"
                        optTemplate = keyValue
                    Case "geopath"
                        optGeo = keyValue
                    Case "setuplang"
                        optSetupLang = keyValue
                    Case "lllang"
                        optLinelistLang = keyValue
                    Case "llpassword"
                        optPassword = keyValue
                    Case "design"
                        optDesign = keyValue
                    Case "stylepath"
                        optStyle = keyValue
                    Case Else
                        AddToReport "option ignored, unknown key: " & keyName
                End Select
            End If
        End If
    Next
End Sub

'@Description("Fill the setup language the caller left empty from the setup itself.")
'@details
'What ExtractAndUpdateLanguages does when a setup is loaded by hand: the
'first language of the setup's Translations sheet is the auto-selected one
'(owner decision). The read goes through EventsDesignerAdvanced.SetupLanguages,
'the one shared language extraction, so the two paths cannot drift.
'
'ONLY THE SETUP LANGUAGE IS RESOLVED HERE
'-------------------------------------------------------------------------------
'This used to fill the linelist language from the same value, and the two are
'not the same vocabulary. RNG_LangSetup takes a COLUMN NAME of the setup's own
'translation table -- "English" -- and RNG_LLForm takes a CODE-Name entry off
'the interface dropdown, "ENG-English", whose prefix is the column the four
'LinelistTranslation tables are keyed on. Writing "English" into RNG_LLForm made
'RNG_LLLanguageCode read "English", no such column existed, and every lookup
'fell back to the tag: sheets came out named LLSHEET_Analysis and every button
'and message in the delivered linelist read as its own tag. The linelist
'language is resolved in ResolveInterfaceLanguage instead, against the codes the
'designer actually carries.
'@param setupPath String. The setup workbook to read.
Private Sub ResolveLanguages(ByVal setupPath As String)
    Dim setupWkb As Workbook
    Dim tradSheet As Worksheet
    Dim languages As BetterArray
    Dim firstLanguage As String

    If LenB(optSetupLang) > 0 Then Exit Sub

    On Error Resume Next
        Set setupWkb = Application.Workbooks.Open(fileName:=setupPath, ReadOnly:=True)
    On Error GoTo 0
    If setupWkb Is Nothing Then Exit Sub

    On Error Resume Next
        Set tradSheet = setupWkb.Worksheets(SHEET_SETUP_TRANSLATIONS)
        If Not tradSheet Is Nothing Then
            Set languages = EventsDesignerAdvanced.SetupLanguages(tradSheet)
            If languages.Length > 0 Then
                firstLanguage = CStr(languages.Item(languages.LowerBound))
            End If
        End If
        setupWkb.Close SaveChanges:=False
        Err.Clear
    On Error GoTo 0

    If LenB(firstLanguage) = 0 Then
        AddToReport "no language found in the setup, the designer's own entry stands"
        Exit Sub
    End If

    optSetupLang = firstLanguage
    AddToReport "setup language resolved from the setup: " & optSetupLang
End Sub

'@Description("Settle the linelist interface language against the codes the designer carries.")
'@details
'RNG_LLForm is a CODE-Name entry -- "ENG-English" -- and only the prefix
'matters: InitTransfer splits it on the dash and writes the prefix as
'RNG_LLLanguageCode, which is the COLUMN the four LinelistTranslation tables
'are keyed on. A value whose prefix names no column translates nothing at all,
'silently, and the delivered linelist reads as a wall of tags.
'
'So the codes are read off the designer's own T_TradLLMsg headers rather than
'assumed, and three candidates are tried in order: what the caller asked for,
'what the copied designer already holds, and the first code the table offers.
'The first one that names a real column wins, and the choice is reported either
'way -- a build that quietly picked a different language than the caller named
'is exactly the outcome this routine exists to make visible.
'@param designerBook Workbook. The designer copy.
Private Sub ResolveInterfaceLanguage(ByVal designerBook As Workbook)
    Dim codes As BetterArray
    Dim wanted As String
    Dim resolved As String

    Set codes = TranslationLanguageCodes(designerBook)

    If codes.Length = 0 Then
        AddToReport "the designer carries no " & TABLE_LL_MESSAGES & " headers, " & _
                    "so the linelist language was left as the designer holds it"
        Exit Sub
    End If

    'What the caller named, then what the designer already holds.
    resolved = MatchingLanguageCode(codes, optLinelistLang)

    If LenB(resolved) = 0 Then
        wanted = ReadMainEntry(designerBook, RNG_LL_FORM)
        resolved = MatchingLanguageCode(codes, wanted)

        If LenB(resolved) > 0 Then
            optLinelistLang = wanted
            AddToReport "linelist language: the designer's own entry stands (" & wanted & ")"
            Exit Sub
        End If
    Else
        AddToReport "linelist language: " & optLinelistLang & " (code " & resolved & ")"
        Exit Sub
    End If

    'Neither answered a column of the table. The first code it offers is a
    'linelist somebody can read, which a wall of tags is not.
    resolved = CStr(codes.Item(codes.LowerBound))
    AddToReport "linelist language: neither the option (" & optLinelistLang & _
                ") nor the designer entry (" & wanted & ") names a column of " & _
                TABLE_LL_MESSAGES & ", so the build fell back to " & resolved
    optLinelistLang = resolved
End Sub

'@Description("The language codes the designer's linelist message table is keyed on.")
'@details
'The header row of T_TradLLMsg, minus its first column, which carries the tag
'itself. Those headers ARE the languages a linelist can be written in, and
'reading them beats a list written down here that would go stale the first time
'a language was added to the designer.
'@param designerBook Workbook. The designer copy.
'@return BetterArray. The codes, empty when the sheet or the table is missing.
Private Function TranslationLanguageCodes(ByVal designerBook As Workbook) As BetterArray
    Dim tradSheet As Worksheet
    Dim headerRow As Range
    Dim result As BetterArray
    Dim counter As Long
    Dim headerText As String

    Set result = New BetterArray
    result.LowerBound = 1
    Set TranslationLanguageCodes = result

    On Error Resume Next
        Set tradSheet = designerBook.Worksheets(SHEET_LL_TRANSLATION)
        If Not tradSheet Is Nothing Then
            Set headerRow = tradSheet.ListObjects(TABLE_LL_MESSAGES).HeaderRowRange
        End If
        Err.Clear
    On Error GoTo 0

    If headerRow Is Nothing Then Exit Function

    For counter = 2 To headerRow.Cells.Count
        headerText = Trim$(CStr(headerRow.Cells(1, counter).Value))
        If LenB(headerText) > 0 Then result.Push headerText
    Next
End Function

'@Description("The code one language entry resolves to, when it names a real column.")
'@details
'An entry is matched on the part before the dash, which is what InitTransfer
'takes; an entry carrying no dash is matched whole, so a caller may hand over
'either "ENG-English" or a bare "ENG". The comparison is case-insensitive
'because the value comes off a worksheet cell somebody typed into.
'@param codes BetterArray. The codes the table offers.
'@param entryValue String. The language entry to judge.
'@return String. The matching code, empty when there is none.
Private Function MatchingLanguageCode(ByVal codes As BetterArray, _
                                      ByVal entryValue As String) As String
    Dim wantedCode As String
    Dim counter As Long
    Dim oneCode As String

    wantedCode = Trim$(entryValue)
    If LenB(wantedCode) = 0 Then Exit Function

    If InStr(1, wantedCode, "-", vbBinaryCompare) > 0 Then
        wantedCode = Trim$(Split(wantedCode, "-")(0))
    End If

    For counter = codes.LowerBound To codes.UpperBound
        oneCode = CStr(codes.Item(counter))
        If StrComp(oneCode, wantedCode, vbTextCompare) = 0 Then
            MatchingLanguageCode = oneCode
            Exit Function
        End If
    Next
End Function

'@Description("Read one named entry off the designer's Main worksheet.")
'@param designerBook Workbook. The designer copy.
'@param rangeName String. The named range carrying the entry.
'@return String. The value, empty when the name does not resolve.
Private Function ReadMainEntry(ByVal designerBook As Workbook, _
                               ByVal rangeName As String) As String
    Dim target As Range

    On Error Resume Next
        Set target = designerBook.Worksheets(SHEET_MAIN).Range(rangeName)
        Err.Clear
    On Error GoTo 0

    If target Is Nothing Then Exit Function
    ReadMainEntry = Trim$(CStr(target.Cells(1, 1).Value))
End Function

'@Description("Clear what the previous run recorded.")
Private Sub ResetRunState()
    runNarrative = vbNullString
    tracePath = vbNullString
    lastSheets = 0
    lastVariables = 0
    lastComponents = 0
    lastLinelist = vbNullString
    lastLog = vbNullString
    lastBuilt = 0
    lastFailed = 0
End Sub

'@Description("Add one line to the narrative of the run.")
'@details
'The line also goes to the trace file when the run has one, and it is opened
'and closed around each line rather than held open. A held handle loses its
'buffer when the process dies, which is the one case the trace exists for.
'@param messageText String. What happened.
Private Sub AddToReport(ByVal messageText As String)
    If LenB(runNarrative) > 0 Then runNarrative = runNarrative & vbLf
    runNarrative = runNarrative & messageText
    WriteTrace messageText
End Sub

'@Description("Mirror one narrative line into the trace file on disk.")
'@details
'Silent about its own failures: a trace that cannot be written is not a reason
'to stop a build, and the caller has the in-memory narrative either way.
'@param messageText String. The line to append.
Private Sub WriteTrace(ByVal messageText As String)
    Dim handle As Integer

    If LenB(tracePath) = 0 Then Exit Sub

    On Error Resume Next
        handle = FreeFile
        Open tracePath For Append As #handle
        Print #handle, Format$(Now, "hh:nn:ss") & "  " & messageText
        Close #handle
        Err.Clear
    On Error GoTo 0
End Sub


'@section Paths
'===============================================================================

'@Description("The first of the four required paths that is not on disk.")
'@details
'Answered as a sentence rather than a raise, so a checkout without the asset
'binaries says which file to fetch instead of dying inside a build phase.
'@param designerPath String. The designer workbook.
'@param setupPath String. The filled setup.
'@param sourceRoot String. The repository root.
'@param formsFolder String. The merged forms folder.
'@return String. The complaint, empty when all four are there.
Private Function FirstMissingPath(ByVal designerPath As String, _
                                  ByVal setupPath As String, _
                                  ByVal sourceRoot As String, _
                                  ByVal formsFolder As String) As String
    If Len(Dir$(designerPath)) = 0 Then
        FirstMissingPath = "no designer workbook at " & designerPath
        Exit Function
    End If

    If Len(Dir$(setupPath)) = 0 Then
        FirstMissingPath = "no setup workbook at " & setupPath
        Exit Function
    End If

    If Not IsFolder(JoinPath(sourceRoot, CLASSES_FOLDER)) Then
        FirstMissingPath = "no " & CLASSES_FOLDER & " folder under " & sourceRoot
        Exit Function
    End If

    If Not IsFolder(formsFolder) Then
        FirstMissingPath = "no merged forms folder at " & formsFolder & _
                           " - run scripts/headless/merge-form-code.R first"
    End If
End Function

'@Description("Join two path parts with the host's separator.")
'@param parentPath String. The folder.
'@param childName String. What sits inside it.
'@return String. The joined path.
Private Function JoinPath(ByVal parentPath As String, ByVal childName As String) As String
    Dim separator As String

    separator = Application.PathSeparator
    If Right$(parentPath, 1) = separator Then
        JoinPath = parentPath & childName
    Else
        JoinPath = parentPath & separator & childName
    End If
End Function

'@Description("True when a path names a folder that exists.")
'@param folderPath String. The path to judge.
'@return Boolean. True for an existing folder.
Private Function IsFolder(ByVal folderPath As String) As Boolean
    Dim attributes As Long

    On Error Resume Next
        attributes = GetAttr(folderPath)
        If Err.Number = 0 Then IsFolder = ((attributes And vbDirectory) = vbDirectory)
        Err.Clear
    On Error GoTo 0
End Function

'@Description("The names inside a folder matching one pattern.")
'@param folderPath String. The folder to read.
'@param pattern String. The pattern, empty for everything.
'@param searchAttributes Optional Long. What Dir$ is asked for.
'@return BetterArray. The names.
Private Function MatchingNames(ByVal folderPath As String, _
                               ByVal pattern As String, _
                               Optional ByVal searchAttributes As Long = 0) As BetterArray
    Dim result As BetterArray
    Dim entryName As String
    Dim searchPath As String

    Set result = New BetterArray
    result.LowerBound = 1

    If LenB(pattern) = 0 Then
        searchPath = JoinPath(folderPath, "*")
    Else
        searchPath = JoinPath(folderPath, pattern)
    End If

    On Error Resume Next
        entryName = Dir$(searchPath, searchAttributes)
        Do While LenB(entryName) > 0
            If entryName <> "." And entryName <> ".." Then result.Push entryName
            entryName = Dir$()
        Loop
    On Error GoTo 0

    Set MatchingNames = result
End Function

'@Description("Create a folder when it is not on disk yet.")
'@details
'Called before the sandbox grant is asked for, and that is the whole point: a
'security-scoped grant is given for a path that EXISTS, so asking for one over
'a folder no class has created yet grants nothing and every later write into it
'raises a dialog. TemporaryRepos.EnsureReady would make it, but not until the
'build is well past the one moment a dialog can still be answered.
'
'A failure is swallowed: TemporaryRepos creates the folder itself later, and
'the only thing lost is the grant.
'@param folderPath String. The folder to create.
Private Sub EnsureFolder(ByVal folderPath As String)
    If IsFolder(folderPath) Then Exit Sub

    On Error Resume Next
        MkDir folderPath
        Err.Clear
    On Error GoTo 0
End Sub

'@Description("Delete one file when it is there.")
'@param filePath String. The file to drop.
Private Sub DeleteFile(ByVal filePath As String)
    On Error Resume Next
        Kill filePath
        Err.Clear
    On Error GoTo 0
End Sub


'@section The VBProject of another workbook
'===============================================================================

'@Description("Import one .bas file into a workbook's VBProject.")
'@details
'A component of that name already in the target is removed first, so a second
'run over the same file does not meet "name conflicts with existing module".
'@param target Workbook. The workbook to inject into.
'@param modulePath String. Full path of the .bas file.
'@throws Whatever the VBProject raises when trust access is off.
Private Sub InjectModule(ByVal target As Workbook, ByVal modulePath As String)
    Dim componentName As String
    Dim proj As Object

    componentName = BaseName(modulePath)
    RemoveModule target, componentName

    Set proj = ProjectOf(target)
    proj.VBComponents.Import modulePath
End Sub

'@Description("Remove a component from a workbook's VBProject when it is there.")
'@param target Workbook. The workbook to clean.
'@param componentName String. The component to drop.
Private Sub RemoveModule(ByVal target As Workbook, ByVal componentName As String)
    Dim comp As Object
    Dim proj As Object

    On Error Resume Next
        Set proj = ProjectOf(target)
        Set comp = proj.VBComponents(componentName)
    On Error GoTo 0

    If comp Is Nothing Then Exit Sub

    On Error Resume Next
        proj.VBComponents.Remove comp
    On Error GoTo 0
End Sub

'@Description("The VBProject of a workbook, as an Object.")
'@details
'The shape CodeTransfer.TargetProject uses, and for the same two reasons.
'The return is Object, so naming a VBIDE type never puts the library in the
'compile - an identifier from an unreferenced library costs the WHOLE
'project its compile, and the static scans cannot see it. And a project the
'host refuses answers with a message naming the workbook, rather than the
'bare automation error.
'@param target Workbook. The workbook to reach into.
'@return Object. The VBProject.
'@throws ProjectError.ErrorUnexpectedState When trust access is off.
Private Function ProjectOf(ByVal target As Workbook) As Object
    On Error GoTo ProjectBlocked
    Set ProjectOf = target.VBProject
    Exit Function

ProjectBlocked:
    ThrowError ProjectError.ErrorUnexpectedState, _
               "The VBProject of '" & target.Name & "' cannot be reached. " & _
               "Turn on Trust access to the VBA project object model."
End Function

'@Description("The file name of a path, without its folder and extension.")
'@param filePath String. The path to reduce.
'@return String. The bare name.
Private Function BaseName(ByVal filePath As String) As String
    Dim parts() As String
    Dim lastPart As String
    Dim dotAt As Long

    'Both separators, because a caller may hand over either shape.
    parts = Split(Replace(filePath, "\", "/"), "/")
    lastPart = parts(UBound(parts))

    dotAt = InStrRev(lastPart, ".")
    If dotAt > 1 Then
        BaseName = Left$(lastPart, dotAt - 1)
    Else
        BaseName = lastPart
    End If
End Function

'@Description("Raise a project error naming this module as its source.")
'@details
'The shape InitTransfer uses. ThrowError is Private in every class and module
'that has one, so a module reaching for another module's copy does not
'compile - this one is ours.
'@param errorCode Long. A ProjectError member.
'@param messageText String. What went wrong, in words.
Private Sub ThrowError(ByVal errorCode As Long, ByVal messageText As String)
    Err.Raise errorCode, MODULE_NAME, messageText
End Sub
