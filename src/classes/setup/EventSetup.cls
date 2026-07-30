VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Workbook-level setup event coordinator"
Option Explicit


'@PredeclaredId
'@Folder("Setup")
'@ModuleDescription("Encapsulates workbook level setup events and caches for high-frequency handlers")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

'@class EventSetup
'@description
'Encapsulates workbook-level setup events and lazily-cached domain managers.
'This class is the single entry point for all workbook event handling in the
'setup workbook. It reacts to worksheet activation, change, and open events,
'manages setup rows (add, insert, delete, sort) through cached helpers
'(LLdictionary, LLChoices, LLExport, Analysis), applies a translation across the
'four content sheets, and provides analysis table value lookups. All domain
'managers are created lazily on first use and cached for the lifetime of the
'instance.
'
'SHEET NAMES ARE DECLARED TWICE ON PURPOSE
'-------------------------------------------------------------------------------
'SetupHelpers.bas declares the same sheet names for the import and clean flow.
'Each file keeps the constants it needs, so neither depends on the other for a
'string. Change one and change the other.
'
'PROTECTION IS APPLIED IN ONE PLACE
'-------------------------------------------------------------------------------
'ProtectSetupSheet is the only routine that decides the protection flags of a
'setup sheet, and BeginSheetUpdate / EndSheetUpdate go through it. The ribbon
'calls it too, so a sheet unlocked on one path and re-locked on the other comes
'back with the flags it started with.
'@depends BetterArray, HiddenNames, Passwords, TranslationObject, UpdatedValues, DropdownLists, LLdictionary, LLVariables, LLChoices, LLExport, Analysis, CustomTable, ChoiceFormula, RibbonDev, ProjectError


'@section State
'===============================================================================
Private Type TEventSetup
    hostBook            As Workbook          'Workbook owning the setup sheets
    updater             As UpdatedValues    'Registry tracker for watched columns
    drop                As DropdownLists    'Cached dropdown manager
    ribbon              As IRibbonUI         'Cached ribbon instance when available
    tsTable             As CustomTable      'Analysis time-series table wrapper
    graphTable          As CustomTable      'Analysis graph table wrapper
    graphLabTable       As CustomTable      'Analysis graph table label wrapper
    spatioTable         As CustomTable      'Analysis spatio-temporal table wrapper
    spatioSpecsTable    As CustomTable     'Spatio-temporal specification table wrapper
    variables           As LLVariables      'Dictionary variable helper
    dictionary          As LLdictionary     'Cached dictionary reference
    choices             As LLChoices        'Cached choices helper
    ana                 As Analysis         'Cached analysis helper
    exports             As LLExport         'Cached exports manager
    passwords           As Passwords        'Cached password helper (lazily created once)
    choicesTable        As CustomTable      'Cached Choices table wrapper
    buttonsOnTranslations As Boolean        'Answer the ribbon buttons were last given
    buttonsAnswered     As Boolean          'Whether that answer has been given once
End Type

Private this As TEventSetup

'@section Constants
'===============================================================================
Private Const SHEET_UPDATED          As String = "__updated"
Private Const SHEET_DICTIONARY       As String = "Dictionary"
Private Const SHEET_CHOICES          As String = "Choices"
Private Const SHEET_ANALYSIS         As String = "Analysis"
Private Const SHEET_EXPORTS          As String = "Exports"
Private Const SHEET_TRANSLATIONS     As String = "Translations"
Private Const SHEET_DROPDOWN         As String = "__variables"
Private Const SHEET_PASSWORDS        As String = "__pass"
Private Const SHEET_CHECKING         As String = "__checkRep"
Private Const COUNTER_NAME           As String = "_SetupTranslationsCounter"

Private Const LIST_TS_DATA           As String = "Tab_TimeSeries_Analysis"
Private Const LIST_GRAPH_TS          As String = "Tab_Graph_TimeSeries"
Private Const LIST_LAB_GRAPH_TS      As String = "Tab_Label_TSGraph"
Private Const LIST_SPATIO_TEMPORAL   As String = "Tab_SpatioTemporal_Analysis"
Private Const LIST_SPATIO_TEMPORAL_SPECS As String = "Tab_SpatioTemporal_Specs"
Private Const TABLE_CHOICES          As String = "Tab_Choices"

Private Const DICT_START_ROW         As Long = 5
Private Const DICT_START_COLUMN      As Long = 1
Private Const CHOICES_START_ROW      As Long = 4
Private Const CHOICES_START_COLUMN   As Long = 1
Private Const EXPORT_START_ROW       As Long = 4
Private Const EXPORT_START_COLUMN    As Long = 1

Private Const KEY_DICTIONARY         As String = "dictionary"
Private Const KEY_CHOICES            As String = "choices"
Private Const KEY_ANALYSIS           As String = "analysis"
Private Const KEY_EXPORTS            As String = "exports"

Private Const DROPDOWN_PREFIX        As String = "dropdown_"
Private Const CLASSNAME              As String = "EventSetup"

'@section Factory helpers
'===============================================================================

'@label:create
'@sub-title Build a configured EventSetup instance bound to a host workbook.
'@details
'Factory method using the predeclared-instance pattern. Creates a new
'EventSetup, calls Configure to bind it to the host workbook and reset all
'internal caches, then returns the initialised object. Raises an error when
'hostBook is Nothing because Configure performs the validation.
'@param hostBook Workbook. The workbook owning the setup sheets.
'@return EventSetup. Configured service instance.
'@throws ProjectError.ObjectNotInitialized When hostBook is Nothing (raised by Configure).
'@export
Public Function Create(ByVal hostBook As Workbook) As EventSetup
    Dim instance As EventSetup

    Set instance = New EventSetup
    instance.Configure hostBook
    Set Create = instance
End Function

'@section Event handling and operations
'===============================================================================
'@description Workbook event handlers and setup row operations.

'@label:configure
'@sub-title Configure the service with its host workbook context.
'@details
'Stores the host workbook reference and resets every lazily-cached domain
'manager so subsequent calls rebuild them from the new workbook. Must be
'called before any event handler or row operation. Raises an error when the
'supplied workbook is Nothing.
'@param hostBook Workbook. The workbook providing setup worksheets and infrastructure.
'@throws ProjectError.ObjectNotInitialized When hostBook is Nothing.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
    If (hostBook Is Nothing) Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook is required to configure EventSetup"
    End If

    Set this.hostBook = hostBook
    ResetInternalCaches
End Sub

'@label:on-workbook-open
'@sub-title React to workbook open notifications.
'@details
'Switches update watcher tags to "yes" so tracked columns are marked as
'needing recalculation, resets the translation counter hidden name to zero,
'disables Application.FormatStaleValues to suppress recalculation prompts,
'and applies the full password protection matrix across all setup sheets.
'Silently skips each step when the required helper cannot be created.
'@export
Public Sub OnWorkbookOpen()
    EnsureUpdatedValues
    If Not (this.updater Is Nothing) Then
        On Error Resume Next
            this.updater.SwitchTagsToYes
        On Error GoTo 0
    End If

    ResetTranslationCounter

    On Error Resume Next
        Application.FormatStaleValues = False
    On Error GoTo 0

    PasswordManager.ApplyProtectionMatrix this.hostBook
End Sub

'@label:on-sheet-activate
'@sub-title Handle worksheet activation events.
'@details
'Refreshes the three ribbon controls whose visibility depends on the active
'sheet, and only when the answer they were last given has changed. When the
'Analysis sheet becomes active, primes the analysis table wrappers and asks the
'update watcher whether the dropdowns need rebuilding.
'@param sh Worksheet. The worksheet that became active.
'@export
Public Sub OnSheetActivate(ByVal sh As Worksheet)
    If (sh Is Nothing) Then Exit Sub

    RefreshSheetButtons sh

    If StrComp(sh.Name, SHEET_ANALYSIS, vbTextCompare) = 0 Then
        'Pre-init analysis tables so first SheetChange doesn't pay init cost
        EnsureAnalysisTable this.tsTable, LIST_TS_DATA, "title"
        EnsureAnalysisTable this.graphTable, LIST_GRAPH_TS, "series title"
        EnsureAnalysisTable this.spatioTable, LIST_SPATIO_TEMPORAL, "section"
        UpdateAnalysisDropdowns
    End If
End Sub

'@label:refresh-sheet-buttons
'@sub-title Invalidate the three sheet-dependent ribbon controls when their answer changes.
'@details
'SetupButtonVisible answers one question for all three controls: whether the
'active sheet is Translations. Invalidating them on every activation asks the
'ribbon to re-read an answer it already has, so the answer is remembered here
'and the three calls are made only when the Translations boundary is crossed.
'The first activation of a session always refreshes, because no answer has been
'given yet.
'@param sh Worksheet. The worksheet that became active.
Private Sub RefreshSheetButtons(ByVal sh As Worksheet)
    Dim onTranslations As Boolean

    onTranslations = (StrComp(sh.Name, SHEET_TRANSLATIONS, vbTextCompare) = 0)

    If this.buttonsAnswered Then
        If onTranslations = this.buttonsOnTranslations Then Exit Sub
    End If

    this.buttonsOnTranslations = onTranslations
    this.buttonsAnswered = True

    EnsureRibbon
    If this.ribbon Is Nothing Then Exit Sub

    this.ribbon.InvalidateControl "btnDelLoRow"
    this.ribbon.InvalidateControl "btnDelLoCol"
    this.ribbon.InvalidateControl "btnSort"
End Sub

'@label:on-sheet-change
'@sub-title Handle worksheet change events fired by Excel.
'@details
'Forwards every change to the update watcher so tracked columns are flagged.
'When the Choices sheet changes, recalculates label formulas via
'HandleChoicesChange and exits. When the Analysis sheet changes, ensures the
'analysis cache tables are initialised, recalculates them, and applies choices
'and geo dropdown validation to the edited cell when it falls inside a
'relevant ListObject.
'@param sh Worksheet. The worksheet raising the change event.
'@param target Range. The edited cells.
'@export
Public Sub OnSheetChange(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    EnsureUpdatedValues
    If Not (this.updater Is Nothing) Then
        this.updater.CheckUpdate sh, target
    End If

    If StrComp(sh.Name, SHEET_CHOICES, vbTextCompare) = 0 Then
        HandleChoicesChange target
        Exit Sub
    End If

    If StrComp(sh.Name, SHEET_ANALYSIS, vbTextCompare) <> 0 Then Exit Sub

    EnsureAnalysisTable this.tsTable, LIST_TS_DATA, "title"
    EnsureAnalysisTable this.graphTable, LIST_GRAPH_TS, "series title"
    EnsureAnalysisTable this.spatioTable, LIST_SPATIO_TEMPORAL, "section"

    If target.ListObject Is Nothing Then Exit Sub

    CalculateAnalysis target

    If target.Rows.Count > 1 Then Exit Sub
    If target.Columns.Count > 1 Then Exit Sub

    AddChoicesDropdown target
    AddGeoDropdown target
    SyncGeoFromSpecs target
End Sub

'@label:update-analysis-dropdowns
'@sub-title Refresh cached dropdowns and related analysis helpers.
'@details
'Conditionally rebuilds dropdown lists used by the Analysis sheet. When
'forceUpdate is True or when the update watcher reports that "control details"
'or "variable name" columns changed, refreshes geo, choices, and special-
'variable dropdowns. When "variable type" or "variable name" changed, refreshes
'the time-variable dropdown. The keys are the Dictionary header names as the
'sheet carries them. It then tells the watcher what it has acted on, and
'invalidates the LLVariables cache so subsequent lookups reflect the latest
'dictionary state.
'@param forceUpdate Optional Boolean. When True, forces a full refresh regardless of update tags. Defaults to False.
'@export
Public Sub UpdateAnalysisDropdowns(Optional ByVal forceUpdate As Boolean = False)
    Dim refreshVariables As Boolean
    Dim refreshTimeVariables As Boolean

    EnsureUpdatedValues
    If (this.updater Is Nothing) Then Exit Sub

    If Not EnsureDictionary() Then Exit Sub

    EnsureDropdownValues
    If this.drop Is Nothing Then Exit Sub

    'VBA runs both sides of Or, so each question is asked in its own statement
    'and a forced refresh asks the watcher nothing at all.
    refreshVariables = forceUpdate
    If Not refreshVariables Then refreshVariables = this.updater.IsUpdated("control details")
    If Not refreshVariables Then refreshVariables = this.updater.IsUpdated("variable name")

    refreshTimeVariables = forceUpdate
    If Not refreshTimeVariables Then refreshTimeVariables = this.updater.IsUpdated("variable type")
    If Not refreshTimeVariables Then refreshTimeVariables = this.updater.IsUpdated("variable name")

    If refreshVariables Then
        AddOrUpdateDrop this.dictionary.GeoVars, "__geo_vars"
        AddOrUpdateDrop this.dictionary.ChoicesVars, "__choice_vars"
        AddOrUpdateDrop this.dictionary.SpecialVars("hf"), "__hfonly_vars"
        AddOrUpdateDrop this.dictionary.SpecialVars("geo"), "__geoonly_vars"
    End If

    If refreshTimeVariables Then
        AddOrUpdateDrop this.dictionary.TimeVars, "__time_vars"
    End If

    'Every flag that reads "yes" now has been acted on, so the next activation
    'rebuilds only what the user has edited since.
    this.updater.AcknowledgeUpdates

    EnsureVars
    If Not this.variables Is Nothing Then
        this.variables.InvalidateCaches
    End If

End Sub

'@label:reset-caches
'@sub-title Clear cached resources so they are lazily rebuilt on demand.
'@export
Public Sub ResetCaches()
    ResetInternalCaches
End Sub

'@label:recalculate-analysis
'@sub-title Recalculate analysis tables in response to structural edits.
'@export
Public Sub RecalculateAnalysis()
    CalculateAnalysis
End Sub

'@label:build-time-series-header
'@sub-title Compose the time-series header displayed in analysis tables.
'@export
Public Function BuildTimeSeriesHeader(ByVal timeVar As String, _
                                      ByVal groupVar As String, _
                                      ByVal sumLabel As String) As String
    BuildTimeSeriesHeader = BuildTimeSeriesHeaderInternal(timeVar, groupVar, sumLabel)
End Function

'@label:analysis-graph-value
'@sub-title Retrieve a value from the analysis graph label table.
'@export
Public Function AnalysisGraphValue(ByVal graphTitle As String, _
                                   Optional ByVal graphCol As String = "Graph ID") As String
    AnalysisGraphValue = AnalysisGraphValueInternal(graphTitle, graphCol)
End Function

'@label:analysis-time-series-value
'@sub-title Retrieve a value from the analysis time-series table.
'@export
Public Function AnalysisTimeSeriesValue(ByVal seriesTitle As String, _
                                        Optional ByVal tsColumn As String = "Series ID") As String
    AnalysisTimeSeriesValue = AnalysisTimeSeriesValueInternal(seriesTitle, tsColumn)
End Function

'@label:spatio-temporal-spec-value
'@sub-title Retrieve a value from the spatio-temporal specifications table.
'@export
Public Function SpatioTemporalSpecValue(ByVal sectionName As String, _
                                        Optional ByVal specColumn As String = "N geo max") As String
    SpatioTemporalSpecValue = SpatioTemporalSpecValueInternal(sectionName, specColumn)
End Function

'@section Row management and sorting
'===============================================================================

'@label:manage-rows
'@sub-title Manage setup worksheet rows using cached helpers.
'@details
'Delegates to the appropriate domain manager (dictionary, choices, analysis or
'exports) to add or remove rows. Each branch lazily creates
'its helper, unlocks the target sheet, invokes the manager, then re-locks the
'sheet in a cleanup block. For exports, the dictionary sheet is also unlocked
'because export row changes can propagate to dictionary columns. Silently exits
'when the sheet name does not match a known key.
'@param sheetName String. Worksheet name to process.
'@param del Optional Boolean. When True, removes rows instead of adding them. Defaults to False.
'@throws ProjectError.SomethingWentWrong When the underlying domain manager operation fails.
'@export
Public Sub ManageRows(ByVal sheetName As String, Optional ByVal del As Boolean = False)
    Dim normalized As String
    Dim sheetPass As Passwords
    Dim dictPass As Passwords
    Dim unlockedSheet As String
    Dim unlockedDict As String

    normalized = NormalisedSheetName(sheetName)
    If LenB(normalized) = 0 Then Exit Sub

    On Error GoTo CleanFail

    Select Case normalized
        Case KEY_DICTIONARY
            If Not EnsureDictionary() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_DICTIONARY)
            unlockedSheet = SHEET_DICTIONARY
            this.dictionary.ManageRows del
        Case KEY_CHOICES
            If Not EnsureChoices() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_CHOICES)
            unlockedSheet = SHEET_CHOICES
            this.choices.ManageRows del
        Case KEY_ANALYSIS
            If Not EnsureAnalysisManager() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_ANALYSIS)
            unlockedSheet = SHEET_ANALYSIS
            this.ana.ManageRows del
        Case KEY_EXPORTS
            If Not EnsureExports() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_EXPORTS)
            unlockedSheet = SHEET_EXPORTS
            If EnsureDictionary() And Not this.dictionary Is Nothing Then
                Set dictPass = BeginSheetUpdate(SHEET_DICTIONARY)
                unlockedDict = SHEET_DICTIONARY
            End If
            this.exports.ManageRows del, this.dictionary
        Case Else
            GoTo CleanExit
    End Select

CleanExit:
    If LenB(unlockedDict) > 0 Then
        EndSheetUpdate dictPass, unlockedDict
        Set dictPass = Nothing
        unlockedDict = vbNullString
    End If

    If LenB(unlockedSheet) > 0 Then
        EndSheetUpdate sheetPass, unlockedSheet
        Set sheetPass = Nothing
        unlockedSheet = vbNullString
    End If

    Exit Sub

CleanFail:
    On Error Resume Next
        If LenB(unlockedDict) > 0 Then EndSheetUpdate dictPass, unlockedDict
        If LenB(unlockedSheet) > 0 Then EndSheetUpdate sheetPass, unlockedSheet
    On Error GoTo 0

    ThrowError ProjectError.SomethingWentWrong, "Failed to manage rows on '" & sheetName & "': " & Err.Description
End Sub

'@label:insert-rows
'@sub-title Insert rows relative to a selection for supported setup sheets.
'@details
'Delegates to the appropriate domain manager to insert rows at the position
'indicated by targetCell. Handles sheet protection around the operation. For
'the Analysis sheet, insertShift is always forced to True to preserve stacked
'tables. For exports, also unlocks the dictionary sheet because insert results
'may propagate to dictionary columns. Silently exits for unrecognised sheets.
'@param sheetName String. Worksheet name identifying which helper should process the insertion.
'@param targetCell Range. The worksheet selection whose height dictates the number of rows to insert.
'@param insertShift Optional Boolean. When True, worksheet rows are inserted to protect stacked tables. Defaults to False.
'@throws ProjectError.SomethingWentWrong When the underlying insert operation fails.
'@export
Public Sub InsertRows(ByVal sheetName As String, _
                      ByVal targetCell As Range, _
                      Optional ByVal insertShift As Boolean = False)
    Dim normalized As String
    Dim sheetPass As Passwords
    Dim dictPass As Passwords
    Dim unlockedSheet As String
    Dim unlockedDict As String
    Dim effectiveShift As Boolean

    normalized = NormalisedSheetName(sheetName)
    If LenB(normalized) = 0 Then Exit Sub
    If targetCell Is Nothing Then Exit Sub

    On Error GoTo CleanFail

    Select Case normalized
        Case KEY_DICTIONARY
            If Not EnsureDictionary() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_DICTIONARY)
            unlockedSheet = SHEET_DICTIONARY
            this.dictionary.InsertRows targetCell, insertShift
        Case KEY_CHOICES
            If Not EnsureChoices() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_CHOICES)
            unlockedSheet = SHEET_CHOICES
            this.choices.InsertRows targetCell, insertShift
        Case KEY_ANALYSIS
            If Not EnsureAnalysisManager() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_ANALYSIS)
            unlockedSheet = SHEET_ANALYSIS
            effectiveShift = True
            this.ana.InsertRows targetCell, effectiveShift
        Case KEY_EXPORTS
            If Not EnsureExports() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_EXPORTS)
            unlockedSheet = SHEET_EXPORTS
            If EnsureDictionary() And Not this.dictionary Is Nothing Then
                Set dictPass = BeginSheetUpdate(SHEET_DICTIONARY)
                unlockedDict = SHEET_DICTIONARY
            End If
            this.exports.InsertRows targetCell, this.dictionary, insertShift
        Case Else
            GoTo CleanExit
    End Select

CleanExit:
    If LenB(unlockedDict) > 0 Then
        EndSheetUpdate dictPass, unlockedDict
        Set dictPass = Nothing
        unlockedDict = vbNullString
    End If

    If LenB(unlockedSheet) > 0 Then
        EndSheetUpdate sheetPass, unlockedSheet
        Set sheetPass = Nothing
        unlockedSheet = vbNullString
    End If

    Exit Sub

CleanFail:
    On Error Resume Next
        If LenB(unlockedDict) > 0 Then EndSheetUpdate dictPass, unlockedDict
        If LenB(unlockedSheet) > 0 Then EndSheetUpdate sheetPass, unlockedSheet
    On Error GoTo 0

    ThrowError ProjectError.SomethingWentWrong, "Failed to insert rows on '" & sheetName & "': " & Err.Description
End Sub

'@label:delete-rows
'@sub-title Delete rows relative to a selection for supported setup sheets.
'@details
'Delegates to the appropriate domain manager to delete rows at the position
'indicated by targetCell. Handles sheet protection around the operation. For
'the Analysis sheet, forceShift is always forced to True. For exports, also
'unlocks the dictionary sheet because deletion results may propagate to
'dictionary columns. Silently exits for unrecognised sheets.
'@param sheetName String. Worksheet name identifying which helper should process the deletion.
'@param targetCell Range. The worksheet selection to remove.
'@param forceShift Optional Boolean. When True, worksheet rows are deleted to preserve stacked tables. Defaults to False.
'@throws ProjectError.SomethingWentWrong When the underlying delete operation fails.
'@export
Public Sub DeleteRows(ByVal sheetName As String, _
                      ByVal targetCell As Range, _
                      Optional ByVal forceShift As Boolean = False)
    Dim normalized As String
    Dim sheetPass As Passwords
    Dim dictPass As Passwords
    Dim unlockedSheet As String
    Dim unlockedDict As String

    normalized = NormalisedSheetName(sheetName)
    If LenB(normalized) = 0 Then Exit Sub
    If targetCell Is Nothing Then Exit Sub

    On Error GoTo CleanFail

    Select Case normalized
        Case KEY_DICTIONARY
            If Not EnsureDictionary() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_DICTIONARY)
            unlockedSheet = SHEET_DICTIONARY
            this.dictionary.DeleteRows targetCell
        Case KEY_CHOICES
            If Not EnsureChoices() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_CHOICES)
            unlockedSheet = SHEET_CHOICES
            this.choices.DeleteRows targetCell
        Case KEY_ANALYSIS
            If Not EnsureAnalysisManager() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_ANALYSIS)
            unlockedSheet = SHEET_ANALYSIS
            this.ana.DeleteRows targetCell, forceShift:=True
        Case KEY_EXPORTS
            If Not EnsureExports() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_EXPORTS)
            unlockedSheet = SHEET_EXPORTS
            Dim dictRef As LLdictionary
            If EnsureDictionary() And Not this.dictionary Is Nothing Then
                Set dictPass = BeginSheetUpdate(SHEET_DICTIONARY)
                unlockedDict = SHEET_DICTIONARY
                Set dictRef = this.dictionary
            End If
            this.exports.DeleteRows targetCell, dictRef, forceShift
        Case Else
            GoTo CleanExit
    End Select

CleanExit:
    If LenB(unlockedDict) > 0 Then
        EndSheetUpdate dictPass, unlockedDict
        Set dictPass = Nothing
        unlockedDict = vbNullString
    End If

    If LenB(unlockedSheet) > 0 Then
        EndSheetUpdate sheetPass, unlockedSheet
        Set sheetPass = Nothing
        unlockedSheet = vbNullString
    End If

    Exit Sub

CleanFail:
    On Error Resume Next
        If LenB(unlockedDict) > 0 Then EndSheetUpdate dictPass, unlockedDict
        If LenB(unlockedSheet) > 0 Then EndSheetUpdate sheetPass, unlockedSheet
    On Error GoTo 0

    ThrowError ProjectError.SomethingWentWrong, "Failed to delete rows on '" & sheetName & "': " & Err.Description
End Sub

'@label:sort-tables
'@sub-title Sort setup worksheets leveraging cached managers.
'@details
'Delegates to the appropriate domain manager to sort the worksheet tables.
'Handles sheet protection around the operation. For exports, also unlocks
'the dictionary sheet because sort results may propagate to dictionary
'columns. Silently exits for unrecognised sheets.
'@param sheetName String. Worksheet name to sort.
'@throws ProjectError.SomethingWentWrong When the sort operation fails.
'@export
Public Sub SortTables(ByVal sheetName As String)
    Dim normalized As String
    Dim sheetPass As Passwords
    Dim dictPass As Passwords
    Dim unlockedSheet As String
    Dim unlockedDict As String

    normalized = NormalisedSheetName(sheetName)
    If LenB(normalized) = 0 Then Exit Sub

    On Error GoTo CleanFail

    Select Case normalized
        Case KEY_DICTIONARY
            If Not EnsureDictionary() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_DICTIONARY)
            unlockedSheet = SHEET_DICTIONARY
            this.dictionary.Sort
        Case KEY_CHOICES
            If Not EnsureChoices() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_CHOICES)
            unlockedSheet = SHEET_CHOICES
            this.choices.Sort
        Case KEY_ANALYSIS
            If Not EnsureAnalysisManager() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_ANALYSIS)
            unlockedSheet = SHEET_ANALYSIS
            this.ana.Sort
        Case KEY_EXPORTS
            If Not EnsureExports() Then GoTo CleanExit
            Set sheetPass = BeginSheetUpdate(SHEET_EXPORTS)
            unlockedSheet = SHEET_EXPORTS
            Dim dictRef As LLdictionary
            If EnsureDictionary() And Not this.dictionary Is Nothing Then
                Set dictPass = BeginSheetUpdate(SHEET_DICTIONARY)
                unlockedDict = SHEET_DICTIONARY
                Set dictRef = this.dictionary
            End If
            this.exports.Sort dictRef
        Case Else
            GoTo CleanExit
    End Select

CleanExit:
    If LenB(unlockedDict) > 0 Then
        EndSheetUpdate dictPass, unlockedDict
        Set dictPass = Nothing
        unlockedDict = vbNullString
    End If

    If LenB(unlockedSheet) > 0 Then
        EndSheetUpdate sheetPass, unlockedSheet
        Set sheetPass = Nothing
        unlockedSheet = vbNullString
    End If

    Exit Sub

CleanFail:
    On Error Resume Next
        If LenB(unlockedDict) > 0 Then EndSheetUpdate dictPass, unlockedDict
        If LenB(unlockedSheet) > 0 Then EndSheetUpdate sheetPass, unlockedSheet
    On Error GoTo 0

    ThrowError ProjectError.SomethingWentWrong, "Failed to sort sheet '" & sheetName & "': " & Err.Description
End Sub

'@section Translations
'===============================================================================

'@label:apply-setup-translation
'@sub-title Rewrite every header and label on the four content sheets.
'@details
'Translates the Dictionary, Choices, Analysis and Exports sheets through the
'managers this service already holds, then drops every cache.
'
'THE CACHE DROP IS THE POINT
'-------------------------------------------------------------------------------
'A translation renames the columns the cached managers were built against. A row
'insert, a sort or a dropdown refresh made after it would look for a column name
'the sheet no longer carries, and the lazy getters answer Nothing on a failure,
'so nothing would report it. Dropping the caches here makes the next caller
'rebuild against the sheet as it now reads.
'
'Each sheet is unlocked, translated and locked again on its own, so a failure
'leaves no sheet unlocked. The failure is raised to the caller.
'@param translator TranslationObject. Label source for the target language.
'@throws ProjectError.ObjectNotInitialized When translator is Nothing.
'@throws ProjectError.SomethingWentWrong When a sheet translation fails.
'@export
Public Sub ApplySetupTranslation(ByVal translator As TranslationObject)
    Dim unlockedSheet As String
    Dim failure As String

    If translator Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "A translator is required to translate the setup"
    End If

    On Error GoTo CleanFail

    If EnsureDictionary() Then
        unlockedSheet = SHEET_DICTIONARY
        UnprotectSetupSheet unlockedSheet
        this.dictionary.Translate translator
        ProtectSetupSheet unlockedSheet
        unlockedSheet = vbNullString
    End If

    If EnsureChoices() Then
        unlockedSheet = SHEET_CHOICES
        UnprotectSetupSheet unlockedSheet
        this.choices.Translate translator
        ProtectSetupSheet unlockedSheet
        unlockedSheet = vbNullString
    End If

    If EnsureAnalysisManager() Then
        unlockedSheet = SHEET_ANALYSIS
        UnprotectSetupSheet unlockedSheet
        this.ana.Translate translator
        ProtectSetupSheet unlockedSheet
        unlockedSheet = vbNullString
    End If

    If EnsureExports() Then
        unlockedSheet = SHEET_EXPORTS
        UnprotectSetupSheet unlockedSheet
        this.exports.Translate translator
        ProtectSetupSheet unlockedSheet
        unlockedSheet = vbNullString
    End If

    ResetInternalCaches
    Exit Sub

CleanFail:
    failure = Err.Description

    'Neither call below can raise, so the recovery finishes inside the handler.
    'ProtectSetupSheet swallows a protect failure and ResetInternalCaches only
    'drops references.
    If LenB(unlockedSheet) > 0 Then ProtectSetupSheet unlockedSheet
    ResetInternalCaches

    ThrowError ProjectError.SomethingWentWrong, "Failed to translate the setup: " & failure
End Sub

'@section Internal helpers
'===============================================================================

'@label:reset-internal-caches
'@sub-title Clear all lazily-cached domain managers and table helpers.
'@details
'Sets every cached reference to Nothing so the next access triggers lazy
'creation. Called by Configure, by ResetCaches and at the end of
'ApplySetupTranslation. The password helper and the Choices table wrapper are
'cleared here too: Configure binds the service to another workbook, and a
'wrapper built against the old one answers for the wrong sheet.
Private Sub ResetInternalCaches()
    Set this.updater = Nothing
    Set this.drop = Nothing
    Set this.ribbon = Nothing
    Set this.tsTable = Nothing
    Set this.graphTable = Nothing
    Set this.graphLabTable = Nothing
    Set this.spatioTable = Nothing
    Set this.spatioSpecsTable = Nothing
    Set this.variables = Nothing
    Set this.dictionary = Nothing
    Set this.choices = Nothing
    Set this.ana = Nothing
    Set this.exports = Nothing
    Set this.passwords = Nothing
    Set this.choicesTable = Nothing

    this.buttonsOnTranslations = False
    this.buttonsAnswered = False
End Sub

'@label:add-or-update-drop
'@sub-title Add or update a dropdown list entry through the cached manager.
'@details
'Adds a new dropdown list entry when the name does not already exist, or
'updates the existing entry. Ensures the dropdown manager is initialised
'before attempting the operation.
'@param lst BetterArray. The dropdown values to store.
'@param dropName String. Name of the dropdown list entry.
Private Sub AddOrUpdateDrop(ByVal lst As BetterArray, ByVal dropName As String)
    EnsureDropdownValues
    If this.drop Is Nothing Then Exit Sub

    If Not this.drop.Exists(dropName) Then
        this.drop.Add lst, dropName
    Else
        this.drop.Update lst, dropName
    End If
End Sub

'@label:ensure-updated-values
'@sub-title Lazily initialise the update watcher from the registry worksheet.
'@details
'Creates an UpdatedValues instance from the "__updated" registry worksheet.
'Exits silently when the worksheet does not exist or creation fails, so
'callers must always check this.updater before use.
Private Sub EnsureUpdatedValues()
    If Not (this.updater Is Nothing) Then Exit Sub

    Dim registrySheet As Worksheet

    Set registrySheet = SheetOrNothing(SHEET_UPDATED)
    If registrySheet Is Nothing Then Exit Sub

    On Error Resume Next
        Set this.updater = UpdatedValues.Create(registrySheet)
    On Error GoTo 0
End Sub

'@label:ensure-dropdown-values
'@sub-title Lazily initialise the dropdown manager from the variables worksheet.
'@details
'Creates an DropdownLists instance from the "__variables" worksheet using
'the standard dropdown prefix. Exits silently when the worksheet does not
'exist or creation fails.
Private Sub EnsureDropdownValues()
    If Not (this.drop Is Nothing) Then Exit Sub

    Dim dropSheet As Worksheet

    Set dropSheet = SheetOrNothing(SHEET_DROPDOWN)
    If (dropSheet Is Nothing) Then Exit Sub

    On Error Resume Next
        Set this.drop = DropdownLists.Create(dropSheet, DROPDOWN_PREFIX)
    On Error GoTo 0
End Sub

'@label:ensure-ribbon
'@sub-title Lazily cache the ribbon instance.
Private Sub EnsureRibbon()
    If Not (this.ribbon Is Nothing) Then Exit Sub
    Set this.ribbon = RibbonDev.ActualRibbon()
End Sub

'@label:ensure-dictionary
'@sub-title Lazily create the dictionary manager.
'@details
'Creates an LLdictionary instance from the "Dictionary" worksheet when not
'already cached. Returns True when the dictionary is available so callers
'can guard subsequent operations.
'@return Boolean. True when the dictionary is available.
Private Function EnsureDictionary() As Boolean
    If Not (this.dictionary Is Nothing) Then
        EnsureDictionary = True
        Exit Function
    End If

    Dim dictSheet As Worksheet

    Set dictSheet = SheetOrNothing(SHEET_DICTIONARY)
    If dictSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set this.dictionary = LLdictionary.Create(dictSheet, DICT_START_ROW, DICT_START_COLUMN)
    On Error GoTo 0
    EnsureDictionary = Not (this.dictionary Is Nothing)
End Function

'@label:ensure-vars
'@sub-title Lazily initialise the variables helper from the cached dictionary.
'@details
'Creates an LLVariables instance from the cached dictionary reference.
'Requires EnsureDictionary to succeed first; exits silently when the
'dictionary is unavailable.
Private Sub EnsureVars()
    If Not (this.variables Is Nothing) Then Exit Sub

    If Not EnsureDictionary() Then Exit Sub

    On Error Resume Next
        Set this.variables = LLVariables.Create(this.dictionary)
    On Error GoTo 0
End Sub

'@label:ensure-choices
'@sub-title Lazily create the choices manager.
'@details
'Creates an LLChoices instance from the "Choices" worksheet when not already
'cached. Returns True when the choices helper is available.
'@return Boolean. True when the choices helper is available.
Private Function EnsureChoices() As Boolean
    If Not (this.choices Is Nothing) Then
        EnsureChoices = True
        Exit Function
    End If

    Dim choiceSheet As Worksheet

    Set choiceSheet = SheetOrNothing(SHEET_CHOICES)
    If choiceSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set this.choices = LLChoices.Create(choiceSheet, CHOICES_START_ROW, CHOICES_START_COLUMN)
    On Error GoTo 0

    EnsureChoices = Not (this.choices Is Nothing)
End Function

'@label:ensure-analysis-manager
'@sub-title Lazily create the analysis manager.
'@details
'Creates an Analysis instance from the "Analysis" worksheet when not already
'cached. Returns True when the analysis helper is available.
'@return Boolean. True when the analysis helper is available.
Private Function EnsureAnalysisManager() As Boolean
    If Not (this.ana Is Nothing) Then
        EnsureAnalysisManager = True
        Exit Function
    End If

    Dim analysisSheet As Worksheet

    Set analysisSheet = SheetOrNothing(SHEET_ANALYSIS)
    If analysisSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set this.ana = Analysis.Create(analysisSheet)
    On Error GoTo 0

    EnsureAnalysisManager = Not (this.ana Is Nothing)
End Function

'@label:ensure-exports
'@sub-title Lazily create the exports manager.
'@details
'Creates an LLExport instance from the "Exports" worksheet when not already
'cached. Returns True when the exports helper is available.
'@return Boolean. True when the exports helper is available.
Private Function EnsureExports() As Boolean
    If Not (this.exports Is Nothing) Then
        EnsureExports = True
        Exit Function
    End If

    Dim exportSheet As Worksheet

    Set exportSheet = SheetOrNothing(SHEET_EXPORTS)
    If exportSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set this.exports = LLExport.Create(exportSheet, EXPORT_START_ROW, EXPORT_START_COLUMN)
    On Error GoTo 0

    EnsureExports = Not (this.exports Is Nothing)
End Function

'@label:ensure-analysis-table
'@sub-title Ensure the cached analysis tables are created lazily and reused.
'@details
'Creates a CustomTable wrapper around the named ListObject on the Analysis
'sheet. The ByRef tableRef parameter is populated on success and reused on
'subsequent calls. Returns True when the table is available.
'@param tableRef ByRef CustomTable. Cached table reference, populated on first call and reused.
'@param listName String. Name of the ListObject on the Analysis worksheet.
'@param keyColumn String. Column to use as the table key.
'@return Boolean. True when the table is available.
Private Function EnsureAnalysisTable(ByRef tableRef As CustomTable, _
                                     ByVal listName As String, _
                                     ByVal keyColumn As String) As Boolean
    If Not (tableRef Is Nothing) Then
        EnsureAnalysisTable = True
        Exit Function
    End If

    Dim lo As ListObject
    Set lo = AnalysisListObject(listName)
    If lo Is Nothing Then Exit Function

    On Error Resume Next
        Set tableRef = CustomTable.Create(lo, keyColumn)
    On Error GoTo 0

    EnsureAnalysisTable = Not (tableRef Is Nothing)
End Function

'@label:ensure-table-calculated
'@sub-title Calculate the associated ListObject once the cache is primed.
'@details
'Ensures the analysis table is initialised via EnsureAnalysisTable, then
'recalculates its data range. Silently exits when the table cannot be created.
'@param tableRef ByRef CustomTable. Cached table reference.
'@param listName String. Name of the ListObject on the Analysis worksheet.
'@param keyColumn String. Column to use as the table key.
Private Sub EnsureTableCalculated(ByRef tableRef As CustomTable, _
                                  ByVal listName As String, _
                                  ByVal keyColumn As String)
    If Not EnsureAnalysisTable(tableRef, listName, keyColumn) Then Exit Sub

    On Error Resume Next
        tableRef.DataRange.Calculate
    On Error GoTo 0
End Sub

'@label:analysis-list-object
'@sub-title Retrieve a ListObject from the Analysis worksheet by name.
'@param loName String. Name of the ListObject to retrieve.
'@return ListObject. The ListObject, or Nothing when not found.
Private Function AnalysisListObject(ByVal loName As String) As ListObject
    Dim analysisSheet As Worksheet
    Dim lo As ListObject

    Set analysisSheet = SheetOrNothing(SHEET_ANALYSIS)
    If analysisSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set lo = analysisSheet.ListObjects(loName)
    On Error GoTo 0

    Set AnalysisListObject = lo
End Function

'@label:sheet-or-nothing
'@sub-title Retrieve a worksheet from the host workbook or Nothing.
'@param sheetName String. Name of the worksheet to retrieve.
'@return Worksheet. The worksheet, or Nothing when the workbook is unset or the sheet does not exist.
Private Function SheetOrNothing(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    If this.hostBook Is Nothing Then Exit Function

    On Error Resume Next
        Set sh = this.hostBook.Worksheets(sheetName)
    On Error GoTo 0

    Set SheetOrNothing = sh
End Function

'@label:normalised-sheet-name
'@sub-title Normalise a sheet name to lowercase for comparison.
'@param sheetName String. Raw sheet name to normalise.
'@return String. Lowercase trimmed sheet name.
Private Function NormalisedSheetName(ByVal sheetName As String) As String
    NormalisedSheetName = LCase$(Trim$(sheetName))
End Function

'@label:password-manager
'@sub-title Build a password helper from the passwords worksheet.
'@details
'Creates an Passwords helper from the "__pass" worksheet. Returns Nothing
'when the sheet does not exist or creation fails.
'@return Passwords. The password helper, or Nothing when unavailable.
Private Function PasswordManager() As Passwords
    Dim passSheet As Worksheet

    Set passSheet = SheetOrNothing(SHEET_PASSWORDS)
    If passSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set PasswordManager = Passwords.Create(passSheet)
    On Error GoTo 0
End Function

'@label:cached-password-manager
'@sub-title Return the cached password helper, creating it on first call.
'@details
'Lazily creates the Passwords instance once and caches it for the lifetime
'of this EventSetup instance. Eliminates repeated Passwords.Create calls
'during BeginSheetUpdate / EndSheetUpdate cycles.
'@return Passwords. The cached password helper, or Nothing when unavailable.
Private Function CachedPasswordManager() As Passwords
    If this.passwords Is Nothing Then
        Set this.passwords = PasswordManager()
    End If
    Set CachedPasswordManager = this.passwords
End Function

'@label:unprotect-setup-sheet
'@sub-title Unlock a setup sheet through the cached password helper.
'@details
'The ribbon and the import flow call this, and so does BeginSheetUpdate, so a
'sheet is always unlocked the same way.
'@param sheetName String. Name of the sheet to unprotect.
'@export
Public Sub UnprotectSetupSheet(ByVal sheetName As String)
    Dim pass As Passwords

    If LenB(sheetName) = 0 Then Exit Sub

    Set pass = CachedPasswordManager()
    If pass Is Nothing Then Exit Sub

    On Error Resume Next
        pass.UnProtect sheetName
    On Error GoTo 0
End Sub

'@label:protect-setup-sheet
'@sub-title Lock a setup sheet with the flags that sheet is entitled to.
'@details
'This is the one place the protection flags of a setup sheet are decided.
'Translations and Analysis are locked with row deletion off, because both carry
'stacked tables a deleted worksheet row would tear apart. Every other sheet
'keeps row deletion. The checking report sheet is left alone: SetupErrors writes
'to it while events are live.
'
'Passwords.Protect folds the flag it is given with the flag already stored for
'the sheet and never widens it back, so two callers passing two different
'answers leave the sheet with whichever answer came first. One rule here is what
'stops that.
'@param sheetName String. Name of the sheet to re-protect.
'@export
Public Sub ProtectSetupSheet(ByVal sheetName As String)
    Dim pass As Passwords
    Dim allowRowDelete As Boolean

    If LenB(sheetName) = 0 Then Exit Sub
    If StrComp(sheetName, SHEET_CHECKING, vbTextCompare) = 0 Then Exit Sub

    Set pass = CachedPasswordManager()
    If pass Is Nothing Then Exit Sub

    allowRowDelete = Not (StrComp(sheetName, SHEET_TRANSLATIONS, vbTextCompare) = 0 _
                          Or StrComp(sheetName, SHEET_ANALYSIS, vbTextCompare) = 0)

    On Error Resume Next
        pass.Protect sheetName, allowDeletingRows:=allowRowDelete
    On Error GoTo 0
End Sub

'@label:begin-sheet-update
'@sub-title Temporarily unlock the provided sheet and return the password helper.
'@details
'Removes sheet protection using the cached password helper so callers can modify
'protected cells. Returns the Passwords instance for later re-locking via
'EndSheetUpdate.
'@param sheetName String. Name of the sheet to unprotect.
'@return Passwords. The password helper for later re-locking.
Private Function BeginSheetUpdate(ByVal sheetName As String) As Passwords
    UnprotectSetupSheet sheetName
    Set BeginSheetUpdate = CachedPasswordManager()
End Function

'@label:end-sheet-update
'@sub-title Restore sheet protection previously removed via BeginSheetUpdate.
'@param pass Passwords. Password helper returned by BeginSheetUpdate.
'@param sheetName String. Name of the sheet to re-protect.
Private Sub EndSheetUpdate(ByVal pass As Passwords, ByVal sheetName As String)
    If pass Is Nothing Then Exit Sub
    ProtectSetupSheet sheetName
End Sub

'@label:calculate-analysis
'@sub-title Recalculate analysis tables, optionally targeting a single table.
'@details
'When a target range is provided, identifies which analysis table it belongs
'to via target.ListObject.Name and recalculates only that table (75% savings
'on per-cell edits). When no target is given (structural edits, post-import),
'recalculates all four tables.
'@param target Optional Range. The edited cell(s); Nothing means recalculate all.
Private Sub CalculateAnalysis(Optional ByVal target As Range = Nothing)

    If Not target Is Nothing Then
        Dim lo As ListObject
        Set lo = target.ListObject
        If lo Is Nothing Then Exit Sub

        Select Case lo.Name
            Case LIST_TS_DATA
                EnsureTableCalculated this.tsTable, LIST_TS_DATA, "title"
                EnsureTableCalculated this.graphTable, LIST_GRAPH_TS, "series title"
            Case LIST_GRAPH_TS
                EnsureTableCalculated this.tsTable, LIST_TS_DATA, "title"
            Case LIST_SPATIO_TEMPORAL, LIST_SPATIO_TEMPORAL_SPECS
                EnsureTableCalculated this.spatioTable, LIST_SPATIO_TEMPORAL, "section"
        End Select
        Exit Sub
    End If

    EnsureTableCalculated this.tsTable, LIST_TS_DATA, "title"
    EnsureTableCalculated this.graphTable, LIST_GRAPH_TS, "series title"
    EnsureTableCalculated this.spatioTable, LIST_SPATIO_TEMPORAL, "section"
End Sub

'@label:handle-choices-change
'@sub-title Recalculate label formulas when the Choices sheet changes.
'@details
'Detects when the edited cell falls within the "translated label" column of
'the Tab_Choices table. When it does, recalculates the "label" column so
'dependent formulas stay in sync. Exits silently when the target is outside
'the relevant range.
'@param target Range. The edited cells on the Choices sheet.
Private Sub HandleChoicesChange(ByVal target As Range)
    Dim choiceList As ListObject
    Dim translatedRange As Range
    Dim labelRange As Range

    If target.ListObject Is Nothing Then Exit Sub

    Set choiceList = target.ListObject
    If StrComp(choiceList.Name, TABLE_CHOICES, vbTextCompare) <> 0 Then Exit Sub

    'Lazily cache the Choices table wrapper (avoids CustomTable.Create on every edit)
    If this.choicesTable Is Nothing Then
        Set this.choicesTable = CustomTable.Create(choiceList)
    End If
    Set translatedRange = this.choicesTable.DataRange(colName:="translated label", includeHeaders:=False, strictSearch:=True, matchCase:=False)
    If translatedRange Is Nothing Then Exit Sub
    If Intersect(target, translatedRange) Is Nothing Then Exit Sub

    Set labelRange = this.choicesTable.DataRange(colName:="label", includeHeaders:=False, strictSearch:=True, matchCase:=False)
    If labelRange Is Nothing Then Exit Sub

    On Error Resume Next
        labelRange.Calculate
    On Error GoTo 0
End Sub

'@label:reset-translation-counter-impl
'@sub-title Reset the translation update counter stored in hidden names.
'@details
'Resets the "_SetupTranslationsCounter" hidden name to zero on the registry
'worksheet. Uses HiddenNames to ensure the name exists before setting its
'value, so the operation is safe on fresh workbooks.
Public Sub ResetTranslationCounter()
    Dim registrySheet As Worksheet
    Dim namesManager As HiddenNames

    Set registrySheet = SheetOrNothing(SHEET_UPDATED)
    If registrySheet Is Nothing Then Exit Sub

    On Error Resume Next
        Set namesManager = HiddenNames.Create(registrySheet)
        namesManager.EnsureName COUNTER_NAME, CLng(0), HiddenNameTypeLong
        namesManager.SetValue COUNTER_NAME, CLng(0)
    On Error GoTo 0
End Sub

'@label:build-time-series-header-impl
'@sub-title Compose the time-series header from variable labels and a separator.
'@details
'Concatenates the summary label, time variable label, and group variable label
'using a horizontal-rule character as separator. Resolves variable names to
'their main labels via VariableLabel. Omits segments when the resolved label
'is empty.
'@param timeVar String. Time variable name.
'@param groupVar String. Group variable name.
'@param sumLabel String. Summary label prefix.
'@return String. The composed header text.
Private Function BuildTimeSeriesHeaderInternal(ByVal timeVar As String, _
                                               ByVal groupVar As String, _
                                               ByVal sumLabel As String) As String
    Dim SEP As String
    Dim header As String
    Dim timeLabel As String
    Dim groupLabel As String

    SEP = " " & ChrW(9472) & " "
    header = sumLabel
    timeLabel = VariableLabel(timeVar)
    If LenB(timeLabel) > 0 Then
        header = header & SEP & timeLabel
    End If

    If LenB(groupVar) > 0 Then
        groupLabel = VariableLabel(groupVar)
        If LenB(groupLabel) > 0 Then
            header = header & SEP & groupLabel
        End If
    End If

    BuildTimeSeriesHeaderInternal = header
End Function

'@label:analysis-graph-value-impl
'@sub-title Retrieve a value from the graph time-series table.
'@param graphTitle String. Graph title used as the row key.
'@param graphCol String. Column to read.
'@return String. The retrieved value, or empty.
Private Function AnalysisGraphValueInternal(ByVal graphTitle As String, _
                                             ByVal graphCol As String) As String
    AnalysisGraphValueInternal = AnalysisTableValue(this.graphLabTable, _
                                                    LIST_LAB_GRAPH_TS, _
                                                    "graph title", _
                                                    graphCol, _
                                                    graphTitle)
End Function

'@label:analysis-time-series-value-impl
'@sub-title Retrieve a value from the time-series data table.
'@param seriesTitle String. Series title used as the row key.
'@param tsColumn String. Column to read.
'@return String. The retrieved value, or empty.
Private Function AnalysisTimeSeriesValueInternal(ByVal seriesTitle As String, _
                                                  ByVal tsColumn As String) As String
    AnalysisTimeSeriesValueInternal = AnalysisTableValue(this.tsTable, _
                                                         LIST_TS_DATA, _
                                                         "title", _
                                                         tsColumn, _
                                                         seriesTitle)
End Function

'@label:spatio-temporal-spec-value-impl
'@sub-title Retrieve a value from the spatio-temporal specifications table.
'@param sectionName String. Section name used as the row key.
'@param specColumn String. Column to read.
'@return String. The retrieved value, or empty.
Private Function SpatioTemporalSpecValueInternal(ByVal sectionName As String, _
                                                 ByVal specColumn As String) As String
    SpatioTemporalSpecValueInternal = AnalysisTableValue(this.spatioSpecsTable, _
                                                         LIST_SPATIO_TEMPORAL_SPECS, _
                                                         "section", _
                                                         specColumn, _
                                                         sectionName)
End Function

'@label:variable-label
'@sub-title Resolve the main label for a variable name via the cached dictionary.
'@details
'Looks up "Main Label" in the cached LLVariables helper. Falls back to the
'raw variable name when the helper is unavailable or the lookup returns empty.
'@param varName String. Variable name to look up.
'@return String. The main label, or the raw variable name as fallback.
Private Function VariableLabel(ByVal varName As String) As String
    Dim label As String

    If LenB(varName) = 0 Then Exit Function

    EnsureVars
    If this.variables Is Nothing Then
        VariableLabel = varName
        Exit Function
    End If

    On Error Resume Next
        label = this.variables.Value("Main Label", varName)
    On Error GoTo 0

    If LenB(label) = 0 Then label = varName
    VariableLabel = label
End Function

'@label:safe-table-value
'@sub-title Safely retrieve a value from a CustomTable, returning empty on error.
'@param tableRef CustomTable. The table to query.
'@param columnName String. Column header to read.
'@param keyName String. Key value to look up.
'@return String. The retrieved value, or empty on error.
Private Function SafeTableValue(ByVal tableRef As CustomTable, _
                                ByVal columnName As String, _
                                ByVal keyName As String) As String
    Dim value As Variant

    If tableRef Is Nothing Then Exit Function

    On Error Resume Next
        value = tableRef.Value(colName:=columnName, keyName:=keyName)
    On Error GoTo 0

    If IsError(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    SafeTableValue = CStr(value)
End Function

'@label:analysis-table-value
'@sub-title Shared lookup helper for analysis ListObjects.
'@details
'Ensures the requested analysis table is initialised via EnsureAnalysisTable,
'then delegates to SafeTableValue for the actual retrieval.
'@param tableRef ByRef CustomTable. Cached table reference.
'@param listName String. Name of the ListObject.
'@param keyColumn String. Column used as the table key.
'@param valueColumn String. Column to read.
'@param keyValue String. Key value to look up.
'@return String. The retrieved value, or empty.
Private Function AnalysisTableValue(ByRef tableRef As CustomTable, _
                                    ByVal listName As String, _
                                    ByVal keyColumn As String, _
                                    ByVal valueColumn As String, _
                                    ByVal keyValue As String) As String
    If Not EnsureAnalysisTable(tableRef, listName, keyColumn) Then Exit Function
    AnalysisTableValue = SafeTableValue(tableRef, valueColumn, keyValue)
End Function

'@label:add-geo-dropdown
'@sub-title Apply geo-type dropdown validation to spatio-temporal analysis rows.
'@details
'When the edited cell falls within the spatio-temporal table, clears the "geo"
'column cell and sets its dropdown validation to either "__hfonly_vars" or
'"__geoonly_vars" depending on the spatial type of the row. Temporarily
'unlocks the Analysis sheet for the modification.
'@param target Range. The edited cell on the Analysis sheet.
Private Sub AddGeoDropdown(ByVal target As Range)
    Dim sectionCell As Range
    Dim geoCell As Range
    Dim spatialTypeCell As Range
    Dim spatialType As String
    Dim pass As Passwords
    Const ANALYSIS_SHEET As String = SHEET_ANALYSIS

    If Not EnsureAnalysisTable(this.spatioTable, LIST_SPATIO_TEMPORAL, "section") Then Exit Sub
    EnsureDropdownValues
    If this.drop Is Nothing Then Exit Sub

    'Only process when the edit is in the "section" column —
    'edits in other columns (e.g. "geo") must not clear the geo cell
    Set sectionCell = this.spatioTable.CellRange("section", target.Row)
    If sectionCell Is Nothing Then Exit Sub
    If target.Column <> sectionCell.Column Then Exit Sub

    Set geoCell = this.spatioTable.CellRange("geo", target.Row)
    If geoCell Is Nothing Then Exit Sub

    'Read spatial type directly from the same row instead of using
    'key-based lookup with target.Value (which fails when the edit
    'is not in the key column)
    Set spatialTypeCell = this.spatioTable.CellRange("spatial type", target.Row)
    If spatialTypeCell Is Nothing Then Exit Sub
    spatialType = CStr(spatialTypeCell.Value)

    Set pass = BeginSheetUpdate(ANALYSIS_SHEET)
    On Error GoTo Cleanup
        geoCell.ClearContents

        Select Case LCase$(spatialType)
            Case "hf"
                this.drop.SetValidation geoCell, "__hfonly_vars"
            Case "geo"
                this.drop.SetValidation geoCell, "__geoonly_vars"
        End Select
Cleanup:
    EndSheetUpdate pass, ANALYSIS_SHEET
End Sub

'@label:sync-geo-from-specs
'@sub-title Propagate spatial type changes from specs to analysis table.
'@details
'When the user changes the "spatial type" column in the spatio-temporal
'specifications table, locates the matching section in the spatio-temporal
'analysis table and clears + revalidates the "geo" column dropdown to
'reflect the new type ("hf" or "geo"). Temporarily unlocks the Analysis
'sheet for the modification.
'@param target Range. The edited cell on the Analysis sheet.
Private Sub SyncGeoFromSpecs(ByVal target As Range)
    Dim sectionCell As Range
    Dim geoCell As Range
    Dim spatialTypeCell As Range
    Dim sectionName As String
    Dim newSpatialType As String
    Dim pass As Passwords
    Const ANALYSIS_SHEET As String = SHEET_ANALYSIS

    'Ensure the specs table is cached
    If Not EnsureAnalysisTable(this.spatioSpecsTable, LIST_SPATIO_TEMPORAL_SPECS, "section") Then Exit Sub

    'Only process when the edit is in the specs table
    If target.ListObject Is Nothing Then Exit Sub
    If StrComp(target.ListObject.Name, LIST_SPATIO_TEMPORAL_SPECS, vbTextCompare) <> 0 Then Exit Sub

    'Only process when the edit is in the "spatial type" column
    Set spatialTypeCell = this.spatioSpecsTable.CellRange("spatial type", target.Row)
    If spatialTypeCell Is Nothing Then Exit Sub
    If target.Column <> spatialTypeCell.Column Then Exit Sub

    'Read the section name from the same row
    Set sectionCell = this.spatioSpecsTable.CellRange("section", target.Row)
    If sectionCell Is Nothing Then Exit Sub
    sectionName = CStr(sectionCell.Value)
    If LenB(sectionName) = 0 Then Exit Sub

    newSpatialType = CStr(target.Value)

    'Ensure the analysis table and dropdown manager are available
    If Not EnsureAnalysisTable(this.spatioTable, LIST_SPATIO_TEMPORAL, "section") Then Exit Sub
    EnsureDropdownValues
    If this.drop Is Nothing Then Exit Sub

    'Find the matching "geo" cell in the analysis table by section name
    Set geoCell = this.spatioTable.ValueRange("geo", sectionName)
    If geoCell Is Nothing Then Exit Sub

    Set pass = BeginSheetUpdate(ANALYSIS_SHEET)
    On Error GoTo Cleanup
        geoCell.ClearContents

        Select Case LCase$(newSpatialType)
            Case "hf"
                this.drop.SetValidation geoCell, "__hfonly_vars"
            Case "geo"
                this.drop.SetValidation geoCell, "__geoonly_vars"
        End Select
Cleanup:
    EndSheetUpdate pass, ANALYSIS_SHEET
End Sub

'@label:add-choices-dropdown
'@sub-title Apply choices dropdown validation to graph series rows.
'@details
'When the user edits a "series title" cell in the graph table, resolves the
'column variable's control details, fetches category lists via GetCategories,
'and sets dropdown validation on the "choice" and "values or percentages"
'cells. Temporarily unlocks the Analysis sheet for the modification.
'@param target Range. The edited cell on the Analysis sheet.
Private Sub AddChoicesDropdown(ByVal target As Range)
    Dim seriesRange As Range
    Dim columnVariable As String
    Dim seriesValue As String
    Dim dropArray As BetterArray
    Dim resolvedChoice As String
    Dim listName As String
    Dim choiceCell As Range
    Dim valueCell As Range
    Dim pass As Passwords

    If Not EnsureAnalysisTable(this.graphTable, LIST_GRAPH_TS, "series title") Then Exit Sub
    If Not EnsureAnalysisTable(this.tsTable, LIST_TS_DATA, "title") Then Exit Sub

    If this.graphTable Is Nothing Then Exit Sub
    If this.tsTable Is Nothing Then Exit Sub

    'AddGeoDropdown builds the dropdown manager on its way in and runs after
    'this one, so a first edit here used to find the cache empty and skip.
    EnsureDropdownValues
    If this.drop Is Nothing Then Exit Sub

    Set seriesRange = this.graphTable.DataRange("series title")
    If seriesRange Is Nothing Then Exit Sub
    If Intersect(target, seriesRange) Is Nothing Then Exit Sub

    seriesValue = CStr(target.Value)
    On Error Resume Next
        columnVariable = CStr(this.graphTable.Value("column", seriesValue))
    On Error GoTo 0
    If columnVariable = vbNullString Then Exit Sub

    Set dropArray = GetCategories(columnVariable, seriesValue, resolvedChoice)

    If dropArray Is Nothing Then Exit Sub
    If dropArray.Length = 0 Then Exit Sub

    listName = resolvedChoice & "__"

    AddOrUpdateDrop dropArray, listName

    Set choiceCell = this.graphTable.CellRange("choice", target.Row)
    Set valueCell = this.graphTable.CellRange("values or percentages", target.Row)

    Set pass = BeginSheetUpdate(SHEET_ANALYSIS)
    On Error GoTo Cleanup
        If Not choiceCell Is Nothing Then
            choiceCell.ClearContents
            this.drop.SetValidation choiceCell, listName
            FormatLockCell choiceCell, False
        End If

        If Not valueCell Is Nothing Then
            this.drop.SetValidation valueCell, "__perc_val"
            FormatLockCell valueCell, False
        End If
Cleanup:
    EndSheetUpdate pass, SHEET_ANALYSIS
End Sub

'@label:get-categories
'@sub-title Resolve choice categories for a graph column variable.
'@details
'When the control details start with "CHOICE_FORMULA", delegates to
'ChoiceFormula to parse the formula and retrieve categories. Otherwise fetches
'categories from the choices helper by name. Appends "Total" when the
'time-series table indicates addTotal is "yes" for the series.
'@param columnVariable String. The variable whose control details determine the category list.
'@param seriesValue String. The series title for looking up the addTotal flag.
'@param resolvedChoice ByRef String. Populated with the resolved choice list name.
'@return BetterArray. The list of categories.
Private Function GetCategories(ByVal columnVariable As String, _
                               ByVal seriesValue As String, _
                               ByRef resolvedChoice As String) As BetterArray
    Dim categoryList As BetterArray
    Dim controlDetails As String
    Dim choiceResolver As Object
    Dim addTotalFlag As String

    Set categoryList = New BetterArray

    EnsureVars
    If this.variables Is Nothing Then GoTo ExitFunction

    controlDetails = Trim$(this.variables.Value("Control Details", columnVariable))
    If controlDetails = vbNullString Then GoTo ExitFunction

    If Left$(controlDetails, Len("CHOICE_FORMULA")) = "CHOICE_FORMULA" Then
        Set choiceResolver = ChoiceFormula.Create(controlDetails)
        Set categoryList = choiceResolver.Categories()
        resolvedChoice = choiceResolver.ChoiceName()
    Else
        resolvedChoice = controlDetails
        Set categoryList = ChoiceCategories(controlDetails)
    End If

    If (Not this.tsTable Is Nothing) Then
        On Error Resume Next
            addTotalFlag = CStr(this.tsTable.Value("add total", seriesValue))
            If StrComp(addTotalFlag, "yes", vbTextCompare) = 0 Then
                categoryList.Push "Total"
            End If
        On Error GoTo 0
    End If

ExitFunction:
    Set GetCategories = categoryList
End Function

'@label:choice-categories
'@sub-title Return the BetterArray of categories for a named choice list.
'@details
'Fetches categories from the cached LLChoices helper. Returns an empty
'BetterArray when the helper is unavailable or the list does not exist.
'@param listName String. Name of the choice list to retrieve.
'@return BetterArray. The categories, or an empty BetterArray.
Private Function ChoiceCategories(ByVal listName As String) As BetterArray
    Dim emptyList As BetterArray

    Set emptyList = New BetterArray
    emptyList.LowerBound = 0

    If Not EnsureChoices() Then
        Set ChoiceCategories = emptyList
        Exit Function
    End If

    On Error Resume Next
        Set ChoiceCategories = this.choices.Categories(listName)
    On Error GoTo 0

    If ChoiceCategories Is Nothing Then
        Set ChoiceCategories = emptyList
    End If
End Function

'@label:format-lock-cell
'@sub-title Toggle font styling and lock state on a single cell.
'@details
'Sets the font colour to blue and italic when locked, or black and normal
'when unlocked. Used to visually indicate whether a cell is user-editable.
'@param cellRng Range. The cell to format.
'@param locked Optional Boolean. Lock state to apply. Defaults to True.
Private Sub FormatLockCell(ByVal cellRng As Range, Optional ByVal locked As Boolean = True)
    If cellRng Is Nothing Then Exit Sub

    cellRng.Font.Color = IIf(locked, RGB(51, 142, 202), vbBlack)
    cellRng.Font.Italic = locked
    cellRng.Locked = locked
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-aligned error from the current class.
'@details
'Wrapper around Err.Raise that standardises the source to CLASSNAME,
'providing a consistent stack trace across all methods in this class.
'@param errNumber ProjectError. The error code to raise.
'@param errorMessage String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errorMessage As String)
  Err.Raise CLng(errNumber), CLASSNAME, "Error: " & errorMessage
End Sub
