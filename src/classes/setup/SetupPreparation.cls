VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetupPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Prepare dropdown lists and updated values registry for setup workbooks"

Option Explicit
Implements ISetupPreparation

'@PredeclaredId
'@Folder("Setup")
'@ModuleDescription("Orchestrates dropdown registration and updated values initialisation for setup workbooks")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends BetterArray, DropdownLists, UpdatedValues, CustomTable, ProjectError, IDevelopment

'@class SetupPreparation
'@description
'Orchestrates the preparation workflow for setup workbooks by registering
'dropdown list objects, initialising the updated values registry, applying
'data validations to setup tables (Dictionary, Exports, Analysis), and
'configuring hidden and protected sheets. Consumed by the setup entry
'point to prepare all worksheet assets before the workbook is ready for
'use. Uses the predeclared factory pattern through the ISetupPreparation
'interface.


'@section State
'===============================================================================
Private Type TSetupPreparation
    hostBook    As Workbook          'Workbook owning setup assets
    dropSheet   As Worksheet         'Worksheet hosting dropdown listobjects
    registrySheet As Worksheet       'Worksheet hosting updated values registry
    dropdowns   As IDropdownLists    'Dropdown manager bound to drop sheet
    updater     As IUpdatedValues    'Updated values watcher bound to registry sheet
End Type

Private this As TSetupPreparation

'@section Constants
'===============================================================================
Private Const SHEET_DROPDOWNS As String = "__variables"
Private Const SHEET_FORMATTER As String = "__formatter"
Private Const SHEET_UPDATED As String = "__updated"
Private Const SHEET_FORMULA As String = "__formula"
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const SHEET_CHOICES As String = "Choices"
Private Const SHEET_ANALYSIS As String = "Analysis"
Private Const SHEET_EXPORTS As String = "Exports"
Private Const SHEET_PASS As String = "__pass"
Private Const SHEET_DEV As String = "Dev"
Private Const SHEET_CHECKING As String = "__checkRep"
Private Const DROPDOWN_PREFIX As String = "dropdown_"
Private Const STATUS_DEFAULT As String = "no"
Private Const SHEET_TRANSLATIONS As String = "Translations"

Private Const TABLE_DICTIONARY As String = "Tab_Dictionary"
Private Const TABLE_EXPORTS As String = "Tab_Export"
Private Const TABLE_ANALYSIS_GLOBAL As String = "Tab_Global_Summary"
Private Const TABLE_ANALYSIS_UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const TABLE_ANALYSIS_BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TABLE_ANALYSIS_TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const TABLE_ANALYSIS_GRAPH_TS As String = "Tab_Graph_TimeSeries"
Private Const TABLE_ANALYSIS_SPATIAL As String = "Tab_Spatial_Analysis"
Private Const TABLE_ANALYSIS_SPATIOTEMP_SPECS As String = "Tab_SpatioTemporal_Specs"
Private Const TABLE_ANALYSIS_SPATIOTEMP As String = "Tab_SpatioTemporal_Analysis"

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create an initialised SetupPreparation instance
'@details
'Factory method on the predeclared instance. Validates the workbook is
'not Nothing via Configure, then returns the new instance through the
'ISetupPreparation interface.
'@param hostBook Workbook. The workbook hosting setup worksheets.
'@return ISetupPreparation. A fully initialised instance ready for use.
'@throws ProjectError.ObjectNotInitialized When hostBook is Nothing.
'@export
Public Function Create(ByVal hostBook As Workbook) As ISetupPreparation
    Dim instance As SetupPreparation

    Set instance = New SetupPreparation
    instance.Configure hostBook
    Set Create = instance.Self
End Function

'@label:self
'@prop-title Current object instance
'@return ISetupPreparation. The current instance cast to the interface.
Public Property Get Self() As ISetupPreparation
    Set Self = Me
End Property

'@section ISetupPreparation implementation
'===============================================================================
'@description Delegated members satisfying the ISetupPreparation contract.
'See the corresponding Public members above for full documentation.


Private Sub ISetupPreparation_Configure(ByVal hostBook As Workbook)
    Configure hostBook
End Sub

Private Sub ISetupPreparation_Prepare(ByRef manager As IDevelopment)
    Prepare manager
End Sub

Private Sub ISetupPreparation_EnsureDropdowns()
    EnsureDropdowns
End Sub

Private Sub ISetupPreparation_EnsureUpdatedRegistry()
    EnsureUpdatedRegistry
End Sub

Private Property Get ISetupPreparation_Dropdowns() As IDropdownLists
    Set ISetupPreparation_Dropdowns = Dropdowns
End Property

Private Property Get ISetupPreparation_UpdatedWatcher() As IUpdatedValues
    Set ISetupPreparation_UpdatedWatcher = UpdatedWatcher
End Property

Private Property Get ISetupPreparation_HostWorkbook() As Workbook
    Set ISetupPreparation_HostWorkbook = HostWorkbook
End Property

'@section Public API
'===============================================================================

'@label:configure
'@sub-title Bind the preparation helper to a workbook hosting setup worksheets
'@details
'Stores the workbook reference and resets all cached worksheet and manager
'objects to Nothing, forcing lazy re-initialisation on next access.
'@param hostBook Workbook. The workbook owning the setup assets.
'@throws ProjectError.ObjectNotInitialized When hostBook is Nothing.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook reference is required"
    End If

    Set this.hostBook = hostBook
    ResetCaches
End Sub

'@label:prepare
'@sub-title Execute the full preparation workflow
'@details
'Runs the complete preparation sequence: registers all dropdown list objects,
'initialises the updated values registry, applies data validations to the
'Dictionary, Exports, and Analysis tables, then registers hidden and
'protected sheets via the supplied development manager.
'@param manager IDevelopment. Development manager providing hidden/protected sheet registration.
'@export
Public Sub Prepare(ByRef manager As IDevelopment)
    EnsureDropdowns
    EnsureUpdatedRegistry
    ApplyValidationsAndUpdates
    EnsureHiddenSheets manager
    EnsureProtectedSheets manager
End Sub

'@label:ensure-dropdowns
'@sub-title Ensure the dropdown manager has registered all default dropdown lists
'@details
'Creates the dropdown manager if needed, then calls RegisterAllDropdowns to
'populate every standard dropdown list on the dropdown worksheet. Safe to
'call multiple times; the manager creation is a no-op when already
'initialised.
'@export
Public Sub EnsureDropdowns()
    EnsureDropdownManager
    RegisterAllDropdowns
End Sub

'@label:ensure-updated-registry
'@sub-title Initialise the updated values registry and register watched tables
'@details
'Creates the update watcher if needed, clears any existing registry data,
'registers the watched setup sheets (Dictionary, Choices, Analysis, Exports),
'and resets all status tags to "yes" (the default ready state).
'@export
Public Sub EnsureUpdatedRegistry()
    EnsureUpdatedValuesWatcher
    ResetUpdatedRegistry
    RegisterWatchedSheets
    ResetRegistryStatuses
End Sub


'@label:dropdowns
'@prop-title Dropdown manager for external callers
'@details
'Returns the IDropdownLists manager bound to the dropdown worksheet.
'Lazily initialises the manager on first access if it was not already
'created by EnsureDropdowns or Prepare.
'@return IDropdownLists. The dropdown manager instance.
'@export
Public Property Get Dropdowns() As IDropdownLists
    EnsureDropdownManager
    Set Dropdowns = this.dropdowns
End Property

'@label:updated-watcher
'@prop-title Updated values watcher for external callers
'@details
'Returns the IUpdatedValues watcher bound to the registry worksheet.
'Lazily initialises the watcher on first access if it was not already
'created by EnsureUpdatedRegistry or Prepare.
'@return IUpdatedValues. The update watcher instance.
'@export
Public Property Get UpdatedWatcher() As IUpdatedValues
    EnsureUpdatedValuesWatcher
    Set UpdatedWatcher = this.updater
End Property

'@label:host-workbook
'@prop-title Workbook hosting setup assets
'@details
'Returns the cached workbook reference stored during Configure.
'Delegates to EnsureWorkbook, which raises an error when the instance
'has not been configured.
'@return Workbook. The bound workbook.
'@throws ProjectError.ObjectNotInitialized When the instance has not been configured.
'@export
Public Property Get HostWorkbook() As Workbook
    Set HostWorkbook = EnsureWorkbook
End Property

'@section Preparation helpers
'===============================================================================

'@label:reset-caches
'@sub-title Reset cached references after reconfiguring the helper
Private Sub ResetCaches()
    Set this.dropdowns = Nothing
    Set this.updater = Nothing
    Set this.dropSheet = Nothing
    Set this.registrySheet = Nothing
End Sub

'@label:ensure-dropdown-manager
'@sub-title Ensure the dropdown manager is initialised for the target workbook
'@details
'Exits immediately when the manager is already initialised. Otherwise
'resolves or creates the dropdown worksheet, then creates a DropdownLists
'instance bound to it with the standard dropdown prefix.
Private Sub EnsureDropdownManager()
    If Not this.dropdowns Is Nothing Then Exit Sub

    Dim targetSheet As Worksheet

    Set targetSheet = EnsureDropdownSheet
    Set this.dropdowns = DropdownLists.Create(targetSheet, DROPDOWN_PREFIX)
End Sub

'@label:ensure-updated-values-watcher
'@sub-title Ensure the updated values watcher is initialised for the registry worksheet
'@details
'Exits immediately when the watcher is already initialised. Otherwise
'resolves or creates the registry worksheet, then creates an UpdatedValues
'instance bound to it.
Private Sub EnsureUpdatedValuesWatcher()
    If Not this.updater Is Nothing Then Exit Sub

    Dim registry As Worksheet

    Set registry = EnsureRegistrySheet
    Set this.updater = UpdatedValues.Create(registry)
End Sub

'@label:reset-updated-registry
'@sub-title Reset the updated values registry prior to registering watched tables
Private Sub ResetUpdatedRegistry()
    If this.updater Is Nothing Then Exit Sub
    this.updater.DeleteUp
End Sub

'@label:apply-validations-and-updates
'@sub-title Apply data validation rules to all setup tables
'@details
'Ensures the dropdown manager is initialised, then delegates to
'ApplyDictionaryValidations, ApplyExportsValidations, and
'ApplyAnalysisValidations to wire each table column to its dropdown list.
Private Sub ApplyValidationsAndUpdates()
    Dim drop As IDropdownLists

    EnsureDropdownManager
    Set drop = this.dropdowns
    ApplyDictionaryValidations drop
    ApplyExportsValidations drop
    ApplyAnalysisValidations drop
End Sub

'@label:ensure-hidden-sheets
'@sub-title Register internal worksheets as hidden via the development manager
'@details
'Calls manager.AddHiddenSheet for each internal worksheet constant
'(updated, variables, formatter, pass, formula, dev).
'@param manager IDevelopment. Development manager to register with.
'@throws ProjectError.ObjectNotInitialized When manager is Nothing.
Private Sub EnsureHiddenSheets(ByRef manager As IDevelopment)
    If manager Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Development manager reference is required"
    End If

    manager.AddHiddenSheet SHEET_UPDATED
    manager.AddHiddenSheet SHEET_DROPDOWNS
    manager.AddHiddenSheet SHEET_FORMATTER
    manager.AddHiddenSheet SHEET_PASS
    manager.AddHiddenSheet SHEET_FORMULA
    manager.AddHiddenSheet SHEET_DEV
End Sub

'@label:ensure-protected-sheets
'@sub-title Register user-facing worksheets as protected via the development manager
'@details
'Calls manager.AddProtectedSheet for each user-facing worksheet, specifying
'whether filtering and sorting are allowed on the protected sheet.
'@param manager IDevelopment. Development manager to register with.
'@throws ProjectError.ObjectNotInitialized When manager is Nothing.
Private Sub EnsureProtectedSheets(ByRef manager As IDevelopment)

    If manager Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Development manager reference is required"
    End If

    manager.AddProtectedSheet SHEET_DICTIONARY, True, True
    manager.AddProtectedSheet SHEET_CHOICES, True, True
    manager.AddProtectedSheet SHEET_ANALYSIS, True, False
    manager.AddProtectedSheet SHEET_EXPORTS, True, False
    manager.AddProtectedSheet SHEET_TRANSLATIONS, False, False

End Sub


'@label:register-all-dropdowns
'@sub-title Register every standard dropdown list for setup worksheets
'@details
'Calls RegisterDropdown for each predefined dropdown (yes/no, formats,
'variable status/type/control, sheet type, export options, analysis
'parameters, geo variables, chart types, etc.). The dropdown names use
'the double-underscore prefix convention.
Private Sub RegisterAllDropdowns()
    RegisterDropdown "__yesno", Array("yes", "no")
    RegisterDropdown "__formats", Array("integer", "round0", "round1", "round2", "round3", _
                                        "percentage0", "percentage1", "percentage2", "percentage3", _
                                        "text", "euros", "dollars", "dd/mm/yyyy", "d-mmm-yyyy", vbNullString)
    RegisterDropdown "__var_status", Array("mandatory", "optional, visible", "optional, hidden", "hidden")
    RegisterDropdown "__var_type", Array("date", "integer", "text", "decimal")
    RegisterDropdown "__sheet_type", Array("vlist1D", "hlist2D")
    RegisterDropdown "__var_control", Array("choice_manual", "choice_formula", "choice_custom", _
                                            "choice_multiple", "formula", "geo", "hf", "list_auto", "case_when")
    RegisterDropdown "__var_print", Array("print, horizontal header", "print, vertical header", "hidden")
    RegisterDropdown "__var_alert", Array("error", "warning", "info")
    RegisterDropdown "__geo_vars", Array(vbNullString, vbNullString)
    RegisterDropdown "__hfonly_vars", Array(vbNullString, vbNullString)
    RegisterDropdown "__geoonly_vars", Array(vbNullString, vbNullString)
    RegisterDropdown "__choice_vars", Array(vbNullString, vbNullString)
    RegisterDropdown "__time_vars", Array(vbNullString, vbNullString)
    RegisterDropdown "__hf_or_geo", Array("hf", "geo")
    RegisterDropdown "__export_status", Array("active", "inactive")
    RegisterDropdown "__export_format", Array("xlsx", "xlsb")
    RegisterDropdown "__export_header", Array("variables names", "variables labels", "variable names + variable labels")
    RegisterDropdown "__percentage_ba", Array("no", "row", "column", "total")
    RegisterDropdown "__missing_ba", Array("no", "row", "column", "all")
    RegisterDropdown "__percentage_ta", Array("no", "row", "column")
    RegisterDropdown "__perc_val", Array("percentages", "values")
    RegisterDropdown "__chart_type", Array("bar", "line", "point")
    RegisterDropdown "__axis_pos", Array("left", "right")
    RegisterDropdown "__swicth_tables", Array( _
        "Add or remove rows of global summary", _
        "Add or remove rows of univariate analysis", _
        "Add or remove rows of bivariate analysis", _
        "Add or remove rows of time series analysis", _
        "Add or remove rows of labels for time series graphs", _
        "Add or remove rows of graph on time series", _
        "Add or remove rows of spatial analysis", _
        "Add or remove rows of spatio-temporal specifications", _
        "Add or remove rows of spatio-temporal analysis", _
        "Add or remove rows of all tables")
    RegisterDropdown "__graphs_titles", Array(vbNullString, vbNullString)
    RegisterDropdown "__series_titles", Array(vbNullString, vbNullString)
    RegisterDropdown "__checking_types", Array("all", "errors", "warnings", "info")
End Sub

'@label:register-dropdown
'@sub-title Register a single dropdown by converting the values into a BetterArray
'@details
'Ensures the dropdown manager is initialised, converts the supplied Variant
'array into a BetterArray (1-based), and adds it to the dropdown manager
'under the given name.
'@param dropdownName String. The name to register the dropdown under.
'@param values Variant. An array of values or a single value to populate the dropdown.
Private Sub RegisterDropdown(ByVal dropdownName As String, ByVal values As Variant)
    Dim listValues As BetterArray
    Dim idx As Long

    EnsureDropdownManager

    Set listValues = New BetterArray
    listValues.LowerBound = 1

    If IsArray(values) Then
        For idx = LBound(values) To UBound(values)
            listValues.Push values(idx)
        Next idx
    Else
        listValues.Push values
    End If

    this.dropdowns.Add listValues, dropdownName
End Sub

'@label:register-watched-sheets
'@sub-title Register watched columns for every relevant setup sheet
'@details
'Iterates over the list of watched sheet names, resolves each to a
'worksheet in the host workbook, and registers it with the update watcher
'via AddSheet. Raises an error when a required sheet is missing.
'@throws ProjectError.ElementNotFound When a watched sheet does not exist.
Private Sub RegisterWatchedSheets()
    Dim watchedSheets As BetterArray
    Dim idx As Long
    Dim sheetName As String
    Dim targetSheet As Worksheet

    EnsureUpdatedValuesWatcher

    Set watchedSheets = WatchedSheetNames

    For idx = watchedSheets.LowerBound To watchedSheets.UpperBound
        sheetName = CStr(watchedSheets.Item(idx))
        Set targetSheet = ResolveWorksheet(sheetName)
        If targetSheet Is Nothing Then
            ThrowError ProjectError.ElementNotFound, "Worksheet '" & sheetName & "' not found for updated values registry"
        End If
        this.updater.AddSheet targetSheet
    Next idx
End Sub

'@label:watched-sheet-names
'@sub-title Return the list of setup sheets whose tables should be watched for updates
'@return BetterArray. Sheet names (Dictionary, Choices, Analysis, Exports).
Private Function WatchedSheetNames() As BetterArray
    Dim names As BetterArray
    Set names = New BetterArray
    names.Push SHEET_DICTIONARY, SHEET_CHOICES, SHEET_ANALYSIS, SHEET_EXPORTS
    Set WatchedSheetNames = names
End Function

'@label:reset-registry-statuses
'@sub-title Reset registry statuses to the default value after registration
Private Sub ResetRegistryStatuses()
    EnsureUpdatedValuesWatcher
    this.updater.SwitchTagsToYes
End Sub

'@section Worksheet helpers
'===============================================================================

'@label:ensure-dropdown-sheet
'@sub-title Ensure the dropdown worksheet exists and cache the reference
'@return Worksheet. The dropdown worksheet.
Private Function EnsureDropdownSheet() As Worksheet
    If Not this.dropSheet Is Nothing Then
        Set EnsureDropdownSheet = this.dropSheet
        Exit Function
    End If

    Dim sh As Worksheet

    Set sh = ResolveOrCreateWorksheet(SHEET_DROPDOWNS, xlSheetVisible)
    Set this.dropSheet = sh
    Set EnsureDropdownSheet = sh
End Function

'@label:ensure-registry-sheet
'@sub-title Ensure the registry worksheet exists and cache the reference
'@return Worksheet. The registry worksheet.
Private Function EnsureRegistrySheet() As Worksheet
    If Not this.registrySheet Is Nothing Then
        Set EnsureRegistrySheet = this.registrySheet
        Exit Function
    End If

    Dim sh As Worksheet

    Set sh = ResolveOrCreateWorksheet(SHEET_UPDATED, addVisibility:=False)
    Set this.registrySheet = sh
    Set EnsureRegistrySheet = sh
End Function

'@label:ensure-workbook
'@sub-title Return the cached workbook reference, raising when missing
'@return Workbook. The host workbook.
'@throws ProjectError.ObjectNotInitialized When the instance has not been configured.
Private Function EnsureWorkbook() As Workbook
    If this.hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "SetupPreparation has not been configured"
    End If
    Set EnsureWorkbook = this.hostBook
End Function

'@label:resolve-worksheet
'@sub-title Resolve an existing worksheet by name without creating it
'@param sheetName String. The worksheet name to look up.
'@return Worksheet. The resolved worksheet, or Nothing when not found.
Private Function ResolveWorksheet(ByVal sheetName As String) As Worksheet
    Dim wb As Workbook
    Dim sh As Worksheet

    Set wb = EnsureWorkbook

    On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0

    Set ResolveWorksheet = sh
End Function

'@label:resolve-or-create-worksheet
'@sub-title Resolve a worksheet if it exists, otherwise create one
'@details
'Looks up the worksheet by name in the host workbook. When not found,
'appends a new worksheet at the end and assigns the given name. Optionally
'applies the supplied visibility state.
'@param sheetName String. The worksheet name to look up or create.
'@param visibility Optional XlSheetVisibility. Visibility state to apply when addVisibility is True.
'@param addVisibility Optional Boolean. When True, applies the visibility parameter. Defaults to True.
'@return Worksheet. The resolved or newly created worksheet.
Private Function ResolveOrCreateWorksheet(ByVal sheetName As String, _
                                          Optional ByVal visibility As XlSheetVisibility, _
                                          Optional ByVal addVisibility As Boolean = True) As Worksheet
    Dim sh As Worksheet
    Dim wb As Workbook

    Set wb = EnsureWorkbook

    On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        sh.Name = sheetName
    End If

    If addVisibility Then sh.Visible = visibility
    Set ResolveOrCreateWorksheet = sh
End Function

'@label:apply-dictionary-validations
'@sub-title Apply data validation rules to the dictionary worksheet tables
'@details
'Resolves the Dictionary worksheet and its Tab_Dictionary ListObject,
'then wires each column (sheet type, status, variable type, format,
'control, etc.) to its corresponding dropdown list via SetValidation.
'@param drop IDropdownLists. The dropdown manager providing validation ranges.
Private Sub ApplyDictionaryValidations(ByVal drop As IDropdownLists)
    Dim sh As Worksheet
    Dim table As ICustomTable

    Set sh = RequireWorksheet(SHEET_DICTIONARY)
    Set table = RequireCustomTable(sh, TABLE_DICTIONARY)

    table.SetValidation colName:="sheet type", dropName:="__sheet_type", drop:=drop, alertType:="error"
    table.SetValidation colName:="editable label", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="status", dropName:="__var_status", drop:=drop, alertType:="error"
    table.SetValidation colName:="personal identifier", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="variable type", dropName:="__var_type", drop:=drop, alertType:="error"
    table.SetValidation colName:="variable format", dropName:="__formats", drop:=drop, alertType:="info"
    table.SetValidation colName:="control", dropName:="__var_control", drop:=drop, alertType:="info"
    table.SetValidation colName:="register book", dropName:="__var_print", drop:=drop, alertType:="info"
    table.SetValidation colName:="unique", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="alert", dropName:="__var_alert", drop:=drop, alertType:="error"
    table.SetValidation colName:="lock cells", dropName:="__yesno", drop:=drop, alertType:="error"
End Sub

'@label:apply-exports-validations
'@sub-title Apply export worksheet validations
'@details
'Resolves the Exports worksheet and its Tab_Export ListObject, then wires
'each column (status, file format, password, etc.) to its dropdown list.
'@param drop IDropdownLists. The dropdown manager providing validation ranges.
Private Sub ApplyExportsValidations(ByVal drop As IDropdownLists)
    Dim sh As Worksheet
    Dim table As ICustomTable

    Set sh = RequireWorksheet(SHEET_EXPORTS)
    Set table = RequireCustomTable(sh, TABLE_EXPORTS)

    table.SetValidation colName:="status", dropName:="__export_status", drop:=drop, alertType:="error"
    table.SetValidation colName:="file format", dropName:="__export_format", drop:=drop, alertType:="error"
    table.SetValidation colName:="password", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="include personal identifiers", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="include p-codes", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="header format", dropName:="__export_header", drop:=drop, alertType:="error"
    table.SetValidation colName:="export metadata", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="export analyses sheets", dropName:="__yesno", drop:=drop, alertType:="error"
End Sub

'@label:apply-analysis-validations
'@sub-title Apply validations to all analysis tables and named ranges
'@details
'Resolves the Analysis worksheet, applies the table-switcher validation
'to the RNG_SelectTable named range, then wires every analysis table
'(global summary, univariate, bivariate, time series, graph, spatial,
'spatio-temporal specs, spatio-temporal) to their respective dropdown
'lists.
'@param drop IDropdownLists. The dropdown manager providing validation ranges.
Private Sub ApplyAnalysisValidations(ByVal drop As IDropdownLists)
    Dim sh As Worksheet
    Dim table As ICustomTable
    Dim targetRange As Range

    Set sh = RequireWorksheet(SHEET_ANALYSIS)

    Set targetRange = ResolveNamedRange(sh, "RNG_SelectTable")
    If Not targetRange Is Nothing Then
        drop.SetValidation targetRange, "__swicth_tables", "error"
    End If

    Set table = RequireCustomTable(sh, TABLE_ANALYSIS_GLOBAL)
    table.SetValidation colName:="format", dropName:="__formats", drop:=drop, alertType:="info"

    Set table = RequireCustomTable(sh, TABLE_ANALYSIS_UNIVARIATE)
    table.SetValidation colName:="add missing data", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="format", dropName:="__formats", drop:=drop, alertType:="info"
    table.SetValidation colName:="add percentage", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="add graph", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="flip coordinates", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="row", dropName:="__choice_vars", drop:=drop, alertType:="error"

    Set table = RequireCustomTable(sh, TABLE_ANALYSIS_BIVARIATE)
    table.SetValidation colName:="add missing data", dropName:="__missing_ba", drop:=drop, alertType:="error"
    table.SetValidation colName:="format", dropName:="__formats", drop:=drop, alertType:="info"
    table.SetValidation colName:="add percentage", dropName:="__percentage_ba", drop:=drop, alertType:="error"
    table.SetValidation colName:="add Graph", dropName:="__perc_val", drop:=drop, alertType:="error"
    table.SetValidation colName:="flip coordinates", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="row", dropName:="__choice_vars", drop:=drop, alertType:="error"
    table.SetValidation colName:="column", dropName:="__choice_vars", drop:=drop, alertType:="error"

    Set table = RequireCustomTable(sh, TABLE_ANALYSIS_TIMESERIES)
    table.SetValidation colName:="add missing data", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="format", dropName:="__formats", drop:=drop, alertType:="info"
    table.SetValidation colName:="add percentage", dropName:="__percentage_ta", drop:=drop, alertType:="error"
    table.SetValidation colName:="add total", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="row", dropName:="__time_vars", drop:=drop, alertType:="error"
    table.SetValidation colName:="column", dropName:="__choice_vars", drop:=drop, alertType:="info"

    Set table = RequireCustomTable(sh, TABLE_ANALYSIS_GRAPH_TS)
    table.SetValidation colName:="plot values or percentages", dropName:="__perc_val", drop:=drop, alertType:="error"
    table.SetValidation colName:="chart type", dropName:="__chart_type", drop:=drop, alertType:="info"
    table.SetValidation colName:="y-axis", dropName:="__axis_pos", drop:=drop, alertType:="error"

    Set table = RequireCustomTable(sh, TABLE_ANALYSIS_SPATIAL)
    table.SetValidation colName:="row", dropName:="__geo_vars", drop:=drop, alertType:="error"
    table.SetValidation colName:="column", dropName:="__choice_vars", drop:=drop, alertType:="info"
    table.SetValidation colName:="add missing data", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="add percentage", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="add graph", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="flip coordinates", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="format", dropName:="__formats", drop:=drop, alertType:="info"

    Set table = RequireCustomTable(sh, TABLE_ANALYSIS_SPATIOTEMP_SPECS)
    table.SetValidation colName:="spatial type", dropName:="__hf_or_geo", drop:=drop, alertType:="error"

    Set table = RequireCustomTable(sh, TABLE_ANALYSIS_SPATIOTEMP)
    table.SetValidation colName:="row", dropName:="__time_vars", drop:=drop, alertType:="error"
    table.SetValidation colName:="column", dropName:="__geo_vars", drop:=drop, alertType:="error"
    table.SetValidation colName:="format", dropName:="__formats", drop:=drop, alertType:="info"
    table.SetValidation colName:="flip coordinates", dropName:="__yesno", drop:=drop, alertType:="error"
    table.SetValidation colName:="add graph", dropName:="__yesno", drop:=drop, alertType:="error"
End Sub

'@label:require-worksheet
'@sub-title Retrieve a worksheet or raise if unavailable
'@param sheetName String. The worksheet name to resolve.
'@return Worksheet. The resolved worksheet.
'@throws ProjectError.ElementNotFound When the worksheet does not exist.
Private Function RequireWorksheet(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    Set sh = ResolveWorksheet(sheetName)
    If sh Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Worksheet '" & sheetName & "' not found in setup workbook"
    End If

    Set RequireWorksheet = sh
End Function

'@label:require-list-object
'@sub-title Retrieve a ListObject from a worksheet or raise when missing
'@param sh Worksheet. The worksheet to search.
'@param listName String. The ListObject name to find.
'@return ListObject. The resolved ListObject.
'@throws ProjectError.ElementNotFound When the ListObject does not exist.
Private Function RequireListObject(ByVal sh As Worksheet, ByVal listName As String) As ListObject
    Dim lo As ListObject

    On Error Resume Next
        Set lo = sh.ListObjects(listName)
    On Error GoTo 0

    If lo Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "ListObject '" & listName & "' not found on worksheet '" & sh.Name & "'"
    End If

    Set RequireListObject = lo
End Function

'@label:require-custom-table
'@sub-title Return a CustomTable helper for the requested ListObject
'@param sh Worksheet. The worksheet hosting the ListObject.
'@param listName String. The ListObject name to wrap.
'@return ICustomTable. A CustomTable wrapper for the ListObject.
Private Function RequireCustomTable(ByVal sh As Worksheet, ByVal listName As String) As ICustomTable
    Dim lo As ListObject

    Set lo = RequireListObject(sh, listName)
    Set RequireCustomTable = CustomTable.Create(lo)
End Function

'@label:resolve-named-range
'@sub-title Resolve a named range within a worksheet
'@param sh Worksheet. The worksheet to search.
'@param rangeName String. The named range to resolve.
'@return Range. The resolved range, or Nothing when not available.
Private Function ResolveNamedRange(ByVal sh As Worksheet, ByVal rangeName As String) As Range
    On Error Resume Next
        Set ResolveNamedRange = sh.Range(rangeName)
    On Error GoTo 0
End Function

'@label:normalize-text
'@sub-title Normalise a string for case-insensitive comparisons
'@param valueText String. The text to normalise.
'@return String. The trimmed, lowercased text.
Private Function NormalizeText(ByVal valueText As String) As String
    NormalizeText = LCase$(Trim$(valueText))
End Function

'@section Error handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a project-scoped error with the supplied message
'@details
'Wrapper around Err.Raise that standardises the source to
'"SetupPreparation", providing a consistent stack trace across
'all methods in this class.
'@param errNumber ProjectError. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, "SetupPreparation", "Error: " & message
End Sub
