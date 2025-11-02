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

'@Folder("Setup")
'@ModuleDescription("Orchestrates dropdown registration and updated values initialisation for setup workbooks")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends BetterArray, DropdownLists, UpdatedValues, ProjectError, IDevelopment

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
'@Description("Create an initialised SetupPreparation instance bound to the supplied workbook.")
Public Function Create(ByVal hostBook As Workbook) As ISetupPreparation
    Dim instance As SetupPreparation

    Set instance = New SetupPreparation
    instance.Configure hostBook
    Set Create = instance.Self
End Function

'@Description("Return the current instance as an ISetupPreparation reference.")
Public Property Get Self() As ISetupPreparation
    Set Self = Me
End Property

'@section ISetupPreparation implementation
'===============================================================================


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
'@Description("Bind the preparation helper to the workbook hosting setup worksheets.")
Public Sub Configure(ByVal hostBook As Workbook)
    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook reference is required"
    End If

    Set this.hostBook = hostBook
    ResetCaches
End Sub

'@Description("Execute the preparation workflow: register dropdowns then initialise the updated values registry.")
Public Sub Prepare(ByRef manager As IDevelopment)
    EnsureDropdowns
    EnsureUpdatedRegistry
    ApplyValidationsAndUpdates
    EnsureHiddenSheets manager
End Sub

'@Description("Ensure the dropdown manager has registered all default dropdown lists.")
Public Sub EnsureDropdowns()
    EnsureDropdownManager
    RegisterAllDropdowns
End Sub

'@Description("Initialise the updated values registry worksheet and register watched tables.")
Public Sub EnsureUpdatedRegistry()
    EnsureUpdatedValuesWatcher
    ResetUpdatedRegistry
    RegisterWatchedSheets
    ResetRegistryStatuses
End Sub


'@Description("Expose the dropdown manager for external callers.")
Public Property Get Dropdowns() As IDropdownLists
    EnsureDropdownManager
    Set Dropdowns = this.dropdowns
End Property

'@Description("Expose the updated values watcher for external callers.")
Public Property Get UpdatedWatcher() As IUpdatedValues
    EnsureUpdatedValuesWatcher
    Set UpdatedWatcher = this.updater
End Property

'@Description("Expose the workbook hosting setup assets.")
Public Property Get HostWorkbook() As Workbook
    Set HostWorkbook = EnsureWorkbook
End Property

'@section Preparation helpers
'===============================================================================
'@Description("Reset cached references after reconfiguring the helper.")
Private Sub ResetCaches()
    Set this.dropdowns = Nothing
    Set this.updater = Nothing
    Set this.dropSheet = Nothing
    Set this.registrySheet = Nothing
End Sub

'@Description("Ensure the dropdown manager is initialised for the target workbook.")
Private Sub EnsureDropdownManager()
    If Not this.dropdowns Is Nothing Then Exit Sub

    Dim targetSheet As Worksheet

    Set targetSheet = EnsureDropdownSheet
    Set this.dropdowns = DropdownLists.Create(targetSheet, DROPDOWN_PREFIX)
End Sub

'@Description("Ensure the updated values watcher is initialised for the registry worksheet.")
Private Sub EnsureUpdatedValuesWatcher()
    If Not this.updater Is Nothing Then Exit Sub

    Dim registry As Worksheet

    Set registry = EnsureRegistrySheet
    Set this.updater = UpdatedValues.Create(registry)
End Sub

'@Description("Reset the updated values registry prior to registering watched tables.")
Private Sub ResetUpdatedRegistry()
    If this.updater Is Nothing Then Exit Sub
    this.updater.DeleteUp
End Sub

Private Sub ApplyValidationsAndUpdates()
    Dim drop As IDropdownLists

    EnsureDropdownManager
    Set drop = this.dropdowns
    ApplyDictionaryValidations drop
    ApplyExportsValidations drop
    ApplyAnalysisValidations drop
End Sub

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

Private Sub EnsureProtectedSheets(ByRef manager As IDevelopment)

    If manager Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Development manager reference is required"
    End If

    manager.AddProtectedSheet SHEET_DICTIONARY
    manager.AddProtectedSheet SHEET_CHOICES
    manager.AddProtectedSheet SHEET_ANALYSIS
    manager.AddProtectedSheet SHEET_EXPORTS
    manager.AddProtectedSheet SHEET_TRANSLATIONS

End Sub


'@Description("Register every dropdown defined in PrepareSetup legacy module.")
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

'@Description("Register a single dropdown by converting the provided values into a BetterArray.")
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

'@Description("Register watched columns for every relevant setup sheet.")
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

'@Description("Return the list of setup sheets whose tables should be watched for updates.")
Private Function WatchedSheetNames() As BetterArray
    Dim names As BetterArray
    Set names = New BetterArray
    names.Push SHEET_DICTIONARY, SHEET_CHOICES, SHEET_ANALYSIS, SHEET_EXPORTS
    Set WatchedSheetNames = names
End Function

'@Description("Reset registry statuses to the default value after registration.")
Private Sub ResetRegistryStatuses()
    EnsureUpdatedValuesWatcher
    this.updater.SwitchTagsToYes
End Sub

'@section Worksheet helpers
'===============================================================================
'@Description("Ensure the dropdown worksheet exists and cache the reference.")
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

'@Description("Ensure the registry worksheet exists and cache the reference.")
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

'@Description("Return the cached workbook reference, raising when missing.")
Private Function EnsureWorkbook() As Workbook
    If this.hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "SetupPreparation has not been configured"
    End If
    Set EnsureWorkbook = this.hostBook
End Function

'@Description("Resolve an existing worksheet by name without creating it.")
Private Function ResolveWorksheet(ByVal sheetName As String) As Worksheet
    Dim wb As Workbook
    Dim sh As Worksheet

    Set wb = EnsureWorkbook

    On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0

    Set ResolveWorksheet = sh
End Function

'@Description("Resolve a worksheet if it exists, otherwise create one with the supplied visibility state.")
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

'@Description("Apply data validation rules to the dictionary worksheet tables.")
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

'@Description("Apply export worksheet validations.")
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

'@Description("Apply validations to analysis tables and named ranges.")
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

'@Description("Retrieve a worksheet or raise if unavailable.")
Private Function RequireWorksheet(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    Set sh = ResolveWorksheet(sheetName)
    If sh Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Worksheet '" & sheetName & "' not found in setup workbook"
    End If

    Set RequireWorksheet = sh
End Function

'@Description("Retrieve a ListObject from a worksheet or raise when missing.")
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

'@Description("Return a CustomTable helper for the requested ListObject.")
Private Function RequireCustomTable(ByVal sh As Worksheet, ByVal listName As String) As ICustomTable
    Dim lo As ListObject

    Set lo = RequireListObject(sh, listName)
    Set RequireCustomTable = CustomTable.Create(lo)
End Function

'@Description("Resolve a named range within a worksheet, returning Nothing when not available.")
Private Function ResolveNamedRange(ByVal sh As Worksheet, ByVal rangeName As String) As Range
    On Error Resume Next
        Set ResolveNamedRange = sh.Range(rangeName)
    On Error GoTo 0
End Function

'@Description("Normalise a string for case-insensitive comparisons.")
Private Function NormalizeText(ByVal valueText As String) As String
    NormalizeText = LCase$(Trim$(valueText))
End Function

'@section Error handling
'===============================================================================
'@Description("Raise a project-scoped error with the supplied message.")
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, "SetupPreparation", "Error: " & message
End Sub
