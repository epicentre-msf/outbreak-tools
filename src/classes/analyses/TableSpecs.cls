VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Analysis table specifications"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Analysis table specifications")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'@section Enumerations
'===============================================================================
'@description
'Analysis table scopes describing the kind of table a setup row defines.
'These live here because TableSpecs is what resolves them, so depending on
'TableSpecs resolves the enum everywhere.
Public Enum AnalysisTableScope
    ScopeGlobalSummary = 1
    ScopeUnivariate = 2
    ScopeBivariate = 3
    ScopeTimeSeries = 4
    ScopeSpatial = 5
    ScopeSpatioTemporal = 6
    ScopeTimeSeriesGraph = 7
End Enum


'@class TableSpecs
'@description
'Represents a single row from an analysis setup ListObject and exposes
'properties that describe the analysis table it defines. Each row in the
'setup table specifies one output table (frequency, cross-tab, time series,
'spatial, or spatio-temporal). The class wraps that row together with the
'header range and a dictionary, providing typed access to the table scope,
'identifiers, navigation to neighboring rows, validation flags (total,
'missing, percentage, graph), and dictionary-based category lookups.
'
'WHERE THE SCOPE COMES FROM
'-------------------------------------------------------------------------------
'The scope is the name of the ListObject the specification row sits in.
'The seven analysis tables carry seven fixed names and each name maps to one
'scope. An earlier version read a type label from a fixed row offset above
'the header. That offset is four rows in the workbook shipped from
'src/bin/setup and two rows in every published release, so on a released file
'the read landed on a blank cell and Create raised for every analysis row.
'The ListObject name holds the same fact and it survives a row being
'inserted.
'
'WHAT IS CACHED AND WHY IT IS SAFE
'-------------------------------------------------------------------------------
'Column names, cell values and the table scope are read once at construction
'time. ValidTable, IsNewSection, TableId, TableSectionId and
'SpatialTableScopes are computed on first access and held. Seal runs at the
'end of Create and stops the ranges changing after that, so a held answer
'cannot go stale within one instance's life.
'
'CATEGORY LOOKUPS
'-------------------------------------------------------------------------------
'RowCategories and ColumnCategories accept an Object parameter. When a
'LinelistSpecs instance is passed, categories are resolved from its
'dictionary and choices; any other object gives an empty BetterArray. That
'keeps TableSpecs usable in the setup error-checking context, where geo and
'translation dependencies are absent. LinelistSpecs is reached through a
'TypeName string, so it is never a compile dependency of this class.
'@depends LLdictionary, LLVariables, BetterArray, Checking
'@version 3.0 (2026-07-29)


Private Type TTableSpecs
    hRng As Range                                'Header range of the setup listobject
    tRng As Range                                'Row range for this table spec
    ' --- Column and cell value caches ---
    colNames As BetterArray                      'Column names from header (original case)
    cellValues As BetterArray                    'Parallel cell values from spec row
    ' --- Scalar caches ---
    cachedScope As Long                          'Cached AnalysisTableScope enum value
    scopeCached As Boolean                       'Whether cachedScope is populated
    dict As LLdictionary                         'Dictionary reference
    vars As LLVariables                          'Variable reader over the same dictionary
    varsResolved As Boolean                      'Whether vars has been resolved once
    cachedValid As Boolean                       'Cached ValidTable result
    validCached As Boolean                       'Whether cachedValid is populated
    cachedNewSection As Boolean                  'Cached IsNewSection result
    newSectionCached As Boolean                  'Whether cachedNewSection is populated
    cachedTableId As String                      'Cached TableId string
    tableIdCached As Boolean                     'Whether cachedTableId is populated
    cachedSectionId As String                    'Cached TableSectionId string
    sectionIdCached As Boolean                   'Whether cachedSectionId is populated
    cachedSpatialScope As String                 'Cached SpatialTableScopes answer
    spatialScopeCached As Boolean                'Whether cachedSpatialScope is populated
    sealed As Boolean
End Type

Private this As TTableSpecs
Private Const CLASS_NAME As String = "TableSpecs"

' Names of the seven analysis ListObjects on the setup worksheet. Each name
' is one table scope. Six other classes hold the same strings beside the code
' that reads them, which is the house rule: a class keeps every constant it
' needs inside itself.
Private Const TABLE_GLOBAL_SUMMARY As String = "Tab_Global_Summary"
Private Const TABLE_UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const TABLE_BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TABLE_TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const TABLE_SPATIAL As String = "Tab_Spatial_Analysis"
Private Const TABLE_SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const TABLE_GRAPH_TIMESERIES As String = "Tab_Graph_TimeSeries"

' Column search terms. Every one is a partial, case-insensitive match against
' the header row, so "row" finds "Group by variable (row)".
Private Const COLUMN_SECTION As String = "section"
Private Const COLUMN_ROW As String = "row"
Private Const COLUMN_COLUMN As String = "column"
Private Const COLUMN_TOTAL As String = "total"
Private Const COLUMN_MISSING As String = "missing"
Private Const COLUMN_PERCENTAGE As String = "percentage"
Private Const COLUMN_GRAPH As String = "graph"
Private Const COLUMN_LABEL As String = "label"
Private Const COLUMN_FUNCTION As String = "function"
Private Const COLUMN_NGEO As String = "n geo"
Private Const COLUMN_GRAPH_TITLE As String = "graph title"
Private Const COLUMN_SERIES_TITLE As String = "series title"

' Column names read through LLVariables, on the dictionary sheet.
Private Const VARIABLE_CONTROL As String = "control"
Private Const VARIABLE_TYPE As String = "variable type"

' Geo prefixes written by LLdictionary.AppendGeoLines.
Private Const PREFIX_ADMIN As String = "adm1_"
Private Const PREFIX_HEALTH_FACILITY As String = "hf_"

' Bounds for the "N geo max" column, which is free text on the setup sheet.
Private Const DEFAULT_GEO_COUNT As Long = 5
Private Const MIN_GEO_COUNT As Long = 1
Private Const MAX_GEO_COUNT As Long = 20

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a TableSpecs instance from a setup row
'@details
'Factory method that creates a new TableSpecs instance representing one row
'from an analysis setup ListObject. Each row in the setup table defines a
'single analysis table (e.g. a univariate frequency table or a time series).
'The TableSpecs object wraps that row and exposes properties to query its
'scope, validity, navigation to neighboring rows, and display flags (total,
'missing, percentage, graph). Category lookups (RowCategories,
'ColumnCategories) are resolved later at call time by passing a
'LinelistSpecs object. Validation is performed immediately via
'CheckRequirements before the instance is populated. After property
'injection, InitCache reads the header and the data row with two array
'reads and resolves the table scope from the ListObject name.
'@param headerRng Range. The single header row Range of the analysis setup
'ListObject. Column names in this range are used by Value() to locate fields
'such as "row", "column", "total", "section", etc.
'@param tableSpecsRng Range. A single data-row Range from the same
'ListObject. Must sit below headerRng, start at the same column and have the
'same number of columns.
'@param dict LLdictionary. The linelist dictionary that provides variable
'definitions, existence checks, and control type lookups needed for
'validation and spatial type detection.
'@param vars LLVariables. Optional reader over the same dictionary. Pass one
'when the caller already holds it: an analysis run then builds a single
'LLVariables for every row instead of one per row and per category lookup.
'When it is omitted the class builds its own on first use.
'@return TableSpecs. A fully initialised TableSpecs instance ready for use.
'@throws InvalidArgument When headerRng is Nothing.
'@throws InvalidArgument When tableSpecsRng is Nothing.
'@throws InvalidArgument When dict is Nothing.
'@throws InvalidArgument When headerRng and tableSpecsRng have different column counts.
'@throws InvalidArgument When headerRng and tableSpecsRng start at different columns.
'@throws InvalidArgument When either range spans more than one row.
'@throws InvalidArgument When headerRng.Row is not above tableSpecsRng.Row.
'@throws ErrorUnexpectedState When the specification row sits outside a known analysis table.
'@export
Public Function Create(ByVal headerRng As Range, _
                       ByVal tableSpecsRng As Range, _
                       ByVal dict As LLdictionary, _
                       Optional ByVal vars As LLVariables) As TableSpecs

    Dim instance As TableSpecs
    CheckRequirements headerRng, tableSpecsRng, dict

    Set instance = New TableSpecs
    Set instance.HeaderRange = headerRng
    Set instance.TableRange = tableSpecsRng
    Set instance.DictionaryObject = dict
    Set instance.Variables = vars
    instance.InitCache
    instance.Seal
    Set Create = instance
End Function

'@label:seal
'@sub-title Seal the instance against further setup writes.
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup routine against post-creation calls.
'@param propName String. The property name for the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@section Internal Properties
'===============================================================================
'@description Properties used by the factory pattern during initialisation.

'@label:header-range
'@prop-title Header range of the setup ListObject
'@details
'Returns the header row Range that contains column names such as "row",
'"column", "total", "section", etc. Used by Value() for all field lookups.
'@return Range. The header row range.
Public Property Get HeaderRange() As Range
    Set HeaderRange = this.hRng
End Property

'@label:header-range-set
'@prop-title Assign the header range
'@param hRng Range. The header row range to store.
Public Property Set HeaderRange(ByVal hRng As Range)
    GuardNotSealed "HeaderRange"
    Set this.hRng = hRng
End Property

'@label:table-range
'@prop-title Table specification row range
'@details
'Returns the single data-row Range from the setup ListObject that
'defines this table specification.
'@return Range. The table specification row range.
Public Property Get TableRange() As Range
    Set TableRange = this.tRng
End Property

'@label:table-range-set
'@prop-title Assign the table specification row range
'@param tRng Range. The data-row range to store.
Public Property Set TableRange(ByVal tRng As Range)
    GuardNotSealed "TableRange"
    Set this.tRng = tRng
End Property

'@label:dictionary-object-set
'@prop-title Assign the dictionary object
'@details
'The read side of this pair is the Dictionary property below, which is the
'one every consumer uses.
'@param dict LLdictionary. The dictionary to store.
Public Property Set DictionaryObject(ByVal dict As LLdictionary)
    GuardNotSealed "DictionaryObject"
    Set this.dict = dict
End Property

'@label:variables-set
'@prop-title Assign the variable reader handed in by the caller
'@details
'Set from the factory with whatever Create received. Passing Nothing leaves
'the class to build its own reader on first use.
'@param vars LLVariables. The variable reader over the same dictionary.
Public Property Set Variables(ByVal vars As LLVariables)
    GuardNotSealed "Variables"
    Set this.vars = vars
    this.varsResolved = Not (vars Is Nothing)
End Property

'@section Cache Initialization
'===============================================================================
'@description One-time cache population called from the factory.

'@label:init-cache
'@sub-title Populate the column and value caches and resolve the scope
'@details
'Reads the header range and the specification row with one array read each
'and builds parallel BetterArray caches of column names and cell values.
'Then resolves the table scope from the name of the ListObject the row sits
'in. Runs exactly once, from Create, before Seal: a later call raises,
'because the derived caches held elsewhere in the class are computed from
'this data.
'@throws SomethingWentWrong When called after the instance is sealed.
'@throws ErrorUnexpectedState When the specification row sits outside a known analysis table.
'@export
Public Sub InitCache()
    Dim headerValues As Variant
    Dim rowValues As Variant
    Dim colCount As Long
    Dim idx As Long

    GuardNotSealed "InitCache"

    Set this.colNames = New BetterArray
    Set this.cellValues = New BetterArray
    this.colNames.LowerBound = 1
    this.cellValues.LowerBound = 1

    colCount = this.hRng.Columns.Count

    ' Two crossings into Excel for the whole row. A single-column range
    ' answers a scalar, so that shape is read on its own.
    headerValues = this.hRng.Value
    rowValues = this.tRng.Value

    If colCount = 1 Then
        this.colNames.Push CellText(headerValues)
        this.cellValues.Push CellText(rowValues)
    Else
        For idx = 1 To colCount
            this.colNames.Push CellText(headerValues(1, idx))
            this.cellValues.Push CellText(rowValues(1, idx))
        Next idx
    End If

    this.cachedScope = ComputeTableScope()
    this.scopeCached = True
End Sub

'@section Core Properties
'===============================================================================
'@description Type identification and identity properties for each table spec.

'@label:table-scope
'@prop-title Analysis table scope for this specification row
'@details
'Returns the cached AnalysisTableScope enum value. The scope is resolved
'once during InitCache from the name of the ListObject the specification row
'sits in. On subsequent calls the cached Long value is returned directly
'without any worksheet access.
'@return AnalysisTableScope. The table scope category.
'@export
Public Property Get TableScope() As AnalysisTableScope
    If this.scopeCached Then
        TableScope = this.cachedScope
    Else
        this.cachedScope = ComputeTableScope()
        this.scopeCached = True
        TableScope = this.cachedScope
    End If
End Property

'@label:table-id
'@prop-title Unique identifier for this analysis table
'@details
'Generates a unique string identifier by combining a short type-based
'prefix with the row offset from the header. The prefix is derived from
'the TableScope: "GS" for GlobalSummary, "UA" for Univariate, "BA" for
'Bivariate, "TS" for TimeSeries, "SA" for Spatial, "SPT" for
'SpatioTemporal and "TSG" for the time series graph block. The row offset is
'computed as TableRange.Row minus HeaderRange.Row, producing a stable
'integer that uniquely locates this row within its analysis block. The
'resulting format is "prefix_tabN", for example "UA_tab3". Seal fixes both
'ranges, so the answer is computed once and held.
'@return String. The unique table identifier such as "GS_tab1" or "BA_tab5".
'@throws ErrorUnexpectedState When the table scope has no prefix.
'@export
Public Property Get TableId() As String
    Dim prefix As String

    If this.tableIdCached Then
        TableId = this.cachedTableId
        Exit Property
    End If

    Select Case TableScope
    Case ScopeGlobalSummary
        prefix = "GS"
    Case ScopeUnivariate
        prefix = "UA"
    Case ScopeBivariate
        prefix = "BA"
    Case ScopeTimeSeries
        prefix = "TS"
    Case ScopeSpatial
        prefix = "SA"
    Case ScopeSpatioTemporal
        prefix = "SPT"
    Case ScopeTimeSeriesGraph
        prefix = "TSG"
    Case Else
        ThrowError ErrorUnexpectedState, _
                   "No table id prefix for scope " & TableScope
    End Select

    this.cachedTableId = prefix & "_tab" & (this.tRng.Row - this.hRng.Row)
    this.tableIdCached = True
    TableId = this.cachedTableId
End Property

'@label:table-section-id
'@prop-title Identifier of the first table in this section
'@details
'Returns the TableId of the first table in the current section group by
'walking backward through the Previous chain. If this table is itself the
'start of a new section (IsNewSection = True) or is a GlobalSummary table,
'it returns its own TableId immediately. The same answer is given when
'Previous finds no valid table left in the section. Otherwise it delegates
'to Previous.TableSectionId, which continues walking backward until a
'section boundary is reached. All tables within the same section share a
'common section identifier.
'@return String. The TableId of the first table in this section.
'@remarks This is a recursive property. Resolving the section id of the k-th
'table in a section builds k instances, and a caller that walks a whole
'block row by row pays that once per row. A caller iterating in order can
'carry the section id in a local and avoid the walk entirely.
'@export
Public Property Get TableSectionId() As String
    Dim prevSpec As TableSpecs

    If this.sectionIdCached Then
        TableSectionId = this.cachedSectionId
        Exit Property
    End If

    If IsNewSection() Or (TableScope = ScopeGlobalSummary) Then
        this.cachedSectionId = TableId()
    Else
        Set prevSpec = Previous()
        If prevSpec Is Nothing Then
            this.cachedSectionId = TableId()
        Else
            this.cachedSectionId = prevSpec.TableSectionId()
        End If
    End If

    this.sectionIdCached = True
    TableSectionId = this.cachedSectionId
End Property

'@label:dictionary
'@prop-title Linelist dictionary bound to this specification
'@details
'Returns the LLdictionary reference stored at construction time. The
'dictionary provides variable existence checks, control type lookups,
'and spatial prefix detection used throughout validation and category
'resolution.
'@return LLdictionary. The linelist dictionary instance.
'@export
Public Property Get Dictionary() As LLdictionary
    Set Dictionary = this.dict
End Property

'@label:geo-count
'@prop-title Number of geographic units this table shows
'@details
'Reads the "N geo max" column and answers a usable count. The column is free
'text on the setup sheet, so the value is clamped: anything that is not a
'number, and an empty cell, answer the default of 5; a value below 1 answers
'1 and a value above 20 answers 20. Without the clamp a "0" builds a table
'with no columns and a large value builds a table that wide.
'@return Long. The number of geographic units, between 1 and 20.
'@export
Public Property Get GeoCount() As Long
    Dim rawValue As String
    Dim count As Long

    count = DEFAULT_GEO_COUNT
    rawValue = Trim$(Value(COLUMN_NGEO))

    If IsNumeric(rawValue) Then count = CLng(Val(rawValue))

    If count < MIN_GEO_COUNT Then count = MIN_GEO_COUNT
    If count > MAX_GEO_COUNT Then count = MAX_GEO_COUNT

    GeoCount = count
End Property

'@label:spatial-table-types
'@prop-title Spatial analysis sub-type
'@details
'Determines the spatial sub-type of a Spatial or SpatioTemporal table by
'checking whether the relevant variable has a health-facility or
'administrative geography prefix in the dictionary. For Spatial tables, the
'variable comes from the "row" field; for SpatioTemporal tables, the spatial
'variable is in the "column" field. The method prepends "hf_" or "adm1_" to
'the variable name and checks whether that prefixed variable exists in the
'dictionary. Returns "hf" for health facility, "geo" for administrative
'geography, or vbNullString if neither prefix is found. The answer is held
'after the first call: it costs up to two dictionary searches and nine
'production call sites ask for it.
'@return String. The spatial sub-type ("hf", "geo", or empty string).
'@remarks This property is meaningful only for Spatial and SpatioTemporal
'table scopes. Calling it on other table scopes will still execute but will
'likely return vbNullString since those variables typically lack geo/hf
'prefixes.
'@export
Public Property Get SpatialTableScopes() As String
    Dim rowVar As String
    Dim cachedDict As LLdictionary

    If this.spatialScopeCached Then
        SpatialTableScopes = this.cachedSpatialScope
        Exit Property
    End If

    Set cachedDict = Dictionary()

    ' For spatio-temporal, the spatial variable is in the "column" field
    If TableScope = ScopeSpatioTemporal Then
        rowVar = Value(COLUMN_COLUMN)
    Else
        rowVar = Value(COLUMN_ROW)
    End If

    this.cachedSpatialScope = Switch( _
        cachedDict.VariableExists(PREFIX_HEALTH_FACILITY & rowVar), "hf", _
        cachedDict.VariableExists(PREFIX_ADMIN & rowVar), "geo", _
        True, vbNullString)
    this.spatialScopeCached = True

    SpatialTableScopes = this.cachedSpatialScope
End Property

'@section Validation Properties
'===============================================================================
'@description Properties that validate table configuration and detect section boundaries.

'@label:valid-table
'@prop-title Whether this table has a valid configuration
'@details
'Returns the cached validation result. On first access, validates whether
'this table specification row contains all required fields and references
'valid variables for its analysis type. Each table scope has different
'validation rules: GlobalSummary requires non-empty "label" and "function"
'fields; Univariate requires the "row" variable to be a choice variable;
'Bivariate requires both "row" and "column" to be choice variables;
'TimeSeries requires the "row" variable to be of type "date"; Spatial
'requires a geographic prefix in the dictionary; SpatioTemporal requires
'a geographic prefix on the column variable and a date-type row variable;
'the time series graph block requires a graph title and a series title, the
'two columns the user picks from a dropdown. Invalid table specs are skipped
'during navigation (Previous, NextSpecs).
'@return Boolean. True if the table specification meets all requirements for its type.
'@depends LLVariables
'@export
Public Property Get ValidTable() As Boolean
    If this.validCached Then
        ValidTable = this.cachedValid
    Else
        this.cachedValid = ComputeValidTable()
        this.validCached = True
        ValidTable = this.cachedValid
    End If
End Property

'@label:is-new-section
'@prop-title Whether this table begins a new section group
'@details
'Returns the cached section-boundary result. On first access, determines
'whether this table specification row begins a new section within the analysis
'setup block. Sections group related tables together. The method checks the
'"section" column: if the current row's section value differs from the
'previous row's section value, or if this is the first data row (where the
'previous row is the header), the table is considered a new section.
'GlobalSummary tables are never considered new sections regardless of their
'section value, because each GlobalSummary table is standalone and does not
'participate in section grouping.
'@return Boolean. True if this table starts a new section.
'@remarks The comparison uses the cell at row offset 0 (prevCell =
'tRng.Cells(0, sectIndex)), which in Excel Range addressing corresponds
'to the row immediately above the current row. The section column is free
'text, so the two values are compared without regard to case: a
'capitalisation slip would otherwise split one section in two and every
'section-keyed named range would follow it.
'@export
Public Property Get IsNewSection() As Boolean
    If this.newSectionCached Then
        IsNewSection = this.cachedNewSection
    Else
        this.cachedNewSection = ComputeIsNewSection()
        this.newSectionCached = True
        IsNewSection = this.cachedNewSection
    End If
End Property

'@section Flag Properties
'===============================================================================
'@description Properties controlling optional table features: totals, missing,
'percentage, and graphs.

'@label:has-total
'@prop-title Whether totals are needed for computation
'@details
'Determines whether a total row/column is needed for this table. The logic
'varies by table scope: GlobalSummary, SpatioTemporal and the time series
'graph block always return False; Univariate and Bivariate always return
'True; TimeSeries returns True when the user explicitly requested
'total="yes" or when percentage is set to "row" or "column" (because
'percentage formulas need the total column as a denominator); Spatial
'returns True only when a column variable is specified. This is distinct
'from TotalRequested: HasTotal can be True even when the user did not ask
'for totals, solely because percentage needs them.
'@return Boolean. True if a total row/column should be created.
'@throws ErrorUnexpectedState When the table scope is unmapped.
'@export
Public Property Get HasTotal() As Boolean
    Select Case TableScope

    Case ScopeGlobalSummary, ScopeSpatioTemporal, ScopeTimeSeriesGraph
        HasTotal = False

    Case ScopeUnivariate
        HasTotal = True

    Case ScopeBivariate
        HasTotal = True

    Case ScopeTimeSeries
        HasTotal = ((FlagValue(COLUMN_TOTAL) = "yes") Or _
                    (FlagValue(COLUMN_PERCENTAGE) = "row") Or _
                    (FlagValue(COLUMN_PERCENTAGE) = "column")) And _
                   (Value(COLUMN_COLUMN) <> vbNullString)

    Case ScopeSpatial
        HasTotal = (Value(COLUMN_COLUMN) <> vbNullString)

    Case Else
        ThrowError ErrorUnexpectedState, _
                   "HasTotal has no rule for scope " & TableScope

    End Select
End Property

'@label:total-requested
'@prop-title Whether total was explicitly requested by the user
'@details
'Returns True only when the user explicitly wrote "yes" in the "total"
'column of the setup row. This is the companion to HasTotal for the
'visibility logic: when HasTotal is True but TotalRequested is False,
'the total column is created for percentage computation but hidden from
'the user's view. If TotalRequested is True, the total column is both
'created and visible.
'@return Boolean. True if the "total" field equals "yes".
'@remarks For Univariate and Bivariate tables, TotalRequested is not
'consulted because HasTotal is unconditionally True and totals are always
'shown. Its one production caller reads it for time series tables.
'@export
Public Property Get TotalRequested() As Boolean
    TotalRequested = (FlagValue(COLUMN_TOTAL) = "yes")
End Property

'@label:has-missing
'@prop-title Whether missing-data rows or columns are included
'@details
'Determines whether a "missing" row or column should be included in the
'output table. The logic varies by table scope: GlobalSummary, SpatioTemporal
'and the time series graph block always return False; Univariate returns True
'when missing="yes"; Bivariate returns True when missing="row", "column", or
'"all"; TimeSeries and Spatial return True when missing="yes" and a column
'variable is present (without a column variable, there is no categorical
'axis to show missing values on).
'@return Boolean. True if missing data should be shown in the output table.
'@throws ErrorUnexpectedState When the table scope is unmapped.
'@export
Public Property Get HasMissing() As Boolean
    Select Case TableScope

    Case ScopeGlobalSummary, ScopeSpatioTemporal, ScopeTimeSeriesGraph
        HasMissing = False

    Case ScopeUnivariate
        HasMissing = (FlagValue(COLUMN_MISSING) = "yes")

    Case ScopeBivariate
        HasMissing = (FlagValue(COLUMN_MISSING) = "row") Or _
                     (FlagValue(COLUMN_MISSING) = "column") Or _
                     (FlagValue(COLUMN_MISSING) = "all")

    Case ScopeTimeSeries
        HasMissing = (FlagValue(COLUMN_MISSING) = "yes") And _
                     (Value(COLUMN_COLUMN) <> vbNullString)

    Case ScopeSpatial
        HasMissing = (FlagValue(COLUMN_MISSING) = "yes") And _
                     (Value(COLUMN_COLUMN) <> vbNullString)

    Case Else
        ThrowError ErrorUnexpectedState, _
                   "HasMissing has no rule for scope " & TableScope

    End Select
End Property

'@label:has-percentage
'@prop-title Whether percentage display is enabled
'@details
'Determines whether percentage values should be computed and displayed for
'this table. The logic varies by table scope: GlobalSummary, SpatioTemporal
'and the time series graph block always return False; Univariate returns True
'when percentage="yes"; Bivariate returns True when percentage="row",
'"column", or "total"; TimeSeries returns True when percentage="row" or
'"column" and HasTotal is True; Spatial returns True when percentage="yes"
'and HasTotal is True. The interplay between HasPercentage and HasTotal for
'TimeSeries is intentional: requesting percentage implicitly forces
'HasTotal=True since percentage needs the total.
'@return Boolean. True if percentage values should be included.
'@throws ErrorUnexpectedState When the table scope is unmapped.
'@export
Public Property Get HasPercentage() As Boolean
    Select Case TableScope

    Case ScopeGlobalSummary, ScopeSpatioTemporal, ScopeTimeSeriesGraph
        HasPercentage = False

    Case ScopeUnivariate
        HasPercentage = (FlagValue(COLUMN_PERCENTAGE) = "yes")

    Case ScopeBivariate
        HasPercentage = (FlagValue(COLUMN_PERCENTAGE) = "row") Or _
                        (FlagValue(COLUMN_PERCENTAGE) = "column") Or _
                        (FlagValue(COLUMN_PERCENTAGE) = "total")

    Case ScopeTimeSeries
        HasPercentage = ((FlagValue(COLUMN_PERCENTAGE) = "row") Or _
                         (FlagValue(COLUMN_PERCENTAGE) = "column")) And HasTotal()

    Case ScopeSpatial
        HasPercentage = (FlagValue(COLUMN_PERCENTAGE) = "yes") And HasTotal()

    Case Else
        ThrowError ErrorUnexpectedState, _
                   "HasPercentage has no rule for scope " & TableScope

    End Select
End Property

'@label:has-graph
'@prop-title Whether a graph should be created for this table
'@details
'Determines whether a chart/graph should be generated alongside the output
'table. The logic varies by table scope: GlobalSummary and the time series
'graph block always return False; Univariate, TimeSeries, Spatial, and
'SpatioTemporal return True when graph="yes"; Bivariate returns True when
'graph="percentages" or "values", the two entries of the __perc_val dropdown
'the setup binds to that column. The actual graph rendering is handled
'downstream by the chart builder; this property only signals whether one
'should be created. "both" stays in the bivariate list because GraphSpecs
'reads it too; the dropdown offers "percentages" and "values".
'@return Boolean. True if a graph should be created for this table.
'@remarks The time series graph block answers False. Every one of its rows
'is already a graph, built by WriteTimeSeriesGraphs, and the term "graph"
'matches its "Graph title (select)" column, which holds a title.
'@throws ErrorUnexpectedState When the table scope is unmapped.
'@export
Public Property Get HasGraph() As Boolean
    Select Case TableScope

    Case ScopeGlobalSummary, ScopeTimeSeriesGraph
        HasGraph = False

    Case ScopeUnivariate
        HasGraph = (FlagValue(COLUMN_GRAPH) = "yes")

    Case ScopeBivariate
        HasGraph = (FlagValue(COLUMN_GRAPH) = "percentages" Or _
                    FlagValue(COLUMN_GRAPH) = "values" Or _
                    FlagValue(COLUMN_GRAPH) = "both")

    Case ScopeTimeSeries
        HasGraph = (FlagValue(COLUMN_GRAPH) = "yes")

    Case ScopeSpatial
        HasGraph = (FlagValue(COLUMN_GRAPH) = "yes")

    Case ScopeSpatioTemporal
        HasGraph = (FlagValue(COLUMN_GRAPH) = "yes")

    Case Else
        ThrowError ErrorUnexpectedState, _
                   "HasGraph has no rule for scope " & TableScope

    End Select
End Property

'@section Navigation Properties
'===============================================================================
'@description Linked-list navigation between sibling specifications.

'@label:previous
'@prop-title Previous valid specification in the same section
'@details
'Returns the nearest valid table specification that precedes this one within
'the same section. Starting from the row immediately above the current table
'range, it walks backward one row at a time. Every candidate, including the
'first, is tested the same way: a valid table is returned, and a candidate
'that starts a new section ends the walk. The walk also stops at the header
'row, to stay inside the data area.
'@return TableSpecs. The nearest valid preceding table in the same section,
'or Nothing when this row starts a section and when the section holds no
'valid table above this row.
'@remarks Each candidate row examined creates a new TableSpecs instance, so
'a long run of invalid rows costs one instance per row.
'@export
Public Property Get Previous() As TableSpecs
    Dim hRng As Range
    Dim prevRng As Range
    Dim candidate As TableSpecs

    If IsNewSection() Then Exit Property

    Set hRng = HeaderRange()
    Set prevRng = TableRange().Offset(-1)

    Do While prevRng.Row > hRng.Row
        Set candidate = Create(hRng, prevRng, Dictionary(), VariablesObject())

        If candidate.ValidTable() Then
            Set Previous = candidate
            Exit Property
        End If

        If candidate.IsNewSection() Then Exit Property

        Set prevRng = prevRng.Offset(-1)
    Loop
End Property

'@label:next-specs
'@prop-title Next valid specification following an anchor range
'@details
'Returns the next valid table specification after this row by walking forward
'through the setup ListObject. Starting from the row immediately below the
'current table range, it creates a TableSpecs for each row until it finds a
'valid table or passes the anchor range. The walk crosses section
'boundaries. If no valid table is found within the data area, the property
'returns Nothing.
'@param anchorRng Range. A Range representing the last row of the setup
'ListObject. The forward walk will not advance past this row.
'@return TableSpecs. The next valid table after this row, or Nothing if no
'valid table exists between this row and the anchor range.
'@remarks This property is the primary iteration mechanism for processing all
'tables in an analysis block. Callers typically start with the first table and
'call NextSpecs repeatedly until Nothing is returned.
'@export
Public Property Get NextSpecs(ByVal anchorRng As Range) As TableSpecs
    Dim hRng As Range
    Dim nextRng As Range
    Dim candidate As TableSpecs

    Set hRng = HeaderRange()
    Set nextRng = TableRange().Offset(1)

    Do While nextRng.Row <= anchorRng.Row
        Set candidate = Create(hRng, nextRng, Dictionary(), VariablesObject())

        If candidate.ValidTable() Then
            Set NextSpecs = candidate
            Exit Property
        End If

        Set nextRng = nextRng.Offset(1)
    Loop
End Property

'@section Data Access
'===============================================================================
'@description Value retrieval and category lookups from the specification row.

'@label:value
'@sub-title Retrieve a value from the specification row by column name
'@details
'Retrieves the cached cell value from this table specification row for a
'given column name. By default the column is located by scanning the cached
'column names for a case-insensitive substring match, so "row" finds
'"Group by variable (row)". Pass strictSearch:=True to require the whole
'header to match the term, which is what a caller wants when two headers on
'the same sheet share a word. Once found, the corresponding value is
'returned from the parallel value cache without any worksheet access. A
'column name absent from the header returns vbNullString.
'@param colName String. The column name to search for in the header.
'@param strictSearch Boolean. Optional. True compares the whole header
'against colName. The default False keeps the substring match every existing
'caller relies on.
'@return String. The cell value, or vbNullString if the column is not found.
'@export
Public Function Value(ByVal colName As String, _
                      Optional ByVal strictSearch As Boolean = False) As String
    Dim idx As Long

    idx = ColumnIndex(colName, strictSearch)
    If idx > 0 Then Value = this.cellValues.Item(idx)
End Function

'@label:row-categories
'@prop-title Row variable categories from the dictionary
'@details
'Returns a BetterArray of category values for the row variable of this table.
'When lData is a LinelistSpecs instance, the categories are retrieved from
'it based on the variable name in the "row" field. When lData is any other
'type (or Nothing), returns an empty BetterArray. Internally delegates to
'CategoriesData("row", lData).
'@param lData Object. A LinelistSpecs instance providing category lookups,
'or any other Object (categories will be empty).
'@return BetterArray. The category strings for the row variable, or an empty
'BetterArray if the row variable is not found or lData is not LinelistSpecs.
'@export
Public Property Get RowCategories(ByRef lData As Object) As BetterArray
    Set RowCategories = CategoriesData(COLUMN_ROW, lData)
End Property

'@label:column-categories
'@prop-title Column variable categories from the dictionary
'@details
'Returns a BetterArray of category values for the column variable of this
'table. When lData is a LinelistSpecs instance, the categories come from it
'based on the variable name in the "column" field. When lData is any other
'type (or Nothing), returns an empty BetterArray. For SpatioTemporal tables,
'the column categories are empty placeholder strings whose count comes from
'GeoCount. These placeholders are filled in later at runtime with actual
'geographic location names. Internally delegates to
'CategoriesData("column", lData).
'@param lData Object. A LinelistSpecs instance providing category lookups,
'or any other Object (categories will be empty).
'@return BetterArray. The category strings for the column variable. For
'SpatioTemporal, contains GeoCount empty strings. Returns an empty
'BetterArray if the column variable is not found or not specified.
'@export
Public Property Get ColumnCategories(ByRef lData As Object) As BetterArray
    Set ColumnCategories = CategoriesData(COLUMN_COLUMN, lData)
End Property

'@section Helpers
'===============================================================================
'@description Private helper methods that support data access and validation.

'@label:cell-text
'@sub-title Coerce one cell value to text, surviving an error value
'@details
'A cell holding #REF!, #N/A, #VALUE! or #NAME? answers an Error variant, and
'CStr on an Error variant raises runtime error 13. The analysis setup sheet
'carries formula columns, so an ordinary edit can produce one. An error
'value, a Null and an empty cell all answer an empty string here.
'@param cellValue Variant. One value read out of a worksheet range.
'@return String. The text of the cell, or vbNullString.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@label:flag-value
'@sub-title Read a flag column, trimmed and lower-cased
'@details
'The flag columns carry values picked from a dropdown, so the normal path is
'already clean. A value pasted into the cell bypasses Excel validation, so
'every flag comparison goes through here and reads " Yes " the same way it
'reads "yes".
'@param colName String. The column name to search for in the header.
'@return String. The flag value, trimmed and lower-cased.
Private Function FlagValue(ByVal colName As String) As String
    FlagValue = LCase$(Trim$(Value(colName)))
End Function

'@label:variables-object
'@sub-title The variable reader over this specification's dictionary
'@details
'Answers the LLVariables handed to Create, or builds one on first use and
'holds it for the life of the instance. LLVariables.Create requires a
'fully-configured dictionary (a Data sheet and a "variable name" column). The
'setup error-checking context may not expose those, so a failure here answers
'Nothing and the callers fall back. The attempt is made once.
'@return LLVariables. The variable reader, or Nothing when one cannot be built.
Private Function VariablesObject() As LLVariables
    If this.varsResolved Then
        Set VariablesObject = this.vars
        Exit Function
    End If

    On Error Resume Next
    Set this.vars = LLVariables.Create(Dictionary())
    On Error GoTo 0

    this.varsResolved = True
    Set VariablesObject = this.vars
End Function

'@label:column-index
'@sub-title Return the 1-based column index of a named header column
'@details
'Scans the cached column names for the search term. The default is a
'case-insensitive substring match; strictSearch compares the whole header.
'Returns the 1-based position within the header range, or -1 if not found.
'@param colName String. The column name to search for.
'@param strictSearch Boolean. True compares the whole header against colName.
'@return Long. The 1-based position within the header range, or -1 if not found.
Private Function ColumnIndex(ByVal colName As String, _
                             Optional ByVal strictSearch As Boolean = False) As Long
    Dim idx As Long

    For idx = 1 To this.colNames.Length
        If strictSearch Then
            If StrComp(this.colNames.Item(idx), colName, vbTextCompare) = 0 Then
                ColumnIndex = idx
                Exit Function
            End If
        ElseIf InStr(1, this.colNames.Item(idx), colName, vbTextCompare) > 0 Then
            ColumnIndex = idx
            Exit Function
        End If
    Next idx

    ColumnIndex = -1
End Function

'@label:categories-data
'@sub-title Build a BetterArray of categories for a row or column variable
'@details
'Shared implementation for RowCategories and ColumnCategories. Reads the
'variable name from the specified field using Value(rowOrCol), then checks
'whether lData is a LinelistSpecs instance via TypeName. If it is, the
'categories are retrieved from lData.Categories(). If not, returns an empty
'BetterArray. Handles a special case for SpatioTemporal tables: when
'rowOrCol is "column", the categories are populated with GeoCount empty
'strings. These empty placeholders represent geographic locations filled in
'at runtime. The resulting BetterArray is cloned before returning to prevent
'external mutations from affecting internal state.
'@param rowOrCol String. Either "row" or "column", indicating which variable
'field to read from the setup row.
'@param lData Object. A LinelistSpecs instance providing category lookups,
'or any other Object (categories will be empty).
'@return BetterArray. A cloned array containing the category strings.
'@depends LLVariables
Private Function CategoriesData(ByVal rowOrCol As String, _
                                ByRef lData As Object) As BetterArray
    Dim varsCat As BetterArray
    Dim nameVar As String
    Dim vars As LLVariables
    Dim counter As Long

    Set varsCat = New BetterArray
    varsCat.LowerBound = 1

    ' The spatio-temporal column path fills placeholders and reads no
    ' variable, so the reader is resolved only on the other branch.
    If (TableScope = ScopeSpatioTemporal) And (rowOrCol = COLUMN_COLUMN) Then
        For counter = 1 To GeoCount()
            varsCat.Push vbNullString
        Next
        Set CategoriesData = varsCat.Clone()
        Exit Function
    End If

    Set vars = VariablesObject()
    If vars Is Nothing Then
        Set CategoriesData = varsCat.Clone()
        Exit Function
    End If

    nameVar = Value(rowOrCol)

    If vars.Contains(nameVar) Then
        ' Only resolve categories when lData is a LinelistSpecs instance
        If Not lData Is Nothing Then
            If TypeName(lData) = "LinelistSpecs" Then
                Set varsCat = lData.Categories(nameVar)
            End If
        End If
    End If

    Set CategoriesData = varsCat.Clone()
End Function

'@label:compute-table-scope
'@sub-title Compute the AnalysisTableScope from the ListObject name
'@details
'Reads the name of the ListObject the specification row sits in and maps it
'to the matching AnalysisTableScope constant. Called once during InitCache to
'populate the scope cache. The seven names are fixed by the setup template
'and the same in the mock, in the shipped file and in every release build.
'@return AnalysisTableScope. The table scope category.
'@throws ErrorUnexpectedState When the row sits outside a ListObject.
'@throws ErrorUnexpectedState When the ListObject is not one of the seven analysis tables.
Private Function ComputeTableScope() As AnalysisTableScope
    Dim specsTable As ListObject

    Set specsTable = this.tRng.ListObject

    If specsTable Is Nothing Then
        ThrowError ErrorUnexpectedState, _
                   "The specification row sits outside an analysis table"
    End If

    Select Case LCase$(specsTable.Name)

    Case LCase$(TABLE_GLOBAL_SUMMARY)
        ComputeTableScope = ScopeGlobalSummary
    Case LCase$(TABLE_UNIVARIATE)
        ComputeTableScope = ScopeUnivariate
    Case LCase$(TABLE_BIVARIATE)
        ComputeTableScope = ScopeBivariate
    Case LCase$(TABLE_TIMESERIES)
        ComputeTableScope = ScopeTimeSeries
    Case LCase$(TABLE_SPATIAL)
        ComputeTableScope = ScopeSpatial
    Case LCase$(TABLE_SPATIOTEMPORAL)
        ComputeTableScope = ScopeSpatioTemporal
    Case LCase$(TABLE_GRAPH_TIMESERIES)
        ComputeTableScope = ScopeTimeSeriesGraph

    Case Else
        ThrowError ErrorUnexpectedState, _
                   "Could not find the table scope. Actual table is: " & specsTable.Name

    End Select
End Function

'@label:compute-valid-table
'@sub-title Compute whether this table has a valid configuration
'@details
'Validates whether this table specification row contains all required fields
'and references valid variables for its analysis type. Called once on first
'access of ValidTable to populate the Boolean cache. Each branch resolves
'only the variables its own rule needs.
'@return Boolean. True if the table specification meets all requirements.
'@throws ErrorUnexpectedState When the table scope is unmapped.
'@depends LLVariables
Private Function ComputeValidTable() As Boolean
    Dim cachedDict As LLdictionary
    Dim vars As LLVariables
    Dim rowVar As String
    Dim colVar As String
    Dim colControl As String
    Dim rowControl As String
    Dim rowType As String
    Dim isValid As Boolean

    ' GlobalSummary validates field presence only, no dictionary variable
    ' lookups needed
    If TableScope = ScopeGlobalSummary Then
        ComputeValidTable = (Value(COLUMN_LABEL) <> vbNullString) And _
                            (Value(COLUMN_FUNCTION) <> vbNullString)
        Exit Function
    End If

    ' The graph block names a graph and a series through two dropdowns, and
    ' every other column of the row is a formula derived from those two.
    If TableScope = ScopeTimeSeriesGraph Then
        ComputeValidTable = (Value(COLUMN_GRAPH_TITLE) <> vbNullString) And _
                            (Value(COLUMN_SERIES_TITLE) <> vbNullString)
        Exit Function
    End If

    Set cachedDict = Dictionary()
    Set vars = VariablesObject()

    If vars Is Nothing Then
        ComputeValidTable = False
        Exit Function
    End If

    rowVar = Value(COLUMN_ROW)
    colVar = Value(COLUMN_COLUMN)

    Select Case TableScope

    Case ScopeUnivariate
        rowControl = vars.Value(colName:=VARIABLE_CONTROL, varName:=rowVar)
        isValid = vars.Contains(rowVar) And _
                  (rowControl = "choice_manual" Or rowControl = "choice_formula")

    Case ScopeBivariate
        rowControl = vars.Value(colName:=VARIABLE_CONTROL, varName:=rowVar)
        colControl = vars.Value(colName:=VARIABLE_CONTROL, varName:=colVar)
        isValid = vars.Contains(rowVar) And vars.Contains(colVar) And _
                  (colControl = "choice_manual" Or colControl = "choice_formula") And _
                  (rowControl = "choice_manual" Or rowControl = "choice_formula")

    Case ScopeTimeSeries
        rowType = vars.Value(colName:=VARIABLE_TYPE, varName:=rowVar)
        colControl = vars.Value(colName:=VARIABLE_CONTROL, varName:=colVar)
        isValid = (colVar = vbNullString Or _
                   colControl = "choice_manual" Or _
                   colControl = "choice_formula") And _
                  (rowType = "date")

    Case ScopeSpatial
        colControl = vars.Value(colName:=VARIABLE_CONTROL, varName:=colVar)
        isValid = (cachedDict.VariableExists(PREFIX_ADMIN & rowVar) Or _
                   cachedDict.VariableExists(PREFIX_HEALTH_FACILITY & rowVar)) And _
                  (colVar = vbNullString Or _
                   colControl = "choice_manual" Or _
                   colControl = "choice_formula")

    Case ScopeSpatioTemporal
        ' LLdictionary.AppendGeoLines rewrites a geo control to geo1..geo4
        ' during Prepare and leaves hf alone, so "hf" is the one control
        ' value still readable here.
        rowType = vars.Value(colName:=VARIABLE_TYPE, varName:=rowVar)
        colControl = vars.Value(colName:=VARIABLE_CONTROL, varName:=colVar)
        isValid = ((colControl = "hf") Or _
                   cachedDict.VariableExists(PREFIX_ADMIN & colVar) Or _
                   cachedDict.VariableExists(PREFIX_HEALTH_FACILITY & colVar)) And _
                  (rowType = "date")

    Case Else
        ThrowError ErrorUnexpectedState, _
                   "ComputeValidTable has no rule for scope " & TableScope

    End Select

    ComputeValidTable = isValid
End Function

'@label:compute-is-new-section
'@sub-title Compute whether this table begins a new section group
'@details
'Determines whether this table specification row begins a new section by
'comparing the current and previous row section values. Called once on first
'access of IsNewSection to populate the Boolean cache. The previous row cell
'value is the only remaining worksheet read (the current row value comes
'from the value cache), and it goes through the same coercion the cache
'used, so an error value in that cell reads as an empty section.
'@return Boolean. True if this table starts a new section.
Private Function ComputeIsNewSection() As Boolean
    Dim sectIndex As Long
    Dim actualSection As String
    Dim previousSection As String
    Dim isSameSection As Boolean
    Dim prevCell As Range

    isSameSection = False

    sectIndex = ColumnIndex(COLUMN_SECTION)

    If sectIndex > 0 Then
        actualSection = Value(COLUMN_SECTION)
        Set prevCell = this.tRng.Cells(0, sectIndex)
        previousSection = CellText(prevCell.Value)
        isSameSection = (StrComp(actualSection, previousSection, vbTextCompare) = 0) And _
                        (prevCell.Row <> this.hRng.Row)
    End If

    ComputeIsNewSection = (Not isSameSection) And (TableScope <> ScopeGlobalSummary)
End Function

'@section Error Handling
'===============================================================================
'@description Logging and error-raising helpers.

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME,
'providing a consistent stack trace across all methods in this class.
'@param errNumber Long. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@label:check-requirements
'@sub-title Validate preconditions for the Create factory
'@details
'Validates all preconditions required by the Create factory before populating
'a new TableSpecs instance. Checks that headerRng, tableSpecsRng, and dict
'are not Nothing, that both ranges span a single row, that they start at the
'same column, that they have the same column count, and that the header sits
'above the specification row.
'
'The column-start check matters because ComputeIsNewSection resolves the
'section index against the header range and then indexes the specification
'range with it: two ranges starting at different columns make every section
'read land on the wrong column.
'@param headerRng Range. The header range to validate.
'@param tableSpecsRng Range. The table specification row range to validate.
'@param dict LLdictionary. The dictionary object to validate.
'@throws InvalidArgument When any precondition fails.
Private Sub CheckRequirements(ByVal headerRng As Range, _
                              ByVal tableSpecsRng As Range, _
                              ByVal dict As LLdictionary)

    If headerRng Is Nothing Then
        ThrowError InvalidArgument, "Header range not defined"
    End If

    If tableSpecsRng Is Nothing Then
        ThrowError InvalidArgument, "Table specification range not defined"
    End If

    If dict Is Nothing Then
        ThrowError InvalidArgument, "Dictionary not defined"
    End If

    If headerRng.Rows.Count <> 1 Then
        ThrowError InvalidArgument, "Header range must span one row"
    End If

    If tableSpecsRng.Rows.Count <> 1 Then
        ThrowError InvalidArgument, "Specification range must span one row"
    End If

    If headerRng.Columns.Count <> tableSpecsRng.Columns.Count Then
        ThrowError InvalidArgument, _
                   "Header range and specification range must have same column count"
    End If

    If headerRng.Column <> tableSpecsRng.Column Then
        ThrowError InvalidArgument, _
                   "Header range and specification range must start at the same column"
    End If

    If headerRng.Row >= tableSpecsRng.Row Then
        ThrowError InvalidArgument, _
                   "Header range must be above the specification range"
    End If
End Sub
