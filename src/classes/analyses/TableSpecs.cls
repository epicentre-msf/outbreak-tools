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
Implements ITableSpecs

'@class TableSpecs
'@description
'Represents a single row from an analysis setup ListObject and exposes
'properties that describe the analysis table it defines. Each row in the
'setup table specifies one output table (frequency, cross-tab, time series,
'spatial, or spatio-temporal). The class wraps that row together with the
'header range and a dictionary, providing typed access to the table scope,
'identifiers, navigation to neighboring rows, validation flags (total,
'missing, percentage, graph), and dictionary-based category lookups.
'Performance-sensitive data is cached eagerly at construction time: column
'names, cell values, table scope, and dictionary reference. Derived booleans
'(ValidTable, IsNewSection) are cached lazily on first access. Category
'lookups (RowCategories, ColumnCategories) accept an Object parameter: when
'a LinelistSpecs instance is passed, categories are resolved from its
'dictionary and choices; otherwise an empty BetterArray is returned. This
'design decouples TableSpecs from LinelistSpecs at instantiation, making it
'usable in the setup error-checking context without geo or translation
'dependencies. Consumers interact through the ITableSpecs interface.
'@depends ILLdictionary, ILLVariables, LLVariables, BetterArray
'@version 2.0 (2026-02-11)
'@author Yves Amevoin

Private Type TTableSpecs
    hRng As Range                                'Header range of the setup listobject
    tRng As Range                                'Row range for this table spec
    ' --- Column and cell value caches (Tier 1 + Tier 3) ---
    colNames As BetterArray                      'Column names from header (original case)
    cellValues As BetterArray                    'Parallel cell values from spec row
    ' --- Scalar caches (Tier 2 + Tier 4 + Tier 5) ---
    cachedScope As Long                          'Cached AnalysisTableScope enum value
    scopeCached As Boolean                       'Whether cachedScope is populated
    dict As ILLdictionary                        'Dictionary reference
    cachedValid As Boolean                       'Cached ValidTable result
    validCached As Boolean                       'Whether cachedValid is populated
    cachedNewSection As Boolean                  'Cached IsNewSection result
    newSectionCached As Boolean                  'Whether cachedNewSection is populated
End Type

Private this As TTableSpecs
Private Const CLASS_NAME As String = "TableSpecs"

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
'injection, InitCache scans the header and data row once to build column
'name and cell value caches, and eagerly resolves the table scope.
'@param headerRng Range. The header row Range of the analysis setup ListObject.
'Column names in this range are used by Value() and ColumnIndex() to locate
'fields such as "row", "column", "total", "section", etc.
'@param tableSpecsRng Range. A single data-row Range from the same ListObject.
'Must be below headerRng and have the same number of columns.
'@param dict ILLdictionary. The linelist dictionary that provides variable
'definitions, existence checks, and control type lookups needed for
'validation and spatial type detection.
'@return ITableSpecs. A fully initialised TableSpecs instance ready for use.
'@throws InvalidArgument When headerRng is Nothing.
'@throws InvalidArgument When tableSpecsRng is Nothing.
'@throws InvalidArgument When dict is Nothing.
'@throws InvalidArgument When headerRng and tableSpecsRng have different column counts.
'@throws InvalidArgument When headerRng.Row is not above tableSpecsRng.Row.
'@throws InvalidArgument When headerRng.Row is less than 1.
'@export
Public Function Create(ByVal headerRng As Range, _
                       ByVal tableSpecsRng As Range, _
                       ByVal dict As ILLdictionary) As ITableSpecs

    Dim instance As TableSpecs
    CheckRequirements headerRng, tableSpecsRng, dict

    Set instance = New TableSpecs
    Set instance.HeaderRange = headerRng
    Set instance.TableRange = tableSpecsRng
    Set instance.DictionaryObject = dict
    instance.InitCache
    Set Create = instance
End Function

'@section Internal Properties
'===============================================================================
'@description Properties used by the factory pattern during initialisation.

'@label:header-range
'@prop-title Header range of the setup ListObject
'@details
'Returns the header row Range that contains column names such as "row",
'"column", "total", "section", etc. Used by Value() and ColumnIndex()
'for all field lookups.
'@return Range. The header row range.
Public Property Get HeaderRange() As Range
    Set HeaderRange = this.hRng
End Property

'@label:header-range-set
'@prop-title Assign the header range
'@param hRng Range. The header row range to store.
Public Property Set HeaderRange(ByVal hRng As Range)
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
    Set this.tRng = tRng
End Property

'@label:dictionary-object
'@prop-title Dictionary object used for validation and variable lookups
'@details
'Returns the ILLdictionary instance that provides variable existence
'checks, control type lookups, and spatial prefix detection. Set during
'factory construction.
'@return ILLdictionary. The dictionary object.
Public Property Get DictionaryObject() As ILLdictionary
    Set DictionaryObject = this.dict
End Property

'@label:dictionary-object-set
'@prop-title Assign the dictionary object
'@param dict ILLdictionary. The dictionary to store.
Public Property Set DictionaryObject(ByVal dict As ILLdictionary)
    Set this.dict = dict
End Property

'@section Cache Initialization
'===============================================================================
'@description One-time cache population called from the factory.

'@label:init-cache
'@sub-title Populate all eager caches from header and data ranges
'@details
'Scans the header range once to build parallel BetterArray caches of column
'names and cell values (Tier 1 + Tier 3). Also eagerly resolves the table
'scope from the cell above the header (Tier 2). Must be called exactly once
'after HeaderRange, TableRange, and DictionaryObject have been assigned.
'Lazy caches (ValidTable, IsNewSection) are NOT populated here; they are
'computed on first access.
'@export
Public Sub InitCache()
    Dim colCount As Long
    Dim idx As Long

    Set this.colNames = New BetterArray
    Set this.cellValues = New BetterArray
    this.colNames.LowerBound = 1
    this.cellValues.LowerBound = 1

    colCount = this.hRng.Columns.Count

    ' Tier 1 + Tier 3: One-time scan of header and data row
    For idx = 1 To colCount
        this.colNames.Push CStr(this.hRng.Cells(1, idx).Value)
        this.cellValues.Push CStr(this.tRng.Cells(1, idx).Value)
    Next idx

    ' Tier 2: Eagerly cache table scope
    this.cachedScope = ComputeTableScope()
    this.scopeCached = True
End Sub

'@section Core Properties
'===============================================================================
'@description Type identification and identity properties for each table spec.

'@label:table-scope
'@prop-title Analysis table scope for this specification row
'@details
'Returns the cached AnalysisTableScope enum value. The scope is computed
'once eagerly during InitCache by reading the type label cell above the
'header range. On subsequent calls the cached Long value is returned
'directly without any worksheet access.
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
'Bivariate, "TS" for TimeSeries, "SA" for Spatial, and "SPT" for
'SpatioTemporal. The row offset is computed as TableRange.Row minus
'HeaderRange.Row, producing a stable integer that uniquely locates this
'row within its analysis block. The resulting format is "prefix_tabN",
'for example "UA_tab3".
'@return String. The unique table identifier such as "GS_tab1" or "BA_tab5".
'@export
Public Property Get TableId() As String
    Dim prefix As String
    Dim lineNum As Long

    lineNum = TableRange.Row - HeaderRange.Row

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
    End Select

    TableId = prefix & "_tab" & lineNum
End Property

'@label:table-section-id
'@prop-title Identifier of the first table in this section
'@details
'Returns the TableId of the first table in the current section group by
'walking backward through the Previous chain. If this table is itself the
'start of a new section (IsNewSection = True) or is a GlobalSummary table,
'it returns its own TableId immediately. Otherwise it delegates to
'Previous.TableSectionId, which continues walking backward until a section
'boundary is reached. All tables within the same section share a common
'section identifier.
'@return String. The TableId of the first table in this section.
'@remarks This is a recursive property. For very long sections with many
'rows, it will traverse one row at a time via Previous. GlobalSummary
'tables short-circuit immediately because they are standalone.
'@export
Public Property Get TableSectionId() As String
    If IsNewSection() Or (TableScope = ScopeGlobalSummary) Then
        TableSectionId = TableId()
    Else
        TableSectionId = Previous.TableSectionId
    End If
End Property

'@label:dictionary
'@prop-title Linelist dictionary bound to this specification
'@details
'Returns the ILLdictionary reference stored at construction time. The
'dictionary provides variable existence checks, control type lookups,
'and spatial prefix detection used throughout validation and category
'resolution.
'@return ILLdictionary. The linelist dictionary instance.
'@export
Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dict
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
'geography, or vbNullString if neither prefix is found.
'@return String. The spatial sub-type ("hf", "geo", or empty string).
'@remarks This property is meaningful only for Spatial and SpatioTemporal
'table scopes. Calling it on other table scopes will still execute but will
'likely return vbNullString since those variables typically lack geo/hf
'prefixes.
'@export
Public Property Get SpatialTableScopes() As String
    Dim rowVar As String
    Dim cachedDict As ILLdictionary
    Dim spatType As String

    Set cachedDict = Dictionary()

    ' For spatio-temporal, the spatial variable is in the "column" field
    If TableScope = ScopeSpatioTemporal Then
        rowVar = Value("column")
    Else
        rowVar = Value("row")
    End If

    spatType = Switch(cachedDict.VariableExists("hf_" & rowVar), "hf", _
                      cachedDict.VariableExists("adm1_" & rowVar), "geo", _
                      True, vbNullString)
    SpatialTableScopes = spatType
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
'a geographic prefix on the column variable and a date-type row variable.
'Invalid table specs are skipped during navigation (Previous, NextSpecs).
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
'to the row immediately above the current row. Section values are compared
'with exact string equality (case-sensitive).
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
'varies by table scope: GlobalSummary and SpatioTemporal always return False;
'Univariate and Bivariate always return True; TimeSeries returns True when
'the user explicitly requested total="yes" or when percentage is set to "row"
'or "column" (because percentage formulas need the total column as a
'denominator); Spatial returns True only when a column variable is specified.
'This is distinct from TotalRequested: HasTotal can be True even when the
'user did not ask for totals, solely because percentage needs them.
'@return Boolean. True if a total row/column should be created.
'@export
Public Property Get HasTotal() As Boolean
    Select Case TableScope

    Case ScopeGlobalSummary, ScopeSpatioTemporal
        HasTotal = False

    Case ScopeUnivariate
        HasTotal = True

    Case ScopeBivariate
        HasTotal = True

    Case ScopeTimeSeries
        HasTotal = ((Value("total") = "yes") Or _
                    (Value("percentage") = "row") Or _
                    (Value("percentage") = "column")) And _
                   (Value("column") <> vbNullString)

    Case ScopeSpatial
        HasTotal = (Value("column") <> vbNullString)

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
'shown. This property is primarily relevant for TimeSeries and Spatial types.
'@export
Public Property Get TotalRequested() As Boolean
    TotalRequested = (Value("total") = "yes")
End Property

'@label:has-missing
'@prop-title Whether missing-data rows or columns are included
'@details
'Determines whether a "missing" row or column should be included in the
'output table. The logic varies by table scope: GlobalSummary and
'SpatioTemporal always return False; Univariate returns True when
'missing="yes"; Bivariate returns True when missing="row", "column", or
'"all"; TimeSeries and Spatial return True when missing="yes" and a column
'variable is present (without a column variable, there is no categorical
'axis to show missing values on).
'@return Boolean. True if missing data should be shown in the output table.
'@export
Public Property Get HasMissing() As Boolean
    Select Case TableScope

    Case ScopeGlobalSummary, ScopeSpatioTemporal
        HasMissing = False

    Case ScopeUnivariate
        HasMissing = (Value("missing") = "yes")

    Case ScopeBivariate
        HasMissing = (Value("missing") = "row") Or _
                     (Value("missing") = "column") Or _
                     (Value("missing") = "all")

    Case ScopeTimeSeries
        HasMissing = (Value("missing") = "yes") And _
                     (Value("column") <> vbNullString)

    Case ScopeSpatial
        HasMissing = (Value("missing") = "yes") And _
                     (Value("column") <> vbNullString)

    End Select
End Property

'@label:has-percentage
'@prop-title Whether percentage display is enabled
'@details
'Determines whether percentage values should be computed and displayed for
'this table. The logic varies by table scope: GlobalSummary and SpatioTemporal
'always return False; Univariate returns True when percentage="yes"; Bivariate
'returns True when percentage="row", "column", or "total"; TimeSeries returns
'True when percentage="row" or "column" and HasTotal is True; Spatial returns
'True when percentage="yes" and HasTotal is True. The interplay between
'HasPercentage and HasTotal for TimeSeries is intentional: requesting
'percentage implicitly forces HasTotal=True since percentage needs the total.
'@return Boolean. True if percentage values should be included.
'@export
Public Property Get HasPercentage() As Boolean
    Select Case TableScope

    Case ScopeGlobalSummary, ScopeSpatioTemporal
        HasPercentage = False

    Case ScopeUnivariate
        HasPercentage = (Value("percentage") = "yes")

    Case ScopeBivariate
        HasPercentage = (Value("percentage") = "row") Or _
                        (Value("percentage") = "column") Or _
                        (Value("percentage") = "total")

    Case ScopeTimeSeries
        HasPercentage = ((Value("percentage") = "row") Or _
                         (Value("percentage") = "column")) And HasTotal()

    Case ScopeSpatial
        HasPercentage = (Value("percentage") = "yes") And HasTotal()

    End Select
End Property

'@label:has-graph
'@prop-title Whether a graph should be created for this table
'@details
'Determines whether a chart/graph should be generated alongside the output
'table. The logic varies by table scope: GlobalSummary always returns False;
'Univariate, TimeSeries, Spatial, and SpatioTemporal return True when
'graph="yes"; Bivariate returns True when graph="percentage", "values", or
'"both", letting the user choose whether the chart shows raw counts,
'percentages, or both. The actual graph rendering is handled downstream by
'the chart builder; this property only signals whether one should be created.
'@return Boolean. True if a graph should be created for this table.
'@export
Public Property Get HasGraph() As Boolean
    Select Case TableScope

    Case ScopeGlobalSummary
        HasGraph = False

    Case ScopeUnivariate
        HasGraph = (Value("graph") = "yes")

    Case ScopeBivariate
        HasGraph = (Value("graph") = "percentage" Or _
                    Value("graph") = "values" Or _
                    Value("graph") = "both")

    Case ScopeTimeSeries
        HasGraph = (Value("graph") = "yes")

    Case ScopeSpatial
        HasGraph = (Value("graph") = "yes")

    Case ScopeSpatioTemporal
        HasGraph = (Value("graph") = "yes")

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
'range, it walks backward one row at a time until it finds a valid table or
'encounters a section boundary (IsNewSection on the candidate). The walk
'stops at the header row to prevent reading outside the data area. This
'property is used by TableSectionId to recursively find the first table in
'a section.
'@return ITableSpecs. The nearest valid preceding table in the same section.
'@throws InvalidArgument When IsNewSection is True, because there is no
'previous table in the same section for a section-starting table.
'@remarks Callers must check IsNewSection before accessing Previous. Each
'candidate row examined creates a new TableSpecs instance, so performance
'degrades with many consecutive invalid rows.
'@export
Public Property Get Previous() As ITableSpecs
    Dim hRng As Range
    Dim tRng As Range
    Dim prevRng As Range
    Dim prevSpec As ITableSpecs

    If IsNewSection() Then
        ThrowError InvalidArgument, _
                   "Table is on a new section, cannot return previous table"
    End If

    Set hRng = HeaderRange()
    Set tRng = TableRange()
    Set prevRng = tRng.Offset(-1)

    ' Create spec for the row just above
    Set prevSpec = Create(hRng, prevRng, Dictionary())
    Set prevRng = prevRng.Offset(-1)

    ' Walk backward until we find a valid table or reach section boundary
    Do While (Not prevSpec.ValidTable()) And (Not (prevRng.Row <= hRng.Row))
        Set prevSpec = Create(hRng, prevRng, Dictionary())
        If prevSpec.IsNewSection() Then Exit Do
        Set prevRng = prevRng.Offset(-1)
    Loop

    Set Previous = prevSpec
End Property

'@label:next-specs
'@prop-title Next valid specification following an anchor range
'@details
'Returns the next valid table specification after this row by walking forward
'through the setup ListObject. Starting from the row immediately below the
'current table range, it creates a TableSpecs for that row. If the row is
'invalid, it continues advancing one row at a time until it finds a valid
'table or passes the anchor range. Unlike Previous, NextSpecs does not stop
'at section boundaries. If no valid table is found within the data area, the
'property returns Nothing.
'@param anchorRng Range. A Range representing the last row of the setup
'ListObject. The forward walk will not advance past this row.
'@return ITableSpecs. The next valid table after this row, or Nothing if no
'valid table exists between this row and the anchor range.
'@remarks This property is the primary iteration mechanism for processing all
'tables in an analysis block. Callers typically start with the first table and
'call NextSpecs repeatedly until Nothing is returned.
'@export
Public Property Get NextSpecs(ByVal anchorRng As Range) As ITableSpecs
    Dim tRng As Range
    Dim hRng As Range
    Dim nextSpec As ITableSpecs

    Set tRng = TableRange()
    Set hRng = HeaderRange()
    Set tRng = tRng.Offset(1)

    If tRng.Row > anchorRng.Row Then Exit Property

    Set nextSpec = Create(hRng, tRng, Dictionary())
    Set tRng = tRng.Offset(1)

    Do While (Not nextSpec.ValidTable()) And (tRng.Row <= anchorRng.Row)
        Set nextSpec = Create(hRng, tRng, Dictionary())
        Set tRng = tRng.Offset(1)
    Loop

    If Not nextSpec.ValidTable() Then Exit Property

    Set NextSpecs = nextSpec
End Property

'@section Data Access
'===============================================================================
'@description Value retrieval and category lookups from the specification row.

'@label:value
'@sub-title Retrieve a value from the specification row by column name
'@details
'Retrieves the cached cell value from this table specification row for a
'given column name. The column is located by scanning the cached column
'names array for a case-insensitive substring match (replicating the
'original Range.Find with xlPart and MatchCase:=False). Once found, the
'corresponding value is returned from the parallel cellValues cache without
'any worksheet access. If the column name does not exist in the header, the
'method returns vbNullString.
'@param colName String. The column name to search for in the header
'(partial, case-insensitive).
'@return String. The cell value, or vbNullString if the column is not found.
'@export
Public Function Value(ByVal colName As String) As String
    Dim idx As Long

    Value = vbNullString

    For idx = 1 To this.colNames.Length
        If InStr(1, this.colNames.Item(idx), colName, vbTextCompare) > 0 Then
            Value = this.cellValues.Item(idx)
            Exit Function
        End If
    Next idx
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
    Set RowCategories = CategoriesData("row", lData)
End Property

'@label:column-categories
'@prop-title Column variable categories from the dictionary
'@details
'Returns a BetterArray of category values for the column variable of this
'table. When lData is a LinelistSpecs instance, the categories come from it
'based on the variable name in the "column" field. When lData is any other
'type (or Nothing), returns an empty BetterArray. For SpatioTemporal tables,
'the column categories are empty placeholder strings whose count is determined
'by the "n geo" field in the setup row (defaulting to 5). These placeholders
'are filled in later at runtime with actual geographic location names.
'Internally delegates to CategoriesData("column", lData).
'@param lData Object. A LinelistSpecs instance providing category lookups,
'or any other Object (categories will be empty).
'@return BetterArray. The category strings for the column variable. For
'SpatioTemporal, contains N empty strings. Returns an empty BetterArray if the
'column variable is not found or not specified.
'@export
Public Property Get ColumnCategories(ByRef lData As Object) As BetterArray
    Set ColumnCategories = CategoriesData("column", lData)
End Property

'@section Helpers
'===============================================================================
'@description Private helper methods that support data access and validation.

'@label:column-exists
'@sub-title Check whether a column name exists in the header cache
'@details
'Delegates to ColumnIndex and checks for a positive return value. Replaces
'the original Range.Find call with a scan of the cached column names array.
'@param colName String. The column name to search for (partial, case-insensitive).
'@return Boolean. True if at least one cached column name contains the search term.
Private Function ColumnExists(ByVal colName As String) As Boolean
    ColumnExists = (ColumnIndex(colName) > 0)
End Function

'@label:column-index
'@sub-title Return the 1-based column index of a named header column
'@details
'Scans the cached column names array for a case-insensitive substring match
'(replicating the original Range.Find with xlPart and MatchCase:=False).
'Returns the 1-based position within the header range, or -1 if not found.
'@param colName String. The column name to search for (partial, case-insensitive).
'@return Long. The 1-based position within the header range, or -1 if not found.
Private Function ColumnIndex(ByVal colName As String) As Long
    Dim idx As Long

    For idx = 1 To this.colNames.Length
        If InStr(1, this.colNames.Item(idx), colName, vbTextCompare) > 0 Then
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
'rowOrCol is "column", the categories are not looked up from the dictionary
'but instead populated with N empty strings (where N comes from the "n geo"
'setup field, defaulting to 5). These empty placeholders represent geographic
'locations filled in at runtime. The resulting BetterArray is cloned before
'returning to prevent external mutations from affecting internal state.
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
    Dim vars As ILLVariables
    Dim nbGeo As Long
    Dim counter As Long

    ' LLVariables.Create may fail if the dictionary lacks a Data sheet or
    ' required columns (e.g. in the setup error-checking context).
    On Error Resume Next
    Set vars = LLVariables.Create(Dictionary())
    On Error GoTo 0

    Set varsCat = New BetterArray
    varsCat.LowerBound = 1

    nameVar = Value(rowOrCol)

    ' For SpatioTemporal, column categories are empty geo placeholders
    If (TableScope = ScopeSpatioTemporal) And (rowOrCol = "column") Then
        nbGeo = 5
        On Error Resume Next
        nbGeo = CLng(Value("n geo"))
        On Error GoTo 0
        For counter = 1 To nbGeo
            varsCat.Push vbNullString
        Next
    ElseIf Not vars Is Nothing Then
        If vars.Contains(nameVar) Then
            ' Only resolve categories when lData is a LinelistSpecs instance
            If Not lData Is Nothing Then
                If TypeName(lData) = "LinelistSpecs" Then
                    Set varsCat = lData.Categories(nameVar)
                End If
            End If
        End If
    End If

    Set CategoriesData = varsCat.Clone()
End Function

'@label:compute-table-scope
'@sub-title Compute the AnalysisTableScope from the setup worksheet
'@details
'Reads the type label cell two rows above the header range, trims and
'lowercases it, then maps it to the corresponding AnalysisTableScope enum
'constant. Called once during InitCache to populate the scope cache.
'@return AnalysisTableScope. The table scope category.
'@throws ErrorUnexpectedState When the type label does not match any known analysis type.
Private Function ComputeTableScope() As AnalysisTableScope
    Dim hRng As Range
    Dim anaType As String

    Set hRng = HeaderRange()
    anaType = LCase(Application.WorksheetFunction.Trim(hRng.Cells(-1, 1).Value))

    Select Case anaType
    Case "global summary"
        ComputeTableScope = ScopeGlobalSummary
    Case "univariate analysis"
        ComputeTableScope = ScopeUnivariate
    Case "bivariate analysis"
        ComputeTableScope = ScopeBivariate
    Case "time series analysis"
        ComputeTableScope = ScopeTimeSeries
    Case "spatial analysis"
        ComputeTableScope = ScopeSpatial
    Case "spatio-temporal analysis"
        ComputeTableScope = ScopeSpatioTemporal
    Case "time series graphs"
        ComputeTableScope = ScopeTimeSeriesGraph
    Case Else
        ThrowError ErrorUnexpectedState, _
                   "Could not find the table scope. Actual type is: " & anaType
    End Select
End Function

'@label:compute-valid-table
'@sub-title Compute whether this table has a valid configuration
'@details
'Validates whether this table specification row contains all required fields
'and references valid variables for its analysis type. Called once on first
'access of ValidTable to populate the Boolean cache.
'@return Boolean. True if the table specification meets all requirements.
'@depends LLVariables
Private Function ComputeValidTable() As Boolean
    Dim cachedDict As ILLdictionary
    Dim vars As ILLVariables
    Dim rowVar As String
    Dim colVar As String
    Dim colControl As String
    Dim rowControl As String
    Dim rowType As String
    Dim isValid As Boolean

    ' GlobalSummary validates field presence only, no dictionary variable
    ' lookups needed
    If TableScope = ScopeGlobalSummary Then
        ComputeValidTable = (Value("label") <> vbNullString) And _
                            (Value("function") <> vbNullString)
        Exit Function
    End If

    Set cachedDict = Dictionary()

    ' LLVariables.Create requires a fully-configured dictionary (Data sheet,
    ' "variable name" column). In the setup error-checking context the
    ' dictionary may not expose these, so fall back to False
    On Error Resume Next
    Set vars = LLVariables.Create(cachedDict)
    On Error GoTo 0

    If vars Is Nothing Then
        ComputeValidTable = False
        Exit Function
    End If

    rowVar = Value("row")
    rowControl = vars.Value(colName:="control", varName:=rowVar)
    colVar = Value("column")
    colControl = vars.Value(colName:="control", varName:=colVar)

    Select Case TableScope

    Case ScopeUnivariate
        isValid = vars.Contains(rowVar) And _
                  (rowControl = "choice_manual" Or rowControl = "choice_formula")

    Case ScopeBivariate
        isValid = vars.Contains(rowVar) And vars.Contains(colVar) And _
                  (colControl = "choice_manual" Or colControl = "choice_formula") And _
                  (rowControl = "choice_manual" Or rowControl = "choice_formula")

    Case ScopeTimeSeries
        rowType = vars.Value(colName:="variable type", varName:=rowVar)
        isValid = (colVar = vbNullString Or _
                   colControl = "choice_manual" Or _
                   colControl = "choice_formula") And _
                  (rowType = "date")

    Case ScopeSpatial
        isValid = (cachedDict.VariableExists("adm1_" & rowVar) Or _
                   cachedDict.VariableExists("hf_" & rowVar)) And _
                  (colVar = vbNullString Or _
                   colControl = "choice_manual" Or _
                   colControl = "choice_formula")

    Case ScopeSpatioTemporal
        rowType = vars.Value(colName:="variable type", varName:=rowVar)
        isValid = ( (colControl = "geo") Or (colControl = "hf") Or _
                   cachedDict.VariableExists("adm1_" & colVar) Or _
                   cachedDict.VariableExists("hf_" & colVar)) And _
                  (rowType = "date")

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
'from the cellValues cache).
'@return Boolean. True if this table starts a new section.
Private Function ComputeIsNewSection() As Boolean
    Dim sectIndex As Long
    Dim actualSection As String
    Dim previousSection As String
    Dim isSameSection As Boolean
    Dim prevCell As Range

    isSameSection = False

    sectIndex = ColumnIndex("section")

    If sectIndex > 0 Then
        actualSection = Value("section")
        Set prevCell = this.tRng.Cells(0, sectIndex)
        previousSection = prevCell.Value
        isSameSection = (actualSection = previousSection) And _
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
'are not Nothing, that the two ranges have the same column count, that the
'header is above the specification row, and that the header row is at least 1.
'If any check fails, a descriptive error is raised via ThrowError.
'@param headerRng Range. The header range to validate.
'@param tableSpecsRng Range. The table specification row range to validate.
'@param dict ILLdictionary. The dictionary object to validate.
'@throws InvalidArgument When any precondition fails.
Private Sub CheckRequirements(ByVal headerRng As Range, _
                              ByVal tableSpecsRng As Range, _
                              ByVal dict As ILLdictionary)

    If headerRng Is Nothing Then
        ThrowError InvalidArgument, "Header range not defined"
    End If

    If tableSpecsRng Is Nothing Then
        ThrowError InvalidArgument, "Table specification range not defined"
    End If

    If dict Is Nothing Then
        ThrowError InvalidArgument, "Dictionary not defined"
    End If

    If headerRng.Columns.Count <> tableSpecsRng.Columns.Count Then
        ThrowError InvalidArgument, _
                   "Header range and specification range must have same column count"
    End If

    If headerRng.Row >= tableSpecsRng.Row Then
        ThrowError InvalidArgument, _
                   "Header range must be above the specification range"
    End If

    If headerRng.Row < 1 Then
        ThrowError InvalidArgument, "Header range must be at row >= 1"
    End If
End Sub

'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the ITableSpecs contract.
'See the corresponding Public members above for full documentation.

Private Property Get ITableSpecs_TableScope() As AnalysisTableScope
    ITableSpecs_TableScope = TableScope()
End Property

Private Property Get ITableSpecs_Dictionary() As ILLdictionary
    Set ITableSpecs_Dictionary = Dictionary()
End Property

Private Property Get ITableSpecs_RowCategories(ByRef lData As Object) As BetterArray
    Set ITableSpecs_RowCategories = RowCategories(lData)
End Property

Private Property Get ITableSpecs_ColumnCategories(ByRef lData As Object) As BetterArray
    Set ITableSpecs_ColumnCategories = ColumnCategories(lData)
End Property

Private Property Get ITableSpecs_IsNewSection() As Boolean
    ITableSpecs_IsNewSection = IsNewSection()
End Property

Private Property Get ITableSpecs_TableId() As String
    ITableSpecs_TableId = TableId()
End Property

Private Property Get ITableSpecs_ValidTable() As Boolean
    ITableSpecs_ValidTable = ValidTable()
End Property

Private Property Get ITableSpecs_HasMissing() As Boolean
    ITableSpecs_HasMissing = HasMissing()
End Property

Private Property Get ITableSpecs_HasTotal() As Boolean
    ITableSpecs_HasTotal = HasTotal()
End Property

Private Property Get ITableSpecs_TotalRequested() As Boolean
    ITableSpecs_TotalRequested = TotalRequested()
End Property

Private Property Get ITableSpecs_HasGraph() As Boolean
    ITableSpecs_HasGraph = HasGraph()
End Property

Private Property Get ITableSpecs_HasPercentage() As Boolean
    ITableSpecs_HasPercentage = HasPercentage()
End Property

Private Property Get ITableSpecs_SpatialTableScopes() As String
    ITableSpecs_SpatialTableScopes = SpatialTableScopes()
End Property

Private Property Get ITableSpecs_Previous() As ITableSpecs
    Set ITableSpecs_Previous = Previous()
End Property

Private Property Get ITableSpecs_NextSpecs(ByVal anchorRng As Range) As ITableSpecs
    Set ITableSpecs_NextSpecs = NextSpecs(anchorRng)
End Property

Private Property Get ITableSpecs_TableSectionId() As String
    ITableSpecs_TableSectionId = TableSectionId()
End Property

Private Function ITableSpecs_Value(ByVal colName As String) As String
    ITableSpecs_Value = Value(colName)
End Function
