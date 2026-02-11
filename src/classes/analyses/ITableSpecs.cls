VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for analysis table specifications"

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface for analysis table specifications")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'@class ITableSpecs
'@description
'Defines the public contract for the TableSpecs class. Exposes read-only
'properties that describe a single analysis table row from the setup
'ListObject: type, identifiers, validation flags, navigation to sibling
'specs, and dictionary-based category lookups.


'@section CoreProperties
'===============================================================================
'@description Type, identity, and navigation properties for each table spec.

' Analysis table scopes matching setup listobject types
Public Enum AnalysisTableScope
    ScopeGlobalSummary = 1
    ScopeUnivariate = 2
    ScopeBivariate = 3
    ScopeTimeSeries = 4
    ScopeSpatial = 5
    ScopeSpatioTemporal = 6
End Enum


'@jump:table-scope
'@prop-title Analysis table scope for this specification row
'@details
'Returns the AnalysisTableScope enum value that categorises this table
'as global summary, univariate, bivariate, time series, spatial, or
'spatio-temporal. The type determines which cross-table layout and
'formula logic apply at build time.
'@return AnalysisTableScope. The table scope category.
Public Property Get TableScope() As AnalysisTableScope: End Property

'@jump:dictionary
'@prop-title Linelist dictionary bound to this specification
'@details
'Returns the dictionary instance that holds the variable definitions and
'category lists used by the analysis table. The dictionary is shared
'across all specs in the same linelist build.
'@return ILLdictionary. The linelist dictionary.
Public Property Get Dictionary() As ILLdictionary: End Property

'@jump:table-id
'@prop-title Unique identifier for this analysis table
'@details
'Returns a string identifier built from a type prefix and the row offset
'within the setup ListObject. Used as a key for tracking, named range
'creation, and graph association.
'@return String. The unique table identifier.
Public Property Get TableId() As String: End Property

'@jump:table-section-id
'@prop-title Identifier of the first table in this section
'@details
'Returns the TableId of the leading table in the current section group.
'Used to associate sibling tables for section-level operations like
'header placement and row alignment.
'@return String. The section leader table identifier.
Public Property Get TableSectionId() As String: End Property

'@jump:is-new-section
'@prop-title Whether this table begins a new section group
'@details
'Returns True when this specification row marks the start of a new
'analysis section. The Analysis class uses this flag to decide when to
'insert section headers on the output worksheet.
'@return Boolean. True when this table starts a new section.
Public Property Get IsNewSection() As Boolean: End Property

'@jump:valid-table
'@prop-title Whether this table has a valid configuration
'@details
'Returns True when the specification row contains all required fields for
'its AnalysisTableScope. Invalid specs are skipped during the build
'process so they do not produce broken output tables.
'@return Boolean. True when the table configuration is complete.
Public Property Get ValidTable() As Boolean: End Property


'@section ValidationProperties
'===============================================================================
'@description Flags controlling optional table features.

'@jump:has-total
'@prop-title Whether totals are needed for computation
'@details
'Returns True when the table requires a total row or column, either
'because the user explicitly set total to "yes" or because a percentage
'mode ("row" or "column") on a time series table implies totals. This
'flag drives cross-table formula generation.
'@return Boolean. True when total computation is needed.
Public Property Get HasTotal() As Boolean: End Property

'@jump:total-requested
'@prop-title Whether total was explicitly requested by the user
'@details
'Returns True only when the setup row has total set to "yes". Unlike
'HasTotal, this flag ignores implied totals from percentage mode. Used
'to decide whether the total column is visible in time series output.
'@return Boolean. True when total was explicitly requested.
Public Property Get TotalRequested() As Boolean: End Property

'@jump:has-missing
'@prop-title Whether missing-data rows or columns are included
'@details
'Returns True when the table should include a row or column for missing
'values. Determined by the "missing" field in the setup row.
'@return Boolean. True when missing data is included.
Public Property Get HasMissing() As Boolean: End Property

'@jump:has-graph
'@prop-title Whether a graph should be created for this table
'@details
'Returns True when the specification row requests a graph to accompany
'the analysis table. The AnaTabIds class uses this flag to register
'graph series and formatting.
'@return Boolean. True when a graph is requested.
Public Property Get HasGraph() As Boolean: End Property

'@jump:has-percentage
'@prop-title Whether percentage display is enabled
'@details
'Returns True when the table should show percentage values. The mode
'(row or column) is read from the setup row and influences how
'cross-table formulas are built.
'@return Boolean. True when percentage is enabled.
Public Property Get HasPercentage() As Boolean: End Property

'@jump:spatial-table-types
'@prop-title Spatial analysis sub-type
'@details
'Returns "hf" for health-facility-based spatial tables or "geo" for
'administrative-boundary-based spatial tables. The sub-type is inferred
'from the variable name prefix. Returns an empty string for non-spatial
'table scopes.
'@return String. The spatial sub-type ("hf" or "geo").
Public Property Get SpatialTableScopes() As String: End Property


'@section NavigationProperties
'===============================================================================
'@description Linked-list navigation between sibling specifications.

'@jump:previous
'@prop-title Previous valid specification in the same section
'@details
'Returns the ITableSpecs instance that precedes this one in the section
'list. Returns Nothing when this is the first table in its section.
'Used by CrossTable to compute start-row offsets and by AnalysisOutput
'to chain table builds.
'@return ITableSpecs. The preceding specification, or Nothing.
Public Property Get Previous() As ITableSpecs: End Property

'@jump:next-specs
'@prop-title Next valid specification following an anchor range
'@details
'Scans the setup ListObject rows after anchorRng to find the next valid
'specification. Returns Nothing when no further valid specs exist. Used
'to advance through the specification chain during the build loop.
'@param anchorRng Range. Anchor cell marking the last row of the current
'ListObject; scanning starts after this point.
'@return ITableSpecs. The next valid specification, or Nothing.
Public Property Get NextSpecs(ByVal anchorRng As Range) As ITableSpecs: End Property


'@section DataAccess
'===============================================================================
'@description Value retrieval and category lookups.

'@jump:value
'@sub-title Retrieve a value from the specification row by column name
'@details
'Looks up the named column in the setup ListObject header and returns
'the string value at the intersection of this specification row and that
'column. Returns an empty string when the column is not found.
'@param colName String. Column header name in the setup ListObject.
'@return String. The cell value as a string.
Public Function Value(ByVal colName As String) As String: End Function

'@jump:row-categories
'@prop-title Row variable categories from the dictionary
'@details
'Returns a BetterArray containing the category labels for the row
'variable defined in this specification. When lData is a LinelistSpecs
'instance, the categories are resolved from its dictionary and choices.
'When lData is any other type (or Nothing), returns an empty BetterArray.
'@param lData Object. A LinelistSpecs instance providing category lookups,
'or any other Object (categories will be empty).
'@return BetterArray. The row category list.
Public Property Get RowCategories(ByRef lData As Object) As BetterArray: End Property

'@jump:column-categories
'@prop-title Column variable categories from the dictionary
'@details
'Returns a BetterArray containing the category labels for the column
'variable defined in this specification. When lData is a LinelistSpecs
'instance, the categories are resolved from its dictionary and choices.
'When lData is any other type (or Nothing), returns an empty BetterArray.
'Returns an empty array for univariate or global summary types.
'@param lData Object. A LinelistSpecs instance providing category lookups,
'or any other Object (categories will be empty).
'@return BetterArray. The column category list.
Public Property Get ColumnCategories(ByRef lData As Object) As BetterArray: End Property
