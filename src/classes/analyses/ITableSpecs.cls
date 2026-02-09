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

' Analysis table types matching setup listobject types
Public Enum AnalysisTablesType
    TypeGlobalSummary = 1
    TypeUnivariate = 2
    TypeBivariate = 3
    TypeTimeSeries = 4
    TypeSpatial = 5
    TypeSpatioTemporal = 6
End Enum

' @description One of the AnalysisTablesType enum values
Public Property Get TableType() As AnalysisTablesType: End Property

' @description The dictionary associated with linelist specifications
Public Property Get Dictionary() As ILLdictionary: End Property

' @description Row categories from the dictionary for the row variable
Public Property Get RowCategories() As BetterArray: End Property

' @description Column categories from the dictionary for the column variable
Public Property Get ColumnCategories() As BetterArray: End Property

' @description Whether this table starts a new section
Public Property Get IsNewSection() As Boolean: End Property

' @description Unique identifier for this table (prefix + row offset)
Public Property Get TableId() As String: End Property

' @description Whether this table has valid configuration for its type
Public Property Get ValidTable() As Boolean: End Property

' @description Whether missing data row/column should be included
Public Property Get HasMissing() As Boolean: End Property

' @description Whether total is needed for computation (True when total="yes"
'              OR percentage="row"/"column" for time series)
Public Property Get HasTotal() As Boolean: End Property

' @description Whether total was explicitly requested by user (total="yes" only).
'              Used to decide visibility of total column in time series.
Public Property Get TotalRequested() As Boolean: End Property

' @description Whether a graph should be created for this table
Public Property Get HasGraph() As Boolean: End Property

' @description Whether percentage is enabled for this table
Public Property Get HasPercentage() As Boolean: End Property

' @description Spatial analysis sub-type ("hf" or "geo") based on variable prefix
Public Property Get SpatialTableTypes() As String: End Property

' @description Previous valid table specification in the same section
Public Property Get Previous() As ITableSpecs: End Property

' @description Next valid table specification (anchorRng is last row of listobject)
Public Property Get NextSpecs(ByVal anchorRng As Range) As ITableSpecs: End Property

' @description Id of the first table in the current section
Public Property Get TableSectionId() As String: End Property

' @description Retrieve value from the table specification row by column name
' @param colName Column name in the setup listobject header
' @return String value at the intersection of this row and the named column
Public Function Value(ByVal colName As String) As String: End Function
