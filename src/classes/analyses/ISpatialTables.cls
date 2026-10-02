VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISpatialTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for spatial table creation at build time"

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface for spatial table creation at build time")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ISpatialTables
'@description
'Defines the public contract for the SpatialTables class. Exposes a check
'for existing spatial ListObjects and a method to create new spatial table
'sets on the "spatial_tables__" worksheet at linelist build time.


'@section Methods
'===============================================================================

'@jump:exists
'@prop-title Whether spatial ListObjects already exist for a variable
'@details
'Returns True when the spatial worksheet already contains the full set
'of admin-level ListObjects (four for geo-type, one for hf-type) for
'the specified variable name. Used to avoid creating duplicate tables
'during the build loop.
'@param varName String. Fully qualified spatial variable name
'(e.g. "cases_sp1").
'@return Boolean. True when the spatial ListObjects already exist.
Public Property Get Exists(ByVal varName As String) As Boolean: End Property

'@jump:add
'@sub-title Create spatial ListObjects for the current cross-table
'@details
'Creates a set of ListObjects on the "spatial_tables__" worksheet for the
'current cross-table row variable. For geo-type tables, four admin-level
'tables (adm1 through adm4) are created with formula, population, and
'attack rate columns. For hf-type tables, a single health-facility table
'is created. Skips creation when the tables already exist.
'@param formObject IFormulas. The formula object whose parsed expression
'is written into each spatial table formula column.
'@export
Public Sub Add(ByVal formObject As IFormulas): End Sub
