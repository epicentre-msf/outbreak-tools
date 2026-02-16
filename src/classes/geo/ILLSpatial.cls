VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSpatial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for spatial analysis table management"

'@Interface
'@Folder("Geo")
'@ModuleDescription("Interface for spatial analysis table management")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ILLSpatial
'@description
'Defines the public contract for the LLSpatial class, exposing methods
'for checking spatial variable existence, updating spatial analysis tables
'from filtered linelist data, sorting tables by formula or attack rate,
'and retrieving ranked geographic or health facility values for use in
'analyses and graphs.


'@section Properties
'===============================================================================

'@jump:exists
'@prop-title Whether a variable has spatial tables
'@details
'Searches the "listofgeovars" ListObject on the spatial sheet for a
'partial match of the supplied variable name. Returns True when the
'variable appears in the spatial variables list, indicating that spatial
'analysis tables exist for it.
'@param varName String. Variable name or partial name to look up.
'@return Boolean. True when the variable exists in the spatial variables list.
'@export
Public Property Get Exists(ByVal varName As String) As Boolean: End Property


'@section Update
'===============================================================================

'@jump:update
'@sub-title Update all spatial tables from filtered linelist data
'@details
'Iterates all HList worksheets, locates their filtered companion sheet,
'finds columns controlled by geo1 and hf, extracts unique values from the
'filtered data, and populates the corresponding spatial ListObjects. After
'updating all tables, recalculates the spatial sheet so formula columns
'reflect the new data.
'@export
Public Sub Update(): End Sub


'@section Ranking
'===============================================================================

'@jump:top-geo-value
'@prop-title Ranked value from a sorted spatial geo table
'@details
'Returns the value at the requested rank position from a sorted spatial
'admin table. The concatValue parameter selects between the name column
'(1) and the concatenated column (other values). Returns an empty string
'when the rank exceeds the available rows or the table does not exist.
'@param adminName String. Admin level name ("adm1", "adm2", "adm3", or "adm4").
'@param orderValue Long. Rank position in the sorted table (1-based).
'@param varName String. Base variable name without admin prefix.
'@param tabId Optional String. Table identifier suffix. Defaults to vbNullString.
'@param concatValue Optional Byte. 1 for the name column, other values for the concat column. Defaults to 1.
'@return String. The value at the requested rank, or empty string if not found.
'@export
Public Property Get TopGeoValue(ByVal adminName As String, _
                                ByVal orderValue As Long, _
                                ByVal varName As String, _
                                Optional ByVal tabId As String = vbNullString, _
                                Optional ByVal concatValue As Byte = 1) As String
End Property

'@jump:top-hf-value
'@prop-title Ranked value from a sorted spatial HF table
'@details
'Returns the value at the requested rank position from a sorted spatial
'health facility table. Returns an empty string when the rank exceeds the
'available rows or the table does not exist.
'@param orderValue Long. Rank position in the sorted table (1-based).
'@param inputVarName String. Base variable name.
'@param tabId Optional String. Table identifier suffix. Defaults to vbNullString.
'@return String. The value at the requested rank, or empty string if not found.
'@export
Public Property Get TopHFValue(ByVal orderValue As Long, _
                               ByVal inputVarName As String, _
                               Optional ByVal tabId As String = vbNullString) As String
End Property


'@section Sort
'===============================================================================

'@jump:sort
'@sub-title Sort spatial tables by formula value or attack rate
'@details
'Locates the spatial variable matching tabId in the "listofgeovars" list,
'then sorts the corresponding admin1 through admin4 ListObjects by either
'the formula column or the attack_rate column in descending order. Skips
'tables that do not exist or have no data.
'@param tabId String. Table identifier used to locate the variable name.
'@param onAR Optional Boolean. True to sort by attack_rate column, False for formula column. Defaults to False.
'@export
Public Sub Sort(ByVal tabId As String, Optional ByVal onAR As Boolean = False): End Sub
