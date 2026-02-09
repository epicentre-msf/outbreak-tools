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

' @description Whether a variable already has spatial ListObjects on the spatial sheet.
' @param varName Fully qualified spatial variable name (e.g. "cases_sp1")
' @return True if the 4 admin-level ListObjects (or 1 HF ListObject) already exist
Public Property Get Exists(ByVal varName As String) As Boolean: End Property

' @description Add a spatial table set for the current cross-table's row variable.
'              Creates ListObjects on the "spatial_tables__" sheet with formula,
'              population, and attack rate columns for each admin level (or HF).
' @param formObject The formula object to write into the formula column
Public Sub Add(ByVal formObject As IFormulas): End Sub
