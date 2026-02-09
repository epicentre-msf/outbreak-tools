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

' @description Whether a variable has spatial tables on the spatial sheet.
'              Searches the "listofgeovars" ListObject for a partial match.
' @param varName Variable name (or part of it) to look up
' @return True if the variable exists in the spatial variables list
Public Property Get Exists(ByVal varName As String) As Boolean: End Property

' @description Update all spatial tables from the HList filtered data.
'              Iterates all HList worksheets, finds geo1/hf control columns,
'              extracts unique values, and populates the spatial ListObjects.
Public Sub Update(): End Sub

' @description Get the top geo value from a sorted spatial table.
' @param adminName Admin level name (e.g. "adm1", "adm2", "adm3", "adm4")
' @param orderValue Rank position (1-based) in the sorted table
' @param varName Base variable name (without admin prefix)
' @param tabId Table identifier suffix
' @param concatValue 1 for name column, other for concat column
' @return The value at the requested rank, or empty string if not found
Public Property Get TopGeoValue(ByVal adminName As String, _
                                ByVal orderValue As Long, _
                                ByVal varName As String, _
                                Optional ByVal tabId As String = vbNullString, _
                                Optional ByVal concatValue As Byte = 1) As String
End Property

' @description Get the top health facility value from a sorted spatial table.
' @param orderValue Rank position (1-based) in the sorted table
' @param inputVarName Base variable name
' @param tabId Table identifier suffix
' @return The value at the requested rank, or empty string if not found
Public Property Get TopHFValue(ByVal orderValue As Long, _
                               ByVal inputVarName As String, _
                               Optional ByVal tabId As String = vbNullString) As String
End Property

' @description Sort spatial tables by formula value or attack rate.
' @param tabId Table identifier to locate the variable name
' @param onAR True to sort by attack_rate column, False for formula column
Public Sub Sort(ByVal tabId As String, Optional ByVal onAR As Boolean = False): End Sub
