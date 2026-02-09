VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for geobase management"

'@Interface
'@Folder("Geo")
'@ModuleDescription("Interface for geobase management")

Option Explicit

'Geo scope: admin hierarchy or health facility
Public Enum GeoScope
    GeoScopeAdmin = 0
    GeoScopeHF = 1
End Enum

'Admin levels for admin-only scope
Public Enum AdminOnlyScope
    LevelAdmin1 = 0
    LevelAdmin2 = 1
    LevelAdmin3 = 2
    LevelAdmin4 = 3
End Enum

'Admin levels including health facility
Public Enum AdminAndHFScope
    LevelGeoAdmin1 = 0
    LevelGeoAdmin2 = 1
    LevelGeoAdmin3 = 2
    LevelGeoAdmin4 = 3
    LevelHF = 4
End Enum

' @description Get the list of values at a given admin level, optionally filtered
'              by parent admin selections. Returns a BetterArray of unique names.
' @param adminLevel Admin level (0-3 for admin1-4)
' @param geoType GeoScope: Admin (0) or HF (1)
' @param adminNames Parent admin selections for filtering
'                    (String for level 2, BetterArray for levels 3-4)
' @return BetterArray of unique values at the requested level
Public Property Get GeoLevel(ByVal adminLevel As Byte, _
                             Optional ByVal geoType As Byte = GeoScopeAdmin, _
                             Optional adminNames As Variant = vbNullString) As BetterArray
End Property

' @description Convert a raw admin column name to its translated display name.
' @param admname Raw column name (e.g. "adm1_name", "hf_name", "geobase")
' @return Translated display name, or the raw name if no translation exists
Public Property Get GeoNames(ByVal admname As String) As String: End Property

' @description Whether the geobase has no imported data
Public Property Get HasNoData() As Boolean: End Property

' @description Translate or un-translate table headers between raw and display names
' @param rawNames True to revert to raw names, False to apply translated names
Public Sub Translate(Optional ByVal rawNames As Boolean = False): End Sub

' @description Import geobase data from an external workbook
' @param fromWkb Source workbook containing geo data sheets
' @param histoOnly True to import only historic data
Public Sub Import(ByVal fromWkb As Workbook, _
                  Optional ByVal histoOnly As Boolean = False)
End Sub

' @description Export geobase tables to a new workbook (one sheet per table)
' @param outputWkb Target workbook to receive exported sheets
' @param onlyHistory True to export only historic tables
Public Sub ExportToFile(ByVal outputWkb As Workbook, _
                        Optional ByVal onlyHistory As Boolean = False)
End Sub

' @description Export the geobase worksheet structure to a workbook (preserving layout)
' @param Wkb Target workbook
' @param llName Linelist name for metadata
' @param setupPath Setup file path for metadata
Public Sub ExportToWkb(ByVal Wkb As Workbook, _
                       Optional ByVal llName As String = vbNullString, _
                       Optional ByVal setupPath As String = vbNullString)
End Sub

' @description Clear all geobase data
Public Sub Clear(): End Sub

' @description Clear historic entries
' @param scope 0 = geo only, 1 = hf only, 3 = both
Public Sub ClearHistoric(Optional ByVal scope As Byte = 3): End Sub

' @description Add a new entry to the historic data
' @param newValue The value to add
' @param scope 0 = geo historic, 1 = hf historic
Public Sub UpdateHistoric(ByVal newValue As String, _
                          Optional ByVal scope As Byte = 0)
End Sub

' @description Update linelist geo column headers and dropdowns after geobase change
' @param pass Password manager for sheet protection
Public Sub Update(ByVal pass As IPasswords): End Sub

' @description Update dictionary labels for geo-related variables
' @param dict The linelist dictionary
Public Sub UpdateDict(ByVal dict As ILLdictionary): End Sub

' @description Update show/hide tables with translated geo labels
' @param sh The show/hide worksheet
' @param dict The linelist dictionary
Public Sub UpdateShowHide(ByVal sh As Worksheet, ByVal dict As ILLdictionary): End Sub

' @description Resolve the admin code from a translated admin display name
' @param selectedAdmin The translated admin name
' @return Admin code string (e.g. "adm1", "adm2", "adm3", "adm4")
Public Function AdminCode(ByVal selectedAdmin As String) As String: End Function
