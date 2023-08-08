VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Geobase")
'@ModuleDescription("Interface of Geo")
'@Interface

Option Explicit

Public Enum GeoScope
    GeoScopeAdmin = 0                            'Geo
    GeoScopeHF = 1                             'Health Facility
End Enum

Public Enum AdminOnlyScope
    LevelAdmin1 = 0
    LevelAdmin2 = 1
    LevelAdmin3 = 2
    LevelAdmin4 = 3
End Enum

Public Enum AdminAndHFScope
    LevelGeoAdmin1 = 0
    LevelGeoAdmin2 = 1
    LevelGeoAdmin3 = 2
    LevelGeoAdmin4 = 3
    LevelHF = 4
End Enum

'Get the list of levels of one admin name (will return either a BetterArray or
' a String(for admin4))
Public Property Get GeoLevel(ByVal adminLevel As Byte, _
                             Optional ByVal geoType As Byte = GeoScopeAdmin, _
                             Optional ByVal adminNames As Variant = vbNullString _
                             ) As BetterArray
End Property

'Convert the code name to the actual name of the current language
Public Property Get GeoNames(ByVal admname As String) As String: End Property
Public Property Get IsEmpty() As Boolean: End Property

'Translation
Public Sub Translate(Optional ByVal rawNames As Boolean = False): End Sub

'Import a new Geobase from a workbook
Public Sub Import(ByVal fromWkb As Workbook, _
                  Optional ByVal histoOnly As Boolean = False)
End Sub

'Export to a workbook
Public Sub ExportToFile(ByVal outputWkb As Workbook, _
                        Optional ByVal onlyHistory As Boolean = False)
End Sub

Public Sub ExportToWkb(ByVal Wkb As Workbook): End Sub
Public Sub Clear(): End Sub
Public Sub ClearHistoric(Optional ByVal scope As Byte = 3): End Sub
'Update historic of geobase
Public Sub UpdateHistoric(ByVal newValue As String, _
                          Optional ByVal scope As Byte = 0)
End Sub
'Update new Geobase
Public Sub Update(ByVal pass As ILLPasswords): End Sub
'Update dictionary with new geobase
Public Sub UpdateDict(ByVal dict As ILLdictionary): End Sub
'Update show hide tables
Public Sub UpdateShowHide(ByVal sh As Worksheet, ByVal dict As ILLdictionary)
End Sub
'Get the admin code from a selected admin
Public Function AdminCode(ByVal selectedAdmin As String) As String:  End Function
