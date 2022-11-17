VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Interface of Geo")
'@IgnoreModule

Option Explicit

Public Enum CustomGeoType
    CustomTypeGeo = 0                            'Geo
    CustomTypeHF = 1                             'Health Facility
End Enum

Public Enum GeoLevel
    LevelAdmin1 = 0
    LevelAdmin2 = 1
    LevelAdmin3 = 2
    LevelAdmin4 = 3
End Enum

'Get the list of levels of one admin name (will return either a BetterArray or
' a String(for admin4))
Public Property Get GeoLevel(ByVal adminLevel As GeoLevel, _
                             Optional ByVal geoType As CustomGeoType = CustomTypeGeo, _
                             Optional ByVal adminNames As Variant = vbNullString _
                             ) As BetterArray
End Property

'Convert the code name to the actual name of the current language
Public Property Get GeoNames(ByVal admname As String) As String: End Property

'Get the Population of one administrative level (only for geoType)
Public Property Get Population(ByVal adminLevel As GeoLevel, _
                               ByVal adminNames As Variant) As String
End Property

'Translation
Public Sub Translate(Optional ByVal rawNames As Boolean = False): End Sub

'Import a new Geobase from a workbook
Public Sub Import(ByVal fromWkb As Workbook): End Sub

'Export to a workbook
Public Sub ExportToFile(ByVal outputWkb As Workbook, Optional ByVal onlyHistory As Boolean = False)
End Sub

Public Sub ExportToWkb(ByVal Wkb As Workbook): End Sub
Public Sub Clear(): End Sub

Public Sub ClearHistoric(): End Sub

'Update new Geobase
'Public Sub Update(ByVal dict As ILLdictionary, pass As IPasswords): End Sub


