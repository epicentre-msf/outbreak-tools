VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoLevelService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides helper-backed retrieval of geobase hierarchical levels"

'@Folder("Implement")
'@ModuleDescription("Uses helper classes to replicate LLGeo GeoLevel behaviour")

Option Explicit

'@section Constants
'===============================================================================

Private Const GEO_SCOPE_ADMIN As Byte = 0
Private Const GEO_SCOPE_HF As Byte = 1

Private Const LEVEL_ADMIN1 As Byte = 0
Private Const LEVEL_ADMIN2 As Byte = 1
Private Const LEVEL_ADMIN3 As Byte = 2
Private Const LEVEL_ADMIN4 As Byte = 3

'@section State
'===============================================================================

Private Type TGeoLevelService
    adapter As GeoLegacyAdapter
    tableNames As BetterArray
End Type

Private this As TGeoLevelService
Private Const CLASSNAME As String = "GeoLevelService"

'@section Factory
'===============================================================================

Public Function Create(ByVal adapter As GeoLegacyAdapter) As GeoLevelService
    Dim service As GeoLevelService

    If adapter Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "GeoLegacyAdapter is required"
    End If

    Set service = New GeoLevelService
    Set service.Adapter = adapter
    Set Create = service
End Function

'@section Dependencies
'===============================================================================

Public Property Set Adapter(ByVal adapter As GeoLegacyAdapter)
    Set this.adapter = adapter
    Set this.tableNames = adapter.TableNames
End Property

Public Property Get Adapter() As GeoLegacyAdapter
    Set Adapter = this.adapter
End Property

'@section Public API
'===============================================================================

Public Function FetchLevels(ByVal adminLevel As Byte, _
                            Optional ByVal geoScope As Byte = GEO_SCOPE_ADMIN, _
                            Optional  adminNames As Variant) As BetterArray
    Select Case geoScope
        Case GEO_SCOPE_ADMIN
            Set FetchLevels = FetchAdminLevels(adminLevel, adminNames)
        Case GEO_SCOPE_HF
            Set FetchLevels = FetchHFLevels(adminLevel, adminNames)
        Case Else
            Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Unknown geo scope: " & CStr(geoScope)
    End Select
End Function

'@section Internal helpers
'===============================================================================

Private Function FetchAdminLevels(ByVal adminLevel As Byte,  adminNames As Variant) As BetterArray
    Dim filters As BetterArray
    Dim targetColumn As String
    Dim tableIndex As Long
    Dim result As BetterArray

    Select Case adminLevel
        Case LEVEL_ADMIN1
            Set result = this.adapter.GetColumnValues(1, "adm1_name")
        Case LEVEL_ADMIN2
            ValidateParent adminLevel, adminNames
            Set filters = BuildFilters(Array(Array("adm1_name", CStr(adminNames))))
            Set result = this.adapter.GetFilteredValues(2, "adm2_name", filters)
        Case LEVEL_ADMIN3
            ValidateParent adminLevel, adminNames
            Set filters = BuildHierarchicalFilters(adminNames, 2)
            Set result = this.adapter.GetFilteredValues(3, "adm3_name", filters)
        Case LEVEL_ADMIN4
            ValidateParent adminLevel, adminNames
            Set filters = BuildHierarchicalFilters(adminNames, 3)
            Set result = this.adapter.GetFilteredValues(4, "adm4_name", filters)
        Case Else
            Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Unknown admin level: " & CStr(adminLevel)
    End Select

    Set FetchAdminLevels = result
End Function

Private Function FetchHFLevels(ByVal adminLevel As Byte,  adminNames As Variant) As BetterArray
    Dim filters As BetterArray
    Dim result As BetterArray

    Select Case adminLevel
        Case LEVEL_ADMIN1
            Set result = this.adapter.GetFilteredValues(5, "adm1_name", Nothing)
        Case LEVEL_ADMIN2
            ValidateParent adminLevel, adminNames
            Set filters = BuildFilters(Array(Array("adm1_name", CStr(adminNames))))
            Set result = this.adapter.GetFilteredValues(5, "adm2_name", filters)
        Case LEVEL_ADMIN3
            ValidateParent adminLevel, adminNames
            Set filters = BuildHierarchicalFilters(adminNames, 2)
            Set result = this.adapter.GetFilteredValues(5, "adm3_name", filters)
        Case LEVEL_ADMIN4
            ValidateParent adminLevel, adminNames
            Set filters = BuildHierarchicalFilters(adminNames, 3)
            Set result = this.adapter.GetFilteredValues(5, "hf_name", filters)
        Case Else
            Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Unknown admin level: " & CStr(adminLevel)
    End Select

    Set FetchHFLevels = result
End Function

Private Sub ValidateParent(ByVal adminLevel As Byte,  adminNames As Variant)
    Select Case adminLevel
        Case LEVEL_ADMIN2
            If TypeName(adminNames) <> "String" Then
                Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Admin level 2 expects a parent name as String"
            End If
        Case LEVEL_ADMIN3, LEVEL_ADMIN4
            If TypeName(adminNames) <> "BetterArray" Then
                Err.Raise ProjectError.InvalidArgument, CLASSNAME, _
                          "Admin level " & CStr(adminLevel + 1) & " expects BetterArray of parent names"
            End If
            If adminNames.Length <> adminLevel Then
                Err.Raise ProjectError.InvalidArgument, CLASSNAME, _
                          "Expected " & CStr(adminLevel) & " parent names, received " & CStr(adminNames.Length)
            End If
        Case Else
            'No validation required
    End Select
End Sub

Private Function BuildFilters( entries As Variant) As BetterArray
    Dim filters As BetterArray
    Dim idx As Long

    Set filters = New BetterArray
    filters.LowerBound = 1

    For idx = LBound(entries) To UBound(entries)
        filters.Push entries(idx)
    Next idx

    Set BuildFilters = filters
End Function

Private Function BuildHierarchicalFilters(ByVal adminNames As BetterArray, ByVal depth As Long) As BetterArray
    Dim filters As BetterArray
    Dim levelIndex As Long
    Dim tagName As String

    Set filters = New BetterArray
    filters.LowerBound = 1

    For levelIndex = 1 To depth
        tagName = "adm" & levelIndex & "_name"
        filters.Push Array(tagName, adminNames.Item(levelIndex))
    Next levelIndex

    Set BuildHierarchicalFilters = filters
End Function

End Class
