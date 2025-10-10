VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLGeoDropdownStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests.Stubs")
'@ModuleDescription("Lightweight ILLGeo stub used to feed dropdown data during tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLGeo

Private Type TGeoStub
    levels As Collection
End Type

Private this As TGeoStub

Private Sub Class_Initialize()
    Set this.levels = New Collection
End Sub

Public Sub SetLevel(ByVal adminLevel As Byte, ByVal values As Variant)
    '@description Store a BetterArray of values associated with the supplied admin level.
    Dim categoryList As BetterArray
    Dim idx As Long
    Dim key As String

    key = CStr(adminLevel)

    If TypeName(values) = "BetterArray" Then
        Set categoryList = values
    Else
        Set categoryList = New BetterArray
        categoryList.LowerBound = 1

        If IsArray(values) Then
            For idx = LBound(values) To UBound(values)
                categoryList.Push values(idx)
            Next idx
        ElseIf LenB(CStr(values)) > 0 Then
            categoryList.Push values
        End If
    End If

    On Error Resume Next
        this.levels.Remove key
    On Error GoTo 0

    this.levels.Add categoryList, key
End Sub

Private Function ResolveLevel(ByVal adminLevel As Byte) As BetterArray
    Dim key As String
    Dim stored As Variant
    Dim emptyList As BetterArray

    key = CStr(adminLevel)

    On Error Resume Next
        Set stored = this.levels.Item(key)
    On Error GoTo 0

    If TypeName(stored) = "BetterArray" Then
        Set ResolveLevel = stored
    Else
        Set emptyList = New BetterArray
        emptyList.LowerBound = 1
        Set ResolveLevel = emptyList
    End If
End Function

Private Property Get ILLGeo_GeoLevel(ByVal adminLevel As Byte, _
                                     Optional ByVal geoType As Byte = GeoScopeAdmin, _
                                     Optional adminNames As Variant = vbNullString) As BetterArray
    Set ILLGeo_GeoLevel = ResolveLevel(adminLevel)
End Property

Private Property Get ILLGeo_GeoNames(ByVal admname As String) As String
    ILLGeo_GeoNames = admname
End Property

Private Property Get ILLGeo_HasNoData() As Boolean
    ILLGeo_HasNoData = (this.levels.Count = 0)
End Property

Private Sub ILLGeo_Translate(Optional ByVal rawNames As Boolean = False)
    'No-op for tests
End Sub

Private Sub ILLGeo_Import(ByVal fromWkb As Workbook, Optional ByVal histoOnly As Boolean = False)
    'Not required for tests
End Sub

Private Sub ILLGeo_ExportToFile(ByVal outputWkb As Workbook, Optional ByVal onlyHistory As Boolean = False)
    'Not required for tests
End Sub

Private Sub ILLGeo_ExportToWkb(ByVal Wkb As Workbook, _
                               Optional ByVal llName As String = vbNullString, _
                               Optional ByVal setupPath As String = vbNullString)
    'Not required for tests
End Sub

Private Sub ILLGeo_Clear()
    Set this.levels = New Collection
End Sub

Private Sub ILLGeo_ClearHistoric(Optional ByVal scope As Byte = 3)
    'Not required for tests
End Sub

Private Sub ILLGeo_UpdateHistoric(ByVal newValue As String, Optional ByVal scope As Byte = 0)
    'Not required for tests
End Sub

Private Sub ILLGeo_Update(ByVal pass As IPasswords)
    'Not required for tests
End Sub

Private Sub ILLGeo_UpdateDict(ByVal dict As ILLdictionary)
    'Not required for tests
End Sub

Private Sub ILLGeo_UpdateShowHide(ByVal sh As Worksheet, ByVal dict As ILLdictionary)
    'Not required for tests
End Sub

Private Property Get ILLGeo_AdminCode(ByVal selectedAdmin As String) As String
    ILLGeo_AdminCode = selectedAdmin
End Property
