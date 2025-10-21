VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoHeaderTranslationCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Caches translated header names for LLGeo tables"

'@Folder("Implement")
'@ModuleDescription("Caches and resolves translated header names for geobase tables")

'@section State
'===============================================================================

Option Explicit

Private Type TGeoHeaderTranslationCache
    targetSheet As Worksheet
    keys As BetterArray
    values As BetterArray
    loaded As Boolean
End Type

Private this As TGeoHeaderTranslationCache
Private Const CLASSNAME As String = "GeoHeaderTranslationCache"

'@section Factory
'===============================================================================

Public Function Create(ByVal targetSheet As Worksheet) As GeoHeaderTranslationCache
    Dim cache As GeoHeaderTranslationCache

    If targetSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Target worksheet is required"
    End If

    Set cache = New GeoHeaderTranslationCache
    Set cache.TargetSheet = targetSheet
    Set Create = cache
End Function

'@section Dependencies
'===============================================================================

Public Property Set TargetSheet(ByVal sh As Worksheet)
    Set this.targetSheet = sh
End Property

Public Property Get TargetSheet() As Worksheet
    Set TargetSheet = this.targetSheet
End Property

'@section Public API
'===============================================================================

Public Sub EnsureLoaded()
    Dim sheet As Worksheet
    Dim mapping As Variant

    If this.loaded Then Exit Sub

    Set sheet = TargetSheet
    Set this.keys = New BetterArray
    this.keys.LowerBound = 1
    Set this.values = New BetterArray
    this.values.LowerBound = 1

    For Each mapping In HeaderMappings()
        AddMapping CStr(mapping(0)), ResolveHeaderValue(sheet, CStr(mapping(1)), CStr(mapping(0)))
    Next mapping

    this.loaded = True
End Sub

Public Function Resolve(ByVal tagName As String) As String
    Dim idx As Long
    Dim normalised As String

    EnsureLoaded

    normalised = LCase$(Trim$(tagName))
    If this.keys Is Nothing Then
        Resolve = tagName
        Exit Function
    End If

    If this.keys.Length > 0 Then
        For idx = this.keys.LowerBound To this.keys.UpperBound
            If StrComp(CStr(this.keys.Item(idx)), normalised, vbTextCompare) = 0 Then
                Resolve = CStr(this.values.Item(idx))
                Exit Function
            End If
        Next idx
    End If

    Resolve = tagName
End Function

Public Sub Reset()
    this.loaded = False
    Set this.keys = Nothing
    Set this.values = Nothing
End Sub

'@section Internal helpers
'===============================================================================

Private Function HeaderMappings() As Variant
    HeaderMappings = Array( _
        Array("adm1_name", "RNG_ADM1NAME"), _
        Array("adm2_name", "RNG_ADM2NAME"), _
        Array("adm3_name", "RNG_ADM3NAME"), _
        Array("adm4_name", "RNG_ADM4NAME"), _
        Array("hf_name", "RNG_HFNAME"), _
        Array("geobase", "RNG_GeoName"))
End Function

Private Function ResolveHeaderValue(ByVal sheet As Worksheet, _
                                    ByVal rangeName As String, _
                                    ByVal fallback As String) As String
    Dim headerRange As Range
    Dim headerValue As Variant

    On Error Resume Next
        Set headerRange = sheet.Range(rangeName)
    On Error GoTo 0

    If headerRange Is Nothing Then
        ResolveHeaderValue = fallback
        Exit Function
    End If

    headerValue = headerRange.Value
    If TypeName(headerValue) = "String" Then
        ResolveHeaderValue = headerValue
    Else
        ResolveHeaderValue = fallback
    End If
End Function

Private Sub AddMapping(ByVal tagName As String, ByVal headerValue As String)
    Dim idx As Long
    Dim normalised As String

    normalised = LCase$(Trim$(tagName))

    If this.keys.Length > 0 Then
        For idx = this.keys.LowerBound To this.keys.UpperBound
            If StrComp(CStr(this.keys.Item(idx)), normalised, vbTextCompare) = 0 Then
                this.values.Item(idx) = headerValue
                Exit Sub
            End If
        Next idx
    End If

    this.keys.Push normalised
    this.values.Push headerValue
End Sub

End Class
