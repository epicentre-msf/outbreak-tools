VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoColumnResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Resolves geobase column indices using translation cache"

'@Folder("Implement")
'@ModuleDescription("Determines geobase and health facility column indices with translation support")

Option Explicit

'@section State
'===============================================================================

Private Type TGeoColumnResolver
    translationCache As GeoHeaderTranslationCache
    geoTables As BetterArray
End Type

Private this As TGeoColumnResolver
Private Const CLASSNAME As String = "GeoColumnResolver"
Private Const HF_TABLE_NAME As String = "T_HF"

'@section Factory
'===============================================================================

Public Function Create(ByVal translationCache As GeoHeaderTranslationCache, _
                       ByVal geoTables As BetterArray) As GeoColumnResolver
    Dim resolver As GeoColumnResolver

    If translationCache Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Translation cache is required"
    End If

    If geoTables Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Geo table name list is required"
    End If

    Set resolver = New GeoColumnResolver
    Set resolver.TranslationCache = translationCache
    Set resolver.GeoTables = geoTables
    Set Create = resolver
End Function

'@section Dependencies
'===============================================================================

Public Property Set TranslationCache(ByVal cache As GeoHeaderTranslationCache)
    Set this.translationCache = cache
End Property

Public Property Set GeoTables(ByVal tableNames As BetterArray)
    Set this.geoTables = tableNames
End Property

'@section Public API
'===============================================================================

Public Function ResolveAdminColumn(ByVal tableIndex As Long, ByVal columnTag As String) As Long
    Dim tableName As String
    Dim tableObject As ListObject

    ValidateTableIndex tableIndex
    this.translationCache.EnsureLoaded

    tableName = CStr(this.geoTables.Item(tableIndex))
    Set tableObject = ResolveTable(tableName)

    ResolveAdminColumn = ResolveColumnInternal(tableObject, columnTag)
End Function

Public Function ResolveHFColumn(ByVal columnTag As String) As Long
    Dim tableObject As ListObject

    this.translationCache.EnsureLoaded
    Set tableObject = ResolveTable(HF_TABLE_NAME)

    ResolveHFColumn = ResolveColumnInternal(tableObject, columnTag)
End Function

'@section Internal helpers
'===============================================================================

Private Sub ValidateTableIndex(ByVal tableIndex As Long)
    If this.geoTables Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Geo table list is not initialised"
    End If

    If tableIndex < this.geoTables.LowerBound Or tableIndex > this.geoTables.UpperBound Then
        Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Table index is out of range"
    End If
End Sub

Private Function ResolveTable(ByVal tableName As String) As ListObject
    Dim sheet As Worksheet
    Dim tableObject As ListObject

    Set sheet = this.translationCache.TargetSheet

    On Error Resume Next
        Set tableObject = sheet.ListObjects(tableName)
    On Error GoTo 0

    If tableObject Is Nothing Then
        Err.Raise ProjectError.ElementNotFound, CLASSNAME, _
                  "ListObject '" & tableName & "' not found on sheet '" & sheet.Name & "'"
    End If

    Set ResolveTable = tableObject
End Function

Private Function ResolveColumnInternal(ByVal tableObject As ListObject, _
                                       ByVal columnTag As String) As Long
    Dim headerRange As Range
    Dim translatedName As String
    Dim columnIndex As Long

    translatedName = this.translationCache.Resolve(columnTag)
    Set headerRange = tableObject.HeaderRowRange

    columnIndex = LocateColumnIndex(headerRange, translatedName)
    If columnIndex = -1 Then
        columnIndex = LocateColumnIndex(headerRange, columnTag)
    End If

    If columnIndex = -1 Then
        Err.Raise ProjectError.ElementNotFound, CLASSNAME, _
                  "Column '" & columnTag & "' not found in table '" & tableObject.Name & "'"
    End If

    ResolveColumnInternal = columnIndex
End Function

Private Function LocateColumnIndex(ByVal headerRange As Range, ByVal columnName As String) As Long
    Dim matchRange As Range

    If LenB(Trim$(columnName)) = 0 Then
        LocateColumnIndex = -1
        Exit Function
    End If

    On Error Resume Next
        Set matchRange = headerRange.Find(What:=columnName, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If matchRange Is Nothing Then
        LocateColumnIndex = -1
    Else
        LocateColumnIndex = matchRange.Column - headerRange.Column + 1
    End If
End Function

End Class
