VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoLegacyAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Adapter exposing helper-based operations for LLGeo"

'@Folder("Implement")
'@ModuleDescription("Provides helper-backed retrieval of geobase column values")

Option Explicit

'@section State
'===============================================================================

Private Type TGeoLegacyAdapter
    services As GeoSpatialServices
    translationCache As GeoHeaderTranslationCache
    columnResolver As GeoColumnResolver
    uniqueExtractor As GeoUniqueExtractor
    targetSheet As Worksheet
    tableNames As BetterArray
End Type

Private this As TGeoLegacyAdapter
Private Const CLASSNAME As String = "GeoLegacyAdapter"

'@section Factory
'===============================================================================

Public Function Create(ByVal geoSheet As Worksheet, ByVal tableNames As BetterArray) As GeoLegacyAdapter
    Dim adapter As GeoLegacyAdapter

    If geoSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Geobase worksheet is required"
    End If

    If tableNames Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Table name collection is required"
    End If

    Set adapter = New GeoLegacyAdapter
    adapter.Initialise geoSheet, tableNames
    Set Create = adapter
End Function

Private Sub Initialise(ByVal geoSheet As Worksheet, ByVal tableNames As BetterArray)
    Set this.services = New GeoSpatialServices
    Set this.translationCache = this.services.BuildTranslationCache(geoSheet)
    Set this.columnResolver = this.services.BuildColumnResolver(this.translationCache, tableNames)
    Set this.uniqueExtractor = this.services.BuildUniqueExtractor
    Set this.targetSheet = geoSheet
    Set this.tableNames = tableNames
End Sub

'@section Public API
'===============================================================================

Public Function GetColumnValues(ByVal tableIndex As Long, _
                                ByVal columnTag As String, _
                                Optional ByVal ignoreBlanks As Boolean = True) As BetterArray
    Dim tableObject As ListObject
    Dim columnIndex As Long
    Dim dataRange As Range

    ValidateColumnTag columnTag
    columnIndex = this.columnResolver.ResolveAdminColumn(tableIndex, columnTag)
    Set tableObject = ResolveTable(tableIndex)

    If tableObject.ListColumns(columnIndex).DataBodyRange Is Nothing Then
        Set GetColumnValues = New BetterArray
        Exit Function
    End If

    Set dataRange = tableObject.ListColumns(columnIndex).DataBodyRange
    Set GetColumnValues = this.uniqueExtractor.Extract(dataRange, ignoreBlanks)
End Function

Public Function GetFilteredValues(ByVal tableIndex As Long, _
                                  ByVal columnTag As String, _
                                  ByVal filterPairs As BetterArray, _
                                  Optional ByVal ignoreBlanks As Boolean = True) As BetterArray
    Dim tableObject As ListObject
    Dim runner As GeoFilterRunner
    Dim columnIndex As Long
    Dim filterIndex As Long
    Dim filterEntry As Variant
    Dim visibleRange As Range

    ValidateColumnTag columnTag
    columnIndex = this.columnResolver.ResolveAdminColumn(tableIndex, columnTag)
    Set tableObject = ResolveTable(tableIndex)

    If filterPairs Is Nothing Or filterPairs.Length = 0 Then
        Set GetFilteredValues = GetColumnValues(tableIndex, columnTag, ignoreBlanks)
        Exit Function
    End If

    Set runner = this.services.BuildFilterRunner(tableObject)

    For filterIndex = filterPairs.LowerBound To filterPairs.UpperBound
        filterEntry = filterPairs.Item(filterIndex)
        If IsArray(filterEntry) Then
            runner.ApplyCriteria ResolveFilterIndex(tableIndex, CStr(filterEntry(0))), _
                                filterEntry(1), False
        End If
    Next filterIndex

    On Error Resume Next
        Set visibleRange = tableObject.ListColumns(columnIndex).DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    runner.Clear

    If visibleRange Is Nothing Then
        Set GetFilteredValues = New BetterArray
    Else
        Set GetFilteredValues = this.uniqueExtractor.Extract(visibleRange, ignoreBlanks)
    End If
End Function

Public Function ResolveHeader(ByVal columnTag As String) As String
    ValidateColumnTag columnTag
    ResolveHeader = this.translationCache.Resolve(columnTag)
End Function

Public Property Get TableNames() As BetterArray
    Set TableNames = this.tableNames.Clone
End Property

Public Property Get Worksheet() As Worksheet
    Set Worksheet = this.targetSheet
End Property

'@section Internal helpers
'===============================================================================

Private Sub ValidateColumnTag(ByVal columnTag As String)
    If LenB(Trim$(columnTag)) = 0 Then
        Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Column tag cannot be empty"
    End If
End Sub

Private Function ResolveTable(ByVal tableIndex As Long) As ListObject
    Dim tableName As String
    Dim tableObject As ListObject

    If tableIndex < this.tableNames.LowerBound Or tableIndex > this.tableNames.UpperBound Then
        Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Table index is out of range"
    End If

    tableName = CStr(this.tableNames.Item(tableIndex))

    On Error Resume Next
        Set tableObject = this.targetSheet.ListObjects(tableName)
    On Error GoTo 0

    If tableObject Is Nothing Then
        Err.Raise ProjectError.ElementNotFound, CLASSNAME, _
                  "ListObject '" & tableName & "' not found on sheet '" & this.targetSheet.Name & "'"
    End If

    Set ResolveTable = tableObject
End Function

Private Function ResolveFilterIndex(ByVal tableIndex As Long, ByVal columnTag As String) As Long
    ResolveFilterIndex = this.columnResolver.ResolveAdminColumn(tableIndex, columnTag)
End Function

End Class
