VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoFilterRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Applies AutoFilter criteria and restores worksheet state"

'@Folder("Implement")
'@ModuleDescription("Runs AutoFilter operations for geobase tables and restores original state")

Option Explicit

'@section State
'===============================================================================

Private Type TGeoFilterRunner
    tableObject As ListObject
    originalFilters As Variant
End Type

Private this As TGeoFilterRunner
Private Const CLASSNAME As String = "GeoFilterRunner"

'@section Factory
'===============================================================================

Public Function Create(ByVal tableObject As ListObject) As GeoFilterRunner
    Dim runner As GeoFilterRunner

    If tableObject Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "ListObject is required"
    End If

    Set runner = New GeoFilterRunner
    Set runner.TableObject = tableObject
    Set Create = runner
End Function

'@section Dependencies
'===============================================================================

Public Property Set TableObject(ByVal tableObject As ListObject)
    Set this.tableObject = tableObject
End Property

Public Property Get TableObject() As ListObject
    Set TableObject = this.tableObject
End Property

'@section Public API
'===============================================================================

Public Function ApplyCriteria(ByVal fieldIndex As Long,  criteriaValues As Variant, _
                              Optional ByVal autoClear As Boolean = True) As Range
    Dim lo As ListObject
    Dim criteria As Variant
    Dim visibleRange As Range

    Set lo = TableObject

    If lo Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "ListObject is not initialised"
    End If

    If fieldIndex < 1 Or fieldIndex > lo.ListColumns.Count Then
        Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Field index is out of range"
    End If

    criteria = BuildCriteria(criteriaValues)

    On Error Resume Next
        lo.Range.AutoFilter fieldIndex, criteria
    On Error GoTo 0

    If Not lo.DataBodyRange Is Nothing Then
        On Error Resume Next
            Set visibleRange = lo.DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
    End If

    If autoClear Then Clear
    Set ApplyCriteria = visibleRange
End Function

Public Sub Clear()
    Dim lo As ListObject
    Set lo = TableObject

    If lo Is Nothing Then Exit Sub

    On Error Resume Next
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo 0
End Sub

Private Function BuildCriteria( rawCriteria As Variant) As Variant
    Dim criteriaArray() As Variant
    Dim arrIndex As Long
    Dim source As BetterArray

    If TypeName(rawCriteria) = "BetterArray" Then
        Set source = rawCriteria
        If source.Length = 0 Then
            BuildCriteria = vbNullString
            Exit Function
        End If

        ReDim criteriaArray(1 To source.Length)
        For arrIndex = source.LowerBound To source.UpperBound
            criteriaArray(arrIndex - source.LowerBound + 1) = source.Item(arrIndex)
        Next arrIndex
        BuildCriteria = criteriaArray
    Else
        BuildCriteria = rawCriteria
    End If
End Function

End Class
