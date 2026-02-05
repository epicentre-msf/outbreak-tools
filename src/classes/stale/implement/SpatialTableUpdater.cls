VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpatialTableUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Applies spatial variable updates using helper services"

'@Folder("Implement")
'@ModuleDescription("Coordinates spatial variable updates via helper services")

Option Explicit

'@section State
'===============================================================================

Private Type TSpatialTableUpdater
    spatialSheet As Worksheet
    filterRunnerFactory As Object 'placeholder for dependency injection
    writeBufferFactory As Object
    performanceTracker As GeoSpatialPerformanceTracker
End Type

Private this As TSpatialTableUpdater
Private Const CLASSNAME As String = "SpatialTableUpdater"

'@section Factory
'===============================================================================

Public Function Create(ByVal spatialSheet As Worksheet, _
                       ByVal tracker As GeoSpatialPerformanceTracker) As SpatialTableUpdater
    Dim updater As SpatialTableUpdater

    If spatialSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Spatial worksheet is required"
    End If

    If tracker Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Performance tracker is required"
    End If

    Set updater = New SpatialTableUpdater
    Set updater.SpatialSheet = spatialSheet
    Set updater.PerformanceTracker = tracker
    Set Create = updater
End Function

'@section Dependencies
'===============================================================================

Public Property Set SpatialSheet(ByVal sh As Worksheet)
    Set this.spatialSheet = sh
End Property

Public Property Get SpatialSheet() As Worksheet
    Set SpatialSheet = this.spatialSheet
End Property

Public Property Set PerformanceTracker(ByVal tracker As GeoSpatialPerformanceTracker)
    Set this.performanceTracker = tracker
End Property

Public Property Get PerformanceTracker() As GeoSpatialPerformanceTracker
    Set PerformanceTracker = this.performanceTracker
End Property

'@section Public API
'===============================================================================

Public Sub ApplyUpdates(ByVal variableMetadata As BetterArray)
    If variableMetadata Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Variable metadata is required"
    End If

    this.performanceTracker.Reset
    this.performanceTracker.BeginMeasurement
    WriteMetadata variableMetadata
    this.performanceTracker.EndMeasurement
End Sub

Private Sub WriteMetadata(ByVal variableMetadata As BetterArray)
    Dim sheet As Worksheet
    Dim rowIndex As Long
    Dim idx As Long

    If variableMetadata.Length = 0 Then Exit Sub

    Set sheet = SpatialSheet
    sheet.Range("A:A").ClearContents

    rowIndex = 1
    For idx = variableMetadata.LowerBound To variableMetadata.UpperBound
        sheet.Cells(rowIndex, 1).Value = variableMetadata.Item(idx)
        rowIndex = rowIndex + 1
    Next idx
End Sub

End Class
