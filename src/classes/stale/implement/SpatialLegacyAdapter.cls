VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpatialLegacyAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Adapter coordinating spatial helper services"

'@Folder("Implement")
'@ModuleDescription("Bridges spatial helpers with legacy LLSpatial workflows")

Option Explicit

'@section State
'===============================================================================

Private Type TSpatialLegacyAdapter
    services As GeoSpatialServices
    scanner As SpatialVariableScanner
    updater As SpatialTableUpdater
    tracker As GeoSpatialPerformanceTracker
End Type

Private this As TSpatialLegacyAdapter
Private Const CLASSNAME As String = "SpatialLegacyAdapter"

'@section Factory
'===============================================================================

Public Function Create(ByVal configSheet As Worksheet, ByVal spatialSheet As Worksheet) As SpatialLegacyAdapter
    Dim adapter As SpatialLegacyAdapter

    If configSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Configuration worksheet is required"
    End If

    If spatialSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Spatial worksheet is required"
    End If

    Set adapter = New SpatialLegacyAdapter
    adapter.Initialise configSheet, spatialSheet
    Set Create = adapter
End Function

Private Sub Initialise(ByVal configSheet As Worksheet, ByVal spatialSheet As Worksheet)
    Set this.services = New GeoSpatialServices
    Set this.tracker = this.services.NewPerformanceTracker
    Set this.scanner = this.services.BuildSpatialScanner(configSheet)
    Set this.updater = this.services.BuildSpatialUpdater(spatialSheet, this.tracker)
End Sub

'@section Public API
'===============================================================================

Public Function RefreshSpatialTables() As BetterArray
    Dim metadata As BetterArray

    Set metadata = this.scanner.ScanVariables
    this.updater.ApplyUpdates metadata
    Set RefreshSpatialTables = metadata.Clone
End Function

Public Property Get DurationMilliseconds() As Double
    DurationMilliseconds = this.tracker.DurationMilliseconds
End Property

Public Property Get Scanner() As SpatialVariableScanner
    Set Scanner = this.scanner
End Property

Public Property Get Updater() As SpatialTableUpdater
    Set Updater = this.updater
End Property

End Class
