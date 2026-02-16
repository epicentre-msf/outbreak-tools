Attribute VB_Name = "CustomSetupFunctions"
Option Explicit
'Custom functions for the setup
'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString
'@Folder("User Define Functions")

Private Function EventService() As IEventSetup
    Set EventService = SetupEventsManager.EventSetupService
End Function

'@section Headers
'===============================================================================
'@description Build the header used in analysis time-series tables.
Public Function TimeSeriesHeader(ByVal timeVar As String, _
                                 ByVal grpVar As String, _
                                 ByVal sumLab As String) As String
Attribute TimeSeriesHeader.VB_Description = "Get the headers for the time series"
    Application.Volatile
    TimeSeriesHeader = EventService.BuildTimeSeriesHeader(timeVar, grpVar, sumLab)
End Function

'@section Analysis lookups
'===============================================================================
'@description Retrieve values from the analysis graph titles table.
Public Function GraphValue(ByVal graphTitle As String, _
                           Optional ByVal graphCol As String = "Graph ID") As String
Attribute GraphValue.VB_Description = "Get a graph value from the label on graph table"
    Application.Volatile
    GraphValue = EventService.AnalysisGraphValue(graphTitle, graphCol)
End Function

'@description Retrieve values from the analysis time-series table.
Public Function TSValue(ByVal tsTitle As String, _
                        Optional ByVal tsCol As String = "Series ID") As String
Attribute TSValue.VB_Description = "Get a time series value from the time series table"
    Application.Volatile
    TSValue = EventService.AnalysisTimeSeriesValue(tsTitle, tsCol)
End Function

'@description Retrieve values from the spatio-temporal specifications table.
Public Function SpatTempValue(ByVal spSection As String, _
                              Optional ByVal spCol As String = "N geo max") As String
Attribute SpatTempValue.VB_Description = "Get the Spatio-temporal Geo max from the label on spatio-temporal table"
    Application.Volatile
    SpatTempValue = EventService.SpatioTemporalSpecValue(spSection, spCol)
End Function
