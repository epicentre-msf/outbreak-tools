VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoSpatialServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Facade exposing helper services for LLGeo and LLSpatial"

'@Folder("Implement")
'@ModuleDescription("Returns configured helper instances for geobase and spatial workflows")

Option Explicit

Public Function BuildTranslationCache(ByVal targetSheet As Worksheet) As GeoHeaderTranslationCache
    Set BuildTranslationCache = GeoHeaderTranslationCache.Create(targetSheet)
End Function

Public Function BuildColumnResolver(ByVal cache As GeoHeaderTranslationCache, _
                                    ByVal tableNames As BetterArray) As GeoColumnResolver
    Set BuildColumnResolver = GeoColumnResolver.Create(cache, tableNames)
End Function

Public Function BuildFilterRunner(ByVal tableObject As ListObject) As GeoFilterRunner
    Set BuildFilterRunner = GeoFilterRunner.Create(tableObject)
End Function

Public Function BuildUniqueExtractor() As GeoUniqueExtractor
    Set BuildUniqueExtractor = New GeoUniqueExtractor
End Function

Public Function BuildSpatialScanner(ByVal configSheet As Worksheet) As SpatialVariableScanner
    Set BuildSpatialScanner = SpatialVariableScanner.Create(configSheet)
End Function

Public Function BuildSpatialUpdater(ByVal spatialSheet As Worksheet, _
                                    ByVal tracker As GeoSpatialPerformanceTracker) As SpatialTableUpdater
    Set BuildSpatialUpdater = SpatialTableUpdater.Create(spatialSheet, tracker)
End Function

Public Function NewPerformanceTracker() As GeoSpatialPerformanceTracker
    Set NewPerformanceTracker = New GeoSpatialPerformanceTracker
End Function

End Class
