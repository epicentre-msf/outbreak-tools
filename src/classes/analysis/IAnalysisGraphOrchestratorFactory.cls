VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisGraphOrchestratorFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Factory interface producing graph specification orchestrators for analysis outputs")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Create an orchestrator for a cross-table graph scenario.
Public Function CreateForCrossTable(ByVal crossTable As ICrossTable, _
                                    Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing) As IGraphSpecsOrchestrator
End Function

'Create an orchestrator for the time-series graph pipeline.
Public Function CreateForTimeSeries(ByVal graphList As ListObject, _
                                    ByVal timeSeriesList As ListObject, _
                                    ByVal titleList As ListObject, _
                                    ByVal linelist As ILinelistSpecs, _
                                    Optional ByVal cache As IGraphSpecsCache = Nothing, _
                                    Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing, _
                                    Optional ByVal timeBuilder As ITimeSeriesGraphBuilder = Nothing, _
                                    Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing, _
                                    Optional ByVal specsFactory As ITimeSeriesTableSpecsFactory = Nothing) As IGraphSpecsOrchestrator
End Function
