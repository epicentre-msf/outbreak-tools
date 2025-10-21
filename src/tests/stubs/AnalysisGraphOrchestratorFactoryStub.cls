VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGraphOrchestratorFactoryStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub factory returning configured GraphSpecsOrchestratorStub instances")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IAnalysisGraphOrchestratorFactory

Private Type TFactoryStubState
    crossCharts As BetterArray
    timeCharts As BetterArray
    lastCrossTable As ICrossTable
    lastGraphList As ListObject
    lastTimeSeriesList As ListObject
    lastTitleList As ListObject
    lastLinelist As ILinelistSpecs
End Type

Private this As TFactoryStubState

Public Sub ConfigureCrossTableCharts(ByVal charts As BetterArray)
    Set this.crossCharts = charts
End Sub

Public Sub ConfigureTimeSeriesCharts(ByVal charts As BetterArray)
    Set this.timeCharts = charts
End Sub

Public Property Get LastCrossTable() As ICrossTable
    Set LastCrossTable = this.lastCrossTable
End Property

Public Property Get LastGraphList() As ListObject
    Set LastGraphList = this.lastGraphList
End Property

Public Property Get LastTimeSeriesList() As ListObject
    Set LastTimeSeriesList = this.lastTimeSeriesList
End Property

Public Property Get LastTitleList() As ListObject
    Set LastTitleList = this.lastTitleList
End Property

Public Property Get LastLinelist() As ILinelistSpecs
    Set LastLinelist = this.lastLinelist
End Property

Private Function CreateStub(ByVal charts As BetterArray, Optional ByVal isComplex As Boolean = False) As GraphSpecsOrchestratorStub
    Dim stub As GraphSpecsOrchestratorStub
    Set stub = New GraphSpecsOrchestratorStub
    stub.Configure charts, isComplex
    Set CreateStub = stub
End Function

Private Function IAnalysisGraphOrchestratorFactory_CreateForCrossTable(ByVal crossTable As ICrossTable, _
                                                                       Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing) As IGraphSpecsOrchestrator
    Set this.lastCrossTable = crossTable
    Set IAnalysisGraphOrchestratorFactory_CreateForCrossTable = CreateStub(this.crossCharts, False)
End Function

Private Function IAnalysisGraphOrchestratorFactory_CreateForTimeSeries(ByVal graphList As ListObject, _
                                                                       ByVal timeSeriesList As ListObject, _
                                                                       ByVal titleList As ListObject, _
                                                                       ByVal linelist As ILinelistSpecs, _
                                                                       Optional ByVal cache As IGraphSpecsCache = Nothing, _
                                                                       Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing, _
                                                                       Optional ByVal timeBuilder As ITimeSeriesGraphBuilder = Nothing, _
                                                                       Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing, _
                                                                       Optional ByVal specsFactory As ITimeSeriesTableSpecsFactory = Nothing) As IGraphSpecsOrchestrator
    Set this.lastGraphList = graphList
    Set this.lastTimeSeriesList = timeSeriesList
    Set this.lastTitleList = titleList
    Set this.lastLinelist = linelist

    Set IAnalysisGraphOrchestratorFactory_CreateForTimeSeries = CreateStub(this.timeCharts, True)
End Function
