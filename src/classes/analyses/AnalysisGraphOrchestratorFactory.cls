VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisGraphOrchestratorFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Default implementation creating GraphSpecsOrchestrator instances for analysis graphs")
'@depends ProjectError, GraphSpecsOrchestrator
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisGraphOrchestratorFactory

Private Const CLASSNAME As String = "AnalysisGraphOrchestratorFactory"

'@section Validation helpers
'===============================================================================
Private Sub ValidateCrossTable(ByVal crossTable As ICrossTable)
    If crossTable Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "crossTable cannot be Nothing"
    End If
End Sub

Private Sub ValidateListObject(ByVal candidate As ListObject, ByVal argumentName As String)
    If candidate Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub ValidateLinelist(ByVal linelist As ILinelistSpecs)
    If linelist Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "linelist cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IAnalysisGraphOrchestratorFactory_CreateForCrossTable(ByVal crossTable As ICrossTable, _
                                                                       Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing) As IGraphSpecsOrchestrator
    ValidateCrossTable crossTable
    Set IAnalysisGraphOrchestratorFactory_CreateForCrossTable = GraphSpecsOrchestrator.CreateForCrossTable(crossTable, crossBuilder)
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
    ValidateListObject graphList, "graphList"
    ValidateListObject(timeSeriesList, "timeSeriesList")
    ValidateListObject(titleList, "titleList")
    ValidateLinelist linelist

    Set IAnalysisGraphOrchestratorFactory_CreateForTimeSeries = GraphSpecsOrchestrator.CreateForTimeSeries( _
        graphList, timeSeriesList, titleList, linelist, cache, crossBuilder, timeBuilder, seriesBuilder, specsFactory)
End Function
