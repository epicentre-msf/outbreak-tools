VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphSpecsOrchestrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis.Graphs")
'@ModuleDescription("Coordinates cross-table and time-series graph builders to produce chart specifications")
'@depends ProjectError, GraphListObjectUtilities
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IGraphSpecsOrchestrator

Private Type TOrchestratorState
    mode As Byte
    crossTable As ICrossTable
    crossBuilder As ICrossTableGraphBuilder
    timeBuilder As ITimeSeriesGraphBuilder
    graphCache As IGraphSpecsCache
    charts As BetterArray
    chartsValid As Boolean
End Type

Private this As TOrchestratorState

Private Const CLASS_NAME As String = "GraphSpecsOrchestrator"
Private Const MODE_UNINITIALISED As Byte = 0
Private Const MODE_SIMPLE As Byte = 1
Private Const MODE_COMPLEX As Byte = 2

'@section Factory
'===============================================================================
Public Function CreateForCrossTable(ByVal table As ICrossTable, _
                                    Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing) As IGraphSpecsOrchestrator
    Dim orchestrator As GraphSpecsOrchestrator

    Set orchestrator = New GraphSpecsOrchestrator
    orchestrator.InitialiseForCrossTable table, crossBuilder
    Set CreateForCrossTable = orchestrator.Self
End Function

Public Function CreateForTimeSeries(ByVal graphList As ListObject, _
                                    ByVal timeSeriesList As ListObject, _
                                    ByVal titleList As ListObject, _
                                    ByVal linelist As ILinelistSpecs, _
                                    Optional ByVal cache As IGraphSpecsCache = Nothing, _
                                    Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing, _
                                    Optional ByVal timeBuilder As ITimeSeriesGraphBuilder = Nothing, _
                                    Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing, _
                                    Optional ByVal specsFactory As ITimeSeriesTableSpecsFactory = Nothing) As IGraphSpecsOrchestrator
    Dim orchestrator As GraphSpecsOrchestrator

    Set orchestrator = New GraphSpecsOrchestrator
    orchestrator.InitialiseForTimeSeries graphList, timeSeriesList, titleList, linelist, cache, crossBuilder, timeBuilder, seriesBuilder, specsFactory
    Set CreateForTimeSeries = orchestrator.Self
End Function

Public Property Get Self() As IGraphSpecsOrchestrator
    Set Self = Me
End Property

'@section Initialisation helpers
'===============================================================================
Friend Sub InitialiseForCrossTable(ByVal table As ICrossTable, _
                                   ByVal crossBuilder As ICrossTableGraphBuilder)

    ValidateCrossTable table

    If crossBuilder Is Nothing Then
        Set this.crossBuilder = CrossTableGraphBuilder.Create
    Else
        Set this.crossBuilder = crossBuilder
    End If

    Set this.crossTable = table
    this.mode = MODE_SIMPLE
    ResetCharts
End Sub

Friend Sub InitialiseForTimeSeries(ByVal graphList As ListObject, _
                                   ByVal timeSeriesList As ListObject, _
                                   ByVal titleList As ListObject, _
                                   ByVal linelist As ILinelistSpecs, _
                                   ByVal cache As IGraphSpecsCache, _
                                   ByVal crossBuilder As ICrossTableGraphBuilder, _
                                   ByVal timeBuilder As ITimeSeriesGraphBuilder, _
                                   ByVal seriesBuilder As IGraphSeriesBuilder, _
                                   ByVal specsFactory As ITimeSeriesTableSpecsFactory)

    ValidateListObject graphList, "graphList"
    ValidateListObject timeSeriesList, "timeSeriesList"
    ValidateListObject titleList, "titleList"
    ValidateLinelist linelist, "linelist"

    If cache Is Nothing Then
        Set this.graphCache = GraphSpecsCache.Create(graphList)
    Else
        Set this.graphCache = cache
    End If

    If crossBuilder Is Nothing Then
        Set this.crossBuilder = CrossTableGraphBuilder.Create
    Else
        Set this.crossBuilder = crossBuilder
    End If

    If timeBuilder Is Nothing Then
        Set this.timeBuilder = TimeSeriesGraphBuilder.Create(this.graphCache, timeSeriesList, titleList, linelist, seriesBuilder, specsFactory)
    Else
        Set this.timeBuilder = timeBuilder
    End If

    this.mode = MODE_COMPLEX
    ResetCharts
End Sub

'@section Public API
'===============================================================================
Public Function Build() As BetterArray
    EnsureChartsPrimed
    Set Build = CloneCharts
End Function

Public Function Charts() As BetterArray
    EnsureChartsPrimed
    Set Charts = CloneCharts
End Function

Public Property Get GraphCount() As Long
    EnsureChartsPrimed

    If this.charts Is Nothing Then
        GraphCount = 0
    Else
        GraphCount = this.charts.Length
    End If
End Property

Public Property Get SeriesCount() As Long
    Dim idx As Long
    Dim chart As IGraphChartSpec
    Dim count As Long

    EnsureChartsPrimed

    If Not this.charts Is Nothing Then
        For idx = this.charts.LowerBound To this.charts.UpperBound
            Set chart = this.charts.Item(idx)
            If Not chart Is Nothing Then count = count + chart.SeriesCount
        Next idx
    End If

    SeriesCount = count
End Property

Public Property Get IsComplex() As Boolean
    IsComplex = (this.mode = MODE_COMPLEX)
End Property

Public Sub Invalidate()
    ResetCharts

    If Not this.timeBuilder Is Nothing Then
        this.timeBuilder.Invalidate
    End If
End Sub

'@section Internal build logic
'===============================================================================
Private Sub EnsureChartsPrimed()
    If this.chartsValid Then Exit Sub

    ResetChartBuffer

    Select Case this.mode
        Case MODE_SIMPLE: BuildSimpleCharts
        Case MODE_COMPLEX: BuildComplexCharts
        Case Else: RaiseError CLng(ProjectError.ErrorUnexpectedState), "Orchestrator not initialised"
    End Select

    this.chartsValid = True
End Sub

Private Sub ResetCharts()
    this.chartsValid = False
    Set this.charts = Nothing
End Sub

Private Sub ResetChartBuffer()
    Set this.charts = New BetterArray
    this.charts.LowerBound = 1
End Sub

Private Sub BuildSimpleCharts()
    Dim specs As ITablesSpecs
    Dim graphId As String
    Dim graphTitle As String
    Dim chart As IGraphChartSpec

    If this.crossTable Is Nothing Then
        RaiseError CLng(ProjectError.ObjectNotInitialized), "Cross table context not initialised"
    End If

    Set specs = this.crossTable.Specifications
    If specs Is Nothing Then
        RaiseError CLng(ProjectError.ObjectNotInitialized), "Table specifications not available"
    End If

    graphId = ResolveCrossGraphId(specs)
    graphTitle = ResolveCrossGraphTitle(specs)

    Set chart = this.crossBuilder.BuildGraph(this.crossTable, graphId, graphTitle)
    If chart Is Nothing Then Exit Sub

    this.charts.Push chart.Clone
End Sub

Private Sub BuildComplexCharts()
    Dim charts As BetterArray
    Dim idx As Long
    Dim chart As IGraphChartSpec

    If this.timeBuilder Is Nothing Then
        RaiseError CLng(ProjectError.ObjectNotInitialized), "Time-series builder not initialised"
    End If

    Set charts = this.timeBuilder.BuildAll

    If charts Is Nothing Or charts.Length = 0 Then Exit Sub

    For idx = charts.LowerBound To charts.UpperBound
        Set chart = charts.Item(idx)
        If Not chart Is Nothing Then this.charts.Push chart.Clone
    Next idx
End Sub

Private Function CloneCharts() As BetterArray
    Dim cloneBuffer As BetterArray
    Dim idx As Long
    Dim chart As IGraphChartSpec

    Set cloneBuffer = New BetterArray
    cloneBuffer.LowerBound = 1

    If Not this.charts Is Nothing Then
        For idx = this.charts.LowerBound To this.charts.UpperBound
            Set chart = this.charts.Item(idx)
            If Not chart Is Nothing Then cloneBuffer.Push chart.Clone
        Next idx
    End If

    Set CloneCharts = cloneBuffer
End Function

Private Function ResolveCrossGraphId(ByVal specs As ITablesSpecs) As String
    Dim idText As String

    idText = Trim$(SafeSpecValue(specs, "graph id"))
    If idText = vbNullString Then idText = Trim$(specs.TableId)
    If idText = vbNullString Then idText = "graph_" & CStr(CrossTableColumnCount)

    ResolveCrossGraphId = idText
End Function

Private Function ResolveCrossGraphTitle(ByVal specs As ITablesSpecs) As String
    Dim titleText As String

    titleText = Trim$(SafeSpecValue(specs, "graph title"))
    If titleText = vbNullString Then titleText = Trim$(SafeSpecValue(specs, "title"))
    ResolveCrossGraphTitle = titleText
End Function

Private Function SafeSpecValue(ByVal specs As ITablesSpecs, ByVal columnName As String) As String
    On Error Resume Next
        SafeSpecValue = CStr(specs.Value(columnName))
    On Error GoTo 0
End Function

Private Function CrossTableColumnCount() As Long
    If this.crossTable Is Nothing Then
        CrossTableColumnCount = 0
    Else
        CrossTableColumnCount = this.crossTable.NumberOfColumns
    End If
End Function

'@section Validation helpers
'===============================================================================
Private Sub ValidateCrossTable(ByVal table As ICrossTable)
    If table Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "table cannot be Nothing"
    End If
End Sub

Private Sub ValidateListObject(ByVal candidate As ListObject, ByVal argumentName As String)
    If candidate Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
        Exit Sub
    End If

    If StrComp(TypeName(candidate), "ListObject", vbBinaryCompare) <> 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " must be a ListObject"
    End If
End Sub

Private Sub ValidateLinelist(ByVal linelist As ILinelistSpecs, ByVal argumentName As String)
    If linelist Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASS_NAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IGraphSpecsOrchestrator_CreateForCrossTable(ByVal table As ICrossTable, _
                                                             Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing) As IGraphSpecsOrchestrator
    Set IGraphSpecsOrchestrator_CreateForCrossTable = CreateForCrossTable(table, crossBuilder)
End Function

Private Function IGraphSpecsOrchestrator_CreateForTimeSeries(ByVal graphList As ListObject, _
                                                             ByVal timeSeriesList As ListObject, _
                                                             ByVal titleList As ListObject, _
                                                             ByVal linelist As ILinelistSpecs, _
                                                             Optional ByVal cache As IGraphSpecsCache = Nothing, _
                                                             Optional ByVal crossBuilder As ICrossTableGraphBuilder = Nothing, _
                                                             Optional ByVal timeBuilder As ITimeSeriesGraphBuilder = Nothing, _
                                                             Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing, _
                                                             Optional ByVal specsFactory As ITimeSeriesTableSpecsFactory = Nothing) As IGraphSpecsOrchestrator
    Set IGraphSpecsOrchestrator_CreateForTimeSeries = CreateForTimeSeries(graphList, timeSeriesList, titleList, linelist, cache, crossBuilder, timeBuilder, seriesBuilder, specsFactory)
End Function

Private Property Get IGraphSpecsOrchestrator_Self() As IGraphSpecsOrchestrator
    Set IGraphSpecsOrchestrator_Self = Self
End Property

Private Function IGraphSpecsOrchestrator_Build() As BetterArray
    Set IGraphSpecsOrchestrator_Build = Build()
End Function

Private Function IGraphSpecsOrchestrator_Charts() As BetterArray
    Set IGraphSpecsOrchestrator_Charts = Charts()
End Function

Private Property Get IGraphSpecsOrchestrator_GraphCount() As Long
    IGraphSpecsOrchestrator_GraphCount = GraphCount
End Property

Private Property Get IGraphSpecsOrchestrator_SeriesCount() As Long
    IGraphSpecsOrchestrator_SeriesCount = SeriesCount
End Property

Private Property Get IGraphSpecsOrchestrator_IsComplex() As Boolean
    IGraphSpecsOrchestrator_IsComplex = IsComplex
End Property

Private Sub IGraphSpecsOrchestrator_Invalidate()
    Invalidate
End Sub
