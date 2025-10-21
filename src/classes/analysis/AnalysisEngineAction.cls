VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisEngineAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Provides a unified IAnalysisAction implementation for both table and graph engine runs.")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisAction

Private Enum AnalysisEngineActionKind
    AnalysisEngineActionKindUndefined = 0
    AnalysisEngineActionKindGraph = 1
    AnalysisEngineActionKindTable = 2
End Enum

Private Type TAnalysisEngineActionState
    kind As AnalysisEngineActionKind
    graphEngine As IAnalysisGraphEngine
    graphPlan As IAnalysisTablePlanResult
    graphInputs As IAnalysisGraphEngineInputs
    tableEngine As IAnalysisTableEngine
    specificationList As ListObject
    outputSheet As Worksheet
    linelist As ILinelistSpecs
    tableTypeCode As Byte
    writerContext As IAnalysisTableWriterContext
    sectionPrefix As String
    headerPrefix As String
End Type

Private this As TAnalysisEngineActionState
Private Const CLASSNAME As String = "AnalysisEngineAction"

'@section Configuration
'===============================================================================
'Prepare the action to execute a graph engine run with shared error handling.
'@param graphEngine The engine instance that will perform the graph execution.
'@param planResult  The table plan providing iteration context for the graph.
'@param inputs      The graph inputs adapter created from the worksheet specification.
Public Sub InitialiseGraph(ByVal graphEngine As IAnalysisGraphEngine, _
                           ByVal planResult As IAnalysisTablePlanResult, _
                           ByVal inputs As IAnalysisGraphEngineInputs)

    ResetState
    ValidateGraphEngine graphEngine
    ValidateGraphPlan planResult
    ValidateGraphInputs inputs

    this.kind = AnalysisEngineActionKindGraph

    Set this.graphEngine = graphEngine
    Set this.graphPlan = planResult
    Set this.graphInputs = inputs
End Sub

'Prepare the action to execute a table engine run with shared error handling.
'@param tableEngine      The table engine responsible for writing analysis tables.
'@param specificationList The listobject that holds the table specification.
'@param outputSheet      Target worksheet for rendered tables.
'@param linelist         Linelist specifications that drive the table data.
'@param tableTypeCode    Encoded table type used by the engine policy resolver.
'@param writerContext    The writer context receiving plan feedback and navigation.
'@param sectionPrefix    Prefix applied to generated GoTo section anchors.
'@param headerPrefix     Prefix applied to generated GoTo header anchors.
Public Sub InitialiseTable(ByVal tableEngine As IAnalysisTableEngine, _
                           ByVal specificationList As ListObject, _
                           ByVal outputSheet As Worksheet, _
                           ByVal linelist As ILinelistSpecs, _
                           ByVal tableTypeCode As Byte, _
                           ByVal writerContext As IAnalysisTableWriterContext, _
                           ByVal sectionPrefix As String, _
                           ByVal headerPrefix As String)

    ResetState
    ValidateTableEngine tableEngine
    ValidateSpecificationList specificationList
    ValidateOutputSheet outputSheet
    ValidateLinelist linelist
    ValidateWriterContext writerContext

    this.kind = AnalysisEngineActionKindTable

    Set this.tableEngine = tableEngine
    Set this.specificationList = specificationList
    Set this.outputSheet = outputSheet
    Set this.linelist = linelist
    this.tableTypeCode = tableTypeCode
    Set this.writerContext = writerContext
    this.sectionPrefix = sectionPrefix
    this.headerPrefix = headerPrefix
End Sub

'@section Execution helpers
'===============================================================================
Private Sub ExecuteGraph()
    If this.graphEngine Is Nothing Then
        RaiseError CLng(ProjectError.InvalidState), "Graph action must be initialised before execution"
    End If

    this.graphEngine.Run this.graphPlan, this.graphInputs
End Sub

Private Sub ExecuteTable()
    If this.tableEngine Is Nothing Then
        RaiseError CLng(ProjectError.InvalidState), "Table action must be initialised before execution"
    End If

    this.tableEngine.Run this.specificationList, _
                         this.outputSheet, _
                         this.linelist, _
                         this.tableTypeCode, _
                         this.writerContext, _
                         this.sectionPrefix, _
                         this.headerPrefix
End Sub

Private Sub ResetState()
    this.kind = AnalysisEngineActionKindUndefined

    Set this.graphEngine = Nothing
    Set this.graphPlan = Nothing
    Set this.graphInputs = Nothing

    Set this.tableEngine = Nothing
    Set this.specificationList = Nothing
    Set this.outputSheet = Nothing
    Set this.linelist = Nothing
    this.tableTypeCode = 0
    Set this.writerContext = Nothing
    this.sectionPrefix = vbNullString
    this.headerPrefix = vbNullString
End Sub

'@section Validation helpers
'===============================================================================
Private Sub ValidateGraphEngine(ByVal graphEngine As IAnalysisGraphEngine)
    If graphEngine Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "graphEngine cannot be Nothing"
    End If
End Sub

Private Sub ValidateGraphPlan(ByVal planResult As IAnalysisTablePlanResult)
    If planResult Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "planResult cannot be Nothing"
    End If
End Sub

Private Sub ValidateGraphInputs(ByVal inputs As IAnalysisGraphEngineInputs)
    If inputs Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "inputs cannot be Nothing"
    End If
End Sub

Private Sub ValidateTableEngine(ByVal tableEngine As IAnalysisTableEngine)
    If tableEngine Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "tableEngine cannot be Nothing"
    End If
End Sub

Private Sub ValidateSpecificationList(ByVal specificationList As ListObject)
    If specificationList Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "specificationList cannot be Nothing"
    End If
End Sub

Private Sub ValidateOutputSheet(ByVal outputSheet As Worksheet)
    If outputSheet Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "outputSheet cannot be Nothing"
    End If
End Sub

Private Sub ValidateLinelist(ByVal linelist As ILinelistSpecs)
    If linelist Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "linelist cannot be Nothing"
    End If
End Sub

Private Sub ValidateWriterContext(ByVal writerContext As IAnalysisTableWriterContext)
    If writerContext Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "writerContext cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Sub IAnalysisAction_Execute()
    Select Case this.kind
        Case AnalysisEngineActionKindGraph
            ExecuteGraph

        Case AnalysisEngineActionKindTable
            ExecuteTable

        Case Else
            RaiseError CLng(ProjectError.InvalidState), "InitialiseGraph or InitialiseTable must be called before Execute"
    End Select
End Sub
