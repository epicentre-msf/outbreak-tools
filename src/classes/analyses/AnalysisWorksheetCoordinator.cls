VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisWorksheetCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Coordinates worksheet pipeline execution with table and graph engines")
'@depends ProjectError, AnalysisWorksheetPipeline, AnalysisTableWriterContext, AnalysisGoToEntryCollectionBuilder, AnalysisWorksheetGraphInputsAdapter, AnalysisTablePlanResult, AnalysisArrayUtilities
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisWorksheetCoordinator
Implements IAnalysisWorksheetHandler

Private Type TCoordinatorState
    pipeline As IAnalysisWorksheetPipeline
    tableEngine As IAnalysisTableEngine
    graphEngine As IAnalysisGraphEngine
    safeRunner As IAnalysisSafeRunner
    currentInputs As IAnalysisWorksheetInputs
    aggregateGoTo As AnalysisGoToEntryCollectionBuilder
    dropdownRequests As BetterArray
    currentSectionPrefix As String
End Type

Private this As TCoordinatorState

Private Const CLASSNAME As String = "AnalysisWorksheetCoordinator"

'@section Dependencies
'===============================================================================
Public Property Set Pipeline(ByVal pipeline As IAnalysisWorksheetPipeline)
    ValidatePipeline pipeline
    Set this.pipeline = pipeline
End Property

Private Property Get PipelineInternal() As IAnalysisWorksheetPipeline
    If this.pipeline Is Nothing Then
        Set this.pipeline = New AnalysisWorksheetPipeline
    End If
    Set PipelineInternal = this.pipeline
End Property

Public Property Set TableEngine(ByVal tableEngine As IAnalysisTableEngine)
    ValidateTableEngine tableEngine
    Set this.tableEngine = tableEngine
End Property

Public Property Get GoToEntries() As AnalysisGoToEntryCollectionBuilder
    Set GoToEntries = this.aggregateGoTo
End Property

Public Property Get DropdownRequests() As BetterArray
    Set DropdownRequests = this.dropdownRequests
End Property

Private Property Get TableEngineInternal() As IAnalysisTableEngine
    If this.tableEngine Is Nothing Then
        Set this.tableEngine = AnalysisTableEngine.Create
    End If
    Set TableEngineInternal = this.tableEngine
End Property

Public Property Set GraphEngine(ByVal graphEngine As IAnalysisGraphEngine)
    ValidateGraphEngine graphEngine
    Set this.graphEngine = graphEngine
End Property

Public Property Set SafeRunner(ByVal runner As IAnalysisSafeRunner)
    ValidateSafeRunner runner
    Set this.safeRunner = runner
End Property

Private Property Get GraphEngineInternal() As IAnalysisGraphEngine
    If this.graphEngine Is Nothing Then
        Set this.graphEngine = New AnalysisGraphEngine
    End If
    Set GraphEngineInternal = this.graphEngine
End Property

Private Property Get SafeRunnerInternal() As IAnalysisSafeRunner
    If this.safeRunner Is Nothing Then
        Set this.safeRunner = New AnalysisSafeRunner
    End If
    Set SafeRunnerInternal = this.safeRunner
End Property

'@section Validation helpers
'===============================================================================
Private Sub ValidatePipeline(ByVal pipeline As IAnalysisWorksheetPipeline)
    If pipeline Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "pipeline cannot be Nothing"
    End If
End Sub

Private Sub ValidateTableEngine(ByVal tableEngine As IAnalysisTableEngine)
    If tableEngine Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "tableEngine cannot be Nothing"
    End If
End Sub

Private Sub ValidateGraphEngine(ByVal graphEngine As IAnalysisGraphEngine)
    If graphEngine Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "graphEngine cannot be Nothing"
    End If
End Sub

Private Sub ValidateSafeRunner(ByVal runner As IAnalysisSafeRunner)
    If runner Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "safeRunner cannot be Nothing"
    End If
End Sub

Private Sub ValidateSequence(ByVal sequence As BetterArray)
    If sequence Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "sequence cannot be Nothing"
    End If
End Sub

Private Sub ValidateInputs(ByVal inputs As IAnalysisWorksheetInputs)
    If inputs Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "inputs cannot be Nothing"
    End If
End Sub

Private Sub ValidateSpec(ByVal spec As IAnalysisWorksheetSpec)
    If spec Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "Worksheet specification cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section IAnalysisWorksheetCoordinator implementation
'===============================================================================
Private Sub IAnalysisWorksheetCoordinator_Run(ByVal sequence As BetterArray, ByVal inputs As IAnalysisWorksheetInputs)
    ValidateSequence sequence
    ValidateInputs inputs

    Set this.currentInputs = inputs
    Set this.aggregateGoTo = New AnalysisGoToEntryCollectionBuilder
    Set this.dropdownRequests = AnalysisArrayUtilities.CreateBuffer()

    PipelineInternal.Run sequence, Me

    Set this.currentInputs = Nothing
End Sub

Public Sub Run(ByVal sequence As BetterArray, ByVal inputs As IAnalysisWorksheetInputs)
    IAnalysisWorksheetCoordinator_Run sequence, inputs
End Sub

'@section IAnalysisWorksheetHandler implementation
'===============================================================================
Private Sub IAnalysisWorksheetHandler_BeginSequence()
    'No-op for now; retained for future use.
End Sub

Private Sub IAnalysisWorksheetHandler_ProcessWorksheet(ByVal spec As IAnalysisWorksheetSpec)
    Dim tableIds As BetterArray
    Dim tableTypes As BetterArray
    Dim idx As Long
    Dim listId As String
    Dim tableTypeCode As Byte
    Dim specList As ListObject
    Dim outputSheet As Worksheet
    Dim writerContext As IAnalysisTableWriterContext
    Dim planResult As IAnalysisTablePlanResult
    Dim emptyLabels As BetterArray

    ValidateSpec spec

    Dim scopeCollector As AnalysisGoToEntryCollectionBuilder
    Set scopeCollector = New AnalysisGoToEntryCollectionBuilder

    Set tableIds = spec.TableIds
    Set tableTypes = spec.TableTypeCodes
    Set outputSheet = this.currentInputs.OutputSheet(spec.Scope)

    For idx = tableIds.LowerBound To tableIds.UpperBound
        listId = CStr(tableIds.Item(idx))
        tableTypeCode = CByte(tableTypes.Item(idx))
        Set specList = this.currentInputs.SpecificationList(listId)

        Set writerContext = New AnalysisTableWriterContext

        Dim engineAction As AnalysisEngineAction
        Set engineAction = New AnalysisEngineAction
        engineAction.InitialiseTable TableEngineInternal, specList, outputSheet, this.currentInputs.Linelist, tableTypeCode, writerContext, spec.SectionGoToPrefix, spec.SectionGoToPrefix

        SafeRunnerInternal.Run engineAction, "TableEngine(" & listId & ")"

        CaptureGoToEntries writerContext, scopeCollector

        If spec.HasGraphs Then
            Set emptyLabels = AnalysisArrayUtilities.CreateBuffer()
            Set planResult = AnalysisTablePlanResult.Create(AnalysisArrayUtilities.CloneItems(writerContext.WrittenItems), emptyLabels, AnalysisArrayUtilities.CreateBuffer(), spec.SectionGoToPrefix, spec.SectionGoToPrefix)
            ProcessGraphs spec, planResult
        End If
    Next idx

    CollectDropdownRequests spec, scopeCollector
End Sub

Private Sub IAnalysisWorksheetHandler_CompleteSequence()
    'Future work: apply accumulated GoTo entries to dropdowns.
End Sub

'@section Internal helpers
'===============================================================================
Private Sub CaptureGoToEntries(ByVal writerContext As IAnalysisTableWriterContext, ByVal scopeCollector As AnalysisGoToEntryCollectionBuilder)
    Dim sections As BetterArray
    Dim headers As BetterArray

    Set sections = writerContext.SectionLabels
    Set headers = writerContext.HeaderLabels

    If Not scopeCollector Is Nothing Then
        scopeCollector.AddEntries sections
        scopeCollector.AddEntries headers
    End If

    If Not this.aggregateGoTo Is Nothing Then
        this.aggregateGoTo.AddEntries sections
        this.aggregateGoTo.AddEntries headers
    End If
End Sub

Private Sub ProcessGraphs(ByVal spec As IAnalysisWorksheetSpec, ByVal planResult As IAnalysisTablePlanResult)
    Dim graphSpecs As BetterArray
    Dim idx As Long
    Dim graphSpec As IAnalysisWorksheetGraphSpec
    Dim adapter As AnalysisWorksheetGraphInputsAdapter

    Set graphSpecs = spec.GraphSpecs

    If graphSpecs Is Nothing Then Exit Sub
    If graphSpecs.Length = 0 Then Exit Sub

    For idx = graphSpecs.LowerBound To graphSpecs.UpperBound
        Set graphSpec = graphSpecs.Item(idx)
        If graphSpec Is Nothing Then GoTo ContinueLoop

        If graphSpec.GraphType = AnalysisWorksheetGraphTimeSeries Then
            Set adapter = New AnalysisWorksheetGraphInputsAdapter
            adapter.Initialise this.currentInputs, graphSpec

            Dim graphAction As AnalysisEngineAction
            Set graphAction = New AnalysisEngineAction
            graphAction.InitialiseGraph GraphEngineInternal, planResult, adapter

            SafeRunnerInternal.Run graphAction, "GraphEngine(" & graphSpec.TableId & ")"
        End If

ContinueLoop:
        Set graphSpec = Nothing
        Set adapter = Nothing
    Next idx
End Sub

Private Sub CollectDropdownRequests(ByVal spec As IAnalysisWorksheetSpec, ByVal scopeCollector As AnalysisGoToEntryCollectionBuilder)
    Dim factory As AnalysisGoToDropdownFactory
    Dim requests As BetterArray
    Dim sectionLabel As String
    Dim headerLabel As String
    Dim graphLabel As String
    Dim idx As Long
    Dim request As IAnalysisGoToDropdownRequest

    If scopeCollector Is Nothing Then Exit Sub

    Set factory = New AnalysisGoToDropdownFactory
    sectionLabel = this.currentInputs.GoToLabel("section")
    headerLabel = this.currentInputs.GoToLabel("header")
    graphLabel = this.currentInputs.GoToLabel("graph")

    Set requests = factory.CreateAll(scopeCollector, spec.SectionGoToPrefix, spec.SectionGoToPrefix, spec.GraphGoToPrefix, spec.GraphGoToSuffix, sectionLabel, headerLabel, graphLabel)

    If requests Is Nothing Then Exit Sub
    If requests.Length = 0 Then Exit Sub

    For idx = requests.LowerBound To requests.UpperBound
        Set request = requests.Item(idx)
        If Not request Is Nothing Then this.dropdownRequests.Push request
    Next idx
End Sub
