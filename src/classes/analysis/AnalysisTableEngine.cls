VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Coordinates sequence, planning and writing stages for analysis table generation")
'@depends ProjectError, AnalysisTableSequenceBuilder, AnalysisTablePlanBuilder, AnalysisTableWriter
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IAnalysisTableEngine

Private Type TEngineState
    sequenceBuilder As IAnalysisTableSequenceBuilder
    planBuilder As IAnalysisTablePlanBuilder
    writer As IAnalysisTableWriter
End Type

Private this As TEngineState
Private Const CLASSNAME As String = "AnalysisTableEngine"

'@section Factory
'===============================================================================
'Create an engine optionally using supplied collaborators.
Public Function Create(Optional ByVal sequenceBuilder As IAnalysisTableSequenceBuilder = Nothing, _
                       Optional ByVal planBuilder As IAnalysisTablePlanBuilder = Nothing, _
                       Optional ByVal writer As IAnalysisTableWriter = Nothing) As IAnalysisTableEngine
    With New AnalysisTableEngine
        If sequenceBuilder Is Nothing Then
            Set .SequenceBuilder = Nothing
        Else
            Set .SequenceBuilder = sequenceBuilder
        End If

        If planBuilder Is Nothing Then
            Set .PlanBuilder = Nothing
        Else
            Set .PlanBuilder = planBuilder
        End If

        If writer Is Nothing Then
            Set .Writer = Nothing
        Else
            Set .Writer = writer
        End If

        Set Create = .Self
    End With
End Function

'Expose the instance as its interface type.
Public Property Get Self() As IAnalysisTableEngine
    Set Self = Me
End Property

'@section Properties
'===============================================================================
Public Property Set SequenceBuilder(ByVal sequenceBuilder As IAnalysisTableSequenceBuilder)
    If Not sequenceBuilder Is Nothing Then ValidateSequenceBuilder sequenceBuilder
    Set this.sequenceBuilder = sequenceBuilder
End Property

Private Property Get SequenceBuilderInternal() As IAnalysisTableSequenceBuilder
    Set SequenceBuilderInternal = this.sequenceBuilder
End Property

Public Property Set PlanBuilder(ByVal planBuilder As IAnalysisTablePlanBuilder)
    If Not planBuilder Is Nothing Then ValidatePlanBuilder planBuilder
    Set this.planBuilder = planBuilder
End Property

Private Property Get PlanBuilderInternal() As IAnalysisTablePlanBuilder
    Set PlanBuilderInternal = this.planBuilder
End Property

Public Property Set Writer(ByVal writer As IAnalysisTableWriter)
    If Not writer Is Nothing Then ValidateWriter writer
    Set this.writer = writer
End Property

Private Property Get WriterInternal() As IAnalysisTableWriter
    Set WriterInternal = this.writer
End Property

'@section Validation helpers
'===============================================================================
Private Sub ValidateSequenceBuilder(ByVal sequenceBuilder As IAnalysisTableSequenceBuilder)
    If sequenceBuilder Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "sequenceBuilder cannot be Nothing"
    End If
End Sub

Private Sub ValidatePlanBuilder(ByVal planBuilder As IAnalysisTablePlanBuilder)
    If planBuilder Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "planBuilder cannot be Nothing"
    End If
End Sub

Private Sub ValidateWriter(ByVal writer As IAnalysisTableWriter)
    If writer Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "writer cannot be Nothing"
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

Private Sub ValidateLinelist(ByVal linelistSpecs As ILinelistSpecs)
    If linelistSpecs Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "linelistSpecs cannot be Nothing"
    End If
End Sub

Private Sub ValidateContext(ByVal writerContext As IAnalysisTableWriterContext)
    If writerContext Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "writerContext cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Internal helpers
'===============================================================================
Private Function BuildSequence(ByVal specificationList As ListObject, _
                               ByVal linelistSpecs As ILinelistSpecs, _
                               ByVal tableTypeCode As Byte) As BetterArray
    Dim builder As IAnalysisTableSequenceBuilder

    Set builder = SequenceBuilderInternal
    If builder Is Nothing Then
        Set builder = AnalysisTableSequenceBuilder.Create(linelistSpecs)
        Set this.sequenceBuilder = builder
    End If

    Set BuildSequence = builder.Build(specificationList, linelistSpecs, tableTypeCode)
End Function

Private Function BuildPlan(ByVal policyResults As BetterArray, _
                           ByVal sectionLabelPrefix As String, _
                           ByVal headerLabelPrefix As String) As IAnalysisTablePlanResult
    Dim builder As IAnalysisTablePlanBuilder

    Set builder = PlanBuilderInternal
    If builder Is Nothing Then
        Set builder = New AnalysisTablePlanBuilder
        Set this.planBuilder = builder
    End If

    Set BuildPlan = builder.Build(policyResults, sectionLabelPrefix, headerLabelPrefix)
End Function

'@section Interface implementation
'===============================================================================
Private Sub IAnalysisTableEngine_Run(ByVal specificationList As ListObject, _
                                     ByVal outputSheet As Worksheet, _
                                     ByVal linelistSpecs As ILinelistSpecs, _
                                     ByVal tableTypeCode As Byte, _
                                     ByVal writerContext As IAnalysisTableWriterContext, _
                                     Optional ByVal sectionLabelPrefix As String = vbNullString, _
                                     Optional ByVal headerLabelPrefix As String = vbNullString)

    Dim sequenceResults As BetterArray
    Dim plan As IAnalysisTablePlanResult

    ValidateSpecificationList specificationList
    ValidateOutputSheet outputSheet
    ValidateLinelist linelistSpecs
    ValidateContext writerContext

    Set sequenceResults = BuildSequence(specificationList, linelistSpecs, tableTypeCode)
    Set plan = BuildPlan(sequenceResults, sectionLabelPrefix, headerLabelPrefix)

    Dim writerInstance As IAnalysisTableWriter
    Set writerInstance = WriterInternal
    If writerInstance Is Nothing Then
        Set writerInstance = AnalysisTableWriter.Create
        Set this.writer = writerInstance
    End If

    writerInstance.Write plan, outputSheet, writerContext
End Sub
