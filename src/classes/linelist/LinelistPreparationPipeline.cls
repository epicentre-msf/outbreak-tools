VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistPreparationPipeline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Coordinates Linelist preparation through explicit, testable stages"

'@Folder("Linelist")
'@ModuleDescription("Coordinates Linelist preparation through explicit, testable stages")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation
'@depends ProjectError

Option Explicit

Implements ILinelistPreparationPipeline

'@section Stage enumeration
'===============================================================================
Public Enum LinelistPreparationStage
    stageCreateTemporarySheets = 1
    stageExportSpecifications = 2
    stageCreateAnalysisSheets = 3
    stageTransferCodeModules = 4
    stageTransferForms = 5
End Enum

'@section State and constants
'===============================================================================
Private Type TPipeline
    context As ILinelistPreparationContext
    steps As Collection                          'Registered steps keyed by stage
    orderedStages As BetterArray                 'Execution order
End Type

Private this As TPipeline
Private Const CLASSNAME As String = "LinelistPreparationPipeline"


'@section Factory helpers
'===============================================================================
Public Function Create(ByVal linelist As ILinelist, _
                       ByVal specs As ILinelistSpecs, _
                       ByVal temporarySheets As ILinelistPreparationStep, _
                       ByVal specExporter As ILinelistPreparationStep, _
                       ByVal analysisSheets As ILinelistPreparationStep, _
                       ByVal codeTransfer As ILinelistPreparationStep, _
                       ByVal formTransfer As ILinelistPreparationStep, _
                       Optional ByVal formatter As ILinelistSheetNameFormatter, _
                       Optional ByVal stateScope As IApplicationState) As ILinelistPreparationPipeline

    Dim pipeline As LinelistPreparationPipeline
    Set pipeline = New LinelistPreparationPipeline
    pipeline.Initialise linelist, specs, temporarySheets, specExporter, analysisSheets, codeTransfer, formTransfer, formatter, stateScope
    Set Create = pipeline
End Function

Friend Sub Initialise(ByVal linelist As ILinelist, _
                      ByVal specs As ILinelistSpecs, _
                      ByVal temporarySheets As ILinelistPreparationStep, _
                      ByVal specExporter As ILinelistPreparationStep, _
                      ByVal analysisSheets As ILinelistPreparationStep, _
                      ByVal codeTransfer As ILinelistPreparationStep, _
                      ByVal formTransfer As ILinelistPreparationStep, _
                      ByVal formatter As ILinelistSheetNameFormatter, _
                      ByVal stateScope As IApplicationState)

    Dim dictionary As ILLdictionary
    Dim actualFormatter As ILinelistSheetNameFormatter
    Dim actualScope As IApplicationState

    EnsureLinelist linelist
    Set dictionary = ResolveDictionary(specs)
    Set actualFormatter = ResolveFormatter(formatter)
    Set actualScope = ResolveScope(stateScope)

    Set this.context = LinelistPreparationContext.Create(linelist, specs, dictionary, actualFormatter, actualScope)
    Set this.steps = New Collection
    Set this.orderedStages = New BetterArray

    RegisterStage stageCreateTemporarySheets, temporarySheets
    RegisterStage stageExportSpecifications, specExporter
    RegisterStage stageCreateAnalysisSheets, analysisSheets
    RegisterStage stageTransferCodeModules, codeTransfer
    RegisterStage stageTransferForms, formTransfer
End Sub


'@section Validation helpers
'===============================================================================
Private Sub EnsureLinelist(ByVal linelist As ILinelist)
    If linelist Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist reference cannot be Nothing"
    End If
End Sub

Private Function ResolveDictionary(ByVal specs As ILinelistSpecs) As ILLdictionary
    Dim dictionary As ILLdictionary

    If specs Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Specifications reference cannot be Nothing"
    End If

    Set dictionary = specs.Dictionary
    If dictionary Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Specifications must expose a dictionary"
    End If

    Set ResolveDictionary = dictionary
End Function

Private Function ResolveFormatter(ByVal formatter As ILinelistSheetNameFormatter) As ILinelistSheetNameFormatter
    If formatter Is Nothing Then
        Set ResolveFormatter = LinelistSheetNameFormatter.Create
    Else
        Set ResolveFormatter = formatter
    End If
End Function

Private Function ResolveScope(ByVal scope As IApplicationState) As IApplicationState
    If scope Is Nothing Then
        Set ResolveScope = ApplicationState.Create(Application)
    Else
        Set ResolveScope = scope
    End If
End Function

Private Sub RegisterStage(ByVal stage As LinelistPreparationStage, ByVal step As ILinelistPreparationStep)
    If step Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Preparation step cannot be Nothing"
    End If

    If HasStage(stage) Then
        ThrowError ProjectError.ElementShouldNotExists, "Stage already registered"
    End If

    this.steps.Add step, CStr(stage)
    this.orderedStages.Push stage
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Stage execution helpers
'===============================================================================
Private Sub CreateTemporarySheets()
    ExecuteStage stageCreateTemporarySheets
End Sub

Private Sub ExportSpecifications()
    ExecuteStage stageExportSpecifications
End Sub

Private Sub CreateAnalysisSheets()
    ExecuteStage stageCreateAnalysisSheets
End Sub

Private Sub TransferCodeModules()
    ExecuteStage stageTransferCodeModules
End Sub

Private Sub TransferForms()
    ExecuteStage stageTransferForms
End Sub

Private Sub ExecuteStage(ByVal stage As LinelistPreparationStage)
    Dim stepObject As ILinelistPreparationStep

    Set stepObject = FetchStage(stage)
    stepObject.Execute this.context
End Sub


'@section Interface implementation
'===============================================================================
Private Sub ILinelistPreparationPipeline_Prepare()
    Dim scope As IApplicationState
    Dim stageIndex As Long
    Dim stageId As LinelistPreparationStage
    Dim stageName As String

    Set scope = this.context.StateScope
    scope.ApplyBusyState

    On Error GoTo PrepareFail
        For stageIndex = this.orderedStages.LowerBound To this.orderedStages.UpperBound
            stageId = CLng(this.orderedStages.Item(stageIndex))
            stageName = StageName(stageId)
            Select Case stageId
                Case stageCreateTemporarySheets: CreateTemporarySheets
                Case stageExportSpecifications: ExportSpecifications
                Case stageCreateAnalysisSheets: CreateAnalysisSheets
                Case stageTransferCodeModules: TransferCodeModules
                Case stageTransferForms: TransferForms
                Case Else
                    ThrowError ProjectError.InvalidArgument, _
                              "Unknown preparation stage encountered during execution"
            End Select
        Next stageIndex
    On Error GoTo 0

    scope.Restore
    Exit Sub

PrepareFail:
    scope.Restore
    ThrowError ProjectError.ErrorUnexpectedState, _
              "Preparation failed during stage " & stageName & ": " & Err.Description
End Sub

Private Function HasStage(ByVal stage As LinelistPreparationStage) As Boolean
    On Error GoTo Missing
        this.steps.Item CStr(stage)
        HasStage = True
        Exit Function
Missing:
    Err.Clear
    HasStage = False
End Function

Private Function FetchStage(ByVal stage As LinelistPreparationStage) As ILinelistPreparationStep
    On Error GoTo Missing
        Set FetchStage = this.steps.Item(CStr(stage))
        Exit Function
Missing:
    ThrowError ProjectError.ElementNotFound, "Preparation stage not registered"
End Function

Private Function StageName(ByVal stage As LinelistPreparationStage) As String
    Select Case stage
        Case stageCreateTemporarySheets: StageName = "CreateTemporarySheets"
        Case stageExportSpecifications: StageName = "ExportSpecifications"
        Case stageCreateAnalysisSheets: StageName = "CreateAnalysisSheets"
        Case stageTransferCodeModules: StageName = "TransferCodeModules"
        Case stageTransferForms: StageName = "TransferForms"
        Case Else: StageName = CStr(stage)
    End Select
End Function
