VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHideOrchestrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Coordinate show/hide state creation, strategy evaluation, and plan persistence"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation
'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Coordinate show/hide state creation, strategy evaluation, and plan persistence")

Option Explicit
Implements IShowHideOrchestrator


'@section State
'===============================================================================

Private Type TShowHideOrchestrator
    dict As ILLdictionary
    strategies As Collection
End Type

Private this As TShowHideOrchestrator
Private Const CLASSNAME As String = "ShowHideOrchestrator"


'@section Factory
'===============================================================================

Public Function Create(ByVal dict As ILLdictionary) As IShowHideOrchestrator
    With New ShowHideOrchestrator
        Set .Dictionary = dict
        Set Create = .Self
    End With
End Function


'@section Public accessors
'===============================================================================

Public Property Get Self() As IShowHideOrchestrator
    Set Self = Me
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
    Set this.dict = dict
End Property

Public Sub RegisterStrategy(ByVal strategy As IShowHideStrategy)
    Dim key As String

    EnsureStrategies

    If strategy Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Cannot register an empty strategy"
    End If

    key = StrategyKey(strategy.TargetLayer)

    On Error Resume Next
        this.strategies.Remove key
    On Error GoTo 0

    this.strategies.Add strategy, key
End Sub

Public Function HasStrategy(ByVal layer As ShowHideWorksheetLayer) As Boolean
    Dim strategy As IShowHideStrategy
    EnsureStrategies
    Set strategy = ResolveStrategy(layer, False)
    HasStrategy = Not (strategy Is Nothing)
End Function

Public Function Apply(ByVal sheetName As String, _
                      ByVal layer As ShowHideWorksheetLayer, _
                      ByVal exportService As IShowHideExport, _
                      Optional ByVal table As ICustomTable) As IShowHidePlan

    Dim builder As IShowHideStateBuilder
    Dim state As IShowHideState
    Dim strategy As IShowHideStrategy
    Dim persistedPlan As IShowHidePlan
    Dim finalPlan As IShowHidePlan

    EnsureReady sheetName, layer, exportService

    Set builder = ShowHideStateBuilder.Create(this.dict, layer, sheetName)

    Set persistedPlan = exportService.ImportPlan(layer)
    If Not (persistedPlan Is Nothing) Then
        builder.ApplyPlan persistedPlan
    End If

    Set state = builder.Build
    Set strategy = ResolveStrategy(layer, True)
    Set finalPlan = strategy.BuildPlan(state, table)

    exportService.ExportPlan finalPlan

    Set Apply = finalPlan
End Function


'@section Helpers
'===============================================================================

Private Sub EnsureStrategies()
    If this.strategies Is Nothing Then
        Set this.strategies = New Collection
    End If
End Sub

Private Sub EnsureReady(ByVal sheetName As String, _
                        ByVal layer As ShowHideWorksheetLayer, _
                        ByVal exportService As IShowHideExport)

    If this.dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "ShowHideOrchestrator requires a dictionary reference"
    End If

    If LenB(Trim$(sheetName)) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                  "Sheet name cannot be blank when applying show/hide coordination"
    End If

    If layer = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                  "Layer value is not configured"
    End If

    If exportService Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Export service reference is required"
    End If

    If exportService.BoundTable Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Export service must be bound to a table before orchestration"
    End If
End Sub

Private Function ResolveStrategy(ByVal layer As ShowHideWorksheetLayer, _
                                 ByVal failWhenMissing As Boolean) As IShowHideStrategy
    Dim key As String
    Dim strategy As IShowHideStrategy

    EnsureStrategies
    key = StrategyKey(layer)

    On Error Resume Next
        Set strategy = this.strategies.Item(key)
    On Error GoTo 0

    If strategy Is Nothing Then
        If failWhenMissing Then
            ThrowError ProjectError.ElementNotFound, _
                      "No show/hide strategy registered for layer value " & CStr(layer)
        End If
    End If

    Set ResolveStrategy = strategy
End Function

Private Function StrategyKey(ByVal layer As ShowHideWorksheetLayer) As String
    StrategyKey = CStr(layer)
End Function

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface
'===============================================================================

Private Sub IShowHideOrchestrator_RegisterStrategy(ByVal strategy As IShowHideStrategy)
    RegisterStrategy strategy
End Sub

Private Function IShowHideOrchestrator_HasStrategy(ByVal layer As ShowHideWorksheetLayer) As Boolean
    IShowHideOrchestrator_HasStrategy = HasStrategy(layer)
End Function

Private Function IShowHideOrchestrator_Apply(ByVal sheetName As String, _
                                             ByVal layer As ShowHideWorksheetLayer, _
                                             ByVal exportService As IShowHideExport, _
                                             Optional ByVal table As ICustomTable) As IShowHidePlan
    Set IShowHideOrchestrator_Apply = Apply(sheetName, layer, exportService, table)
End Function
