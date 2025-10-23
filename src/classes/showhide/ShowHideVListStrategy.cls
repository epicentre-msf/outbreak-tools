VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHideVListStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Strategy building visibility plans for VList worksheets"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation
'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Strategy building visibility plans for VList worksheets")

Option Explicit
Implements IShowHideStrategy

Private Const CLASSNAME As String = "ShowHideVListStrategy"
Private Const STRATEGY_NAME As String = "VList"


'@section Factory
'===============================================================================

Public Function Create() As IShowHideStrategy
    With New ShowHideVListStrategy
        Set Create = .Self
    End With
End Function


'@section Public accessors
'===============================================================================

Public Property Get Self() As IShowHideStrategy
    Set Self = Me
End Property

Public Property Get TargetLayer() As ShowHideWorksheetLayer
    TargetLayer = ShowHideLayerVList
End Property

Public Property Get StrategyName() As String
    StrategyName = STRATEGY_NAME
End Property

Public Function BuildPlan(ByVal state As IShowHideState, _
                          ByVal table As ICustomTable) As IShowHidePlan
    ValidateDependencies state
    Set BuildPlan = BuildPlanInternal(state)
End Function


'@section Helpers
'===============================================================================

Private Function BuildPlanInternal(ByVal state As IShowHideState) As IShowHidePlan
    Dim rules As BetterArray
    Dim plan As IShowHidePlan
    Dim idx As Long
    Dim rule As IShowHideRule

    Set plan = ShowHidePlan.Create(TargetLayer)
    Set rules = state.Rules

    If rules Is Nothing Then
        Set BuildPlanInternal = plan
        Exit Function
    End If

    If rules.Length = 0 Then
        Set BuildPlanInternal = plan
        Exit Function
    End If

    For idx = rules.LowerBound To rules.UpperBound
        Set rule = rules.Item(idx)
        If ShouldIncludeRule(rule) Then
            plan.AddVisibility rule.FieldKey, rule.HeaderText, _
                                rule.EffectiveHidden(TargetLayer)
        End If
    Next idx

    Set BuildPlanInternal = plan
End Function

Private Function ShouldIncludeRule(ByVal rule As IShowHideRule) As Boolean
    Dim include As Boolean
    include = True

    If Not rule.AppliesToLayer(TargetLayer) Then
        include = rule.HasLayerVisibility(TargetLayer) _
                  Or rule.DictionaryHidden _
                  Or rule.Mandatory
    End If

    ShouldIncludeRule = include
End Function

Private Sub ValidateDependencies(ByVal state As IShowHideState)
    If state Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "ShowHideVListStrategy requires a state provider"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface
'===============================================================================

Private Property Get IShowHideStrategy_TargetLayer() As ShowHideWorksheetLayer
    IShowHideStrategy_TargetLayer = TargetLayer
End Property

Private Property Get IShowHideStrategy_StrategyName() As String
    IShowHideStrategy_StrategyName = StrategyName
End Property

Private Function IShowHideStrategy_BuildPlan(ByVal state As IShowHideState, _
                                             ByVal table As ICustomTable) As IShowHidePlan
    Set IShowHideStrategy_BuildPlan = BuildPlan(state, table)
End Function
