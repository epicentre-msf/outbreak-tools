VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHideStateBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Build show/hide states for a dictionary sheet and layer combination"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation
'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Build show/hide states for a dictionary sheet and layer combination")

Option Explicit
Implements IShowHideStateBuilder


'@section State
'===============================================================================

Private Type TShowHideStateBuilder
    dict As ILLdictionary
    vars As ILLVariables
    targetLayer As ShowHideWorksheetLayer
    sourceSheetName As String
    plan As IShowHidePlan
    hasPlan As Boolean
    hasVisibility As Boolean
    hasStatus As Boolean
    hasRegisterBook As Boolean
    hasCrfStatus As Boolean
    hasListAuto As Boolean
End Type

Private this As TShowHideStateBuilder
Private Const CLASSNAME As String = "ShowHideStateBuilder"
Private Const STATUS_COLUMN As String = "status"
Private Const VISIBILITY_COLUMN As String = "visibility"
Private Const REGISTER_BOOK_COLUMN As String = "register book"
Private Const CRF_STATUS_COLUMN As String = "crf status"
Private Const LIST_AUTO_COLUMN As String = "list auto"
Private Const MAIN_LABEL_COLUMN As String = "main label"
Private Const SHEET_TYPE_COLUMN As String = "sheet type"
Private Const SHEET_NAME_COLUMN As String = "sheet name"


'@section Factory
'===============================================================================

Public Function Create(ByVal dict As ILLdictionary, _
                       ByVal layer As ShowHideWorksheetLayer, _
                       ByVal sheetName As String, _
                       Optional ByVal vars As ILLVariables) As IShowHideStateBuilder

    With New ShowHideStateBuilder
        Set .Dictionary = dict
        .TargetLayer = layer
        .SourceSheetName = sheetName
        If Not (vars Is Nothing) Then Set .Variables = vars
        Set Create = .Self
    End With
End Function


'@section Public accessors
'===============================================================================

Public Property Get Self() As IShowHideStateBuilder
    Set Self = Me
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
    Set this.dict = dict
    InvalidateColumnCache
End Property

Public Property Set Variables(ByVal vars As ILLVariables)
    Set this.vars = vars
End Property

Public Property Get TargetLayer() As ShowHideWorksheetLayer
    TargetLayer = this.targetLayer
End Property

Public Property Let TargetLayer(ByVal layer As ShowHideWorksheetLayer)
    this.targetLayer = layer
End Property

Public Property Get SourceSheetName() As String
    SourceSheetName = this.sourceSheetName
End Property

Public Property Let SourceSheetName(ByVal sheetName As String)
    this.sourceSheetName = Trim$(sheetName)
End Property

Public Sub ApplyPlan(ByVal plan As IShowHidePlan)
    EnsurePlanValid plan
    Set this.plan = plan
    this.hasPlan = True
End Sub


'@section Execution
'===============================================================================

Public Function Build() As IShowHideState
    Dim state As IShowHideState
    Dim rule As IShowHideRule
    Dim varNames As BetterArray
    Dim idx As Long
    Dim varName As String
    Dim sheetType As String
    Dim hiddenDefault As Boolean
    Dim headerText As String
    Dim mandatory As Boolean
    Dim dictionaryHidden As Boolean
    Dim forceHidden As Boolean

    EnsureReady
    CacheColumnAvailability

    Set state = ShowHideState.Create()
    Set varNames = this.vars.VariableNames

    If Not (varNames Is Nothing) Then
        For idx = varNames.LowerBound To varNames.UpperBound
            varName = CStr(varNames.Item(idx))
            If LenB(varName) = 0 Then GoTo NextVariable
            If Not AppliesToSheet(varName) Then GoTo NextVariable

            sheetType = NormaliseText(ColumnValue(SHEET_TYPE_COLUMN, varName))
            If Not MatchesTargetLayer(sheetType) Then GoTo NextVariable

            headerText = DetermineHeader(varName)
            hiddenDefault = DetermineHiddenDefault(varName)
            mandatory = DetermineMandatory(varName)
            dictionaryHidden = DetermineDictionaryHidden(varName)
            forceHidden = DetermineForceHidden(varName)

            Set rule = ShowHideRule.Create(varName, headerText)
            rule.Mandatory = mandatory
            rule.DictionaryHidden = dictionaryHidden
            rule.ForceHidden = forceHidden
            rule.MarkLayerApplicability this.targetLayer, True
            rule.SetLayerVisibility this.targetLayer, hiddenDefault
            state.AddRule rule
NextVariable:
        Next idx
    End If

    ApplyPersistedPlan state
    Set Build = state
End Function


'@section Helpers
'===============================================================================

Private Sub EnsureReady()
    EnsureDictionaryAssigned
    EnsureLayerAssigned
    EnsureSheetNameProvided
    EnsureDictionaryPrepared
    EnsureVariablesReady
End Sub

Private Sub EnsureDictionaryAssigned()
    If this.dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "ShowHideStateBuilder requires a dictionary reference"
    End If
End Sub

Private Sub EnsureLayerAssigned()
    If this.targetLayer = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                  "ShowHideStateBuilder target layer is not configured"
    End If
End Sub

Private Sub EnsureSheetNameProvided()
    If LenB(Me.SourceSheetName) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                  "ShowHideStateBuilder requires a source sheet name"
    End If
End Sub

Private Sub EnsureDictionaryPrepared()
    If Not this.dict.Prepared Then
        this.dict.Prepare
    End If
End Sub

Private Sub EnsureVariablesReady()
    If this.vars Is Nothing Then
        Set this.vars = LLVariables.Create(this.dict)
    End If
End Sub

Private Sub CacheColumnAvailability()
    this.hasVisibility = this.dict.ColumnExists(VISIBILITY_COLUMN)
    this.hasStatus = this.dict.ColumnExists(STATUS_COLUMN)
    this.hasRegisterBook = this.dict.ColumnExists(REGISTER_BOOK_COLUMN)
    this.hasCrfStatus = this.dict.ColumnExists(CRF_STATUS_COLUMN)
    this.hasListAuto = this.dict.ColumnExists(LIST_AUTO_COLUMN)
End Sub

Private Sub InvalidateColumnCache()
    this.hasVisibility = False
    this.hasStatus = False
    this.hasRegisterBook = False
    this.hasCrfStatus = False
    this.hasListAuto = False
End Sub

Private Function AppliesToSheet(ByVal varName As String) As Boolean
    Dim varSheetName As String

    varSheetName = NormaliseText(this.vars.SheetName(varName, False))
    AppliesToSheet = (StrComp(varSheetName, NormaliseText(this.sourceSheetName), vbTextCompare) = 0)
End Function

Private Function MatchesTargetLayer(ByVal sheetType As String) As Boolean
    Dim lowered As String

    lowered = NormaliseText(sheetType)

    Select Case this.targetLayer
        Case ShowHideLayerHList, ShowHideLayerPrinted
            MatchesTargetLayer = (InStr(1, lowered, "hlist", vbBinaryCompare) > 0)
        Case ShowHideLayerVList
            MatchesTargetLayer = (InStr(1, lowered, "vlist", vbBinaryCompare) > 0)
        Case ShowHideLayerCRF
            MatchesTargetLayer = (InStr(1, lowered, "crf", vbBinaryCompare) > 0)
        Case ShowHideLayerDictionary
            MatchesTargetLayer = True
        Case Else
            MatchesTargetLayer = False
    End Select
End Function

Private Function DetermineHeader(ByVal varName As String) As String
    Dim headerValue As String

    headerValue = ColumnValue(MAIN_LABEL_COLUMN, varName)
    If LenB(headerValue) = 0 Then
        DetermineHeader = varName
    Else
        DetermineHeader = headerValue
    End If
End Function

Private Function DetermineHiddenDefault(ByVal varName As String) As Boolean
    Dim statusValue As String

    Select Case this.targetLayer
        Case ShowHideLayerPrinted
            statusValue = RegisterBookValue(varName)
        Case ShowHideLayerCRF
            statusValue = CrfStatusValue(varName)
        Case Else
            statusValue = VisibilityStatusValue(varName)
    End Select

    DetermineHiddenDefault = StatusIsHidden(statusValue, this.targetLayer)
End Function

Private Function DetermineMandatory(ByVal varName As String) As Boolean
    Dim statusValue As String
    Dim controlType As String

    statusValue = VisibilityStatusValue(varName)
    controlType = NormaliseText(this.vars.ControlType(varName, False))

    If StatusIsMandatory(statusValue) Then
        DetermineMandatory = True
    ElseIf (this.targetLayer = ShowHideLayerHList) And (controlType = "geo1") Then
        DetermineMandatory = True
    Else
        DetermineMandatory = False
    End If
End Function

Private Function DetermineDictionaryHidden(ByVal varName As String) As Boolean
    Dim statusValue As String

    statusValue = VisibilityStatusValue(varName)
    DetermineDictionaryHidden = (NormaliseText(statusValue) = "always hidden")
End Function

Private Function DetermineForceHidden(ByVal varName As String) As Boolean
    Dim controlType As String
    Dim listAutoValue As String
    Dim forceForLayer As Boolean

    forceForLayer = (this.targetLayer = ShowHideLayerCRF) Or (this.targetLayer = ShowHideLayerPrinted)

    If Not forceForLayer Then
        DetermineForceHidden = False
        Exit Function
    End If

    controlType = NormaliseText(this.vars.ControlType(varName, False))
    listAutoValue = NormaliseText(ColumnValue(LIST_AUTO_COLUMN, varName))

    If controlType = "formula" Then
        DetermineForceHidden = True
    ElseIf controlType = "choice_formula" Then
        DetermineForceHidden = True
    ElseIf controlType = "case_when" Then
        DetermineForceHidden = True
    ElseIf listAutoValue = "yes" Then
        DetermineForceHidden = True
    Else
        DetermineForceHidden = False
    End If
End Function

Private Sub ApplyPersistedPlan(ByVal state As IShowHideState)
    Dim actions As BetterArray
    Dim idx As Long
    Dim action As Variant
    Dim fieldKey As String
    Dim headerText As String
    Dim hiddenValue As Boolean
    Dim rule As IShowHideRule

    If Not this.hasPlan Then Exit Sub
    If this.plan Is Nothing Then Exit Sub

    Set actions = this.plan.Actions
    If actions Is Nothing Then Exit Sub
    If actions.Length = 0 Then Exit Sub

    For idx = actions.LowerBound To actions.UpperBound
        action = actions.Item(idx)
        fieldKey = CStr(action(0))
        headerText = CStr(action(1))
        hiddenValue = CBool(action(2))

        If Not state.HasRule(fieldKey) Then
            GoTo NextPlanAction
        End If

        Set rule = state.Rule(fieldKey)
        If LenB(headerText) > 0 Then rule.HeaderText = headerText

        If rule.Mandatory Then hiddenValue = False
        rule.SetLayerVisibility this.targetLayer, hiddenValue
NextPlanAction:
    Next idx
End Sub

Private Function VisibilityStatusValue(ByVal varName As String) As String
    Dim statusValue As String

    If this.hasVisibility Then
        statusValue = ColumnValue(VISIBILITY_COLUMN, varName)
    ElseIf this.hasStatus Then
        statusValue = ColumnValue(STATUS_COLUMN, varName)
    Else
        statusValue = vbNullString
    End If

    VisibilityStatusValue = statusValue
End Function

Private Function RegisterBookValue(ByVal varName As String) As String
    If this.hasRegisterBook Then
        RegisterBookValue = ColumnValue(REGISTER_BOOK_COLUMN, varName)
    Else
        RegisterBookValue = vbNullString
    End If
End Function

Private Function CrfStatusValue(ByVal varName As String) As String
    If this.hasCrfStatus Then
        CrfStatusValue = ColumnValue(CRF_STATUS_COLUMN, varName)
    Else
        CrfStatusValue = vbNullString
    End If
End Function

Private Function ColumnValue(ByVal colName As String, ByVal varName As String) As String
    Dim value As String

    On Error Resume Next
        value = this.vars.Value(colName, varName)
    On Error GoTo 0

    ColumnValue = value
End Function

Private Function StatusIsHidden(ByVal statusValue As String, _
                                ByVal layer As ShowHideWorksheetLayer) As Boolean
    Dim lowered As String

    lowered = NormaliseText(statusValue)

    Select Case lowered
        Case "hidden", "hidden on creation", "always hidden", "optional, hidden", "do not print"
            StatusIsHidden = True
        Case Else
            If layer = ShowHideLayerPrinted Then
                StatusIsHidden = (InStr(1, lowered, "hidden", vbBinaryCompare) > 0)
            Else
                StatusIsHidden = False
            End If
    End Select
End Function

Private Function StatusIsMandatory(ByVal statusValue As String) As Boolean
    Dim lowered As String

    lowered = NormaliseText(statusValue)

    Select Case lowered
        Case "mandatory", "always visible"
            StatusIsMandatory = True
        Case Else
            StatusIsMandatory = False
    End Select
End Function

Private Function NormaliseText(ByVal value As String) As String
    NormaliseText = LCase$(Trim$(value))
End Function

Private Sub EnsurePlanValid(ByVal plan As IShowHidePlan)
    If plan Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Cannot apply an empty persistence plan"
    End If

    If plan.TargetLayer <> this.targetLayer Then
        ThrowError ProjectError.InvalidArgument, _
                  "Plan target layer does not match the builder configuration"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface
'===============================================================================

Private Sub IShowHideStateBuilder_ApplyPlan(ByVal plan As IShowHidePlan)
    ApplyPlan plan
End Sub

Private Function IShowHideStateBuilder_Build() As IShowHideState
    Set IShowHideStateBuilder_Build = Build()
End Function

Private Property Set IShowHideStateBuilder_Dictionary(ByVal dict As ILLdictionary)
    Set Dictionary = dict
End Property

Private Property Set IShowHideStateBuilder_Variables(ByVal vars As ILLVariables)
    Set Variables = vars
End Property

Private Property Get IShowHideStateBuilder_TargetLayer() As ShowHideWorksheetLayer
    IShowHideStateBuilder_TargetLayer = TargetLayer
End Property

Private Property Let IShowHideStateBuilder_TargetLayer(ByVal layer As ShowHideWorksheetLayer)
    TargetLayer = layer
End Property

Private Property Get IShowHideStateBuilder_SourceSheetName() As String
    IShowHideStateBuilder_SourceSheetName = SourceSheetName
End Property

Private Property Let IShowHideStateBuilder_SourceSheetName(ByVal sheetName As String)
    SourceSheetName = sheetName
End Property
