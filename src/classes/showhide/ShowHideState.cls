VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHideState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Aggregate show/hide state storage leveraging IShowHideRule entries"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation
'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Aggregate show/hide state storage leveraging IShowHideRule entries")

Option Explicit
Implements IShowHideState


'@section State
'===============================================================================

Private Type TShowHideState
    rules As Collection
    order As BetterArray
End Type

Private this As TShowHideState
Private Const CLASSNAME As String = "ShowHideState"


'@section Factory
'===============================================================================

'@label:Create
'@sub-title Instantiate an empty state container
'@return IShowHideState ready to accept rules.
Public Function Create() As IShowHideState
    With New ShowHideState
        .Initialise
        Set Create = .Self
    End With
End Function


'@section Public accessors
'===============================================================================

'@label:Self
'@sub-title Return the current instance as interface reference
Public Property Get Self() As IShowHideState
    Set Self = Me
End Property

'@label:AddRule
'@sub-title Register a new rule in the state container
'@param rule IShowHideRule to add.
Public Sub AddRule(ByVal rule As IShowHideRule)
    Dim key As String

    EnsureStateReady

    If rule Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Cannot add an empty show/hide rule"
    End If

    key = ValidatedFieldKey(rule.FieldKey)

    If HasRule(key) Then
        ThrowError ProjectError.ElementShouldNotExists, _
                  "A show/hide rule already exists for field '" & key & "'"
    End If

    this.rules.Add rule, key
    this.order.Push key
End Sub

'@label:HasRule
'@sub-title Determine whether a rule is present for the supplied key
Public Function HasRule(ByVal fieldKey As String) As Boolean
    Dim rule As IShowHideRule

    EnsureStateReady

    Set rule = Nothing
    Set rule = InternalRule(fieldKey, False)
    HasRule = Not (rule Is Nothing)
End Function

'@label:Rule
'@sub-title Retrieve the rule for the supplied key
'@param fieldKey Identifier of the column.
'@return IShowHideRule reference; raises when absent.
Public Function Rule(ByVal fieldKey As String) As IShowHideRule
    Dim result As IShowHideRule

    Set result = InternalRule(fieldKey, True)
    Set Rule = result
End Function

'@label:Rules
'@sub-title Return all registered rules preserving insertion order
'@return BetterArray containing IShowHideRule entries.
Public Function Rules() As BetterArray
    Dim keysArr As BetterArray
    Dim idx As Long
    Dim key As String
    Dim rule As IShowHideRule
    Dim copyArr As BetterArray

    EnsureStateReady

    Set keysArr = this.order

    Set copyArr = New BetterArray
    copyArr.LowerBound = 1

    If Not (keysArr Is Nothing) Then
        For idx = keysArr.LowerBound To keysArr.UpperBound
            key = CStr(keysArr.Item(idx))
            Set rule = InternalRule(key, True)
            copyArr.Push rule
        Next idx
    End If

    Set Rules = copyArr
End Function

'@label:UpdateLayerVisibility
'@sub-title Persist visibility choice for a given field and layer
Public Sub UpdateLayerVisibility(ByVal fieldKey As String, _
                                 ByVal layer As ShowHideWorksheetLayer, _
                                 ByVal hidden As Boolean)
    Dim rule As IShowHideRule

    Set rule = InternalRule(fieldKey, True)
    rule.SetLayerVisibility layer, hidden
End Sub

'@label:MarkLayerAsApplicable
'@sub-title Flag the rule as relevant for the provided layer
Public Sub MarkLayerAsApplicable(ByVal fieldKey As String, _
                                 ByVal layer As ShowHideWorksheetLayer, _
                                 ByVal applies As Boolean)
    Dim rule As IShowHideRule

    Set rule = InternalRule(fieldKey, True)
    rule.MarkLayerApplicability layer, applies
End Sub

'@label:SetDictionaryHidden
'@sub-title Override dictionary hidden flag for a given rule
Public Sub SetDictionaryHidden(ByVal fieldKey As String, ByVal hidden As Boolean)
    Dim rule As IShowHideRule

    Set rule = InternalRule(fieldKey, True)
    rule.DictionaryHidden = hidden
End Sub

'@label:SetForceHidden
'@sub-title Override forced hidden flag for a given rule
Public Sub SetForceHidden(ByVal fieldKey As String, ByVal hidden As Boolean)
    Dim rule As IShowHideRule

    Set rule = InternalRule(fieldKey, True)
    rule.ForceHidden = hidden
End Sub

'@label:EvaluateHidden
'@sub-title Compute final hidden status for a field/layer combination
Public Function EvaluateHidden(ByVal fieldKey As String, _
                               ByVal layer As ShowHideWorksheetLayer) As Boolean
    Dim rule As IShowHideRule

    Set rule = InternalRule(fieldKey, True)
    EvaluateHidden = rule.EffectiveHidden(layer)
End Function


'@section Helpers
'===============================================================================

'@label:Initialise
'@sub-title Prepare internal collections
Public Sub Initialise()
    Set this.rules = New Collection
    Set this.order = New BetterArray
    this.order.LowerBound = 1
End Sub

'@label:EnsureStateReady
'@sub-title Lazily initialise internal state
Private Sub EnsureStateReady()
    If this.rules Is Nothing Then Initialise
End Sub

'@label:InternalRule
'@sub-title Lookup a rule and optionally raise when missing
'@param fieldKey Column identifier.
'@param failWhenMissing Raise ProjectError.ElementNotFound when True and rule absent.
'@return IShowHideRule or Nothing.
Private Function InternalRule(ByVal fieldKey As String, ByVal failWhenMissing As Boolean) As IShowHideRule
    Dim key As String
    Dim candidate As IShowHideRule

    key = ValidatedFieldKey(fieldKey)
    Set candidate = Nothing

    Err.Clear
    On Error Resume Next
        Set candidate = this.rules.Item(key)
    On Error GoTo 0
    Err.Clear

    If candidate Is Nothing Then
        If failWhenMissing Then
            ThrowError ProjectError.ElementNotFound, _
                      "No show/hide rule registered for field '" & key & "'"
        End If
    End If

    Set InternalRule = candidate
End Function

'@label:ValidatedFieldKey
'@sub-title Trim and validate a field key input
'@return Sanitised key suitable for use as collection index.
Private Function ValidatedFieldKey(ByVal fieldKey As String) As String
    Dim trimmed As String

    trimmed = Trim$(fieldKey)
    If LenB(trimmed) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Field key cannot be blank"
    End If

    ValidatedFieldKey = trimmed
End Function

'@label:ThrowError
'@sub-title Raise a project specific error
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface
'===============================================================================

Private Sub IShowHideState_AddRule(ByVal rule As IShowHideRule)
    AddRule rule
End Sub

Private Function IShowHideState_HasRule(ByVal fieldKey As String) As Boolean
    IShowHideState_HasRule = HasRule(fieldKey)
End Function

Private Function IShowHideState_Rule(ByVal fieldKey As String) As IShowHideRule
    Set IShowHideState_Rule = Rule(fieldKey)
End Function

Private Function IShowHideState_Rules() As BetterArray
    Set IShowHideState_Rules = Rules()
End Function

Private Sub IShowHideState_UpdateLayerVisibility(ByVal fieldKey As String, _
                                                 ByVal layer As ShowHideWorksheetLayer, _
                                                 ByVal hidden As Boolean)
    UpdateLayerVisibility fieldKey, layer, hidden
End Sub

Private Sub IShowHideState_MarkLayerAsApplicable(ByVal fieldKey As String, _
                                                 ByVal layer As ShowHideWorksheetLayer, _
                                                 ByVal applies As Boolean)
    MarkLayerAsApplicable fieldKey, layer, applies
End Sub

Private Sub IShowHideState_SetDictionaryHidden(ByVal fieldKey As String, ByVal hidden As Boolean)
    SetDictionaryHidden fieldKey, hidden
End Sub

Private Sub IShowHideState_SetForceHidden(ByVal fieldKey As String, ByVal hidden As Boolean)
    SetForceHidden fieldKey, hidden
End Sub

Private Function IShowHideState_EvaluateHidden(ByVal fieldKey As String, _
                                               ByVal layer As ShowHideWorksheetLayer) As Boolean
    IShowHideState_EvaluateHidden = EvaluateHidden(fieldKey, layer)
End Function
