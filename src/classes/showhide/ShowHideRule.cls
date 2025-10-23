VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHideRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Concrete implementation of IShowHideRule storing per-layer visibility flags"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, HungarianNotation, ProcedureNotUsed
'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Concrete implementation of IShowHideRule storing per-layer visibility flags")

Option Explicit
Implements IShowHideRule

'@section State
'===============================================================================

Private Type TShowHideRule
    fieldKey As String
    headerText As String
    mandatory As Boolean
    dictionaryHidden As Boolean
    forceHidden As Boolean
    appliesToLayer(ShowHideLayerHList To ShowHideLayerDictionary) As Boolean
    userHidden(ShowHideLayerHList To ShowHideLayerDictionary) As Boolean
    hasUserChoice(ShowHideLayerHList To ShowHideLayerDictionary) As Boolean
End Type

Private this As TShowHideRule
Private Const CLASSNAME As String = "ShowHideRule"


'@section Factory
'===============================================================================

'@label:Create
'@sub-title Build a new rule instance tied to a column identifier
'@param fieldKey Column identifier (dictionary name or internal id).
'@param header Column header displayed in worksheets.
'@return IShowHideRule ready for consumption.
Public Function Create(ByVal fieldKey As String, ByVal header As String) As IShowHideRule
    Dim trimmedKey As String
    Dim trimmedHeader As String

    trimmedKey = Trim$(fieldKey)
    trimmedHeader = Trim$(header)

    If LenB(trimmedKey) = 0 Then
        ThrowError ProjectError.InvalidArgument, "ShowHideRule requires a non-empty field key"
    End If

    If LenB(trimmedHeader) = 0 Then
        ThrowError ProjectError.InvalidArgument, "ShowHideRule requires a non-empty header text"
    End If

    With New ShowHideRule
        .Initialise trimmedKey, trimmedHeader
        Set Create = .Self
    End With
End Function


'@section Public accessors
'===============================================================================

'@label:Self
'@sub-title Return the current instance as an interface reference
Public Property Get Self() As IShowHideRule
    Set Self = Me
End Property

'@label:FieldKey
'@sub-title Retrieve the unique identifier associated with the column
Public Property Get FieldKey() As String
    FieldKey = this.fieldKey
End Property

'@sub-title Assign the unique identifier associated with the column
Public Property Let FieldKey(ByVal value As String)
    Dim trimmedValue As String

    trimmedValue = Trim$(value)
    If LenB(trimmedValue) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Field key cannot be blank"
    End If
    this.fieldKey = trimmedValue
End Property

'@label:HeaderText
'@sub-title Return the worksheet header associated with the column
Public Property Get HeaderText() As String
    HeaderText = this.headerText
End Property

'@sub-title Assign the worksheet header associated with the column
Public Property Let HeaderText(ByVal value As String)
    Dim trimmedValue As String

    trimmedValue = Trim$(value)
    If LenB(trimmedValue) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Header text cannot be blank"
    End If
    this.headerText = trimmedValue
End Property

'@label:Mandatory
'@sub-title Retrieve the mandatory flag for the column
Public Property Get Mandatory() As Boolean
    Mandatory = this.mandatory
End Property

'@sub-title Assign the mandatory flag for the column
Public Property Let Mandatory(ByVal value As Boolean)
    this.mandatory = value
End Property

'@label:DictionaryHidden
'@sub-title True when the dictionary definition hides the column globally
Public Property Get DictionaryHidden() As Boolean
    DictionaryHidden = this.dictionaryHidden
End Property

'@sub-title Update the dictionary hidden status
Public Property Let DictionaryHidden(ByVal value As Boolean)
    this.dictionaryHidden = value
End Property

'@label:ForceHidden
'@sub-title True when the column must be hidden regardless of other settings
Public Property Get ForceHidden() As Boolean
    ForceHidden = this.forceHidden
End Property

'@sub-title Enforce forced hidden flag
Public Property Let ForceHidden(ByVal value As Boolean)
    this.forceHidden = value
End Property


'@section Visibility management
'===============================================================================

'@label:SetLayerVisibility
'@sub-title Store the explicit visibility chosen for a given layer
'@param layer Target worksheet layer.
'@param hidden True when the column should be hidden on that layer.
Public Sub SetLayerVisibility(ByVal layer As ShowHideWorksheetLayer, ByVal hidden As Boolean)
    EnsureValidLayer layer
    this.userHidden(layer) = hidden
    this.hasUserChoice(layer) = True
    this.appliesToLayer(layer) = True
End Sub

'@label:LayerVisibility
'@sub-title Retrieve the stored visibility for the provided layer
'@param layer Target worksheet layer.
'@return True when hidden; False otherwise (defaults to False when unset).
Public Property Get LayerVisibility(ByVal layer As ShowHideWorksheetLayer) As Boolean
    EnsureValidLayer layer
    LayerVisibility = this.userHidden(layer)
End Property

'@label:HasLayerVisibility
'@sub-title Indicate whether a user choice exists for the layer
Public Function HasLayerVisibility(ByVal layer As ShowHideWorksheetLayer) As Boolean
    EnsureValidLayer layer
    HasLayerVisibility = this.hasUserChoice(layer)
End Function

'@label:AppliesToLayer
'@sub-title Check whether the rule should be considered for the provided layer
Public Function AppliesToLayer(ByVal layer As ShowHideWorksheetLayer) As Boolean
    EnsureValidLayer layer
    AppliesToLayer = this.appliesToLayer(layer)
End Function

'@label:MarkLayerApplicability
'@sub-title Explicitly mark a layer as receiving show/hide evaluation without setting user choice
Public Sub MarkLayerApplicability(ByVal layer As ShowHideWorksheetLayer, ByVal applies As Boolean)
    EnsureValidLayer layer
    this.appliesToLayer(layer) = applies
End Sub

'@label:EffectiveHidden
'@sub-title Compute final hidden state after applying forced and dictionary rules
'@param layer Target worksheet layer.
'@return Boolean True when the column must be hidden for that layer.
Public Function EffectiveHidden(ByVal layer As ShowHideWorksheetLayer) As Boolean
    Dim finalState As Boolean

    EnsureValidLayer layer

    If this.dictionaryHidden Then
        finalState = True
    ElseIf this.forceHidden And (layer = ShowHideLayerCRF Or layer = ShowHideLayerPrinted) Then
        finalState = True
    ElseIf this.hasUserChoice(layer) Then
        finalState = this.userHidden(layer)
    Else
        finalState = False
    End If

    EffectiveHidden = finalState
End Function


'@section Helpers
'===============================================================================

'@label:Initialise
'@sub-title Prime internal state once Create validates the arguments
Public Sub Initialise(ByVal fieldKey As String, ByVal header As String)
    Dim layer As ShowHideWorksheetLayer

    this.fieldKey = fieldKey
    this.headerText = header

    For layer = ShowHideLayerHList To ShowHideLayerDictionary
        this.appliesToLayer(layer) = False
        this.userHidden(layer) = False
        this.hasUserChoice(layer) = False
    Next layer
End Sub

'@label:EnsureValidLayer
'@sub-title Guard against invalid layer identifiers
'@param layer Candidate layer value.
Private Sub EnsureValidLayer(ByVal layer As ShowHideWorksheetLayer)
    If layer < ShowHideLayerHList Or layer > ShowHideLayerDictionary Then
        ThrowError ProjectError.InvalidArgument, _
                  "ShowHideRule layer is outside expected range"
    End If
End Sub

'@label:ThrowError
'@sub-title Raise a project-specific error
'@param errNumber ProjectError code describing the failure.
'@param message Human readable explanation.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface
'===============================================================================

Private Property Get IShowHideRule_FieldKey() As String
    IShowHideRule_FieldKey = FieldKey
End Property

Private Property Let IShowHideRule_FieldKey(ByVal value As String)
    FieldKey = value
End Property

Private Property Get IShowHideRule_HeaderText() As String
    IShowHideRule_HeaderText = HeaderText
End Property

Private Property Let IShowHideRule_HeaderText(ByVal value As String)
    HeaderText = value
End Property

Private Property Get IShowHideRule_Mandatory() As Boolean
    IShowHideRule_Mandatory = Mandatory
End Property

Private Property Let IShowHideRule_Mandatory(ByVal value As Boolean)
    Mandatory = value
End Property

Private Property Get IShowHideRule_DictionaryHidden() As Boolean
    IShowHideRule_DictionaryHidden = DictionaryHidden
End Property

Private Property Let IShowHideRule_DictionaryHidden(ByVal value As Boolean)
    DictionaryHidden = value
End Property

Private Property Get IShowHideRule_ForceHidden() As Boolean
    IShowHideRule_ForceHidden = ForceHidden
End Property

Private Property Let IShowHideRule_ForceHidden(ByVal value As Boolean)
    ForceHidden = value
End Property

Private Sub IShowHideRule_SetLayerVisibility(ByVal layer As ShowHideWorksheetLayer, ByVal hidden As Boolean)
    SetLayerVisibility layer, hidden
End Sub

Private Property Get IShowHideRule_LayerVisibility(ByVal layer As ShowHideWorksheetLayer) As Boolean
    IShowHideRule_LayerVisibility = LayerVisibility(layer)
End Property

Private Function IShowHideRule_HasLayerVisibility(ByVal layer As ShowHideWorksheetLayer) As Boolean
    IShowHideRule_HasLayerVisibility = HasLayerVisibility(layer)
End Function

Private Sub IShowHideRule_MarkLayerApplicability(ByVal layer As ShowHideWorksheetLayer, ByVal applies As Boolean)
    MarkLayerApplicability layer, applies
End Sub

Private Function IShowHideRule_EffectiveHidden(ByVal layer As ShowHideWorksheetLayer) As Boolean
    IShowHideRule_EffectiveHidden = EffectiveHidden(layer)
End Function

Private Function IShowHideRule_AppliesToLayer(ByVal layer As ShowHideWorksheetLayer) As Boolean
    IShowHideRule_AppliesToLayer = AppliesToLayer(layer)
End Function
