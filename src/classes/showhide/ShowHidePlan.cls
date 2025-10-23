VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHidePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Collection of column visibility actions for a specific worksheet layer"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation
'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Collection of column visibility actions for a specific worksheet layer")

Option Explicit
Implements IShowHidePlan


'@section State
'===============================================================================

Private Type TShowHidePlan
    layer As ShowHideWorksheetLayer
    actions As Collection
    order As BetterArray
End Type

Private this As TShowHidePlan
Private Const CLASSNAME As String = "ShowHidePlan"
Private Const ACTION_FIELD As Long = 0
Private Const ACTION_HEADER As Long = 1
Private Const ACTION_HIDDEN As Long = 2


'@section Factory
'===============================================================================

'@label:Create
'@sub-title Instantiate an empty plan bound to a target layer
'@param layer Worksheet layer that will receive the actions.
'@return IShowHidePlan containing no actions.
Public Function Create(ByVal layer As ShowHideWorksheetLayer) As IShowHidePlan
    With New ShowHidePlan
        .Initialise layer
        Set Create = .Self
    End With
End Function


'@section Public accessors
'===============================================================================

'@label:Self
Public Property Get Self() As IShowHidePlan
    Set Self = Me
End Property

'@label:TargetLayer
Public Property Get TargetLayer() As ShowHideWorksheetLayer
    TargetLayer = this.layer
End Property

'@label:AddVisibility
'@sub-title Record a visibility change for a column
Public Sub AddVisibility(ByVal fieldKey As String, _
                         ByVal headerText As String, _
                         ByVal hidden As Boolean)
    Dim key As String
    Dim entry As Variant

    EnsureStateReady

    key = ValidatedFieldKey(fieldKey)

    If ContainsField(key) Then
        ThrowError ProjectError.ElementShouldNotExists, _
                  "Visibility already specified for field '" & key & "'"
    End If

    entry = BuildEntry(key, headerText, hidden)
    this.actions.Add entry, key
    this.order.Push key
End Sub

'@label:Actions
'@sub-title Retrieve recorded actions preserving insertion order
Public Function Actions() As BetterArray
    Dim output As BetterArray
    Dim idx As Long
    Dim key As String
    Dim entry As Variant

    EnsureStateReady

    Set output = New BetterArray
    output.LowerBound = 1

    For idx = this.order.LowerBound To this.order.UpperBound
        key = CStr(this.order.Item(idx))
        entry = this.actions.Item(key)
        output.Push entry
    Next idx

    Set Actions = output
End Function

'@label:ContainsField
'@sub-title Check if an action already exists for a column
Public Function ContainsField(ByVal fieldKey As String) As Boolean
    Dim key As String
    Dim exists As Boolean

    EnsureStateReady
    key = ValidatedFieldKey(fieldKey)

    exists = False
    Err.Clear
    On Error Resume Next
        Call this.actions.Item(key)
        exists = (Err.Number = 0)
    On Error GoTo 0
    Err.Clear

    ContainsField = exists
End Function

'@label:VisibilityFor
'@sub-title Return the hidden flag recorded for a column
Public Function VisibilityFor(ByVal fieldKey As String) As Boolean
    Dim key As String
    Dim entry As Variant

    EnsureStateReady
    key = ValidatedFieldKey(fieldKey)

    Err.Clear
    On Error Resume Next
        entry = this.actions.Item(key)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Clear
            ThrowError ProjectError.ElementNotFound, _
                      "No visibility entry exists for field '" & key & "'"
        End If
    On Error GoTo 0
    Err.Clear

    VisibilityFor = CBool(entry(ACTION_HIDDEN))
End Function


'@section Helpers
'===============================================================================

'@label:Initialise
'@sub-title Prepare the plan for usage
Public Sub Initialise(ByVal layer As ShowHideWorksheetLayer)
    EnsureValidLayer layer
    this.layer = layer
    Set this.actions = New Collection
    Set this.order = New BetterArray
    this.order.LowerBound = 1
End Sub

'@label:EnsureStateReady
Private Sub EnsureStateReady()
    If this.actions Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "ShowHidePlan is not initialised. Use Create or Initialise before recording actions"
    End If
End Sub

'@label:BuildEntry
Private Function BuildEntry(ByVal fieldKey As String, _
                            ByVal headerText As String, _
                            ByVal hidden As Boolean) As Variant
    Dim trimmedHeader As String
    Dim entry As Variant

    trimmedHeader = Trim$(headerText)
    If LenB(trimmedHeader) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                  "Header text cannot be blank when registering visibility"
    End If

    entry = Array(fieldKey, trimmedHeader, hidden)
    BuildEntry = entry
End Function

'@label:ValidatedFieldKey
Private Function ValidatedFieldKey(ByVal fieldKey As String) As String
    Dim trimmed As String

    trimmed = Trim$(fieldKey)
    If LenB(trimmed) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Field key cannot be blank"
    End If

    ValidatedFieldKey = trimmed
End Function

'@label:EnsureValidLayer
Private Sub EnsureValidLayer(ByVal layer As ShowHideWorksheetLayer)
    If layer < ShowHideLayerHList Or layer > ShowHideLayerDictionary Then
        ThrowError ProjectError.InvalidArgument, _
                  "Target layer is outside the expected range"
    End If
End Sub

'@label:ThrowError
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface
'===============================================================================

Private Property Get IShowHidePlan_TargetLayer() As ShowHideWorksheetLayer
    IShowHidePlan_TargetLayer = TargetLayer
End Property

Private Sub IShowHidePlan_AddVisibility(ByVal fieldKey As String, _
                                        ByVal headerText As String, _
                                        ByVal hidden As Boolean)
    AddVisibility fieldKey, headerText, hidden
End Sub

Private Function IShowHidePlan_Actions() As BetterArray
    Set IShowHidePlan_Actions = Actions()
End Function

Private Function IShowHidePlan_ContainsField(ByVal fieldKey As String) As Boolean
    IShowHidePlan_ContainsField = ContainsField(fieldKey)
End Function

Private Function IShowHidePlan_VisibilityFor(ByVal fieldKey As String) As Boolean
    IShowHidePlan_VisibilityFor = VisibilityFor(fieldKey)
End Function
