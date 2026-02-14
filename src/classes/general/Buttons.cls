VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Buttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Buttons of the linelist"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Buttons of the linelist")
'@depends Checking

Option Explicit
Implements IButtons

'@class Buttons
'@description
'Light wrapper around worksheet shapes for creating, positioning,
'and formatting linelist buttons. Supports large, small, and geo
'button scopes with automatic placement offset calculations.
'@depends Checking, ILLFormat, ProjectError
'@version 1.0 (2026-02-09)
'@author Yves Amevoin

'@section State
'===============================================================================

Private Type TButton
    rng As Range
    shapeName As String
    scope As Byte
    hasChecks As Boolean
End Type

Private Type TDimensions
    width As Double
    height As Double
End Type

Private Type TPosition
    leftPos As Double
    topPos As Double
End Type

Private this As TButton

Private Const CLASSNAME As String = "Buttons"
Private Const BUTTONWIDTH As Long = 105
Private Const BUTTONHEIGHT As Long = 30
Private Const BUTTONGAP As Long = 5
Private Const BUTTON_PADDING As Long = 3
Private Const LARGE_MULTIPLIER As Long = 2

Private checkCounter As Long
Private internalChecks As IChecking

'@section Factory
'===============================================================================
'Initialisation and self-reference.

'@label:create
'@sub-title Factory returning an initialised interface
'@param outrng Range. Anchor cell for the button.
'@param codeName String. Name assigned to the shape.
'@param scope Byte. ButtonScope controlling size and format.
'@return IButtons. Ready to use.
'@export
Public Function Create(ByVal outrng As Range, ByVal codeName As String, _
                       ByVal scope As Byte) As IButtons

    If outrng Is Nothing Then
        LogCheck "Missing output range while creating button", checkingError
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Unable to create button: no anchor range provided"
    End If

    With New Buttons
        Set .OutputRange = outrng
        .Name = codeName
        .ShapeScope = scope
        Set Create = .Self
    End With
End Function

'@section Properties
'===============================================================================
'State accessors.

'@label:self
'@prop-title Self-reference as IButtons
'@return IButtons. Self-reference.
Public Property Get Self() As IButtons
    Set Self = Me
End Property

'@label:output-range
'@prop-title Anchor range used for placement
'@return Range. Top-left anchor cell.
'@export
Public Property Get OutputRange() As Range
    Set OutputRange = this.rng
End Property

'@label:output-range-set
'@prop-title Update the anchor range used for placement
'@param rng Range. Providing the top-left anchor cell.
'@export
Public Property Set OutputRange(ByVal rng As Range)
    Set this.rng = rng
End Property

'@label:name
'@prop-title Shape name assigned in the worksheet
'@return String. Shape identifier.
'@export
Public Property Get Name() As String
    Name = this.shapeName
End Property

'@label:name-set
'@prop-title Store the shape name assigned in the worksheet
'@param nm String. Identifier for the shape.
Public Property Let Name(ByVal nm As String)
    this.shapeName = nm
End Property

'@label:shape-scope
'@prop-title Formatting scope applied to the button
'@return Byte. ButtonScope value.
'@export
Public Property Get ShapeScope() As Byte
    If this.scope = 0 Then
        LogCheck "Shape scope not initialised; defaulting to small", checkingWarning
        this.scope = ButtonScopeSmall
    End If
    ShapeScope = this.scope
End Property

'@label:shape-scope-set
'@prop-title Store the formatting scope applied to the button
'@param scope Byte. ButtonScope value.
Public Property Let ShapeScope(ByVal scope As Byte)
    Select Case scope
        Case ButtonScopeLarge, ButtonScopeSmall, ButtonScopeGeo
            this.scope = scope
        Case Else
            LogCheck "Unknown button scope " & CStr(scope) & " provided; defaulting to small", checkingWarning
            this.scope = ButtonScopeSmall
    End Select
End Property

'@section Shape Queries
'===============================================================================
'Shape existence and retrieval.

'@label:shape-exists
'@prop-title Whether the button shape already exists on the worksheet
'@return Boolean. True when the named shape is present.
Private Property Get ShapeExists() As Boolean
    Dim targetShape As Shape
    Dim targetSheet As Worksheet

    Set targetSheet = TargetWorksheet()

    On Error Resume Next
        Set targetShape = targetSheet.Shapes(Name)
    On Error GoTo 0

    ShapeExists = Not targetShape Is Nothing
End Property

'@label:shape-object
'@prop-title Retrieve the underlying shape or raise when absent
'@return Shape. Object representing the button.
Private Property Get ShapeObject() As Shape
    Dim sheet As Worksheet
    Dim shapeName As String

    Set sheet = TargetWorksheet()
    shapeName = Name

    If Not ShapeExists Then
        LogCheck "Shape " & shapeName & " not found on worksheet " & sheet.Name, checkingError
        ThrowError ProjectError.ElementNotFound, _
                   "Shape object " & shapeName & " does not exist in worksheet " & sheet.Name
    End If

    Set ShapeObject = sheet.Shapes(shapeName)
End Property

'@section Operations
'===============================================================================
'Button creation and formatting.

'@label:add
'@sub-title Create the worksheet shape at the configured location
'@param actionCommand String. Macro to invoke on click. Defaults to vbNullString.
'@param shapeLabel String. Label displayed on the button. Defaults to vbNullString.
'@export
Public Sub Add(Optional ByVal actionCommand As String = vbNullString, _
               Optional ByVal shapeLabel As String = vbNullString)

    Dim anchor As Range
    Dim dims As TDimensions
    Dim pos As TPosition
    Dim sheet As Worksheet
    Dim shapeName As String
    Dim labelText As String
    Dim newShape As Shape

    If ShapeExists Then
        LogCheck "Button " & Name & " already exists; skipping creation", checkingWarning
        Exit Sub
    End If

    Set anchor = TargetRange()
    Set sheet = TargetWorksheet()
    shapeName = Name
    dims = ResolveDimensions(ShapeScope)
    pos = ResolvePosition(anchor)
    labelText = EffectiveLabel(shapeLabel, shapeName)

    Set newShape = sheet.Shapes.AddShape(msoShapeRoundedRectangle, _
                                         pos.leftPos, pos.topPos, _
                                         dims.width, dims.height)

    With newShape
        .Name = shapeName
        .OnAction = actionCommand
        .Placement = xlFreeFloating
        .Shadow.Type = msoShadow24
        .TextFrame2.TextRange.Characters.Text = labelText
    End With

End Sub

'@label:format
'@sub-title Apply the appropriate visual design to the button
'@param design ILLFormat. Formatter responsible for styling the shape.
'@export
Public Sub Format(ByVal design As ILLFormat)
    Dim shp As Shape
    Dim formatToken As Variant

    If design Is Nothing Then
        LogCheck "No design provided when formatting button " & Name, checkingError
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Unable to format button: design provider is missing"
    End If

    Set shp = ShapeObject()

    Select Case ShapeScope
        Case ButtonScopeLarge
            formatToken = AdminLargeButton
        Case ButtonScopeSmall
            formatToken = LinelistSmallButton
        Case ButtonScopeGeo
            formatToken = LinelistGeoButton
        Case Else
            LogCheck "Unsupported button scope " & CStr(ShapeScope) & " for button " & Name, checkingError
            ThrowError ProjectError.InvalidArgument, _
                       "Unsupported button scope: " & CStr(ShapeScope)
    End Select

    design.ApplyFormat shp, scope:=formatToken
End Sub

'@section Checkings
'===============================================================================
'Diagnostic logging.

'@label:log-check
'@sub-title Append a diagnostic entry to the internal checking log
'@param label String. Message describing the issue or action.
'@param scope Byte. Severity scope for the log entry.
Private Sub LogCheck(ByVal label As String, _
                     Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "Buttons " & Name & " trace"

    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)

    internalChecks.Add CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1
    this.hasChecks = True
End Sub

'@label:has-checkings
'@prop-title Whether diagnostic entries have been recorded
'@return Boolean. True when at least one entry exists.
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Object. The checking instance, or Nothing.
'@export
Private Property Get CheckingValues() As Object
    If Not HasCheckings Then Exit Property
    Set CheckingValues = internalChecks
End Property

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a typed project error
'@param errNumb Long. ProjectError error code.
'@param errorMessage String. Descriptive error message.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, "Error: " & errorMessage
End Sub

'@section Helpers
'===============================================================================
'Internal utilities.

'@label:target-range
'@sub-title Ensure the internal anchor range is available
'@return Range. Anchor used for calculations.
Private Function TargetRange() As Range
    If this.rng Is Nothing Then
        LogCheck "Button anchor range has not been configured", checkingError
        ThrowError ProjectError.ObjectNotInitialized, _
                   "The button anchor range is not set"
    End If
    Set TargetRange = this.rng
End Function

'@label:target-worksheet
'@sub-title Retrieve the worksheet hosting the button
'@return Worksheet. Containing the anchor range.
Private Function TargetWorksheet() As Worksheet
    Set TargetWorksheet = TargetRange().Parent
End Function

'@label:resolve-dimensions
'@sub-title Resolve button dimensions from the scope definition
'@param scope Byte. ButtonScope value.
'@return TDimensions. Width and height in points.
Private Function ResolveDimensions(ByVal scope As Byte) As TDimensions
    Dim dims As TDimensions

    Select Case scope
        Case ButtonScopeLarge
            dims.width = BUTTONWIDTH * LARGE_MULTIPLIER
            dims.height = BUTTONHEIGHT * LARGE_MULTIPLIER
        Case ButtonScopeSmall, ButtonScopeGeo
            dims.width = BUTTONWIDTH
            dims.height = BUTTONHEIGHT
        Case Else
            dims.width = BUTTONWIDTH
            dims.height = BUTTONHEIGHT
            LogCheck "Unknown button scope " & CStr(scope) & " defaulted to standard dimensions", checkingWarning
    End Select

    ResolveDimensions = dims
End Function

'@label:resolve-position
'@sub-title Compute the top-left coordinates for the button
'@param anchor Range. Positional reference.
'@return TPosition. Left and top offsets.
Private Function ResolvePosition(ByVal anchor As Range) As TPosition
    Dim position As TPosition

    position.leftPos = anchor.Left + BUTTON_PADDING + (anchor.Column - 1) * BUTTONGAP
    position.topPos = anchor.Top + BUTTON_PADDING + (anchor.Row - 1) * BUTTONGAP

    ResolvePosition = position
End Function

'@label:effective-label
'@sub-title Determine the label displayed on the button
'@param providedLabel String. Caller-provided label.
'@param fallbackName String. Fallback when the label is blank.
'@return String. Label text.
Private Function EffectiveLabel(ByVal providedLabel As String, ByVal fallbackName As String) As String
    If LenB(providedLabel) = 0 Then
        EffectiveLabel = fallbackName
    Else
        EffectiveLabel = providedLabel
    End If
End Function

'@section Interface
'===============================================================================
'IButtons delegation wiring.

Private Sub IButtons_Add(Optional ByVal actionCommand As String = vbNullString, _
                         Optional ByVal shapeLabel As String = vbNullString)
    Add actionCommand:=actionCommand, shapeLabel:=shapeLabel
End Sub

Private Sub IButtons_Format(design As ILLFormat)
    Format design
End Sub

Private Property Get IButtons_OutputRange() As Range
    Set IButtons_OutputRange = OutputRange
End Property

Private Property Get IButtons_Name() As String
    IButtons_Name = Name()
End Property

Private Property Set IButtons_OutputRange(ByVal rng As Range)
    Set OutputRange = rng
End Property

Private Property Get IButtons_HasCheckings() As Boolean
    IButtons_HasCheckings = HasCheckings
End Property

Private Property Get IButtons_CheckingValues() As Object
    Set IButtons_CheckingValues = CheckingValues
End Property

Private Property Get IButtons_ShapeScope() As Byte
    IButtons_ShapeScope = ShapeScope()
End Property