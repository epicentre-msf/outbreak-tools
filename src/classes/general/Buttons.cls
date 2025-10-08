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
'@label:Buttons
'@section Overview
'@details Provides a light wrapper around worksheet shapes so linelist buttons can be created, positioned, and formatted consistently.

Option Explicit
Implements IButtons

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

'@label:Create
'@jump:Create
'@sub-title Instantiate a button helper bound to an output range
'@details Validates the target range before initialising the internal state so callers receive a ready-to-use instance.
'@param outrng Range identifying the anchor cell for the button.
'@param codeName String name assigned to the shape in the worksheet.
'@param scope ButtonScope controlling the button size and formatting category.
'@return IButtons fully initialised helper.
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

'@section PublicAccessors
'===============================================================================

'@label:Self
'@jump:Self
'@sub-title Return the current helper through its interface
'@return IButtons reference to this instance.
Public Property Get Self() As IButtons
    Set Self = Me
End Property

'@label:OutputRange
'@sub-title Retrieve the anchor range used for placement
'@return Range providing the top-left anchor cell.
Public Property Get OutputRange() As Range
    Set OutputRange = this.rng
End Property

'@sub-title Update the anchor range used for placement
'@param rng Range providing the top-left anchor cell.
Public Property Set OutputRange(ByVal rng As Range)
    Set this.rng = rng
End Property

'@label:Name
'@jump:Name
'@sub-title Retrieve the shape name assigned in the worksheet
'@return String shape identifier.
Public Property Get Name() As String
    Name = this.shapeName
End Property

'@sub-title Store the shape name assigned in the worksheet
'@param nm String identifier for the shape.
Public Property Let Name(ByVal nm As String)
    this.shapeName = nm
End Property

'@label:ShapeScope
'@jump:ShapeScope
'@sub-title Retrieve the formatting scope applied to the button
'@return ButtonScope value describing the button flavour.
Public Property Get ShapeScope() As Byte
    If this.scope = 0 Then
        LogCheck "Shape scope not initialised; defaulting to small", checkingWarning
        this.scope = ButtonScopeSmall
    End If
    ShapeScope = this.scope
End Property

'@sub-title Store the formatting scope applied to the button
'@param scope ButtonScope describing the button flavour.
Public Property Let ShapeScope(ByVal scope As Byte)
    Select Case scope
        Case ButtonScopeLarge, ButtonScopeSmall, ButtonScopeGeo
            this.scope = scope
        Case Else
            LogCheck "Unknown button scope " & CStr(scope) & " provided; defaulting to small", checkingWarning
            this.scope = ButtonScopeSmall
    End Select
End Property

'@section ShapeQueries
'===============================================================================

'@sub-title Determine whether the button shape already exists on the worksheet
'@return Boolean True when the named shape is present.
Private Property Get ShapeExists() As Boolean
    Dim targetShape As Shape
    Dim targetSheet As Worksheet

    Set targetSheet = TargetWorksheet()

    On Error Resume Next
        Set targetShape = targetSheet.Shapes(Name)
    On Error GoTo 0

    ShapeExists = Not targetShape Is Nothing
End Property

'@sub-title Retrieve the underlying shape or raise when absent
'@return Shape object representing the button.
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

'@section ButtonOperations
'===============================================================================

'@label:Add
'@jump:Add
'@sub-title Create the worksheet shape at the configured location
'@details Respects the scope when determining size, applies spacing offsets, and optionally wires up a macro command.
'@param actionCommand Optional String macro to invoke when the button is clicked.
'@param shapeLabel Optional String label displayed on the button; defaults to the shape name.
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

'@label:Format
'@jump:Format
'@sub-title Apply the appropriate visual design to the button
'@details Delegates to an `ILLFormat` provider while mapping the internal ButtonScope to the expected format scope.
'@param design ILLFormat formatter responsible for styling the shape.
'@depends ILLFormat
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

'@label:LogCheck
'@sub-title Append a diagnostic entry to the internal checking log
'@param label String message describing the issue or action.
'@param scope Byte severity scope for the log entry.
Private Sub LogCheck(ByVal label As String, _
                     Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "Buttons " & Name & " trace"

    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)

    internalChecks.Add CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1
    this.hasChecks = True
End Sub

'@label:HasCheckings
'@prop-title Determine whether any checks have been recorded
'@return Boolean True when entries exist.
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:CheckingValues
'@prop-title Retrieve the internal checking log
'@return Object IChecking instance containing accumulated entries.
'@export
Private Property Get CheckingValues() As Object
    If Not HasCheckings Then Exit Property
    Set CheckingValues = internalChecks
End Property

'@section ErrorHandling
'===============================================================================

'@sub-title Raise a class-specific error message
'@param errNumb ProjectError enumeration describing the failure.
'@param errorMessage String describing the failure.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, "Error: " & errorMessage
End Sub

'@section PrivateHelpers
'===============================================================================

'@sub-title Ensure the internal anchor range is available
'@return Range anchor used for calculations.
Private Function TargetRange() As Range
    If this.rng Is Nothing Then
        LogCheck "Button anchor range has not been configured", checkingError
        ThrowError ProjectError.ObjectNotInitialized, _
                   "The button anchor range is not set"
    End If
    Set TargetRange = this.rng
End Function

'@sub-title Retrieve the worksheet hosting the button
'@return Worksheet containing the anchor range.
Private Function TargetWorksheet() As Worksheet
    Set TargetWorksheet = TargetRange().Parent
End Function

'@sub-title Resolve button dimensions from the scope definition
'@param scope ButtonScope describing the button flavour.
'@return TDimensions with width and height in points.
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

'@sub-title Compute the top-left coordinates for the button
'@param anchor Range used as the positional reference.
'@return TPosition containing left/top offsets.
Private Function ResolvePosition(ByVal anchor As Range) As TPosition
    Dim position As TPosition

    position.leftPos = anchor.Left + BUTTON_PADDING + (anchor.Column - 1) * BUTTONGAP
    position.topPos = anchor.Top + BUTTON_PADDING + (anchor.Row - 1) * BUTTONGAP

    ResolvePosition = position
End Function

'@sub-title Determine the label displayed on the button
'@param providedLabel Optional caller-provided label.
'@param fallbackName String used when the label is blank.
'@return String label text.
Private Function EffectiveLabel(ByVal providedLabel As String, ByVal fallbackName As String) As String
    If LenB(providedLabel) = 0 Then
        EffectiveLabel = fallbackName
    Else
        EffectiveLabel = providedLabel
    End If
End Function

'@section Interface
'===============================================================================

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