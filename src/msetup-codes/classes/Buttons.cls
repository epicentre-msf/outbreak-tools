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

'@IgnoreModule UseMeaningfulName, HungarianNotation
'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Buttons of the linelist")

Option Explicit
Implements IButtons

Private Type TButton
    rng As Range
    shapename As String
    scope As ButtonScope
End Type

Private this As TButton
Private Const CLASSNAME As String = "Buttons"
Private Const BUTTONWIDTH As Long = 105
Private Const BUTTONHEIGHT As Long = 30
Private Const BUTTONGAP As Long = 5

Public Function Create(ByVal outrng As Range, ByVal codeName As String, _
                       ByVal scope As ButtonScope) As IButtons
    
    'Check if the output range is provided
    If outrng Is Nothing Then ThrowError "I don't know where to write the button, no range found."

    'Create a new button object and set properties
    With New Buttons
        Set .OutputRange = outrng
        .Name = codeName
        .ShapeScope = scope
        Set Create = .Self()
    End With
End Function

'Encapsulated objects
Public Property Get Self() As IButtons
    Set Self = Me
End Property

Public Property Get OutputRange() As Range
    Set OutputRange = this.rng
End Property

Public Property Set OutputRange(ByVal rng As Range)
    Set this.rng = rng
End Property

Public Property Get Name() As String
    Name = this.shapename
End Property

Public Property Let Name(ByVal nm As String)
    this.shapename = nm
End Property

Public Property Let ShapeScope(ByVal scope As ButtonScope)
    this.scope = scope
End Property

Public Property Get ShapeScope() As ButtonScope
    ShapeScope = this.scope
End Property

Private Property Get ShapeExists() As Boolean
    Dim shp As Shape
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim codeName As String

    codeName = Name()
    Set cellRng = OutputRange()
    Set sh = cellRng.Parent

    On Error Resume Next
    Set shp = sh.Shapes(codeName)
    ShapeExists = Not (shp Is Nothing)
    On Error GoTo 0
End Property

Private Property Get ShapeObject() As Shape
    Dim cellRng As Range
    Dim codeName As String
    Dim sh As Worksheet

    Set cellRng = OutputRange()
    Set sh = cellRng.Parent
    codeName = Name()

    If Not ShapeExists() Then ThrowError "Shape object " & codeName & "does not exists in worksheet " & sh.Name
    Set ShapeObject = sh.Shapes(codeName)

End Property

'Add the button
Public Sub Add(Optional ByVal actionCommand As String = vbNullString, Optional ByVal shapeLabel As String = vbNullString)

    Dim height As Long
    Dim width As Long
    Dim topGap As Long
    Dim leftGap As Long
    Dim cellRng As Range
    Dim sh As Worksheet
    Dim codeName As String


    If ShapeExists() Then Exit Sub

    Set cellRng = OutputRange()
    codeName = Name()
  
    'Long button
    If ShapeScope = ButtonScopeLarge Then
        height = 2 * BUTTONHEIGHT
        width = 2 * BUTTONWIDTH
    Else
        'Short Button
        height = BUTTONHEIGHT
        width = BUTTONWIDTH
    End If
    
    Set sh = cellRng.Parent
    topGap = BUTTONGAP * (cellRng.Row - 1) + 3   'Gap between multiple buttons on the same column
    leftGap = BUTTONGAP * (cellRng.Column - 1) + 3

    'Add the button
    sh.Shapes.AddShape(msoShapeRoundedRectangle, cellRng.Left + leftGap, cellRng.Top + topGap, width, height).Name = codeName

    'Set The text to display on the shape
    With sh.Shapes(codeName)
        .OnAction = actionCommand
        .Placement = xlFreeFloating
        .Shadow.Type = msoShadow24
        .TextFrame2.TextRange.Characters.Text = shapeLabel
    End With
End Sub

'Format the button using the provided design
Public Sub Format(Optional ByVal fontColor As Variant = vbNullString, _
                  Optional ByVal interiorColor As Variant = vbNullString, _
                  Optional ByVal isBold As Boolean = False, _
                  Optional ByVal fontSize As Double = 12)
    Dim shp As Shape

    Set shp = ShapeObject()

    With shp
        'Fill color
        If interiorColor <> vbNullString Then
            .Fill.ForeColor.RGB = interiorColor
            .Fill.BackColor.RGB = interiorColor
        End If

        With .TextFrame2
            'Text alignment
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoTrue

            With .TextRange
                .ParagraphFormat.Alignment = msoAlignCenter

                'Font Zize and Font Shape
                If FontSize <> 0 Then .Font.Size = FontSize
                If isBold Then .Font.Bold = isBold
                'Text Color
                If fontColor <> vbNullString Then .Font.Fill.ForeColor.RGB = fontColor
            End With
        End With
    End With
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Exposed methods
Private Sub IButtons_Add(Optional ByVal actionCommand As String = vbNullString, _
                        Optional ByVal shapeLabel As String = vbNullString)
    Add actionCommand:=actionCommand, shapeLabel:=shapeLabel
End Sub

Private Sub IButtons_Format(Optional ByVal fontColor As Variant = vbWhite, _
                   Optional ByVal interiorColor As Variant = RGB(24, 149, 237), _
                   Optional ByVal isBold As Boolean = False, _
                   Optional ByVal fontSize As Double = 12)
    Format fontColor:=fontColor, interirorColor:=interiorColor, isBold:=isBold, fontSize:=fontSize
End Sub
