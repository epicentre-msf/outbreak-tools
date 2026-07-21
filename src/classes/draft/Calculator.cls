VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Running-total arithmetic; a second probe class for the AppleScript test loop"

'@PredeclaredId
'@Folder("Draft")
'@ModuleDescription("Running-total arithmetic; a second probe class for the AppleScript test loop")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

Private Type TCalculator
    total As Double
End Type

Private this As TCalculator

'@section Factory
'===============================================================================

'@sub-title Build a Calculator seeded with an optional starting total.
'@param startValue Optional Double. The initial running total. Defaults to 0.
'@return Calculator ready for use.
Public Function Create(Optional ByVal startValue As Double = 0) As Calculator
    Dim instance As Calculator
    Set instance = New Calculator
    instance.Configure startValue
    Set Create = instance
End Function

'@sub-title Seed the running total on a fresh instance.
'@param startValue Double. The initial running total.
Public Sub Configure(ByVal startValue As Double)
    this.total = startValue
End Sub

'@section Members
'===============================================================================

'@prop-title The current running total.
'@return Double.
Public Property Get Total() As Double
    Total = this.total
End Property

'@sub-title Add a value to the running total.
'@param value Double. The amount to add.
Public Sub Add(ByVal value As Double)
    this.total = this.total + value
End Sub

'@sub-title Subtract a value from the running total.
'@param value Double. The amount to subtract.
Public Sub Subtract(ByVal value As Double)
    this.total = this.total - value
End Sub

'@sub-title Multiply the running total by a factor.
'@param factor Double. The multiplier.
Public Sub Multiply(ByVal factor As Double)
    this.total = this.total * factor
End Sub

'@sub-title Divide the running total by a divisor, guarding against zero.
'@details Leaves the total untouched and reports False when divisor is zero, so
'the loop can prove a guarded path as well as the happy path.
'@param divisor Double. The value to divide by.
'@return Boolean. True when the division was applied, False when divisor was 0.
Public Function SafeDivide(ByVal divisor As Double) As Boolean
    If divisor = 0 Then
        SafeDivide = False
        Exit Function
    End If

    this.total = this.total / divisor
    SafeDivide = True
End Function

'@sub-title Reset the running total to zero.
Public Sub Reset()
    this.total = 0
End Sub
