VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for worksheet progress bars"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface for worksheet progress bars used while building the linelist")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ProcedureNotUsed, UseMeaningfulName, ExcelMemberMayReturnNothing

Option Explicit

'@section Configuration
'===============================================================================

'@jump:Attach
'@sub-title Bind the progress bar to a worksheet range
'@details The range must be a contiguous area whose cells will be tinted to reflect progress.
'@param barRange Range hosting the visual progress indicator.
'@param maximum Optional Long maximum value. Defaults to 100.
Public Sub Attach(ByVal barRange As Range, Optional ByVal maximum As Long = 100)
End Sub

'@jump:ConfigureColors
'@sub-title Configure the colors used to render completed and remaining segments
'@param completedColor Long color applied to completed cells.
'@param pendingColor Long color applied to remaining cells.
Public Sub ConfigureColors(ByVal completedColor As Long, ByVal pendingColor As Long)
End Sub

'@jump:ConfigureValueFormat
'@sub-title Configure the number format used to display progress text
'@param formatPattern String VBA format pattern used when rendering the value and maximum.
Public Sub ConfigureValueFormat(ByVal formatPattern As String)
End Sub

'@jump:AttachStatusCell
'@sub-title Bind a cell that will display textual status updates
'@param statusRange Single-cell range receiving status messages passed to Update/StepBy/Complete/Reset.
Public Sub AttachStatusCell(ByVal statusRange As Range)
End Sub

'@section Progress
'===============================================================================

'@jump:Update
'@sub-title Update the current value and re-render the bar
'@param currentValue Long absolute progress value.
'@param statusMessage Optional String message pushed to the status cell (when configured).
Public Sub Update(ByVal currentValue As Long, Optional ByVal statusMessage As String = vbNullString)
End Sub

'@jump:StepBy
'@sub-title Increment the current value by a step
'@param stepValue Optional Long increment. Defaults to 1.
'@param statusMessage Optional String message pushed to the status cell (when configured).
Public Sub StepBy(Optional ByVal stepValue As Long = 1, Optional ByVal statusMessage As String = vbNullString)
End Sub

'@jump:Complete
'@sub-title Mark the progress as complete
'@param statusMessage Optional String message pushed to the status cell (when configured).
Public Sub Complete(Optional ByVal statusMessage As String = vbNullString)
End Sub

'@jump:Reset
'@sub-title Reset the current value to zero and clear the bar
'@param statusMessage Optional String message pushed to the status cell (when configured).
Public Sub Reset(Optional ByVal statusMessage As String = vbNullString)
End Sub

'@section Properties
'===============================================================================

'@jump:BarRange
'@sub-title Retrieve the range hosting the progress bar
'@return Range currently linked to the bar.
Public Property Get BarRange() As Range
End Property

'@jump:StatusCell
'@sub-title Retrieve the status cell (if any)
'@return Range single-cell range receiving status messages.
Public Property Get StatusCell() As Range
End Property

'@jump:Maximum
'@sub-title Maximum expected value
'@return Long maximum value used for computations.
Public Property Get Maximum() As Long
End Property

'@sub-title Update the maximum expected value
'@param value Long new maximum value (must be greater than zero).
Public Property Let Maximum(ByVal value As Long)
End Property

'@jump:Value
'@sub-title Current value
'@return Long current progress value.
Public Property Get Value() As Long
End Property

'@jump:PercentComplete
'@sub-title Compute the current completion percentage
'@return Double ratio between 0 and 1 describing completion.
Public Property Get PercentComplete() As Double
End Property
