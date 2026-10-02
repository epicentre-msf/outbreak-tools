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

'@class IProgressBar
'@description
'Defines the public contract for the ProgressBar class, exposing a
'cell-based visual progress indicator on a worksheet. Supports
'configurable colours, number formatting, status messages, and
'step-based or absolute value updates.


'@section Configuration
'===============================================================================
'@description Bind the bar to a range and configure visual properties.

'@jump:attach
'@sub-title Bind the progress bar to a worksheet range
'@details
'The range must be a contiguous area whose cells will be tinted
'to reflect progress.
'@param barRange Range. Hosting the visual progress indicator.
'@param maximum Long. Maximum expected value. Defaults to 100.
Public Sub Attach(ByVal barRange As Range, Optional ByVal maximum As Long = 100)
End Sub

'@jump:configure-colors
'@sub-title Configure the colours for completed and remaining segments
'@param completedColor Long. Colour applied to completed cells.
'@param pendingColor Long. Colour applied to remaining cells.
Public Sub ConfigureColors(ByVal completedColor As Long, ByVal pendingColor As Long)
End Sub

'@jump:configure-value-format
'@sub-title Configure the number format for progress text
'@param formatPattern String. VBA format pattern.
Public Sub ConfigureValueFormat(ByVal formatPattern As String)
End Sub

'@jump:attach-status-cell
'@sub-title Bind a cell for textual status messages
'@param statusRange Range. Single-cell range receiving status messages.
Public Sub AttachStatusCell(ByVal statusRange As Range)
End Sub


'@section Progress
'===============================================================================
'@description Update, step, complete, and reset the progress value.

'@jump:update
'@sub-title Update the current value and re-render the bar
'@param currentValue Long. Absolute progress value.
'@param statusMessage String. Message pushed to the status cell. Defaults to vbNullString.
Public Sub Update(ByVal currentValue As Long, Optional ByVal statusMessage As String = vbNullString)
End Sub

'@jump:step-by
'@sub-title Increment the current value by a step
'@param stepValue Long. Increment. Defaults to 1.
'@param statusMessage String. Message pushed to the status cell. Defaults to vbNullString.
Public Sub StepBy(Optional ByVal stepValue As Long = 1, Optional ByVal statusMessage As String = vbNullString)
End Sub

'@jump:complete
'@sub-title Mark the progress as complete
'@param statusMessage String. Message pushed to the status cell. Defaults to vbNullString.
Public Sub Complete(Optional ByVal statusMessage As String = vbNullString)
End Sub

'@jump:reset
'@sub-title Reset the current value to zero and clear the bar
'@param statusMessage String. Message pushed to the status cell. Defaults to vbNullString.
Public Sub Reset(Optional ByVal statusMessage As String = vbNullString)
End Sub


'@section Properties
'===============================================================================
'@description Read-only and read-write state accessors.

'@jump:bar-range
'@prop-title Range hosting the progress bar
'@return Range. The bar range.
Public Property Get BarRange() As Range
End Property

'@jump:status-cell
'@prop-title Status cell receiving text messages
'@return Range. The status cell, or Nothing.
Public Property Get StatusCell() As Range
End Property

'@jump:maximum
'@prop-title Maximum expected value
'@return Long. The maximum value.
Public Property Get Maximum() As Long
End Property

'@jump:maximum-set
'@prop-title Update the maximum expected value
'@param value Long. New maximum (must be greater than zero).
Public Property Let Maximum(ByVal value As Long)
End Property

'@jump:value
'@prop-title Current progress value
'@return Long. The current value.
Public Property Get Value() As Long
End Property

'@jump:percent-complete
'@prop-title Compute the current completion percentage
'@return Double. Ratio between 0 and 1.
Public Property Get PercentComplete() As Double
End Property
