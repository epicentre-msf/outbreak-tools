VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Worksheet progress bar helper"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Worksheet progress bar helper used while generating linelist artefacts")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, ExcelMemberMayReturnNothing, HungarianNotation
'@depends IProgressBar, ProjectError

Option Explicit

Implements IProgressBar

'@class ProgressBar
'@description
'Cell-based visual progress indicator on a worksheet. Tints cells
'to reflect completion ratio and optionally writes status text to
'a companion cell.
'@depends ProjectError
'@version 1.0 (2026-02-09)
'@author Yves Amevoin

Private Type TProgressBar
    barRange As Range
    statusRange As Range
    maximum As Long
    value As Long
    isInitialised As Boolean
    completedColor As Long
    pendingColor As Long
    valueFormat As String
    statusText As String
End Type

Private this As TProgressBar

Private Const CLASSNAME As String = "ProgressBar"
Private Const DEFAULT_MAXIMUM As Long = 100
Private Const DEFAULT_VALUE_FORMAT As String = "#,##0"
Private Const DEFAULT_COMPLETED_COLOR As Long = &H4DB870 '  Clean green
Private Const DEFAULT_PENDING_COLOR As Long = &HEEF1F5   '  Light grey

'@section Lifecycle
'===============================================================================
'Initialisation and factory.

'@label:class-initialize
'@sub-title Set default values for colours, format, and maximum
Private Sub Class_Initialize()
    this.maximum = DEFAULT_MAXIMUM
    this.completedColor = DEFAULT_COMPLETED_COLOR
    this.pendingColor = DEFAULT_PENDING_COLOR
    this.valueFormat = DEFAULT_VALUE_FORMAT
End Sub

'@label:create
'@sub-title Factory returning an initialised interface
'@param barRange Range. Range that will host the indicator.
'@param maximum Long. Maximum value. Defaults to DEFAULT_MAXIMUM.
'@return IProgressBar. Ready to use.
'@export
Public Function Create(ByVal barRange As Range, Optional ByVal maximum As Long = DEFAULT_MAXIMUM) As IProgressBar
    With New ProgressBar
        .Attach barRange, maximum
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Self-reference as IProgressBar
'@return IProgressBar. Self-reference.
Public Property Get Self() As IProgressBar
    Set Self = Me
End Property

'@section Public API
'===============================================================================

'@label:attach
'@sub-title Bind the progress bar to a worksheet range
'@param barRange Range. Hosting the progress indicator.
'@param maximum Long. Maximum value. Defaults to DEFAULT_MAXIMUM.
'@export
Public Sub Attach(ByVal barRange As Range, Optional ByVal maximum As Long = DEFAULT_MAXIMUM)
    ValidateBarRange barRange

    With this
        Set .barRange = barRange
        .maximum = ResolveMaximum(maximum)
        .value = 0
        .isInitialised = True
    End With

    PrepareBarRange
    RenderProgress
    RenderStatus vbNullString
End Sub

'@label:configure-colors
'@sub-title Configure the tint for completed and pending cells
'@param completedColor Long. Colour applied to completed cells.
'@param pendingColor Long. Colour applied to remaining cells.
'@export
Public Sub ConfigureColors(ByVal completedColor As Long, ByVal pendingColor As Long)
    this.completedColor = completedColor
    this.pendingColor = pendingColor

    If this.isInitialised Then
        RenderProgress
    End If
End Sub

'@label:configure-value-format
'@sub-title Configure the number format for progress text
'@param formatPattern String. VBA format pattern.
'@export
Public Sub ConfigureValueFormat(ByVal formatPattern As String)
    If LenB(formatPattern) = 0 Then
        ThrowError ProjectError.InvalidArgument, "formatPattern cannot be blank"
    End If

    this.valueFormat = formatPattern

    If this.isInitialised Then
        RenderProgress
    End If
End Sub

'@label:attach-status-cell
'@sub-title Bind a cell for textual status messages
'@param statusRange Range. Single-cell range that will receive status text.
'@export
Public Sub AttachStatusCell(ByVal statusRange As Range)
    ValidateStatusRange statusRange
    Set this.statusRange = statusRange
    RenderStatus vbNullString
End Sub

'@label:update
'@sub-title Update progress to a specific value
'@param currentValue Long. Absolute progress value.
'@param statusMessage String. Message pushed to the status cell. Defaults to vbNullString.
'@export
Public Sub Update(ByVal currentValue As Long, Optional ByVal statusMessage As String = vbNullString)
    EnsureInitialised

    this.value = ClampValue(currentValue)
    RenderProgress
    RenderStatus statusMessage
End Sub

'@label:step-by
'@sub-title Increment progress by the provided step
'@param stepValue Long. Increment. Defaults to 1.
'@param statusMessage String. Message pushed to the status cell. Defaults to vbNullString.
'@export
Public Sub StepBy(Optional ByVal stepValue As Long = 1, Optional ByVal statusMessage As String = vbNullString)
    EnsureInitialised

    Dim targetValue As Long
    targetValue = this.value + stepValue

    Update targetValue, statusMessage
End Sub

'@label:complete
'@sub-title Set progress to its maximum state
'@param statusMessage String. Message pushed to the status cell. Defaults to vbNullString.
'@export
Public Sub Complete(Optional ByVal statusMessage As String = vbNullString)
    EnsureInitialised
    Update this.maximum, statusMessage
End Sub

'@label:reset
'@sub-title Reset progress to zero
'@param statusMessage String. Message pushed to the status cell. Defaults to vbNullString.
'@export
Public Sub Reset(Optional ByVal statusMessage As String = vbNullString)
    EnsureInitialised

    this.value = 0
    RenderProgress
    RenderStatus statusMessage
End Sub

'@section Properties
'===============================================================================
'Read-only and read-write state accessors.

'@label:bar-range
'@prop-title Range hosting the progress bar
'@return Range. The bar range.
'@export
Public Property Get BarRange() As Range
    Set BarRange = this.barRange
End Property

'@label:status-cell
'@prop-title Status cell receiving text messages
'@return Range. The status cell, or Nothing.
'@export
Public Property Get StatusCell() As Range
    If this.statusRange Is Nothing Then Exit Property
    Set StatusCell = this.statusRange
End Property

'@label:maximum
'@prop-title Maximum expected value
'@return Long. The maximum value.
'@export
Public Property Get Maximum() As Long
    Maximum = this.maximum
End Property

'@label:maximum-set
'@prop-title Update the maximum expected value
'@param value Long. New maximum (must be greater than zero).
'@export
Public Property Let Maximum(ByVal value As Long)
    this.maximum = ResolveMaximum(value)
    this.value = ClampValue(this.value)

    If this.isInitialised Then
        RenderProgress
    End If
End Property

'@label:value
'@prop-title Current progress value
'@return Long. The current value.
'@export
Public Property Get Value() As Long
    Value = this.value
End Property

'@label:percent-complete
'@prop-title Compute the current completion percentage
'@return Double. Ratio between 0 and 1.
'@export
Public Property Get PercentComplete() As Double
    PercentComplete = ComputeRatio(this.value)
End Property

'@section Interface
'===============================================================================
'IProgressBar delegation wiring.

Private Sub IProgressBar_Attach(ByVal barRange As Range, Optional ByVal maximum As Long = DEFAULT_MAXIMUM)
    Attach barRange, maximum
End Sub

Private Sub IProgressBar_AttachStatusCell(ByVal statusRange As Range)
    AttachStatusCell statusRange
End Sub

Private Property Get IProgressBar_BarRange() As Range
    Set IProgressBar_BarRange = BarRange
End Property

Private Sub IProgressBar_Complete(Optional ByVal statusMessage As String = vbNullString)
    Complete statusMessage
End Sub

Private Sub IProgressBar_ConfigureColors(ByVal completedColor As Long, ByVal pendingColor As Long)
    ConfigureColors completedColor, pendingColor
End Sub

Private Sub IProgressBar_ConfigureValueFormat(ByVal formatPattern As String)
    ConfigureValueFormat formatPattern
End Sub

Private Sub IProgressBar_Reset(Optional ByVal statusMessage As String = vbNullString)
    Reset statusMessage
End Sub

Private Sub IProgressBar_StepBy(Optional ByVal stepValue As Long = 1, Optional ByVal statusMessage As String = vbNullString)
    StepBy stepValue, statusMessage
End Sub

Private Property Get IProgressBar_StatusCell() As Range
    Set IProgressBar_StatusCell = StatusCell
End Property

Private Sub IProgressBar_Update(ByVal currentValue As Long, Optional ByVal statusMessage As String = vbNullString)
    Update currentValue, statusMessage
End Sub

Private Property Get IProgressBar_Maximum() As Long
    IProgressBar_Maximum = Maximum
End Property

Private Property Let IProgressBar_Maximum(ByVal value As Long)
    Maximum = value
End Property

Private Property Get IProgressBar_Value() As Long
    IProgressBar_Value = Value
End Property

Private Property Get IProgressBar_PercentComplete() As Double
    IProgressBar_PercentComplete = PercentComplete
End Property

'@section Internal Helpers
'===============================================================================
'Validation, clamping, rendering, and error handling.

'@label:ensure-initialised
'@sub-title Guard against use before Attach is called
Private Sub EnsureInitialised()
    If Not this.isInitialised Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Attach must be called before updating the progress bar"
    End If
End Sub

'@label:validate-bar-range
'@sub-title Validate the bar range argument
Private Sub ValidateBarRange(ByVal barRange As Range)
    If barRange Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "barRange cannot be Nothing"
    End If

    If barRange.Areas.Count > 1 Then
        ThrowError ProjectError.InvalidArgument, "barRange must be a contiguous range"
    End If

    If barRange.Cells.Count = 0 Then
        ThrowError ProjectError.InvalidArgument, "barRange must contain at least one cell"
    End If
End Sub

'@label:validate-status-range
'@sub-title Validate the status range argument
Private Sub ValidateStatusRange(ByVal statusRange As Range)
    If statusRange Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "statusRange cannot be Nothing"
    End If

    If statusRange.Cells.Count <> 1 Then
        ThrowError ProjectError.InvalidArgument, "statusRange must reference a single cell"
    End If
End Sub

'@label:clamp-value
'@sub-title Clamp a value to the 0..maximum range
'@return Long. The clamped value.
Private Function ClampValue(ByVal candidateValue As Long) As Long
    If candidateValue < 0 Then
        ClampValue = 0
    ElseIf candidateValue > this.maximum Then
        ClampValue = this.maximum
    Else
        ClampValue = candidateValue
    End If
End Function

'@label:compute-ratio
'@sub-title Compute the completion ratio for a value
'@return Double. Ratio between 0 and 1.
Private Function ComputeRatio(ByVal candidateValue As Long) As Double
    Dim ratio As Double

    ratio = candidateValue / this.maximum

    ComputeRatio = ClampRatio(ratio)
End Function

'@label:clamp-ratio
'@sub-title Clamp a ratio to 0..1
'@return Double. The clamped ratio.
Private Function ClampRatio(ByVal ratio As Double) As Double
    If ratio < 0# Then
        ClampRatio = 0#
    ElseIf ratio > 1# Then
        ClampRatio = 1#
    Else
        ClampRatio = ratio
    End If
End Function

'@label:render-progress
'@sub-title Re-render the bar and value text
Private Sub RenderProgress()
    Dim ratio As Double
    ratio = ComputeRatio(this.value)
    PaintBar ratio
    WriteValue
End Sub

'@label:render-status
'@sub-title Write the latest status text to the status cell
Private Sub RenderStatus(ByVal statusMessage As String)
    If LenB(statusMessage) > 0 Then
        this.statusText = statusMessage
    End If

    If this.statusRange Is Nothing Then Exit Sub

    this.statusRange.Value = this.statusText
End Sub

'@label:prepare-bar-range
'@sub-title Clear all cells and set them to the pending colour
Private Sub PrepareBarRange()
    Dim cell As Range

    For Each cell In this.barRange.Cells
        cell.Value = vbNullString
        cell.Interior.Color = this.pendingColor
    Next cell
End Sub

'@label:paint-bar
'@sub-title Tint cells according to the completion ratio
Private Sub PaintBar(ByVal ratio As Double)
    Dim filledCount As Long
    Dim totalCells As Long
    Dim cell As Range
    Dim index As Long

    If this.barRange Is Nothing Then Exit Sub

    totalCells = this.barRange.Cells.Count
    If totalCells = 0 Then Exit Sub

    filledCount = CLng(Fix(ratio * totalCells))

    If ratio > 0# And filledCount = 0 Then
        filledCount = 1
    ElseIf filledCount > totalCells Then
        filledCount = totalCells
    End If

    index = 0
    For Each cell In this.barRange.Cells
        index = index + 1
        If index <= filledCount Then
            cell.Interior.Color = this.completedColor
        Else
            cell.Interior.Color = this.pendingColor
        End If
    Next cell
End Sub

'@label:write-value
'@sub-title Write the formatted progress text to the first cell
Private Sub WriteValue()
    If this.barRange Is Nothing Then Exit Sub

    Dim formattedValue As String
    formattedValue = Format$(this.value, this.valueFormat) & " / " & Format$(this.maximum, this.valueFormat)

    With this.barRange.Cells(1)
        .Value = formattedValue
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

'@label:resolve-maximum
'@sub-title Validate and return a positive maximum
'@return Long. The validated maximum.
Private Function ResolveMaximum(ByVal candidateMaximum As Long) As Long
    If candidateMaximum <= 0 Then
        ThrowError ProjectError.InvalidArgument, "Maximum must be greater than zero"
    End If

    ResolveMaximum = candidateMaximum
End Function

'@label:throw-error
'@sub-title Raise a typed project error
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, "Error: " & message
End Sub
