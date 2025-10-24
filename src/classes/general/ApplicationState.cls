VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Guards Excel application settings while the linelist runs heavy operations"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Guards Excel application settings while the linelist runs heavy operations")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends ProjectError

Option Explicit

Implements IApplicationState

'@section State and constants
'===============================================================================
Private Type TScopeState
    screenUpdating As Boolean                   'Cached ScreenUpdating value
    displayAlerts As Boolean                    'Cached DisplayAlerts value
    calculationMode As XlCalculation            'Cached Calculation mode
    enableAnimations As Boolean                 'Cached EnableAnimations value
    supportsAnimations As Boolean               'Flag indicating whether EnableAnimations is available
    supportsEvents As Boolean
    cursor As XlMousePointer                    'Cached cursor value
    snapshotCaptured As Boolean                 'Flag indicating a snapshot is ready for restoration
    isBusy As Boolean
    calcBeforeSave As Boolean                    'Flag indicating busy mode is currently applied
End Type

Private this As TScopeState
Private excelApp As Application

Private Const CLASSNAME As String = "ApplicationState"


'@section Factory helpers
'===============================================================================
'@sub-title Create a scope bound to an Excel application instance
'@details Captures the current application state immediately so callers can
'@details switch to busy mode without an additional call.
'@param excelApplication Optional target application. Defaults to Excel.Application
'@return A ready-to-use IApplicationState instance.
Public Function Create(Optional ByVal excelApplication As Application) As IApplicationState
    Dim scope As ApplicationState

    Set scope = New ApplicationState
    scope.BindApplication excelApplication
    scope.CaptureSnapshot

    Set Create = scope.Self
End Function

'@sub-title Self reference
Public Property Get Self() As IApplicationState
    Set Self = Me
End Property


'@section Introspection helpers
'===============================================================================
'@sub-title Application currently guarded
Public Property Get ApplicationObject() As Application
    Set ApplicationObject = TargetApplication
End Property

'@sub-title Indicates whether busy state is active
Public Property Get IsBusy() As Boolean
    IsBusy = this.isBusy
End Property

'@sub-title Indicates whether a snapshot is available
Public Property Get HasSnapshot() As Boolean
    HasSnapshot = this.snapshotCaptured
End Property


'@section Internal workflow
'===============================================================================
'@sub-title Bind a target application
Friend Sub BindApplication(ByVal excelApplication As Application)
    If excelApplication Is Nothing Then
        Set excelApp = Nothing
    ElseIf StrComp(TypeName(excelApplication), "Application", vbTextCompare) = 0 Then
        Set excelApp = excelApplication
    Else
        ThrowError ProjectError.InvalidArgument, _
                  "excelApplication must be an Excel Application reference"
    End If
End Sub

'@sub-title Capture current application settings
Friend Sub CaptureSnapshot()
    CaptureSnapshotCore
End Sub


'@section Core logic
'===============================================================================
Private Sub ApplyBusyStateCore(Optional ByVal suppressEvents As Boolean = False, _
                               Optional ByVal calculateOnSave As Boolean = True, _
                               Optional ByVal busyCursor As XlMousePointer = xlNorthWestArrow)
    Dim target As Application

    If this.isBusy Then Exit Sub

    EnsureSnapshot "ApplyBusyState"
    Set target = TargetApplication

    On Error GoTo ApplyFail
        target.ScreenUpdating = False
        target.DisplayAlerts = False
        target.Calculation = xlCalculationManual
        target.EnableEvents = (Not suppressEvents)
        target.CalculateBeforeSave = calculateOnSave
        target.Cursor = busyCursor
        If this.supportsAnimations Then
            SafeSetEnableAnimations target, False
        End If
        this.isBusy = True
    On Error GoTo 0
    Exit Sub

ApplyFail:
    this.isBusy = False
    ThrowError ProjectError.ErrorUnexpectedState, _
              "Unable to apply busy state: " & Err.Description
End Sub

Private Sub CaptureSnapshotCore()
    Dim target As Application

    EnsureNotBusy "CaptureSnapshot"
    Set target = TargetApplication

    this.screenUpdating = target.ScreenUpdating
    this.displayAlerts = target.DisplayAlerts
    this.calculationMode = target.Calculation
    this.supportsAnimations = TryReadEnableAnimations(target, this.enableAnimations)
    this.supportsEvents = target.EnableEvents
    this.calcBeforeSave = target.CalculateBeforeSave
    this.cursor = target.Cursor
    this.snapshotCaptured = True
End Sub

Private Sub RefreshSnapshotCore()
    EnsureNotBusy "RefreshSnapshot"
    CaptureSnapshotCore
End Sub

Private Function TargetApplication() As Application
    If excelApp Is Nothing Then
        Set excelApp = Application
    End If

    If excelApp Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Excel application reference is required"
    End If

    Set TargetApplication = excelApp
End Function

Private Sub EnsureSnapshot(ByVal operationName As String)
    If Not this.snapshotCaptured Then
        CaptureSnapshotCore
    End If

    If Not this.snapshotCaptured Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                  operationName & " requires a captured snapshot. Unexpected error"
    End If
End Sub

Private Sub EnsureNotBusy(ByVal operationName As String)
    If this.isBusy Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                  operationName & " cannot execute while the scope is busy. Call Restore first."
    End If
End Sub

Private Sub RestoreInternal(ByVal silent As Boolean)
    Dim target As Application

    If Not this.snapshotCaptured Then Exit Sub

    Set target = TargetApplication

    On Error GoTo RestoreFail
        target.ScreenUpdating = this.screenUpdating
        target.DisplayAlerts = this.displayAlerts
        target.Calculation = this.calculationMode
        target.EnableEvents = this.supportsEvents
        target.CalculateBeforeSave = this.calcBeforeSave
        target.Cursor = this.cursor
        If this.supportsAnimations Then
            SafeSetEnableAnimations target, this.enableAnimations
        End If
        this.isBusy = False
    On Error GoTo 0
    Exit Sub

RestoreFail:
    this.isBusy = False
    If silent Then
        Err.Clear
    Else
        ThrowError ProjectError.ErrorUnexpectedState, _
                  "Unable to restore application state: " & Err.Description
    End If
End Sub

Private Sub SafeSetEnableAnimations(ByVal target As Application, ByVal value As Boolean)
    On Error GoTo Handler
        target.EnableAnimations = value
    On Error GoTo 0
    Exit Sub
Handler:
    'Mac Excel may not expose EnableAnimations; mark capability as unavailable for future calls.
    this.supportsAnimations = False
    Err.Clear
End Sub

Private Function TryReadEnableAnimations(ByVal target As Application, _
                                         ByRef value As Boolean) As Boolean
    On Error GoTo Unavailable
        value = target.EnableAnimations
        TryReadEnableAnimations = True
    On Error GoTo 0
    Exit Function

Unavailable:
    value = False
    TryReadEnableAnimations = False
    Err.Clear
End Function

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, "ApplicationState", message
End Sub

'@section Interface implementation
'===============================================================================
Private Sub IApplicationState_ApplyBusyState(Optional ByVal suppressEvents As Boolean = False, _
                                             Optional ByVal calculateOnSave As Boolean = True, _
                                             Optional ByVal busyCursor As XlMousePointer = xlNorthWestArrow)
    ApplyBusyStateCore suppressEvents, calculateOnSave, busyCursor
End Sub

Private Sub IApplicationState_Restore()
    RestoreInternal False
End Sub

Private Sub IApplicationState_RefreshSnapshot()
    RefreshSnapshotCore
End Sub

Private Property Get IApplicationState_ApplicationObject() As Application
    Set IApplicationState_ApplicationObject = TargetApplication
End Property

Private Property Get IApplicationState_IsBusy() As Boolean
    IApplicationState_IsBusy = this.isBusy
End Property

Private Property Get IApplicationState_HasSnapshot() As Boolean
    IApplicationState_HasSnapshot = this.snapshotCaptured
End Property
