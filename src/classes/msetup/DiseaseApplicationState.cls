VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseApplicationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Guards Excel Application state for disease setup workflows"

'@Folder("Msetup")
'@ModuleDescription("Guards Excel Application state for disease setup workflows")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDiseaseApplicationState

'@section State
'===============================================================================

Private Type TApplicationState
    ExcelApp As Application
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
    EnableAnimations As Boolean
    CalculateBeforeSave As Boolean
    DisplayStatusBar As Boolean
    StatusBarValue As Variant
    HasCustomStatusBar As Boolean
    HasSnapshot As Boolean
    IsActive As Boolean
End Type

Private this As TApplicationState

'@section Lifecycle
'===============================================================================

'@sub-title Capture the current Application state and apply optimised settings
'@param excelApp Application reference to guard.
'@param disableStatusBar Optional Boolean. When True, StatusBar is cleared while busy.
Public Sub BeginGuard(ByVal excelApp As Application, Optional ByVal disableStatusBar As Boolean = True)
    ValidateApplication excelApp
    ActivateGuard excelApp, disableStatusBar
End Sub

'@sub-title Restore the Application to its original state.
'@remarks Safe to call multiple times; subsequent calls simply exit.
Public Sub Restore()
    If Not this.IsActive Then Exit Sub
    RestoreState
End Sub

'@section Accessors
'===============================================================================

'@sub-title Indicates whether the guard currently holds the Application in busy mode.
'@return Boolean True when BeginGuard has been called and Restore not yet invoked.
Public Property Get Active() As Boolean
    Active = this.IsActive
End Property

'@sub-title Retrieve the guarded Application reference.
'@return Application reference used during BeginGuard or Nothing when inactive.
Public Property Get HostApplication() As Application
    Set HostApplication = this.ExcelApp
End Property

'@section Internal Helpers
'===============================================================================

'@sub-title Ensure the supplied Application reference is valid.
Private Sub ValidateApplication(ByVal excelApp As Application)
    If excelApp Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Application reference cannot be Nothing"
    End If
End Sub

'@sub-title Capture the current state and apply optimised values.
Private Sub ActivateGuard(ByVal excelApp As Application, ByVal disableStatusBar As Boolean)
    On Error GoTo Fail

    If this.IsActive Then RestoreState

    CaptureState excelApp
    ApplyOptimisedState excelApp, disableStatusBar

    this.IsActive = True
    Exit Sub
Fail:
    RestoreState
    ThrowError ProjectError.SomethingWentWrong, "Unable to initialise Application guard: " & Err.Description
End Sub

'@sub-title Capture Application properties so they can be restored later.
Private Sub CaptureState(ByVal excelApp As Application)
    With excelApp
        this.ScreenUpdating = .ScreenUpdating
        this.EnableEvents = .EnableEvents
        this.DisplayAlerts = .DisplayAlerts
        this.Calculation = .Calculation
        this.EnableAnimations = .EnableAnimations
        this.CalculateBeforeSave = .CalculateBeforeSave
        this.DisplayStatusBar = .DisplayStatusBar

        On Error Resume Next
            this.StatusBarValue = .StatusBar
            this.HasCustomStatusBar = (Err.Number = 0)
        On Error GoTo 0
    End With

    Set this.ExcelApp = excelApp
    this.HasSnapshot = True
End Sub

'@sub-title Apply optimised Application settings for performant workflows.
Private Sub ApplyOptimisedState(ByVal excelApp As Application, ByVal disableStatusBar As Boolean)
    With excelApp
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .EnableAnimations = False
        .CalculateBeforeSave = False
        .Calculation = xlCalculationManual

        If disableStatusBar Then
            .DisplayStatusBar = True
            On Error Resume Next
                .StatusBar = False
            On Error GoTo 0
        End If
    End With
End Sub

'@sub-title Restore the Application to the captured state.
Private Sub RestoreState()
    On Error Resume Next

    If this.HasSnapshot And Not this.ExcelApp Is Nothing Then
        With this.ExcelApp
            .ScreenUpdating = this.ScreenUpdating
            .EnableEvents = this.EnableEvents
            .DisplayAlerts = this.DisplayAlerts
            .EnableAnimations = this.EnableAnimations
            .CalculateBeforeSave = this.CalculateBeforeSave
            .Calculation = this.Calculation
            .DisplayStatusBar = this.DisplayStatusBar

            If this.HasCustomStatusBar Then
                .StatusBar = this.StatusBarValue
            Else
                .StatusBar = False
            End If
        End With
    End If

    On Error GoTo 0
    Set this.ExcelApp = Nothing
    this.IsActive = False
    this.HasSnapshot = False
    this.HasCustomStatusBar = False
    this.StatusBarValue = Empty
End Sub

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised error raising aligned with ProjectError codes.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseApplicationState", errMessage
End Sub

'@section ClassLifecycle
'===============================================================================

Private Sub Class_Terminate()
    On Error Resume Next
        Restore
    On Error GoTo 0
End Sub

'@section Interface
'===============================================================================

Private Sub IDiseaseApplicationState_BeginGuard(ByVal excelApp As Application, Optional ByVal disableStatusBar As Boolean = True)
    BeginGuard excelApp, disableStatusBar
End Sub

Private Sub IDiseaseApplicationState_Restore()
    Restore
End Sub

Private Property Get IDiseaseApplicationState_Active() As Boolean
    IDiseaseApplicationState_Active = Active
End Property

Private Property Get IDiseaseApplicationState_HostApplication() As Application
    Set IDiseaseApplicationState_HostApplication = HostApplication
End Property
