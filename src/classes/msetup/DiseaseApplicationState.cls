VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseApplicationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Guards and restores Excel Application state for disease workflows"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Guards and restores Excel Application state for disease workflows")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'DiseaseApplicationState.cls without interface -- single implementation, no
'test fake, no compile cycle (see project-rules.md 4.4).

'@section Types
'===============================================================================

Private Type TApplicationState
    App As Application
    IsGuarding As Boolean
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
    DisplayStatusBar As Boolean
    StatusBarManaged As Boolean
End Type

Private this As TApplicationState

'@section Factory
'===============================================================================

'@sub-title Build a fresh, unarmed state guard.
'@return DiseaseApplicationState ready to arm with BeginGuard.
Public Function Create() As DiseaseApplicationState
    Set Create = New DiseaseApplicationState
End Function

'@section Commands
'===============================================================================

'@sub-title Snapshot the Application state then disable interactive toggles.
'@param app Application whose state is captured and guarded.
'@param disableStatusBar Optional Boolean; when True the status bar is also hidden (default True).
Public Sub BeginGuard(ByVal app As Application, Optional ByVal disableStatusBar As Boolean = True)
    ValidateApplication app

    ' A re-entrant BeginGuard must not overwrite the originally captured state,
    ' otherwise Restore would put back the already-disabled values.
    If this.IsGuarding Then Exit Sub

    Set this.App = app
    this.ScreenUpdating = app.ScreenUpdating
    this.DisplayAlerts = app.DisplayAlerts
    this.EnableEvents = app.EnableEvents
    this.Calculation = app.Calculation

    app.ScreenUpdating = False
    app.DisplayAlerts = False
    app.EnableEvents = False

    ' Assigning Calculation can flush pending calculations and raise 1004 when no
    ' workbook context is available; isolate it so the guard still arms.
    On Error Resume Next
    app.Calculation = xlCalculationManual
    On Error GoTo 0

    If disableStatusBar Then
        this.DisplayStatusBar = app.DisplayStatusBar
        this.StatusBarManaged = True
        app.DisplayStatusBar = False
    End If

    this.IsGuarding = True
End Sub

'@sub-title Restore the Application state captured by BeginGuard.
'@details Safe to call repeatedly: a Restore with nothing guarded is a no-op.
Public Sub Restore()
    ' Idempotent guard: nothing to restore means a harmless no-op (never raises).
    If Not this.IsGuarding Then Exit Sub

    If this.App Is Nothing Then
        this.IsGuarding = False
        Exit Sub
    End If

    this.App.EnableEvents = this.EnableEvents
    this.App.DisplayAlerts = this.DisplayAlerts

    ' Mirror BeginGuard: shield the Calculation restore from a 1004 flush error.
    On Error Resume Next
    this.App.Calculation = this.Calculation
    On Error GoTo 0

    If this.StatusBarManaged Then
        this.App.DisplayStatusBar = this.DisplayStatusBar
    End If

    ' ScreenUpdating is restored last so the UI repaints only once, after every
    ' other toggle is back to its captured value.
    this.App.ScreenUpdating = this.ScreenUpdating

    this.StatusBarManaged = False
    this.IsGuarding = False
    Set this.App = Nothing
End Sub

'@section Internal Helpers
'===============================================================================

'@sub-title Ensure the Application reference is valid.
Private Sub ValidateApplication(ByVal app As Application)
    If app Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Application reference cannot be Nothing"
    End If
End Sub

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised error raising aligned with ProjectError codes.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseApplicationState", errMessage
End Sub
