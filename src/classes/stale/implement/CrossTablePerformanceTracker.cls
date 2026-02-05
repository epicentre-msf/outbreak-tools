VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTablePerformanceTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Captures timing information for cross table helper operations"

'@Folder("Implement")
'@ModuleDescription("Lightweight timer used to profile cross table helper workflows")

'@section State
'===============================================================================

Option Explicit

Private Type TCrossTablePerformanceTracker
    startTime As Double
    elapsedSeconds As Double
    running As Boolean
End Type

Private this As TCrossTablePerformanceTracker
Private Const CLASSNAME As String = "CrossTablePerformanceTracker"
Private Const SECONDS_PER_DAY As Double = 86400#

'@section Public API
'===============================================================================

Public Sub BeginMeasurement()
    If this.running Then
        Err.Raise ProjectError.ErrorUnexpectedState, CLASSNAME, _
                  "Measurement already in progress"
    End If

    this.startTime = Timer
    this.running = True
End Sub

Public Sub EndMeasurement()
    Dim stopTime As Double

    If Not this.running Then
        Err.Raise ProjectError.ErrorUnexpectedState, CLASSNAME, _
                  "BeginMeasurement must be called before EndMeasurement"
    End If

    stopTime = Timer
    If stopTime < this.startTime Then
        this.elapsedSeconds = (stopTime + SECONDS_PER_DAY) - this.startTime
    Else
        this.elapsedSeconds = stopTime - this.startTime
    End If

    this.running = False
End Sub

Public Sub Reset()
    this.startTime = 0
    this.elapsedSeconds = 0
    this.running = False
End Sub

Public Property Get DurationMilliseconds() As Double
    DurationMilliseconds = this.elapsedSeconds * 1000#
End Property

Public Property Get IsRunning() As Boolean
    IsRunning = this.running
End Property

End Class
