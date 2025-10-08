VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistPreparationStepStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Records execution order for Linelist preparation pipeline tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

Implements ILinelistPreparationStep

Private Type TState
    label As String
    callLog As BetterArray
    shouldFail As Boolean
    failMessage As String
End Type

Private this As TState

Public Sub Initialise(ByVal label As String, ByVal log As BetterArray)
    this.label = label
    Set this.callLog = log
    this.shouldFail = False
    this.failMessage = vbNullString
End Sub

Public Property Let ShouldFail(ByVal value As Boolean)
    this.shouldFail = value
End Property

Public Property Let FailMessage(ByVal value As String)
    this.failMessage = value
End Property

Private Sub ILinelistPreparationStep_Execute(ByVal context As ILinelistPreparationContext)
    If Not this.callLog Is Nothing Then
        this.callLog.Push this.label
    End If

    If this.shouldFail Then
        Err.Raise vbObjectError + 123, this.label, _
                  IIf(Len(this.failMessage) = 0, "Step failure", this.failMessage)
    End If
End Sub

