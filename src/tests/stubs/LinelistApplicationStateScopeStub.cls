VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationStateStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stubbed application state scope for pipeline tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements IApplicationState

Private Type TScope
    applyCount As Long
    restoreCount As Long
    refreshCount As Long
    shouldFailApply As Boolean
    shouldFailRestore As Boolean
    application As Application
End Type

Private this As TScope

Public Property Get ApplyCount() As Long
    ApplyCount = this.applyCount
End Property

Public Property Get RestoreCount() As Long
    RestoreCount = this.restoreCount
End Property

Public Property Get RefreshCount() As Long
    RefreshCount = this.refreshCount
End Property

Public Property Let ShouldFailApply(ByVal value As Boolean)
    this.shouldFailApply = value
End Property

Public Property Let ShouldFailRestore(ByVal value As Boolean)
    this.shouldFailRestore = value
End Property

Public Property Get ApplicationObject() As Application
    Set ApplicationObject = this.application
End Property

Public Property Set ApplicationObject(ByVal value As Application)
    Set this.application = value
End Property

Private Sub IApplicationState_ApplyBusyState()
    If this.shouldFailApply Then
        Err.Raise vbObjectError + 321, "ApplyBusyState", "Apply failure"
    End If
    this.applyCount = this.applyCount + 1
End Sub

Private Sub IApplicationState_Restore()
    If this.shouldFailRestore Then
        Err.Raise vbObjectError + 322, "Restore", "Restore failure"
    End If
    this.restoreCount = this.restoreCount + 1
End Sub

Private Sub IApplicationState_RefreshSnapshot()
    this.refreshCount = this.refreshCount + 1
End Sub

Private Property Get IApplicationState_ApplicationObject() As Application
    Set IApplicationState_ApplicationObject = this.application
End Property

Private Property Get IApplicationState_IsBusy() As Boolean
    IApplicationState_IsBusy = (this.applyCount > this.restoreCount)
End Property

Private Property Get IApplicationState_HasSnapshot() As Boolean
    IApplicationState_HasSnapshot = True
End Property
