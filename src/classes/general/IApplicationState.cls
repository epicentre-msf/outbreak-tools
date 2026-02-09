VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IApplicationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Application state guard for linelist busy operations"

'@Interface
'@Folder("General")
'@ModuleDescription("Application state guard for linelist busy operations")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@class IApplicationState
'@description
'Defines the public contract for the ApplicationState class, exposing
'busy-state toggling and snapshot capture of Excel Application settings.
'Callers apply a locked-down busy mode (no screen updates, manual calc,
'alerts off) and later restore the original state from the snapshot.


'@section Busy State
'===============================================================================
'@description Transition to and from the busy state.

'@jump:apply-busy-state
'@sub-title Transition to busy state
'@details
'Applies busy settings (ScreenUpdating False, DisplayAlerts False,
'Calculation Manual, EnableAnimations False when available) over the
'captured application snapshot.
'@param suppressEvents Boolean. When True, disables Application.EnableEvents. Defaults to False.
'@param calculateOnSave Boolean. Value for CalculateBeforeSave. Defaults to True.
'@param busyCursor XlMousePointer. Cursor shown while busy. Defaults to xlNorthWestArrow.
'@param blockSecurity Boolean. When True, forces AutomationSecurity to msoAutomationSecurityForceDisable. Defaults to False.
Public Sub ApplyBusyState(Optional ByVal suppressEvents As Boolean = False, _
                          Optional ByVal calculateOnSave As Boolean = True, _
                          Optional ByVal busyCursor As XlMousePointer = xlNorthWestArrow, _
                          Optional ByVal blockSecurity As Boolean = False)
End Sub

'@jump:restore
'@sub-title Restore the application to its pre-busy state
'@details
'Restores all captured application settings, even when ApplyBusyState
'was never called. Clears the busy flag.
Public Sub Restore()
End Sub

'@jump:refresh-snapshot
'@sub-title Refresh the stored snapshot with current application values
'@details
'Re-captures the application state so subsequent Restore invocations
'return to the freshly captured values. Cannot be called while busy.
Public Sub RefreshSnapshot()
End Sub


'@section Introspection
'===============================================================================
'@description Read-only state accessors.

'@jump:application-object
'@prop-title Application reference currently guarded
'@return Application. The Excel Application instance.
Public Property Get ApplicationObject() As Application
End Property

'@jump:is-busy
'@prop-title Whether the busy state is currently active
'@return Boolean. True when ApplyBusyState has been invoked without a matching Restore.
Public Property Get IsBusy() As Boolean
End Property

'@jump:has-snapshot
'@prop-title Whether a snapshot exists and restoration is possible
'@return Boolean. True when a snapshot has been captured.
Public Property Get HasSnapshot() As Boolean
End Property
