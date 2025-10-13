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
'@Folder("Linelist")
'@ModuleDescription("Application state guard for linelist busy operations")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@section Busy state operations
'===============================================================================
'@sub-title Transition to busy state
'@details Applies the busy settings (ScreenUpdating False, DisplayAlerts False,
'@details Calculation Manual, EnableAnimations False when available) over the
'@details captured application snapshot.
'@param busyCursor Optional cursor to display while busy; defaults to xlNorthWestArrow.
Public Sub ApplyBusyState(Optional ByVal suppressEvent As Boolean = False, Optional ByVal busyCursor As XlMousePointer = xlNorthWestArrow)
End Sub

'@sub-title Restore to original state
'@details Restores the captured application settings, even when ApplyBusyState
'@details was never called.
Public Sub Restore()
End Sub

'@sub-title Refresh snapshot
'@details Refreshes the stored snapshot with the application's current state so
'@details subsequent Restore invocations return to the freshly captured values.
Public Sub RefreshSnapshot()
End Sub

'@section Introspection
'===============================================================================
'@sub-title Application reference
'@return Application currently guarded by the scope.
Public Property Get ApplicationObject() As Application
End Property

'@sub-title Indicates whether ApplyBusyState was invoked and has not yet been
'@details balanced by Restore.
Public Property Get IsBusy() As Boolean
End Property

'@sub-title Indicates whether a snapshot exists and the scope is ready to
'@details restore values.
Public Property Get HasSnapshot() As Boolean
End Property
