VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseApplicationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface guarding and restoring Excel Application state for disease workflows"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface guarding and restoring Excel Application state for disease workflows")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Commands
'===============================================================================

'@sub-title Snapshot the Application state then disable interactive toggles.
'@details Captures ScreenUpdating, DisplayAlerts, EnableEvents and Calculation,
'         then disables them (Calculation set to manual) so a disease workflow
'         can run silently. Re-entrant calls keep the originally captured state.
'@param app Application whose state is captured and guarded.
'@param disableStatusBar Optional Boolean; when True the status bar is also hidden (default True).
Public Sub BeginGuard(ByVal app As Application, Optional ByVal disableStatusBar As Boolean = True)
End Sub

'@sub-title Restore the Application state captured by BeginGuard.
'@details Safe to call repeatedly: a Restore with nothing guarded is a no-op and
'         never raises. ScreenUpdating is restored last so the UI repaints once.
Public Sub Restore()
End Sub
