VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseApplicationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing temporary Application state guards for disease workflows"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface describing temporary Application state guards for disease workflows")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Lifecycle
'===============================================================================

'@sub-title Capture current Application state and apply the optimised settings
'@param excelApp Application reference to guard.
'@param disableStatusBar Optional Boolean to clear the StatusBar during the busy window.
Public Sub BeginGuard(ByVal excelApp As Application, Optional ByVal disableStatusBar As Boolean = True)
End Sub

'@sub-title Restore the Application to its original state.
'@remarks This method is safe to call multiple times; subsequent calls will be ignored.
Public Sub Restore()
End Sub

'@section State
'===============================================================================

'@sub-title Indicates whether optimised settings are active.
'@return Boolean True when BeginGuard successfully captured and applied the guard.
Public Property Get Active() As Boolean
End Property

'@sub-title Retrieve the guarded Application reference.
'@return Application reference supplied to BeginGuard or Nothing if not initialised.
Public Property Get HostApplication() As Application
End Property
