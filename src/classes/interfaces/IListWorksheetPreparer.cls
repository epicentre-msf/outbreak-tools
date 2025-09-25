VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListWorksheetPreparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Linelist.Shared")
'@ModuleDescription("Prepares and restores worksheet state before and after list building")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Lifecycle
'===============================================================================
'Acquire resources and toggle Excel state prior to building.
Public Sub Begin(ByVal context As IListBuildContext)
End Sub

'Restore state after a successful build.
Public Sub Complete(ByVal context As IListBuildContext)
End Sub

'Restore state after a failure.
Public Sub Abort(ByVal context As IListBuildContext, _
                 ByVal errNumber As Long, _
                 ByVal messageText As String)
End Sub

