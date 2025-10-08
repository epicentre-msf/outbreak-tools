VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistLifecycleManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Coordinates reset and disposal operations for the linelist output workbook"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Coordinates reset and disposal operations for the linelist output workbook")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, SuperfluousAnnotationArgument

Option Explicit

'@section Lifecycle operations
'===============================================================================
'@sub-title Clears workbook state and temporary artefacts so a new preparation run can begin.
Public Sub Reset(Optional ByVal saveChanges As Boolean = False)
End Sub

'@sub-title Performs a terminal shutdown, releasing resources and clearing references.
Public Sub Dispose(Optional ByVal saveChanges As Boolean = False)
End Sub
