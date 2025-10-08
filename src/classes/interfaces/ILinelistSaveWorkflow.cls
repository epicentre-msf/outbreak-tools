VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistSaveWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Coordinates the final save sequence for the Linelist workbook"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Coordinates the final save sequence for the Linelist workbook")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, SuperfluousAnnotationArgument

Option Explicit

'@section Save operations
'===============================================================================
'@sub-title Executes the save process using the supplied context.
Public Sub Save(ByVal context As ILinelistPreparationContext)
End Sub

