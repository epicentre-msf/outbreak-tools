VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistWorkbookAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Exposes the workbook currently being assembled by the linelist"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Exposes the workbook currently being assembled by the linelist")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, SuperfluousAnnotationArgument

Option Explicit

'@section Workbook access
'===============================================================================
'@sub-title Returns the in-flight output workbook used by the linelist builder.
Public Property Get OutputWorkbook() As Workbook
End Property
