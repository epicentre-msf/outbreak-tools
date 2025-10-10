VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistCodeTransferPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Describes the VBA components that must be transferred into the linelist workbook"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Describes the VBA components that must be transferred into the linelist workbook")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@section Component collections
'===============================================================================
'@sub-title Returns the list of class modules that must be cloned into the target workbook.
Public Function ClassComponentNames() As BetterArray
End Function

'@sub-title Returns the list of standard modules that must be cloned into the target workbook.
Public Function ModuleComponentNames() As BetterArray
End Function

'@sub-title Returns the userform names that must exist in the target workbook.
Public Function FormNames() As BetterArray
End Function

'@section Workbook module metadata
'===============================================================================
'@sub-title Name of the workbook module whose code must be synchronised.
Public Property Get WorkbookModuleName() As String
End Property
