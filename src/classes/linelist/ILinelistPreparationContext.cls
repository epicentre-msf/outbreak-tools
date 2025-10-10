VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistPreparationContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Runtime context shared across Linelist preparation steps"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Runtime context shared across Linelist preparation steps")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@section Core dependencies
'===============================================================================
Public Property Get Linelist() As ILinelist
End Property

Public Property Get Specifications() As ILinelistSpecs
End Property

Public Property Get Dictionary() As ILLdictionary
End Property

Public Property Get SheetFormatter() As ILinelistSheetNameFormatter
End Property

Public Property Get StateScope() As IApplicationState
End Property
