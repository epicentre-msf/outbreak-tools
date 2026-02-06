VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for building HList or VList data entry sheets"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Contract for building HList or VList data entry sheets")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@description Layer modes for list building. HList is horizontal, VList is vertical.
Public Enum ListBuilderLayer
    ListBuilderLayerHList = 1
    ListBuilderLayerVList = 2
End Enum

'@section Core Methods
'===============================================================================

'@description Build the data entry sheet: format, freeze panes, sections, buttons, ListObjects, code transfer, protect.
Public Sub Build()
End Sub
