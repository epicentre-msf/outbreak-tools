VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListLayoutStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Linelist.Shared")
'@ModuleDescription("Applies orientation-specific layout logic during list building")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Lifecycle
'===============================================================================
'Configure strategy state ahead of section processing.
Public Sub Initialise(ByVal context As IListBuildContext)
End Sub

'Apply layout operations for a single section descriptor.
Public Sub ApplySection(ByVal context As IListBuildContext, _
                        ByVal section As IListSectionDescriptor)
End Sub

'Perform any required tidy-up once all sections have been processed.
Public Sub Finalise(ByVal context As IListBuildContext)
End Sub

'Clear cached state when upstream dependencies change.
Public Sub Invalidate()
End Sub

