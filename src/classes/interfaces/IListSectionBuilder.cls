VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListSectionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Linelist.Shared")
'@ModuleDescription("Produces section descriptors consumed by linelist builders")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Factory
'===============================================================================
'Generate the list of sections that should be processed for the supplied context.
Public Function BuildSections(ByVal context As IListBuildContext) As BetterArray
End Function

'Reset any cached state to honour dictionary or layout changes.
Public Sub Invalidate()
End Sub

