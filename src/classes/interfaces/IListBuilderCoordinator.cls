VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListBuilderCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Linelist.Shared")
'@ModuleDescription("Coordinates layout strategies, section builders, and worksheet preparation")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Factory
'===============================================================================
'Create a coordinator from the supplied collaborators.
Public Function Create(ByVal layout As IListLayoutStrategy, _
                       ByVal sectionBuilder As IListSectionBuilder, _
                       ByVal preparer As IListWorksheetPreparer) As IListBuilderCoordinator
End Function

'Expose the instance as the interface type.
Public Property Get Self() As IListBuilderCoordinator
End Property

'@section Operations
'===============================================================================
'Execute the build pipeline and return the processed section descriptors.
Public Function Build(ByVal context As IListBuildContext) As BetterArray
End Function

'Invalidate cached collaborators after upstream data changes.
Public Sub Invalidate()
End Sub

