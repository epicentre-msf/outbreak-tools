VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListSectionBuilderStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub IListSectionBuilder returning preconfigured sections")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IListSectionBuilder

Private Type TBuilderState
    sections As BetterArray
    buildCalls As Long
    invalidateCalls As Long
End Type

Private this As TBuilderState

Public Sub SetSections(ByVal sectionList As BetterArray)
    If sectionList Is Nothing Then
        Set this.sections = Nothing
    Else
        Set this.sections = sectionList.Clone()
    End If
End Sub

Public Property Get BuildCalls() As Long
    BuildCalls = this.buildCalls
End Property

Public Property Get InvalidateCalls() As Long
    InvalidateCalls = this.invalidateCalls
End Property

'@section Interface implementation
'===============================================================================
Private Function IListSectionBuilder_BuildSections(ByVal context As IListBuildContext) As BetterArray
    this.buildCalls = this.buildCalls + 1

    If this.sections Is Nothing Then
        Set IListSectionBuilder_BuildSections = Nothing
    Else
        Set IListSectionBuilder_BuildSections = this.sections.Clone()
    End If
End Function

Private Sub IListSectionBuilder_Invalidate()
    this.invalidateCalls = this.invalidateCalls + 1
End Sub

