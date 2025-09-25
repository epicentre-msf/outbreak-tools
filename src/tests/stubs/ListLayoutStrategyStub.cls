VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListLayoutStrategyStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub IListLayoutStrategy capturing invocation order")
'@IgnoreModule UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IListLayoutStrategy

Private Type TLayoutState
    initialiseCount As Long
    finaliseCount As Long
    invalidateCount As Long
    appliedSections As BetterArray
    raiseOnApplyIndex As Long
End Type

Private this As TLayoutState

Public Sub Class_Initialize()
    Set this.appliedSections = New BetterArray
    this.appliedSections.LowerBound = 1
End Sub

Public Property Get InitialiseCount() As Long
    InitialiseCount = this.initialiseCount
End Property

Public Property Get FinaliseCount() As Long
    FinaliseCount = this.finaliseCount
End Property

Public Property Get InvalidateCount() As Long
    InvalidateCount = this.invalidateCount
End Property

Public Property Get AppliedSections() As BetterArray
    Set AppliedSections = this.appliedSections.Clone()
End Property

Public Property Let RaiseOnApplyIndex(ByVal sectionIndex As Long)
    this.raiseOnApplyIndex = sectionIndex
End Property

'@section Interface implementation
'===============================================================================
Private Sub IListLayoutStrategy_Initialise(ByVal context As IListBuildContext)
    this.initialiseCount = this.initialiseCount + 1
End Sub

Private Sub IListLayoutStrategy_ApplySection(ByVal context As IListBuildContext, _
                                             ByVal section As IListSectionDescriptor)
    If this.raiseOnApplyIndex > 0 Then
        If this.appliedSections.Length + 1 = this.raiseOnApplyIndex Then
            Err.Raise CLng(ProjectError.ErrorUnexpectedState), "ListLayoutStrategyStub", _
                      "Simulated failure during ApplySection"
        End If
    End If

    this.appliedSections.Push section
End Sub

Private Sub IListLayoutStrategy_Finalise(ByVal context As IListBuildContext)
    this.finaliseCount = this.finaliseCount + 1
End Sub

Private Sub IListLayoutStrategy_Invalidate()
    this.invalidateCount = this.invalidateCount + 1
    Set this.appliedSections = New BetterArray
    this.appliedSections.LowerBound = 1
End Sub
