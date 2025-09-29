VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBuilderCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Linelist.Shared")
'@ModuleDescription("Coordinates worksheet preparation, section building, and layout strategies for linelist outputs")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IListBuilderCoordinator

Private Type TCoordinatorState
    layout As IListLayoutStrategy
    sectionBuilder As IListSectionBuilder
    preparer As IListWorksheetPreparer
End Type

Private this As TCoordinatorState
Private Const CLASSNAME As String = "ListBuilderCoordinator"

'@section Factory
'===============================================================================
'Create a coordinator with the supplied collaborators.
Public Function Create(ByVal layout As IListLayoutStrategy, _
                       ByVal sectionBuilder As IListSectionBuilder, _
                       ByVal preparer As IListWorksheetPreparer) As IListBuilderCoordinator

    ValidateLayout layout
    ValidateSectionBuilder sectionBuilder
    ValidatePreparer preparer

    With New ListBuilderCoordinator
        Set .LayoutStrategy = layout
        Set .SectionBuilder = sectionBuilder
        Set .WorksheetPreparer = preparer
        Set Create = .Self
    End With
End Function

'Expose the instance as its interface type.
Public Property Get Self() As IListBuilderCoordinator
    Set Self = Me
End Property

'Internal helpers used by the factory to initialise state.
Friend Property Set LayoutStrategy(ByVal layout As IListLayoutStrategy)
    ValidateLayout layout
    Set this.layout = layout
End Property

Friend Property Get LayoutStrategy() As IListLayoutStrategy
    Set LayoutStrategy = this.layout
End Property

Friend Property Set SectionBuilder(ByVal sectionBuilder As IListSectionBuilder)
    ValidateSectionBuilder sectionBuilder
    Set this.sectionBuilder = sectionBuilder
End Property

Friend Property Set WorksheetPreparer(ByVal preparer As IListWorksheetPreparer)
    ValidatePreparer preparer
    Set this.preparer = preparer
End Property

'@section Validation helpers
'===============================================================================
Private Sub ValidateContext(ByVal context As IListBuildContext)
    If context Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Build context cannot be Nothing"
    End If

    If Trim$(context.WorksheetName) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Worksheet name cannot be empty when building"
    End If
End Sub

Private Sub ValidateLayout(ByVal layout As IListLayoutStrategy)
    If layout Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "layout cannot be Nothing"
    End If

    If NormalizeValue(TypeName(layout)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "layout must expose a concrete implementation"
    End If
End Sub

Private Sub ValidateSectionBuilder(ByVal sectionBuilder As IListSectionBuilder)
    If sectionBuilder Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "sectionBuilder cannot be Nothing"
    End If

    If NormalizeValue(TypeName(sectionBuilder)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "sectionBuilder must expose a concrete implementation"
    End If
End Sub

Private Sub ValidatePreparer(ByVal preparer As IListWorksheetPreparer)
    If preparer Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "preparer cannot be Nothing"
    End If

    If NormalizeValue(TypeName(preparer)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "preparer must expose a concrete implementation"
    End If
End Sub

Private Sub ValidateSectionDescriptor(ByVal section As IListSectionDescriptor)
    If section Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Section descriptor cannot be Nothing"
    End If

    If NormalizeValue(TypeName(section)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Section descriptor must expose a concrete implementation"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

Private Function NormalizeValue(ByVal valueText As String) As String
    NormalizeValue = LCase$(Trim$(valueText))
End Function

'@section Core orchestration
'===============================================================================
'Execute the build pipeline returning the sections processed.
Public Function Build(ByVal context As IListBuildContext) As BetterArray
    ValidateContext context

    EnsureDependenciesReady

    On Error GoTo BuildFail

    this.preparer.Begin context
    this.layout.Initialise context

    Dim sections As BetterArray
    Set sections = this.sectionBuilder.BuildSections(context)

    Dim processed As BetterArray
    Set processed = New BetterArray
    processed.LowerBound = 1

    If sections Is Nothing Then
        ThrowError CLng(ProjectError.ElementNotFound), _
                  "Section builder returned Nothing"
    End If

    If sections.Length = 0 Then
        ThrowError CLng(ProjectError.ElementNotFound), _
                  "Section builder returned no sections"
    End If

    Dim idx As Long
    Dim section As IListSectionDescriptor

    For idx = sections.LowerBound To sections.UpperBound
        Set section = sections.Item(idx)
        ValidateSectionDescriptor section

        this.layout.ApplySection context, section
        processed.Push section
    Next idx

    this.layout.Finalise context
    this.preparer.Complete context

    Set Build = processed.Clone()
    Exit Function

BuildFail:
    On Error Resume Next
        If Not this.preparer Is Nothing Then this.preparer.Abort context, Err.Number, Err.Description
    On Error GoTo 0

    ThrowError CLng(ProjectError.ErrorUnexpectedState), _
              "List builder failed: " & Err.Description
End Function

Private Sub EnsureDependenciesReady()
    ValidateLayout this.layout
    ValidateSectionBuilder this.sectionBuilder
    ValidatePreparer this.preparer
End Sub

'@section Interface implementation
'===============================================================================
Private Function IListBuilderCoordinator_Create(ByVal layout As IListLayoutStrategy, _
                                                ByVal sectionBuilder As IListSectionBuilder, _
                                                ByVal preparer As IListWorksheetPreparer) As IListBuilderCoordinator
    Set IListBuilderCoordinator_Create = Create(layout, sectionBuilder, preparer)
End Function

Private Property Get IListBuilderCoordinator_Self() As IListBuilderCoordinator
    Set IListBuilderCoordinator_Self = Self
End Property

Private Function IListBuilderCoordinator_Build(ByVal context As IListBuildContext) As BetterArray
    Set IListBuilderCoordinator_Build = Build(context)
End Function

Private Sub IListBuilderCoordinator_Invalidate()
    On Error Resume Next
        If Not this.layout Is Nothing Then this.layout.Invalidate
        If Not this.sectionBuilder Is Nothing Then this.sectionBuilder.Invalidate
    On Error GoTo 0
End Sub
