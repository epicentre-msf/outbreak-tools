Attribute VB_Name = "TestListBuilderCoordinator"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Unit tests covering the ListBuilderCoordinator orchestration")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private Coordinator As IListBuilderCoordinator
Private LayoutStub As ListLayoutStrategyStub
Private SectionBuilderStub As ListSectionBuilderStub
Private PreparerStub As ListWorksheetPreparerStub
Private ContextStub As ListBuildContextStub

'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@section Test lifecycle
'===============================================================================
'@TestInitialize
Private Sub TestInitialize()
    Set LayoutStub = New ListLayoutStrategyStub
    Set SectionBuilderStub = New ListSectionBuilderStub
    Set PreparerStub = New ListWorksheetPreparerStub
    Set ContextStub = New ListBuildContextStub
    ContextStub.Configure "Sheet_A", CByte(2)
    Set Coordinator = ListBuilderCoordinator.Create(LayoutStub, SectionBuilderStub, PreparerStub)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Coordinator = Nothing
    Set LayoutStub = Nothing
    Set SectionBuilderStub = Nothing
    Set PreparerStub = Nothing
    Set ContextStub = Nothing
End Sub

'@section Helpers
'===============================================================================
Private Function SectionList(ParamArray sectionNames() As Variant) As BetterArray
    Dim list As BetterArray
    Dim idx As Long

    Set list = New BetterArray
    list.LowerBound = 1

    For idx = LBound(sectionNames) To UBound(sectionNames)
        list.Push ListSectionDescriptor.Create(Trim$(CStr(sectionNames(idx))), 10 + idx, 3)
    Next idx

    Set SectionList = list
End Function

'@section Tests
'===============================================================================
'@TestMethod("ListBuilderCoordinator")
Private Sub TestBuildCoordinatesCollaborators()
    Dim result As BetterArray

    SectionBuilderStub.SetSections SectionList("sec_one", "sec_two")

    Set result = Coordinator.Build(ContextStub)

    Assert.AreEqual 1&, PreparerStub.BeginCount, "Preparer.Begin should be called once"
    Assert.AreEqual 1&, PreparerStub.CompleteCount, "Preparer.Complete should be called once"
    Assert.AreEqual 0&, PreparerStub.AbortCount, "Preparer.Abort should not be called on success"
    Assert.AreEqual 1&, LayoutStub.InitialiseCount, "Layout.Initialise should be called once"
    Assert.AreEqual 1&, LayoutStub.FinaliseCount, "Layout.Finalise should be called once"
    Assert.AreEqual 2&, LayoutStub.AppliedSections.Length, "ApplySection should be called for each section"
    Assert.AreEqual 2&, result.Length, "Build should return the processed sections"
End Sub

'@TestMethod("ListBuilderCoordinator")
Private Sub TestBuildRaisesWhenNoSectionsReturned()
    On Error GoTo ExpectError

    SectionBuilderStub.SetSections Nothing
    Coordinator.Build ContextStub
    Assert.Fail "Build should raise when no sections are returned"
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ElementNotFound), Err.Number
    Assert.AreEqual 1&, PreparerStub.AbortCount, "Abort should be called when build fails"
    Err.Clear
End Sub

'@TestMethod("ListBuilderCoordinator")
Private Sub TestBuildAbortWhenLayoutFails()
    On Error GoTo ExpectError

    LayoutStub.RaiseOnApplyIndex = 2
    SectionBuilderStub.SetSections SectionList("sec_one", "sec_two", "sec_three")

    Coordinator.Build ContextStub
    Assert.Fail "Build should raise when layout fails"
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), Err.Number
    Assert.AreEqual 1&, PreparerStub.AbortCount, "Abort should be invoked on failure"
    Err.Clear
End Sub

'@TestMethod("ListBuilderCoordinator")
Private Sub TestInvalidateDelegatesToCollaborators()
    Coordinator.Invalidate
    Assert.AreEqual 1&, LayoutStub.InvalidateCount, "Layout should be invalidated"
    Assert.AreEqual 1&, SectionBuilderStub.InvalidateCalls, "Section builder should be invalidated"
End Sub

