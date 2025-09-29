Attribute VB_Name = "TestListLayoutStrategies"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Smoke tests for horizontal and vertical list layout strategies")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private Section As IListSectionDescriptor

'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Section = ListSectionDescriptor.Create("section-one", 5, 2)
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Section = Nothing
    Set Assert = Nothing
End Sub

'@section Tests
'===============================================================================
'@TestMethod("ListLayoutStrategies")
Private Sub TestHorizontalStrategyRecordsSections()
    Dim strategy As HorizontalListLayoutStrategy
    Dim context As ListBuildContextStub

    Set strategy = New HorizontalListLayoutStrategy
    Set context = New ListBuildContextStub
    context.Configure "Sheet_A", CByte(2)

    strategy.Initialise context
    strategy.ApplySection context, Section
    strategy.Finalise context

    Assert.AreEqual 1&, strategy.ProcessedSections.Length
    Assert.AreEqual 1&, strategy.InitialiseCount
    Assert.AreEqual 1&, strategy.FinaliseCount
End Sub

'@TestMethod("ListLayoutStrategies")
Private Sub TestVerticalStrategyRecordsSections()
    Dim strategy As VerticalListLayoutStrategy
    Dim context As ListBuildContextStub

    Set strategy = New VerticalListLayoutStrategy
    Set context = New ListBuildContextStub
    context.Configure "Sheet_B", CByte(3)

    strategy.Initialise context
    strategy.ApplySection context, Section
    strategy.Finalise context

    Assert.AreEqual 1&, strategy.ProcessedSections.Length
    Assert.AreEqual 1&, strategy.InitialiseCount
    Assert.AreEqual 1&, strategy.FinaliseCount
End Sub

