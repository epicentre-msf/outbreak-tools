Attribute VB_Name = "TestAnalysisArrayUtilities"
Attribute VB_Description = "Unit tests for AnalysisArrayUtilities cloning helpers"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests exercising AnalysisArrayUtilities buffer and cloning behaviour")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisArrayUtilities"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Tests
'===============================================================================
'@TestMethod("AnalysisArrayUtilities")
Public Sub TestCreateBufferInitialisesLowerBound()
    CustomTestSetTitles Assert, "AnalysisArrayUtilities", "TestCreateBufferInitialisesLowerBound"
    Dim buffer As BetterArray

    Set buffer = AnalysisArrayUtilities.CreateBuffer()

    Assert.AreEqual 1&, buffer.LowerBound, "Buffer should default to lower bound of 1"
    Assert.AreEqual 0&, buffer.Length, "New buffer should start empty"
End Sub

'@TestMethod("AnalysisArrayUtilities")
Public Sub TestCloneItemsPerformsShallowCopy()
    CustomTestSetTitles Assert, "AnalysisArrayUtilities", "TestCloneItemsPerformsShallowCopy"
    Dim source As BetterArray
    Dim stub As CloneableItemStub
    Dim clone As BetterArray
    Dim clonedStub As CloneableItemStub

    Set source = AnalysisArrayUtilities.CreateBuffer()
    source.Push "alpha"

    Set stub = New CloneableItemStub
    stub.Configure "stub-1"
    source.Push stub

    Set clone = AnalysisArrayUtilities.CloneItems(source)

    Assert.AreEqual source.Length, clone.Length, "Shallow clone should preserve the element count"
    Assert.AreEqual "alpha", CStr(clone.Item(1)), "First element should match original value"

    Set clonedStub = clone.Item(2)
    Assert.IsTrue stub Is clonedStub, "Shallow clone should preserve object references"
End Sub

'@TestMethod("AnalysisArrayUtilities")
Public Sub TestCloneCloneableCreatesDistinctInstances()
    CustomTestSetTitles Assert, "AnalysisArrayUtilities", "TestCloneCloneableCreatesDistinctInstances"
    Dim source As BetterArray
    Dim first As CloneableItemStub
    Dim second As CloneableItemStub
    Dim clones As BetterArray
    Dim cloneOne As CloneableItemStub
    Dim cloneTwo As CloneableItemStub

    Set source = AnalysisArrayUtilities.CreateBuffer()

    Set first = New CloneableItemStub
    first.Configure "first"
    source.Push first

    Set second = New CloneableItemStub
    second.Configure "second"
    source.Push second

    Set clones = AnalysisArrayUtilities.CloneCloneable(source)

    Assert.AreEqual source.Length, clones.Length, "CloneCloneable should preserve element count"

    Set cloneOne = clones.Item(1)
    Set cloneTwo = clones.Item(2)

    Assert.IsFalse first Is cloneOne, "CloneCloneable should return new instances for cloneable objects"
    Assert.IsFalse second Is cloneTwo, "CloneCloneable should return new instances for each element"

    Assert.AreEqual first.Identifier, cloneOne.Identifier, "Clone should carry over identifier values"
    Assert.AreEqual second.Identifier, cloneTwo.Identifier, "Clone should preserve metadata for each element"
End Sub

'@TestMethod("AnalysisArrayUtilities")
Public Sub TestCloneCloneableRaisesWhenCloneMissing()
    CustomTestSetTitles Assert, "AnalysisArrayUtilities", "TestCloneCloneableRaisesWhenCloneMissing"
    Dim badSource As BetterArray
    Dim badObject As Collection
    Dim raisedError As Boolean

    Set badSource = AnalysisArrayUtilities.CreateBuffer()
    Set badObject = New Collection
    badObject.Add "value"
    badSource.Push badObject

    On Error Resume Next
        AnalysisArrayUtilities.CloneCloneable badSource
        raisedError = (Err.Number = ProjectError.ErrorUnexpectedState)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "CloneCloneable should raise when items do not expose a Clone method"
End Sub
