Attribute VB_Name = "TestAnalysisGoToEntryCollectionBuilder"
Attribute VB_Description = "Unit tests for AnalysisGoToEntryCollectionBuilder"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests verifying AnalysisGoToEntryCollectionBuilder behaviour")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisGoToEntryCollectionBuilder"
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

'@TestMethod("AnalysisGoToEntryCollectionBuilder")
Public Sub TestBuilderCreatesTypedEntries()
    CustomTestSetTitles Assert, "AnalysisGoToEntryCollectionBuilder", "TestBuilderCreatesTypedEntries"
    Dim builder As AnalysisGoToEntryCollectionBuilder
    Dim entry As IAnalysisGoToEntry

    Set builder = New AnalysisGoToEntryCollectionBuilder
    builder.AddSection "Section A", "sec: ", "tableA"
    builder.AddHeader "Header A", "hdr: "
    builder.AddGraph "Graph A", "gr: "

    Dim manualEntry As IAnalysisGoToEntry
    Set manualEntry = AnalysisGoToEntry.Create("section", "Section B", "sec: ")
    builder.AddEntry manualEntry

    Assert.AreEqual 2&, builder.SectionEntries.Length, "Section entries should contain two items"
    Assert.AreEqual 1&, builder.HeaderEntries.Length, "Header entries should contain one item"
    Assert.AreEqual 1&, builder.GraphEntries.Length, "Graph entries should contain one item"

    Set entry = builder.SectionEntries.Item(1)
    Assert.AreEqual "section", entry.Scope, "Section entry should expose scope metadata"
    Assert.AreEqual "sec: Section A", entry.DisplayText, "Section entry display text should include prefix"
    Assert.AreEqual "tableA", entry.Suffix, "Section entry should expose suffix metadata"

    Set entry = builder.SectionEntries.Item(2)
    Assert.AreEqual "Section B", entry.Label, "Manual entry should be captured"
End Sub
