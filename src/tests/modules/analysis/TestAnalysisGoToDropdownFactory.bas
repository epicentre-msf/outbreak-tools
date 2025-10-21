Attribute VB_Name = "TestAnalysisGoToDropdownFactory"
Attribute VB_Description = "Unit tests for AnalysisGoToDropdownFactory"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests verifying dropdown request factory output")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisGoToDropdownFactory"
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

'@TestMethod("AnalysisGoToDropdownFactory")
Public Sub TestCreateRequestBuildsValues()
    CustomTestSetTitles Assert, "AnalysisGoToDropdownFactory", "TestCreateRequestBuildsValues"

    Dim builder As AnalysisGoToEntryCollectionBuilder
    Dim factory As AnalysisGoToDropdownFactory
    Dim request As IAnalysisGoToDropdownRequest

    Set builder = New AnalysisGoToEntryCollectionBuilder
    builder.AddSection "Section A", "sec: ", "tableA"
    builder.AddSection "Section B", "sec: ", "tableB"

    Set factory = New AnalysisGoToDropdownFactory
    Set request = factory.CreateRequest("section", builder.SectionEntries, "ua_", "_01", "GoTo")

    Assert.IsTrue Not request Is Nothing, "Factory should produce a request"
    Assert.IsTrue request.HasEntries, "Request should contain entries"
    Assert.AreEqual "ua_gotosection_01", request.ListName, "List name should include prefix and suffix"
    Assert.AreEqual "ua_go_to_section_01", request.CellName, "Cell name should include prefix and suffix"
    Assert.AreEqual 2&, request.Values.Length, "Two dropdown values expected"
    Assert.AreEqual "sec: Section A", CStr(request.Values.Item(1)), "Values should use display text"
End Sub
