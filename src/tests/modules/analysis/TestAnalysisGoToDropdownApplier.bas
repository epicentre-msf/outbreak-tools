Attribute VB_Name = "TestAnalysisGoToDropdownApplier"
Attribute VB_Description = "Unit tests for AnalysisGoToDropdownApplier"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests verifying dropdown applier wiring")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TEST_WORKSHEET As String = "GoToApplierTest"

Private Assert As ICustomTest
Private HostSheet As Worksheet

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisGoToDropdownApplier"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet TEST_WORKSHEET
    RestoreApp
    Set Assert = Nothing
    Set HostSheet = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set HostSheet = EnsureWorksheet(TEST_WORKSHEET)
    ClearWorksheet HostSheet
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    ClearWorksheet HostSheet
End Sub

Private Function BuildRequest() As IAnalysisGoToDropdownRequest
    Dim builder As AnalysisGoToEntryCollectionBuilder
    Dim factory As AnalysisGoToDropdownFactory

    Set builder = New AnalysisGoToEntryCollectionBuilder
    builder.AddSection "Section A", "sec: ", "tableA"

    Set factory = New AnalysisGoToDropdownFactory
    Set BuildRequest = factory.CreateRequest("section", builder.SectionEntries, "ua_", "", "GoTo Section")
End Function

'@TestMethod("AnalysisGoToDropdownApplier")
Public Sub TestApplyWritesDropdownAndFormatsCell()
    CustomTestSetTitles Assert, "AnalysisGoToDropdownApplier", "TestApplyWritesDropdownAndFormatsCell"

    Dim request As IAnalysisGoToDropdownRequest
    Dim dropdowns As DropdownListsStub
    Dim formatter As LLFormatLogStub
    Dim applier As AnalysisGoToDropdownApplier
    Dim targetCell As Range

    Set request = BuildRequest()
    Set dropdowns = New DropdownListsStub
    dropdowns.Initialise HostSheet
    Set formatter = New LLFormatLogStub
    Set applier = New AnalysisGoToDropdownApplier
    Set targetCell = HostSheet.Range("B2")

    applier.Apply request, dropdowns, formatter, targetCell

    Assert.AreEqual request.ListName, dropdowns.LastAddedListName, "Dropdown list name should match request"
    Assert.AreEqual request.ListName, dropdowns.ValidationListName(1), "Validation should reference created list"
    Assert.AreEqual request.LabelText, CStr(targetCell.Value), "Cell value should match label text"
    Assert.AreEqual request.CellName, targetCell.Name.Name, "Cell name should match request"
    Assert.AreEqual 1&, formatter.AppliedScopes.Count, "Formatter should be invoked once"
End Sub
