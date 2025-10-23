Attribute VB_Name = "TestShowHideOrchestrator"

Option Explicit

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private DictionarySheet As Worksheet
Private Dict As ILLdictionary
Private Variables As ILLVariables
Private Orchestrator As IShowHideOrchestrator
Private ExportService As IShowHideExport
Private ExportTable As ListObject
Private ExportCustomTable As ICustomTable

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ShowHideOrchestrator"
Private Const DICTIONARY_SHEET As String = "DictionaryFixture"
Private Const PLAN_TABLE_NAME As String = "tbl_showhide_state"


'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestShowHideOrchestrator"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, FixtureWorkbook
    Set DictionarySheet = FixtureWorkbook.Worksheets(DICTIONARY_SHEET)
    Set Dict = LLdictionary.Create(DictionarySheet, 1, 1)
    Dict.Prepare
    Set Variables = LLVariables.Create(Dict)
    InitialiseExportTable FixtureWorkbook.Worksheets.Add
    Set ExportCustomTable = CustomTable.Create(ExportTable)
    Set ExportService = ShowHideExport.Create(ExportCustomTable)
    Set Orchestrator = ShowHideOrchestrator.Create(Dict)
    RegisterStrategies
    Assert.BeginTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then
            TestHelpers.DeleteWorkbook FixtureWorkbook
        End If
    On Error GoTo 0

    Set ExportCustomTable = Nothing
    Set ExportTable = Nothing
    Set ExportService = Nothing
    Set Orchestrator = Nothing
    Set Variables = Nothing
    Set Dict = Nothing
    Set DictionarySheet = Nothing
    Set FixtureWorkbook = Nothing
End Sub


'@TestMethod("ShowHide")
Public Sub TestApplyHListExportsPlan()
    CustomTestSetTitles Assert, TESTMODULE, "TestApplyHListExportsPlan"

    Dim plan As IShowHidePlan

    Set plan = Orchestrator.Apply("hlist2D-sheet1", ShowHideLayerHList, ExportService, ExportCustomTable)

    Assert.IsNotNothing plan, "Orchestrator should return a plan instance"
    Assert.IsTrue ExportTable.ListRows.Count > 0, _
                  "Export table should receive rows for the evaluated plan"

    Dim hiddenValue As String
    hiddenValue = PlanHiddenValue("hid_end_h2")

    Assert.AreEqual "true", hiddenValue, _
                    "Hidden variables in the dictionary should remain hidden in the exported plan"
End Sub


'@TestMethod("ShowHide")
Public Sub TestApplyRespectsPersistedOverrides()
    CustomTestSetTitles Assert, TESTMODULE, "TestApplyRespectsPersistedOverrides"

    'Initial export populates defaults
    Orchestrator.Apply "hlist2D-sheet1", ShowHideLayerHList, ExportService, ExportCustomTable

    'Override a hidden variable to visible
    UpdatePlanRow "hid_end_h2", "false"

    'Re-run orchestration to ensure persisted plan is respected
    Orchestrator.Apply "hlist2D-sheet1", ShowHideLayerHList, ExportService, ExportCustomTable

    Assert.AreEqual "false", PlanHiddenValue("hid_end_h2"), _
                    "Persisted plan should override the dictionary default visibility"
End Sub


'@section Helpers
'===============================================================================

Private Sub RegisterStrategies()
    Orchestrator.RegisterStrategy ShowHideHListStrategy.Create
    Orchestrator.RegisterStrategy ShowHideVListStrategy.Create
    Orchestrator.RegisterStrategy ShowHideCRFStrategy.Create
    Orchestrator.RegisterStrategy ShowHidePrintedStrategy.Create
    Orchestrator.RegisterStrategy ShowHideDictionaryStrategy.Create
End Sub

Private Sub InitialiseExportTable(ByVal targetSheet As Worksheet)
    Dim headers As Variant
    Dim headerRange As Range

    headers = TestHelpers.RowsToMatrix(Array(Array("layer", "field_key", "header_text", "hidden_flag")))
    targetSheet.Cells.Clear

    Set headerRange = targetSheet.Range("A1").Resize(1, UBound(headers, 2))
    headerRange.Value = headers

    Set ExportTable = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                  Source:=headerRange, _
                                                  XlListObjectHasHeaders:=xlYes)
    ExportTable.Name = PLAN_TABLE_NAME
End Sub

Private Function PlanHiddenValue(ByVal fieldKey As String) As String
    Dim row As ListRow

    For Each row In ExportTable.ListRows
        If StrComp(CStr(row.Range.Cells(1, 2).Value), fieldKey, vbTextCompare) = 0 Then
            PlanHiddenValue = CStr(row.Range.Cells(1, 4).Value)
            Exit Function
        End If
    Next row

    PlanHiddenValue = vbNullString
End Function

Private Sub UpdatePlanRow(ByVal fieldKey As String, ByVal hiddenFlag As String)
    Dim row As ListRow

    For Each row In ExportTable.ListRows
        If StrComp(CStr(row.Range.Cells(1, 2).Value), fieldKey, vbTextCompare) = 0 Then
            row.Range.Cells(1, 4).Value = hiddenFlag
            Exit Sub
        End If
    Next row
End Sub
