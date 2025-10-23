Attribute VB_Name = "TestShowHideExport"

Option Explicit

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private FixtureSheet As Worksheet
Private FixtureTable As ListObject
Private FixtureCustomTable As ICustomTable
Private ExportSubject As IShowHideExport

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ShowHideExport"


'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestShowHideExport"
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
    Set FixtureSheet = FixtureWorkbook.Worksheets(1)
    InitialiseTable FixtureSheet
    Set FixtureCustomTable = CustomTable.Create(FixtureTable)
    Set ExportSubject = ShowHideExport.Create(FixtureCustomTable)
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

    Set FixtureWorkbook = Nothing
    Set FixtureSheet = Nothing
    Set FixtureTable = Nothing
    Set FixtureCustomTable = Nothing
    Set ExportSubject = Nothing
End Sub


'@TestMethod("ShowHide")
Public Sub TestExportPlanWritesRows()
    CustomTestSetTitles Assert, TESTMODULE, "TestExportPlanWritesRows"

    Dim plan As IShowHidePlan
    Dim lo As ListObject

    Set plan = ShowHidePlan.Create(ShowHideLayerHList)
    plan.AddVisibility "field_a", "Field A", True
    plan.AddVisibility "field_b", "Field B", False

    ExportSubject.ExportPlan plan

    Set lo = FixtureTable

    Assert.AreEqual 2, lo.ListRows.Count, _
                     "ExportPlan should append one row per action"
    Assert.AreEqual "hlist", CStr(lo.DataBodyRange.Cells(1, 1).Value), _
                     "Layer token should be stored in the first column"
    Assert.AreEqual "field_b", CStr(lo.DataBodyRange.Cells(2, 2).Value), _
                     "Field key should be persisted in the second column"
    Assert.AreEqual "false", CStr(lo.DataBodyRange.Cells(2, 4).Value), _
                     "Hidden flag should be stored as a boolean string"
End Sub


'@TestMethod("ShowHide")
Public Sub TestImportPlanReadsRows()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportPlanReadsRows"

    Dim planA As IShowHidePlan
    Dim planB As IShowHidePlan
    Dim imported As IShowHidePlan
    Dim actions As BetterArray

    Set planA = ShowHidePlan.Create(ShowHideLayerHList)
    planA.AddVisibility "field_a", "Field A", False
    ExportSubject.ExportPlan planA

    Set planB = ShowHidePlan.Create(ShowHideLayerPrinted)
    planB.AddVisibility "field_p", "Field Printed", True
    ExportSubject.ExportPlan planB

    Set imported = ExportSubject.ImportPlan(ShowHideLayerPrinted)
    Set actions = imported.Actions

    Assert.AreEqual ShowHideLayerPrinted, imported.TargetLayer, _
                     "Imported plan should target the requested layer"
    Assert.AreEqual 1, actions.Length, _
                     "ImportPlan should return only rows matching the target layer"
    Assert.IsTrue CBool(actions.Item(actions.LowerBound)(2)), _
                  "Hidden flag should be converted back to Boolean"
End Sub


'@TestMethod("ShowHide")
Public Sub TestClearLayerRemovesRows()
    CustomTestSetTitles Assert, TESTMODULE, "TestClearLayerRemovesRows"

    Dim planA As IShowHidePlan
    Dim planB As IShowHidePlan
    Dim lo As ListObject

    Set planA = ShowHidePlan.Create(ShowHideLayerHList)
    planA.AddVisibility "field_a", "Field A", False
    planA.AddVisibility "field_b", "Field B", True
    ExportSubject.ExportPlan planA

    Set planB = ShowHidePlan.Create(ShowHideLayerCRF)
    planB.AddVisibility "field_c", "Field C", True
    ExportSubject.ExportPlan planB

    ExportSubject.ClearLayer ShowHideLayerHList
    Set lo = FixtureTable

    Assert.AreEqual 1, lo.ListRows.Count, _
                     "ClearLayer should delete only the rows matching the layer token"
    Assert.AreEqual "crf", CStr(lo.DataBodyRange.Cells(1, 1).Value), _
                     "Remaining rows should correspond to the other layers"
End Sub


'@section Helpers
'===============================================================================

Private Sub InitialiseTable(ByVal targetSheet As Worksheet)
    Dim headerValues As Variant
    Dim headerRange As Range

    headerValues = TestHelpers.RowsToMatrix(Array(Array("layer", "field_key", "header_text", "hidden_flag")))
    targetSheet.Cells.Clear

    Set headerRange = targetSheet.Range("A1").Resize(1, UBound(headerValues, 2))
    headerRange.Value = headerValues

    Set FixtureTable = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                   Source:=headerRange, _
                                                   XlListObjectHasHeaders:=xlYes)
    FixtureTable.Name = "tbl_showhide_state"
End Sub
