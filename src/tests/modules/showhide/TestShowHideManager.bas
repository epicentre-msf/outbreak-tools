Attribute VB_Name = "TestShowHideManager"
Attribute VB_Description = "Tests for ShowHideManager class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for ShowHideManager class")

Option Explicit

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private Dict As ILLdictionary

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ShowHideManager"
Private Const DICTIONARY_SHEET As String = "DictionaryFixture"


'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestShowHideManager"
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
    Set Dict = LLdictionary.Create(FixtureWorkbook.Worksheets(DICTIONARY_SHEET), 1, 1)
    Dict.Prepare
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

    Set Dict = Nothing
    Set FixtureWorkbook = Nothing
End Sub


'@section Build tests
'===============================================================================

'@TestMethod("ShowHide")
Public Sub TestBuildCreatesEntriesForVList()
    CustomTestSetTitles Assert, TESTMODULE, "TestBuildCreatesEntriesForVList"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")

    Assert.IsTrue sut.EntryCount > 0, _
                  "Manager should contain entries for vlist1D-sheet1"
    Assert.IsTrue sut.HasField("opt_vis_v1"), _
                  "Should contain opt_vis_v1 variable"
    Assert.IsTrue sut.HasField("mand_v1"), _
                  "Should contain mand_v1 variable"
    Assert.AreEqual ShowHideLayerVList, sut.TargetLayer, _
                     "TargetLayer should be VList"
    Assert.AreEqual "vlist1D-sheet1", sut.SheetName, _
                     "SheetName should match the provided value"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildCreatesEntriesForVList", Err.Number, Err.Description
End Sub


'@TestMethod("ShowHide")
Public Sub TestBuildCreatesEntriesForHList()
    CustomTestSetTitles Assert, TESTMODULE, "TestBuildCreatesEntriesForHList"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Set sut = ShowHideManager.Create(Dict, ShowHideLayerHList, "hlist2D-sheet1")

    Assert.IsTrue sut.EntryCount > 0, _
                  "Manager should contain entries for hlist2D-sheet1"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildCreatesEntriesForHList", Err.Number, Err.Description
End Sub


'@TestMethod("ShowHide")
Public Sub TestBuildReturnsEmptyForUnknownSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestBuildReturnsEmptyForUnknownSheet"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Set sut = ShowHideManager.Create(Dict, ShowHideLayerHList, "nonexistent_sheet")

    Assert.AreEqual CLng(0), sut.EntryCount, _
                     "Manager should have zero entries for a non-existent sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildReturnsEmptyForUnknownSheet", Err.Number, Err.Description
End Sub


'@section Mandatory tests
'===============================================================================

'@TestMethod("ShowHide")
Public Sub TestMandatoryFlaggedCorrectly()
    CustomTestSetTitles Assert, TESTMODULE, "TestMandatoryFlaggedCorrectly"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim idx As Long

    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")
    idx = sut.IndexOf("mand_v1")

    Assert.IsTrue idx > 0, "mand_v1 should exist in the manager"
    Assert.IsTrue sut.IsMandatory(idx), _
                  "mand_v1 should be flagged as mandatory"
    Assert.IsFalse sut.IsHidden(idx), _
                   "Mandatory entries must always return False from IsHidden"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestMandatoryFlaggedCorrectly", Err.Number, Err.Description
End Sub


'@TestMethod("ShowHide")
Public Sub TestSetHiddenIgnoresMandatory()
    CustomTestSetTitles Assert, TESTMODULE, "TestSetHiddenIgnoresMandatory"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim idx As Long

    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")
    idx = sut.IndexOf("mand_v1")

    sut.SetHidden idx, True

    Assert.IsFalse sut.IsHidden(idx), _
                   "SetHidden should silently ignore mandatory entries"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetHiddenIgnoresMandatory", Err.Number, Err.Description
End Sub


'@section Visibility tests
'===============================================================================

'@TestMethod("ShowHide")
Public Sub TestIsHiddenReturnsEffectiveState()
    CustomTestSetTitles Assert, TESTMODULE, "TestIsHiddenReturnsEffectiveState"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim idx As Long

    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")
    idx = sut.IndexOf("opt_vis_v1")

    Assert.IsTrue idx > 0, "opt_vis_v1 should exist"
    Assert.IsFalse sut.IsHidden(idx), _
                   "Optional visible variables should default to visible"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIsHiddenReturnsEffectiveState", Err.Number, Err.Description
End Sub


'@TestMethod("ShowHide")
Public Sub TestSetHiddenUpdatesState()
    CustomTestSetTitles Assert, TESTMODULE, "TestSetHiddenUpdatesState"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim idx As Long

    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")
    idx = sut.IndexOf("opt_vis_v1")

    sut.SetHidden idx, True
    Assert.IsTrue sut.IsHidden(idx), _
                  "IsHidden should return True after SetHidden(True)"

    sut.SetHidden idx, False
    Assert.IsFalse sut.IsHidden(idx), _
                   "IsHidden should return False after SetHidden(False)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetHiddenUpdatesState", Err.Number, Err.Description
End Sub


'@TestMethod("ShowHide")
Public Sub TestSetAllOptionalHiddenHidesNonMandatory()
    CustomTestSetTitles Assert, TESTMODULE, "TestSetAllOptionalHiddenHidesNonMandatory"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim mandIdx As Long
    Dim optIdx As Long

    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")
    mandIdx = sut.IndexOf("mand_v1")
    optIdx = sut.IndexOf("opt_vis_v1")

    sut.SetAllOptionalHidden True

    Assert.IsFalse sut.IsHidden(mandIdx), _
                   "Mandatory entries should remain visible after SetAllOptionalHidden"
    Assert.IsTrue sut.IsHidden(optIdx), _
                  "Optional entries should be hidden after SetAllOptionalHidden(True)"

    sut.SetAllOptionalHidden False
    Assert.IsFalse sut.IsHidden(optIdx), _
                   "Optional entries should be visible after SetAllOptionalHidden(False)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetAllOptionalHiddenHidesNonMandatory", Err.Number, Err.Description
End Sub


'@TestMethod("ShowHide")
Public Sub TestForceHiddenOnCRFLayer()
    CustomTestSetTitles Assert, TESTMODULE, "TestForceHiddenOnCRFLayer"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim idx As Long

    Set sut = ShowHideManager.Create(Dict, ShowHideLayerCRF, "hlist2D-sheet2")
    idx = sut.IndexOf("val_of_text_h2")

    Assert.IsTrue idx > 0, "val_of_text_h2 should exist on hlist2D-sheet2 CRF"
    Assert.IsTrue sut.IsHidden(idx), _
                  "Formula-based variables must be force-hidden on CRF layer"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestForceHiddenOnCRFLayer", Err.Number, Err.Description
End Sub


'@section Lookup tests
'===============================================================================

'@TestMethod("ShowHide")
Public Sub TestIndexOfFindsEntry()
    CustomTestSetTitles Assert, TESTMODULE, "TestIndexOfFindsEntry"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim idx As Long

    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")
    idx = sut.IndexOf("opt_vis_v1")

    Assert.IsTrue idx > 0, "IndexOf should return a positive index for existing field"
    Assert.AreEqual "opt_vis_v1", sut.FieldKey(idx), _
                     "FieldKey at the found index should match the lookup key"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIndexOfFindsEntry", Err.Number, Err.Description
End Sub


'@TestMethod("ShowHide")
Public Sub TestIndexOfReturnsZeroForMissing()
    CustomTestSetTitles Assert, TESTMODULE, "TestIndexOfReturnsZeroForMissing"
    On Error GoTo TestFail

    Dim sut As IShowHideManager

    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")

    Assert.AreEqual CLng(0), sut.IndexOf("nonexistent_var"), _
                     "IndexOf should return 0 for a missing field key"
    Assert.IsFalse sut.HasField("nonexistent_var"), _
                   "HasField should return False for a missing field key"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestIndexOfReturnsZeroForMissing", Err.Number, Err.Description
End Sub


'@section Plan tests
'===============================================================================

'@TestMethod("ShowHide")
Public Sub TestPlanReturnsCorrectData()
    CustomTestSetTitles Assert, TESTMODULE, "TestPlanReturnsCorrectData"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim planData As BetterArray
    Dim entry As Variant

    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")
    Set planData = sut.Plan()

    Assert.AreEqual sut.EntryCount, planData.Length, _
                     "Plan should return one entry per managed variable"

    ' Verify plan entry structure: Array(headerText, fieldKey, effectiveHidden)
    entry = planData.Item(planData.LowerBound)
    Assert.IsTrue LenB(CStr(entry(0))) > 0, _
                  "Plan entry(0) should contain non-empty header text"
    Assert.IsTrue LenB(CStr(entry(1))) > 0, _
                  "Plan entry(1) should contain non-empty field key"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPlanReturnsCorrectData", Err.Number, Err.Description
End Sub


'@section Persistence tests
'===============================================================================

'@TestMethod("ShowHide")
Public Sub TestExportPlanWritesRows()
    CustomTestSetTitles Assert, TESTMODULE, "TestExportPlanWritesRows"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim lo As ListObject
    Dim persistSheet As Worksheet

    Set persistSheet = CreatePersistenceSheet()
    Set lo = CreatePersistenceTable(persistSheet)
    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")

    sut.ExportPlan lo

    Assert.AreEqual sut.EntryCount, lo.ListRows.Count, _
                     "ExportPlan should write one row per entry"
    Assert.AreEqual "vlist", CStr(lo.DataBodyRange.Cells(1, 1).Value), _
                     "Layer token should be stored in the first column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestExportPlanWritesRows", Err.Number, Err.Description
End Sub


'@TestMethod("ShowHide")
Public Sub TestImportPlanUpdatesHiddenState()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportPlanUpdatesHiddenState"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim lo As ListObject
    Dim persistSheet As Worksheet
    Dim optIdx As Long

    Set persistSheet = CreatePersistenceSheet()
    Set lo = CreatePersistenceTable(persistSheet)
    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")
    optIdx = sut.IndexOf("opt_vis_v1")

    ' Export initial state, modify the table to hide opt_vis_v1, then re-import
    sut.ExportPlan lo
    SetPersistedHiddenFlag lo, "opt_vis_v1", "true"

    ' Re-create manager and import
    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")
    sut.ImportPlan lo
    optIdx = sut.IndexOf("opt_vis_v1")

    Assert.IsTrue sut.IsHidden(optIdx), _
                  "ImportPlan should update hidden state from persisted data"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestImportPlanUpdatesHiddenState", Err.Number, Err.Description
End Sub


'@TestMethod("ShowHide")
Public Sub TestImportPlanCannotHideMandatory()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportPlanCannotHideMandatory"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim lo As ListObject
    Dim persistSheet As Worksheet
    Dim mandIdx As Long

    Set persistSheet = CreatePersistenceSheet()
    Set lo = CreatePersistenceTable(persistSheet)
    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")

    ' Export, then force mand_v1 to hidden in the table
    sut.ExportPlan lo
    SetPersistedHiddenFlag lo, "mand_v1", "true"

    ' Re-create and import
    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")
    sut.ImportPlan lo
    mandIdx = sut.IndexOf("mand_v1")

    Assert.IsFalse sut.IsHidden(mandIdx), _
                   "ImportPlan must not hide mandatory variables"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestImportPlanCannotHideMandatory", Err.Number, Err.Description
End Sub


'@TestMethod("ShowHide")
Public Sub TestExportClearsLayerBeforeWriting()
    CustomTestSetTitles Assert, TESTMODULE, "TestExportClearsLayerBeforeWriting"
    On Error GoTo TestFail

    Dim sut As IShowHideManager
    Dim lo As ListObject
    Dim persistSheet As Worksheet
    Dim initialCount As Long

    Set persistSheet = CreatePersistenceSheet()
    Set lo = CreatePersistenceTable(persistSheet)
    Set sut = ShowHideManager.Create(Dict, ShowHideLayerVList, "vlist1D-sheet1")

    ' Export twice to verify old rows are cleared
    sut.ExportPlan lo
    initialCount = lo.ListRows.Count
    sut.ExportPlan lo

    Assert.AreEqual initialCount, lo.ListRows.Count, _
                     "Second export should not double the row count"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestExportClearsLayerBeforeWriting", Err.Number, Err.Description
End Sub


'@section Helpers
'===============================================================================

Private Function CreatePersistenceSheet() As Worksheet
    Set CreatePersistenceSheet = FixtureWorkbook.Worksheets.Add
End Function

Private Function CreatePersistenceTable(ByVal targetSheet As Worksheet) As ListObject
    Dim headerValues As Variant
    Dim headerRange As Range

    headerValues = TestHelpers.RowsToMatrix(Array(Array("layer", "field_key", "header_text", "hidden_flag")))
    targetSheet.Cells.Clear

    Set headerRange = targetSheet.Range("A1").Resize(1, UBound(headerValues, 2))
    headerRange.Value = headerValues

    Set CreatePersistenceTable = targetSheet.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=headerRange, _
        XlListObjectHasHeaders:=xlYes)
    CreatePersistenceTable.Name = "tbl_showhide_persist"
End Function

' @description Set the hidden_flag column for a specific field_key row in the persistence table.
Private Sub SetPersistedHiddenFlag(ByVal lo As ListObject, ByVal fieldKey As String, ByVal flagValue As String)
    Dim idx As Long

    For idx = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(idx, 2).Value) = fieldKey Then
            lo.DataBodyRange.Cells(idx, 4).Value = flagValue
            Exit Sub
        End If
    Next idx
End Sub
