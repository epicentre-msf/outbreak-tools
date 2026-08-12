Attribute VB_Name = "TestShowHide"
Attribute VB_Description = "Tests for ShowHide, ShowHideLayout and ShowHideStore"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for ShowHide, ShowHideLayout and ShowHideStore")

Option Explicit

'@description
'Drives the three classes of the show/hide feature. ShowHide says which
'variables a sheet offers and which of them are hidden, ShowHideLayout writes
'that to a worksheet, and ShowHideStore carries the choices from one session to
'the next.
'
'ONE WORKBOOK PER MODULE, AND A POOL OF SHEETS INSIDE IT
'-------------------------------------------------------------------------------
'The dictionary fixture is written and prepared once, in ModuleInitialize.
'Preparing a dictionary derives four columns over seventy odd rows, and the
'three classes only read it, so doing that per test bought nothing.
'
'The scratch worksheets are made once as well, and a test that needs one asks
'ScratchSheet for the next of the pool. **Nothing here deletes a worksheet.**
'OBTHeadless.ResetOutputSheet records that Worksheet.Delete is unreliable on
'macOS Excel and hard-clears its own output sheet instead, and forty deletions
'a run is not the place to test that. A pool sheet is hard-cleared on the way
'out: tables, contents, then the row and column geometry the layout tests
'write.
'
'THE FIXTURE FAILS QUIETLY AND EVERY TEST SAYS SO
'-------------------------------------------------------------------------------
'An error escaping ModuleInitialize is a modal dialog, which stops the whole
'run. The setup captures its error into two module fields and FixtureReady
'reports it as each test's own failure.
'
'WHAT THE FIXTURE DICTIONARY CARRIES
'-------------------------------------------------------------------------------
'  hid_v1, hid_h2       status "hidden", so the user never sees them
'  mand_v1, mand_h2     status "mandatory"
'  opt_hid_v1           status "optional, hidden"
'  opt_vis_v1           status "optional, visible"
'  vis_hidd_reg_h2      register book "hidden", so it prints hidden alone
'  val_of_text_h2       a formula on hlist2D-sheet2, locked off the CRF
'@depends ShowHide, ShowHideLayout, ShowHideStore, LLdictionary, CustomTest

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private Dict As LLdictionary
Private ScratchSheets As Collection
Private ScratchTaken As Long
Private SetupError As Long
Private SetupMessage As String

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ShowHide"
Private Const DICTIONARY_SHEET As String = "DictionaryFixture"

'How many scratch worksheets the pool holds. The test that takes the most asks
'for three: one for the store, one to write, one to read back into.
Private Const SCRATCH_POOL_SIZE As Long = 4

'How far the geometry reset of a pool sheet reaches, in rows and in columns
Private Const SCRATCH_RESET_SPAN As Long = 300

Private Const VLIST_SHEET As String = "vlist1D-sheet1"
Private Const HLIST_SHEET As String = "hlist2D-sheet1"
Private Const HLIST_SHEET_TWO As String = "hlist2D-sheet2"


'@section Lifecycle
'===============================================================================

'@sub-title Build the assertion harness, the fixture workbook and the dictionary.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    Dim counter As Long

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestShowHide"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set FixtureWorkbook = NewWorkbook()
        DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, FixtureWorkbook
        Set Dict = LLdictionary.Create(FixtureWorkbook.Worksheets(DICTIONARY_SHEET), 1, 1)
        Dict.Prepare

        'The whole scratch pool is made here and reused for the run
        Set ScratchSheets = New Collection
        For counter = 1 To SCRATCH_POOL_SIZE
            ScratchSheets.Add FixtureWorkbook.Worksheets.Add
        Next counter

        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print the results and drop the fixture workbook.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set ScratchSheets = Nothing
    Set Dict = Nothing
    Set FixtureWorkbook = Nothing
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Open the list of sheets this test will take.
'@details
'There is no BeginTest call here on purpose. BeginTest opens the checking with
'whatever titles are pending, and the Flush in TestCleanup has just reset those
'to the default, so every result of the module would be filed under the default
'label.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    ScratchTaken = 0
End Sub

'@sub-title Flush the results and hand the scratch sheets back.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    ScratchTaken = 0
End Sub


'@section Entry list
'===============================================================================

'@TestMethod("ShowHide")
Public Sub TestDesignerHiddenVariablesNeverReachTheList()
    CustomTestSetTitles Assert, TESTMODULE, "TestDesignerHiddenVariablesNeverReachTheList"
    If Not FixtureReady("TestDesignerHiddenVariablesNeverReachTheList") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)

    Assert.IsTrue sut.EntryCount > 0, _
                  "The VList sheet should offer entries"
    Assert.IsFalse sut.HasField("hid_v1"), _
                   "A variable the designer hid must never enter the list"
    Assert.IsFalse sut.HasField("hid_beg_v1"), _
                   "A variable the designer hid must never enter the list"
    Assert.IsTrue sut.HasField("opt_vis_v1"), _
                  "An optional variable belongs in the list"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDesignerHiddenVariablesNeverReachTheList", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestMandatoryEntryRefusesTheUser()
    CustomTestSetTitles Assert, TESTMODULE, "TestMandatoryEntryRefusesTheUser"
    If Not FixtureReady("TestMandatoryEntryRefusesTheUser") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim idx As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    idx = sut.IndexOf("mand_v1")

    Assert.IsTrue idx > 0, "mand_v1 should be in the list"
    Assert.IsTrue sut.IsMandatory(idx), "mand_v1 is a mandatory variable"
    Assert.IsFalse sut.IsFree(idx), "A mandatory entry does not follow the user"
    Assert.IsFalse sut.IsHidden(idx), "A mandatory entry is always visible"

    sut.SetHidden idx, True
    Assert.IsFalse sut.IsHidden(idx), _
                   "SetHidden leaves a mandatory entry visible"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestMandatoryEntryRefusesTheUser", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestFreeEntryFollowsTheUser()
    CustomTestSetTitles Assert, TESTMODULE, "TestFreeEntryFollowsTheUser"
    If Not FixtureReady("TestFreeEntryFollowsTheUser") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim visibleIdx As Long
    Dim hiddenIdx As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    visibleIdx = sut.IndexOf("opt_vis_v1")
    hiddenIdx = sut.IndexOf("opt_hid_v1")

    Assert.IsTrue sut.IsFree(visibleIdx), "opt_vis_v1 follows the user"
    Assert.IsFalse sut.IsHidden(visibleIdx), _
                   "An optional visible variable starts visible"
    Assert.IsTrue sut.IsHidden(hiddenIdx), _
                  "An optional hidden variable starts hidden"

    sut.SetHidden visibleIdx, True
    Assert.IsTrue sut.IsHidden(visibleIdx), "SetHidden True hides a free entry"

    sut.SetHidden visibleIdx, False
    Assert.IsFalse sut.IsHidden(visibleIdx), "SetHidden False shows it again"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFreeEntryFollowsTheUser", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSetAllOptionalHiddenSkipsMandatory()
    CustomTestSetTitles Assert, TESTMODULE, "TestSetAllOptionalHiddenSkipsMandatory"
    If Not FixtureReady("TestSetAllOptionalHiddenSkipsMandatory") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim mandIdx As Long
    Dim optIdx As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    mandIdx = sut.IndexOf("mand_v1")
    optIdx = sut.IndexOf("opt_vis_v1")

    sut.SetAllOptionalHidden True
    Assert.IsFalse sut.IsHidden(mandIdx), _
                   "A mandatory entry stays visible when everything else hides"
    Assert.IsTrue sut.IsHidden(optIdx), _
                  "A free entry hides with the rest"

    sut.SetAllOptionalHidden False
    Assert.IsFalse sut.IsHidden(optIdx), "And shows again"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetAllOptionalHiddenSkipsMandatory", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestUnknownSheetGivesAnEmptyList()
    CustomTestSetTitles Assert, TESTMODULE, "TestUnknownSheetGivesAnEmptyList"
    If Not FixtureReady("TestUnknownSheetGivesAnEmptyList") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide

    Set sut = ShowHide.Create(Dict, ShowHideLayerHList, "no_such_sheet")

    Assert.AreEqual CLng(0), sut.EntryCount, _
                     "A sheet the dictionary does not name offers nothing"
    Assert.AreEqual CLng(0), sut.IndexOf("mand_h2"), _
                     "IndexOf answers 0 on an empty list"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUnknownSheetGivesAnEmptyList", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestLookupIgnoresCaseAndStraySpaces()
    CustomTestSetTitles Assert, TESTMODULE, "TestLookupIgnoresCaseAndStraySpaces"
    If Not FixtureReady("TestLookupIgnoresCaseAndStraySpaces") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim idx As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    idx = sut.IndexOf("opt_vis_v1")

    Assert.AreEqual idx, sut.IndexOf("  OPT_VIS_V1  "), _
                     "A key is trimmed and lower cased on the way in"
    Assert.AreEqual CLng(0), sut.IndexOf("   "), _
                     "A key of spaces alone answers 0"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLookupIgnoresCaseAndStraySpaces", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestLayerFilterFollowsTheSheetType()
    CustomTestSetTitles Assert, TESTMODULE, "TestLayerFilterFollowsTheSheetType"
    If Not FixtureReady("TestLayerFilterFollowsTheSheetType") Then Exit Sub
    On Error GoTo TestFail

    Assert.IsTrue ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET).EntryCount > 0, _
                  "An hlist2D sheet offers its entries on the HList layer"
    Assert.IsTrue ShowHide.Create(Dict, ShowHideLayerPrinted, HLIST_SHEET).EntryCount > 0, _
                  "A printed sheet is derived from the same hlist2D sheet"
    Assert.AreEqual CLng(0), _
                    ShowHide.Create(Dict, ShowHideLayerVList, HLIST_SHEET).EntryCount, _
                    "An hlist2D sheet has nothing on the VList layer"
    Assert.AreEqual CLng(0), _
                    ShowHide.Create(Dict, ShowHideLayerHList, VLIST_SHEET).EntryCount, _
                    "And a vlist1D sheet has nothing on the HList layer"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLayerFilterFollowsTheSheetType", Err.Number, Err.Description
End Sub


'@section The CRF layer
'===============================================================================

'@TestMethod("ShowHide")
Public Sub TestCrfLayerOffersEntries()
    CustomTestSetTitles Assert, TESTMODULE, "TestCrfLayerOffersEntries"
    If Not FixtureReady("TestCrfLayerOffersEntries") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide

    'The CRF filter used to test the dictionary sheet type for the word "crf",
    'which the vocabulary has never held, so every CRF list came back empty.
    Set sut = ShowHide.Create(Dict, ShowHideLayerCRF, HLIST_SHEET_TWO)

    Assert.IsTrue sut.EntryCount > 0, _
                  "A CRF is derived from an hlist2D sheet, so it offers its entries"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCrfLayerOffersEntries", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestCalculatedColumnIsLockedOnCrf()
    CustomTestSetTitles Assert, TESTMODULE, "TestCalculatedColumnIsLockedOnCrf"
    If Not FixtureReady("TestCalculatedColumnIsLockedOnCrf") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim idx As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerCRF, HLIST_SHEET_TWO)
    idx = sut.IndexOf("val_of_text_h2")

    Assert.IsTrue idx > 0, "val_of_text_h2 is on hlist2D-sheet2"
    Assert.IsTrue sut.IsLocked(idx), _
                  "A calculated column is held hidden on a CRF"
    Assert.IsTrue sut.IsHidden(idx), "A locked entry reads as hidden"
    Assert.IsFalse sut.IsFree(idx), "A locked entry does not follow the user"

    sut.SetHidden idx, False
    Assert.IsTrue sut.IsHidden(idx), "SetHidden leaves a locked entry hidden"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCalculatedColumnIsLockedOnCrf", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestPrintedLayerReadsRegisterBook()
    CustomTestSetTitles Assert, TESTMODULE, "TestPrintedLayerReadsRegisterBook"
    If Not FixtureReady("TestPrintedLayerReadsRegisterBook") Then Exit Sub
    On Error GoTo TestFail

    Dim printed As ShowHide
    Dim onSheet As ShowHide
    Dim idx As Long

    Set printed = ShowHide.Create(Dict, ShowHideLayerPrinted, HLIST_SHEET)
    Set onSheet = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET)

    idx = printed.IndexOf("vis_hidd_reg_h2")
    Assert.IsTrue idx > 0, "vis_hidd_reg_h2 is on the printed layer"
    Assert.IsTrue printed.IsHidden(idx), _
                  "A register book of 'hidden' starts the entry hidden in print"
    Assert.IsFalse printed.AuthoredVertical(idx), _
                   "Its header direction was never set to vertical"

    idx = onSheet.IndexOf("vis_hidd_reg_h2")
    Assert.IsFalse onSheet.IsHidden(idx), _
                   "The same variable is visible on the data entry sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPrintedLayerReadsRegisterBook", Err.Number, Err.Description
End Sub


'@section The worksheet
'===============================================================================

'@TestMethod("ShowHide")
Public Sub TestAxisIsRowOnVListAndCrf()
    CustomTestSetTitles Assert, TESTMODULE, "TestAxisIsRowOnVListAndCrf"
    If Not FixtureReady("TestAxisIsRowOnVListAndCrf") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet

    Set sh = ScratchSheet()

    Assert.AreEqual ShowHideAxisRow, _
                    ShowHideLayout.Create(sh, ShowHideLayerVList).Axis, _
                    "A VList sheet holds one variable per row"
    Assert.AreEqual ShowHideAxisRow, _
                    ShowHideLayout.Create(sh, ShowHideLayerCRF).Axis, _
                    "A CRF holds one variable per row"
    Assert.AreEqual ShowHideAxisColumn, _
                    ShowHideLayout.Create(sh, ShowHideLayerHList).Axis, _
                    "A data entry sheet holds one variable per column"
    Assert.AreEqual ShowHideAxisColumn, _
                    ShowHideLayout.Create(sh, ShowHideLayerPrinted).Axis, _
                    "A printed sheet holds one variable per column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAxisIsRowOnVListAndCrf", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSetHiddenWritesTheRightAxis()
    CustomTestSetTitles Assert, TESTMODULE, "TestSetHiddenWritesTheRightAxis"
    If Not FixtureReady("TestSetHiddenWritesTheRightAxis") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim onRows As ShowHideLayout
    Dim onColumns As ShowHideLayout

    Set sh = ScratchSheet()
    Set onRows = ShowHideLayout.Create(sh, ShowHideLayerCRF)
    Set onColumns = ShowHideLayout.Create(sh, ShowHideLayerHList)

    onRows.SetHidden 3, True
    Assert.IsTrue sh.Rows(3).Hidden, "The CRF layer hides row 3"
    Assert.IsFalse sh.Columns(3).Hidden, "And leaves column 3 alone"
    Assert.IsTrue onRows.IsHidden(3), "IsHidden reads the same row back"

    onRows.SetHidden 3, False
    onColumns.SetHidden 4, True
    Assert.IsTrue sh.Columns(4).Hidden, "The HList layer hides column 4"
    Assert.IsFalse sh.Rows(4).Hidden, "And leaves row 4 alone"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetHiddenWritesTheRightAxis", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSetSizeRefusesAnEmptySize()
    CustomTestSetTitles Assert, TESTMODULE, "TestSetSizeRefusesAnEmptySize"
    If Not FixtureReady("TestSetSizeRefusesAnEmptySize") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim layout As ShowHideLayout

    Set sh = ScratchSheet()
    Set layout = ShowHideLayout.Create(sh, ShowHideLayerHList)

    layout.SetSize 2, 20
    Assert.AreEqual CDbl(20), layout.Size(2), "SetSize writes a real width"

    'Excel hides a column that is set to width 0, and a blank size cell reads
    'as 0, so a missing size used to make the entry vanish.
    layout.SetSize 2, 0
    Assert.AreEqual CDbl(20), layout.Size(2), "A size of zero is refused"
    Assert.IsFalse sh.Columns(2).Hidden, "And the column is still visible"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetSizeRefusesAnEmptySize", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSizeWhenShownReadsThroughAHiddenColumn()
    CustomTestSetTitles Assert, TESTMODULE, "TestSizeWhenShownReadsThroughAHiddenColumn"
    If Not FixtureReady("TestSizeWhenShownReadsThroughAHiddenColumn") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim layout As ShowHideLayout

    Set sh = ScratchSheet()
    Set layout = ShowHideLayout.Create(sh, ShowHideLayerHList)

    layout.SetSize 5, 18
    layout.SetHidden 5, True

    Assert.AreEqual CDbl(0), layout.Size(5), _
                     "Excel reports a hidden column as width 0"
    Assert.AreEqual CDbl(18), layout.SizeWhenShown(5), _
                     "SizeWhenShown answers the width the column has when visible"
    Assert.IsTrue sh.Columns(5).Hidden, _
                  "And leaves the column hidden afterwards"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSizeWhenShownReadsThroughAHiddenColumn", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestApplyPutsTheSheetInStepWithTheEntries()
    CustomTestSetTitles Assert, TESTMODULE, "TestApplyPutsTheSheetInStepWithTheEntries"
    If Not FixtureReady("TestApplyPutsTheSheetInStepWithTheEntries") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim idx As Long
    Dim position As Long

    Set sh = ScratchSheet()
    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET)
    Set layout = ShowHideLayout.Create(sh, ShowHideLayerHList)

    idx = entries.IndexOf("opt_vis_h2")
    position = entries.PositionIndex(idx)
    Assert.IsTrue position > 0, "opt_vis_h2 has a column index"

    entries.SetHidden idx, True
    Assert.IsTrue entries.Apply(layout) > 0, "Apply reports the positions it set"
    Assert.IsTrue sh.Columns(position).Hidden, _
                  "The column the list hides is hidden on the sheet"

    entries.SetHidden idx, False
    entries.Apply layout
    Assert.IsFalse sh.Columns(position).Hidden, "And shown again"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestApplyPutsTheSheetInStepWithTheEntries", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestAdoptReadsTheSheetBackIntoTheEntries()
    CustomTestSetTitles Assert, TESTMODULE, "TestAdoptReadsTheSheetBackIntoTheEntries"
    If Not FixtureReady("TestAdoptReadsTheSheetBackIntoTheEntries") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim idx As Long
    Dim mandIdx As Long

    Set sh = ScratchSheet()
    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET)
    Set layout = ShowHideLayout.Create(sh, ShowHideLayerHList)

    idx = entries.IndexOf("opt_vis_h2")
    mandIdx = entries.IndexOf("mand_h2")

    'The user hid the column by hand rather than through the form
    sh.Columns(entries.PositionIndex(idx)).Hidden = True
    sh.Columns(entries.PositionIndex(mandIdx)).Hidden = True

    entries.Adopt layout

    Assert.IsTrue entries.IsHidden(idx), "Adopt reads a hand hidden column back"
    Assert.IsFalse entries.IsHidden(mandIdx), _
                   "A mandatory entry keeps what the dictionary said"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAdoptReadsTheSheetBackIntoTheEntries", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestResetToAuthoredPutsTheChoicesBack()
    CustomTestSetTitles Assert, TESTMODULE, "TestResetToAuthoredPutsTheChoicesBack"
    If Not FixtureReady("TestResetToAuthoredPutsTheChoicesBack") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim visIdx As Long
    Dim hidIdx As Long
    Dim visPos As Long
    Dim hidPos As Long

    Set sh = ScratchSheet()
    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET)
    Set layout = ShowHideLayout.Create(sh, ShowHideLayerHList)

    visIdx = entries.IndexOf("opt_vis_h2")
    hidIdx = entries.IndexOf("opt_hid_h2")
    visPos = entries.PositionIndex(visIdx)
    hidPos = entries.PositionIndex(hidIdx)

    'The user turned both free entries around and the sheet followed
    entries.SetHidden visIdx, True
    entries.SetHidden hidIdx, False
    entries.Apply layout
    Assert.IsTrue sh.Columns(visPos).Hidden, _
                  "The user's choice landed on the sheet"

    Assert.IsTrue entries.ResetToAuthored(layout) > 0, _
                  "ResetToAuthored reports the positions it set"

    Assert.IsFalse entries.IsHidden(visIdx), _
                   "An optional visible variable is back where the dictionary started it"
    Assert.IsTrue entries.IsHidden(hidIdx), _
                  "And an optional hidden variable is back hidden"
    Assert.IsFalse sh.Columns(visPos).Hidden, _
                   "The sheet shows the authored visible column again"
    Assert.IsTrue sh.Columns(hidPos).Hidden, _
                  "And hides the authored hidden one"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestResetToAuthoredPutsTheChoicesBack", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestResetToAuthoredRestoresThePrintedHeaderDirection()
    CustomTestSetTitles Assert, TESTMODULE, "TestResetToAuthoredRestoresThePrintedHeaderDirection"
    If Not FixtureReady("TestResetToAuthoredRestoresThePrintedHeaderDirection") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim idx As Long
    Dim position As Long

    Set sh = ScratchSheet()
    Set entries = ShowHide.Create(Dict, ShowHideLayerPrinted, HLIST_SHEET)

    'The printed header sits one row above the PRINTSTART anchor
    sh.Names.Add Name:="table1_PRINTSTART", _
                 RefersTo:="='" & sh.Name & "'!" & sh.Cells(5, 1).Address
    Set layout = ShowHideLayout.Create(sh, ShowHideLayerPrinted, _
                                       baseTableName:="table1")

    idx = entries.IndexOf("opt_vis_h2")
    position = entries.PositionIndex(idx)
    Assert.IsFalse entries.AuthoredVertical(idx), _
                   "The dictionary never asked this header to be turned"

    'Excel reports a flat cell as xlHorizontal, and IsVertical used to read
    'that as turned, so every flat header answered vertical.
    Assert.IsFalse layout.IsVertical(position), _
                   "A header never turned reads flat"

    'The user turned the header by hand
    layout.SetOrientation position, True
    Assert.IsTrue layout.IsVertical(position), _
                  "The turned header reads back vertical"

    entries.ResetToAuthored layout

    Assert.AreEqual CLng(0), layout.FailureCount, _
                    "The reset was refused by nothing"
    Assert.IsFalse layout.IsVertical(position), _
                   "The reset lays the header back flat"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestResetToAuthoredRestoresThePrintedHeaderDirection", Err.Number, Err.Description
End Sub


'@section The store
'===============================================================================

'@TestMethod("ShowHide")
Public Sub TestStoreBuildsItsOwnTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestStoreBuildsItsOwnTable"
    If Not FixtureReady("TestStoreBuildsItsOwnTable") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim store As ShowHideStore

    Set sh = ScratchSheet()
    Assert.AreEqual CLng(0), CLng(sh.ListObjects.Count), _
                     "The sheet starts with no table on it"

    'Nothing in a linelist ever built this table, so every save and every load
    'returned without writing and the user's choices were thrown away.
    Set store = ShowHideStore.CreateOnSheet(sh)

    Assert.IsTrue store.HasTable, "The store makes the table it needs"
    Assert.AreEqual CLng(1), CLng(sh.ListObjects.Count), _
                     "One table lands on the sheet"
    Assert.AreEqual CLng(7), CLng(store.Table.ListColumns.Count), _
                     "It carries the seven columns the store writes"
    Assert.AreEqual CLng(0), store.RowCount, "And it starts empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestStoreBuildsItsOwnTable", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestStoreDropsThePerSheetTablesOfTheOldDesign()
    CustomTestSetTitles Assert, TESTMODULE, "TestStoreDropsThePerSheetTablesOfTheOldDesign"
    If Not FixtureReady("TestStoreDropsThePerSheetTablesOfTheOldDesign") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim store As ShowHideStore

    Set sh = ScratchSheet()

    'The first design tiled one table per sheet across this worksheet. Both
    'designs read ListObjects(1), so the two cannot live together.
    sh.Cells(1, 1).Value = "main label"
    sh.Cells(1, 2).Value = "variable name"
    sh.ListObjects.Add(xlSrcRange, sh.Range(sh.Cells(1, 1), sh.Cells(1, 2)), , xlYes) _
      .Name = "ShowHideTable_oldone"

    Set store = ShowHideStore.CreateOnSheet(sh)

    Assert.IsTrue store.HasTable, "The store still gets its table"
    Assert.AreEqual CLng(1), CLng(sh.ListObjects.Count), _
                     "The table of the old design is gone"
    Assert.AreEqual "show_hide_state", store.Table.Name, _
                     "And the one left is the store's"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestStoreDropsThePerSheetTablesOfTheOldDesign", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestCreateReadsBothStoreSheetNames()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReadsBothStoreSheetNames"
    If Not FixtureReady("TestCreateReadsBothStoreSheetNames") Then Exit Sub
    On Error GoTo TestFail

    Dim oldSh As Worksheet
    Dim newSh As Worksheet
    Dim store As ShowHideStore

    'The two sheets live in the fixture workbook: opening and closing an
    'extra workbook mid-module leaves another window holding the screen,
    'and PrintResults then fails at cleanup. ModuleCleanup drops the whole
    'fixture workbook, so nothing lingers.
    'A linelist generated before the internal-sheet rename carries the old
    'trailing name alone, and Create still has to find its store there.
    Set oldSh = FixtureWorkbook.Worksheets.Add
    oldSh.Name = "show_hide__"

    Set store = ShowHideStore.Create(FixtureWorkbook)

    Assert.IsTrue store.HasTable, "The old trailing name still opens the store"
    Assert.AreEqual CLng(1), CLng(oldSh.ListObjects.Count), _
                     "The table lands on the old-named sheet"

    'Once the workbook carries the leading name, that sheet wins.
    Set newSh = FixtureWorkbook.Worksheets.Add
    newSh.Name = "__show_hide"

    Set store = ShowHideStore.Create(FixtureWorkbook)

    Assert.IsTrue store.HasTable, "The leading name opens the store too"
    Assert.AreEqual CLng(1), CLng(newSh.ListObjects.Count), _
                     "The table lands on the leading-named sheet first"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReadsBothStoreSheetNames", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSavedChoiceComesBackOnLoad()
    CustomTestSetTitles Assert, TESTMODULE, "TestSavedChoiceComesBackOnLoad"
    If Not FixtureReady("TestSavedChoiceComesBackOnLoad") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim saved As ShowHide
    Dim reloaded As ShowHide
    Dim idx As Long

    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set saved = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)

    idx = saved.IndexOf("opt_vis_v1")
    saved.SetHidden idx, True
    store.Save saved

    Assert.AreEqual saved.EntryCount, store.RowCount, _
                     "Save writes one row per entry"

    Set reloaded = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    Assert.IsFalse reloaded.IsHidden(reloaded.IndexOf("opt_vis_v1")), _
                   "A fresh list starts where the dictionary says"

    Assert.IsTrue store.Load(reloaded) > 0, "Load reports the rows it matched"
    Assert.IsTrue reloaded.IsHidden(reloaded.IndexOf("opt_vis_v1")), _
                  "The choice the user made comes back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSavedChoiceComesBackOnLoad", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSavingOneLayerLeavesTheOthersAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestSavingOneLayerLeavesTheOthersAlone"
    If Not FixtureReady("TestSavingOneLayerLeavesTheOthersAlone") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim onVList As ShowHide
    Dim onHList As ShowHide

    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set onVList = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    Set onHList = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET)

    store.Save onVList
    store.Save onHList

    Assert.AreEqual onVList.EntryCount + onHList.EntryCount, store.RowCount, _
                     "Four layers share one table"
    Assert.IsTrue store.HasLayer(ShowHideLayerVList), "The VList rows are still there"
    Assert.IsTrue store.HasLayer(ShowHideLayerHList), "And the HList rows landed"

    'A second save of the same layer replaces its rows rather than doubling them
    store.Save onHList
    Assert.AreEqual onVList.EntryCount + onHList.EntryCount, store.RowCount, _
                     "Saving twice does not double the rows"

    store.Clear ShowHideLayerHList
    Assert.IsFalse store.HasLayer(ShowHideLayerHList), "Clear drops one layer"
    Assert.IsTrue store.HasLayer(ShowHideLayerVList), "And keeps the other"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSavingOneLayerLeavesTheOthersAlone", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestLoadRefusesToHideAMandatoryEntry()
    CustomTestSetTitles Assert, TESTMODULE, "TestLoadRefusesToHideAMandatoryEntry"
    If Not FixtureReady("TestLoadRefusesToHideAMandatoryEntry") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim entries As ShowHide

    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set entries = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)

    store.Save entries
    SetStoredFlag store, "mand_v1", "true"

    Set entries = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    store.Load entries

    Assert.IsFalse entries.IsHidden(entries.IndexOf("mand_v1")), _
                   "A saved file cannot hide a mandatory variable"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLoadRefusesToHideAMandatoryEntry", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestBlankSizeLeavesTheSheetAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestBlankSizeLeavesTheSheetAlone"
    If Not FixtureReady("TestBlankSizeLeavesTheSheetAlone") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim entries As ShowHide
    Dim sh As Worksheet
    Dim layout As ShowHideLayout
    Dim position As Long

    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set sh = ScratchSheet()
    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET)
    Set layout = ShowHideLayout.Create(sh, ShowHideLayerHList)

    position = entries.PositionIndex(entries.IndexOf("opt_vis_h2"))
    layout.SetSize position, 22

    'Save with no layout, so every entry_size cell is left blank
    store.Save entries
    store.Load entries, layout

    Assert.AreEqual CDbl(22), layout.Size(position), _
                     "A blank size leaves the column the width it had"
    Assert.IsFalse sh.Columns(position).Hidden, _
                   "And a blank size never hides a column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBlankSizeLeavesTheSheetAlone", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSaveRecordsTheSizeOfAHiddenEntry()
    CustomTestSetTitles Assert, TESTMODULE, "TestSaveRecordsTheSizeOfAHiddenEntry"
    If Not FixtureReady("TestSaveRecordsTheSizeOfAHiddenEntry") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim entries As ShowHide
    Dim sh As Worksheet
    Dim layout As ShowHideLayout
    Dim rebuilt As ShowHide
    Dim target As Worksheet
    Dim targetLayout As ShowHideLayout
    Dim position As Long

    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set sh = ScratchSheet()
    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET)
    Set layout = ShowHideLayout.Create(sh, ShowHideLayerHList)

    position = entries.PositionIndex(entries.IndexOf("opt_vis_h2"))

    'A hidden column reports width 0, which used to be the width that shipped
    layout.SetSize position, 24
    layout.SetHidden position, True
    entries.Adopt layout

    store.Save entries, layout

    'Read it into a second sheet, the way a migration import does
    Set target = ScratchSheet()
    Set targetLayout = ShowHideLayout.Create(target, ShowHideLayerHList)
    Set rebuilt = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET)

    store.Load rebuilt, targetLayout
    rebuilt.Apply targetLayout

    Assert.IsTrue target.Columns(position).Hidden, _
                  "The entry comes back hidden"
    Assert.AreEqual CDbl(24), targetLayout.SizeWhenShown(position), _
                     "And it comes back at the width it really had"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSaveRecordsTheSizeOfAHiddenEntry", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestReadingASheetWithNoTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestReadingASheetWithNoTable"
    If Not FixtureReady("TestReadingASheetWithNoTable") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim entries As ShowHide

    Set store = ShowHideStore.CreateForRead(ScratchSheet())
    Set entries = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)

    Assert.IsFalse store.HasTable, _
                   "Reading a sheet that carries no table provisions nothing"
    Assert.AreEqual CLng(0), store.Load(entries), _
                     "And loading from it changes nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestReadingASheetWithNoTable", Err.Number, Err.Description
End Sub


'@section Saved layouts
'===============================================================================
'@description
'A saved layout is a named copy of the show/hide rows, kept in the same table
'behind the layout_name column. The current state has no name.

'@TestMethod("ShowHide")
Public Sub TestASavedLayoutSitsBesideTheCurrentState()
    CustomTestSetTitles Assert, TESTMODULE, "TestASavedLayoutSitsBesideTheCurrentState"
    If Not FixtureReady("TestASavedLayoutSitsBesideTheCurrentState") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim entries As ShowHide
    Dim reloaded As ShowHide

    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set entries = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)

    store.Save entries
    entries.SetHidden entries.IndexOf("opt_vis_v1"), True
    store.Save entries, layoutName:="compact"

    Assert.AreEqual entries.EntryCount * 2, store.RowCount, _
                     "The current state and the layout each hold their rows"
    Assert.IsTrue store.HasLayout("compact"), "The layout is stored under its name"
    Assert.AreEqual CLng(1), store.LayoutCount(), "And it is the only one"

    Set reloaded = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    store.Load reloaded
    Assert.IsFalse reloaded.IsHidden(reloaded.IndexOf("opt_vis_v1")), _
                   "The current state still reads what it held before the layout"

    Set reloaded = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    Assert.IsTrue store.Load(reloaded, layoutName:="compact") > 0, _
                  "Loading by name reports the rows it matched"
    Assert.IsTrue reloaded.IsHidden(reloaded.IndexOf("opt_vis_v1")), _
                  "And the layout reads the choice it was saved with"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASavedLayoutSitsBesideTheCurrentState", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSaveReplacesByKeySoSheetsShareALayer()
    CustomTestSetTitles Assert, TESTMODULE, "TestSaveReplacesByKeySoSheetsShareALayer"
    If Not FixtureReady("TestSaveReplacesByKeySoSheetsShareALayer") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim firstSheet As ShowHide
    Dim secondSheet As ShowHide

    'Both data entry sheets write hlist rows. Replacing the whole layer on
    'every save dropped the first sheet's rows the moment the second saved.
    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set firstSheet = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET)
    Set secondSheet = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET_TWO)

    store.Save firstSheet
    store.Save secondSheet

    Assert.AreEqual firstSheet.EntryCount + secondSheet.EntryCount, store.RowCount, _
                     "Two sheets of one layer keep their rows side by side"

    store.Save secondSheet
    Assert.AreEqual firstSheet.EntryCount + secondSheet.EntryCount, store.RowCount, _
                     "Saving one sheet again replaces its own rows alone"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSaveReplacesByKeySoSheetsShareALayer", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestALayoutCarriesSizesAndHeaderDirections()
    CustomTestSetTitles Assert, TESTMODULE, "TestALayoutCarriesSizesAndHeaderDirections"
    If Not FixtureReady("TestALayoutCarriesSizesAndHeaderDirections") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim sh As Worksheet
    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim reloaded As ShowHide
    Dim position As Long

    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set sh = ScratchSheet()
    Set entries = ShowHide.Create(Dict, ShowHideLayerPrinted, HLIST_SHEET)

    sh.Names.Add Name:="table1_PRINTSTART", _
                 RefersTo:="='" & sh.Name & "'!" & sh.Cells(5, 1).Address
    Set layout = ShowHideLayout.Create(sh, ShowHideLayerPrinted, _
                                       baseTableName:="table1")

    position = entries.PositionIndex(entries.IndexOf("opt_vis_h2"))
    layout.SetSize position, 27
    layout.SetOrientation position, True
    entries.Adopt layout
    store.Save entries, layout, "register"

    'The user works on, and the sheet drifts away from the saved layout
    layout.SetSize position, 15
    layout.SetOrientation position, False

    Set reloaded = ShowHide.Create(Dict, ShowHideLayerPrinted, HLIST_SHEET)
    Assert.IsTrue store.Load(reloaded, layout, "register") > 0, _
                  "The named load reports the rows it matched"
    Assert.AreEqual CDbl(27), layout.SizeWhenShown(position), _
                     "The saved size lands back on the sheet"
    Assert.IsTrue layout.IsVertical(position), _
                  "And the header is turned the way the layout kept it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALayoutCarriesSizesAndHeaderDirections", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestDeleteLayoutDropsItsRowsAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestDeleteLayoutDropsItsRowsAlone"
    If Not FixtureReady("TestDeleteLayoutDropsItsRowsAlone") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim entries As ShowHide

    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set entries = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)

    store.Save entries
    store.Save entries, layoutName:="one"
    store.Save entries, layoutName:="two"

    store.DeleteLayout "one"

    Assert.IsFalse store.HasLayout("one"), "The deleted layout is gone"
    Assert.IsTrue store.HasLayout("two"), "Its neighbour stays"
    Assert.AreEqual entries.EntryCount * 2, store.RowCount, _
                     "And the current state keeps its rows"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDeleteLayoutDropsItsRowsAlone", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestRenameLayoutKeepsTheRows()
    CustomTestSetTitles Assert, TESTMODULE, "TestRenameLayoutKeepsTheRows"
    If Not FixtureReady("TestRenameLayoutKeepsTheRows") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim entries As ShowHide
    Dim reloaded As ShowHide
    Dim raised As Long

    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set entries = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)

    entries.SetHidden entries.IndexOf("opt_vis_v1"), True
    store.Save entries, layoutName:="draft"
    store.Save entries, layoutName:="other"

    store.RenameLayout "draft", "final"

    Assert.IsFalse store.HasLayout("draft"), "The old name is gone"
    Assert.IsTrue store.HasLayout("final"), "The new name is there"

    Set reloaded = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    store.Load reloaded, layoutName:="final"
    Assert.IsTrue reloaded.IsHidden(reloaded.IndexOf("opt_vis_v1")), _
                  "And the rows moved with the name"

    'Renaming onto a stored name would merge two layouts, so it is refused
    On Error Resume Next
    Err.Clear
    store.RenameLayout "final", "other"
    raised = Err.Number
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), raised, _
                     "Renaming onto a name the store holds is refused"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRenameLayoutKeepsTheRows", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestTheCapRefusesTheEleventhName()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheCapRefusesTheEleventhName"
    If Not FixtureReady("TestTheCapRefusesTheEleventhName") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim entries As ShowHide
    Dim counter As Long
    Dim raised As Long

    'The small sheet keeps the eleven table rewrites cheap
    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET_TWO)

    For counter = 1 To store.MaxSavedLayouts
        store.Save entries, layoutName:="layout" & CStr(counter)
    Next counter

    Assert.AreEqual store.MaxSavedLayouts, store.LayoutCount(), _
                     "The store holds its maximum of saved layouts"

    On Error Resume Next
    Err.Clear
    store.Save entries, layoutName:="one too many"
    raised = Err.Number
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), raised, _
                     "The name past the cap is refused"

    store.Save entries, layoutName:="layout1"
    Assert.AreEqual store.MaxSavedLayouts, store.LayoutCount(), _
                     "A name already stored still saves at the cap"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheCapRefusesTheEleventhName", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestAnOldTableGainsTheLayoutColumn()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnOldTableGainsTheLayoutColumn"
    If Not FixtureReady("TestAnOldTableGainsTheLayoutColumn") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim store As ShowHideStore
    Dim entries As ShowHide
    Dim headers As Variant
    Dim counter As Long

    'The six column table a linelist generated before saved layouts carries
    Set sh = ScratchSheet()
    headers = Array("layer", "field_key", "header_text", "hidden_flag", _
                    "entry_size", "orientation")
    For counter = LBound(headers) To UBound(headers)
        sh.Cells(1, counter + 1).Value = CStr(headers(counter))
    Next counter
    sh.Cells(2, 1).Value = "vlist"
    sh.Cells(2, 2).Value = "opt_vis_v1"
    sh.Cells(2, 4).Value = "true"
    sh.ListObjects.Add(xlSrcRange, sh.Range(sh.Cells(1, 1), sh.Cells(2, 6)), , xlYes) _
      .Name = "show_hide_state"

    Set store = ShowHideStore.CreateOnSheet(sh)

    Assert.AreEqual CLng(7), CLng(store.Table.ListColumns.Count), _
                     "Opening the store adds the layout column"
    Assert.AreEqual CLng(0), store.LayoutCount(), _
                     "The old rows name no layout"

    Set entries = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    Assert.IsTrue store.Load(entries) > 0, _
                  "And they read as the current state"
    Assert.IsTrue entries.IsHidden(entries.IndexOf("opt_vis_v1")), _
                  "With the choice they carried"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnOldTableGainsTheLayoutColumn", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestLayoutsTravelBetweenStores()
    CustomTestSetTitles Assert, TESTMODULE, "TestLayoutsTravelBetweenStores"
    If Not FixtureReady("TestLayoutsTravelBetweenStores") Then Exit Sub
    On Error GoTo TestFail

    Dim sourceStore As ShowHideStore
    Dim targetStore As ShowHideStore
    Dim entries As ShowHide
    Dim reloaded As ShowHide
    Dim skippedNames As BetterArray

    'The exporter and the importer both come down to this pair of stores
    Set sourceStore = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set targetStore = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set entries = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)

    sourceStore.Save entries
    entries.SetHidden entries.IndexOf("opt_vis_v1"), True
    sourceStore.Save entries, layoutName:="compact"

    Set skippedNames = targetStore.MergeLayoutsFrom(sourceStore)

    Assert.AreEqual CLng(0), skippedNames.Length, "Nothing is skipped under the cap"
    Assert.IsTrue targetStore.HasLayout("compact"), "The layout came across"
    Assert.AreEqual entries.EntryCount, targetStore.RowCount, _
                     "And the source's current state stayed home"

    'A second merge of the same name replaces rather than doubles
    targetStore.MergeLayoutsFrom sourceStore
    Assert.AreEqual entries.EntryCount, targetStore.RowCount, _
                     "Merging the same layout again does not double its rows"

    Set reloaded = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    targetStore.Load reloaded, layoutName:="compact"
    Assert.IsTrue reloaded.IsHidden(reloaded.IndexOf("opt_vis_v1")), _
                  "The travelled layout reads what it was saved with"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLayoutsTravelBetweenStores", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestMergeTakesCollisionsAndSkipsPastTheCap()
    CustomTestSetTitles Assert, TESTMODULE, "TestMergeTakesCollisionsAndSkipsPastTheCap"
    If Not FixtureReady("TestMergeTakesCollisionsAndSkipsPastTheCap") Then Exit Sub
    On Error GoTo TestFail

    Dim sourceStore As ShowHideStore
    Dim targetStore As ShowHideStore
    Dim entries As ShowHide
    Dim reloaded As ShowHide
    Dim skippedNames As BetterArray
    Dim counter As Long

    Set sourceStore = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set targetStore = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET_TWO)

    'The target is full, and one of its names is also in the file
    For counter = 1 To targetStore.MaxSavedLayouts - 1
        targetStore.Save entries, layoutName:="local" & CStr(counter)
    Next counter
    targetStore.Save entries, layoutName:="shared"

    entries.SetHidden entries.IndexOf("lauto_drop_h2"), True
    sourceStore.Save entries, layoutName:="shared"
    sourceStore.Save entries, layoutName:="fresh"

    Set skippedNames = targetStore.MergeLayoutsFrom(sourceStore)

    Assert.AreEqual CLng(1), skippedNames.Length, "One name found no room"
    Assert.AreEqual "fresh", CStr(skippedNames.Item(skippedNames.LowerBound)), _
                     "It is the new one, because a collision costs no slot"
    Assert.IsFalse targetStore.HasLayout("fresh"), "The skipped layout stayed out"

    Set reloaded = ShowHide.Create(Dict, ShowHideLayerHList, HLIST_SHEET_TWO)
    targetStore.Load reloaded, layoutName:="shared"
    Assert.IsTrue reloaded.IsHidden(reloaded.IndexOf("lauto_drop_h2")), _
                  "The collided name carries the file's rows"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestMergeTakesCollisionsAndSkipsPastTheCap", Err.Number, Err.Description
End Sub


'@section Hiding a whole section
'===============================================================================
'@description
'A section is a span of positions, which is what SectionMap records for it.
'These tests use the three VList entries the dictionary fixture files next to
'each other in its Status section:
'
'  opt_hid_v1   free, starts hidden
'  opt_vis_v1   free, starts visible
'  mand_v1      mandatory, always visible
'
'Every position is read off the entry list through PositionOf. `column index`
'is DERIVED by LLdictionary.Prepare and appears nowhere in the fixture rows, so
'a test that wrote the three numbers down would be asserting a number it had
'guessed. The test below pins the one thing the spans need - that the three run
'consecutively in that order - and the rest build their spans from the values.

'@TestMethod("ShowHide")
Public Sub TestTheSpanTheSectionTestsUse()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheSpanTheSectionTestsUse"
    If Not FixtureReady("TestTheSpanTheSectionTestsUse") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim firstPos As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    firstPos = PositionOf(sut, "opt_hid_v1")

    Assert.IsTrue firstPos > 0, "opt_hid_v1 has a position on the sheet"
    Assert.AreEqual firstPos + 1, PositionOf(sut, "opt_vis_v1"), _
                    "opt_vis_v1 sits at the next position"
    Assert.AreEqual firstPos + 2, PositionOf(sut, "mand_v1"), _
                    "and mand_v1 at the one after that"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSpanTheSectionTestsUse", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSetHiddenInRangeHidesEveryFreeEntry()
    CustomTestSetTitles Assert, TESTMODULE, "TestSetHiddenInRangeHidesEveryFreeEntry"
    If Not FixtureReady("TestSetHiddenInRangeHidesEveryFreeEntry") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim hidPos As Long
    Dim visPos As Long
    Dim mandPos As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    hidPos = PositionOf(sut, "opt_hid_v1")
    visPos = PositionOf(sut, "opt_vis_v1")
    mandPos = PositionOf(sut, "mand_v1")

    sut.SetHiddenInRange hidPos, visPos, True

    Assert.IsTrue sut.IsHidden(sut.IndexOf("opt_hid_v1")), _
                  "An entry of the span that was already hidden stays hidden"
    Assert.IsTrue sut.IsHidden(sut.IndexOf("opt_vis_v1")), _
                  "And one that was visible hides with it"

    sut.SetHiddenInRange hidPos, visPos, False

    Assert.IsFalse sut.IsHidden(sut.IndexOf("opt_vis_v1")), _
                   "And the whole span shows again"
    Assert.IsFalse sut.IsHidden(sut.IndexOf("opt_hid_v1")), _
                   "Including the entry the dictionary started hidden"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetHiddenInRangeHidesEveryFreeEntry", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSetHiddenInRangeLeavesEntriesOutsideItAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestSetHiddenInRangeLeavesEntriesOutsideItAlone"
    If Not FixtureReady("TestSetHiddenInRangeLeavesEntriesOutsideItAlone") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim hidPos As Long
    Dim visPos As Long
    Dim mandPos As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    hidPos = PositionOf(sut, "opt_hid_v1")
    visPos = PositionOf(sut, "opt_vis_v1")
    mandPos = PositionOf(sut, "mand_v1")

    'Hiding one section must not reach the section beside it.
    sut.SetHiddenInRange hidPos, hidPos, True

    Assert.IsFalse sut.IsHidden(sut.IndexOf("opt_vis_v1")), _
                   "The entry one position past the span is untouched"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetHiddenInRangeLeavesEntriesOutsideItAlone", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestAMandatoryEntryHoldsItsSectionOpen()
    CustomTestSetTitles Assert, TESTMODULE, "TestAMandatoryEntryHoldsItsSectionOpen"
    If Not FixtureReady("TestAMandatoryEntryHoldsItsSectionOpen") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim hidPos As Long
    Dim visPos As Long
    Dim mandPos As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    hidPos = PositionOf(sut, "opt_hid_v1")
    visPos = PositionOf(sut, "opt_vis_v1")
    mandPos = PositionOf(sut, "mand_v1")

    sut.SetHiddenInRange hidPos, mandPos, True

    Assert.IsTrue sut.IsHidden(sut.IndexOf("opt_vis_v1")), _
                  "The free entries of the span hide"
    Assert.IsFalse sut.IsHidden(sut.IndexOf("mand_v1")), _
                   "A mandatory entry stays visible when its section hides"
    Assert.AreEqual CByte(ShowHideRangeHidden), sut.RangeState(hidPos, mandPos), _
                    "And the span still reads as hidden, because the state " & _
                    "is read from the entries the user owns"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAMandatoryEntryHoldsItsSectionOpen", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestSetHiddenInRangeCountsWhatItChanged()
    CustomTestSetTitles Assert, TESTMODULE, "TestSetHiddenInRangeCountsWhatItChanged"
    If Not FixtureReady("TestSetHiddenInRangeCountsWhatItChanged") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim hidPos As Long
    Dim visPos As Long
    Dim mandPos As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    hidPos = PositionOf(sut, "opt_hid_v1")
    visPos = PositionOf(sut, "opt_vis_v1")
    mandPos = PositionOf(sut, "mand_v1")

    'opt_hid_v1 starts hidden and opt_vis_v1 starts visible, so hiding the span
    'moves one entry. mand_v1 is in the span too and moves nothing.
    Assert.AreEqual CLng(1), sut.SetHiddenInRange(hidPos, mandPos, True), _
                    "The count is the entries that actually moved"
    Assert.AreEqual CLng(0), sut.SetHiddenInRange(hidPos, mandPos, True), _
                    "And hiding an already hidden span moves nothing"
    Assert.AreEqual CLng(2), sut.SetHiddenInRange(hidPos, mandPos, False), _
                    "Showing it again moves both free entries"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSetHiddenInRangeCountsWhatItChanged", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestRangeStateReadsTheFreeEntries()
    CustomTestSetTitles Assert, TESTMODULE, "TestRangeStateReadsTheFreeEntries"
    If Not FixtureReady("TestRangeStateReadsTheFreeEntries") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim hidPos As Long
    Dim visPos As Long
    Dim mandPos As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    hidPos = PositionOf(sut, "opt_hid_v1")
    visPos = PositionOf(sut, "opt_vis_v1")
    mandPos = PositionOf(sut, "mand_v1")

    Assert.AreEqual CByte(ShowHideRangeShown), sut.RangeState(visPos, visPos), _
                    "A span holding one visible free entry reads as shown"
    Assert.AreEqual CByte(ShowHideRangeMixed), sut.RangeState(hidPos, visPos), _
                    "One hidden and one visible reads as mixed"

    sut.SetHiddenInRange hidPos, visPos, True
    Assert.AreEqual CByte(ShowHideRangeHidden), sut.RangeState(hidPos, visPos), _
                    "And hiding the lot reads as hidden"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRangeStateReadsTheFreeEntries", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestASpanTheUserOwnsNothingInIsFixed()
    CustomTestSetTitles Assert, TESTMODULE, "TestASpanTheUserOwnsNothingInIsFixed"
    If Not FixtureReady("TestASpanTheUserOwnsNothingInIsFixed") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim hidPos As Long
    Dim visPos As Long
    Dim mandPos As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    hidPos = PositionOf(sut, "opt_hid_v1")
    visPos = PositionOf(sut, "opt_vis_v1")
    mandPos = PositionOf(sut, "mand_v1")

    'A span with nothing free in it answers Fixed, so a caller can tell "there is
    'nothing here to toggle" from "this is showing and you may hide it".
    Assert.AreEqual CByte(ShowHideRangeFixed), sut.RangeState(mandPos, mandPos), _
                    "A span holding only a mandatory entry is fixed"
    Assert.AreEqual CByte(ShowHideRangeEmpty), sut.RangeState(500, 510), _
                    "And a span holding no entry at all is empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASpanTheUserOwnsNothingInIsFixed", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestRangeStateReadsItsBoundsEitherWay()
    CustomTestSetTitles Assert, TESTMODULE, "TestRangeStateReadsItsBoundsEitherWay"
    If Not FixtureReady("TestRangeStateReadsItsBoundsEitherWay") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As ShowHide
    Dim hidPos As Long
    Dim visPos As Long
    Dim mandPos As Long

    Set sut = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    hidPos = PositionOf(sut, "opt_hid_v1")
    visPos = PositionOf(sut, "opt_vis_v1")
    mandPos = PositionOf(sut, "mand_v1")

    sut.SetHiddenInRange visPos, hidPos, True

    Assert.IsTrue sut.IsHidden(sut.IndexOf("opt_vis_v1")), _
                  "A span handed over backwards still hides what it covers"
    Assert.AreEqual sut.RangeState(hidPos, visPos), sut.RangeState(visPos, hidPos), _
                    "And reads the same state either way round"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRangeStateReadsItsBoundsEitherWay", Err.Number, Err.Description
End Sub

'@TestMethod("ShowHide")
Public Sub TestACollapsedSectionTravelsThroughTheStore()
    CustomTestSetTitles Assert, TESTMODULE, "TestACollapsedSectionTravelsThroughTheStore"
    If Not FixtureReady("TestACollapsedSectionTravelsThroughTheStore") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ShowHideStore
    Dim collapsed As ShowHide
    Dim reloaded As ShowHide
    Dim hidPos As Long
    Dim mandPos As Long

    'This is the whole of what "the layout is exported and imported" comes to.
    'Hiding a section writes nothing of its own: the choice lands on each member
    'entry, so the six columns the store already carries take a collapsed
    'section from one workbook to the next with no column and no code of their
    'own. LLExporter.AddShowHide and LLImporter.ImportShowHide drive exactly
    'this pair of calls.
    Set store = ShowHideStore.CreateOnSheet(ScratchSheet())
    Set collapsed = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    hidPos = PositionOf(collapsed, "opt_hid_v1")
    mandPos = PositionOf(collapsed, "mand_v1")

    collapsed.SetHiddenInRange hidPos, mandPos, True
    store.Save collapsed

    Set reloaded = ShowHide.Create(Dict, ShowHideLayerVList, VLIST_SHEET)
    Assert.IsFalse reloaded.IsHidden(reloaded.IndexOf("opt_vis_v1")), _
                   "A fresh list starts where the dictionary says"

    Assert.IsTrue store.Load(reloaded) > 0, "Load reports the rows it matched"

    Assert.AreEqual CByte(ShowHideRangeHidden), reloaded.RangeState(hidPos, mandPos), _
                    "The collapsed section comes back collapsed"
    Assert.IsFalse reloaded.IsHidden(reloaded.IndexOf("mand_v1")), _
                   "And its mandatory entry comes back visible"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestACollapsedSectionTravelsThroughTheStore", Err.Number, Err.Description
End Sub


'@section Fixture helpers
'===============================================================================

'@fun-title Where one variable of an entry list sits.
'@param entries ShowHide. The list to read.
'@param fieldKey String. The variable to look for.
'@return Long. The position, or 0 when the list does not carry the variable.
Private Function PositionOf(ByVal entries As ShowHide, ByVal fieldKey As String) As Long
    Dim idx As Long

    idx = entries.IndexOf(fieldKey)
    If idx = 0 Then Exit Function

    PositionOf = entries.PositionIndex(idx)
End Function

'@fun-title Report a fixture that could not be built, once per test.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is there.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If SetupError = 0 And Not Dict Is Nothing Then
        FixtureReady = True
        Exit Function
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The fixture could not be built - " & SetupMessage
End Function

'@fun-title Hand the running test the next worksheet of the pool.
'@details
'The sheet is hard-cleared on the way out, so a test writing geometry never
'sees what the one before it left. Nothing is deleted: Worksheet.Delete is
'unreliable on macOS Excel and hangs the headless run.
'@return Worksheet. An empty worksheet of the fixture workbook.
Private Function ScratchSheet() As Worksheet
    Dim sh As Worksheet

    ScratchTaken = ScratchTaken + 1
    If ScratchTaken > ScratchSheets.Count Then
        Err.Raise vbObjectError + 3001, "TestShowHide", _
                  "The scratch pool holds " & CStr(ScratchSheets.Count) & _
                  " sheets and this test asked for " & CStr(ScratchTaken)
    End If

    Set sh = ScratchSheets.Item(ScratchTaken)
    ResetScratchSheet sh
    Set ScratchSheet = sh
End Function

'@sub-title Put one pool sheet back to empty.
'@details
'Tables first, then the cells, then the row and column geometry the layout
'tests write. A hidden column or a width left behind would be read as the next
'test's own state.
'@param sh Worksheet. The sheet to clear.
Private Sub ResetScratchSheet(ByVal sh As Worksheet)
    Dim span As Range

    On Error Resume Next
        Do While sh.ListObjects.Count > 0
            sh.ListObjects(1).Delete
        Loop

        sh.Cells.Clear

        'The geometry reset is bounded. Writing a width to all 16,384 columns is
        'slow enough to be felt eighteen times over, and the largest position the
        'dictionary fixture produces is well under a hundred.
        Set span = sh.Range(sh.Cells(1, 1), sh.Cells(SCRATCH_RESET_SPAN, SCRATCH_RESET_SPAN))
        span.EntireColumn.Hidden = False
        span.EntireRow.Hidden = False
        span.EntireColumn.ColumnWidth = sh.StandardWidth
        span.EntireRow.RowHeight = sh.StandardHeight
        span.Orientation = 0
    On Error GoTo 0
End Sub

'@sub-title Write a hidden flag straight into the stored table.
'@param store ShowHideStore. The store holding the table.
'@param fieldKey String. The variable name whose row to change.
'@param flagValue String. What to put in the hidden_flag cell.
Private Sub SetStoredFlag(ByVal store As ShowHideStore, _
                          ByVal fieldKey As String, _
                          ByVal flagValue As String)
    Dim counter As Long
    Dim body As Range

    Set body = store.Table.DataBodyRange
    If body Is Nothing Then Exit Sub

    For counter = 1 To body.Rows.Count
        If StrComp(CStr(body.Cells(counter, 2).Value), fieldKey, vbTextCompare) = 0 Then
            body.Cells(counter, 4).Value = flagValue
            Exit Sub
        End If
    Next counter
End Sub
