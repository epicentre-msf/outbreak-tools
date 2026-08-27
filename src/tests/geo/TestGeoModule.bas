Attribute VB_Name = "TestGeoModule"
Attribute VB_Description = "Tests for the GeoModule geo picker driver"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the GeoModule geo picker driver")

'@description
'Drives GeoModule, the module behind the F_Geo geo picker of a running
'linelist. This suite covers the paths a user reaches by pressing the wrong
'thing: a workbook whose geobase manager cannot be built, a scope the form has
'no layout for, and a cascade level outside the four admin levels. It also
'covers the tab strip the first open settles and the parent read the cascade
'caption is built from.
'
'THREE WINDOWS ARE HELD BACK FOR THE WHOLE MODULE
'-------------------------------------------------------------------------------
'GeoModule puts three windows on the screen and waits for a hand: the picker of
'LoadGeo, the failure box every handler reports through, and the name prompt of
'AddAdminName. A headless run has nobody to close one, so it stops there and
'writes no result file. ModuleInitialize raises all three seams before anything
'else runs and ModuleCleanup puts them back. The name prompt is held back as
'insurance: no test here reaches it, and a run that ever did would wedge.
'
'THE FAILURE IS READ BACK OFF THE __log WORKSHEET
'-------------------------------------------------------------------------------
'ReportGeoError reports through EventLinelist.Fail, which writes its line before
'it shows its box. GeoSuppressBox empties the fallback, which is what keeps the
'box shut, and the line is written either way. So every failure test here reads
'the rows the act added to __log and asserts the source and the reason
'GeoModule handed over. The rows before the act are counted first, because the
'driver workbook keeps its log between runs.
'
'THE GEOBASE MANAGER IS ARRANGED, BOTH WAYS
'-------------------------------------------------------------------------------
'EventLinelist builds its geobase manager off a worksheet named Geo in the
'workbook it was handed, and LinelistEventsManager hands it ThisWorkbook. So the
'manager is Nothing exactly while ThisWorkbook carries no such worksheet.
'ArrangeNoGeoManager deletes it, ArrangeGeoManager builds it through
'GeoTestFixture, and both drop the held service so the next read builds a fresh
'one. Every test says which of the two it wants, so the tests hold in any order.
'
'ONLY ONE OF THE ARGUMENT GUARDS IS REACHED WITHOUT A GEOBASE
'-------------------------------------------------------------------------------
'All three entry points read the manager before they look at their arguments, so
'the unknown scope and the bad levels sit behind a manager that answers. That is
'why the guard tests below arrange a geobase first.
'
'THE PICKER IS BROUGHT UP ONCE, IN ModuleInitialize
'-------------------------------------------------------------------------------
'F_Geo.UserForm_Initialize raises over a workbook with no translation helper, so
'the first mention of the form anywhere in the run raises. SettleTheForm below
'stands a translation worksheet up for one call, brings the form up on it, and
'renames the worksheet so EventLinelist stops finding it. Read its doc block
'before touching the module lifecycle: a translation worksheet left standing is
'what turns every failure box back on.
'
'THE CURSOR INVARIANT IS MEASURED TWO WAYS
'-------------------------------------------------------------------------------
'The north-west arrow is the standing pointer of a linelist session.
'ShowAdminList and AddAdminName write it on every exit they own, so their tests
'park the pointer on xlWait first and assert the arrow came back. LoadGeo holds
'the pointer through the shared busy state, which puts back the pointer it found
'on the way in, so its tests park the arrow and assert the arrow.
'@depends GeoModule, LinelistEventsManager, EventLinelist, LLGeo, HiddenNames, GeoTestFixture, CustomTest, TestHelpersLite, BetterArray

Private Const TESTMODULE As String = "TestGeoModule"
Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'The worksheet EventLinelist builds its geobase manager from.
Private Const GEO_SHEET As String = "Geo"

'LoadGeo builds its dropdown manager off this worksheet before it reads
'anything else, so a workbook without it fails there rather than on the
'manager the test is aiming at.
Private Const DROPDOWN_SHEET As String = "__dropdown_lists"

'The user log of the workbook, and the two columns the entries land in.
Private Const LOG_SHEET As String = "__log"
Private Const LOG_DETAIL_COLUMN As Long = 5

'The separator the picker joins an admin path with.
Private Const SEPARATOR As String = " | "

'The worksheet EventLinelist reads its translation helper from, the five tables
'that helper asks for, the hidden name holding the language code and the code
'itself. They stand up for the length of one call in ModuleInitialize and the
'worksheet is renamed straight after. See SettleTheForm.
Private Const TRANSLATION_SHEET As String = "LinelistTranslation"
Private Const SETTLED_TRANSLATION_SHEET As String = "__geo_trads"
Private Const LANGUAGE_NAME As String = "RNG_LLLanguageCode"
Private Const LANGUAGE_CODE As String = "ENG"

'The four workbook names GeoFormCache reads the search lists through. They are
'dropped and written again on every arrange: a name over a worksheet that was
'deleted keeps its entry and answers #REF.
Private Const CONCAT_GEO_NAME As String = "adm4_concat"
Private Const CONCAT_HF_NAME As String = "hf_concat"
Private Const HISTORIC_GEO_NAME As String = "histo_geo"
Private Const HISTORIC_HF_NAME As String = "histo_hf"

'The column of the geobase worksheet the three added lists are written in. The
'fixture tables stop well before it.
Private Const SPARE_COLUMN As Long = 60

Private Assert As CustomTest

'Whether the picker came up. The three tests that read a control of the form
'say so in their own line rather than raising into their handler.
Private formIsSettled As Boolean


'@section Module lifecycle
'===============================================================================

'@sub-title Initialise the test module before any tests run.
'@details
'The three seams are raised first, before any worksheet work that could raise,
'because a raise with the picker or the failure box still armed is what stops a
'headless run. The dropdown worksheet is left as it is when it is already
'there, so a suite that wrote one keeps what it wrote.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp

    GeoModule.GeoSuppressBox True
    GeoModule.GeoSuppressShow True
    GeoModule.GeoStubAdminName False

    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    EnsureWorksheet DROPDOWN_SHEET, ThisWorkbook, clearSheet:=False, _
                    visibility:=xlSheetHidden

    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName TESTMODULE

    SettleTheForm
End Sub

'@sub-title Tear down the module after all tests complete.
'@details
'The three seams go back to what production runs on. Module-level state in
'GeoModule outlives a test module, and a picker held back inside a later suite
'is a silent wrong-green.
'The geobase worksheet, the dropdown worksheet and the four form names are
'dropped, so the workbook comes out of this module holding what it held going
'in. The shared GeoFormCache instance is emptied for the same reason: LoadGeo
'loads it from ThisWorkbook, and the lists it holds are read off names that are
'about to go.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    GeoModule.GeoSuppressShow False
    GeoModule.GeoSuppressBox False
    GeoModule.GeoClearAdminNameStub

    Application.Cursor = xlNorthwestArrow

    GeoFormCache.Clear
    DropFormNames
    DropWorkbookName LANGUAGE_NAME
    DeleteWorksheet GEO_SHEET
    DeleteWorksheet DROPDOWN_SHEET
    DeleteWorksheet SETTLED_TRANSLATION_SHEET
    LinelistEventsManager.DisposeEventLinelist

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Reset state before each individual test.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Clean up after each individual test.
'@details
'Flushes the assertions of the test to the output sheet, and puts the standing
'pointer back so a test that parked it on xlWait and then failed leaves the
'next one a clean start.
'@TestCleanup
Public Sub TestCleanup()
    Application.Cursor = xlNorthwestArrow

    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Bringing the picker up
'===============================================================================

'@sub-title Build the picker once, while a translation helper can be reached.
'@details
'F_Geo.UserForm_Initialize asks EventLinelist for the translation helper of the
'workbook and raises when it answers Nothing. The helper is read off a
'worksheet named LinelistTranslation, and the driver workbook carries none, so
'the very first mention of F_Geo anywhere in the run raises. That is what
'stands between this suite and every control of the form.
'
'A translation worksheet that stays would cost more than it gives. It is also
'what carries EventLinelist.ShowMessage past its early exit, so every failure
'this suite drives would raise a real box and the run would stop on the first
'one with nobody to press OK.
'
'So the worksheet stands up for one call. The form takes its two translation
'objects in UserForm_Initialize and keeps them for the rest of the run, and the
'worksheet is renamed straight after: EventLinelist looks it up by name and
'finds nothing under the new one, while the tables the form holds stay alive
'under it. The held service is dropped either side, because the flag that
'remembers a failed translation build lives on the service.
'
'A raise here would cost the whole module its results, so the build is scoped
'under a handler and what it answers is written into formIsSettled. The three
'tests that read a control of the form assert that flag first, so a workbook
'this stops working on says so in a line of its own.
Private Sub SettleTheForm()
    Dim frameControl As Object

    'Scoped to the settling alone, and every step of it is allowed to fail:
    'what matters is the answer in formIsSettled, which the tests read.
    On Error Resume Next

    BuildTranslationSheet
    LinelistEventsManager.DisposeEventLinelist

    'The first mention of F_Geo is what runs UserForm_Initialize.
    Set frameControl = F_Geo.FRM_Geo
    formIsSettled = Not (frameControl Is Nothing)

    On Error GoTo 0

    HideTranslationSheet
    LinelistEventsManager.DisposeEventLinelist
End Sub

'@sub-title Stand up the translation worksheet the picker asks for.
'@details
'LLTranslation.Create asks for five tables by name and answers nothing about
'their contents, so each one is a header row and one empty row. The language
'code it reads the columns under lives in a workbook hidden name, and it raises
'when that name holds nothing.
Private Sub BuildTranslationSheet()
    Dim sh As Worksheet
    Dim store As HiddenNames

    Set sh = EnsureWorksheet(TRANSLATION_SHEET, ThisWorkbook, clearSheet:=True, _
                             visibility:=xlSheetHidden)

    AddTranslationTable sh, sh.Range("A1"), "T_TradLLMsg"
    AddTranslationTable sh, sh.Range("D1"), "T_TradLLShapes"
    AddTranslationTable sh, sh.Range("G1"), "T_TradLLForms"
    AddTranslationTable sh, sh.Range("J1"), "Tab_Translations"
    AddTranslationTable sh, sh.Range("M1"), "T_TradLLRibbon"

    Set store = HiddenNames.Create(ThisWorkbook)
    store.EnsureName LANGUAGE_NAME, LANGUAGE_CODE, HiddenNameTypeString
    store.SetValue LANGUAGE_NAME, LANGUAGE_CODE
End Sub

'@sub-title Write one empty translation table.
'@param sh Worksheet. The translation worksheet.
'@param startCell Range. The header cell of the table.
'@param tableName String. The name LLTranslation looks the table up under.
Private Sub AddTranslationTable(ByVal sh As Worksheet, ByVal startCell As Range, _
                                ByVal tableName As String)
    Dim tradTable As ListObject

    startCell.Value = "tag"
    startCell.Offset(0, 1).Value = LANGUAGE_CODE

    Set tradTable = sh.ListObjects.Add(xlSrcRange, startCell.Resize(2, 2), , xlYes)
    tradTable.Name = tableName
End Sub

'@sub-title Put the translation worksheet out of EventLinelist's reach.
'@details
'A rename rather than a delete: the form holds two translation objects over the
'tables of this worksheet and they have to stay readable, while EventLinelist
'resolves the worksheet by name and finds nothing under the new one.
Private Sub HideTranslationSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets(TRANSLATION_SHEET).Name = SETTLED_TRANSLATION_SHEET
    On Error GoTo 0
End Sub


'@section Arranging the geobase manager
'===============================================================================

'@sub-title Leave the workbook with no geobase manager to build.
'@details
'EventLinelist looks for a worksheet named Geo and answers Nothing when it
'finds none. The flag that remembers the failed build lives on the service, so
'the held service is dropped and the next read builds a fresh one.
Private Sub ArrangeNoGeoManager()
    DropFormNames
    DeleteWorksheet GEO_SHEET
    LinelistEventsManager.DisposeEventLinelist
End Sub

'@sub-title Give the workbook a geobase the manager can be built from.
'@details
'The fixture writes the nine tables LLGeo reads, filled with three admin 1
'values, two admin 2 under each, two admin 3 under each of those and one
'admin 4 per admin 3.
'It leaves the adm4_concat name behind and nothing for the other three, so the
'three are written here: LoadGeo asks GeoFormCache for both historic lists on
'every open, and a list built over a missing name is one more thing between a
'failing test and its reason.
Private Sub ArrangeGeoManager()
    Dim sh As Worksheet

    DropFormNames
    Set sh = GeoTestFixture.PrepareGeoFixture(GEO_SHEET, ThisWorkbook, _
                                              withData:=True)
    WriteFormNames sh
    LinelistEventsManager.DisposeEventLinelist
End Sub

'@sub-title Drop the four workbook names the geo form reads its lists through.
'@details
'A name written over a worksheet that is later deleted keeps its entry in the
'workbook and answers #REF, and EnsureConcatName in the fixture adds its name
'only when the workbook carries none. So the four are dropped by hand before
'every build.
Private Sub DropFormNames()
    DropWorkbookName CONCAT_GEO_NAME
    DropWorkbookName CONCAT_HF_NAME
    DropWorkbookName HISTORIC_GEO_NAME
    DropWorkbookName HISTORIC_HF_NAME
End Sub

'@sub-title Drop one workbook name when the workbook carries it.
'@param nameText String. The name to drop.
Private Sub DropWorkbookName(ByVal nameText As String)
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0
End Sub

'@sub-title Write the facility and historic lists the geo form reads.
'@details
'They sit in three spare columns of the geobase worksheet, well right of the
'nine tables, so a rebuild of the tables leaves them alone. The admin concat
'name comes from the fixture itself and is left as it is.
'@param sh Worksheet. The geobase worksheet.
Private Sub WriteFormNames(ByVal sh As Worksheet)
    WriteColumn sh.Cells(1, SPARE_COLUMN), "P1 | D1 | C1 | V1", "P1 | D1 | C2 | V2"
    sh.Parent.Names.Add Name:=HISTORIC_GEO_NAME, _
                        RefersTo:=sh.Range(sh.Cells(1, SPARE_COLUMN), _
                                           sh.Cells(2, SPARE_COLUMN))

    WriteColumn sh.Cells(1, SPARE_COLUMN + 1), "HF1", "HF2"
    sh.Parent.Names.Add Name:=CONCAT_HF_NAME, _
                        RefersTo:=sh.Range(sh.Cells(1, SPARE_COLUMN + 1), _
                                           sh.Cells(2, SPARE_COLUMN + 1))

    WriteColumn sh.Cells(1, SPARE_COLUMN + 2), "HF1"
    sh.Parent.Names.Add Name:=HISTORIC_HF_NAME, _
                        RefersTo:=sh.Range(sh.Cells(1, SPARE_COLUMN + 2), _
                                           sh.Cells(1, SPARE_COLUMN + 2))
End Sub

'@fun-title The geobase manager the module reads on every press.
'@details
'The same read GeoModule makes. Every failure test asserts over it first, so a
'workbook that stopped answering the way this module expects says so in its own
'line rather than through the assertions below it.
'@return LLGeo. The manager, or Nothing.
Private Function GeoManagerNow() As LLGeo
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Function

    Set GeoManagerNow = linelistEvents.GeoManager()
End Function


'@section Reading the log back
'===============================================================================

'@fun-title The number of log rows the workbook carries.
'@details
'Answers zero while the log worksheet is missing, which is the state before the
'first failure of the run builds it.
'@return Long. The last written row of the detail column.
Private Function LogRowCount() As Long
    Dim sh As Worksheet

    If Not WorksheetExists(LOG_SHEET) Then Exit Function

    Set sh = ThisWorkbook.Worksheets(LOG_SHEET)
    LogRowCount = sh.Cells(sh.Rows.Count, LOG_DETAIL_COLUMN).End(xlUp).Row
End Function

'@fun-title The log lines written after a given row, joined into one string.
'@details
'The log is a plain append, so everything below the row counted before the act
'was written by the act. Reading the rows added, rather than the whole sheet,
'is what keeps a line from an earlier test out of the answer.
'@param sinceRow Long. The last row the log held before the act.
'@return String. The detail cells written after that row.
Private Function LogTextAfter(ByVal sinceRow As Long) As String
    Dim sh As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim joined As String

    If Not WorksheetExists(LOG_SHEET) Then Exit Function

    Set sh = ThisWorkbook.Worksheets(LOG_SHEET)
    lastRow = sh.Cells(sh.Rows.Count, LOG_DETAIL_COLUMN).End(xlUp).Row

    For rowIndex = sinceRow + 1 To lastRow
        joined = joined & CStr(sh.Cells(rowIndex, LOG_DETAIL_COLUMN).Value) & vbLf
    Next rowIndex

    LogTextAfter = joined
End Function

'@sub-title Assert the log lines of one act name the source and the reason.
'@details
'ReportGeoError owns two pieces of the line: the procedure that failed and the
'reason behind it. The rest of the line is composed by LLLog, so the assertion
'reads the two pieces rather than the whole sentence.
'@param logText String. The lines the act added to the log.
'@param source String. The procedure the failure should name.
'@param reason String. The reason the failure should carry.
Private Sub AssertFailureLine(ByVal logText As String, ByVal source As String, _
                              ByVal reason As String)
    Assert.IsTrue InStr(1, logText, source, vbTextCompare) > 0, _
                  "The failure line should name " & source & _
                  " - the lines read [" & logText & "]"
    Assert.IsTrue InStr(1, logText, reason, vbTextCompare) > 0, _
                  "The failure line should carry the reason [" & reason & _
                  "] - the lines read [" & logText & "]"
End Sub

'@sub-title Assert the standing pointer of the session came back.
'@param message String. What the assertion is about.
Private Sub AssertRestingPointer(ByVal message As String)
    Assert.AreEqual CLng(xlNorthwestArrow), CLng(Application.Cursor), message
End Sub


'@section Reading the form
'===============================================================================

'@fun-title The tab strip of one frame of the picker.
'@details
'ShowFirstGeoPage finds it by type, because the two frames give their strips
'different names, and this reads it the same way.
'@param frameControl Object. The FRM_Geo or FRM_Facility frame.
'@return Object. The tab strip, or Nothing when the frame carries none.
Private Function TabStripOf(ByVal frameControl As Object) As Object
    Dim ctrl As Object

    For Each ctrl In frameControl.Controls
        If TypeName(ctrl) = "MultiPage" Then
            Set TabStripOf = ctrl
            Exit Function
        End If
    Next ctrl
End Function


'@section The manager that could not be built
'===============================================================================

'@sub-title LoadGeo over a workbook with no geobase reports and stops.
'@details
'Arranges a workbook with no geobase worksheet. Acts by opening the picker for
'the admin scope. Asserts the manager was Nothing, the failure line names
'LoadGeo and the reason, and the standing pointer was left alone.
'@TestMethod("GeoModule")
Public Sub TestLoadGeoReportsAManagerItCannotBuild()
    CustomTestSetTitles Assert, TESTMODULE, "TestLoadGeoReportsAManagerItCannotBuild"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeNoGeoManager
    Assert.IsNothing GeoManagerNow(), _
                     "A workbook with no geobase worksheet builds no manager"

    Application.Cursor = xlNorthwestArrow
    logRows = LogRowCount()

    GeoModule.LoadGeo GeoScopeAdmin

    AssertFailureLine LogTextAfter(logRows), "LoadGeo", _
                      "The geobase manager could not be built"
    AssertRestingPointer "LoadGeo leaves the standing pointer where it found it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLoadGeoReportsAManagerItCannotBuild", _
                         Err.Number, Err.Description
End Sub

'@sub-title ShowAdminList over a workbook with no geobase reports and stops.
'@details
'Arranges a workbook with no geobase worksheet. Acts by asking for the admin 2
'list. Asserts the failure line names ShowAdminList and the reason, and the
'pointer parked on the hourglass came back to the arrow.
'@TestMethod("GeoModule")
Public Sub TestShowAdminListReportsAManagerItCannotBuild()
    CustomTestSetTitles Assert, TESTMODULE, "TestShowAdminListReportsAManagerItCannotBuild"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeNoGeoManager
    Assert.IsNothing GeoManagerNow(), _
                     "A workbook with no geobase worksheet builds no manager"

    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.ShowAdminList 2, "P1", GeoScopeAdmin, SEPARATOR

    AssertFailureLine LogTextAfter(logRows), "ShowAdminList", _
                      "The geobase manager could not be built"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestShowAdminListReportsAManagerItCannotBuild", _
                         Err.Number, Err.Description
End Sub

'@sub-title AddAdminName over a workbook with no geobase reports and stops.
'@details
'Arranges a workbook with no geobase worksheet. Acts by asking to add a name at
'level 2. Asserts the failure line names AddAdminName and the reason, and the
'pointer parked on the hourglass came back to the arrow.
'@TestMethod("GeoModule")
Public Sub TestAddAdminNameReportsAManagerItCannotBuild()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddAdminNameReportsAManagerItCannotBuild"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeNoGeoManager
    Assert.IsNothing GeoManagerNow(), _
                     "A workbook with no geobase worksheet builds no manager"

    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.AddAdminName 2

    AssertFailureLine LogTextAfter(logRows), "AddAdminName", _
                      "The geobase manager could not be built"
    AssertRestingPointer "AddAdminName puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddAdminNameReportsAManagerItCannotBuild", _
                         Err.Number, Err.Description
End Sub


'@section The argument guards
'===============================================================================

'@sub-title LoadGeo refuses a scope the form has no layout for.
'@details
'GeoScopeBoth is on the geo scope enum and the picker carries no layout for it.
'An open under it used to configure nothing and show the form as the previous
'open left it.
'Arranges a geobase. Acts by opening the picker under GeoScopeBoth. Asserts the
'failure line names LoadGeo and the scope, and the standing pointer is back.
'@TestMethod("GeoModule")
Public Sub TestLoadGeoRefusesAScopeTheFormHasNoLayoutFor()
    CustomTestSetTitles Assert, TESTMODULE, "TestLoadGeoRefusesAScopeTheFormHasNoLayoutFor"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"

    Application.Cursor = xlNorthwestArrow
    logRows = LogRowCount()

    GeoModule.LoadGeo GeoScopeBoth

    AssertFailureLine LogTextAfter(logRows), "LoadGeo", _
                      "Unknown geo scope " & GeoScopeBoth
    AssertRestingPointer "LoadGeo leaves the standing pointer where it found it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLoadGeoRefusesAScopeTheFormHasNoLayoutFor", _
                         Err.Number, Err.Description
End Sub

'@sub-title ShowAdminList refuses a scope the form has no lists for.
'@details
'Arranges a geobase. Acts by asking for the admin 2 list under GeoScopeBoth.
'Asserts the failure line names ShowAdminList and the scope, and the pointer is
'back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestShowAdminListRefusesAScopeTheFormHasNoListsFor()
    CustomTestSetTitles Assert, TESTMODULE, "TestShowAdminListRefusesAScopeTheFormHasNoListsFor"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"

    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.ShowAdminList 2, "P1", GeoScopeBoth, SEPARATOR

    AssertFailureLine LogTextAfter(logRows), "ShowAdminList", _
                      "The cascade knows no scope " & GeoScopeBoth
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestShowAdminListRefusesAScopeTheFormHasNoListsFor", _
                         Err.Number, Err.Description
End Sub

'@sub-title ShowAdminList refuses level 1, which has no list above it.
'@details
'The cascade fills levels 2 to 4 from the levels above them. Level 1 is filled
'by the open itself and has no parent to read.
'Arranges a geobase. Acts by asking for level 1. Asserts the failure line names
'ShowAdminList and the level, and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestShowAdminListRefusesALevelBelowTheCascade()
    CustomTestSetTitles Assert, TESTMODULE, "TestShowAdminListRefusesALevelBelowTheCascade"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"

    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.ShowAdminList 1, "P1", GeoScopeAdmin, SEPARATOR

    AssertFailureLine LogTextAfter(logRows), "ShowAdminList", _
                      "The cascade knows no level 1"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestShowAdminListRefusesALevelBelowTheCascade", _
                         Err.Number, Err.Description
End Sub

'@sub-title ShowAdminList refuses a level past the four admin levels.
'@details
'Arranges a geobase. Acts by asking for level 5. Asserts the failure line names
'ShowAdminList and the level, and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestShowAdminListRefusesALevelAboveTheCascade()
    CustomTestSetTitles Assert, TESTMODULE, "TestShowAdminListRefusesALevelAboveTheCascade"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"

    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.ShowAdminList 5, "P1", GeoScopeAdmin, SEPARATOR

    AssertFailureLine LogTextAfter(logRows), "ShowAdminList", _
                      "The cascade knows no level 5"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestShowAdminListRefusesALevelAboveTheCascade", _
                         Err.Number, Err.Description
End Sub

'@sub-title AddAdminName refuses a level past the four admin levels.
'@details
'The picker adds a name at level 2, 3 or 4. Level 5 is refused before the name
'prompt is reached, which is what this measures.
'Arranges a geobase. Acts by asking to add a name at level 5. Asserts the
'failure line names AddAdminName and the level, and the pointer is back on the
'arrow.
'@TestMethod("GeoModule")
Public Sub TestAddAdminNameRefusesALevelAboveTheCascade()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddAdminNameRefusesALevelAboveTheCascade"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"

    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.AddAdminName 5

    AssertFailureLine LogTextAfter(logRows), "AddAdminName", _
                      "The picker adds no admin 5"
    AssertRestingPointer "AddAdminName puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddAdminNameRefusesALevelAboveTheCascade", _
                         Err.Number, Err.Description
End Sub


'@section The tab strip the first open settles
'===============================================================================

'@sub-title The first open settles the tab strip and the next one leaves it.
'@details
'The picker keeps its state between opens through its default instance, so the
'tab it comes up on the very first time is the tab the form was saved on. The
'first open of each frame puts the strip back on the four admin lists, and
'every later open gives the user back the tab they left.
'The flag that remembers this lives in GeoModule and nothing outside can put it
'back, so the two halves are one test: split in two they would hold only in the
'order they happen to run in.
'Arranges a geobase and holds the picker back. Acts by opening the admin scope,
'moving the strip to its second page, and opening again. Asserts the first open
'settled the strip and the second left it where the user put it.
'@TestMethod("GeoModule")
Public Sub TestTheFirstGeoOpenSettlesTheTabStripAndTheNextLeavesIt()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFirstGeoOpenSettlesTheTabStripAndTheNextLeavesIt"
    On Error GoTo TestFail

    Dim tabStrip As Object

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"

    GeoModule.LoadGeo GeoScopeAdmin

    Set tabStrip = TabStripOf(F_Geo.FRM_Geo)
    Assert.IsNotNothing tabStrip, "The geo frame of the picker carries a tab strip"
    If tabStrip Is Nothing Then Exit Sub

    Assert.AreEqual CLng(0), CLng(tabStrip.Value), _
                    "The first open puts the strip on the four admin lists"

    tabStrip.Value = 1
    GeoModule.LoadGeo GeoScopeAdmin

    Assert.AreEqual CLng(1), CLng(tabStrip.Value), _
                    "A later open gives the user back the tab they left"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFirstGeoOpenSettlesTheTabStripAndTheNextLeavesIt", _
                         Err.Number, Err.Description
End Sub


'@section The parent read behind the cascade caption
'===============================================================================

'@sub-title A parent list with nothing selected reads as an empty name.
'@details
'The caption of the picker joins the parents above the clicked value. A list
'with nothing selected answers Null, and the read turns that into an empty
'string, so the caption shows the separator with nothing in front of it.
'Arranges a geobase and empties the admin 1 list. Acts by asking for the
'admin 3 list under D1. Asserts the caption opens on the separator and the
'pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestTheCascadeCaptionReadsAnUnselectedParentAsEmpty()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheCascadeCaptionReadsAnUnselectedParentAsEmpty"
    On Error GoTo TestFail

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"

    F_Geo.LST_Adm1.Clear
    Application.Cursor = xlWait

    GeoModule.ShowAdminList 3, "D1", GeoScopeAdmin, SEPARATOR

    Assert.AreEqual SEPARATOR & "D1", CStr(F_Geo.TXT_Msg.Value), _
                    "An admin 1 with nothing selected reads as an empty name"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheCascadeCaptionReadsAnUnselectedParentAsEmpty", _
                         Err.Number, Err.Description
End Sub

'@sub-title A parent list with a selection reads as its value.
'@details
'Arranges a geobase and puts P1 in the admin 1 list, selected. Acts by asking
'for the admin 3 list under D1. Asserts the caption joins the selected parent
'with the clicked value, and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestTheCascadeCaptionReadsTheSelectedParent()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheCascadeCaptionReadsTheSelectedParent"
    On Error GoTo TestFail

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"

    With F_Geo.LST_Adm1
        .Clear
        .AddItem "P1"
        .ListIndex = 0
    End With
    Application.Cursor = xlWait

    GeoModule.ShowAdminList 3, "D1", GeoScopeAdmin, SEPARATOR

    Assert.AreEqual "P1" & SEPARATOR & "D1", CStr(F_Geo.TXT_Msg.Value), _
                    "The caption joins the selected parent with the clicked value"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheCascadeCaptionReadsTheSelectedParent", _
                         Err.Number, Err.Description
End Sub
