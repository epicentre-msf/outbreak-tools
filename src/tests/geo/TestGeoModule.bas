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
'caption is built from. Over a filled geobase it covers the open of the picker
'in both scopes, the cascade of admin lists at levels 2 to 4 in both scopes,
'and the name prompt of the picker.
'
'THREE WINDOWS ARE HELD BACK FOR THE WHOLE MODULE
'-------------------------------------------------------------------------------
'GeoModule puts three windows on the screen and waits for a hand: the picker of
'LoadGeo, the failure box every handler reports through, and the name prompt of
'AddAdminName. A headless run has nobody to close one, so it stops there and
'writes no result file. ModuleInitialize raises all three seams before anything
'else runs and ModuleCleanup puts them back. The name prompt is answered
'through GeoStubAdminName: each AddAdminName test sets the answer it wants and
'TestCleanup puts the stub back on the cancel answer.
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
'THE SHAPE OF THE FIXTURE GEOBASE THE CASCADE RUNS OVER
'-------------------------------------------------------------------------------
'GeoTestFixture fills three admin 1 values (P1, P2 and the number 3), two
'admin 2 under each (D1 to D6), two admin 3 under each of those (C1 to C12)
'and one admin 4 per admin 3 (V1 to V12). The three health facilities HF1 to
'HF3 all sit under P1 and D1, one per commune C1 to C3. The assertions below
'name those values, so a change to the fixture shows up here first.
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
'THE NAME PROMPT NEEDS THE TRANSLATION WORKSHEET BACK FOR THE LENGTH OF ONE ACT
'-------------------------------------------------------------------------------
'AddAdminName reads the translation helper of the service before it walks the
'parents and raises when it answers Nothing. So the three AddAdminName tests
'put the worksheet back under its real name for the act alone, and TestCleanup
'renames it away again. While it stands, EventLinelist.Warn and .Fail show a
'real box, and that is what keeps two of the prompt's paths out of this suite:
'the warning on a parent with nothing selected and the warning on a duplicate
'name both end in a box with nobody to press OK. The three paths covered end
'before any box: the cancel, the blank answer and the name that lands.
'
'THE CURSOR INVARIANT IS MEASURED TWO WAYS
'-------------------------------------------------------------------------------
'The north-west arrow is the standing pointer of a linelist session.
'ShowAdminList and AddAdminName write it on every exit they own, so their tests
'park the pointer on xlWait first and assert the arrow came back. LoadGeo holds
'the pointer through the shared busy state, which puts back the pointer it found
'on the way in, so its tests park the arrow and assert the arrow.
'
'A SELECTION MADE BY HAND RUNS THE CLICK OF ITS LIST
'-------------------------------------------------------------------------------
'SelectInList sets the ListIndex of a picker list, and an MSForms list runs its
'Click on that, the way a click by hand does. So selecting P1 in LST_Adm1 runs
'the cascade to admin 2 during the arrange. Every arrange below selects the
'parents first and writes its stand-in entries after, so the cascade of the
'arrange leaves nothing the act is measured against.
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

'The admin 2 table of the fixture geobase, the one the name prompt writes to.
Private Const ADMIN2_TABLE As String = "T_ADM2"

'An entry written into a picker list by an arrange, so a test can tell a list
'the act emptied from a list the act left alone.
Private Const STAND_IN_ENTRY As String = "stand-in"

Private Assert As CustomTest

'Whether the picker came up. The tests that read a control of the form say so
'in their own line rather than raising into their handler.
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
'about to go. The translation worksheet is renamed away first, in case a test
'stopped while it stood under its real name.
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
    HideTranslationSheet
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
'next one a clean start. The name prompt goes back on the cancel answer, and
'the translation worksheet is put out of reach again, so a test that stood it
'up and then failed leaves the failure boxes shut for the next one.
'@TestCleanup
Public Sub TestCleanup()
    Application.Cursor = xlNorthwestArrow
    GeoModule.GeoStubAdminName False
    PutTranslationOutOfReach

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
'under a handler and what it answers is written into formIsSettled. The tests
'that read a control of the form assert that flag first, so a workbook this
'stops working on says so in a line of its own.
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

'@sub-title Give the translation worksheet its real name back for one act.
'@details
'AddAdminName asks the service for its translation helper and raises when it
'answers Nothing, so the three tests that drive the prompt call this before
'the act. The held service is dropped, because the flag that remembers a
'failed translation build lives on it and the next read builds afresh.
'While the worksheet stands under this name, every failure box of the
'linelist is a real box: the act it covers has to end before one.
Private Sub StandTranslationUp()
    On Error Resume Next
    ThisWorkbook.Worksheets(SETTLED_TRANSLATION_SHEET).Name = TRANSLATION_SHEET
    On Error GoTo 0

    LinelistEventsManager.DisposeEventLinelist
End Sub

'@sub-title Rename the translation worksheet away and drop the service holding it.
'@details
'The reverse of StandTranslationUp. TestCleanup calls it after every test, so
'a test that stood the worksheet up and stopped early leaves the boxes shut.
'The dispose matters as much as the rename: a held service keeps the helper it
'built while the worksheet stood, and its boxes with it.
Private Sub PutTranslationOutOfReach()
    HideTranslationSheet
    LinelistEventsManager.DisposeEventLinelist
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

'@fun-title The translation helper the service holds right now.
'@details
'The read AddAdminName makes before it walks the parents. The three prompt
'tests assert over it after StandTranslationUp, so a worksheet that stopped
'building a helper says so in its own line.
'@return LLTranslation. The helper, or Nothing.
Private Function TranslationNow() As LLTranslation
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Function

    Set TranslationNow = linelistEvents.Translation()
End Function

'@fun-title The number of data rows of one fixture geobase table.
'@param tableName String. The table to count.
'@return Long. Its data rows.
Private Function GeoTableRows(ByVal tableName As String) As Long
    GeoTableRows = ThisWorkbook.Worksheets(GEO_SHEET).ListObjects(tableName).ListRows.Count
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

'@sub-title Assert the act wrote nothing to the log.
'@param sinceRow Long. The last row the log held before the act.
'@param message String. What the assertion is about.
Private Sub AssertNoFailureLine(ByVal sinceRow As Long, ByVal message As String)
    Assert.AreEqual sinceRow, LogRowCount(), _
                    message & " - the lines read [" & LogTextAfter(sinceRow) & "]"
End Sub

'@sub-title Assert the standing pointer of the session came back.
'@param message String. What the assertion is about.
Private Sub AssertRestingPointer(ByVal message As String)
    Assert.AreEqual CLng(xlNorthwestArrow), CLng(Application.Cursor), message
End Sub


'@section Reading and writing the form
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

'@sub-title Leave one picker list holding one entry, selected.
'@details
'The way a parent is arranged for the cascade: the list holds the value alone
'and the value is selected, so ListValueOf in GeoModule reads it. Setting the
'ListIndex runs the Click of the list, which cascades to the level below the
'way a click by hand does. See the module description.
'@param listControl Object. The list to write.
'@param entry String. The one entry it holds afterwards.
Private Sub SelectInList(ByVal listControl As Object, ByVal entry As String)
    listControl.Clear
    listControl.AddItem entry
    listControl.ListIndex = 0
End Sub

'@sub-title Leave one picker list holding a stand-in entry, nothing selected.
'@details
'A list the act is expected to empty is given something to lose first.
'@param listControl Object. The list to write.
Private Sub FillWithStandIn(ByVal listControl As Object)
    listControl.Clear
    listControl.AddItem STAND_IN_ENTRY
End Sub

'@fun-title Whether one picker list holds an entry.
'@param listControl Object. The list to read.
'@param entry String. The entry looked for, compared without regard to case.
'@return Boolean. True when the list holds it.
Private Function ListHolds(ByVal listControl As Object, ByVal entry As String) As Boolean
    Dim counter As Long

    For counter = 0 To listControl.ListCount - 1
        If StrComp(CStr(listControl.List(counter)), entry, vbTextCompare) = 0 Then
            ListHolds = True
            Exit Function
        End If
    Next counter
End Function

'@fun-title The selected value of one picker list, as text.
'@details
'A list with nothing selected answers Null, which reads as an empty string
'here, the way ListValueOf in GeoModule reads it.
'@param listControl Object. The list to read.
'@return String. The selected value, or an empty string.
Private Function SelectedTextOf(ByVal listControl As Object) As String
    If IsNull(listControl.Value) Then Exit Function
    SelectedTextOf = CStr(listControl.Value)
End Function

'@sub-title Assert one picker list holds exactly the entries named.
'@details
'The entries are compared as a set, because the cascade answers a level in the
'order the geobase table holds it and the table is sorted on its concat column
'by any add. The count pins the size and each entry pins its presence.
'@param listControl Object. The list to read.
'@param listName String. The name of the list, for the message.
'@param entries Variant. The entries expected, as an array of strings.
Private Sub AssertListIs(ByVal listControl As Object, ByVal listName As String, _
                         ByVal entries As Variant)
    Dim counter As Long

    Assert.AreEqual CLng(UBound(entries) - LBound(entries) + 1), CLng(listControl.ListCount), _
                    listName & " should hold " & (UBound(entries) - LBound(entries) + 1) & " entries"

    For counter = LBound(entries) To UBound(entries)
        Assert.IsTrue ListHolds(listControl, CStr(entries(counter))), _
                      listName & " should hold " & CStr(entries(counter))
    Next counter
End Sub

'@sub-title Assert one picker list is empty.
'@param listControl Object. The list to read.
'@param message String. What the assertion is about.
Private Sub AssertListEmpty(ByVal listControl As Object, ByVal message As String)
    Assert.AreEqual CLng(0), CLng(listControl.ListCount), message
End Sub


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

    SelectInList F_Geo.LST_Adm1, "P1"
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


'@section The open of the picker over a geobase
'===============================================================================

'@sub-title The admin open fills the picker from the geobase.
'@details
'Arranges a geobase, a stale caption in the message box and the facility
'frame in front. Acts by opening the admin scope with the picker held back.
'Asserts the four level captions read what the manager answers for them, the
'admin 1 list holds the three admin 1 values, the historic list holds the two
'entries written for it, the geo frame is in front with the facility frame
'hidden, the message box is empty, no failure was logged and the standing
'pointer was left alone.
'@TestMethod("GeoModule")
Public Sub TestTheAdminOpenFillsThePickerFromTheGeobase()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheAdminOpenFillsThePickerFromTheGeobase"
    On Error GoTo TestFail

    Dim geoObj As LLGeo
    Dim logRows As Long

    ArrangeGeoManager
    Set geoObj = GeoManagerNow()
    Assert.IsNotNothing geoObj, "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"
    If geoObj Is Nothing Then Exit Sub

    With F_Geo
        .TXT_Msg.Value = "stale"
        .FRM_Geo.Visible = False
        .FRM_Facility.Visible = True
        .LBL_Geo1.Visible = False
        .LBL_Fac1.Visible = True
        .LST_Adm1.Clear
        .LST_Histo.Clear
    End With

    Application.Cursor = xlNorthwestArrow
    logRows = LogRowCount()

    GeoModule.LoadGeo GeoScopeAdmin

    AssertNoFailureLine logRows, "An admin open over a filled geobase logs no failure"

    With F_Geo
        Assert.AreEqual geoObj.GeoNames("adm1_name"), CStr(.LBL_Adm1.Caption), _
                        "The admin 1 caption reads the level name of the geobase"
        Assert.AreEqual geoObj.GeoNames("adm2_name"), CStr(.LBL_Adm2.Caption), _
                        "The admin 2 caption reads the level name of the geobase"
        Assert.AreEqual geoObj.GeoNames("adm3_name"), CStr(.LBL_Adm3.Caption), _
                        "The admin 3 caption reads the level name of the geobase"
        Assert.AreEqual geoObj.GeoNames("adm4_name"), CStr(.LBL_Adm4.Caption), _
                        "The admin 4 caption reads the level name of the geobase"
        Assert.IsTrue LenB(CStr(.LBL_Adm1.Caption)) > 0, _
                      "The admin 1 caption holds a name"

        AssertListIs .LST_Adm1, "LST_Adm1", Array("P1", "P2", "3")
        AssertListIs .LST_Histo, "LST_Histo", _
                     Array("P1 | D1 | C1 | V1", "P1 | D1 | C2 | V2")

        Assert.IsTrue .FRM_Geo.Visible, "The admin open shows the geo frame"
        Assert.IsFalse .FRM_Facility.Visible, "The admin open hides the facility frame"
        Assert.IsTrue .LBL_Geo1.Visible, "The admin open shows the geo label"
        Assert.IsFalse .LBL_Fac1.Visible, "The admin open hides the facility label"

        Assert.AreEqual vbNullString, CStr(.TXT_Msg.Value), _
                        "The open empties the message box"
    End With

    AssertRestingPointer "LoadGeo leaves the standing pointer where it found it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheAdminOpenFillsThePickerFromTheGeobase", _
                         Err.Number, Err.Description
End Sub

'@sub-title The facility open fills the picker from the facility table.
'@details
'Arranges a geobase, a stale caption in the message box and the geo frame in
'front. Acts by opening the facility scope with the picker held back. Asserts
'the four facility captions read what the manager answers, the facility admin
'1 list holds the one admin 1 the facilities sit under, the facility historic
'list holds its one entry, the facility frame is in front with the geo frame
'hidden, the message box is empty, no failure was logged and the standing
'pointer was left alone.
'@TestMethod("GeoModule")
Public Sub TestTheFacilityOpenFillsThePickerFromTheFacilityTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFacilityOpenFillsThePickerFromTheFacilityTable"
    On Error GoTo TestFail

    Dim geoObj As LLGeo
    Dim logRows As Long

    ArrangeGeoManager
    Set geoObj = GeoManagerNow()
    Assert.IsNotNothing geoObj, "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"
    If geoObj Is Nothing Then Exit Sub

    With F_Geo
        .TXT_Msg.Value = "stale"
        .FRM_Geo.Visible = True
        .FRM_Facility.Visible = False
        .LBL_Geo1.Visible = True
        .LBL_Fac1.Visible = False
        .LST_AdmF1.Clear
        .LST_HistoF.Clear
    End With

    Application.Cursor = xlNorthwestArrow
    logRows = LogRowCount()

    GeoModule.LoadGeo GeoScopeHF

    AssertNoFailureLine logRows, "A facility open over a filled geobase logs no failure"

    With F_Geo
        Assert.AreEqual geoObj.GeoNames("hf_name"), CStr(.LBL_Adm4F.Caption), _
                        "The facility caption reads the level name of the geobase"
        Assert.AreEqual geoObj.GeoNames("adm3_name"), CStr(.LBL_Adm3F.Caption), _
                        "The facility admin 3 caption reads the level name of the geobase"
        Assert.AreEqual geoObj.GeoNames("adm2_name"), CStr(.LBL_Adm2F.Caption), _
                        "The facility admin 2 caption reads the level name of the geobase"
        Assert.AreEqual geoObj.GeoNames("adm1_name"), CStr(.LBL_Adm1F.Caption), _
                        "The facility admin 1 caption reads the level name of the geobase"

        AssertListIs .LST_AdmF1, "LST_AdmF1", Array("P1")
        AssertListIs .LST_HistoF, "LST_HistoF", Array("HF1")

        Assert.IsTrue .FRM_Facility.Visible, "The facility open shows the facility frame"
        Assert.IsFalse .FRM_Geo.Visible, "The facility open hides the geo frame"
        Assert.IsTrue .LBL_Fac1.Visible, "The facility open shows the facility label"
        Assert.IsFalse .LBL_Geo1.Visible, "The facility open hides the geo label"

        Assert.AreEqual vbNullString, CStr(.TXT_Msg.Value), _
                        "The open empties the message box"
    End With

    AssertRestingPointer "LoadGeo leaves the standing pointer where it found it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFacilityOpenFillsThePickerFromTheFacilityTable", _
                         Err.Number, Err.Description
End Sub


'@section The cascade over a geobase
'===============================================================================

'@sub-title Admin 2 fills under one admin 1 and the caption is that admin 1.
'@details
'Arranges a geobase. Acts by asking for the admin 2 list under P2. Asserts the
'list holds the two districts of P2, the caption is P2, no failure was logged
'and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestTheCascadeFillsAdminTwoUnderOneParent()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheCascadeFillsAdminTwoUnderOneParent"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"

    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.ShowAdminList 2, "P2", GeoScopeAdmin, SEPARATOR

    AssertNoFailureLine logRows, "A cascade over a filled geobase logs no failure"
    AssertListIs F_Geo.LST_Adm2, "LST_Adm2", Array("D3", "D4")
    Assert.AreEqual "P2", CStr(F_Geo.TXT_Msg.Value), _
                    "The caption of admin 2 is the admin 1 clicked"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheCascadeFillsAdminTwoUnderOneParent", _
                         Err.Number, Err.Description
End Sub

'@sub-title Admin 3 fills under two parents and the caption reads admin 1 first.
'@details
'Arranges a geobase with P1 selected in the admin 1 list. Acts by asking for
'the admin 3 list under D2. Asserts the list holds the two communes of D2, the
'caption joins the parents in geo order, no failure was logged and the pointer
'is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestTheCascadeFillsAdminThreeInGeoOrder()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheCascadeFillsAdminThreeInGeoOrder"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"

    SelectInList F_Geo.LST_Adm1, "P1"
    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.ShowAdminList 3, "D2", GeoScopeAdmin, SEPARATOR

    AssertNoFailureLine logRows, "A cascade over a filled geobase logs no failure"
    AssertListIs F_Geo.LST_Adm3, "LST_Adm3", Array("C3", "C4")
    Assert.AreEqual "P1" & SEPARATOR & "D2", CStr(F_Geo.TXT_Msg.Value), _
                    "The geo caption reads admin 1 first"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheCascadeFillsAdminThreeInGeoOrder", _
                         Err.Number, Err.Description
End Sub

'@sub-title Admin 4 fills under three parents and the caption reads admin 1 first.
'@details
'Arranges a geobase with P1 and D1 selected in the admin 1 and admin 2 lists.
'Acts by asking for the admin 4 list under C2. Asserts the list holds the one
'village of C2, the caption joins the three parents in geo order, no failure
'was logged and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestTheCascadeFillsAdminFourInGeoOrder()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheCascadeFillsAdminFourInGeoOrder"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"

    SelectInList F_Geo.LST_Adm1, "P1"
    SelectInList F_Geo.LST_Adm2, "D1"
    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.ShowAdminList 4, "C2", GeoScopeAdmin, SEPARATOR

    AssertNoFailureLine logRows, "A cascade over a filled geobase logs no failure"
    AssertListIs F_Geo.LST_Adm4, "LST_Adm4", Array("V2")
    Assert.AreEqual "P1" & SEPARATOR & "D1" & SEPARATOR & "C2", CStr(F_Geo.TXT_Msg.Value), _
                    "The geo caption reads admin 1 first, down to the clicked value"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheCascadeFillsAdminFourInGeoOrder", _
                         Err.Number, Err.Description
End Sub

'@sub-title The facility cascade fills admin 3 and reads the caption deepest first.
'@details
'The facility caption joins the parents from the deepest level up, the order
'CMD_Copier_Click splits back out. The two caption branches of ShowAdminList
'are asymmetric and this is the first test that measures the facility one.
'Arranges a geobase with P1 selected in the facility admin 1 list. Acts by
'asking for the facility admin 3 list under D1. Asserts the list holds the
'three communes the facilities sit in, the caption reads D1 then P1, no
'failure was logged and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestTheFacilityCascadeFillsAdminThreeDeepestFirst()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFacilityCascadeFillsAdminThreeDeepestFirst"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"

    SelectInList F_Geo.LST_AdmF1, "P1"
    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.ShowAdminList 3, "D1", GeoScopeHF, SEPARATOR

    AssertNoFailureLine logRows, "A facility cascade over a filled geobase logs no failure"
    AssertListIs F_Geo.LST_AdmF3, "LST_AdmF3", Array("C1", "C2", "C3")
    Assert.AreEqual "D1" & SEPARATOR & "P1", CStr(F_Geo.TXT_Msg.Value), _
                    "The facility caption reads the deepest level first"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFacilityCascadeFillsAdminThreeDeepestFirst", _
                         Err.Number, Err.Description
End Sub

'@sub-title The facility cascade fills the facilities and reads the caption deepest first.
'@details
'Arranges a geobase with P1 and D1 selected in the facility admin 1 and admin
'2 lists. Acts by asking for the facility level 4 list under C1. Asserts the
'list holds the one facility of C1, the caption reads C1 then D1 then P1, no
'failure was logged and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestTheFacilityCascadeFillsTheFacilitiesDeepestFirst()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFacilityCascadeFillsTheFacilitiesDeepestFirst"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"

    SelectInList F_Geo.LST_AdmF1, "P1"
    SelectInList F_Geo.LST_AdmF2, "D1"
    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.ShowAdminList 4, "C1", GeoScopeHF, SEPARATOR

    AssertNoFailureLine logRows, "A facility cascade over a filled geobase logs no failure"
    AssertListIs F_Geo.LST_AdmF4, "LST_AdmF4", Array("HF1")
    Assert.AreEqual "C1" & SEPARATOR & "D1" & SEPARATOR & "P1", CStr(F_Geo.TXT_Msg.Value), _
                    "The facility caption reads the deepest level first, up to admin 1"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFacilityCascadeFillsTheFacilitiesDeepestFirst", _
                         Err.Number, Err.Description
End Sub

'@sub-title The cascade empties every list below the level it fills.
'@details
'The lists from the given level down hold children of a selection that just
'changed, so they are emptied before the level refills.
'Arranges a geobase with P1 selected in the admin 1 list and a stand-in entry
'in the admin 3 and admin 4 lists. Acts by asking for the admin 2 list under
'P1. Asserts admin 2 holds the two districts of P1 and admin 3 and admin 4 are
'empty.
'@TestMethod("GeoModule")
Public Sub TestTheCascadeEmptiesTheListsBelowTheLevelItFills()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheCascadeEmptiesTheListsBelowTheLevelItFills"
    On Error GoTo TestFail

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"

    SelectInList F_Geo.LST_Adm1, "P1"
    FillWithStandIn F_Geo.LST_Adm3
    FillWithStandIn F_Geo.LST_Adm4
    Application.Cursor = xlWait

    GeoModule.ShowAdminList 2, "P1", GeoScopeAdmin, SEPARATOR

    AssertListIs F_Geo.LST_Adm2, "LST_Adm2", Array("D1", "D2")
    AssertListEmpty F_Geo.LST_Adm3, "The cascade to admin 2 empties admin 3"
    AssertListEmpty F_Geo.LST_Adm4, "The cascade to admin 2 empties admin 4"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheCascadeEmptiesTheListsBelowTheLevelItFills", _
                         Err.Number, Err.Description
End Sub

'@sub-title A parent with no children leaves the level empty and logs nothing.
'@details
'The manager answers an empty table for a parent the geobase does not hold,
'and the cascade writes nothing into the list then: an MSForms list refuses
'an empty array. The list stays as the clear left it.
'Arranges a geobase. Acts by asking for the admin 2 list under a name the
'geobase does not hold. Asserts the list is empty, the caption still names the
'value clicked, no failure was logged and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestTheCascadeLeavesTheListEmptyUnderAnUnknownParent()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheCascadeLeavesTheListEmptyUnderAnUnknownParent"
    On Error GoTo TestFail

    Dim logRows As Long

    ArrangeGeoManager
    Assert.IsNotNothing GeoManagerNow(), _
                        "The geobase fixture builds a manager"
    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"

    FillWithStandIn F_Geo.LST_Adm2
    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.ShowAdminList 2, "Nowhere", GeoScopeAdmin, SEPARATOR

    AssertNoFailureLine logRows, "A parent with no children is an ordinary answer"
    AssertListEmpty F_Geo.LST_Adm2, "A parent with no children leaves admin 2 empty"
    Assert.AreEqual "Nowhere", CStr(F_Geo.TXT_Msg.Value), _
                    "The caption still names the value clicked"
    AssertRestingPointer "ShowAdminList puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheCascadeLeavesTheListEmptyUnderAnUnknownParent", _
                         Err.Number, Err.Description
End Sub


'@section The name prompt
'===============================================================================

'@fun-title The error number dropping the selection of one list raises.
'@details
'AddAdminName drops the selection of its level with ListIndex = -1. An MSForms
'list may run its Click on that, and the Click of an admin list hands its
'value to ShowAdminList, which is Null with nothing selected. The prompt tests
'try the same drop here first, while the failure boxes are still shut, so a
'raise ends in a line of their own.
'@param listControl Object. The list to deselect.
'@return Long. The error number the drop raised, zero when it raised nothing.
Private Function DeselectError(ByVal listControl As Object) As Long
    On Error Resume Next
    listControl.ListIndex = -1
    DeselectError = Err.Number
    On Error GoTo 0
End Function

'@fun-title Stand the picker at admin 2 under P1 with D1 selected.
'@details
'The arrange the three prompt tests share. P1 is selected in the admin 1 list
'and D1 in the admin 2 list, the admin 3 list holds a stand-in entry, and the
'translation worksheet is put back under its real name, because AddAdminName
'reads the helper before it walks the parents.
'The drop of the admin 2 selection is tried once before the worksheet comes
'back, see DeselectError. A drop that raises answers False and leaves the
'worksheet out of reach, so the caller stops before the act.
'The caller asserts the manager and the helper both answer before it acts.
'@return Boolean. True when the picker stands ready for the prompt.
Private Function PromptArrangedAtAdminTwo() As Boolean
    ArrangeGeoManager
    SelectInList F_Geo.LST_Adm1, "P1"
    SelectInList F_Geo.LST_Adm2, "D1"

    If DeselectError(F_Geo.LST_Adm2) <> 0 Then Exit Function
    SelectInList F_Geo.LST_Adm2, "D1"

    FillWithStandIn F_Geo.LST_Adm3
    StandTranslationUp
    PromptArrangedAtAdminTwo = True
End Function

'@sub-title A cancelled prompt changes nothing.
'@details
'A cancelled box answers the Boolean False, and the stub hands that over.
'The double click still leaves the user standing at the level with nothing
'chosen there, so the admin 3 list is empty and the admin 2 selection is
'dropped, and the caption shows the parent path.
'Arranges the picker at admin 2 under P1 with the prompt answering a cancel.
'Acts by asking to add a name at level 2. Asserts the admin 2 table kept its
'rows, the admin 3 list is empty, admin 2 has nothing selected, the caption is
'P1, no failure was logged and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestACancelledPromptChangesNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestACancelledPromptChangesNothing"
    On Error GoTo TestFail

    Dim promptReady As Boolean
    Dim rowsBefore As Long
    Dim logRows As Long

    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"
    promptReady = PromptArrangedAtAdminTwo()
    Assert.IsTrue promptReady, _
                  "Dropping the admin 2 selection raises nothing, so the prompt is reachable with the boxes live"
    If Not promptReady Then Exit Sub
    Assert.IsNotNothing GeoManagerNow(), "The geobase fixture builds a manager"
    Assert.IsNotNothing TranslationNow(), "The translation worksheet builds a helper"

    GeoModule.GeoStubAdminName False
    rowsBefore = GeoTableRows(ADMIN2_TABLE)
    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.AddAdminName 2

    AssertNoFailureLine logRows, "A cancelled prompt logs no failure"
    Assert.AreEqual rowsBefore, GeoTableRows(ADMIN2_TABLE), _
                    "A cancelled prompt writes no row"
    AssertListEmpty F_Geo.LST_Adm3, "The double click empties the level below"
    Assert.AreEqual CLng(-1), CLng(F_Geo.LST_Adm2.ListIndex), _
                    "The double click drops the selection of its level"
    Assert.AreEqual "P1", CStr(F_Geo.TXT_Msg.Value), _
                    "The caption shows the parent path"
    AssertRestingPointer "AddAdminName puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestACancelledPromptChangesNothing", _
                         Err.Number, Err.Description
End Sub

'@sub-title A blank answer changes nothing.
'@details
'Arranges the picker at admin 2 under P1 with the prompt answering spaces.
'Acts by asking to add a name at level 2. Asserts the admin 2 table kept its
'rows, the admin 3 list is empty, admin 2 has nothing selected, the caption is
'P1, no failure was logged and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestABlankAnswerChangesNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestABlankAnswerChangesNothing"
    On Error GoTo TestFail

    Dim promptReady As Boolean
    Dim rowsBefore As Long
    Dim logRows As Long

    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"
    promptReady = PromptArrangedAtAdminTwo()
    Assert.IsTrue promptReady, _
                  "Dropping the admin 2 selection raises nothing, so the prompt is reachable with the boxes live"
    If Not promptReady Then Exit Sub
    Assert.IsNotNothing GeoManagerNow(), "The geobase fixture builds a manager"
    Assert.IsNotNothing TranslationNow(), "The translation worksheet builds a helper"

    GeoModule.GeoStubAdminName "   "
    rowsBefore = GeoTableRows(ADMIN2_TABLE)
    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.AddAdminName 2

    AssertNoFailureLine logRows, "A blank answer logs no failure"
    Assert.AreEqual rowsBefore, GeoTableRows(ADMIN2_TABLE), _
                    "A blank answer writes no row"
    AssertListEmpty F_Geo.LST_Adm3, "The double click empties the level below"
    Assert.AreEqual CLng(-1), CLng(F_Geo.LST_Adm2.ListIndex), _
                    "The double click drops the selection of its level"
    Assert.AreEqual "P1", CStr(F_Geo.TXT_Msg.Value), _
                    "The caption shows the parent path"
    AssertRestingPointer "AddAdminName puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestABlankAnswerChangesNothing", _
                         Err.Number, Err.Description
End Sub

'@sub-title A good name lands in the geobase and comes back selected.
'@details
'The name is written under the parents, the admin 2 list refills with it in
'place, and it is selected, which runs the Click of the list and writes the
'caption down to it. The name is handed over with spaces around it, so the
'trim of the prompt answer is measured too.
'Arranges the picker at admin 2 under P1 with the prompt answering a name P1
'does not hold. Acts by asking to add a name at level 2. Asserts the admin 2
'table gained one row, the admin 2 list holds the two districts and the new
'name, the new name is the selected value, the caption opens on P1, no failure
'was logged and the pointer is back on the arrow.
'@TestMethod("GeoModule")
Public Sub TestAGoodNameLandsAndComesBackSelected()
    CustomTestSetTitles Assert, TESTMODULE, "TestAGoodNameLandsAndComesBackSelected"
    On Error GoTo TestFail

    Dim promptReady As Boolean
    Dim rowsBefore As Long
    Dim logRows As Long

    Assert.IsTrue formIsSettled, "The picker came up in ModuleInitialize"
    promptReady = PromptArrangedAtAdminTwo()
    Assert.IsTrue promptReady, _
                  "Dropping the admin 2 selection raises nothing, so the prompt is reachable with the boxes live"
    If Not promptReady Then Exit Sub
    Assert.IsNotNothing GeoManagerNow(), "The geobase fixture builds a manager"
    Assert.IsNotNothing TranslationNow(), "The translation worksheet builds a helper"

    GeoModule.GeoStubAdminName " DX "
    rowsBefore = GeoTableRows(ADMIN2_TABLE)
    Application.Cursor = xlWait
    logRows = LogRowCount()

    GeoModule.AddAdminName 2

    AssertNoFailureLine logRows, "A good name logs no failure"
    Assert.AreEqual rowsBefore + 1, GeoTableRows(ADMIN2_TABLE), _
                    "A good name writes one row into the admin 2 table"
    AssertListIs F_Geo.LST_Adm2, "LST_Adm2", Array("D1", "D2", "DX")
    Assert.AreEqual "DX", SelectedTextOf(F_Geo.LST_Adm2), _
                    "The new name comes back selected, trimmed"
    Assert.AreEqual "P1", Left$(CStr(F_Geo.TXT_Msg.Value), 2), _
                    "The caption opens on the parent path"
    AssertRestingPointer "AddAdminName puts the standing pointer back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAGoodNameLandsAndComesBackSelected", _
                         Err.Number, Err.Description
End Sub
