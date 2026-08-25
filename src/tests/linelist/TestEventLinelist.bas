Attribute VB_Name = "TestEventLinelist"
Attribute VB_Description = "Tests for EventLinelist class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for EventLinelist class")

'@description
'Drives EventLinelist, the class every event of a generated linelist comes
'through. It routes sheet changes, sheet deactivation, selection changes and
'double clicks, and it holds the domain managers those handlers need.
'
'WHAT THIS MODULE AIMS AT
'-------------------------------------------------------------------------------
'The held managers and the flags that guard them. A manager is built once per
'session, and a build that fails leaves its field Nothing; without a flag saying
'the attempt was made, the guard "the field is Nothing" reads as "never tried"
'and every keystroke rebuilds and re-fails the same manager. The translation
'helper is the one manager with a public reader, so it is the seam the flag can
'be measured through.
'
'THE DEACTIVATE TESTS COVER A SHIPPED BUG
'-------------------------------------------------------------------------------
'The workbook hidden name store used to be assigned at the foot of
'EnsureTranslation, below two early exits. A workbook with no
'LinelistTranslation worksheet took the exit at the sheet lookup and never
'reached the line, so the store stayed Nothing for the whole session and the
'list-auto flag was never cleared. TestDeactivateClearsTheListAutoFlagWithoutATranslationSheet
'is red against that code and green against the split.
'
'THE FIXTURE IS TWO WORKBOOKS
'-------------------------------------------------------------------------------
'A bare one, which is what a workbook missing its translation worksheet looks
'like, and a seeded one carrying the five translation tables and the two
'language codes. Several tests seed the bare workbook part way through, because
'the flag under test is what decides whether the class notices.
'
'THE DISPATCH TESTS COVER A SHIPPED BUG TOO
'-------------------------------------------------------------------------------
'The HList change handler used to resolve <table>_go_to_section as a range name
'above the header and multiple choice branches. The builder stores that id as a
'string, so the lookup raised on every sheet it makes and both branches were
'dead. The dispatch fixture carries no go-to caption on purpose: the toggle and
'the header restore working there is what proves the reorder.
'THE GEOBASE MANAGER TESTS COVER A REPORTED BUG
'-------------------------------------------------------------------------------
'LLGeo caches the five admin level labels on the instance, and the geo and form
'modules used to hold an LLGeo of their own that nothing dropped. After an
'in-session geobase import the F_Geo level captions read the previous geobase
'while the lists under them read the new one. GeoManager makes the service the
'one owner, so the ResetCaches every import already runs refreshes every reader.
'TestTheHeldGeoManagerKeepsItsLevelLabels states the hazard and
'TestResetCachesRereadsTheLevelLabels is the regression test.
'@depends EventLinelist, CustomTest, HiddenNames, LLTranslation, TranslationObject, LLLog, LLGeo, GeoTestFixture

Private Assert As CustomTest
Private FixtureWkb As Workbook

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "EventLinelist"
Private Const TRANS_SHEET_NAME As String = "LinelistTranslation"
'The flag lives in the hidden names of the SHEET that was edited. It used to be
'one workbook-level name, which could not say which sheet needed rebuilding.
Private Const LISTAUTO_FLAG As String = "update_listauto"

'The geobase worksheet the class looks for, and the T_NAMES table the level
'labels are translated from.
Private Const GEO_SHEET_NAME As String = "Geo"
Private Const GEO_NAMES_TABLE As String = "T_NAMES"

'The log worksheet and the two columns the funnel tests read: the output
'column carries the action and the outcome word, the detail column carries
'the text behind the date.
Private Const LOG_SHEET As String = "__log"

'The title the log writes above a block. LLLog heads a bundle with the section
'of the action rather than the action itself, and every code these tests use
'falls to the last case of that mapping. It is spelled out here rather than
'read off the class, because nothing in the project reads it that way and a
'test is the wrong place to be the first.
Private Const LOG_SECTION_LIFECYCLE As String = "linelist lifecycle"
Private Const LOG_OUTPUT_COLUMN As Long = 3
Private Const LOG_DETAIL_COLUMN As Long = 5

'What OnDoubleClick answers when the click asked for no geo picker.
Private Const GEOSCOPENONE As Long = -1

'@section Lifecycle
'===============================================================================

'@sub-title Set up the assertion harness.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestEventLinelist"
End Sub

'@sub-title Print results and tear down.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    HandBackTheScreen

    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Give the screen back to the workbook the harness writes into.
'@details
'Every test builds a workbook of its own, and writing a ListObject or a hidden
'name into it brings it to the front. CustomTest.PrintResults writes into a
'worksheet of this workbook and raises 1004 while another workbook holds the
'screen, and a raise inside a lifecycle hook is a modal dialog that stops the
'whole headless run.
Private Sub HandBackTheScreen()
    On Error Resume Next
        ThisWorkbook.Activate
    On Error GoTo 0
End Sub

'@sub-title Build a fresh bare workbook before each test.
'@details
'The workbook starts with no translation worksheet. A test that wants one calls
'SeedTranslationSheet, and several call it part way through on purpose.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureWkb = NewWorkbook()
End Sub

'@sub-title Flush assert state and drop the fixture workbook.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not FixtureWkb Is Nothing Then DeleteWorkbook FixtureWkb
    On Error GoTo 0

    Set FixtureWkb = Nothing
    HandBackTheScreen
End Sub


'@section Test Fixture Helpers
'===============================================================================

'@sub-title Build a translation worksheet with the five tables and two languages
'@details
'LLTranslation.Create validates all five tables, so a fixture short of one of
'them gives a Nothing helper for a reason that has nothing to do with the test.
'@param targetWkb Workbook. The workbook to seed.
Private Sub SeedTranslationSheet(ByVal targetWkb As Workbook)
    Dim sh As Worksheet

    Set sh = targetWkb.Worksheets.Add
    sh.Name = TRANS_SHEET_NAME

    SeedTable sh, 1, "T_TradLLMsg", "en", "fr", Array( _
        Array("MSG_GoToSection", "Go to section", "Aller a la section"), _
        Array("MSG_NotModify", "Do not modify", "Ne pas modifier"), _
        Array("MSG_Error", "Error", "Erreur"), _
        Array("MSG_ErrUpdate", "Update failed", "Echec de la mise a jour"))

    SeedTable sh, 7, "T_TradLLShapes", "en", "fr", _
              Array(Array("SHP_Advanced", "Advanced", "Avance"))
    SeedTable sh, 13, "T_TradLLForms", "en", "fr", _
              Array(Array("FRM_Title", "Form title", "Titre du formulaire"))
    SeedTable sh, 19, "Tab_Translations", "en", "fr", _
              Array(Array("DICT_Var1", "Variable 1", "Variable un"))
    SeedTable sh, 25, "T_TradLLRibbon", "en", "fr", _
              Array(Array("RIB_Advanced", "Ribbon advanced", "Ruban avance"))

    'TransObject reads both language codes from the WORKBOOK store.
    SetWorkbookName targetWkb, "RNG_LLLanguageCode", "en"
    SetWorkbookName targetWkb, "RNG_DictionaryLanguage", "en"
End Sub

'@sub-title Build the geobase worksheet the class looks for.
'@details
'LLGeo.Create refuses a sheet missing one of its nine tables, so the fixture is
'built whole. It carries data because the two label tests translate through it,
'and the five level labels land in the workbook store of the fixture workbook.
'@param targetWkb Workbook. The workbook to seed.
'@return Worksheet. The geobase worksheet.
Private Function SeedGeoSheet(ByVal targetWkb As Workbook) As Worksheet
    Set SeedGeoSheet = GeoTestFixture.PrepareGeoFixture(GEO_SHEET_NAME, targetWkb, _
                                                        withData:=True)
End Function

'@sub-title Give the admin1 level a new label, the way a geobase import does.
'@details
'An import reverts the headers, rewrites T_NAMES and translates again, and the
'instance that runs it is its own. This walks the same three steps through a
'second LLGeo over the one geobase worksheet, which is what puts a new label in
'the workbook store under any manager already held.
'@param geoSheet Worksheet. The geobase worksheet.
'@param newLabel String. The label admin1 takes.
Private Sub MoveAdmin1Label(ByVal geoSheet As Worksheet, ByVal newLabel As String)
    Dim mover As LLGeo

    Set mover = LLGeo.Create(geoSheet)
    mover.Translate rawNames:=True

    GeoTestFixture.GeoFixtureWriteCell geoSheet, _
                                       geoSheet.ListObjects(GEO_NAMES_TABLE), _
                                       1, 1, newLabel

    mover.Translate rawNames:=False
End Sub

'@sub-title Write one translation table and name it.
'@details
'The tables sit six columns apart. A table with a neighbour one column away has
'nowhere to grow into.
'@param sh Worksheet. The host worksheet.
'@param startColumn Long. The column the label column sits in.
'@param tableName String. The name to give the ListObject.
'@param langOne String. The header of the first value column.
'@param langTwo String. The header of the second value column.
'@param entries Variant. An array of label, first value and second value.
Private Sub SeedTable(ByVal sh As Worksheet, ByVal startColumn As Long, _
                      ByVal tableName As String, ByVal langOne As String, _
                      ByVal langTwo As String, ByVal entries As Variant)
    Dim idx As Long
    Dim lastRow As Long
    Dim rng As Range

    sh.Cells(1, startColumn).Value = "label"
    sh.Cells(1, startColumn + 1).Value = langOne
    sh.Cells(1, startColumn + 2).Value = langTwo

    For idx = LBound(entries) To UBound(entries)
        sh.Cells(idx + 2, startColumn).Value = entries(idx)(0)
        sh.Cells(idx + 2, startColumn + 1).Value = entries(idx)(1)
        sh.Cells(idx + 2, startColumn + 2).Value = entries(idx)(2)
    Next idx

    lastRow = UBound(entries) - LBound(entries) + 2
    Set rng = sh.Range(sh.Cells(1, startColumn), sh.Cells(lastRow, startColumn + 2))
    sh.ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, _
                       XlListObjectHasHeaders:=xlYes).Name = tableName
End Sub

'@sub-title Give a workbook-level hidden name its value.
'@param targetWkb Workbook. The workbook carrying the name.
'@param nameId String. The hidden name.
'@param value String. The value to store.
Private Sub SetWorkbookName(ByVal targetWkb As Workbook, _
                            ByVal nameId As String, ByVal value As String)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(targetWkb)
    If store.HasName(nameId) Then
        store.SetValue nameId, value
    Else
        store.EnsureName nameId, value, HiddenNameTypeString
    End If
End Sub

'@sub-title Read one workbook-level hidden name back.
'@param targetWkb Workbook. The workbook carrying the name.
'@param nameId String. The hidden name.
'@return String. The stored value.
Private Function WorkbookNameValue(ByVal targetWkb As Workbook, _
                                   ByVal nameId As String) As String
    Dim store As HiddenNames

    Set store = HiddenNames.Create(targetWkb)
    WorkbookNameValue = store.ValueAsString(nameId)
End Function

'@sub-title Give a sheet-level hidden name its value.
'@details
'The automatic lists are flagged on the sheet that was edited, not on the
'workbook, so the flag tests reach a worksheet store.
'@param targetSheet Worksheet. The sheet carrying the name.
'@param nameId String. The hidden name.
'@param value String. The value to store.
Private Sub SetSheetName(ByVal targetSheet As Worksheet, _
                         ByVal nameId As String, ByVal value As String)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(targetSheet)
    If store.HasName(nameId) Then
        store.SetValue nameId, value
    Else
        store.EnsureName nameId, value, HiddenNameTypeString
    End If
End Sub

'@sub-title Read one sheet-level hidden name back.
'@param targetSheet Worksheet. The sheet carrying the name.
'@param nameId String. The hidden name.
'@return String. The stored value.
Private Function SheetNameValue(ByVal targetSheet As Worksheet, _
                                ByVal nameId As String) As String
    Dim store As HiddenNames

    Set store = HiddenNames.Create(targetSheet)
    SheetNameValue = store.ValueAsString(nameId)
End Function

'@sub-title Write four admin values across one row and answer the first cell.
'@param values Variant. The values to write, left to right.
'@return Range. The leftmost cell of the run.
Private Function SeedGeoRow(ByVal values As Variant) As Range
    Dim sh As Worksheet
    Dim idx As Long

    Set sh = FixtureWkb.Worksheets(1)

    For idx = LBound(values) To UBound(values)
        sh.Cells(1, idx - LBound(values) + 1).Value = values(idx)
    Next idx

    Set SeedGeoRow = sh.Cells(1, 1)
End Function


'@sub-title Build an HList data entry sheet the change handler can dispatch on.
'@details
'A ListObject with its header on row 8, one variable column, the label cell
'above the header named after the variable the way VarWriter names it, and the
'three hidden names the dispatch reads: sheet_type, table_name and the
'variable's control string.
'@param varName String. The variable heading the single column.
'@param varControl String. The control string stored for that variable.
'@return Worksheet. The seeded sheet.
Private Function SeedHListSheet(ByVal varName As String, _
                                ByVal varControl As String) As Worksheet
    Dim sh As Worksheet
    Dim store As HiddenNames

    Set sh = FixtureWkb.Worksheets(1)

    sh.Cells(7, 2).Value = "Label of " & varName
    sh.Cells(8, 2).Value = varName
    sh.ListObjects.Add(SourceType:=xlSrcRange, _
                       Source:=sh.Range(sh.Cells(8, 2), sh.Cells(12, 2)), _
                       XlListObjectHasHeaders:=xlYes).Name = "table1"

    'The header refusal writes this name back into a scribbled header cell.
    FixtureWkb.Names.Add Name:=varName, _
                         RefersTo:="='" & sh.Name & "'!" & sh.Cells(7, 2).Address

    Set store = HiddenNames.Create(sh)
    store.EnsureName "sheet_type", "HList", HiddenNameTypeString
    store.EnsureName "table_name", "table1", HiddenNameTypeString
    store.EnsureName varName & " -- control", varControl, HiddenNameTypeString

    Set SeedHListSheet = sh
End Function


'@sub-title Build an HList sheet and the filtered companion its store names.
'@details
'The source table holds two columns with its header on row 8, the way the
'builder lays a data entry sheet out. The filtered sheet carries a header-only
'table at the same address, which is the layout UpdateFilterTables relies on.
'@param sourceRows Variant. One entry per data row: an array of two values.
'@return Worksheet. The filtered sheet.
Private Function SeedFilteredPair(ByVal sourceRows As Variant) As Worksheet
    Dim sh As Worksheet
    Dim filtsh As Worksheet
    Dim store As HiddenNames
    Dim rowIdx As Long
    Dim lastRow As Long

    Set sh = FixtureWkb.Worksheets(1)

    sh.Cells(8, 2).Value = "var1"
    sh.Cells(8, 3).Value = "var2"

    For rowIdx = LBound(sourceRows) To UBound(sourceRows)
        sh.Cells(9 + rowIdx - LBound(sourceRows), 2).Value = sourceRows(rowIdx)(0)
        sh.Cells(9 + rowIdx - LBound(sourceRows), 3).Value = sourceRows(rowIdx)(1)
    Next rowIdx

    lastRow = 9 + UBound(sourceRows) - LBound(sourceRows)
    sh.ListObjects.Add(SourceType:=xlSrcRange, _
                       Source:=sh.Range(sh.Cells(8, 2), sh.Cells(lastRow, 3)), _
                       XlListObjectHasHeaders:=xlYes).Name = "table1"

    Set filtsh = FixtureWkb.Worksheets.Add(After:=sh)
    filtsh.Name = "filtered1"
    filtsh.Cells(8, 2).Value = "var1"
    filtsh.Cells(8, 3).Value = "var2"
    filtsh.ListObjects.Add(SourceType:=xlSrcRange, _
                           Source:=filtsh.Range(filtsh.Cells(8, 2), filtsh.Cells(8, 3)), _
                           XlListObjectHasHeaders:=xlYes).Name = "ftable1"

    Set store = HiddenNames.Create(sh)
    store.EnsureName "sheet_type", "HList", HiddenNameTypeString
    store.EnsureName "table_name", "table1", HiddenNameTypeString
    store.EnsureName "filtered_sheet", "filtered1", HiddenNameTypeString

    Set SeedFilteredPair = filtsh
End Function

'@sub-title Build an HList sheet whose filtered companion cannot be found.
'@details
'The sheet carries a table and the HList tag with no filtered_sheet name, so
'the refresh cannot place its companion. This is the sheet the walk has to
'skip and name while the healthy pair still syncs.
'@param sheetName String. The name of the sheet to add.
'@return Worksheet. The seeded sheet.
Private Function SeedBrokenHListSheet(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet
    Dim store As HiddenNames

    Set sh = FixtureWkb.Worksheets.Add
    sh.Name = sheetName

    sh.Cells(8, 2).Value = "var1"
    sh.Cells(9, 2).Value = "kept"
    sh.ListObjects.Add SourceType:=xlSrcRange, _
                       Source:=sh.Range(sh.Cells(8, 2), sh.Cells(9, 2)), _
                       XlListObjectHasHeaders:=xlYes

    Set store = HiddenNames.Create(sh)
    store.EnsureName "sheet_type", "HList", HiddenNameTypeString

    Set SeedBrokenHListSheet = sh
End Function

'@sub-title The first row whose cell in a column of the log sheet holds a text.
'@details
'A loop rather than Range.Find, which inherits LookIn and SearchOrder from
'the last search of the Excel session. Answers 0 on a miss.
'@param sh Worksheet. The log worksheet.
'@param columnIndex Long. The column to read.
'@param searched String. The text to look for.
'@return Long. The first matching row, or 0.
Private Function LogRowOfText(ByVal sh As Worksheet, ByVal columnIndex As Long, _
                              ByVal searched As String) As Long
    Dim rowIndex As Long
    Dim lastRow As Long

    lastRow = sh.Cells(sh.Rows.Count, LOG_OUTPUT_COLUMN).End(xlUp).Row
    For rowIndex = 1 To lastRow
        If InStr(1, CStr(sh.Cells(rowIndex, columnIndex).Value), searched, vbTextCompare) > 0 Then
            LogRowOfText = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

'@fun-title The row of one variable in a VarLabelTable answer.
'@param labelRows BetterArray. The table VarLabelTable gave back.
'@param varName String. The variable name to look for.
'@return Variant. The row's array of title, name and label, or Empty.
Private Function VarLabelRowOf(ByVal labelRows As BetterArray, _
                               ByVal varName As String) As Variant
    Dim idx As Long
    Dim rowData As Variant

    For idx = labelRows.LowerBound To labelRows.UpperBound
        rowData = labelRows.Item(idx)
        If CStr(rowData(LBound(rowData) + 1)) = varName Then
            VarLabelRowOf = rowData
            Exit Function
        End If
    Next idx
End Function


'@section Factory Tests
'===============================================================================

'@sub-title A workbook gives an instance bound to it.
'@TestMethod("EventLinelist")
Public Sub TestCreateReturnsAnEventLinelist()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsAnEventLinelist"
    On Error GoTo TestFail

    Dim sut As EventLinelist

    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.IsTrue (Not sut Is Nothing), "Create gives back an instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsAnEventLinelist", Err.Number, Err.Description
End Sub

'@sub-title A Nothing workbook is refused, and the number says why.
'@TestMethod("EventLinelist")
Public Sub TestCreateRejectsNothingWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim errNumber As Long

    On Error Resume Next
        Set sut = EventLinelist.Create(Nothing)
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing workbook is refused and the number names the reason"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingWorkbook", Err.Number, Err.Description
End Sub


'@section The shared translation helper
'===============================================================================

'@sub-title A seeded workbook gives a translation helper.
'@TestMethod("EventLinelist")
Public Sub TestTranslationGivesTheHelper()
    CustomTestSetTitles Assert, TESTMODULE, "TestTranslationGivesTheHelper"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim helper As LLTranslation

    SeedTranslationSheet FixtureWkb
    Set sut = EventLinelist.Create(FixtureWkb)

    Set helper = sut.Translation

    Assert.IsTrue (Not helper Is Nothing), _
                  "A workbook carrying the five tables gives a helper"
    Assert.AreEqual TRANS_SHEET_NAME, helper.Wksh().Name, _
                    "The helper is bound to the translation worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTranslationGivesTheHelper", Err.Number, Err.Description
End Sub

'@sub-title The helper is held, so two reads give one object.
'@details
'This is what lets the button, ribbon and geo modules share it. Two objects
'here would mean each caller paid for its own build.
'@TestMethod("EventLinelist")
Public Sub TestTranslationIsHeldAcrossCalls()
    CustomTestSetTitles Assert, TESTMODULE, "TestTranslationIsHeldAcrossCalls"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstRead As LLTranslation
    Dim secondRead As LLTranslation

    SeedTranslationSheet FixtureWkb
    Set sut = EventLinelist.Create(FixtureWkb)

    Set firstRead = sut.Translation
    Set secondRead = sut.Translation

    Assert.IsTrue (Not firstRead Is Nothing), "The first read gives a helper"
    Assert.IsTrue (firstRead Is secondRead), _
                  "The second read gives the same object as the first"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTranslationIsHeldAcrossCalls", Err.Number, Err.Description
End Sub

'@sub-title A workbook with no translation worksheet gives Nothing, twice, quietly.
'@details
'An event handler has nothing above it to catch a raise, so a workbook in this
'state costs the events their labels and raises nothing.
'@TestMethod("EventLinelist")
Public Sub TestTranslationIsNothingWithoutTheSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestTranslationIsNothingWithoutTheSheet"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstRead As LLTranslation
    Dim secondRead As LLTranslation

    Set sut = EventLinelist.Create(FixtureWkb)

    Set firstRead = sut.Translation
    Set secondRead = sut.Translation

    Assert.IsTrue (firstRead Is Nothing), _
                  "A workbook with no translation worksheet gives Nothing"
    Assert.IsTrue (secondRead Is Nothing), _
                  "Asking a second time gives Nothing and raises nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTranslationIsNothingWithoutTheSheet", Err.Number, Err.Description
End Sub

'@sub-title A failed build is attempted once, and ResetCaches buys another go.
'@details
'This is the tried flag measured end to end. The worksheet appears between the
'second read and the third, and the class is meant to ignore it until it is
'asked to forget what it holds. Without the flag the second read would build a
'helper, which is the per-keystroke rebuild this guards against.
'@TestMethod("EventLinelist")
Public Sub TestTranslationStopsRetryingAfterAFailedBuild()
    CustomTestSetTitles Assert, TESTMODULE, "TestTranslationStopsRetryingAfterAFailedBuild"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim beforeSeed As LLTranslation
    Dim afterSeed As LLTranslation
    Dim afterReset As LLTranslation

    Set sut = EventLinelist.Create(FixtureWkb)

    Set beforeSeed = sut.Translation

    SeedTranslationSheet FixtureWkb
    Set afterSeed = sut.Translation

    sut.ResetCaches
    Set afterReset = sut.Translation

    Assert.IsTrue (beforeSeed Is Nothing), _
                  "The build fails while the workbook has no translation worksheet"
    Assert.IsTrue (afterSeed Is Nothing), _
                  "The failed build is not attempted again on the next read"
    Assert.IsTrue (Not afterReset Is Nothing), _
                  "ResetCaches clears the flag and the next read builds the helper"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTranslationStopsRetryingAfterAFailedBuild", Err.Number, Err.Description
End Sub

'@sub-title ResetCaches drops a helper that was built.
'@TestMethod("EventLinelist")
Public Sub TestResetCachesDropsTheHeldTranslation()
    CustomTestSetTitles Assert, TESTMODULE, "TestResetCachesDropsTheHeldTranslation"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstRead As LLTranslation
    Dim afterReset As LLTranslation

    SeedTranslationSheet FixtureWkb
    Set sut = EventLinelist.Create(FixtureWkb)

    Set firstRead = sut.Translation
    sut.ResetCaches
    Set afterReset = sut.Translation

    Assert.IsTrue (Not afterReset Is Nothing), "The helper is built again"
    Assert.IsTrue (Not (firstRead Is afterReset)), _
                  "The helper built after ResetCaches is a different object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestResetCachesDropsTheHeldTranslation", Err.Number, Err.Description
End Sub


'@section The shared geobase manager
'===============================================================================

'@sub-title A workbook carrying a geobase gives a geo manager.
'@TestMethod("EventLinelist")
Public Sub TestGeoManagerGivesTheManager()
    CustomTestSetTitles Assert, TESTMODULE, "TestGeoManagerGivesTheManager"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim manager As LLGeo

    SeedGeoSheet FixtureWkb
    Set sut = EventLinelist.Create(FixtureWkb)

    Set manager = sut.GeoManager()

    Assert.IsTrue (Not manager Is Nothing), _
                  "A workbook carrying the nine geobase tables gives a manager"
    Assert.AreEqual GEO_SHEET_NAME, manager.Wksh().Name, _
                    "The manager is bound to the geobase worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoManagerGivesTheManager", Err.Number, Err.Description
End Sub

'@sub-title The manager is held, so two reads give one object.
'@details
'This is what lets the geo and form modules share it. Two objects here would
'mean each caller paid for its own build, and LLGeo.Create walks two whole
'Names collections.
'@TestMethod("EventLinelist")
Public Sub TestGeoManagerIsHeldAcrossCalls()
    CustomTestSetTitles Assert, TESTMODULE, "TestGeoManagerIsHeldAcrossCalls"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstRead As LLGeo
    Dim secondRead As LLGeo

    SeedGeoSheet FixtureWkb
    Set sut = EventLinelist.Create(FixtureWkb)

    Set firstRead = sut.GeoManager()
    Set secondRead = sut.GeoManager()

    Assert.IsTrue (Not firstRead Is Nothing), "The first read gives a manager"
    Assert.IsTrue (firstRead Is secondRead), _
                  "The second read gives the same object as the first"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoManagerIsHeldAcrossCalls", Err.Number, Err.Description
End Sub

'@sub-title A workbook with no geobase worksheet gives Nothing, twice, quietly.
'@details
'The accessor swallows the failed build, so a module reading it guards on
'Nothing. The build used to raise, and the guard is what stands in for that.
'@TestMethod("EventLinelist")
Public Sub TestGeoManagerIsNothingWithoutTheSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestGeoManagerIsNothingWithoutTheSheet"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstRead As LLGeo
    Dim secondRead As LLGeo

    Set sut = EventLinelist.Create(FixtureWkb)

    Set firstRead = sut.GeoManager()
    Set secondRead = sut.GeoManager()

    Assert.IsTrue (firstRead Is Nothing), _
                  "A workbook with no geobase worksheet gives Nothing"
    Assert.IsTrue (secondRead Is Nothing), _
                  "Asking a second time gives Nothing and raises nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoManagerIsNothingWithoutTheSheet", Err.Number, Err.Description
End Sub

'@sub-title A failed build is attempted once, and ResetCaches buys another go.
'@details
'The tried flag measured end to end over a manager whose build reads a
'worksheet carrying nine ListObjects. The geobase appears between the second
'read and the third, and the class is meant to ignore it until it is asked to
'forget what it holds.
'@TestMethod("EventLinelist")
Public Sub TestGeoManagerStopsRetryingAfterAFailedBuild()
    CustomTestSetTitles Assert, TESTMODULE, "TestGeoManagerStopsRetryingAfterAFailedBuild"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim beforeSeed As LLGeo
    Dim afterSeed As LLGeo
    Dim afterReset As LLGeo

    Set sut = EventLinelist.Create(FixtureWkb)

    Set beforeSeed = sut.GeoManager()

    SeedGeoSheet FixtureWkb
    Set afterSeed = sut.GeoManager()

    sut.ResetCaches
    Set afterReset = sut.GeoManager()

    Assert.IsTrue (beforeSeed Is Nothing), _
                  "The build fails while the workbook has no geobase worksheet"
    Assert.IsTrue (afterSeed Is Nothing), _
                  "The failed build is not attempted again on the next read"
    Assert.IsTrue (Not afterReset Is Nothing), _
                  "ResetCaches clears the flag and the next read builds the manager"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoManagerStopsRetryingAfterAFailedBuild", Err.Number, Err.Description
End Sub

'@sub-title ResetCaches drops a manager that was built.
'@TestMethod("EventLinelist")
Public Sub TestResetCachesDropsTheHeldGeoManager()
    CustomTestSetTitles Assert, TESTMODULE, "TestResetCachesDropsTheHeldGeoManager"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstRead As LLGeo
    Dim afterReset As LLGeo

    SeedGeoSheet FixtureWkb
    Set sut = EventLinelist.Create(FixtureWkb)

    Set firstRead = sut.GeoManager()
    sut.ResetCaches
    Set afterReset = sut.GeoManager()

    Assert.IsTrue (Not afterReset Is Nothing), "The manager is built again"
    Assert.IsTrue (Not (firstRead Is afterReset)), _
                  "The manager built after ResetCaches is a different object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestResetCachesDropsTheHeldGeoManager", Err.Number, Err.Description
End Sub

'@sub-title A held manager keeps the level label it cached.
'@details
'LLGeo reads the five admin level labels once per instance. A geobase import
'runs through an instance of its own and moves the store under everything else,
'so the manager the service still holds answers the label of the geobase that
'has gone. This states the hazard the single owner exists to close.
'@TestMethod("EventLinelist")
Public Sub TestTheHeldGeoManagerKeepsItsLevelLabels()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheHeldGeoManagerKeepsItsLevelLabels"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim geoSheet As Worksheet
    Dim held As LLGeo
    Dim beforeMove As String

    Set geoSheet = SeedGeoSheet(FixtureWkb)
    Set sut = EventLinelist.Create(FixtureWkb)

    Set held = sut.GeoManager()
    held.Translate rawNames:=False
    beforeMove = held.GeoNames("adm1_name")

    MoveAdmin1Label geoSheet, "Region"

    Assert.AreEqual "Province", beforeMove, _
                    "The held manager reads the label T_NAMES carried at the start"
    Assert.AreEqual "Province", held.GeoNames("adm1_name"), _
                    "The held manager still answers the old label after the store moved"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheHeldGeoManagerKeepsItsLevelLabels", Err.Number, Err.Description
End Sub

'@sub-title ResetCaches makes the next read answer the new level label.
'@details
'This is the regression test for the reported caption bug. HandleImportGeobase
'follows every geobase import with ResetCaches on both its paths, and one owner
'means that one call refreshes every reader in the workbook.
'@TestMethod("EventLinelist")
Public Sub TestResetCachesRereadsTheLevelLabels()
    CustomTestSetTitles Assert, TESTMODULE, "TestResetCachesRereadsTheLevelLabels"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim geoSheet As Worksheet
    Dim held As LLGeo
    Dim afterReset As LLGeo
    Dim cachedLabel As String

    Set geoSheet = SeedGeoSheet(FixtureWkb)
    Set sut = EventLinelist.Create(FixtureWkb)

    Set held = sut.GeoManager()
    held.Translate rawNames:=False

    'The read is what loads the label cache of the held manager, which is the
    'state the import then moves the store under.
    cachedLabel = held.GeoNames("adm1_name")

    MoveAdmin1Label geoSheet, "Region"

    sut.ResetCaches
    Set afterReset = sut.GeoManager()

    Assert.AreEqual "Province", cachedLabel, _
                    "The held manager cached the label T_NAMES carried at the start"
    Assert.IsTrue (Not afterReset Is Nothing), "The manager is built again"
    Assert.AreEqual "Region", afterReset.GeoNames("adm1_name"), _
                    "The read after ResetCaches answers the label of the new geobase"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestResetCachesRereadsTheLevelLabels", Err.Number, Err.Description
End Sub


'@section Sheet deactivation
'===============================================================================

'@sub-title The list-auto flag is cleared on a workbook with no translation worksheet.
'@details
'This is the shipped bug. The workbook hidden name store used to be assigned at
'the foot of EnsureTranslation, below the exit taken when the translation
'worksheet is missing, so the store stayed Nothing and this flag was read and
'written by nothing for the whole session. Against that code the flag still
'reads "yes" here.
'@TestMethod("EventLinelist")
Public Sub TestDeactivateClearsTheListAutoFlagWithoutATranslationSheet()
    CustomTestSetTitles Assert, TESTMODULE, _
                        "TestDeactivateClearsTheListAutoFlagWithoutATranslationSheet"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim dataWksh As Worksheet

    Set dataWksh = FixtureWkb.Worksheets(1)
    SetSheetName dataWksh, LISTAUTO_FLAG, "yes"

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.OnSheetDeactivate dataWksh

    Assert.AreEqual "no", SheetNameValue(dataWksh, LISTAUTO_FLAG), _
                    "The flag is cleared even though the workbook has no translation worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, _
                         "TestDeactivateClearsTheListAutoFlagWithoutATranslationSheet", _
                         Err.Number, Err.Description
End Sub

'@sub-title A flag that is not "yes" is left where it stands.
'@TestMethod("EventLinelist")
Public Sub TestDeactivateLeavesTheFlagWhenItIsNotYes()
    CustomTestSetTitles Assert, TESTMODULE, "TestDeactivateLeavesTheFlagWhenItIsNotYes"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim dataWksh As Worksheet

    Set dataWksh = FixtureWkb.Worksheets(1)
    SetSheetName dataWksh, LISTAUTO_FLAG, "no"

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.OnSheetDeactivate dataWksh

    Assert.AreEqual "no", SheetNameValue(dataWksh, LISTAUTO_FLAG), _
                    "A flag reading no is left where it stands"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDeactivateLeavesTheFlagWhenItIsNotYes", _
                         Err.Number, Err.Description
End Sub

'@sub-title A Nothing sheet is ignored and raises nothing.
'@details
'The handler now takes the sheet object the worksheet event hands it, so
'Nothing is the one argument it has to refuse.
'@TestMethod("EventLinelist")
Public Sub TestDeactivateIgnoresNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestDeactivateIgnoresNothing"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim dataWksh As Worksheet
    Dim errNumber As Long

    Set dataWksh = FixtureWkb.Worksheets(1)
    SetSheetName dataWksh, LISTAUTO_FLAG, "yes"
    Set sut = EventLinelist.Create(FixtureWkb)

    On Error Resume Next
        sut.OnSheetDeactivate Nothing
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual 0&, errNumber, "A Nothing sheet raises nothing"
    Assert.AreEqual "yes", SheetNameValue(dataWksh, LISTAUTO_FLAG), _
                    "A Nothing sheet leaves the flag of every sheet alone"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDeactivateIgnoresNothing", Err.Number, Err.Description
End Sub


'@section Workbook open
'===============================================================================

'@sub-title Opening the workbook parks the pointer on the north-west arrow.
'@details
'The arrow is the standing cursor of a linelist session. Every busy state of
'the events shows the same arrow and ApplicationState restores the cursor it
'snapshots, so this one write at open is what keeps the pointer still while
'the user moves over a data entry sheet. The snapshot used to hold the default
'cursor, and every selection flicked the pointer twice.
'@TestMethod("EventLinelist")
Public Sub TestOpeningTheWorkbookParksThePointerOnTheArrow()
    CustomTestSetTitles Assert, TESTMODULE, "TestOpeningTheWorkbookParksThePointerOnTheArrow"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim heldCalculation As Long

    heldCalculation = Application.Calculation
    Application.Cursor = xlDefault

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.OnWorkbookOpen

    Assert.AreEqual CLng(xlNorthwestArrow), CLng(Application.Cursor), _
                    "The open parks the pointer on the north-west arrow"

    RestoreAfterOpen heldCalculation
    Exit Sub
TestFail:
    RestoreAfterOpen heldCalculation
    CustomTestLogFailure Assert, "TestOpeningTheWorkbookParksThePointerOnTheArrow", _
                         Err.Number, Err.Description
End Sub

'@sub-title Put back what OnWorkbookOpen changed on the application.
'@param heldCalculation Long. The calculation mode found before the call.
Private Sub RestoreAfterOpen(ByVal heldCalculation As Long)
    On Error Resume Next
    Application.Cursor = xlDefault
    If heldCalculation <> 0 Then Application.Calculation = heldCalculation
    Application.OnKey "^+g"
    On Error GoTo 0
End Sub


'@section Geobase recalculation
'===============================================================================

'@sub-title The geobase columns recalculate on demand and nothing else moves.
'@details
'The workbook runs on manual calculation, so a geobase import leaves the
'concat and p-code columns of the data entry sheets, and the admin level
'labels under ADM_UNIT_LIST, holding the values of the geobase before.
'RecalculateGeoColumns is what the import handler calls to refresh them. The
'plain column of the fixture keeps its stale value, which is what says the
'pass covers the geobase cells alone.
'@TestMethod("EventLinelist")
Public Sub TestRecalculateGeoColumnsRefreshesTheGeobaseCells()
    CustomTestSetTitles Assert, TESTMODULE, "TestRecalculateGeoColumnsRefreshesTheGeobaseCells"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim sh As Worksheet
    Dim heldCalculation As Long

    heldCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual

    Set sh = SeedGeoColumnSheet()

    'The three helper cells the formulas read, settled once.
    sh.Cells(1, 8).Value = 1
    sh.Cells(2, 8).Value = 2
    sh.Cells(3, 8).Value = 3
    sh.Calculate

    'The helper cells move, the way a geobase import moves the level tables.
    'Under manual calculation the formula cells hold their old values.
    sh.Cells(1, 8).Value = 10
    sh.Cells(2, 8).Value = 20
    sh.Cells(3, 8).Value = 30

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.RecalculateGeoColumns

    Assert.AreEqual 10#, CDbl(sh.Cells(9, 3).Value), _
                    "The concat column recalculates"
    Assert.AreEqual 20#, CDbl(sh.Cells(9, 4).Value), _
                    "The p-code column recalculates"
    Assert.AreEqual 3#, CDbl(sh.Cells(9, 5).Value), _
                    "The plain column keeps its value: the pass covers the " & _
                    "geobase cells alone"
    Assert.AreEqual 10#, CDbl(sh.Cells(1, 6).Value), _
                    "The admin level labels under ADM_UNIT_LIST recalculate"

    Application.Calculation = heldCalculation
    Exit Sub
TestFail:
    On Error Resume Next
    Application.Calculation = heldCalculation
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestRecalculateGeoColumnsRefreshesTheGeobaseCells", _
                         Err.Number, Err.Description
End Sub

'@sub-title Build an HList sheet carrying geobase formula columns.
'@details
'A data table with its header on row 8: an admin column, a concat column, a
'p-code column and a plain one, each formula column reading one helper cell in
'column H. Column F carries the four label cells the ADM_UNIT_LIST name
'covers, the first one a formula on the same helper cell as the concat.
'@return Worksheet. The seeded sheet.
Private Function SeedGeoColumnSheet() As Worksheet
    Dim sh As Worksheet
    Dim store As HiddenNames

    Set sh = FixtureWkb.Worksheets(1)

    sh.Cells(8, 2).Value = "adm1_geo"
    sh.Cells(8, 3).Value = "concat_adm1_geo"
    sh.Cells(8, 4).Value = "pcode_adm1_geo"
    sh.Cells(8, 5).Value = "plain_var"
    sh.Cells(9, 2).Value = "P1"
    sh.Cells(9, 3).Formula = "=$H$1"
    sh.Cells(9, 4).Formula = "=$H$2"
    sh.Cells(9, 5).Formula = "=$H$3"

    sh.ListObjects.Add(SourceType:=xlSrcRange, _
                       Source:=sh.Range(sh.Cells(8, 2), sh.Cells(9, 5)), _
                       XlListObjectHasHeaders:=xlYes).Name = "geotable"

    sh.Cells(1, 6).Formula = "=$H$1"
    FixtureWkb.Names.Add Name:="ADM_UNIT_LIST", _
                         RefersTo:="='" & sh.Name & "'!$F$1:$F$4"

    Set store = HiddenNames.Create(sh)
    store.EnsureName "sheet_type", "HList", HiddenNameTypeString

    Set SeedGeoColumnSheet = sh
End Function


'@section Double click routing
'===============================================================================

'@sub-title A sheet that is not a spatio-temporal analysis asks for no picker.
'@TestMethod("EventLinelist")
Public Sub TestDoubleClickAnswersNoneOnAPlainSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestDoubleClickAnswersNoneOnAPlainSheet"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim plainSheet As Worksheet

    Set plainSheet = FixtureWkb.Worksheets(1)
    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.AreEqual GEOSCOPENONE, sut.OnDoubleClick(plainSheet, plainSheet.Cells(1, 1)), _
                    "A sheet with no sheet_type asks for no picker"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDoubleClickAnswersNoneOnAPlainSheet", _
                         Err.Number, Err.Description
End Sub

'@sub-title Nothing arguments ask for no picker.
'@TestMethod("EventLinelist")
Public Sub TestDoubleClickAnswersNoneForNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestDoubleClickAnswersNoneForNothing"
    On Error GoTo TestFail

    Dim sut As EventLinelist

    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.AreEqual GEOSCOPENONE, sut.OnDoubleClick(Nothing, Nothing), _
                    "A click with no sheet and no cell asks for no picker"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDoubleClickAnswersNoneForNothing", _
                         Err.Number, Err.Description
End Sub


'@section HList change dispatch
'===============================================================================

'@sub-title A choice edit appends to the value the cell held.
'@details
'The previous value comes from the copy taken when the cell was selected, so
'the toggle costs no Application.Undo round trip. The sheet carries no go-to
'caption, which is the state that used to send every in-table edit into the
'error label before the toggle ran.
'@TestMethod("EventLinelist")
Public Sub TestAChoiceEditAppendsToTheHeldValue()
    CustomTestSetTitles Assert, TESTMODULE, "TestAChoiceEditAppendsToTheHeldValue"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim sh As Worksheet
    Dim cellRng As Range

    Set sh = SeedHListSheet("mchoice_var", "choice_multiple")
    Set sut = EventLinelist.Create(FixtureWkb)
    Set cellRng = sh.Cells(9, 2)

    cellRng.Value = "a"
    sut.OnSelectionChange sh, cellRng
    cellRng.Value = "b"
    sut.OnSheetChange sh, cellRng

    Assert.AreEqual "a, b", CStr(cellRng.Value), _
                    "The new pick lands after the value the cell held"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAChoiceEditAppendsToTheHeldValue", _
                         Err.Number, Err.Description
End Sub

'@sub-title Picking a choice the cell already holds restores the cell.
'@TestMethod("EventLinelist")
Public Sub TestAChoiceAlreadyPickedIsKeptOnce()
    CustomTestSetTitles Assert, TESTMODULE, "TestAChoiceAlreadyPickedIsKeptOnce"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim sh As Worksheet
    Dim cellRng As Range

    Set sh = SeedHListSheet("mchoice_var", "choice_multiple")
    Set sut = EventLinelist.Create(FixtureWkb)
    Set cellRng = sh.Cells(9, 2)

    cellRng.Value = "a, b"
    sut.OnSelectionChange sh, cellRng
    cellRng.Value = "a"
    sut.OnSheetChange sh, cellRng

    Assert.AreEqual "a, b", CStr(cellRng.Value), _
                    "A choice picked a second time keeps the cell as it stood"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAChoiceAlreadyPickedIsKeptOnce", _
                         Err.Number, Err.Description
End Sub

'@sub-title A scribbled header cell is restored from the label cell's name.
'@details
'Red against the old branch order: the go-to lookup above this branch raised
'on a sheet with no go-to range and the restore never ran.
'@TestMethod("EventLinelist")
Public Sub TestAHeaderEditIsRestored()
    CustomTestSetTitles Assert, TESTMODULE, "TestAHeaderEditIsRestored"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim sh As Worksheet
    Dim cellRng As Range

    Set sh = SeedHListSheet("mchoice_var", "choice_multiple")
    Set sut = EventLinelist.Create(FixtureWkb)
    Set cellRng = sh.Cells(8, 2)

    cellRng.Value = "scribble"
    sut.OnSheetChange sh, cellRng

    Assert.AreEqual "mchoice_var", CStr(cellRng.Value), _
                    "The header cell gets the variable name back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAHeaderEditIsRestored", _
                         Err.Number, Err.Description
End Sub


'@section Geo concatenation
'===============================================================================

'@sub-title Each admin level joins the cells to its left with a pipe.
'@TestMethod("EventLinelist")
Public Sub TestGeoConcatJoinsTheLevelsItIsGiven()
    CustomTestSetTitles Assert, TESTMODULE, "TestGeoConcatJoinsTheLevelsItIsGiven"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstCell As Range

    Set firstCell = SeedGeoRow(Array("Region", "District", "Chiefdom", "Village"))
    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.AreEqual "Region", sut.GeoConcat(firstCell, 1), _
                    "One level answers the cell itself"
    Assert.AreEqual "Region | District", sut.GeoConcat(firstCell, 2), _
                    "Two levels join the cell and its neighbour"
    Assert.AreEqual "Region | District | Chiefdom", sut.GeoConcat(firstCell, 3), _
                    "Three levels join three cells"
    Assert.AreEqual "Region | District | Chiefdom | Village", sut.GeoConcat(firstCell, 4), _
                    "Four levels join four cells"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoConcatJoinsTheLevelsItIsGiven", _
                         Err.Number, Err.Description
End Sub

'@sub-title A run with a gap in it answers empty.
'@details
'A half-filled geo row would otherwise concatenate to a value that looks real
'and matches no geobase entry.
'@TestMethod("EventLinelist")
Public Sub TestGeoConcatIsEmptyWhenALevelIsMissing()
    CustomTestSetTitles Assert, TESTMODULE, "TestGeoConcatIsEmptyWhenALevelIsMissing"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstCell As Range

    Set firstCell = SeedGeoRow(Array("Region", vbNullString, "Chiefdom", "Village"))
    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.AreEqual vbNullString, sut.GeoConcat(firstCell, 2), _
                    "Two levels with the second empty answer empty"
    Assert.AreEqual vbNullString, sut.GeoConcat(firstCell, 3), _
                    "Three levels with one empty answer empty"
    Assert.AreEqual vbNullString, sut.GeoConcat(firstCell, 4), _
                    "Four levels with one empty answer empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoConcatIsEmptyWhenALevelIsMissing", _
                         Err.Number, Err.Description
End Sub

'@sub-title A level outside 1 to 4 answers empty.
'@details
'The formula writer only authors levels 1 to 4, so this branch is only
'reachable from a hand-typed formula. It used to answer the first cell alone,
'a value that looks real and matches no geobase entry.
'@TestMethod("EventLinelist")
Public Sub TestGeoConcatAnswersEmptyOutsideTheLevels()
    CustomTestSetTitles Assert, TESTMODULE, "TestGeoConcatAnswersEmptyOutsideTheLevels"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstCell As Range

    Set firstCell = SeedGeoRow(Array("Region", "District", "Chiefdom", "Village"))
    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.AreEqual vbNullString, sut.GeoConcat(firstCell, 0), _
                    "Level zero answers empty on a fully filled row"
    Assert.AreEqual vbNullString, sut.GeoConcat(firstCell, 5), _
                    "A level past four answers empty on a fully filled row"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestGeoConcatAnswersEmptyOutsideTheLevels", _
                         Err.Number, Err.Description
End Sub


'@section Spatial name cleaning
'===============================================================================
'The spatial formulas carry the variable name with the tag the formula writer
'put in front -- "concat_adm1_" or "hf_" -- and the spatial tables are keyed on
'the bare name. BareSpatialName takes the tag off whole. The delegates used to
'split on "_" and take the third piece, which cut "concat_adm1_case_zone" down
'to "case", and to count the tag's characters off the front, which raised on
'any name shorter than the tag.

'@sub-title The authored tag comes off whole and nothing else does.
'@TestMethod("EventLinelist")
Public Sub TestBareSpatialNameTakesTheTagOffWhole()
    CustomTestSetTitles Assert, TESTMODULE, "TestBareSpatialNameTakesTheTagOffWhole"
    On Error GoTo TestFail

    Dim sut As EventLinelist

    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.AreEqual "myvar", sut.BareSpatialName("concat_adm1_myvar"), _
                    "The geo tag comes off the front"
    Assert.AreEqual "case_zone", sut.BareSpatialName("concat_adm1_case_zone"), _
                    "A bare name carrying underscores survives whole"
    Assert.AreEqual "myvar", sut.BareSpatialName("concat_adm4_myvar"), _
                    "The geo tag comes off at every level. ChangeAdminLevel " & _
                    "rewrites the level inside the quoted argument, and the " & _
                    "strip used to know the adm1 spelling alone"
    Assert.AreEqual "myvar", sut.BareSpatialName("concat_adm2_myvar"), _
                    "and level 2 strips the same way"
    Assert.AreEqual "facility", sut.BareSpatialName("hf_facility"), _
                    "The facility tag comes off the front"
    Assert.AreEqual "health_post", sut.BareSpatialName("hf_health_post"), _
                    "A facility name carrying underscores survives whole"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBareSpatialNameTakesTheTagOffWhole", _
                         Err.Number, Err.Description
End Sub

'@sub-title A name with no tag passes through, however short.
'@details
'The character-counting form raised "invalid procedure call" on any name
'shorter than the tag it counted off. A short or untagged name is not this
'member's error to raise: it passes through, and the table lookup answers its
'usual empty string.
'@TestMethod("EventLinelist")
Public Sub TestBareSpatialNamePassesUntaggedNamesThrough()
    CustomTestSetTitles Assert, TESTMODULE, "TestBareSpatialNamePassesUntaggedNamesThrough"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim errNumber As Long
    Dim shortAnswer As String

    Set sut = EventLinelist.Create(FixtureWkb)

    On Error Resume Next
        shortAnswer = sut.BareSpatialName("ad")
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual 0&, errNumber, "A name shorter than the tag raises nothing"
    Assert.AreEqual "ad", shortAnswer, "And comes back unchanged"
    Assert.AreEqual "plainvar", sut.BareSpatialName("plainvar"), _
                    "A name with no tag passes through unchanged"
    Assert.AreEqual vbNullString, sut.BareSpatialName("concat_adm1_"), _
                    "A tag with nothing behind it answers empty"
    Assert.AreEqual "concat_admin_myvar", sut.BareSpatialName("concat_admin_myvar"), _
                    "A name whose tag carries no level digit passes through unchanged"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBareSpatialNamePassesUntaggedNamesThrough", _
                         Err.Number, Err.Description
End Sub


'@section The managers the buttons read
'===============================================================================
'The button module used to build its own workbook store, its own password
'manager and one worksheet store per reader, on every click. It reads the held
'ones now, so the three accessors below are what the modules stand on. Each
'test asks twice and compares the objects: the same object back is what proves
'the store was held rather than rebuilt.

'@sub-title The workbook store is built once and handed back.
'@TestMethod("EventLinelist")
Public Sub TestWorkbookNamesIsHeldAcrossCalls()
    CustomTestSetTitles Assert, TESTMODULE, "TestWorkbookNamesIsHeldAcrossCalls"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstRead As HiddenNames
    Dim secondRead As HiddenNames

    Set sut = EventLinelist.Create(FixtureWkb)
    Set firstRead = sut.WorkbookNames()
    Set secondRead = sut.WorkbookNames()

    Assert.IsNotNothing firstRead, "A workbook gives a store"
    Assert.IsTrue firstRead Is secondRead, "The second read is the same store"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWorkbookNamesIsHeldAcrossCalls", _
                         Err.Number, Err.Description
End Sub

'@sub-title Each worksheet gets its own store, held across calls.
'@TestMethod("EventLinelist")
Public Sub TestSheetNamesGivesOneStorePerSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetNamesGivesOneStorePerSheet"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstSheet As Worksheet
    Dim secondSheet As Worksheet
    Dim firstRead As HiddenNames
    Dim secondRead As HiddenNames
    Dim otherRead As HiddenNames

    Set firstSheet = FixtureWkb.Worksheets(1)
    Set secondSheet = FixtureWkb.Worksheets.Add
    Set sut = EventLinelist.Create(FixtureWkb)

    Set firstRead = sut.SheetNames(firstSheet)
    Set secondRead = sut.SheetNames(firstSheet)
    Set otherRead = sut.SheetNames(secondSheet)

    Assert.IsNotNothing firstRead, "A worksheet gives a store"
    Assert.IsTrue firstRead Is secondRead, "The second read is the same store"
    Assert.IsTrue Not (firstRead Is otherRead), "Another sheet gives another store"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetNamesGivesOneStorePerSheet", _
                         Err.Number, Err.Description
End Sub

'@sub-title A missing worksheet answers Nothing.
'@details
'The button module reads a sheet tag through this and treats an empty tag as
'"the wrong sheet", so the answer has to come back rather than raise.
'@TestMethod("EventLinelist")
Public Sub TestSheetNamesIgnoresNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetNamesIgnoresNothing"
    On Error GoTo TestFail

    Dim sut As EventLinelist

    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.IsNothing sut.SheetNames(Nothing), "Nothing in answers Nothing out"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetNamesIgnoresNothing", _
                         Err.Number, Err.Description
End Sub

'@sub-title A workbook with no password worksheet answers Nothing.
'@TestMethod("EventLinelist")
Public Sub TestPasswordManagerIsNothingWithoutTheSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestPasswordManagerIsNothingWithoutTheSheet"
    On Error GoTo TestFail

    Dim sut As EventLinelist

    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.IsNothing sut.PasswordManager(), _
                     "No __pass worksheet gives no manager"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPasswordManagerIsNothingWithoutTheSheet", _
                         Err.Number, Err.Description
End Sub

'@sub-title A warning with no translation sheet under it shows nothing.
'@details
'Warn and Fail are the one warning box and the one failure box of the linelist
'event surface, and four files used to write their own. A workbook with no
'usable translation sheet has no text to show, so both stay quiet there. This
'test is what holds that line: a box raised here would stop the whole run on a
'modal dialog with nobody to press OK.
'@TestMethod("EventLinelist")
Public Sub TestWarnStaysQuietWithoutATranslationSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestWarnStaysQuietWithoutATranslationSheet"
    On Error GoTo TestFail

    Dim sut As EventLinelist

    Set sut = EventLinelist.Create(FixtureWkb)

    sut.Warn "MSG_NotModify"
    sut.Fail "MSG_ErrUpdate", "some detail"

    Assert.IsNothing sut.Translation(), _
                     "The fixture carries no translation helper, and both calls returned"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWarnStaysQuietWithoutATranslationSheet", _
                         Err.Number, Err.Description
End Sub


'@section The user log funnel
'===============================================================================
'Fail and Warn write their line on the log worksheet before the box shows,
'so the whole button, ribbon and geo surface logs its refusals and failures
'through the two methods it already reports through, and the workbook open
'writes the first line of the session. The fixture carries no translation
'sheet, so the boxes stay quiet and the suite drives the funnel headless.

'@sub-title The log is built once, held, and its sheet is very hidden.
'@TestMethod("EventLinelist")
Public Sub TestUserLogIsHeldAcrossCalls()
    CustomTestSetTitles Assert, TESTMODULE, "TestUserLogIsHeldAcrossCalls"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim firstRead As LLLog
    Dim secondRead As LLLog

    Set sut = EventLinelist.Create(FixtureWkb)
    Set firstRead = sut.UserLog()
    Set secondRead = sut.UserLog()

    Assert.IsNotNothing firstRead, "A workbook gives a log"
    Assert.IsTrue firstRead Is secondRead, "The second read is the same log"
    Assert.AreEqual CLng(xlSheetVeryHidden), _
                    CLng(FixtureWkb.Worksheets(LOG_SHEET).Visible), _
                    "The log sheet grew on the fixture and is very hidden"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestUserLogIsHeldAcrossCalls", _
                         Err.Number, Err.Description
End Sub

'@sub-title Fail writes its failure line even while the box stays quiet.
'@TestMethod("EventLinelist")
Public Sub TestFailWritesTheFailureLine()
    CustomTestSetTitles Assert, TESTMODULE, "TestFailWritesTheFailureLine"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim logsh As Worksheet
    Dim titleRow As Long
    Dim entryRow As Long

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.Fail "MSG_ErrUpdate", "some detail"

    Set logsh = FixtureWkb.Worksheets(LOG_SHEET)

    'The block is headed by the SECTION of the action, and the action itself
    'rides on the entry line ahead of the detail. The log carries three titles
    'for the whole sheet, so the action has to be on the line to say which
    'event a row is.
    'The action is the plain word "failed"; the message and the detail follow
    'it. The message CODE used to fill the action slot, which put a lookup key
    'where every other line carries an action. This fixture has no translation
    'sheet, so the message reads as the code itself, which is the documented
    'fallback of a workbook that cannot translate.
    titleRow = LogRowOfText(logsh, LOG_OUTPUT_COLUMN, LOG_SECTION_LIFECYCLE)
    entryRow = LogRowOfText(logsh, LOG_OUTPUT_COLUMN, "Error")

    Assert.IsTrue (titleRow > 0), "The section of the action heads the block"
    Assert.IsTrue (entryRow > titleRow), "The failure line sits under it"
    Assert.AreEqual "failed: MSG_ErrUpdate: some detail", _
                    CStr(logsh.Cells(entryRow, LOG_DETAIL_COLUMN).Value), _
                    "The action, the message and the detail ride beside the date"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFailWritesTheFailureLine", _
                         Err.Number, Err.Description
End Sub

'@sub-title Warn writes its warning line even while the box stays quiet.
'@TestMethod("EventLinelist")
Public Sub TestWarnWritesTheWarningLine()
    CustomTestSetTitles Assert, TESTMODULE, "TestWarnWritesTheWarningLine"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim logsh As Worksheet
    Dim titleRow As Long
    Dim entryRow As Long

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.Warn "MSG_NotModify"

    Set logsh = FixtureWkb.Worksheets(LOG_SHEET)
    titleRow = LogRowOfText(logsh, LOG_OUTPUT_COLUMN, LOG_SECTION_LIFECYCLE)
    entryRow = LogRowOfText(logsh, LOG_OUTPUT_COLUMN, "Warning")

    Assert.IsTrue (titleRow > 0), "The section of the action heads the block"
    Assert.IsTrue (entryRow > titleRow), "The warning line sits under it"
    Assert.AreEqual "refused: MSG_NotModify", _
                    CStr(logsh.Cells(entryRow, LOG_DETAIL_COLUMN).Value), _
                    "The refusal reads as an action and the message behind it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWarnWritesTheWarningLine", _
                         Err.Number, Err.Description
End Sub

'@sub-title A refusal names the procedure that raised it.
'@details
'The reported log had four different refusals inside ClickShowHideSection all
'writing the same line. VBA carries no call stack, so the caller names itself
'and the name opens the entry.
'@TestMethod("EventLinelist")
Public Sub TestWarnCarriesTheProcedureThatRefused()
    CustomTestSetTitles Assert, TESTMODULE, "TestWarnCarriesTheProcedureThatRefused"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim logsh As Worksheet
    Dim entryRow As Long

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.Warn "MSG_SectionTitleCell", "ClickShowHideSection"

    Set logsh = FixtureWkb.Worksheets(LOG_SHEET)
    entryRow = LogRowOfText(logsh, LOG_OUTPUT_COLUMN, "Warning")

    Assert.AreEqual "ClickShowHideSection > refused: MSG_SectionTitleCell", _
                    CStr(logsh.Cells(entryRow, LOG_DETAIL_COLUMN).Value), _
                    "The procedure, the refusal and the message read in that order"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWarnCarriesTheProcedureThatRefused", _
                         Err.Number, Err.Description
End Sub

'@sub-title A reason meant for the log alone reaches the log and not the box.
'@details
'Three handlers of the buttons module hold a raw Err.Description. It is worth
'having in a log a user sends on and worth little inside a box on a field
'machine, so it goes in logDetail and the box argument is left empty.
'The box stays quiet here for the fixture's own reason -- no translation
'sheet -- so what this test can hold is that the reason reaches the line.
'@TestMethod("EventLinelist")
Public Sub TestFailLogsAReasonKeptOutOfTheBox()
    CustomTestSetTitles Assert, TESTMODULE, "TestFailLogsAReasonKeptOutOfTheBox"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim logsh As Worksheet
    Dim entryRow As Long

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.Fail "MSG_ErrAddRows", fallback:=vbNullString, _
             source:="ClickAddRows", logDetail:="Application-defined error"

    Set logsh = FixtureWkb.Worksheets(LOG_SHEET)
    entryRow = LogRowOfText(logsh, LOG_OUTPUT_COLUMN, "Error")

    Assert.AreEqual _
        "ClickAddRows > failed: MSG_ErrAddRows: Application-defined error", _
        CStr(logsh.Cells(entryRow, LOG_DETAIL_COLUMN).Value), _
        "The log-only reason rides behind the message"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFailLogsAReasonKeptOutOfTheBox", _
                         Err.Number, Err.Description
End Sub

'@sub-title The workbook open writes the first info line of the session.
'@details
'OnWorkbookOpen turns application events back on, so the test re-enters the
'busy state right after the call and the harness keeps the screen.
'@TestMethod("EventLinelist")
Public Sub TestWorkbookOpenWritesTheOpenLine()
    CustomTestSetTitles Assert, TESTMODULE, "TestWorkbookOpenWritesTheOpenLine"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim logsh As Worksheet
    Dim titleRow As Long
    Dim entryRow As Long

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.OnWorkbookOpen
    BusyApp

    Set logsh = FixtureWkb.Worksheets(LOG_SHEET)
    titleRow = LogRowOfText(logsh, LOG_OUTPUT_COLUMN, "open")
    entryRow = LogRowOfText(logsh, LOG_OUTPUT_COLUMN, "Info")

    Assert.IsTrue (titleRow > 0), "The open action heads the block"
    Assert.IsTrue (entryRow > titleRow), "The info line sits under it"
    Assert.AreEqual "open: " & FixtureWkb.Name, _
                    CStr(logsh.Cells(entryRow, LOG_DETAIL_COLUMN).Value), _
                    "The entry line names the action and the workbook that opened"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWorkbookOpenWritesTheOpenLine", _
                         Err.Number, Err.Description
End Sub


'@section The dictionary managers and the variable-labels table
'===============================================================================
'The variable-labels button used to build a dictionary and a variable reader on
'every click, and stage its rows on the __temp worksheet. Both managers are
'held on the service now, and VarLabelTable builds the rows in memory: one row
'per hlist2D variable, carrying the pivot block title of its table, the
'variable name and the main label.

'@sub-title A workbook with no Dictionary sheet answers Nothing, once.
'@TestMethod("EventLinelist")
Public Sub TestDictionaryStopsRetryingAfterAFailedBuild()
    CustomTestSetTitles Assert, TESTMODULE, "TestDictionaryStopsRetryingAfterAFailedBuild"
    On Error GoTo TestFail

    Dim sut As EventLinelist

    Set sut = EventLinelist.Create(FixtureWkb)
    Assert.IsNothing sut.Dictionary(), _
                     "A workbook with no Dictionary sheet answers Nothing"

    'The sheet arrives after the failed build. The tried flag holds the answer.
    DictionaryTestFixture.PrepareDictionaryFixture "Dictionary", FixtureWkb
    Assert.IsNothing sut.Dictionary(), _
                     "The session keeps the answer of the first build"

    Set sut = EventLinelist.Create(FixtureWkb)
    Assert.IsNotNothing sut.Dictionary(), "A fresh service sees the sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDictionaryStopsRetryingAfterAFailedBuild", _
                         Err.Number, Err.Description
End Sub

'@sub-title The dictionary and the variable reader are built once and handed back.
'@TestMethod("EventLinelist")
Public Sub TestDictionaryAndVariablesAreHeldAcrossCalls()
    CustomTestSetTitles Assert, TESTMODULE, "TestDictionaryAndVariablesAreHeldAcrossCalls"
    On Error GoTo TestFail

    Dim sut As EventLinelist

    DictionaryTestFixture.PrepareDictionaryFixture "Dictionary", FixtureWkb
    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.IsNotNothing sut.Dictionary(), "A workbook with the sheet gives a dictionary"
    Assert.IsTrue sut.Dictionary() Is sut.Dictionary(), _
                  "The second read is the same dictionary"
    Assert.IsNotNothing sut.Variables(), "And a variable reader stands on it"
    Assert.IsTrue sut.Variables() Is sut.Variables(), _
                  "The second read is the same reader"
    Assert.IsTrue sut.Variables().Dictionary Is sut.Dictionary(), _
                  "The reader reads the held dictionary"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDictionaryAndVariablesAreHeldAcrossCalls", _
                         Err.Number, Err.Description
End Sub

'@sub-title A workbook with no dictionary gives an empty table.
'@TestMethod("EventLinelist")
Public Sub TestVarLabelTableIsEmptyWithoutADictionary()
    CustomTestSetTitles Assert, TESTMODULE, "TestVarLabelTableIsEmptyWithoutADictionary"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim labelRows As BetterArray

    Set sut = EventLinelist.Create(FixtureWkb)
    Set labelRows = sut.VarLabelTable()

    Assert.IsNotNothing labelRows, "The table itself always comes back"
    Assert.AreEqual CLng(0), labelRows.Length, _
                    "A workbook with no dictionary has no rows to offer"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestVarLabelTableIsEmptyWithoutADictionary", _
                         Err.Number, Err.Description
End Sub

'@sub-title One row per hlist2D variable: pivot title, name and label.
'@TestMethod("EventLinelist")
Public Sub TestVarLabelTableListsTheHListVariables()
    CustomTestSetTitles Assert, TESTMODULE, "TestVarLabelTableListsTheHListVariables"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim labelRows As BetterArray
    Dim pivotSh As Worksheet
    Dim pivotStore As HiddenNames
    Dim tabName As String
    Dim rowData As Variant

    DictionaryTestFixture.PrepareDictionaryFixture "Dictionary", FixtureWkb
    Set pivotSh = FixtureWkb.Worksheets.Add
    Set sut = EventLinelist.Create(FixtureWkb)

    '`table name` is derived by Prepare, so the fixture writes it nowhere and a
    'test that wrote it down would be asserting a guess. Read it off the reader.
    sut.Dictionary().Prepare
    tabName = sut.Variables().Value(colName:="table name", varName:="mand_h2")

    'The custom pivot sheet is reached through RNG_CustomPivot, and each table's
    'title is a hidden name CustomPivotTable writes on that sheet.
    SetWorkbookName FixtureWkb, "RNG_CustomPivot", pivotSh.Name
    Set pivotStore = HiddenNames.Create(pivotSh)
    pivotStore.EnsureName "pivot_title_" & tabName, "Case listing", HiddenNameTypeString

    Set labelRows = sut.VarLabelTable()

    Assert.AreEqual DictionaryTestFixture.DictionaryFieldEquals("Sheet Type", "hlist2D").Length, _
                    labelRows.Length, _
                    "The table holds one row per hlist2D variable of the dictionary"

    rowData = VarLabelRowOf(labelRows, "mand_h2")
    Assert.IsFalse IsEmpty(rowData), "mand_h2 is on a data entry sheet, so it has a row"
    Assert.AreEqual "Case listing", CStr(rowData(LBound(rowData))), _
                    "The row opens with the pivot title of its table"
    Assert.AreEqual "Mandatory variable on hlist2D", _
                    CStr(rowData(LBound(rowData) + 2)), _
                    "And closes with the main label"

    rowData = VarLabelRowOf(labelRows, "mand_v1")
    Assert.IsTrue IsEmpty(rowData), "A vlist1D variable has no row here"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestVarLabelTableListsTheHListVariables", _
                         Err.Number, Err.Description
End Sub


'@section The filter tables
'===============================================================================
'UpdateFilterTables rewrites each filtered table from the visible rows of its
'source table that hold data. The fixture is the pair SeedFilteredPair builds:
'a two-column HList table and a header-only companion at the same address.

'@sub-title The filtered table holds the visible rows that carry data.
'@details
'One source row is hidden and two are blank, which is what a table padded by
'ClickAddRows looks like. The hidden row and the blank rows stay out, and the
'kept rows keep their order.
'@TestMethod("EventLinelist")
Public Sub TestFilterTablesKeepVisibleFilledRows()
    CustomTestSetTitles Assert, TESTMODULE, "TestFilterTablesKeepVisibleFilledRows"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim filtsh As Worksheet
    Dim bodyRng As Range

    Set filtsh = SeedFilteredPair(Array( _
        Array("a", 1), _
        Array("b", 2), _
        Array(vbNullString, vbNullString), _
        Array("d", 4), _
        Array(vbNullString, vbNullString)))
    FixtureWkb.Worksheets(1).Rows(10).Hidden = True

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.UpdateFilterTables calculate:=False

    Set bodyRng = filtsh.ListObjects(1).DataBodyRange
    Assert.AreEqual 2&, bodyRng.Rows.Count, _
                    "The visible rows holding data are the rows kept"
    Assert.AreEqual "a", CStr(bodyRng.Cells(1, 1).Value), _
                    "The first kept row is the first visible row with data"
    Assert.AreEqual "d", CStr(bodyRng.Cells(2, 1).Value), _
                    "The hidden row and the blank rows stay out"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFilterTablesKeepVisibleFilledRows", _
                         Err.Number, Err.Description
End Sub

'@sub-title A hidden column keeps its values in the filtered table.
'@details
'Show/hide hides variable columns on a data entry sheet. The filtered table
'feeds the analysis formulas, so a hidden variable still has to travel.
'@TestMethod("EventLinelist")
Public Sub TestFilterTablesCarryHiddenColumns()
    CustomTestSetTitles Assert, TESTMODULE, "TestFilterTablesCarryHiddenColumns"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim filtsh As Worksheet
    Dim bodyRng As Range

    Set filtsh = SeedFilteredPair(Array(Array("a", 1), Array("b", 2)))
    FixtureWkb.Worksheets(1).Columns(3).Hidden = True

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.UpdateFilterTables calculate:=False

    Set bodyRng = filtsh.ListObjects(1).DataBodyRange
    Assert.AreEqual 2&, bodyRng.Rows.Count, _
                    "Both rows travel while a column is hidden"
    Assert.AreEqual "1", CStr(bodyRng.Cells(1, 2).Value), _
                    "The hidden column keeps its values"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFilterTablesCarryHiddenColumns", _
                         Err.Number, Err.Description
End Sub

'@sub-title Stale rows leave when every source row is hidden.
'@details
'The filtered table starts with rows from an earlier pass. Hiding every source
'row is what a filter matching nothing does, and the rewrite has to leave the
'table header-only in that state.
'@TestMethod("EventLinelist")
Public Sub TestFilterTablesClearStaleRowsWhenEverythingIsHidden()
    CustomTestSetTitles Assert, TESTMODULE, _
                        "TestFilterTablesClearStaleRowsWhenEverythingIsHidden"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim filtsh As Worksheet

    Set filtsh = SeedFilteredPair(Array(Array("a", 1), Array("b", 2)))
    filtsh.ListObjects(1).Resize filtsh.Range("B8:C10")
    filtsh.Range("B9").Value = "old1"
    filtsh.Range("B10").Value = "old2"
    FixtureWkb.Worksheets(1).Rows("9:10").Hidden = True

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.UpdateFilterTables calculate:=False

    Assert.IsNothing filtsh.ListObjects(1).DataBodyRange, _
                     "The stale rows are gone and the table is header-only"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, _
                         "TestFilterTablesClearStaleRowsWhenEverythingIsHidden", _
                         Err.Number, Err.Description
End Sub

'@sub-title A sheet the refresh cannot place is skipped, named and logged.
'@details
'The refresh used to stop at the first broken sheet and say nothing about
'which one it was. It skips that sheet now, finishes the healthy pair, and the
'failure line of the run names the sheet in the user log. The box stays quiet
'because the fixture carries no translation sheet, which is what lets this run
'headless.
'@TestMethod("EventLinelist")
Public Sub TestFilterTablesSkipAndNameABrokenSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestFilterTablesSkipAndNameABrokenSheet"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim filtsh As Worksheet
    Dim logsh As Worksheet
    Dim bodyRng As Range

    Set filtsh = SeedFilteredPair(Array(Array("a", 1), Array("b", 2)))
    SeedBrokenHListSheet "hlist_broken"

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.UpdateFilterTables calculate:=False

    Set bodyRng = filtsh.ListObjects(1).DataBodyRange
    Assert.AreEqual 2&, bodyRng.Rows.Count, _
                    "The healthy pair still syncs past a broken sheet"

    Set logsh = FixtureWkb.Worksheets(LOG_SHEET)
    Assert.IsTrue (LogRowOfText(logsh, LOG_OUTPUT_COLUMN, _
                                LOG_SECTION_LIFECYCLE) > 0), _
                  "The refresh failure sits under its section in the user log"
    Assert.IsTrue (LogRowOfText(logsh, LOG_DETAIL_COLUMN, "MSG_ErrUpdate") > 0), _
                  "The message code rides on the entry line"
    Assert.IsTrue (LogRowOfText(logsh, LOG_DETAIL_COLUMN, "hlist_broken") > 0), _
                  "The failure line names the sheet that was skipped"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, _
                         "TestFilterTablesSkipAndNameABrokenSheet", _
                         Err.Number, Err.Description
End Sub


'@section The quiet state of the events manager
'===============================================================================
'LinelistEventsManager owns Application.EnableEvents for a running linelist.
'The quiet state is the events-only half of it, for work that needs silence and
'nothing else: the import resizing every data entry table, the geo form writing
'a place into the cells beside it. Both used to write the flag themselves, and a
'run that ended between the two lines left worksheet events dead for the whole
'session -- no dropdown cascade, no checking, no autofill.
'
'The four tests below state the whole contract. What none of them can reach is a
'VBA state reset, which empties the counter and leaves Excel as the work left
'it; there is no way to make VBA drop its state from inside a test.

'@sub-title The quiet state takes the events and gives them back.
'@TestMethod("EventLinelist")
Public Sub TestQuietStateTakesTheEventsAndGivesThemBack()
    CustomTestSetTitles Assert, TESTMODULE, "TestQuietStateTakesTheEventsAndGivesThemBack"
    On Error GoTo TestFail

    LinelistEventsManager.LLEnterQuietState
    Assert.IsFalse Application.EnableEvents, _
                   "A quiet stretch should silence the worksheet events"

    LinelistEventsManager.LLExitQuietState
    Assert.IsTrue Application.EnableEvents, _
                  "The end of the stretch should give the events back"

    Exit Sub
TestFail:
    Application.EnableEvents = True
    CustomTestLogFailure Assert, _
                         "TestQuietStateTakesTheEventsAndGivesThemBack", _
                         Err.Number, Err.Description
End Sub

'@sub-title Only the outermost exit gives the events back.
'@details
'The import opens a stretch and the walk inside it opens another. An inner exit
'that handed the events back would raise the sheet-change handler on every row
'the outer work still had to write.
'@TestMethod("EventLinelist")
Public Sub TestQuietStateGivesTheEventsBackOnTheOutermostExitOnly()
    CustomTestSetTitles Assert, TESTMODULE, "TestQuietStateGivesTheEventsBackOnTheOutermostExitOnly"
    On Error GoTo TestFail

    LinelistEventsManager.LLEnterQuietState
    LinelistEventsManager.LLEnterQuietState

    LinelistEventsManager.LLExitQuietState
    Assert.IsFalse Application.EnableEvents, _
                   "An inner exit should leave the events where the outer work put them"

    LinelistEventsManager.LLExitQuietState
    Assert.IsTrue Application.EnableEvents, _
                  "The outermost exit is what gives the events back"

    Exit Sub
TestFail:
    Application.EnableEvents = True
    CustomTestLogFailure Assert, _
                         "TestQuietStateGivesTheEventsBackOnTheOutermostExitOnly", _
                         Err.Number, Err.Description
End Sub

'@sub-title An exit with nothing open changes nothing.
'@details
'FormLogicGeo calls the exit from its error label, which is reached whether or
'not the raise happened inside a stretch. An exit that wrote the flag blind
'would turn events on in the middle of somebody else's work.
'@TestMethod("EventLinelist")
Public Sub TestQuietStateExitWithNothingOpenChangesNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestQuietStateExitWithNothingOpenChangesNothing"
    On Error GoTo TestFail

    Application.EnableEvents = False
    LinelistEventsManager.LLExitQuietState

    Assert.IsFalse Application.EnableEvents, _
                   "An exit with no stretch open should write nothing at all"

    Application.EnableEvents = True

    Exit Sub
TestFail:
    Application.EnableEvents = True
    CustomTestLogFailure Assert, _
                         "TestQuietStateExitWithNothingOpenChangesNothing", _
                         Err.Number, Err.Description
End Sub

'@sub-title The quiet state puts back what it found, not what it assumed.
'@details
'This is what makes it safe under the busy state. A stretch that opens while the
'events are already off must hand that back and let whoever turned them off
'decide when they return. A stretch that assumed True would hand the events to
'the user in the middle of a generation.
'@TestMethod("EventLinelist")
Public Sub TestQuietStatePutsBackWhatItFound()
    CustomTestSetTitles Assert, TESTMODULE, "TestQuietStatePutsBackWhatItFound"
    On Error GoTo TestFail

    Application.EnableEvents = False

    LinelistEventsManager.LLEnterQuietState
    LinelistEventsManager.LLExitQuietState

    Assert.IsFalse Application.EnableEvents, _
                   "The stretch should hand back the events it was given"

    Application.EnableEvents = True

    Exit Sub
TestFail:
    Application.EnableEvents = True
    CustomTestLogFailure Assert, _
                         "TestQuietStatePutsBackWhatItFound", _
                         Err.Number, Err.Description
End Sub

'@sub-title The quiet state takes the events and nothing else.
'@details
'This is what the show/hide session runs under. The form is modal and its own
'writes raise no worksheet event, so all it needs is silence -- and the busy
'state would give it far more than that. The busy state turns the screen off and
'sets the pointer to the arrow, and giving both back at the end of every store
'write and every log line is what made the form flicker on open, on close and on
'the step into the sections form.
'
'The pointer and the screen are read on both sides of the stretch. A quiet state
'that touched either would put that flicker straight back.
'@TestMethod("EventLinelist")
Public Sub TestQuietStateLeavesThePointerAndTheScreenAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestQuietStateLeavesThePointerAndTheScreenAlone"
    On Error GoTo TestFail

    Dim pointerBefore As XlMousePointer

    pointerBefore = Application.Cursor
    Application.Cursor = xlWait
    Application.ScreenUpdating = True

    LinelistEventsManager.LLEnterQuietState
    Assert.AreEqual CLng(xlWait), CLng(Application.Cursor), _
                     "A quiet stretch has no business with the pointer"
    Assert.IsTrue Application.ScreenUpdating, _
                  "A quiet stretch has no business with the screen"

    LinelistEventsManager.LLExitQuietState
    Assert.AreEqual CLng(xlWait), CLng(Application.Cursor), _
                     "And it leaves the pointer where it found it on the way out"
    Assert.IsTrue Application.ScreenUpdating, _
                  "And the screen with it"

    Application.Cursor = pointerBefore

    Exit Sub
TestFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    CustomTestLogFailure Assert, _
                         "TestQuietStateLeavesThePointerAndTheScreenAlone", _
                         Err.Number, Err.Description
End Sub


'@section The resting pointer
'===============================================================================
'OnWorkbookOpen parks the pointer of a generated linelist on the north-west
'arrow, and every busy state of the session shows that same arrow. The two being
'equal is what makes an event leave no visible change: ApplicationState
'snapshots the standing pointer and puts it back, so arrow follows arrow.
'
'A modal form breaks it. Excel hands the pointer back on the default cursor once
'the form closes, so the standing pointer is no longer the arrow, and from there
'every selection on a data entry sheet flicks it twice -- to the arrow going in
'and to the default coming out. The form is long gone by then. LLRestPointer is
'called on the way out of each form session to put the invariant back.

'@sub-title The rest call parks the pointer on the arrow.
'@TestMethod("EventLinelist")
Public Sub TestRestPointerParksThePointerOnTheArrow()
    CustomTestSetTitles Assert, TESTMODULE, "TestRestPointerParksThePointerOnTheArrow"
    On Error GoTo TestFail

    Dim pointerBefore As XlMousePointer

    pointerBefore = Application.Cursor

    'What a closed modal form leaves behind.
    Application.Cursor = xlDefault

    LinelistEventsManager.LLRestPointer
    Assert.AreEqual CLng(xlNorthwestArrow), CLng(Application.Cursor), _
                     "The rest call should put the pointer back on the arrow"

    Application.Cursor = pointerBefore

    Exit Sub
TestFail:
    Application.Cursor = xlDefault
    CustomTestLogFailure Assert, _
                         "TestRestPointerParksThePointerOnTheArrow", _
                         Err.Number, Err.Description
End Sub

'@sub-title Calling it on a pointer already at rest changes nothing.
'@details
'The three show/hide handlers call it from a label every path reaches, the path
'that opened no form included. It has to be safe to call when there is nothing
'to put right.
'@TestMethod("EventLinelist")
Public Sub TestRestPointerIsSafeWhenThePointerIsAlreadyAtRest()
    CustomTestSetTitles Assert, TESTMODULE, "TestRestPointerIsSafeWhenThePointerIsAlreadyAtRest"
    On Error GoTo TestFail

    Dim pointerBefore As XlMousePointer

    pointerBefore = Application.Cursor

    Application.Cursor = xlNorthwestArrow
    LinelistEventsManager.LLRestPointer
    LinelistEventsManager.LLRestPointer

    Assert.AreEqual CLng(xlNorthwestArrow), CLng(Application.Cursor), _
                     "A pointer already at rest should stay where it is"

    Application.Cursor = pointerBefore

    Exit Sub
TestFail:
    Application.Cursor = xlDefault
    CustomTestLogFailure Assert, _
                         "TestRestPointerIsSafeWhenThePointerIsAlreadyAtRest", _
                         Err.Number, Err.Description
End Sub


'@section The resting state
'===============================================================================
'A linelist runs on manual calculation and rests its pointer on the north-west
'arrow. Both are what every busy state of the session applies, and
'ApplicationState puts back what it snapshots, so with the resting values equal
'to the busy ones an event leaves neither of them changed.
'
'The open used to set them inside a busy state whose snapshot was taken before
'it, so the exit restored the host Excel's own calculation mode over the top. A
'linelist opened into a session that already had a workbook up therefore ran on
'automatic calculation, and every event of the workbook ended in a full
'recalculation on the way out of its busy state. LinelistEventsManager calls
'this once more after the open closes, which is why the sub exists apart from
'OnWorkbookOpen and why it is worth a test of its own.

'@sub-title The resting state is manual calculation and the arrow.
'@TestMethod("EventLinelist")
Public Sub TestTheRestingStateIsManualCalculationAndTheArrow()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheRestingStateIsManualCalculationAndTheArrow"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim heldCalculation As Long

    heldCalculation = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    Application.Cursor = xlDefault

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.ApplyRestingState

    Assert.AreEqual CLng(xlCalculationManual), CLng(Application.Calculation), _
                    "The resting state is manual calculation"
    Assert.AreEqual CLng(xlNorthwestArrow), CLng(Application.Cursor), _
                    "The resting state parks the pointer on the north-west arrow"

    RestoreAfterOpen heldCalculation
    Exit Sub
TestFail:
    RestoreAfterOpen heldCalculation
    CustomTestLogFailure Assert, "TestTheRestingStateIsManualCalculationAndTheArrow", _
                         Err.Number, Err.Description
End Sub


'@section What a handler is worth taking application state for
'===============================================================================
'The two questions LinelistEventsManager asks before it picks the state to run a
'handler under. Screen updating off and back on repaints the window whether or
'not anything changed, so a handler that turns out to have no work costs the
'user a flicker for nothing. Both answers come off values the handlers were
'reading anyway.

'@sub-title A sheet whose flag says yes has a rebuild waiting on it.
'@TestMethod("EventLinelist")
Public Sub TestSheetHasListAutoWorkFollowsTheFlag()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetHasListAutoWorkFollowsTheFlag"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim dataWksh As Worksheet

    Set dataWksh = FixtureWkb.Worksheets(1)
    SetSheetName dataWksh, LISTAUTO_FLAG, "yes"
    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.IsTrue sut.SheetHasListAutoWork(dataWksh), _
                  "A sheet flagged yes has a rebuild waiting on it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetHasListAutoWorkFollowsTheFlag", _
                         Err.Number, Err.Description
End Sub

'@sub-title A sheet nobody edited, and Nothing, are both worth no state.
'@details
'This is the answer that matters. It is what the manager reads on every sheet
'the user leaves, and reading False is what keeps the busy state, and the
'repaint that comes with it, off the ordinary path.
'@TestMethod("EventLinelist")
Public Sub TestSheetHasListAutoWorkIsFalseWithoutAnEdit()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetHasListAutoWorkIsFalseWithoutAnEdit"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim dataWksh As Worksheet

    Set dataWksh = FixtureWkb.Worksheets(1)
    SetSheetName dataWksh, LISTAUTO_FLAG, "no"
    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.IsFalse sut.SheetHasListAutoWork(dataWksh), _
                   "A sheet flagged no has nothing waiting on it"
    Assert.IsFalse sut.SheetHasListAutoWork(Nothing), _
                   "A Nothing sheet has nothing waiting on it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetHasListAutoWorkIsFalseWithoutAnEdit", _
                         Err.Number, Err.Description
End Sub

'@sub-title A cell under a geo column starts the admin cascade.
'@TestMethod("EventLinelist")
Public Sub TestSelectionIsHeavyUnderAGeoColumn()
    CustomTestSetTitles Assert, TESTMODULE, "TestSelectionIsHeavyUnderAGeoColumn"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim sh As Worksheet

    Set sh = SeedHListSheet("adm2_var", "geo2")
    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.IsTrue sut.SelectionIsHeavy(sh, sh.Cells(9, 2)), _
                  "A cell under a geo2 column starts the admin cascade"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSelectionIsHeavyUnderAGeoColumn", _
                         Err.Number, Err.Description
End Sub

'@sub-title An ordinary cell, the header row and Nothing all start nothing.
'@details
'The three answers the arrow keys give. A selection that starts no cascade runs
'under the quiet state, which writes nothing the user can see, and that is what
'stops a data entry sheet flickering once per key.
'@TestMethod("EventLinelist")
Public Sub TestSelectionIsNotHeavyAnywhereElse()
    CustomTestSetTitles Assert, TESTMODULE, "TestSelectionIsNotHeavyAnywhereElse"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim sh As Worksheet

    Set sh = SeedHListSheet("plain_var", "choice_multiple")
    Set sut = EventLinelist.Create(FixtureWkb)

    Assert.IsFalse sut.SelectionIsHeavy(sh, sh.Cells(9, 2)), _
                   "A cell under an ordinary column starts no cascade"
    Assert.IsFalse sut.SelectionIsHeavy(sh, sh.Cells(8, 2)), _
                   "The header row itself starts no cascade"
    Assert.IsFalse sut.SelectionIsHeavy(sh, Nothing), _
                   "A Nothing selection starts no cascade"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSelectionIsNotHeavyAnywhereElse", _
                         Err.Number, Err.Description
End Sub


'@section Where the open lands
'===============================================================================
'A linelist is saved on the sheet the generation stopped on, which is Geo, so
'the open has to move the user off it. The instruction sheet is where they
'belong, and it is not always there to move them to: a build with the
'instructions turned off carries it very hidden.
'
'That case used to be turned away. The fallback to the first visible worksheet
'ran only when no instruction sheet was found at all, and a hidden one is found,
'so the visibility test below it exited and the user was left on Geo. This is
'that case.

'@sub-title A hidden instruction sheet lands the user on the first visible sheet.
'@TestMethod("EventLinelist")
Public Sub TestTheOpenLandsOnTheFirstVisibleSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheOpenLandsOnTheFirstVisibleSheet"
    On Error GoTo TestFail

    Dim sut As EventLinelist
    Dim hiddenSheet As Worksheet
    Dim landingSheet As Worksheet
    Dim heldCalculation As Long

    heldCalculation = Application.Calculation

    'The workbook the fixture hands over holds one sheet, so the sheet to land
    'on is added before the first one is taken out of sight. A workbook cannot
    'hide its last visible sheet.
    Set hiddenSheet = FixtureWkb.Worksheets(1)
    Set landingSheet = FixtureWkb.Worksheets.Add(After:=hiddenSheet)
    landingSheet.Name = "landing"
    hiddenSheet.Visible = xlSheetVeryHidden

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.OnWorkbookOpen

    Assert.AreEqual "landing", FixtureWkb.ActiveSheet.Name, _
                    "The open lands on the first sheet the user can see"

    RestoreAfterOpen heldCalculation
    Exit Sub
TestFail:
    RestoreAfterOpen heldCalculation
    CustomTestLogFailure Assert, "TestTheOpenLandsOnTheFirstVisibleSheet", _
                         Err.Number, Err.Description
End Sub
