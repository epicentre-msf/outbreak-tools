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
'@depends EventLinelist, CustomTest, HiddenNames, LLTranslation, TranslationObject

Private Assert As CustomTest
Private FixtureWkb As Workbook

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "EventLinelist"
Private Const TRANS_SHEET_NAME As String = "LinelistTranslation"
Private Const LISTAUTO_FLAG As String = "RNG_UpdateListAuto"

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
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    RestoreApp
    Set Assert = Nothing
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
    Dim dataSheet As Worksheet

    Set dataSheet = FixtureWkb.Worksheets(1)
    SetWorkbookName FixtureWkb, LISTAUTO_FLAG, "yes"

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.OnSheetDeactivate dataSheet

    Assert.AreEqual "no", WorkbookNameValue(FixtureWkb, LISTAUTO_FLAG), _
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
    Dim dataSheet As Worksheet

    Set dataSheet = FixtureWkb.Worksheets(1)
    SetWorkbookName FixtureWkb, LISTAUTO_FLAG, "no"

    Set sut = EventLinelist.Create(FixtureWkb)
    sut.OnSheetDeactivate dataSheet

    Assert.AreEqual "no", WorkbookNameValue(FixtureWkb, LISTAUTO_FLAG), _
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
    Dim errNumber As Long

    SetWorkbookName FixtureWkb, LISTAUTO_FLAG, "yes"
    Set sut = EventLinelist.Create(FixtureWkb)

    On Error Resume Next
        sut.OnSheetDeactivate Nothing
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual 0&, errNumber, "A Nothing sheet raises nothing"
    Assert.AreEqual "yes", WorkbookNameValue(FixtureWkb, LISTAUTO_FLAG), _
                    "A Nothing sheet leaves the flag alone"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDeactivateIgnoresNothing", Err.Number, Err.Description
End Sub


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
