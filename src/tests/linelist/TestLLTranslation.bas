Attribute VB_Name = "TestLLTranslation"
Attribute VB_Description = "Tests for LLTranslation class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLTranslation class")

'@description
'Drives LLTranslation, which owns the translation tables of a linelist: the
'messages, the shapes, the form labels, the ribbon labels, and the dictionary
'table that carries the variable labels. Every linelist form module builds one,
'and so do the build service, the exporter and LinelistSpecs.
'
'The class has never been audited. This suite pins the behaviour that is worth
'keeping while the audit is still owed, and the open questions it raised are in
'.obt/gotchas/linelist-folder.md.
'
'The suite this replaced qualified ten helper calls with TestHelpers, which the
'run does not import, and its module hooks were Private. Two of its tests
'asserted only that an error number differed from zero, and one asserted True.
'@depends LLTranslation, CustomTest, TranslationObject, HiddenNames

Private Assert As CustomTest
Private FixtureWkb As Workbook
Private TransSheet As Worksheet

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "LLTranslation"
Private Const TRANS_SHEET_NAME As String = "LinelistTranslation"

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
    Assert.SetModuleName "TestLLTranslation"
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

'@sub-title Build a fresh translation workbook before each test.
'@details
'There is no BeginTest call here on purpose. BeginTest opens the checking with
'whatever titles are pending at that moment, and the Flush in TestCleanup has
'just reset those to the default, so every result of this module was filed
'under the default label and the per-module count could not be read. Letting
'the first assertion of each test open the checking picks up the titles that
'CustomTestSetTitles set at the top of the test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureWkb = NewWorkbook()
    SeedTranslationSheet FixtureWkb
    Set TransSheet = FixtureWkb.Worksheets(TRANS_SHEET_NAME)
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

    Set TransSheet = Nothing
    Set FixtureWkb = Nothing
End Sub

'@section Test Fixture Helpers
'===============================================================================

'@sub-title Build a translation worksheet with the required tables and languages
'@details
'The four required tables are seeded plus T_TradLLRibbon, which Create treats
'as optional and TransObject requires.
'@param targetWkb Workbook. The workbook to seed.
Private Sub SeedTranslationSheet(ByVal targetWkb As Workbook)
    Dim sh As Worksheet
    Dim wkbNames As HiddenNames

    Set sh = targetWkb.Worksheets.Add
    sh.Name = TRANS_SHEET_NAME

    SeedTable sh, 1, "T_TradLLMsg", Array( _
        Array("MSG_GoToSection", "Go to section"), _
        Array("MSG_AnaPeriod", "Analysis period"), _
        Array("MSG_GoToHead", "Go to header"), _
        Array("MSG_NoDevide", "Do not split"), _
        Array("MSG_Devide", "Split"), _
        Array("MSG_GoToGraph", "Go to graph"), _
        Array("MSG_ComputeOnFiltered", "Compute on filtered"), _
        Array("LLSHEET_CustomChoice", "Custom dropdown"), _
        Array("LLSHEET_Analysis", "Analysis"), _
        Array("LLSHEET_TemporalAnalysis", "Temporal"), _
        Array("LLSHEET_SpatialAnalysis", "Spatial"), _
        Array("LLSHEET_SpatioTemporalAnalysis", "SpatioTemporal"), _
        Array("LLSHEET_CustomPivotTable", "Custom pivot"), _
        Array("MSG_W", "W"), _
        Array("MSG_Q", "Q"), _
        Array("MSG_InfoStart", "Info start"), _
        Array("MSG_InfoEnd", "Info end"))

    SeedTable sh, 4, "T_TradLLShapes", Array(Array("SHP_Advanced", "Advanced"))
    SeedTable sh, 7, "T_TradLLForms", Array(Array("FRM_Title", "Form title"))
    SeedTable sh, 10, "Tab_Translations", Array(Array("DICT_Var1", "Variable 1"))
    SeedTable sh, 13, "T_TradLLRibbon", Array(Array("RIB_Advanced", "Ribbon advanced"))

    'TransObject reads both language codes from the WORKBOOK store.
    Set wkbNames = HiddenNames.Create(targetWkb)
    wkbNames.EnsureName "RNG_LLLanguageCode", "en", HiddenNameTypeString
    wkbNames.EnsureName "RNG_DictionaryLanguage", "en", HiddenNameTypeString
End Sub

'@sub-title Write one translation table and name it.
'@param sh Worksheet. The host worksheet.
'@param startColumn Long. The column the label column sits in.
'@param tableName String. The name to give the ListObject.
'@param entries Variant. An array of label and value pairs.
Private Sub SeedTable(ByVal sh As Worksheet, ByVal startColumn As Long, _
                      ByVal tableName As String, ByVal entries As Variant)
    Dim idx As Long
    Dim lastRow As Long
    Dim rng As Range

    sh.Cells(1, startColumn).Value = "label"
    sh.Cells(1, startColumn + 1).Value = "en"

    For idx = LBound(entries) To UBound(entries)
        sh.Cells(idx + 2, startColumn).Value = entries(idx)(0)
        sh.Cells(idx + 2, startColumn + 1).Value = entries(idx)(1)
    Next idx

    lastRow = UBound(entries) - LBound(entries) + 2
    Set rng = sh.Range(sh.Cells(1, startColumn), sh.Cells(lastRow, startColumn + 1))
    sh.ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, _
                       XlListObjectHasHeaders:=xlYes).Name = tableName
End Sub

'@section Factory Tests
'===============================================================================

'@sub-title A seeded translation worksheet gives an instance bound to it.
'@TestMethod("LLTranslation")
Public Sub TestCreateReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsInstance"
    On Error GoTo TestFail

    Dim sut As LLTranslation

    Set sut = LLTranslation.Create(TransSheet)

    Assert.IsTrue (Not sut Is Nothing), "Create gives back an instance"
    Assert.AreEqual TRANS_SHEET_NAME, sut.Wksh().Name, _
                    "The instance is bound to the worksheet it was given"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsInstance", Err.Number, Err.Description
End Sub

'@sub-title A Nothing worksheet is refused, and the number says why.
'@TestMethod("LLTranslation")
Public Sub TestCreateRejectsNothingSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingSheet"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim errNumber As Long

    On Error Resume Next
        Set sut = LLTranslation.Create(Nothing)
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing worksheet is refused and the number names the reason"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingSheet", Err.Number, Err.Description
End Sub

'@sub-title A worksheet with no translation tables is refused.
'@TestMethod("LLTranslation")
Public Sub TestCreateRejectsMissingTables()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsMissingTables"
    On Error GoTo TestFail

    Dim emptyWkb As Workbook
    Dim sut As LLTranslation
    Dim errNumber As Long

    Set emptyWkb = NewWorkbook()

    On Error Resume Next
        Set sut = LLTranslation.Create(emptyWkb.Worksheets(1))
        errNumber = Err.Number
    On Error GoTo 0

    DeleteWorkbook emptyWkb

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "A worksheet with no translation table is refused"

    Exit Sub
TestFail:
    On Error Resume Next
        If Not emptyWkb Is Nothing Then DeleteWorkbook emptyWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestCreateRejectsMissingTables", Err.Number, Err.Description
End Sub

'@sub-title The host worksheet cannot be swapped after creation.
'@details
'The instance caches two hidden name stores taken from that worksheet and its
'workbook, so a swap would leave the stores pointing at a worksheet the
'instance had given up.
'@TestMethod("LLTranslation")
Public Sub TestInternalSheetIsSetAtCreationOnly()
    CustomTestSetTitles Assert, TESTMODULE, "TestInternalSheetIsSetAtCreationOnly"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim errNumber As Long

    Set sut = LLTranslation.Create(TransSheet)

    On Error Resume Next
        Set sut.InternalSheet = FixtureWkb.Worksheets(1)
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.SomethingWentWrong), errNumber, _
                    "The worksheet is set at creation only"
    Assert.AreEqual TRANS_SHEET_NAME, sut.Wksh().Name, _
                    "The refused write left the original worksheet in place"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestInternalSheetIsSetAtCreationOnly", _
                         Err.Number, Err.Description
End Sub

'@section Scope routing
'===============================================================================

'@sub-title Each scope reads the table that belongs to it.
'@details
'A scope reading the wrong table gives a translation object that answers for
'every label it is asked about and gets each one wrong, so a routing test is
'worth more than an emptiness check. Each fixture table carries one
'label nothing else carries.
'@TestMethod("LLTranslation")
Public Sub TestEachScopeReadsItsOwnTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestEachScopeReadsItsOwnTable"
    On Error GoTo TestFail

    Dim sut As LLTranslation

    Set sut = LLTranslation.Create(TransSheet)

    Assert.AreEqual "Go to section", _
                    sut.TransObject(TranslationOfMessages).TranslatedValue("MSG_GoToSection"), _
                    "The messages scope reads T_TradLLMsg"
    Assert.AreEqual "Advanced", _
                    sut.TransObject(TranslationOfShapes).TranslatedValue("SHP_Advanced"), _
                    "The shapes scope reads T_TradLLShapes"
    Assert.AreEqual "Form title", _
                    sut.TransObject(TranslationOfForms).TranslatedValue("FRM_Title"), _
                    "The forms scope reads T_TradLLForms"
    Assert.AreEqual "Variable 1", _
                    sut.TransObject(TranslationOfDictionary).TranslatedValue("DICT_Var1"), _
                    "The dictionary scope reads Tab_Translations"
    Assert.AreEqual "Ribbon advanced", _
                    sut.TransObject(TranslationOfRibbon).TranslatedValue("RIB_Advanced"), _
                    "The ribbon scope reads T_TradLLRibbon"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEachScopeReadsItsOwnTable", Err.Number, Err.Description
End Sub

'@sub-title The default scope is the messages scope.
'@TestMethod("LLTranslation")
Public Sub TestTransObjectDefaultsToMessages()
    CustomTestSetTitles Assert, TESTMODULE, "TestTransObjectDefaultsToMessages"
    On Error GoTo TestFail

    Dim sut As LLTranslation

    Set sut = LLTranslation.Create(TransSheet)

    Assert.AreEqual "Go to section", sut.TransObject().TranslatedValue("MSG_GoToSection"), _
                    "A call with no scope reads the messages table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransObjectDefaultsToMessages", _
                         Err.Number, Err.Description
End Sub

'@sub-title A scope outside the five is refused.
'@TestMethod("LLTranslation")
Public Sub TestTransObjectRejectsUnknownScope()
    CustomTestSetTitles Assert, TESTMODULE, "TestTransObjectRejectsUnknownScope"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim answer As TranslationObject
    Dim errNumber As Long

    Set sut = LLTranslation.Create(TransSheet)

    On Error Resume Next
        Set answer = sut.TransObject(200)
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A scope outside the enum raises and names itself"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransObjectRejectsUnknownScope", _
                         Err.Number, Err.Description
End Sub

'@sub-title Create accepts a sheet with no ribbon table and that scope then raises.
'@details
'This pins the gap as it stands today. CheckRequirements calls T_TradLLRibbon
'optional and TransObject requires it, so a workbook can pass creation and fail
'later at the one call site that asks for the ribbon scope. The test states the
'behaviour so a change to either side trips it.
'@TestMethod("LLTranslation")
Public Sub TestRibbonTableIsOptionalAtCreationAndRequiredOnRead()
    CustomTestSetTitles Assert, TESTMODULE, "TestRibbonTableIsOptionalAtCreationAndRequiredOnRead"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim answer As TranslationObject
    Dim errNumber As Long

    TransSheet.ListObjects("T_TradLLRibbon").Delete

    Set sut = LLTranslation.Create(TransSheet)
    Assert.IsTrue (Not sut Is Nothing), "Create accepts a sheet with no ribbon table"

    On Error Resume Next
        Set answer = sut.TransObject(TranslationOfRibbon)
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "Asking for the ribbon scope on that sheet raises"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRibbonTableIsOptionalAtCreationAndRequiredOnRead", _
                         Err.Number, Err.Description
End Sub

'@section Language codes
'===============================================================================

'@sub-title InitialiseSheetNames writes the two codes and Value reads them back.
'@TestMethod("LLTranslation")
Public Sub TestSheetNamesRoundTripThroughValue()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetNamesRoundTripThroughValue"
    On Error GoTo TestFail

    Dim sut As LLTranslation

    Set sut = LLTranslation.Create(TransSheet)
    sut.InitialiseSheetNames "en", "fr"

    Assert.AreEqual "en", sut.Value("dictlang"), "The dictionary language reads back"
    Assert.AreEqual "fr", sut.Value("lllang"), "The interface language reads back"
    Assert.AreEqual "en", sut.Value("DictLang"), "The key is read without regard to case"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetNamesRoundTripThroughValue", _
                         Err.Number, Err.Description
End Sub

'@sub-title Writing the codes twice keeps the second value.
'@details
'EnsureName leaves an existing name alone, so the SetValue beside it is what
'makes a second generation over the same workbook write the new language.
'@TestMethod("LLTranslation")
Public Sub TestSheetNamesAreOverwrittenOnASecondWrite()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetNamesAreOverwrittenOnASecondWrite"
    On Error GoTo TestFail

    Dim sut As LLTranslation

    Set sut = LLTranslation.Create(TransSheet)
    sut.InitialiseSheetNames "en", "en"
    sut.InitialiseSheetNames "fr", "es"

    Assert.AreEqual "fr", sut.Value("dictlang"), "The second dictionary language won"
    Assert.AreEqual "es", sut.Value("lllang"), "The second interface language won"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetNamesAreOverwrittenOnASecondWrite", _
                         Err.Number, Err.Description
End Sub

'@sub-title An unknown key answers an empty string.
'@details
'This is the behaviour today. It is the same silent shape that hid the
'debugging password fault in LinelistSpecs, and it is written down here so the
'audit that is owed can decide about it.
'@TestMethod("LLTranslation")
Public Sub TestValueOfAnUnknownKeyIsEmpty()
    CustomTestSetTitles Assert, TESTMODULE, "TestValueOfAnUnknownKeyIsEmpty"
    On Error GoTo TestFail

    Dim sut As LLTranslation

    Set sut = LLTranslation.Create(TransSheet)

    Assert.AreEqual vbNullString, sut.Value("nosuchkey"), _
                    "An unknown key answers an empty string today"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueOfAnUnknownKeyIsEmpty", Err.Number, Err.Description
End Sub

'@section Export and Import
'===============================================================================

'@sub-title Export copies every table onto a sheet of the target workbook.
'@TestMethod("LLTranslation")
Public Sub TestExportCarriesEveryTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestExportCarriesEveryTable"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim targetWkb As Workbook
    Dim exportSheet As Worksheet

    Set sut = LLTranslation.Create(TransSheet)
    Set targetWkb = NewWorkbook()

    sut.Export targetWkb, Hide:=xlSheetVisible

    Set exportSheet = Nothing
    On Error Resume Next
        Set exportSheet = targetWkb.Worksheets(TRANS_SHEET_NAME)
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.IsTrue (Not exportSheet Is Nothing), _
                  "Export creates a sheet named after the host worksheet"
    Assert.AreEqual CLng(TransSheet.ListObjects.Count), _
                    CLng(exportSheet.ListObjects.Count), _
                    "Every table on the host sheet reaches the export"

    DeleteWorkbook targetWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not targetWkb Is Nothing Then DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportCarriesEveryTable", Err.Number, Err.Description
End Sub

'@sub-title ExportDictionary writes the dictionary table on its own.
'@TestMethod("LLTranslation")
Public Sub TestExportDictionaryWritesOneTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestExportDictionaryWritesOneTable"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim targetWkb As Workbook
    Dim exportSheet As Worksheet

    Set sut = LLTranslation.Create(TransSheet)
    Set targetWkb = NewWorkbook()

    sut.ExportDictionary targetWkb, sheetName:="Translations", Hide:=xlSheetVisible

    Set exportSheet = Nothing
    On Error Resume Next
        Set exportSheet = targetWkb.Worksheets("Translations")
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.IsTrue (Not exportSheet Is Nothing), "ExportDictionary creates the sheet"
    Assert.AreEqual CLng(1), CLng(exportSheet.ListObjects.Count), _
                    "The dictionary table travels on its own"
    Assert.AreEqual "DICT_Var1", CStr(exportSheet.Cells(2, 1).Value), _
                    "The dictionary rows came across"

    DeleteWorkbook targetWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not targetWkb Is Nothing Then DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportDictionaryWritesOneTable", _
                         Err.Number, Err.Description
End Sub

'@sub-title Import takes the values of the tables the source carries.
'@details
'The suite this replaced asserted True here, so it proved that Import raised
'nothing and nothing else. What matters is that the value moved.
'@TestMethod("LLTranslation")
Public Sub TestImportTakesTheValuesItFinds()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportTakesTheValuesItFinds"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim sourceWkb As Workbook
    Dim sourceSh As Worksheet

    Set sut = LLTranslation.Create(TransSheet)

    Set sourceWkb = NewWorkbook()
    Set sourceSh = sourceWkb.Worksheets.Add
    sourceSh.Name = TRANS_SHEET_NAME
    SeedTable sourceSh, 1, "T_TradLLMsg", Array(Array("MSG_GoToSection", "Updated section"))

    sut.Import sourceWkb

    On Error GoTo TestFail
    Assert.AreEqual "Updated section", _
                    sut.TransObject(TranslationOfMessages).TranslatedValue("MSG_GoToSection"), _
                    "The value of the matching table came across"
    Assert.AreEqual "Advanced", _
                    sut.TransObject(TranslationOfShapes).TranslatedValue("SHP_Advanced"), _
                    "A table the source does not carry is left alone"

    DeleteWorkbook sourceWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not sourceWkb Is Nothing Then DeleteWorkbook sourceWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestImportTakesTheValuesItFinds", Err.Number, Err.Description
End Sub

'@sub-title Import from a workbook with no matching sheet changes nothing.
'@details
'It returns quietly today. The test states that so the audit can decide whether
'the caller deserves to hear about it.
'@TestMethod("LLTranslation")
Public Sub TestImportWithNoMatchingSheetChangesNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportWithNoMatchingSheetChangesNothing"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim sourceWkb As Workbook

    Set sut = LLTranslation.Create(TransSheet)
    Set sourceWkb = NewWorkbook()

    sut.Import sourceWkb
    sut.Import Nothing

    On Error GoTo TestFail
    Assert.AreEqual "Go to section", _
                    sut.TransObject(TranslationOfMessages).TranslatedValue("MSG_GoToSection"), _
                    "A source with no translation sheet leaves the tables alone"

    DeleteWorkbook sourceWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not sourceWkb Is Nothing Then DeleteWorkbook sourceWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestImportWithNoMatchingSheetChangesNothing", _
                         Err.Number, Err.Description
End Sub

'@sub-title ImportDictionary refuses a Nothing table.
'@TestMethod("LLTranslation")
Public Sub TestImportDictionaryRejectsNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportDictionaryRejectsNothing"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim errNumber As Long

    Set sut = LLTranslation.Create(TransSheet)

    On Error Resume Next
        sut.ImportDictionary Nothing
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing dictionary table is refused"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestImportDictionaryRejectsNothing", _
                         Err.Number, Err.Description
End Sub

'@section Hidden names
'===============================================================================

'@sub-title InitialiseHiddenNames writes the translated values onto a workbook.
'@TestMethod("LLTranslation")
Public Sub TestInitialiseHiddenNamesCreatesNames()
    CustomTestSetTitles Assert, TESTMODULE, "TestInitialiseHiddenNamesCreatesNames"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim targetWkb As Workbook
    Dim targetNames As HiddenNames

    Set sut = LLTranslation.Create(TransSheet)
    Set targetWkb = NewWorkbook()

    sut.InitialiseHiddenNames targetWkb

    Set targetNames = HiddenNames.Create(targetWkb)

    On Error GoTo TestFail
    Assert.IsTrue targetNames.HasName("RNG_GoToSection"), "RNG_GoToSection was created"
    Assert.IsTrue targetNames.HasName("RNG_Week"), "RNG_Week was created"
    Assert.IsTrue targetNames.HasName("RNG_SPTSheet"), "RNG_SPTSheet was created"
    Assert.AreEqual "Go to section", targetNames.ValueAsString("RNG_GoToSection"), _
                    "RNG_GoToSection carries the translated message"
    Assert.AreEqual "SpatioTemporal", targetNames.ValueAsString("RNG_SPTSheet"), _
                    "The sheet names carry their translated value"

    DeleteWorkbook targetWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not targetWkb Is Nothing Then DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestInitialiseHiddenNamesCreatesNames", _
                         Err.Number, Err.Description
End Sub

'@sub-title InitialiseHiddenNames refuses a Nothing workbook.
'@TestMethod("LLTranslation")
Public Sub TestInitialiseHiddenNamesRejectsNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestInitialiseHiddenNamesRejectsNothing"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim errNumber As Long

    Set sut = LLTranslation.Create(TransSheet)

    On Error Resume Next
        sut.InitialiseHiddenNames Nothing
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing workbook is refused"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestInitialiseHiddenNamesRejectsNothing", _
                         Err.Number, Err.Description
End Sub
