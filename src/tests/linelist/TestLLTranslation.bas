Attribute VB_Name = "TestLLTranslation"
Attribute VB_Description = "Tests for LLTranslation class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLTranslation class")

'@description
'Drives LLTranslation, which owns the five translation tables of a linelist:
'the messages, the shapes, the form labels, the ribbon labels, and
'Tab_Translations. That last one is the setup's own translation table, copied
'onto the linelist translation worksheet while the linelist is built, and it is
'what translates the dictionary, the choices and the analyses. Every linelist
'form module builds one of these, and so do InitTransfer, the exporter and
'LinelistSpecs.
'
'THE FIXTURE SPEAKS TWO LANGUAGES
'-------------------------------------------------------------------------------
'Every table carries an "en" column and a "fr" column with a different value in
'each. A fixture with one language lets a routing test pass while the class
'reads the wrong column, and it leaves the empty-language case impossible to
'reach.
'
'The tables are spaced six columns apart. An import that takes the source
'headers grows the host table by a column, and a table with a neighbour one
'column away has nowhere to grow into.
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

'@sub-title Build a translation worksheet with the five tables and two languages
'@param targetWkb Workbook. The workbook to seed.
Private Sub SeedTranslationSheet(ByVal targetWkb As Workbook)
    Dim sh As Worksheet

    Set sh = targetWkb.Worksheets.Add
    sh.Name = TRANS_SHEET_NAME

    'MSG_GoToSec is the code the go-to captions are really built from, and it is
    'what InitialiseHiddenNames reads for RNG_GoToSection. MSG_GoToSection stays
    'beside it because the translation tests below use it as their sample row.
    SeedTable sh, 1, "T_TradLLMsg", "en", "fr", Array( _
        Array("MSG_GoToSec", "Go to section", "Aller a la section"), _
        Array("MSG_GoToSection", "Go to section", "Aller a la section"), _
        Array("MSG_AnaPeriod", "Analysis period", "Periode d analyse"), _
        Array("MSG_GoToHead", "Go to header", "Aller a l entete"), _
        Array("MSG_NoDevide", "Do not split", "Ne pas diviser"), _
        Array("MSG_Devide", "Split", "Diviser"), _
        Array("MSG_GoToGraph", "Go to graph", "Aller au graphique"), _
        Array("MSG_ComputeOnFiltered", "Compute on filtered", "Calculer sur le filtre"), _
        Array("LLSHEET_CustomChoice", "Custom dropdown", "Liste personnalisee"), _
        Array("LLSHEET_Analysis", "Analysis", "Analyse"), _
        Array("LLSHEET_TemporalAnalysis", "Temporal", "Temporel"), _
        Array("LLSHEET_SpatialAnalysis", "Spatial", "Spatial FR"), _
        Array("LLSHEET_SpatioTemporalAnalysis", "SpatioTemporal", "Spatio temporel"), _
        Array("LLSHEET_CustomPivotTable", "Custom pivot", "Pivot personnalise"), _
        Array("MSG_W", "W", "S"), _
        Array("MSG_Q", "Q", "T"), _
        Array("MSG_InfoStart", "Info start", "Debut info"), _
        Array("MSG_InfoEnd", "Info end", "Fin info"))

    SeedTable sh, 7, "T_TradLLShapes", "en", "fr", _
              Array(Array("SHP_Advanced", "Advanced", "Avance"))
    SeedTable sh, 13, "T_TradLLForms", "en", "fr", _
              Array(Array("FRM_Title", "Form title", "Titre du formulaire"))
    SeedTable sh, 19, "Tab_Translations", "en", "fr", _
              Array(Array("DICT_Var1", "Variable 1", "Variable un"))
    SeedTable sh, 25, "T_TradLLRibbon", "en", "fr", _
              Array(Array("RIB_Advanced", "Ribbon advanced", "Ruban avance"))

    'TransObject reads both language codes from the WORKBOOK store.
    SetWorkbookLanguage targetWkb, "RNG_LLLanguageCode", "en"
    SetWorkbookLanguage targetWkb, "RNG_DictionaryLanguage", "en"
End Sub

'@sub-title Write one translation table and name it.
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

'@sub-title Give a workbook-level language hidden name its value.
'@param targetWkb Workbook. The workbook carrying the name.
'@param nameId String. The hidden name.
'@param value String. The language code to store.
Private Sub SetWorkbookLanguage(ByVal targetWkb As Workbook, _
                                ByVal nameId As String, ByVal value As String)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(targetWkb)
    If store.HasName(nameId) Then
        store.SetValue nameId, value
    Else
        store.EnsureName nameId, value, HiddenNameTypeString
    End If
End Sub

'@sub-title Replace one tag of the messages table with another.
'@details
'A loop rather than Range.Find, which inherits LookIn and SearchOrder from the
'last search of the Excel session.
'@param oldTag String. The tag to replace.
'@param newTag String. The tag to write in its place.
Private Sub RetagMessage(ByVal oldTag As String, ByVal newTag As String)
    Dim tagColumn As Range
    Dim cellRng As Range

    Set tagColumn = TransSheet.ListObjects("T_TradLLMsg").ListColumns(1).DataBodyRange

    For Each cellRng In tagColumn
        If StrComp(CStr(cellRng.Value), oldTag, vbBinaryCompare) = 0 Then
            cellRng.Value = newTag
            Exit Sub
        End If
    Next
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

'@sub-title The ribbon table is required at creation.
'@details
'It used to be optional there and required on read, so a sheet passed Create
'and the ribbon callback died later at the one call site that asks for that
'scope, with no error handler above it. The designer template ships the table.
'@TestMethod("LLTranslation")
Public Sub TestCreateRequiresTheRibbonTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRequiresTheRibbonTable"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim errNumber As Long

    TransSheet.ListObjects("T_TradLLRibbon").Delete

    On Error Resume Next
        Set sut = LLTranslation.Create(TransSheet)
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "A sheet with no ribbon table is refused at creation"
    Assert.IsTrue (sut Is Nothing), "Nothing comes back from the refused create"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRequiresTheRibbonTable", _
                         Err.Number, Err.Description
End Sub

'@sub-title The host worksheet cannot be swapped after creation.
'@details
'The instance caches a hidden name store taken from that worksheet's workbook
'and five translation objects taken from its tables, so a swap would leave all
'six pointing at a worksheet the instance had given up.
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

'@section Language codes
'===============================================================================

'@sub-title The interface code picks the column for four scopes and the dictionary code for the fifth.
'@details
'Every table of the fixture carries a different value under "en" and under
'"fr", so a scope reading the wrong column is visible here.
'@TestMethod("LLTranslation")
Public Sub TestTheLanguageCodeDecidesTheColumn()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheLanguageCodeDecidesTheColumn"
    On Error GoTo TestFail

    Dim sut As LLTranslation

    SetWorkbookLanguage FixtureWkb, "RNG_LLLanguageCode", "fr"

    Set sut = LLTranslation.Create(TransSheet)

    Assert.AreEqual "Aller a la section", _
                    sut.TransObject(TranslationOfMessages).TranslatedValue("MSG_GoToSection"), _
                    "The messages scope answers in the interface language"
    Assert.AreEqual "Avance", _
                    sut.TransObject(TranslationOfShapes).TranslatedValue("SHP_Advanced"), _
                    "The shapes scope answers in the interface language"
    Assert.AreEqual "Ruban avance", _
                    sut.TransObject(TranslationOfRibbon).TranslatedValue("RIB_Advanced"), _
                    "The ribbon scope answers in the interface language"
    Assert.AreEqual "Variable 1", _
                    sut.TransObject(TranslationOfDictionary).TranslatedValue("DICT_Var1"), _
                    "The dictionary scope keeps its own code, which is still en"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheLanguageCodeDecidesTheColumn", _
                         Err.Number, Err.Description
End Sub

'@sub-title An empty interface language code raises.
'@details
'An empty code used to give a translation object whose language matched no
'column header, and every tag then translated to itself: buttons reading
'SHP_Advanced, worksheets named LLSHEET_Analysis, and no error anywhere.
'@TestMethod("LLTranslation")
Public Sub TestAnAbsentLanguageCodeRaises()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnAbsentLanguageCodeRaises"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim store As HiddenNames
    Dim answer As TranslationObject
    Dim errNumber As Long
    Dim dictErrNumber As Long

    Set store = HiddenNames.Create(FixtureWkb)
    store.RemoveName "RNG_LLLanguageCode"
    store.RemoveName "RNG_DictionaryLanguage"

    Set sut = LLTranslation.Create(TransSheet)

    On Error Resume Next
        Set answer = sut.TransObject(TranslationOfMessages)
        errNumber = Err.Number
    On Error GoTo 0

    On Error Resume Next
        Set answer = sut.TransObject(TranslationOfDictionary)
        dictErrNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "An empty interface language code raises"
    Assert.AreEqual CLng(ProjectError.ElementNotFound), dictErrNumber, _
                    "An empty dictionary language code raises"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnAbsentLanguageCodeRaises", _
                         Err.Number, Err.Description
End Sub

'@sub-title A scope is built once, and Refresh drops what is held.
'@details
'Each translation object holds its own copy of its table. A caller that keeps
'the instance keeps the copy, which is the whole point of holding them here.
'@TestMethod("LLTranslation")
Public Sub TestScopeIsBuiltOnceAndRefreshDropsIt()
    CustomTestSetTitles Assert, TESTMODULE, "TestScopeIsBuiltOnceAndRefreshDropsIt"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim firstRead As TranslationObject
    Dim secondRead As TranslationObject
    Dim afterRefresh As TranslationObject

    Set sut = LLTranslation.Create(TransSheet)

    Set firstRead = sut.TransObject(TranslationOfMessages)
    Set secondRead = sut.TransObject(TranslationOfMessages)

    Assert.IsTrue (firstRead Is secondRead), _
                  "The second read of a scope gives the object the first one built"

    sut.Refresh
    Set afterRefresh = sut.TransObject(TranslationOfMessages)

    Assert.IsTrue (Not (afterRefresh Is firstRead)), _
                  "Refresh drops the held object and the next read builds another"
    Assert.AreEqual "Go to section", afterRefresh.TranslatedValue("MSG_GoToSection"), _
                    "The rebuilt object still reads its table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestScopeIsBuiltOnceAndRefreshDropsIt", _
                         Err.Number, Err.Description
End Sub

'@section Export and Import
'===============================================================================

'@sub-title Export carries every table and the values inside them.
'@details
'A count alone passes for an export that wrote five empty tables.
'@TestMethod("LLTranslation")
Public Sub TestExportCarriesEveryTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestExportCarriesEveryTable"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim targetWkb As Workbook
    Dim exportSheet As Worksheet
    Dim exported As LLTranslation

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

    'The exported sheet is read back through the class, which is what proves
    'the five tables arrived with their values and their language columns.
    SetWorkbookLanguage targetWkb, "RNG_LLLanguageCode", "en"
    SetWorkbookLanguage targetWkb, "RNG_DictionaryLanguage", "en"
    Set exported = LLTranslation.Create(exportSheet)

    Assert.AreEqual "Go to section", _
                    exported.TransObject(TranslationOfMessages).TranslatedValue("MSG_GoToSection"), _
                    "The messages table carried its values"
    Assert.AreEqual "Advanced", _
                    exported.TransObject(TranslationOfShapes).TranslatedValue("SHP_Advanced"), _
                    "The shapes table carried its values"
    Assert.AreEqual "Form title", _
                    exported.TransObject(TranslationOfForms).TranslatedValue("FRM_Title"), _
                    "The forms table carried its values"
    Assert.AreEqual "Variable 1", _
                    exported.TransObject(TranslationOfDictionary).TranslatedValue("DICT_Var1"), _
                    "The dictionary table carried its values"
    Assert.AreEqual "Ribbon advanced", _
                    exported.TransObject(TranslationOfRibbon).TranslatedValue("RIB_Advanced"), _
                    "The ribbon table carried its values"

    DeleteWorkbook targetWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not targetWkb Is Nothing Then DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportCarriesEveryTable", Err.Number, Err.Description
End Sub

'@sub-title Exporting twice into one workbook works the second time.
'@details
'Clearing the cells of the sheet leaves the ListObject objects standing, each
'one still holding its table name, and the second export asks Excel for those
'same five names.
'@TestMethod("LLTranslation")
Public Sub TestExportTwiceIntoOneWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "TestExportTwiceIntoOneWorkbook"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim targetWkb As Workbook
    Dim exportSheet As Worksheet
    Dim errNumber As Long

    Set sut = LLTranslation.Create(TransSheet)
    Set targetWkb = NewWorkbook()

    sut.Export targetWkb, Hide:=xlSheetVisible

    On Error Resume Next
        sut.Export targetWkb, Hide:=xlSheetVisible
        errNumber = Err.Number
    On Error GoTo 0

    Set exportSheet = Nothing
    On Error Resume Next
        Set exportSheet = targetWkb.Worksheets(TRANS_SHEET_NAME)
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(0), errNumber, "The second export into one workbook raises nothing"
    Assert.IsTrue (Not exportSheet Is Nothing), "The sheet is still there after the second export"
    Assert.AreEqual CLng(TransSheet.ListObjects.Count), _
                    CLng(exportSheet.ListObjects.Count), _
                    "The second export leaves five tables, one per source table"

    DeleteWorkbook targetWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not targetWkb Is Nothing Then DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportTwiceIntoOneWorkbook", Err.Number, Err.Description
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

'@sub-title Import takes the values of the tables the source carries and leaves the rest standing.
'@details
'The dictionary assertion is the one that matters. The dictionary table used
'to be emptied before the loop had discovered what the source carried, and it
'was emptied through a CustomTable with no key column, which takes the tag
'column with the rest.
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
    SeedTable sourceSh, 1, "T_TradLLMsg", "en", "fr", _
              Array(Array("MSG_GoToSection", "Updated section", "Section mise a jour"))

    sut.Import sourceWkb

    On Error GoTo TestFail
    Assert.AreEqual "Updated section", _
                    sut.TransObject(TranslationOfMessages).TranslatedValue("MSG_GoToSection"), _
                    "The value of the matching table came across"
    Assert.AreEqual "Advanced", _
                    sut.TransObject(TranslationOfShapes).TranslatedValue("SHP_Advanced"), _
                    "A table the source does not carry is left alone"
    Assert.AreEqual "Variable 1", _
                    sut.TransObject(TranslationOfDictionary).TranslatedValue("DICT_Var1"), _
                    "The dictionary table is left alone when the source has none"

    DeleteWorkbook sourceWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not sourceWkb Is Nothing Then DeleteWorkbook sourceWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestImportTakesTheValuesItFinds", Err.Number, Err.Description
End Sub

'@sub-title Import says so when there is nothing to import from.
'@details
'Export raises on a Nothing workbook and Import used to return quietly on both
'a Nothing workbook and a source with no matching worksheet. One pair of
'operations, one contract.
'@TestMethod("LLTranslation")
Public Sub TestImportRaisesWhenThereIsNothingToImport()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportRaisesWhenThereIsNothingToImport"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim sourceWkb As Workbook
    Dim nothingErrNumber As Long
    Dim missingSheetErrNumber As Long

    Set sut = LLTranslation.Create(TransSheet)
    Set sourceWkb = NewWorkbook()

    On Error Resume Next
        sut.Import Nothing
        nothingErrNumber = Err.Number
    On Error GoTo 0

    On Error Resume Next
        sut.Import sourceWkb
        missingSheetErrNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), nothingErrNumber, _
                    "A Nothing source workbook is refused"
    Assert.AreEqual CLng(ProjectError.ElementNotFound), missingSheetErrNumber, _
                    "A source with no translation worksheet is refused"
    Assert.AreEqual "Go to section", _
                    sut.TransObject(TranslationOfMessages).TranslatedValue("MSG_GoToSection"), _
                    "The host tables are left standing"

    DeleteWorkbook sourceWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not sourceWkb Is Nothing Then DeleteWorkbook sourceWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestImportRaisesWhenThereIsNothingToImport", _
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

'@sub-title ImportDictionary takes the language columns of the source.
'@details
'This is the most consequential thing the class does during a build: the host
'dictionary table comes out carrying the source workbook's language columns.
'The source here speaks en and es, the host speaks en and fr, and the host
'answers in es afterwards.
'@TestMethod("LLTranslation")
Public Sub TestImportDictionaryTakesTheSourceHeaders()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportDictionaryTakesTheSourceHeaders"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim sourceWkb As Workbook
    Dim sourceSh As Worksheet

    Set sut = LLTranslation.Create(TransSheet)

    Set sourceWkb = NewWorkbook()
    Set sourceSh = sourceWkb.Worksheets.Add
    sourceSh.Name = "SourceDictionary"
    SeedTable sourceSh, 1, "Tab_SourceTranslations", "en", "es", _
              Array(Array("DICT_Var1", "Variable one", "Variable uno"))

    sut.ImportDictionary sourceSh.ListObjects("Tab_SourceTranslations")

    Assert.AreEqual "Variable one", _
                    sut.TransObject(TranslationOfDictionary).TranslatedValue("DICT_Var1"), _
                    "The values of the source reached the host table"

    SetWorkbookLanguage FixtureWkb, "RNG_DictionaryLanguage", "es"
    sut.Refresh

    Assert.AreEqual "Variable uno", _
                    sut.TransObject(TranslationOfDictionary).TranslatedValue("DICT_Var1"), _
                    "The host answers in a language only the source carried"

    DeleteWorkbook sourceWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not sourceWkb Is Nothing Then DeleteWorkbook sourceWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestImportDictionaryTakesTheSourceHeaders", _
                         Err.Number, Err.Description
End Sub

'@section Hidden names
'===============================================================================

'@sub-title InitialiseHiddenNames writes all seventeen translated values onto a workbook.
'@details
'Three of the seventeen used to be checked, so a typo in any of the other
'fourteen tag strings was invisible.
'@TestMethod("LLTranslation")
Public Sub TestInitialiseHiddenNamesCreatesNames()
    CustomTestSetTitles Assert, TESTMODULE, "TestInitialiseHiddenNamesCreatesNames"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim targetWkb As Workbook
    Dim targetNames As HiddenNames
    Dim expected As Variant
    Dim idx As Long

    Set sut = LLTranslation.Create(TransSheet)
    Set targetWkb = NewWorkbook()

    sut.InitialiseHiddenNames targetWkb

    Set targetNames = HiddenNames.Create(targetWkb)

    expected = Array( _
        Array("RNG_GoToSection", "Go to section"), _
        Array("RNG_AnaPeriod", "Analysis period"), _
        Array("RNG_GoToHeader", "Go to header"), _
        Array("RNG_NoDevide", "Do not split"), _
        Array("RNG_Devide", "Split"), _
        Array("RNG_GoToGraph", "Go to graph"), _
        Array("RNG_OnFiltered", "Compute on filtered"), _
        Array("RNG_CustomDrop", "Custom dropdown"), _
        Array("RNG_UASheet", "Analysis"), _
        Array("RNG_TSSheet", "Temporal"), _
        Array("RNG_SPSheet", "Spatial"), _
        Array("RNG_SPTSheet", "SpatioTemporal"), _
        Array("RNG_CustomPivot", "Custom pivot"), _
        Array("RNG_Week", "W"), _
        Array("RNG_Quarter", "Q"), _
        Array("RNG_InfoStart", "Info start"), _
        Array("RNG_InfoEnd", "Info end"))

    On Error GoTo TestFail
    For idx = LBound(expected) To UBound(expected)
        Assert.AreEqual CStr(expected(idx)(1)), _
                        targetNames.ValueAsString(CStr(expected(idx)(0))), _
                        CStr(expected(idx)(0)) & " carries its translated value"
    Next idx

    DeleteWorkbook targetWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not targetWkb Is Nothing Then DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestInitialiseHiddenNamesCreatesNames", _
                         Err.Number, Err.Description
End Sub

'@sub-title A tag absent from the messages table gives a hidden name holding the tag.
'@details
'TranslatedValue answers with the tag itself when it finds no row for it, so
'the hidden name carries "MSG_Q" and every reader downstream shows that. This
'test states the behaviour as it is. Refusing to write such a name would need
'the seventeen tags of the shipped designer checked one by one first, because
'a tag the template never carried would then stop a generation.
'@TestMethod("LLTranslation")
Public Sub TestAnAbsentTagGivesAHiddenNameHoldingTheTag()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnAbsentTagGivesAHiddenNameHoldingTheTag"
    On Error GoTo TestFail

    Dim sut As LLTranslation
    Dim targetWkb As Workbook
    Dim targetNames As HiddenNames

    RetagMessage "MSG_Q", "MSG_QUARTER_RENAMED"

    Set sut = LLTranslation.Create(TransSheet)
    Set targetWkb = NewWorkbook()

    sut.InitialiseHiddenNames targetWkb
    Set targetNames = HiddenNames.Create(targetWkb)

    On Error GoTo TestFail
    Assert.AreEqual "MSG_Q", targetNames.ValueAsString("RNG_Quarter"), _
                    "A tag with no row gives a hidden name holding the tag"
    Assert.AreEqual "W", targetNames.ValueAsString("RNG_Week"), _
                    "The tags that are there are unaffected"

    DeleteWorkbook targetWkb
    Exit Sub
TestFail:
    On Error Resume Next
        If Not targetWkb Is Nothing Then DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestAnAbsentTagGivesAHiddenNameHoldingTheTag", _
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
