Attribute VB_Name = "TestLLTranslation"
Attribute VB_Description = "Tests for LLTranslation class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for LLTranslation class")

Option Explicit

Private Assert As ICustomTest
Private FixtureWkb As Workbook
Private TransSheet As Worksheet

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "LLTranslation"
Private Const TRANS_SHEET_NAME As String = "LinelistTranslation"


'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLLTranslation"
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
    Set FixtureWkb = TestHelpers.NewWorkbook
    SeedTranslationSheet FixtureWkb
    Set TransSheet = FixtureWkb.Worksheets(TRANS_SHEET_NAME)
    Assert.BeginTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not FixtureWkb Is Nothing Then TestHelpers.DeleteWorkbook FixtureWkb
    On Error GoTo 0

    Set TransSheet = Nothing
    Set FixtureWkb = Nothing
End Sub


'@section Test Fixture Helpers
'===============================================================================

'@sub-title Build a translation worksheet with required ListObjects and HiddenNames
Private Sub SeedTranslationSheet(ByVal targetWkb As Workbook)
    Dim sh As Worksheet
    Dim rng As Range
    Dim wkbNames As IHiddenNames

    Set sh = targetWkb.Worksheets.Add
    sh.Name = TRANS_SHEET_NAME

    'T_TradLLMsg: label column + language column
    sh.Cells(1, 1).Value = "label"
    sh.Cells(1, 2).Value = "en"
    sh.Cells(2, 1).Value = "MSG_GoToSection"
    sh.Cells(2, 2).Value = "Go to section"
    sh.Cells(3, 1).Value = "MSG_AnaPeriod"
    sh.Cells(3, 2).Value = "Analysis period"
    sh.Cells(4, 1).Value = "MSG_GoToHead"
    sh.Cells(4, 2).Value = "Go to header"
    sh.Cells(5, 1).Value = "MSG_NoDevide"
    sh.Cells(5, 2).Value = "Don't split"
    sh.Cells(6, 1).Value = "MSG_Devide"
    sh.Cells(6, 2).Value = "Split"
    sh.Cells(7, 1).Value = "MSG_GoToGraph"
    sh.Cells(7, 2).Value = "Go to graph"
    sh.Cells(8, 1).Value = "MSG_ComputeOnFiltered"
    sh.Cells(8, 2).Value = "Compute on filtered"
    sh.Cells(9, 1).Value = "LLSHEET_CustomChoice"
    sh.Cells(9, 2).Value = "Custom dropdown"
    sh.Cells(10, 1).Value = "LLSHEET_Analysis"
    sh.Cells(10, 2).Value = "Analysis"
    sh.Cells(11, 1).Value = "LLSHEET_TemporalAnalysis"
    sh.Cells(11, 2).Value = "Temporal"
    sh.Cells(12, 1).Value = "LLSHEET_SpatialAnalysis"
    sh.Cells(12, 2).Value = "Spatial"
    sh.Cells(13, 1).Value = "LLSHEET_SpatioTemporalAnalysis"
    sh.Cells(13, 2).Value = "SpatioTemporal"
    sh.Cells(14, 1).Value = "LLSHEET_CustomPivotTable"
    sh.Cells(14, 2).Value = "Custom pivot"
    sh.Cells(15, 1).Value = "MSG_W"
    sh.Cells(15, 2).Value = "W"
    sh.Cells(16, 1).Value = "MSG_Q"
    sh.Cells(16, 2).Value = "Q"
    sh.Cells(17, 1).Value = "MSG_InfoStart"
    sh.Cells(17, 2).Value = "Info start"
    sh.Cells(18, 1).Value = "MSG_InfoEnd"
    sh.Cells(18, 2).Value = "Info end"

    Set rng = sh.Range(sh.Cells(1, 1), sh.Cells(18, 2))
    sh.ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, XlListObjectHasHeaders:=xlYes).Name = "T_TradLLMsg"

    'T_TradLLShapes: label column + language column
    sh.Cells(1, 4).Value = "label"
    sh.Cells(1, 5).Value = "en"
    sh.Cells(2, 4).Value = "SHP_Advanced"
    sh.Cells(2, 5).Value = "Advanced"

    Set rng = sh.Range(sh.Cells(1, 4), sh.Cells(2, 5))
    sh.ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, XlListObjectHasHeaders:=xlYes).Name = "T_TradLLShapes"

    'T_TradLLForms: label column + language column
    sh.Cells(1, 7).Value = "label"
    sh.Cells(1, 8).Value = "en"
    sh.Cells(2, 7).Value = "FRM_Title"
    sh.Cells(2, 8).Value = "Form title"

    Set rng = sh.Range(sh.Cells(1, 7), sh.Cells(2, 8))
    sh.ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, XlListObjectHasHeaders:=xlYes).Name = "T_TradLLForms"

    'Tab_Translations (dictionary): label column + language column
    sh.Cells(1, 10).Value = "label"
    sh.Cells(1, 11).Value = "en"
    sh.Cells(2, 10).Value = "DICT_Var1"
    sh.Cells(2, 11).Value = "Variable 1"

    Set rng = sh.Range(sh.Cells(1, 10), sh.Cells(2, 11))
    sh.ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, XlListObjectHasHeaders:=xlYes).Name = "Tab_Translations"

    'Workbook-level HiddenNames for language codes
    Set wkbNames = HiddenNames.Create(targetWkb)
    wkbNames.EnsureName "RNG_LLLanguageCode", "en", HiddenNameTypeString
    wkbNames.EnsureName "RNG_DictionaryLanguage", "en", HiddenNameTypeString
End Sub


'@section Factory Tests
'===============================================================================

'@TestMethod("LLTranslation")
Public Sub TestCreateReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsInstance"
    On Error GoTo TestFail

    Dim sut As ILLTranslation
    Set sut = LLTranslation.Create(TransSheet)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsInstance", Err.Number, Err.Description
End Sub

'@TestMethod("LLTranslation")
Public Sub TestCreateRejectsNothingSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingSheet"
    On Error GoTo ExpectError

    Dim sut As ILLTranslation
    Set sut = LLTranslation.Create(Nothing)

    CustomTestLogFailure Assert, "TestCreateRejectsNothingSheet", , _
                         "Expected error when sheet is Nothing"
    Exit Sub
ExpectError:
    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when sheet is Nothing"
End Sub

'@TestMethod("LLTranslation")
Public Sub TestCreateRejectsMissingTables()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsMissingTables"
    On Error GoTo ExpectError

    'Create a sheet without the required ListObjects
    Dim emptyWkb As Workbook
    Dim emptySheet As Worksheet

    Set emptyWkb = TestHelpers.NewWorkbook
    Set emptySheet = emptyWkb.Worksheets(1)

    Dim sut As ILLTranslation
    Set sut = LLTranslation.Create(emptySheet)

    TestHelpers.DeleteWorkbook emptyWkb

    CustomTestLogFailure Assert, "TestCreateRejectsMissingTables", , _
                         "Expected error when required tables are missing"
    Exit Sub
ExpectError:
    On Error Resume Next
    If Not emptyWkb Is Nothing Then TestHelpers.DeleteWorkbook emptyWkb
    On Error GoTo 0

    Assert.IsTrue Err.Number <> 0, _
                  "Should raise an error when required tables are missing"
End Sub


'@section Property Tests
'===============================================================================

'@TestMethod("LLTranslation")
Public Sub TestTransObjectReturnsTranslationObject()
    CustomTestSetTitles Assert, TESTMODULE, "TestTransObjectReturnsTranslationObject"
    On Error GoTo TestFail

    Dim sut As ILLTranslation
    Set sut = LLTranslation.Create(TransSheet)

    Dim transObj As ITranslationObject
    Set transObj = sut.TransObject(TranslationOfMessages)

    Assert.IsTrue Not transObj Is Nothing, _
                  "TransObject should return a non-Nothing ITranslationObject"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransObjectReturnsTranslationObject", Err.Number, Err.Description
End Sub

'@TestMethod("LLTranslation")
Public Sub TestWkshReturnsHostSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestWkshReturnsHostSheet"
    On Error GoTo TestFail

    Dim sut As ILLTranslation
    Set sut = LLTranslation.Create(TransSheet)

    Assert.IsTrue sut.Wksh.Name = TRANS_SHEET_NAME, _
                  "Wksh should return the translation worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWkshReturnsHostSheet", Err.Number, Err.Description
End Sub


'@section Export/Import Tests
'===============================================================================

'@TestMethod("LLTranslation")
Public Sub TestExportCreatesSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestExportCreatesSheet"
    On Error GoTo TestFail

    Dim sut As ILLTranslation
    Set sut = LLTranslation.Create(TransSheet)

    Dim targetWkb As Workbook
    Set targetWkb = TestHelpers.NewWorkbook

    sut.Export targetWkb, Hide:=xlSheetVisible

    Dim exportSheet As Worksheet
    Set exportSheet = Nothing

    On Error Resume Next
    Set exportSheet = targetWkb.Worksheets(TRANS_SHEET_NAME)
    On Error GoTo 0

    Assert.IsTrue Not exportSheet Is Nothing, _
                  "Export should create a sheet named '" & TRANS_SHEET_NAME & "'"

    Assert.IsTrue exportSheet.ListObjects.Count > 0, _
                  "Exported sheet should contain at least one ListObject"

    TestHelpers.DeleteWorkbook targetWkb

    Exit Sub
TestFail:
    On Error Resume Next
    If Not targetWkb Is Nothing Then TestHelpers.DeleteWorkbook targetWkb
    On Error GoTo 0

    CustomTestLogFailure Assert, "TestExportCreatesSheet", Err.Number, Err.Description
End Sub

'@TestMethod("LLTranslation")
Public Sub TestImportSkipsMissingTables()
    CustomTestSetTitles Assert, TESTMODULE, "TestImportSkipsMissingTables"
    On Error GoTo TestFail

    Dim sut As ILLTranslation
    Set sut = LLTranslation.Create(TransSheet)

    'Create a source workbook with only one matching table
    Dim sourceWkb As Workbook
    Dim sourceSh As Worksheet
    Dim rng As Range

    Set sourceWkb = TestHelpers.NewWorkbook
    Set sourceSh = sourceWkb.Worksheets.Add
    sourceSh.Name = TRANS_SHEET_NAME

    'Only T_TradLLMsg present (missing T_TradLLShapes, T_TradLLForms, Tab_Translations)
    sourceSh.Cells(1, 1).Value = "label"
    sourceSh.Cells(1, 2).Value = "en"
    sourceSh.Cells(2, 1).Value = "MSG_GoToSection"
    sourceSh.Cells(2, 2).Value = "Updated section"

    Set rng = sourceSh.Range(sourceSh.Cells(1, 1), sourceSh.Cells(2, 2))
    sourceSh.ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, XlListObjectHasHeaders:=xlYes).Name = "T_TradLLMsg"

    'Import should succeed without error
    sut.Import sourceWkb

    TestHelpers.DeleteWorkbook sourceWkb

    Assert.IsTrue True, _
                  "Import should complete without error when some tables are missing"

    Exit Sub
TestFail:
    On Error Resume Next
    If Not sourceWkb Is Nothing Then TestHelpers.DeleteWorkbook sourceWkb
    On Error GoTo 0

    CustomTestLogFailure Assert, "TestImportSkipsMissingTables", Err.Number, Err.Description
End Sub

'@TestMethod("LLTranslation")
Public Sub TestInitialiseHiddenNamesCreatesNames()
    CustomTestSetTitles Assert, TESTMODULE, "TestInitialiseHiddenNamesCreatesNames"
    On Error GoTo TestFail

    Dim sut As ILLTranslation
    Set sut = LLTranslation.Create(TransSheet)

    Dim targetWkb As Workbook
    Set targetWkb = TestHelpers.NewWorkbook

    sut.InitialiseHiddenNames targetWkb

    Dim targetNames As IHiddenNames
    Set targetNames = HiddenNames.Create(targetWkb)

    Assert.IsTrue targetNames.HasName("RNG_GoToSection"), _
                  "InitialiseHiddenNames should create RNG_GoToSection"

    Assert.IsTrue targetNames.HasName("RNG_Week"), _
                  "InitialiseHiddenNames should create RNG_Week"

    Assert.AreEqual "Go to section", targetNames.ValueAsString("RNG_GoToSection"), _
                    "RNG_GoToSection should contain the translated value of MSG_GoToSection"

    TestHelpers.DeleteWorkbook targetWkb

    Exit Sub
TestFail:
    On Error Resume Next
    If Not targetWkb Is Nothing Then TestHelpers.DeleteWorkbook targetWkb
    On Error GoTo 0

    CustomTestLogFailure Assert, "TestInitialiseHiddenNamesCreatesNames", Err.Number, Err.Description
End Sub
