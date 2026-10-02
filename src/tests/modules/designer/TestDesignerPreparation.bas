Attribute VB_Name = "TestDesignerPreparation"
Attribute VB_Description = "Unit tests for DesignerPreparation class"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates DesignerPreparation for persisted flags, sheet hiding, dropdown creation, T_Multi and Main validation.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private MainSheet As Worksheet
Private TranslationSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    TestHelpers.BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDesignerPreparation"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    TestHelpers.RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    TestHelpers.BusyApp

    Set FixtureWorkbook = TestHelpers.NewWorkbook
    Set MainSheet = TestHelpers.EnsureWorksheet("Main", FixtureWorkbook)
    Set TranslationSheet = TestHelpers.EnsureWorksheet("DesignerTranslation", FixtureWorkbook)

    FixtureWorkbook.Names.Add Name:="RNG_MainLangCode", RefersTo:=TranslationSheet.Range("A1")
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        TestHelpers.DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set TranslationSheet = Nothing
    Set MainSheet = Nothing
    Set FixtureWorkbook = Nothing

    TestHelpers.RestoreApp
End Sub


'@section DesignerPreparation Tests
'===============================================================================
'@TestMethod("DesignerPreparation")
Public Sub TestPrepareSeedsFlags()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareSeedsFlags"
    On Error GoTo Fail

    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    Assert.IsTrue subject.GetFlag("chkAlert"), "Alert flag should default to on."
    Assert.IsTrue subject.GetFlag("chkInstruct"), "Instruction flag should default to on."

    subject.SetFlag "chkAlert", False
    Assert.IsFalse subject.GetFlag("chkAlert"), "Alert flag should persist changes."
    Assert.AreEqual "No", subject.HiddenStore.ValueAsString("chkAlert"), "Hidden name should store No for disabled flags."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSeedsFlags"
End Sub

'@TestMethod("DesignerPreparation")
Public Sub TestPrepareHidesInternalSheets()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareHidesInternalSheets"
    On Error GoTo Fail

    'Arrange: create the internal sheets on the fixture workbook
    Dim passSheet As Worksheet
    Dim formatterSheet As Worksheet
    Dim formulaSheet As Worksheet

    Set passSheet = TestHelpers.EnsureWorksheet("__pass", FixtureWorkbook)
    Set formatterSheet = TestHelpers.EnsureWorksheet("__formatter", FixtureWorkbook)
    Set formulaSheet = TestHelpers.EnsureWorksheet("__formula", FixtureWorkbook)

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: internal sheets should be VeryHidden
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(passSheet.Visible), "__pass should be VeryHidden."
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(formatterSheet.Visible), "__formatter should be VeryHidden."
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(formulaSheet.Visible), "__formula should be VeryHidden."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareHidesInternalSheets"
End Sub

'@TestMethod("DesignerPreparation")
Public Sub TestPrepareHidesTranslationSheets()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareHidesTranslationSheets"
    On Error GoTo Fail

    'Arrange: LinelistTranslation sheet
    Dim llTransSheet As Worksheet
    Set llTransSheet = TestHelpers.EnsureWorksheet("LinelistTranslation", FixtureWorkbook)

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: translation sheets should be Hidden (not VeryHidden)
    Assert.AreEqual CLng(xlSheetHidden), CLng(llTransSheet.Visible), "LinelistTranslation should be Hidden."
    Assert.AreEqual CLng(xlSheetHidden), CLng(TranslationSheet.Visible), "DesignerTranslation should be Hidden."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareHidesTranslationSheets"
End Sub

'@TestMethod("DesignerPreparation")
Public Sub TestPrepareCreatesWorkbookFlags()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareCreatesWorkbookFlags"
    On Error GoTo Fail

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: workbook-level HiddenNames should exist
    Dim wkbNames As IHiddenNames
    Set wkbNames = subject.HiddenStore

    Assert.AreEqual "Yes", wkbNames.ValueAsString("chkAlert"), "chkAlert should be Yes."
    Assert.AreEqual "Yes", wkbNames.ValueAsString("chkInstruct"), "chkInstruct should be Yes."
    Assert.IsTrue LenB(wkbNames.ValueAsString("RNG_LastOpenedDate")) > 0, "RNG_LastOpenedDate should be set."

    'Language flags should exist with empty defaults
    Assert.AreEqual vbNullString, wkbNames.ValueAsString("TAG_DES_LANG"), "TAG_DES_LANG should default to empty."
    Assert.AreEqual vbNullString, wkbNames.ValueAsString("RNG_LLLanguageCode"), "RNG_LLLanguageCode should default to empty."
    Assert.AreEqual vbNullString, wkbNames.ValueAsString("RNG_DictionaryLanguage"), "RNG_DictionaryLanguage should default to empty."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareCreatesWorkbookFlags"
End Sub

'@TestMethod("DesignerPreparation")
Public Sub TestPrepareCreatesGeoFlags()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareCreatesGeoFlags"
    On Error GoTo Fail

    'Arrange: create Geo sheet on the fixture workbook
    Dim geoSheet As Worksheet
    Set geoSheet = TestHelpers.EnsureWorksheet("Geo", FixtureWorkbook)

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: Geo worksheet-level HiddenNames should exist
    Dim geoStore As IHiddenNames
    Set geoStore = HiddenNames.Create(geoSheet)

    Assert.AreEqual vbNullString, geoStore.ValueAsString("RNG_GeoLangCode"), "RNG_GeoLangCode should default to empty."
    Assert.AreEqual vbNullString, geoStore.ValueAsString("RNG_GeoName"), "RNG_GeoName should default to empty."
    Assert.AreEqual vbNullString, geoStore.ValueAsString("RNG_MetaLang"), "RNG_MetaLang should default to empty."
    Assert.AreEqual "empty", geoStore.ValueAsString("RNG_GeoUpdated"), "RNG_GeoUpdated should default to empty."
    Assert.AreEqual vbNullString, geoStore.ValueAsString("RNG_PastingGeoCol"), "RNG_PastingGeoCol should default to empty."
    Assert.AreEqual vbNullString, geoStore.ValueAsString("RNG_FormLoaded"), "RNG_FormLoaded should default to empty."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareCreatesGeoFlags"
End Sub

'@TestMethod("DesignerPreparation")
Public Sub TestPrepareSkipsGeoWhenSheetMissing()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareSkipsGeoWhenSheetMissing"
    On Error GoTo Fail

    'Arrange: do NOT create a Geo sheet

    'Act: should not raise an error
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: workbook-level flags should still be created
    Assert.IsTrue subject.GetFlag("chkAlert"), "Preparation should succeed without Geo sheet."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSkipsGeoWhenSheetMissing"
End Sub


'@section Dropdown Tests
'===============================================================================
'@TestMethod("DesignerPreparation.Dropdowns")
Public Sub TestPrepareCreatesDropdownSheet()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareCreatesDropdownSheet"
    On Error GoTo Fail

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: __dropdowns sheet should exist and be VeryHidden
    Assert.IsTrue TestHelpers.WorksheetExists("__dropdowns", FixtureWorkbook), _
                  "__dropdowns worksheet should be created."
    Assert.AreEqual CLng(xlSheetVeryHidden), _
                    CLng(FixtureWorkbook.Worksheets("__dropdowns").Visible), _
                    "__dropdowns should be VeryHidden."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareCreatesDropdownSheet"
End Sub

'@TestMethod("DesignerPreparation.Dropdowns")
Public Sub TestPrepareRegistersAllDropdowns()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareRegistersAllDropdowns"
    On Error GoTo Fail

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: all 4 dropdowns should be registered
    Dim drop As IDropdownLists
    Set drop = subject.Dropdowns

    Assert.IsTrue drop.Exists("__setup_languages"), "Setup languages dropdown should exist."
    Assert.IsTrue drop.Exists("__interface_languages"), "Interface languages dropdown should exist."
    Assert.IsTrue drop.Exists("__epiweek_start"), "Epiweek start dropdown should exist."
    Assert.IsTrue drop.Exists("__design_values"), "Design values dropdown should exist."
    Assert.AreEqual 4&, drop.Length, "Exactly 4 dropdowns should be registered."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareRegistersAllDropdowns"
End Sub

'@TestMethod("DesignerPreparation.Dropdowns")
Public Sub TestInterfaceLanguagesContainsExpectedValues()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestInterfaceLanguagesContainsExpectedValues"
    On Error GoTo Fail

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: interface languages should contain the 5 expected values
    Dim values As BetterArray
    Set values = subject.Dropdowns.Values("__interface_languages")

    Assert.AreEqual 5&, values.Length, "Interface languages should have 5 entries."
    Assert.AreEqual "English", CStr(values.Item(values.LowerBound + 1)), _
                    "Second entry should be English."
    Exit Sub

Fail:
    ReportTestFailure "TestInterfaceLanguagesContainsExpectedValues"
End Sub

'@TestMethod("DesignerPreparation.Dropdowns")
Public Sub TestEpiweekStartContainsSevenDays()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestEpiweekStartContainsSevenDays"
    On Error GoTo Fail

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: epiweek start should contain 1 through 7
    Dim values As BetterArray
    Set values = subject.Dropdowns.Values("__epiweek_start")

    Assert.AreEqual 7&, values.Length, "Epiweek start should have 7 entries."
    Assert.AreEqual "1", CStr(values.Item(values.LowerBound)), "First entry should be 1."
    Assert.AreEqual "7", CStr(values.Item(values.UpperBound)), "Last entry should be 7."
    Exit Sub

Fail:
    ReportTestFailure "TestEpiweekStartContainsSevenDays"
End Sub

'@TestMethod("DesignerPreparation.Dropdowns")
Public Sub TestDesignValuesMatchesLLFormat()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestDesignValuesMatchesLLFormat"
    On Error GoTo Fail

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: design values should match LLFormat constants
    Dim values As BetterArray
    Set values = subject.Dropdowns.Values("__design_values")

    Assert.AreEqual 3&, values.Length, "Design values should have 3 entries."
    Assert.AreEqual "design 1", CStr(values.Item(values.LowerBound)), "First design should be design 1."
    Assert.AreEqual "design 2", CStr(values.Item(values.LowerBound + 1)), "Second design should be design 2."
    Assert.AreEqual "user defined", CStr(values.Item(values.LowerBound + 2)), "Third design should be user defined."
    Exit Sub

Fail:
    ReportTestFailure "TestDesignValuesMatchesLLFormat"
End Sub

'@TestMethod("DesignerPreparation.Dropdowns")
Public Sub TestDropdownsPropertyLazilyInitialises()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestDropdownsPropertyLazilyInitialises"
    On Error GoTo Fail

    'Arrange: create without calling Prepare
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)

    'Act: access Dropdowns property directly (lazy init)
    Dim drop As IDropdownLists
    Set drop = subject.Dropdowns

    'Assert: should have created the dropdown sheet and manager
    Assert.IsTrue Not drop Is Nothing, "Dropdowns property should return a valid manager."
    Assert.IsTrue TestHelpers.WorksheetExists("__dropdowns", FixtureWorkbook), _
                  "__dropdowns worksheet should be created lazily."
    Exit Sub

Fail:
    ReportTestFailure "TestDropdownsPropertyLazilyInitialises"
End Sub

'@TestMethod("DesignerPreparation.Dropdowns")
Public Sub TestDropdownUpdateReplacesValues()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestDropdownUpdateReplacesValues"
    On Error GoTo Fail

    'Arrange: create dropdown sheet and register initial __setup_languages
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Act: update the dropdown with new language values (mimics ExtractAndUpdateLanguages)
    Dim langValues As BetterArray
    Set langValues = New BetterArray
    langValues.LowerBound = 1
    langValues.Push "English", "Francais", "Espanol"

    Dim dropSheet As Worksheet
    Set dropSheet = FixtureWorkbook.Worksheets("__dropdowns")

    Dim drop As IDropdownLists
    Set drop = DropdownLists.Create(dropSheet)
    drop.Update langValues, "__setup_languages"

    'Assert: the dropdown should contain the updated values
    Dim result As BetterArray
    Set result = drop.Values("__setup_languages")

    Assert.AreEqual 3&, result.Length, "Updated dropdown should have 3 entries."
    Assert.AreEqual "English", CStr(result.Item(result.LowerBound)), _
                    "First language should be English."
    Assert.AreEqual "Francais", CStr(result.Item(result.LowerBound + 1)), _
                    "Second language should be Francais."
    Assert.AreEqual "Espanol", CStr(result.Item(result.LowerBound + 2)), _
                    "Third language should be Espanol."
    Exit Sub

Fail:
    ReportTestFailure "TestDropdownUpdateReplacesValues"
End Sub


'@section T_Multi Validation Tests
'===============================================================================
'@TestMethod("DesignerPreparation.MultiValidation")
Public Sub TestPrepareAppliesMultiValidations()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareAppliesMultiValidations"
    On Error GoTo Fail

    'Arrange: create GenerateMultiple sheet with T_Multi table
    Dim multiSheet As Worksheet
    Set multiSheet = TestHelpers.EnsureWorksheet("GenerateMultiple", FixtureWorkbook)
    CreateMultiTable multiSheet

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: validation should be applied to the 3 expected columns
    Dim lo As ListObject
    Set lo = multiSheet.ListObjects("T_Multi")

    Dim langCol As Range
    Set langCol = lo.ListColumns("language of the interface").DataBodyRange

    Dim epiCol As Range
    Set epiCol = lo.ListColumns("epiweek start").DataBodyRange

    Dim designCol As Range
    Set designCol = lo.ListColumns("design").DataBodyRange

    Assert.AreEqual CLng(xlValidateList), CLng(langCol.Cells(1).Validation.Type), _
                    "Language of the interface should have list validation."
    Assert.AreEqual CLng(xlValidateList), CLng(epiCol.Cells(1).Validation.Type), _
                    "Epiweek start should have list validation."
    Assert.AreEqual CLng(xlValidateList), CLng(designCol.Cells(1).Validation.Type), _
                    "Design should have list validation."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareAppliesMultiValidations"
End Sub

'@TestMethod("DesignerPreparation.MultiValidation")
Public Sub TestPrepareSkipsMultiWhenSheetMissing()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareSkipsMultiWhenSheetMissing"
    On Error GoTo Fail

    'Arrange: do NOT create GenerateMultiple sheet

    'Act: should not raise an error
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: dropdowns should still be created
    Assert.IsTrue subject.Dropdowns.Exists("__epiweek_start"), _
                  "Preparation should succeed without GenerateMultiple sheet."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSkipsMultiWhenSheetMissing"
End Sub


'@section Main Validation Tests
'===============================================================================
'@TestMethod("DesignerPreparation.MainValidation")
Public Sub TestPrepareAppliesMainValidations()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareAppliesMainValidations"
    On Error GoTo Fail

    'Arrange: create named ranges on the Main worksheet
    FixtureWorkbook.Names.Add Name:="RNG_LangSetup", RefersTo:=MainSheet.Range("H1")
    FixtureWorkbook.Names.Add Name:="RNG_LLForm", RefersTo:=MainSheet.Range("H2")
    FixtureWorkbook.Names.Add Name:="RNG_LLDesign", RefersTo:=MainSheet.Range("H3")

    'Act
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: all three ranges should have list validation
    Assert.AreEqual CLng(xlValidateList), CLng(MainSheet.Range("H1").Validation.Type), _
                    "RNG_LangSetup should have list validation."
    Assert.AreEqual CLng(xlValidateList), CLng(MainSheet.Range("H2").Validation.Type), _
                    "RNG_LLForm should have list validation."
    Assert.AreEqual CLng(xlValidateList), CLng(MainSheet.Range("H3").Validation.Type), _
                    "RNG_LLDesign should have list validation."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareAppliesMainValidations"
End Sub

'@TestMethod("DesignerPreparation.MainValidation")
Public Sub TestPrepareSkipsMainValidationsWhenRangesMissing()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareSkipsMainValidationsWhenRangesMissing"
    On Error GoTo Fail

    'Arrange: do NOT create named ranges on the Main sheet

    'Act: should not raise an error
    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    'Assert: preparation should still succeed
    Assert.IsTrue subject.Dropdowns.Exists("__setup_languages"), _
                  "Preparation should succeed without Main named ranges."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSkipsMainValidationsWhenRangesMissing"
End Sub


'@section Internal helpers
'===============================================================================

'@label:create-multi-table
'@sub-title Create a T_Multi ListObject with the expected headers
'@details
'Writes the T_Multi header row and one empty data row on the supplied
'worksheet, then converts the range to a ListObject named T_Multi.
'@param sh Worksheet. The worksheet to create the table on.
Private Sub CreateMultiTable(ByVal sh As Worksheet)
    Dim headers As Variant
    headers = Array("setups", "geobases", "output folders", "output files", _
                    "output file password", "output file debugging password", _
                    "language of the dictionary", "language of the interface", _
                    "epiweek start", "design", "result")

    Dim idx As Long
    For idx = LBound(headers) To UBound(headers)
        sh.Cells(1, idx - LBound(headers) + 1).Value = headers(idx)
    Next idx

    'Add one empty data row so DataBodyRange exists
    sh.Cells(2, 1).Value = vbNullString

    Dim dataRange As Range
    Set dataRange = sh.Range(sh.Cells(1, 1), sh.Cells(2, UBound(headers) - LBound(headers) + 1))

    Dim lo As ListObject
    Set lo = sh.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=dataRange, _
        XlListObjectHasHeaders:=xlYes)
    lo.Name = "T_Multi"
End Sub

Private Sub ReportTestFailure(ByVal context As String)
    Dim message As String

    If Assert Is Nothing Then Exit Sub

    message = context & " failed with error " & Err.Number & " (" & Err.Source & "): " & Err.Description
    Assert.LogFailure message
    Err.Clear
End Sub
