Attribute VB_Name = "TestDesignerRibbon"
Attribute VB_Description = "Unit tests for designer ribbon helpers"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates designer ribbon helpers for entry clearing, translation, persisted flags, dropdown creation, and T_Multi validation.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private EntrySheet As Worksheet
Private TranslationSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    TestHelpers.BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDesignerRibbon"
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
    Set EntrySheet = TestHelpers.EnsureWorksheet("Main", FixtureWorkbook)
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
    Set EntrySheet = Nothing
    Set FixtureWorkbook = Nothing

    TestHelpers.RestoreApp
End Sub


'@section DesignerEntry Tests
'===============================================================================
'@TestMethod("DesignerEntry")
Public Sub TestClearUsesEntryManager()
    CustomTestSetTitles Assert, "DesignerEntry", "TestClearUsesEntryManager"
    On Error GoTo Fail

    Dim subject As IDesignerEntry
    Dim stub As DesignerMainStub

    Set stub = New DesignerMainStub
    Set subject = DesignerEntry.Create(EntrySheet)
    subject.UseEntryManager stub

    subject.Clear

    Assert.IsTrue stub.ClearRequested, "ClearInputRanges should be invoked."
    Assert.IsTrue stub.ClearedWithValues, "Entry manager should clear values."
    Exit Sub

Fail:
    ReportTestFailure "TestClearUsesEntryManager"
End Sub

'@TestMethod("DesignerEntry")
Public Sub TestTranslateUpdatesLanguageCode()
    CustomTestSetTitles Assert, "DesignerEntry", "TestTranslateUpdatesLanguageCode"
    On Error GoTo Fail

    Dim subject As IDesignerEntry
    Dim translator As DesignerTranslationStub

    Set translator = New DesignerTranslationStub

    Set subject = DesignerEntry.Create(EntrySheet)
    subject.UseTranslator translator
    subject.Translate "ENG"

    Assert.IsTrue translator.TranslateRequested, "Translator should be invoked."
    Assert.AreEqual EntrySheet.Name, translator.TargetSheet.Name, "Translation target should be the entry sheet."
    Exit Sub

Fail:
    ReportTestFailure "TestTranslateUpdatesLanguageCode"
End Sub


'@section DesignerEntry AddInfo/ValueOf Tests
'===============================================================================
'@TestMethod("DesignerEntry.Info")
Public Sub TestAddInfoWritesToNamedRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestAddInfoWritesToNamedRange"
    On Error GoTo Fail

    'Arrange: create the setuppath named range on the entry sheet
    FixtureWorkbook.Names.Add Name:="RNG_PathDico", RefersTo:=EntrySheet.Range("B1")

    Dim subject As IDesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    subject.AddInfo "/path/to/setup.xlsb", "setuppath"

    'Assert
    Assert.AreEqual "/path/to/setup.xlsb", CStr(EntrySheet.Range("B1").value), _
                    "AddInfo should write the value to the named range."
    Assert.AreEqual CLng(vbWhite), CLng(EntrySheet.Range("B1").Interior.Color), _
                    "AddInfo should set the cell background to white."
    Exit Sub

Fail:
    ReportTestFailure "TestAddInfoWritesToNamedRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestAddInfoEditionRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestAddInfoEditionRange"
    On Error GoTo Fail

    'Arrange
    FixtureWorkbook.Names.Add Name:="RNG_Edition", RefersTo:=EntrySheet.Range("C1")

    Dim subject As IDesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act: write a status message to the edition range
    subject.AddInfo "File loaded", "edition"

    'Assert
    Assert.AreEqual "File loaded", CStr(EntrySheet.Range("C1").value), _
                    "Edition range should contain the message."
    Exit Sub

Fail:
    ReportTestFailure "TestAddInfoEditionRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestValueOfReadsFromNamedRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValueOfReadsFromNamedRange"
    On Error GoTo Fail

    'Arrange: create the geopath named range and write a value
    FixtureWorkbook.Names.Add Name:="RNG_PathGeo", RefersTo:=EntrySheet.Range("D1")
    EntrySheet.Range("D1").value = "/path/to/geo.xlsx"

    Dim subject As IDesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    Dim result As String
    result = subject.ValueOf("geopath")

    'Assert
    Assert.AreEqual "/path/to/geo.xlsx", result, _
                    "ValueOf should return the value from the named range."
    Exit Sub

Fail:
    ReportTestFailure "TestValueOfReadsFromNamedRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestValueOfReturnsEmptyForUnknownRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValueOfReturnsEmptyForUnknownRange"
    On Error GoTo Fail

    Dim subject As IDesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act: request an unknown info name
    Dim result As String
    result = subject.ValueOf("nonexistent")

    'Assert
    Assert.AreEqual vbNullString, result, _
                    "ValueOf should return empty for unknown info names."
    Exit Sub

Fail:
    ReportTestFailure "TestValueOfReturnsEmptyForUnknownRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestTranslateMessageReturnsTranslatedText()
    CustomTestSetTitles Assert, "DesignerEntry", "TestTranslateMessageReturnsTranslatedText"
    On Error GoTo Fail

    'Arrange
    Dim translator As DesignerTranslationStub
    Set translator = New DesignerTranslationStub
    translator.SetMessage "MSG_ChemFich", "File path loaded"

    Dim subject As IDesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)
    subject.UseTranslator translator

    'Act
    Dim result As String
    result = subject.TranslateMessage("MSG_ChemFich")

    'Assert
    Assert.AreEqual "File path loaded", result, _
                    "TranslateMessage should return the translated text."
    Exit Sub

Fail:
    ReportTestFailure "TestTranslateMessageReturnsTranslatedText"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestTranslateMessageFallsBackToRawCode()
    CustomTestSetTitles Assert, "DesignerEntry", "TestTranslateMessageFallsBackToRawCode"
    On Error GoTo Fail

    'Arrange: translator with no messages registered
    Dim translator As DesignerTranslationStub
    Set translator = New DesignerTranslationStub

    Dim subject As IDesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)
    subject.UseTranslator translator

    'Act
    Dim result As String
    result = subject.TranslateMessage("MSG_Unknown")

    'Assert: stub returns raw msgCode for unknown codes
    Assert.AreEqual "MSG_Unknown", result, _
                    "TranslateMessage should fall back to the raw message code."
    Exit Sub

Fail:
    ReportTestFailure "TestTranslateMessageFallsBackToRawCode"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestAddInfoSilentlySkipsMissingRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestAddInfoSilentlySkipsMissingRange"
    On Error GoTo Fail

    'Arrange: do NOT create the RNG_PathDico named range

    Dim subject As IDesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act: should not raise an error
    subject.AddInfo "/some/path", "setuppath"

    'Assert: no error was raised (test completes)
    Assert.IsTrue True, "AddInfo should silently skip when the named range is missing."
    Exit Sub

Fail:
    ReportTestFailure "TestAddInfoSilentlySkipsMissingRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestValueOfLLDirRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValueOfLLDirRange"
    On Error GoTo Fail

    'Arrange: create the lldir named range and write a value
    FixtureWorkbook.Names.Add Name:="RNG_LLDir", RefersTo:=EntrySheet.Range("E1")
    EntrySheet.Range("E1").value = "/output/folder"

    Dim subject As IDesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    Dim result As String
    result = subject.ValueOf("lldir")

    'Assert
    Assert.AreEqual "/output/folder", result, _
                    "ValueOf lldir should return the linelist directory path."
    Exit Sub

Fail:
    ReportTestFailure "TestValueOfLLDirRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestValueOfLLNameRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValueOfLLNameRange"
    On Error GoTo Fail

    'Arrange
    FixtureWorkbook.Names.Add Name:="RNG_LLName", RefersTo:=EntrySheet.Range("F1")
    EntrySheet.Range("F1").value = "my_linelist"

    Dim subject As IDesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    Dim result As String
    result = subject.ValueOf("llname")

    'Assert
    Assert.AreEqual "my_linelist", result, _
                    "ValueOf llname should return the linelist name."
    Exit Sub

Fail:
    ReportTestFailure "TestValueOfLLNameRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestValueOfTempPathRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValueOfTempPathRange"
    On Error GoTo Fail

    'Arrange
    FixtureWorkbook.Names.Add Name:="RNG_LLTemp", RefersTo:=EntrySheet.Range("G1")
    EntrySheet.Range("G1").value = "/path/to/template.xlsb"

    Dim subject As IDesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    Dim result As String
    result = subject.ValueOf("temppath")

    'Assert
    Assert.AreEqual "/path/to/template.xlsb", result, _
                    "ValueOf temppath should return the template file path."
    Exit Sub

Fail:
    ReportTestFailure "TestValueOfTempPathRange"
End Sub

'@TestMethod("DesignerEntry.Dropdowns")
Public Sub TestDropdownUpdateReplacesValues()
    CustomTestSetTitles Assert, "DesignerEntry", "TestDropdownUpdateReplacesValues"
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
