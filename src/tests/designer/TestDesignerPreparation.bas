Attribute VB_Name = "TestDesignerPreparation"
Attribute VB_Description = "Unit tests for DesignerPreparation class"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates DesignerPreparation for persisted flags, the stored ribbon file of the multi generation, sheet hiding, dropdown creation, T_Multi and Main validation.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private MainSheet As Worksheet
Private TranslationSheet As Worksheet
Private TranslationsSource As Workbook

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
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
    RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    BusyApp

    Set FixtureWorkbook = NewWorkbook
    Set MainSheet = EnsureWorksheet("Main", FixtureWorkbook)
    Set TranslationSheet = EnsureWorksheet("DesignerTranslation", FixtureWorkbook)

    'The translations source is a fixture workbook handed to Prepare, so the
    'import runs without the file picker.
    Set TranslationsSource = NewWorkbook
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        DeleteWorkbook FixtureWorkbook
        DeleteWorkbook TranslationsSource
    On Error GoTo 0

    Set TranslationsSource = Nothing
    Set TranslationSheet = Nothing
    Set MainSheet = Nothing
    Set FixtureWorkbook = Nothing

    RestoreApp
End Sub


'@section DesignerPreparation Tests
'===============================================================================
'@TestMethod("DesignerPreparation")
Public Sub TestPrepareSeedsFlags()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareSeedsFlags"
    On Error GoTo Fail

    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    Assert.IsTrue subject.GetFlag("chkAlert"), "Alert flag should default to on."
    Assert.IsTrue subject.GetFlag("chkInstruct"), "Instruction flag should default to on."

    subject.SetFlag "chkAlert", False
    Assert.IsFalse subject.GetFlag("chkAlert"), "Alert flag should persist changes."
    Assert.AreEqual "No", subject.HiddenStore.ValueAsString("chkAlert"), "Hidden name should store No for disabled flags."

    'The build path switch is off by default and takes a write on a
    'workbook whose names predate it
    Assert.IsFalse subject.GetFlag("chkBuildInPlace"), "Build-in-place flag should default to off."
    subject.SetFlag "chkBuildInPlace", True
    Assert.IsTrue subject.GetFlag("chkBuildInPlace"), "Build-in-place flag should persist changes."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSeedsFlags"
End Sub

'@TestMethod("DesignerPreparation.Silence")
Public Sub TestPrepareSeedsTheSilenceSwitch()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareSeedsTheSilenceSwitch"
    On Error GoTo Fail

    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    Assert.AreEqual Messenger.SwitchOffValue, _
        subject.HiddenStore.ValueAsString(Messenger.SwitchName), _
        "The silence switch should be seeded No."
    Assert.IsFalse Messenger.ReadStoredSwitch(FixtureWorkbook), _
        "A designer seeded No should read as not silent."

    'Someone switching the designer to silent has to survive another Prepare,
    'because EnsureName leaves a name that is already there alone.
    subject.HiddenStore.SetValue Messenger.SwitchName, Messenger.SwitchOnValue
    subject.Prepare Nothing, TranslationsSource
    Assert.IsTrue Messenger.ReadStoredSwitch(FixtureWorkbook), _
        "A second Prepare should keep the value the switch holds."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSeedsTheSilenceSwitch"
End Sub

'@TestMethod("DesignerPreparation.Silence")
Public Sub TestSwitchArrivesOnADesignerThatPredatesIt()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestSwitchArrivesOnADesignerThatPredatesIt"
    On Error GoTo Fail

    'A designer built before the switch existed holds no such name, and it has
    'to behave exactly as it always has until something ensures the flags.
    Assert.IsFalse Messenger.ReadStoredSwitch(FixtureWorkbook), _
        "A workbook with no switch name should read as not silent."

    'EnsureDefaultFlags runs on the first read of any flag, so touching the
    'alert flag is what brings the switch in.
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    Assert.IsTrue subject.GetFlag("chkAlert", True), "Reading a flag should seed the defaults."
    Assert.AreEqual Messenger.SwitchOffValue, _
        subject.HiddenStore.ValueAsString(Messenger.SwitchName), _
        "Reading a flag should bring the silence switch in, seeded No."
    Exit Sub

Fail:
    ReportTestFailure "TestSwitchArrivesOnADesignerThatPredatesIt"
End Sub

'@TestMethod("DesignerPreparation")
Public Sub TestPrepareHidesInternalSheets()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareHidesInternalSheets"
    On Error GoTo Fail

    'Arrange: create the internal sheets on the fixture workbook. The
    '__check sheet starts visible, the way a finished generation run
    'leaves it, and Prepare is expected to leave it that way.
    Dim passSheet As Worksheet
    Dim formatterSheet As Worksheet
    Dim formulaSheet As Worksheet
    Dim checkSheet As Worksheet
    Dim geoSheet As Worksheet

    Set passSheet = EnsureWorksheet("__pass", FixtureWorkbook)
    Set formatterSheet = EnsureWorksheet("__formatter", FixtureWorkbook)
    Set formulaSheet = EnsureWorksheet("__formula", FixtureWorkbook)
    Set checkSheet = EnsureWorksheet("__check", FixtureWorkbook)
    Set geoSheet = EnsureWorksheet("Geo", FixtureWorkbook)

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: the sheets nobody edits by hand go away, and the two the
    'designer works with stay on screen. __formatter carries the formats a
    'user picks and __check carries the generation report, so Prepare leaves
    'both visible. That is what "Unhide some designer worksheet at
    'preparation" set, and it is the behaviour this test now holds.
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(passSheet.Visible), "__pass should be VeryHidden."
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(formulaSheet.Visible), "__formula should be VeryHidden."
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(geoSheet.Visible), _
                    "Geo should be VeryHidden: the geobase is imported on the linelist, never here."
    Assert.AreEqual CLng(xlSheetVisible), CLng(formatterSheet.Visible), "__formatter should stay visible."
    Assert.AreEqual CLng(xlSheetVisible), CLng(checkSheet.Visible), "__check should stay visible."
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
    Set llTransSheet = EnsureWorksheet("LinelistTranslation", FixtureWorkbook)

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

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
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: workbook-level HiddenNames should exist
    Dim wkbNames As HiddenNames
    Set wkbNames = subject.HiddenStore

    Assert.AreEqual "Yes", wkbNames.ValueAsString("chkAlert"), "chkAlert should be Yes."
    Assert.AreEqual "Yes", wkbNames.ValueAsString("chkInstruct"), "chkInstruct should be Yes."
    Assert.IsTrue LenB(wkbNames.ValueAsString("RNG_LastOpenedDate")) > 0, "RNG_LastOpenedDate should be set."

    'Language flags should exist with empty defaults
    Assert.AreEqual vbNullString, wkbNames.ValueAsString("RNG_LLLanguageCode"), "RNG_LLLanguageCode should default to empty."
    Assert.AreEqual vbNullString, wkbNames.ValueAsString("RNG_DictionaryLanguage"), "RNG_DictionaryLanguage should default to empty."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareCreatesWorkbookFlags"
End Sub

'@TestMethod("DesignerPreparation.Formatter")
Public Sub TestFormatterFlagStartsCleared()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestFormatterFlagStartsCleared"
    On Error GoTo Fail

    'A designer that never met the styles import button reads the setup's
    'formatter, so the flag answers False with no hidden name in the workbook
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)

    Assert.IsFalse subject.FormatterImported, _
                   "A workbook with no formatter flag should read False."
    Exit Sub

Fail:
    ReportTestFailure "TestFormatterFlagStartsCleared"
End Sub

'@TestMethod("DesignerPreparation.Formatter")
Public Sub TestFormatterFlagIsSetAndCleared()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestFormatterFlagIsSetAndCleared"
    On Error GoTo Fail

    'Arrange
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)

    'Act: the styles import sets the flag
    subject.FormatterImported = True

    'Assert
    Assert.IsTrue subject.FormatterImported, _
                  "The flag should read True after a styles import."
    Assert.AreEqual "Yes", subject.HiddenStore.ValueAsString("TAG_FORMATTER_IMPORTED"), _
                    "The hidden name should store Yes."

    'Act: loading a setup file clears it
    subject.FormatterImported = False

    'Assert
    Assert.IsFalse subject.FormatterImported, _
                   "The flag should read False once a setup file is loaded."
    Assert.AreEqual "No", subject.HiddenStore.ValueAsString("TAG_FORMATTER_IMPORTED"), _
                    "The hidden name should store No."
    Exit Sub

Fail:
    ReportTestFailure "TestFormatterFlagIsSetAndCleared"
End Sub


'@TestMethod("DesignerPreparation.RibbonTemplate")
Public Sub TestRibbonTemplatePathIsKeptAndCleared()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestRibbonTemplatePathIsKeptAndCleared"
    On Error GoTo Fail

    'Arrange: a designer that was never given a ribbon file
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)

    Assert.AreEqual vbNullString, subject.RibbonTemplatePath, _
                    "A workbook with no ribbon file should read empty."

    'Act: the multi generation stores the file the user picked
    subject.RibbonTemplatePath = "C:\templates\ribbon.xlsb"

    'Assert: a second reader of the same workbook finds it
    Dim reader As DesignerPreparation
    Set reader = DesignerPreparation.Create(FixtureWorkbook)
    Assert.AreEqual "C:\templates\ribbon.xlsb", reader.RibbonTemplatePath, _
                    "The ribbon file should outlive the object that stored it."

    'Act: the button clears it so the linelists carry buttons
    subject.RibbonTemplatePath = vbNullString

    'Assert
    Assert.AreEqual vbNullString, subject.RibbonTemplatePath, _
                    "A cleared ribbon file should read empty."
    Assert.IsFalse subject.HiddenStore.HasName("TAG_RIBBON_TEMPLATE"), _
                   "Clearing the ribbon file should take its hidden name off the workbook."
    Exit Sub

Fail:
    ReportTestFailure "TestRibbonTemplatePathIsKeptAndCleared"
End Sub

'@TestMethod("DesignerPreparation")
Public Sub TestPrepareCreatesGeoFlags()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareCreatesGeoFlags"
    On Error GoTo Fail

    'Arrange: create Geo sheet on the fixture workbook
    Dim geoSheet As Worksheet
    Set geoSheet = EnsureWorksheet("Geo", FixtureWorkbook)

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: Geo worksheet-level HiddenNames should exist
    Dim geoStore As HiddenNames
    Set geoStore = HiddenNames.Create(geoSheet)

    Assert.AreEqual vbNullString, geoStore.ValueAsString("RNG_GeoLangCode"), "RNG_GeoLangCode should default to empty."
    Assert.AreEqual vbNullString, geoStore.ValueAsString("RNG_GeoName"), "RNG_GeoName should default to empty."
    Assert.AreEqual vbNullString, geoStore.ValueAsString("RNG_MetaLang"), "RNG_MetaLang should default to empty."
    Assert.AreEqual "empty", geoStore.ValueAsString("RNG_GeoUpdated"), "RNG_GeoUpdated should default to empty."
    Assert.AreEqual vbNullString, geoStore.ValueAsString("RNG_FormLoaded"), "RNG_FormLoaded should default to empty."

    'RNG_PastingGeoCol stays a cell-based named range. It is a paste anchor.
    Dim pastingRng As Range
    On Error Resume Next
    Set pastingRng = geoSheet.Range("RNG_PastingGeoCol")
    On Error GoTo Fail
    Assert.IsNotNothing pastingRng, "RNG_PastingGeoCol should exist as a cell-based range."
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
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: workbook-level flags should still be created
    Assert.IsTrue subject.GetFlag("chkAlert"), "Preparation should succeed without Geo sheet."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSkipsGeoWhenSheetMissing"
End Sub


'@section Version Tests
'===============================================================================
'The version is owner hand work on the Dev worksheet, and preparation carries
'it to the workbook-level name LLGeo reads for the metadata sheet.

'@TestMethod("DesignerPreparation.Version")
Public Sub TestPrepareCopiesDevVersion()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareCopiesDevVersion"
    On Error GoTo Fail

    'Arrange: a Dev sheet carrying the version as a value name
    Dim devSheet As Worksheet
    Set devSheet = EnsureWorksheet("Dev", FixtureWorkbook)
    devSheet.Names.Add Name:="RNG_Version", RefersTo:="=""1.4.2""", Visible:=False

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert
    Assert.AreEqual "1.4.2", subject.HiddenStore.ValueAsString("RNG_DesignerVersion"), _
                    "The workbook name should carry the version of the Dev sheet."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareCopiesDevVersion"
End Sub

'@TestMethod("DesignerPreparation.Version")
Public Sub TestPrepareReadsVersionWrittenInACell()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareReadsVersionWrittenInACell"
    On Error GoTo Fail

    'Arrange: a Dev sheet whose RNG_Version points at a cell instead
    Dim devSheet As Worksheet
    Set devSheet = EnsureWorksheet("Dev", FixtureWorkbook)
    devSheet.Range("B2").Value = "2.0.1"
    devSheet.Names.Add Name:="RNG_Version", RefersTo:=devSheet.Range("B2"), Visible:=False

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: the cell value travels, not the address it sits at
    Assert.AreEqual "2.0.1", subject.HiddenStore.ValueAsString("RNG_DesignerVersion"), _
                    "A version typed in a cell should reach the workbook name."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareReadsVersionWrittenInACell"
End Sub

'@TestMethod("DesignerPreparation.Version")
Public Sub TestPrepareSkipsVersionWhenDevSheetMissing()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareSkipsVersionWhenDevSheetMissing"
    On Error GoTo Fail

    'Arrange: do NOT create a Dev sheet

    'Act: preparation should run through
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: no version name is written, which is what LLGeo reads as
    '"(not found)"
    Assert.IsFalse subject.HiddenStore.HasName("RNG_DesignerVersion"), _
                   "A designer with no Dev sheet should carry no version name."
    Assert.IsTrue subject.GetFlag("chkAlert"), _
                  "Preparation should succeed without a Dev sheet."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSkipsVersionWhenDevSheetMissing"
End Sub

'@TestMethod("DesignerPreparation.Version")
Public Sub TestPrepareSkipsVersionWhenNameMissing()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareSkipsVersionWhenNameMissing"
    On Error GoTo Fail

    'Arrange: a Dev sheet with no RNG_Version on it
    Dim devSheet As Worksheet
    Set devSheet = EnsureWorksheet("Dev", FixtureWorkbook)

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert
    Assert.IsFalse subject.HiddenStore.HasName("RNG_DesignerVersion"), _
                   "A Dev sheet with no RNG_Version should leave the workbook name alone."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSkipsVersionWhenNameMissing"
End Sub


'@section Dropdown Tests
'===============================================================================
'@TestMethod("DesignerPreparation.Dropdowns")
Public Sub TestPrepareCreatesDropdownSheet()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareCreatesDropdownSheet"
    On Error GoTo Fail

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: __dropdowns sheet should exist and be VeryHidden
    Assert.IsTrue WorksheetExists("__dropdowns", FixtureWorkbook), _
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
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: all 4 dropdowns should be registered
    Dim drop As DropdownLists
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
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: interface languages should contain the 5 expected values
    Dim values As BetterArray
    Set values = subject.Dropdowns.Values("__interface_languages")

    Assert.AreEqual 5&, values.Length, "Interface languages should have 5 entries."
    Assert.AreEqual "ENG-English", CStr(values.Item(values.LowerBound + 1)), _
                    "Second entry should be ENG-English."
    Exit Sub

Fail:
    ReportTestFailure "TestInterfaceLanguagesContainsExpectedValues"
End Sub

'@TestMethod("DesignerPreparation.Dropdowns")
Public Sub TestEpiweekStartContainsSevenDays()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestEpiweekStartContainsSevenDays"
    On Error GoTo Fail

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: epiweek start should contain 1 through 6 and then 0
    Dim values As BetterArray
    Set values = subject.Dropdowns.Values("__epiweek_start")

    Assert.AreEqual 7&, values.Length, "Epiweek start should have 7 entries."
    Assert.AreEqual "1", CStr(values.Item(values.LowerBound)), "First entry should be 1."
    Assert.AreEqual "0", CStr(values.Item(values.UpperBound)), "Last entry should be 0."
    Exit Sub

Fail:
    ReportTestFailure "TestEpiweekStartContainsSevenDays"
End Sub

'@TestMethod("DesignerPreparation.Dropdowns")
Public Sub TestDesignValuesMatchesLLFormat()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestDesignValuesMatchesLLFormat"
    On Error GoTo Fail

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

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
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)

    'Act: access Dropdowns property directly (lazy init)
    Dim drop As DropdownLists
    Set drop = subject.Dropdowns

    'Assert: should have created the dropdown sheet and manager
    Assert.IsTrue Not drop Is Nothing, "Dropdowns property should return a valid manager."
    Assert.IsTrue WorksheetExists("__dropdowns", FixtureWorkbook), _
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
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Act: update the dropdown with new language values (mimics ExtractAndUpdateLanguages)
    Dim langValues As BetterArray
    Set langValues = New BetterArray
    langValues.LowerBound = 1
    langValues.Push "English", "Francais", "Espanol"

    Dim dropSheet As Worksheet
    Set dropSheet = FixtureWorkbook.Worksheets("__dropdowns")

    Dim drop As DropdownLists
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
    Set multiSheet = EnsureWorksheet("GenerateMultiple", FixtureWorkbook)
    CreateMultiTable multiSheet

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

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
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

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
    FixtureWorkbook.Names.Add Name:="RNG_DesignLL", RefersTo:=MainSheet.Range("H3")

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: all three ranges should have list validation
    Assert.AreEqual CLng(xlValidateList), CLng(MainSheet.Range("H1").Validation.Type), _
                    "RNG_LangSetup should have list validation."
    Assert.AreEqual CLng(xlValidateList), CLng(MainSheet.Range("H2").Validation.Type), _
                    "RNG_LLForm should have list validation."
    Assert.AreEqual CLng(xlValidateList), CLng(MainSheet.Range("H3").Validation.Type), _
                    "RNG_DesignLL should have list validation."
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
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: preparation should still succeed
    Assert.IsTrue subject.Dropdowns.Exists("__setup_languages"), _
                  "Preparation should succeed without Main named ranges."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSkipsMainValidationsWhenRangesMissing"
End Sub


'@section Translations Import Tests
'===============================================================================
'@TestMethod("DesignerPreparation.Import")
Public Sub TestPrepareImportsTranslationTables()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareImportsTranslationTables"
    On Error GoTo Fail

    'Arrange: the same table on both sides, with the values on the source
    AddTradMsgTable TranslationSheet, "placeholder", vbNullString
    AddTradMsgTable EnsureWorksheet("DesignerTranslation", TranslationsSource), _
                    "MSG_Test", "Hello"

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: the source values should land in the fixture table
    Dim lo As ListObject
    Set lo = TranslationSheet.ListObjects("T_tradMsg")

    Assert.AreEqual "MSG_Test", CStr(lo.DataBodyRange.Cells(1, 1).Value), _
                    "The import should bring the source tag."
    Assert.AreEqual "Hello", CStr(lo.DataBodyRange.Cells(1, 2).Value), _
                    "The import should bring the source value."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareImportsTranslationTables"
End Sub

'@TestMethod("DesignerPreparation.Import")
Public Sub TestPrepareImportsDropdownTranslations()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareImportsDropdownTranslations"
    On Error GoTo Fail

    'Arrange: T_tradDrop on both sides. DesignerTranslation refuses to build
    'without that table and the import used to walk past it
    Dim sourceSheet As Worksheet
    Set sourceSheet = EnsureWorksheet("DesignerTranslation", TranslationsSource)

    AddTradTable TranslationSheet, "T_tradDrop", 1, "placeholder", vbNullString
    AddTradTable sourceSheet, "T_tradDrop", 1, "DROP_Test", "Choice"

    'Act
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: the source values should land in the fixture table
    Dim lo As ListObject
    Set lo = TranslationSheet.ListObjects("T_tradDrop")

    Assert.AreEqual "DROP_Test", CStr(lo.DataBodyRange.Cells(1, 1).Value), _
                    "The import should bring the source dropdown tag."
    Assert.AreEqual "Choice", CStr(lo.DataBodyRange.Cells(1, 2).Value), _
                    "The import should bring the source dropdown value."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareImportsDropdownTranslations"
End Sub

'@TestMethod("DesignerPreparation.Import")
Public Sub TestPrepareSkipsImportWhenSourceSheetMissing()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareSkipsImportWhenSourceSheetMissing"
    On Error GoTo Fail

    'Arrange: the fixture has the table, the source workbook has no
    'translation sheet at all
    AddTradMsgTable TranslationSheet, "MSG_Kept", "Still here"

    'Act: should run through with no raise
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing, TranslationsSource

    'Assert: the fixture table should keep its rows
    Dim lo As ListObject
    Set lo = TranslationSheet.ListObjects("T_tradMsg")

    Assert.AreEqual "MSG_Kept", CStr(lo.DataBodyRange.Cells(1, 1).Value), _
                    "A source with no translation sheet should leave the target alone."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSkipsImportWhenSourceSheetMissing"
End Sub


'@section Sealing Tests
'===============================================================================
'@TestMethod("DesignerPreparation.Sealing")
Public Sub TestConfigureIsSetOnce()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestConfigureIsSetOnce"
    On Error GoTo Fail

    'Arrange: Create configures and seals the instance
    Dim subject As DesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)

    'Act: a second Configure should raise
    Dim raisedNumber As Long
    On Error Resume Next
    subject.Configure FixtureWorkbook
    raisedNumber = Err.Number
    On Error GoTo Fail

    'Assert
    Assert.IsTrue raisedNumber <> 0, "Configure should raise on a sealed instance."
    Assert.IsNotNothing subject.HostWorkbook, "The first binding should survive the refused call."
    Exit Sub

Fail:
    ReportTestFailure "TestConfigureIsSetOnce"
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

'@label:add-trad-msg-table
'@sub-title Create a two-column T_tradMsg ListObject with one data row
'@param sh Worksheet. The worksheet to create the table on.
'@param tagValue String. The value of the tag cell.
'@param engValue String. The value of the ENG cell.
Private Sub AddTradMsgTable(ByVal sh As Worksheet, ByVal tagValue As String, _
                            ByVal engValue As String)
    AddTradTable sh, "T_tradMsg", 1, tagValue, engValue
End Sub

'@label:add-trad-table
'@sub-title Create a two-column translation ListObject with one data row
'@details
'Writes tag/ENG headers and one data row from the given column of the
'supplied worksheet, then converts the range to a ListObject of the given
'name, the shape ImportTranslations looks for.
'@param sh Worksheet. The worksheet to create the table on.
'@param tableName String. The ListObject name.
'@param firstColumn Long. The column the two-column table starts on.
'@param tagValue String. The value of the tag cell.
'@param engValue String. The value of the ENG cell.
Private Sub AddTradTable(ByVal sh As Worksheet, ByVal tableName As String, _
                         ByVal firstColumn As Long, ByVal tagValue As String, _
                         ByVal engValue As String)
    sh.Cells(1, firstColumn).Value = "tag"
    sh.Cells(1, firstColumn + 1).Value = "ENG"
    sh.Cells(2, firstColumn).Value = tagValue
    sh.Cells(2, firstColumn + 1).Value = engValue

    Dim lo As ListObject
    Set lo = sh.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=sh.Range(sh.Cells(1, firstColumn), sh.Cells(2, firstColumn + 1)), _
        XlListObjectHasHeaders:=xlYes)
    lo.Name = tableName
End Sub

Private Sub ReportTestFailure(ByVal context As String)
    Dim message As String

    If Assert Is Nothing Then Exit Sub

    message = context & " failed with error " & Err.Number & " (" & Err.Source & "): " & Err.Description
    Assert.LogFailure message
    Err.Clear
End Sub
