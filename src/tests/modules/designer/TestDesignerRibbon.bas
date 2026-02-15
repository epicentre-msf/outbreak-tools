Attribute VB_Name = "TestDesignerRibbon"
Attribute VB_Description = "Unit tests for designer ribbon helpers"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates designer ribbon helpers for entry clearing, translation, and persisted flags.")
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
    translator.SetLanguageRange TranslationSheet.Range("A1")

    Set subject = DesignerEntry.Create(EntrySheet)
    subject.UseTranslator translator
    subject.Translate "ENG"

    Assert.AreEqual "ENG", TranslationSheet.Range("A1").Value, "Language code should be written to translation sheet."
    Assert.IsTrue translator.TranslateRequested, "Translator should be invoked."
    Assert.AreEqual EntrySheet.Name, translator.TargetSheet.Name, "Translation target should be the entry sheet."
    Exit Sub

Fail:
    ReportTestFailure "TestTranslateUpdatesLanguageCode"
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


'@section Internal helpers
'===============================================================================

Private Sub ReportTestFailure(ByVal context As String)
    Dim message As String

    If Assert Is Nothing Then Exit Sub

    message = context & " failed with error " & Err.Number & " (" & Err.Source & "): " & Err.Description
    Assert.LogFailure message
    Err.Clear
End Sub
