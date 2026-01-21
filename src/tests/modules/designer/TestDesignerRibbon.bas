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


'@section Tests
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

'@TestMethod("DesignerPreparation")
Public Sub TestPrepareSeedsFlags()
    CustomTestSetTitles Assert, "DesignerPreparation", "TestPrepareSeedsFlags"
    On Error GoTo Fail

    Dim subject As IDesignerPreparation
    Set subject = DesignerPreparation.Create(FixtureWorkbook)
    subject.Prepare Nothing

    Assert.IsFalse subject.GetFlag("chkAlert"), "Alert flag should default to off."
    Assert.IsFalse subject.GetFlag("chkInstruct"), "Instruction flag should default to off."

    subject.SetFlag "chkAlert", True
    Assert.IsTrue subject.GetFlag("chkAlert"), "Alert flag should persist changes."
    Assert.AreEqual "Yes", subject.HiddenStore.ValueAsString("chkAlert"), "Hidden name should store Yes for enabled flags."
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareSeedsFlags"
End Sub
