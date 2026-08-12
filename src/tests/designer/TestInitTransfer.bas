Attribute VB_Name = "TestInitTransfer"
Attribute VB_Description = "Unit tests for the InitTransfer module"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates InitTransfer: the content-decided setup anchor rows, the language sync onto the designer hidden names, the design name read, the formatter export of both branches and the missing-translations report entry.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private SourceWorkbook As Workbook

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'The hidden names SyncDesignerLanguageNames writes on the designer workbook
Private Const TAG_DICT_LANG As String = "RNG_DictionaryLanguage"
Private Const TAG_LL_LANG_CODE As String = "RNG_LLLanguageCode"

'The first header of the setup dictionary table, the content probe anchor
Private Const DICT_FIRST_HEADER As String = "variable name"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestInitTransfer"
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
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        DeleteWorkbook FixtureWorkbook
        If Not SourceWorkbook Is Nothing Then DeleteWorkbook SourceWorkbook
    On Error GoTo 0

    Set SourceWorkbook = Nothing
    Set FixtureWorkbook = Nothing

    'A mid-test workbook close can hand the screen to another workbook
    ThisWorkbook.Activate

    RestoreApp
End Sub


'@section Anchor Row Tests
'===============================================================================
'@TestMethod("InitTransfer.AnchorRows")
Public Sub TestSetupStartRowReadsHeaderAtRowOne()
    CustomTestSetTitles Assert, "InitTransfer", "TestSetupStartRowReadsHeaderAtRowOne"
    On Error GoTo Fail

    'Arrange: the header sits at row 1, the exported .xlsx layout
    Dim sh As Worksheet
    Set sh = EnsureWorksheet("Dictionary", FixtureWorkbook)
    sh.Cells(1, 1).Value = DICT_FIRST_HEADER

    'Act and assert
    Assert.AreEqual CLng(1), InitTransfer.SetupStartRow(sh, 5, 1, DICT_FIRST_HEADER), _
                    "A header at row 1 should anchor the table at row 1."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestSetupStartRowReadsHeaderAtRowOne", Err.Number, Err.Description
End Sub

'@TestMethod("InitTransfer.AnchorRows")
Public Sub TestSetupStartRowFindsHeaderAtOtherRow()
    CustomTestSetTitles Assert, "InitTransfer", "TestSetupStartRowFindsHeaderAtOtherRow"
    On Error GoTo Fail

    'Arrange: an unsaved fixture workbook hints at row 1, and the header
    'really sits at row 5, the .xlsb layout. The content probe has to win
    'over the hint.
    Dim sh As Worksheet
    Set sh = EnsureWorksheet("Dictionary", FixtureWorkbook)
    sh.Cells(1, 1).Value = "some metadata"
    sh.Cells(5, 1).Value = DICT_FIRST_HEADER

    'Act and assert
    Assert.AreEqual CLng(5), InitTransfer.SetupStartRow(sh, 5, 1, DICT_FIRST_HEADER), _
                    "The row carrying the header should win over the extension hint."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestSetupStartRowFindsHeaderAtOtherRow", Err.Number, Err.Description
End Sub

'@TestMethod("InitTransfer.AnchorRows")
Public Sub TestSetupStartRowMatchesHeaderCaseInsensitive()
    CustomTestSetTitles Assert, "InitTransfer", "TestSetupStartRowMatchesHeaderCaseInsensitive"
    On Error GoTo Fail

    'Arrange: the header carries capitals and padding
    Dim sh As Worksheet
    Set sh = EnsureWorksheet("Dictionary", FixtureWorkbook)
    sh.Cells(5, 1).Value = "  Variable Name  "

    'Act and assert
    Assert.AreEqual CLng(5), InitTransfer.SetupStartRow(sh, 5, 1, DICT_FIRST_HEADER), _
                    "The header probe should match without regard to case or padding."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestSetupStartRowMatchesHeaderCaseInsensitive", Err.Number, Err.Description
End Sub

'@TestMethod("InitTransfer.AnchorRows")
Public Sub TestSetupStartRowKeepsHintWhenNoProbeMatches()
    CustomTestSetTitles Assert, "InitTransfer", "TestSetupStartRowKeepsHintWhenNoProbeMatches"
    On Error GoTo Fail

    'Arrange: a blank sheet matches neither candidate row. An unsaved
    'fixture workbook reads as the .xlsx layout, so the hint is row 1.
    Dim sh As Worksheet
    Set sh = EnsureWorksheet("Dictionary", FixtureWorkbook)

    'Act and assert
    Assert.AreEqual CLng(1), InitTransfer.SetupStartRow(sh, 5, 1, DICT_FIRST_HEADER), _
                    "With no header found, the hinted row should be kept."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestSetupStartRowKeepsHintWhenNoProbeMatches", Err.Number, Err.Description
End Sub


'@section Language Sync Tests
'===============================================================================
'@TestMethod("InitTransfer.LanguageSync")
Public Sub TestSyncWritesDictionaryAndInterfaceCodes()
    CustomTestSetTitles Assert, "InitTransfer", "TestSyncWritesDictionaryAndInterfaceCodes"
    On Error GoTo Fail

    'Arrange
    ArrangeMainLanguageRanges "ENG", "FRA-Francais"

    'Act
    InitTransfer.SyncDesignerLanguageNames FixtureWorkbook

    'Assert: both codes land as workbook-level hidden names
    Dim store As HiddenNames
    Set store = HiddenNames.Create(FixtureWorkbook)
    Assert.AreEqual "ENG", store.ValueAsString(TAG_DICT_LANG), _
                    "The dictionary language should land on its hidden name."
    Assert.AreEqual "FRA", store.ValueAsString(TAG_LL_LANG_CODE), _
                    "The interface language should land as the code before the dash."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestSyncWritesDictionaryAndInterfaceCodes", Err.Number, Err.Description
End Sub

'@TestMethod("InitTransfer.LanguageSync")
Public Sub TestSyncPassesBareInterfaceValueThrough()
    CustomTestSetTitles Assert, "InitTransfer", "TestSyncPassesBareInterfaceValueThrough"
    On Error GoTo Fail

    'Arrange: an interface value with no dash
    ArrangeMainLanguageRanges "ENG", "ENG"

    'Act
    InitTransfer.SyncDesignerLanguageNames FixtureWorkbook

    'Assert
    Dim store As HiddenNames
    Set store = HiddenNames.Create(FixtureWorkbook)
    Assert.AreEqual "ENG", store.ValueAsString(TAG_LL_LANG_CODE), _
                    "A bare interface value should pass through whole."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestSyncPassesBareInterfaceValueThrough", Err.Number, Err.Description
End Sub

'@TestMethod("InitTransfer.LanguageSync")
Public Sub TestSyncHandlesEmptyInterfaceValue()
    CustomTestSetTitles Assert, "InitTransfer", "TestSyncHandlesEmptyInterfaceValue"
    On Error GoTo Fail

    'Arrange: RNG_LLForm left empty. The Split guard has to keep the sync
    'from raising on the empty answer.
    ArrangeMainLanguageRanges "ENG", vbNullString

    'Act
    InitTransfer.SyncDesignerLanguageNames FixtureWorkbook

    'Assert
    Dim store As HiddenNames
    Set store = HiddenNames.Create(FixtureWorkbook)
    Assert.AreEqual vbNullString, store.ValueAsString(TAG_LL_LANG_CODE), _
                    "An empty interface value should land as an empty code with no raise."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestSyncHandlesEmptyInterfaceValue", Err.Number, Err.Description
End Sub


'@section Design Name Tests
'===============================================================================
'@TestMethod("InitTransfer.DesignName")
Public Sub TestDesignerDesignNameReadsFormatSheet()
    CustomTestSetTitles Assert, "InitTransfer", "TestDesignerDesignNameReadsFormatSheet"
    On Error GoTo Fail

    'Arrange: a format sheet whose DESIGNTYPE range carries a padded name
    Dim sh As Worksheet
    Set sh = EnsureWorksheet("__formatter", FixtureWorkbook)
    sh.Range("B2").Value = "  design 2  "
    FixtureWorkbook.Names.Add Name:="DESIGNTYPE", RefersTo:=sh.Range("B2")

    'Act and assert
    Assert.AreEqual "design 2", InitTransfer.DesignerDesignName(FixtureWorkbook), _
                    "The design name should come back trimmed from DESIGNTYPE."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestDesignerDesignNameReadsFormatSheet", Err.Number, Err.Description
End Sub

'@TestMethod("InitTransfer.DesignName")
Public Sub TestDesignerDesignNameEmptyWithoutFormatSheet()
    CustomTestSetTitles Assert, "InitTransfer", "TestDesignerDesignNameEmptyWithoutFormatSheet"
    On Error GoTo Fail

    'Act and assert: a workbook with no __formatter sheet names no design
    Assert.AreEqual vbNullString, InitTransfer.DesignerDesignName(FixtureWorkbook), _
                    "A missing format sheet should answer an empty design name."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestDesignerDesignNameEmptyWithoutFormatSheet", Err.Number, Err.Description
End Sub

'@TestMethod("InitTransfer.DesignName")
Public Sub TestDesignerDesignNameEmptyWithoutRange()
    CustomTestSetTitles Assert, "InitTransfer", "TestDesignerDesignNameEmptyWithoutRange"
    On Error GoTo Fail

    'Arrange: the format sheet exists and DESIGNTYPE is missing
    EnsureWorksheet "__formatter", FixtureWorkbook

    'Act and assert
    Assert.AreEqual vbNullString, InitTransfer.DesignerDesignName(FixtureWorkbook), _
                    "A format sheet without DESIGNTYPE should answer an empty design name."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestDesignerDesignNameEmptyWithoutRange", Err.Number, Err.Description
End Sub


'@section Formatter Branch Tests
'===============================================================================
'@TestMethod("InitTransfer.Formatter")
Public Sub TestExportFormatterFromSetupWritesDesignName()
    CustomTestSetTitles Assert, "InitTransfer", "TestExportFormatterFromSetupWritesDesignName"
    On Error GoTo Fail

    'Arrange: the setup workbook holds the format sheet
    Set SourceWorkbook = NewWorkbook
    PrepareLLFormatFixture "__formatter", SourceWorkbook

    'Act: the setup branch of the transfer, with the design of the run
    InitTransfer.ExportFormatterFromSetup SourceWorkbook, FixtureWorkbook, "design 2"

    'Assert: the sheet lands very hidden and DESIGNTYPE names the design of
    'the run. The export used to keep the source's own DESIGNTYPE value, so
    'the shipped sheet named a design the build never applied.
    AssertExportedFormatter "design 2", "the setup branch"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestExportFormatterFromSetupWritesDesignName", Err.Number, Err.Description
End Sub

'@TestMethod("InitTransfer.Formatter")
Public Sub TestExportFormatterFromDesignerWritesDesignName()
    CustomTestSetTitles Assert, "InitTransfer", "TestExportFormatterFromDesignerWritesDesignName"
    On Error GoTo Fail

    'Arrange: the designer workbook holds the format sheet
    Set SourceWorkbook = NewWorkbook
    PrepareLLFormatFixture "__formatter", SourceWorkbook

    'Act: the designer branch of the transfer
    InitTransfer.ExportFormatterFromDesigner SourceWorkbook, FixtureWorkbook, "design 2"

    'Assert
    AssertExportedFormatter "design 2", "the designer branch"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestExportFormatterFromDesignerWritesDesignName", Err.Number, Err.Description
End Sub


'@section Report Entry Tests
'===============================================================================
'@TestMethod("InitTransfer.ReportEntries")
Public Sub TestMissingTranslationsSheetFilesReportEntry()
    CustomTestSetTitles Assert, "InitTransfer", "TestMissingTranslationsSheetFilesReportEntry"
    On Error GoTo Fail

    'Arrange: a setup workbook with no Translations sheet
    Set SourceWorkbook = NewWorkbook

    'Act
    InitTransfer.ImportSetupTranslationTable SourceWorkbook, FixtureWorkbook

    'Assert: the skip speaks through the transfer record
    Assert.IsTrue InitTransfer.HasCheckings(), _
                  "A setup without its Translations sheet should file a report entry."
    Assert.IsTrue InitTransfer.CheckingValues().KeyExists("setup translations"), _
                  "The entry should carry the setup translations key."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestMissingTranslationsSheetFilesReportEntry", Err.Number, Err.Description
End Sub


'@section Test helpers
'===============================================================================

'@sub-title Create the two Main language ranges and fill them
Private Sub ArrangeMainLanguageRanges(ByVal dictLang As String, ByVal llForm As String)
    Dim mainSheet As Worksheet
    Set mainSheet = EnsureWorksheet("Main", FixtureWorkbook)

    FixtureWorkbook.Names.Add Name:="RNG_LangSetup", RefersTo:=mainSheet.Range("A1")
    FixtureWorkbook.Names.Add Name:="RNG_LLForm", RefersTo:=mainSheet.Range("A2")
    mainSheet.Range("A1").Value = dictLang
    mainSheet.Range("A2").Value = llForm
End Sub

'@sub-title Assert the exported format sheet of the target workbook
Private Sub AssertExportedFormatter(ByVal designName As String, ByVal branchName As String)
    Dim targetSheet As Worksheet

    Assert.IsTrue WorksheetExists("__formatter", FixtureWorkbook), _
                  "The format sheet should land on the target through " & branchName & "."

    Set targetSheet = FixtureWorkbook.Worksheets("__formatter")
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(targetSheet.Visible), _
                    "The exported format sheet should be very hidden."
    Assert.AreEqual designName, CStr(targetSheet.Range("DESIGNTYPE").Cells(1, 1).Value), _
                    "DESIGNTYPE should name the design of the run through " & branchName & "."
End Sub
