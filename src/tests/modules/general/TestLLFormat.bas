Attribute VB_Name = "TestLLFormat"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private FormatWorkbook As Workbook
Private FormatTemplate As Worksheet
Private FormatSheet As Worksheet
Private FormatUnderTest As ILLFormat

Private Const FORMAT_SHEET_NAME As String = "LLFormatFixture_Test"
Private Const IMPORT_SHEET_NAME As String = "LLFormatImport_Test"
Private Const DEFAULT_DESIGN As String = "design 1"
Private Const SECONDARY_DESIGN As String = "design 2"
Private Const LABEL_ANALYSIS_BASE_FONT_SIZE As String = "analysis base font size"
Private Const LABEL_MISSING_FONT_COLOR As String = "missing font color"

'@ModuleInitialize
Private Sub ModuleInitialize()
    On Error GoTo FailTemplate

    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set FormatWorkbook = ThisWorkbook
    Set FormatTemplate = LLFormatTestFixture.LLFormatTemplate(FormatWorkbook)

    ValidateFixture FormatTemplate
    Exit Sub

FailTemplate:
    If Not Assert Is Nothing Then
        Assert.Fail "LLFormat fixture worksheet validation failed: " & Err.Description
    Else
        Err.Raise Err.Number, "TestLLFormat.ModuleInitialize", Err.Description
    End If
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        LLFormatTestFixture.DeleteLLFormatFixture FORMAT_SHEET_NAME, FormatWorkbook
        LLFormatTestFixture.DeleteLLFormatFixture IMPORT_SHEET_NAME, FormatWorkbook
    On Error GoTo 0

    Set FormatUnderTest = Nothing
    Set FormatSheet = Nothing
    Set FormatTemplate = Nothing
    Set FormatWorkbook = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    LLFormatTestFixture.DeleteLLFormatFixture FORMAT_SHEET_NAME, FormatWorkbook
    LLFormatTestFixture.DeleteLLFormatFixture IMPORT_SHEET_NAME, FormatWorkbook

    Set FormatSheet = LLFormatTestFixture.PrepareLLFormatFixture(FORMAT_SHEET_NAME, FormatWorkbook)
    FormatSheet.Range("DESIGNTYPE").Value = DEFAULT_DESIGN

    Set FormatUnderTest = LLFormat.Create(FormatSheet)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    On Error Resume Next
        LLFormatTestFixture.DeleteLLFormatFixture IMPORT_SHEET_NAME, FormatWorkbook
        LLFormatTestFixture.DeleteLLFormatFixture FORMAT_SHEET_NAME, FormatWorkbook
    On Error GoTo 0

    Set FormatUnderTest = Nothing
    Set FormatSheet = Nothing
End Sub

'@TestMethod("LLFormat")
Private Sub TestCreateFallsBackToDefaultDesign()
    Dim sut As ILLFormat

    Set sut = LLFormat.Create(FormatSheet, designType:="unknown design")

    Assert.AreEqual 11, sut.DesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE, False), _
                     "Fallback design should still produce values"
End Sub

'@TestMethod("LLFormat")
Private Sub TestDesignValueReturnsConfiguredColour()
    Dim colorValue As Long

    colorValue = CLng(FormatUnderTest.DesignValue(LABEL_MISSING_FONT_COLOR))

    Assert.AreEqual RGB(255, 0, 0), colorValue, _
                     "DesignValue should return configured color for the default design"
End Sub

'@TestMethod("LLFormat")
Private Sub TestApplyFormatPercentSetsNumberFormat()
    Dim target As Range

    Set target = FormatSheet.Range("H1")
    target.Clear
    target.Value = 0.25

    FormatUnderTest.ApplyFormat target, AnalysisPercent

    Assert.AreEqual "0.00%", target.NumberFormat, _
                     "Percent scope should enforce 2 decimal percent format"
End Sub

'@TestMethod("LLFormat")
Private Sub TestImportCopiesDesignColours()
    Dim importSheet As Worksheet
    Dim colorValue As Long

    Set importSheet = LLFormatTestFixture.PrepareLLFormatFixture(IMPORT_SHEET_NAME, FormatWorkbook)
    LLFormatTestFixture.FixtureCell(importSheet, LABEL_MISSING_FONT_COLOR, SECONDARY_DESIGN).Interior.Color = RGB(0, 255, 0)
    importSheet.Range("DESIGNTYPE").Value = SECONDARY_DESIGN

    FormatUnderTest.Import importSheet

    colorValue = CLng(FormatUnderTest.DesignValue(LABEL_MISSING_FONT_COLOR))
    Assert.AreEqual RGB(0, 255, 0), colorValue, _
                     "Import should copy interior colours for alternate designs"
End Sub

'@section Fixture Validation
'===============================================================================
Private Sub ValidateFixture(ByVal template As Worksheet)
    Dim fixtureFormat As ILLFormat
    Dim colorValue As Long

    Set fixtureFormat = LLFormat.Create(template)

    Assert.AreEqual 11, fixtureFormat.DesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE, False), _
                     "Fixture must expose expected base font size value"

    colorValue = CLng(fixtureFormat.DesignValue(LABEL_MISSING_FONT_COLOR))
    Assert.AreEqual RGB(255, 0, 0), colorValue, _
                     "Fixture must expose expected default font colour"

    Set fixtureFormat = Nothing
End Sub
