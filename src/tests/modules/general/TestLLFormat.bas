Attribute VB_Name = "TestLLFormat"
Attribute VB_Description = "Behavioural tests for LLFormat"
Option Explicit



'@Folder("CustomTests")
'@ModuleDescription("Behavioural tests for LLFormat")
'@details Exercises LLFormat creation, value lookups, formatting scopes, and import behaviour.
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private FormatWorkbook As Workbook
Private FormatSheet As Worksheet
Private FormatUnderTest As ILLFormat

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FORMAT_SHEET_NAME As String = "LLFormatFixture_Test"
Private Const IMPORT_SHEET_NAME As String = "LLFormatImport_Test"
Private Const LABEL_ANALYSIS_BASE_FONT_SIZE As String = "analysis base font size"
Private Const LABEL_MISSING_FONT_COLOR As String = "missing font color"

Private Function FixtureDefaultDesign() As String
    If FormatSheet Is Nothing Then
        Err.Raise vbObjectError + 601, "TestLLFormat.FixtureDefaultDesign", _
                  "Format sheet must be initialised before requesting the default design"
    End If
    FixtureDefaultDesign = LLFormatTestFixture.DefaultDesignName(FormatSheet)
End Function

Private Function FixtureSecondaryDesign() As String
    Dim names As Collection
    If FormatSheet Is Nothing Then
        Err.Raise vbObjectError + 602, "TestLLFormat.FixtureSecondaryDesign", _
                  "Format sheet must be initialised before requesting design names"
    End If
    Set names = LLFormatTestFixture.DesignNames(FormatSheet)
    If names.Count < 2 Then
        Err.Raise vbObjectError + 600, "TestLLFormat.FixtureSecondaryDesign", _
                  "Fixture does not expose a secondary design column"
    End If
    FixtureSecondaryDesign = CStr(names.Item(2))
End Function

Private Function ExpectedDesignColour(ByVal labelName As String, _
                                      Optional ByVal designName As String = vbNullString) As Long
    ExpectedDesignColour = LLFormatTestFixture.DesignColour(FormatSheet, labelName, designName)
End Function

Private Function ExpectedDesignValue(ByVal labelName As String, _
                                     Optional ByVal designName As String = vbNullString) As Variant
    ExpectedDesignValue = LLFormatTestFixture.DesignNumericValue(FormatSheet, labelName, designName)
End Function

Private Function RequireNumericLong(ByVal candidate As Variant, ByVal context As String) As Long
    On Error GoTo ConversionError

    If IsObject(candidate) Then
        CustomTestLogFailure Assert, context & " returned an object reference; numeric value expected"
        RequireNumericLong = 0
        Exit Function
    End If

    If IsNull(candidate) Then
        CustomTestLogFailure Assert, context & " returned Null; numeric value expected"
        RequireNumericLong = 0
        Exit Function
    End If

    If VarType(candidate) = vbEmpty Then
        CustomTestLogFailure Assert, context & " returned Empty; numeric value expected"
        RequireNumericLong = 0
        Exit Function
    End If

    If IsError(candidate) Then
        CustomTestLogFailure Assert, context & " returned an error value; numeric value expected"
        RequireNumericLong = 0
        Exit Function
    End If

    If VarType(candidate) = vbString Then
        If LenB(Trim$(CStr(candidate))) = 0 Then
            CustomTestLogFailure Assert, context & " returned an empty string; numeric value expected"
            RequireNumericLong = 0
            Exit Function
        End If
    End If

    RequireNumericLong = CLng(candidate)
    Exit Function

ConversionError:
    CustomTestLogFailure Assert, context & " could not be converted to Long (type: " & TypeName(candidate) & ")"
    RequireNumericLong = 0
End Function

Private Function RequireNumericDouble(ByVal candidate As Variant, ByVal context As String) As Double
    On Error GoTo ConversionError

    If IsObject(candidate) Then
        CustomTestLogFailure Assert, context & " returned an object reference; numeric value expected"
        RequireNumericDouble = 0#
        Exit Function
    End If

    If IsNull(candidate) Then
        CustomTestLogFailure Assert, context & " returned Null; numeric value expected"
        RequireNumericDouble = 0#
        Exit Function
    End If

    If VarType(candidate) = vbEmpty Then
        CustomTestLogFailure Assert, context & " returned Empty; numeric value expected"
        RequireNumericDouble = 0#
        Exit Function
    End If

    If IsError(candidate) Then
        CustomTestLogFailure Assert, context & " returned an error value; numeric value expected"
        RequireNumericDouble = 0#
        Exit Function
    End If

    If VarType(candidate) = vbString Then
        If LenB(Trim$(CStr(candidate))) = 0 Then
            CustomTestLogFailure Assert, context & " returned an empty string; numeric value expected"
            RequireNumericDouble = 0#
            Exit Function
        End If
    End If

    RequireNumericDouble = CDbl(candidate)
    Exit Function

ConversionError:
    CustomTestLogFailure Assert, context & " could not be converted to Double (type: " & TypeName(candidate) & ")"
    RequireNumericDouble = 0#
End Function

'@ModuleInitialize
'@description Configure common test state and build the assertion helper.
Public Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLFormat"
    Set FormatWorkbook = ThisWorkbook
End Sub

'@ModuleCleanup
'@description Tear down shared resources and print accumulated results.
Public Sub ModuleCleanup()
    On Error Resume Next
        LLFormatTestFixture.DeleteLLFormatFixture FORMAT_SHEET_NAME, FormatWorkbook
        LLFormatTestFixture.DeleteLLFormatFixture IMPORT_SHEET_NAME, FormatWorkbook
    On Error GoTo 0
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    Set FormatUnderTest = Nothing
    Set FormatSheet = Nothing
    Set FormatWorkbook = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
'@description Prepare a fresh LL format worksheet and system under test for each test.
Public Sub TestInitialize()
    LLFormatTestFixture.DeleteLLFormatFixture FORMAT_SHEET_NAME, FormatWorkbook
    LLFormatTestFixture.DeleteLLFormatFixture IMPORT_SHEET_NAME, FormatWorkbook

    Set FormatSheet = LLFormatTestFixture.PrepareLLFormatFixture(FORMAT_SHEET_NAME, FormatWorkbook)
    FormatSheet.Range("DESIGNTYPE").Value = FixtureDefaultDesign()

    Set FormatUnderTest = LLFormat.Create(FormatSheet)
End Sub

'@TestCleanup
'@description Flush assertions and remove any worksheets created during the test run.
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    On Error Resume Next
        LLFormatTestFixture.DeleteLLFormatFixture IMPORT_SHEET_NAME, FormatWorkbook
        LLFormatTestFixture.DeleteLLFormatFixture FORMAT_SHEET_NAME, FormatWorkbook
        LLFormatTestFixture.DeleteLLFormatFixture "LLFormatFixture_DesignRange", FormatWorkbook
        TestHelpers.DeleteWorksheet "LLFormat_AllAnalysis_Test"
    On Error GoTo 0

    Set FormatUnderTest = Nothing
    Set FormatSheet = Nothing
End Sub

'@TestMethod("LLFormat")
'@description Creating with an unknown design should use the default design values.
Public Sub TestCreateFallsBackToDefaultDesign()
    CustomTestSetTitles Assert, "LLFormat", "TestCreateFallsBackToDefaultDesign"
    Dim sut As ILLFormat

    Set sut = LLFormat.Create(FormatSheet, designType:="unknown design")

    Dim fixtureValue As Variant
    Dim expectedLong As Long
    Dim defaultLong As Long
    Dim fallbackLong As Long

    fixtureValue = ExpectedDesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE)
    expectedLong = RequireNumericLong(fixtureValue, _
                                      "Fixture value for '" & LABEL_ANALYSIS_BASE_FONT_SIZE & "'")
    defaultLong = RequireNumericLong(FormatUnderTest.DesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE, False), _
                                     "Default design value for '" & LABEL_ANALYSIS_BASE_FONT_SIZE & "'")

    Assert.AreEqual expectedLong, defaultLong, _
                     "Default design should match the fixture value for '" & LABEL_ANALYSIS_BASE_FONT_SIZE & "'"

    fallbackLong = RequireNumericLong(sut.DesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE, False), _
                                      "Fallback design value for '" & LABEL_ANALYSIS_BASE_FONT_SIZE & "'")

    Assert.AreEqual expectedLong, fallbackLong, _
                     "Fallback design should still produce values"
End Sub

'@TestMethod("LLFormat")
'@description DesignValue should return the configured colour for the default design.
Public Sub TestDesignValueReturnsConfiguredColour()
    CustomTestSetTitles Assert, "LLFormat", "TestDesignValueReturnsConfiguredColour"
    Dim colorValue As Long

    colorValue = RequireNumericLong(FormatUnderTest.DesignValue(LABEL_MISSING_FONT_COLOR), _
                                    "Default colour value for '" & LABEL_MISSING_FONT_COLOR & "'")

    Assert.AreEqual ExpectedDesignColour(LABEL_MISSING_FONT_COLOR), colorValue, _
                     "DesignValue should return configured color for the default design"
End Sub

'@TestMethod("LLFormat")
'@description DesignValue should expose the stored numeric value when colour is not requested.
Public Sub TestDesignValueReturnsCellValue()
    CustomTestSetTitles Assert, "LLFormat", "TestDesignValueReturnsCellValue"
    Dim expectedLong As Long
    Dim actualLong As Long

    expectedLong = RequireNumericLong(ExpectedDesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE), _
                                      "Fixture value for '" & LABEL_ANALYSIS_BASE_FONT_SIZE & "'")
    actualLong = RequireNumericLong(FormatUnderTest.DesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE, False), _
                                    "Design value for '" & LABEL_ANALYSIS_BASE_FONT_SIZE & "'")

    Assert.AreEqual expectedLong, actualLong, _
                     "DesignValue should return the configured numeric value when returnColor is False"
End Sub

'@TestMethod("LLFormat")
'@description Missing labels should return fallback values and log a checking entry.
Public Sub TestDesignValueMissingLabelFallsBackAndLogs()
    CustomTestSetTitles Assert, "LLFormat", "TestDesignValueMissingLabelFallsBackAndLogs"
    Dim colourValue As Long
    Dim numericValue As Long
    Dim keys As BetterArray
    Dim firstKey As String
    Dim logEntry As IChecking

    colourValue = RequireNumericLong(FormatUnderTest.DesignValue("missing label"), _
                                     "Fallback colour for missing labels")
    Assert.AreEqual CLng(vbBlack), colourValue, _
                     "Missing labels should return the fallback colour"

    numericValue = RequireNumericLong(FormatUnderTest.DesignValue("missing label", False), _
                                      "Fallback numeric value for missing labels")
    Assert.AreEqual 0, numericValue, _
                     "Missing labels should return the fallback numeric value"

    Assert.IsTrue FormatUnderTest.HasCheckings, _
                  "Missing labels should enqueue checking information"

    Set logEntry = FormatUnderTest.CheckingValues
    Assert.ObjectExists logEntry, "Checking", "Checking log should be provided for missing labels"

    Set keys = logEntry.ListOfKeys
    Assert.IsTrue (keys.Length > 0), "Checking log should contain at least one entry"
    firstKey = CStr(keys.Item(keys.LowerBound))
    Assert.IsTrue InStr(1, logEntry.ValueOf(firstKey, checkingLabel), "missing label", vbTextCompare) > 0, _
                  "Checking log should reference the missing label"
End Sub

'@TestMethod("LLFormat")
'@description Applying the analysis section scope should honour design-driven styling.
Public Sub TestApplyFormatAnalysisSectionUsesDesignSettings()
    CustomTestSetTitles Assert, "LLFormat", "TestApplyFormatAnalysisSectionUsesDesignSettings"
    Dim target As Range
    Dim applied As Range
    Dim expectedFontColour As Long
    Dim expectedFontSize As Double

    Set target = FormatSheet.Range("G10")
    target.Value = "Section title"

    FormatUnderTest.ApplyFormat target, AnalysisSection

    Set applied = target.Parent.Range(target.Cells(1, 1), target.Cells(1, 7))
    expectedFontColour = ExpectedDesignColour("table sections font color")
    expectedFontSize = RequireNumericDouble(ExpectedDesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE), _
                                            "Fixture value for '" & LABEL_ANALYSIS_BASE_FONT_SIZE & "'") + 5

    Dim appliedFontSize As Double
    appliedFontSize = RequireNumericDouble(applied.Font.Size, "Applied analysis section font size")

    Assert.AreEqual expectedFontColour, CLng(applied.Font.Color), "Section font colour should come from design"
    Assert.IsTrue applied.Font.Bold, "Section text should be bold"
    Assert.AreEqual CLng(expectedFontSize), CLng(appliedFontSize), "Section font size should add the section boost"
    Assert.IsTrue applied.Cells(1, 1).WrapText, "Section header should enable wrapping"
End Sub

'@TestMethod("LLFormat")
'@description Applying the analysis one-cell scope should apply missing-value formatting.
Public Sub TestApplyFormatAnalysisOneCellAppliesMissingColours()
    CustomTestSetTitles Assert, "LLFormat", "TestApplyFormatAnalysisOneCellAppliesMissingColours"
    Dim target As Range

    Set target = FormatSheet.Range("H15")
    target.Value = "Missing value"

    FormatUnderTest.ApplyFormat target, AnalysisOneCell

    Assert.AreEqual ExpectedDesignColour(LABEL_MISSING_FONT_COLOR), CLng(target.Font.Color), _
                     "Missing cell font colour should come from the design"
    Assert.AreEqual ExpectedDesignColour("missing interior color"), CLng(target.Interior.Color), _
                     "Missing cell interior colour should come from the design"
    Assert.IsTrue target.Font.Bold, "Missing cell should be bold"
End Sub

'@TestMethod("LLFormat")
'@description Applying the all-analysis scope should set worksheet font and dimensions.
Public Sub TestApplyFormatAllAnalysisSheetUsesDesignDimensions()
    CustomTestSetTitles Assert, "LLFormat", "TestApplyFormatAllAnalysisSheetUsesDesignDimensions"
    Dim tempSheet As Worksheet
    Dim expectedFontSize As Double
    Dim expectedColumnWidth As Double

    Set tempSheet = EnsureWorksheet("LLFormat_AllAnalysis_Test")
    tempSheet.Cells.Clear

    FormatUnderTest.ApplyFormat tempSheet, AllAnalysisSheet

    expectedFontSize = RequireNumericDouble(ExpectedDesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE), _
                                            "Fixture value for '" & LABEL_ANALYSIS_BASE_FONT_SIZE & "'")
    If expectedFontSize = 0 Then expectedFontSize = 9
    expectedColumnWidth = RequireNumericDouble(ExpectedDesignValue("default analysis column width"), _
                                               "Fixture value for 'default analysis column width'")
    If expectedColumnWidth = 0 Then expectedColumnWidth = 25

    Dim actualFontSize As Double
    actualFontSize = RequireNumericDouble(tempSheet.Cells.Font.Size, "All analysis sheet font size")

    Assert.AreEqual CLng(expectedFontSize), CLng(actualFontSize), _
                     "Worksheet font size should match the design value"
    Assert.AreEqual expectedColumnWidth, tempSheet.Columns(1).ColumnWidth, _
                     "Worksheet column width should match the design value"
    Assert.AreEqual 25, tempSheet.Rows(2).RowHeight, "Row height for row 2 should match specification"

    TestHelpers.DeleteWorksheet "LLFormat_AllAnalysis_Test"
End Sub

'@TestMethod("LLFormat")
'@description Preparing a fixture should define the DESIGNTYPE named range with the default.
Public Sub TestPrepareFixtureDefinesDesignTypeRange()
    CustomTestSetTitles Assert, "LLFormat", "TestPrepareFixtureDefinesDesignTypeRange"
    Dim sheetName As String
    Dim fixtureSheet As Worksheet
    Dim designValue As String

    sheetName = "LLFormatFixture_DesignRange"
    On Error Resume Next
        LLFormatTestFixture.DeleteLLFormatFixture sheetName, FormatWorkbook
    On Error GoTo 0

    Set fixtureSheet = LLFormatTestFixture.PrepareLLFormatFixture(sheetName, FormatWorkbook)
    designValue = CStr(fixtureSheet.Range("DESIGNTYPE").Value)

    Assert.AreEqual FixtureDefaultDesign(), designValue, _
                     "Prepared fixture should seed the design type named range"

    LLFormatTestFixture.DeleteLLFormatFixture sheetName, FormatWorkbook
End Sub
'@TestMethod("LLFormat")
'@description Percent scope formatting should enforce a two-decimal percent number format.
Public Sub TestApplyFormatPercentSetsNumberFormat()
    CustomTestSetTitles Assert, "LLFormat", "TestApplyFormatPercentSetsNumberFormat"
    Dim target As Range

    Set target = FormatSheet.Range("H1")
    target.Clear
    target.Value = 0.25

    FormatUnderTest.ApplyFormat target, AnalysisPercent

    Assert.AreEqual "0.00%", target.NumberFormat, _
                     "Percent scope should enforce 2 decimal percent format"
End Sub

'@TestMethod("LLFormat")
'@description Importing from another sheet should copy font and interior colours for designs.
Public Sub TestImportCopiesDesignColours()
    CustomTestSetTitles Assert, "LLFormat", "TestImportCopiesDesignColours"
    Dim importSheet As Worksheet
    Dim colorValue As Long

    Dim secondaryDesign As String
    secondaryDesign = FixtureSecondaryDesign()

    Set importSheet = LLFormatTestFixture.PrepareLLFormatFixture(IMPORT_SHEET_NAME, FormatWorkbook)
    With LLFormatTestFixture.FixtureCell(importSheet, LABEL_MISSING_FONT_COLOR, secondaryDesign)
        .Interior.Color = RGB(0, 255, 0)
        .Font.Color = RGB(0, 0, 255)
    End With
    importSheet.Range("DESIGNTYPE").Value = secondaryDesign

    FormatUnderTest.Import importSheet

    colorValue = RequireNumericLong(FormatUnderTest.DesignValue(LABEL_MISSING_FONT_COLOR), _
                                    "Imported colour value for '" & LABEL_MISSING_FONT_COLOR & "'")
    Assert.AreEqual RGB(0, 255, 0), colorValue, _
                     "Import should copy interior colours for alternate designs"

    Dim formatCell As Range
    Set formatCell = LLFormatTestFixture.FixtureCell(FormatSheet, LABEL_MISSING_FONT_COLOR, secondaryDesign)
    Assert.AreEqual RGB(0, 0, 255), CLng(formatCell.Font.Color), _
                     "Import should copy font colours for alternate designs"
    Assert.AreEqual secondaryDesign, CStr(FormatSheet.Range("DESIGNTYPE").Value), _
                     "Design type cell should update to imported design"
End Sub

