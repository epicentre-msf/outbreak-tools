Attribute VB_Name = "TestLLFormat"
Attribute VB_Description = "Behavioural tests for LLFormat"
Option Explicit

'===============================================================================
' @ModuleDescription Behavioural tests for LLFormat
'
' @description
'   Comprehensive test suite for the LLFormat class (ILLFormat interface), which
'   manages linelist visual formatting including design colours, font sizes, cell
'   styling, and import/export of format data between workbooks. Tests are grouped
'   into seven logical areas:
'
'     1. Factory / Creation  -- fallback to default design when an unknown design
'        name is supplied.
'     2. DesignValue lookups -- colour retrieval, numeric value retrieval, and
'        graceful fallback with checking log when a label is missing.
'     3. ApplyFormat scopes  -- AnalysisSection, AnalysisOneCell, AllAnalysisSheet,
'        and AnalysisPercent formatting applied to ranges and worksheets.
'     4. Fixture preparation -- verifying that LLFormatTestFixture correctly seeds
'        the DESIGNTYPE named range.
'     5. Import behaviour    -- copying design colours and font colours from an
'        import sheet into the active format instance.
'     6. Export behaviour    -- seven tests covering new sheet creation, table data
'        transfer, style copying, ListObject creation, named range creation, design
'        value preservation, existing-sheet import fallback, and the Nothing guard.
'     7. Error handling      -- InvalidArgument when Export receives Nothing.
'
'   Each test follows Arrange / Act / Assert. A fresh fixture worksheet is built
'   before every test via TestInitialize and torn down in TestCleanup, ensuring
'   test isolation.
'
' @depends LLFormat, ILLFormat, Checking, IChecking, BetterArray, CustomTest,
'          TestHelpers, LLFormatTestFixture
'===============================================================================

'@Folder("CustomTests")
'@ModuleDescription("Behavioural tests for LLFormat")
'@details Exercises LLFormat creation, value lookups, formatting scopes, import and export behaviour.
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private FormatWorkbook As Workbook
Private FormatSheet As Worksheet
Private FormatUnderTest As ILLFormat

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FORMAT_SHEET_NAME As String = "LLFormatFixture_Test"
Private Const IMPORT_SHEET_NAME As String = "LLFormatImport_Test"
Private Const EXPORT_SHEET_NAME As String = "LLFormatExport_Test"
Private Const LABEL_ANALYSIS_BASE_FONT_SIZE As String = "analysis base font size"
Private Const LABEL_MISSING_FONT_COLOR As String = "missing font color"

'-------------------------------------------------------------------------------
' Private helpers
'-------------------------------------------------------------------------------

' @sub-title FixtureDefaultDesign
' @description Retrieves the default design name from the fixture sheet via the
'   LLFormatTestFixture helper. Raises an error if FormatSheet has not been
'   initialised, protecting against out-of-order test execution.
Private Function FixtureDefaultDesign() As String
    If FormatSheet Is Nothing Then
        Err.Raise vbObjectError + 601, "TestLLFormat.FixtureDefaultDesign", _
                  "Format sheet must be initialised before requesting the default design"
    End If
    FixtureDefaultDesign = LLFormatTestFixture.DefaultDesignName(FormatSheet)
End Function

' @sub-title FixtureSecondaryDesign
' @description Returns the second design name exposed by the fixture, used by
'   tests that need to verify behaviour against a non-default design column.
'   Raises if FormatSheet is Nothing or the fixture has fewer than two designs.
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

' @sub-title ExpectedDesignColour
' @description Convenience wrapper that delegates to LLFormatTestFixture.DesignColour
'   to obtain the expected interior colour for a given label/design pair. When
'   designName is omitted the fixture returns the default design colour.
Private Function ExpectedDesignColour(ByVal labelName As String, _
                                      Optional ByVal designName As String = vbNullString) As Long
    ExpectedDesignColour = LLFormatTestFixture.DesignColour(FormatSheet, labelName, designName)
End Function

' @sub-title ExpectedDesignValue
' @description Convenience wrapper that delegates to LLFormatTestFixture.DesignNumericValue
'   to obtain the expected numeric cell value for a given label/design pair.
Private Function ExpectedDesignValue(ByVal labelName As String, _
                                     Optional ByVal designName As String = vbNullString) As Variant
    ExpectedDesignValue = LLFormatTestFixture.DesignNumericValue(FormatSheet, labelName, designName)
End Function

' @sub-title RequireNumericLong
' @description Type-safe conversion helper that attempts to convert an arbitrary
'   Variant to Long. When the candidate is an object, Null, Empty, an error value,
'   an empty string, or otherwise unconvertible, the function logs a descriptive
'   test failure via CustomTestLogFailure and returns 0 so the calling test can
'   continue with meaningful assertion messages rather than crashing.
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

' @sub-title RequireNumericDouble
' @description Type-safe conversion helper that attempts to convert an arbitrary
'   Variant to Double. Mirrors the guard logic of RequireNumericLong but returns
'   0# (Double zero) on failure. Used for font sizes and column widths where
'   fractional precision matters.
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

' @sub-title VerifyTableStructureMatches
' @description Asserts that a source and target ListObject share the same column
'   count and row count. Used by export tests to confirm that structural fidelity
'   is preserved across workbook boundaries.
Private Sub VerifyTableStructureMatches(ByVal sourceTable As ListObject, _
                                       ByVal targetTable As ListObject, _
                                       ByVal context As String)
    Assert.AreEqual sourceTable.ListColumns.Count, targetTable.ListColumns.Count, _
                     context & ": Column count should match between source and target"
    Assert.AreEqual sourceTable.DataBodyRange.Rows.Count, targetTable.DataBodyRange.Rows.Count, _
                     context & ": Row count should match between source and target"
End Sub

' @sub-title VerifyCellFormatting
' @description Asserts that font colour, interior colour, and bold state match
'   between a source and target cell. Provides per-label failure messages so
'   formatting regressions are easy to localise.
Private Sub VerifyCellFormatting(ByVal sourceCell As Range, _
                                ByVal targetCell As Range, _
                                ByVal labelName As String)
    Assert.AreEqual CLng(sourceCell.Font.Color), CLng(targetCell.Font.Color), _
                     "Font color for '" & labelName & "' should match between source and target"
    Assert.AreEqual CLng(sourceCell.Interior.Color), CLng(targetCell.Interior.Color), _
                     "Interior color for '" & labelName & "' should match between source and target"
    Assert.AreEqual sourceCell.Font.Bold, targetCell.Font.Bold, _
                     "Bold formatting for '" & labelName & "' should match between source and target"
End Sub

'-------------------------------------------------------------------------------
' Test lifecycle
'-------------------------------------------------------------------------------

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

'===============================================================================
'@section Factory and Creation Tests
'===============================================================================

'@TestMethod("LLFormat")
'@description Creating with an unknown design should use the default design values.
' @sub-title TestCreateFallsBackToDefaultDesign
' @details Tests that LLFormat.Create gracefully handles an unrecognised design name
'   by falling back to the default design. Arranges by creating an LLFormat with
'   "unknown design", then compares the analysis base font size returned by the
'   fallback instance against both the fixture expectation and the default-design
'   instance. Asserts that all three values are equal, confirming the fallback
'   path produces identical output to the explicit default.
Public Sub TestCreateFallsBackToDefaultDesign()
    CustomTestSetTitles Assert, "LLFormat", "TestCreateFallsBackToDefaultDesign"
    On Error GoTo TestFail

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

    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestCreateFallsBackToDefaultDesign", Err.Number, Err.Description
End Sub

'===============================================================================
'@section DesignValue Lookup Tests
'===============================================================================

'@TestMethod("LLFormat")
'@description DesignValue should return the configured colour for the default design.
' @sub-title TestDesignValueReturnsConfiguredColour
' @details Verifies that DesignValue returns the interior colour (Long) for a
'   known label when the default returnColor parameter is used. Arranges by
'   reading the "missing font color" from the fixture, acts by calling
'   FormatUnderTest.DesignValue with the same label, and asserts that the returned
'   colour matches the fixture expectation.
Public Sub TestDesignValueReturnsConfiguredColour()
    CustomTestSetTitles Assert, "LLFormat", "TestDesignValueReturnsConfiguredColour"
    On Error GoTo TestFail

    Dim colorValue As Long

    colorValue = RequireNumericLong(FormatUnderTest.DesignValue(LABEL_MISSING_FONT_COLOR), _
                                    "Default colour value for '" & LABEL_MISSING_FONT_COLOR & "'")

    Assert.AreEqual ExpectedDesignColour(LABEL_MISSING_FONT_COLOR), colorValue, _
                     "DesignValue should return configured color for the default design"

    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestDesignValueReturnsConfiguredColour", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description DesignValue should expose the stored numeric value when colour is not requested.
' @sub-title TestDesignValueReturnsCellValue
' @details Verifies that calling DesignValue with returnColor:=False returns the
'   raw numeric cell value rather than its interior colour. Arranges by obtaining
'   the expected font size from the fixture, acts by calling DesignValue with
'   False for the "analysis base font size" label, and asserts numeric equality.
Public Sub TestDesignValueReturnsCellValue()
    CustomTestSetTitles Assert, "LLFormat", "TestDesignValueReturnsCellValue"
    On Error GoTo TestFail

    Dim expectedLong As Long
    Dim actualLong As Long

    expectedLong = RequireNumericLong(ExpectedDesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE), _
                                      "Fixture value for '" & LABEL_ANALYSIS_BASE_FONT_SIZE & "'")
    actualLong = RequireNumericLong(FormatUnderTest.DesignValue(LABEL_ANALYSIS_BASE_FONT_SIZE, False), _
                                    "Design value for '" & LABEL_ANALYSIS_BASE_FONT_SIZE & "'")

    Assert.AreEqual expectedLong, actualLong, _
                     "DesignValue should return the configured numeric value when returnColor is False"

    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestDesignValueReturnsCellValue", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description Missing labels should return fallback values and log a checking entry.
' @sub-title TestDesignValueMissingLabelFallsBackAndLogs
' @details Tests the edge case where a label does not exist in the format table.
'   Arranges by using the standard FormatUnderTest, acts by requesting DesignValue
'   with "missing label" for both colour and numeric modes. Asserts that the colour
'   fallback is vbBlack, the numeric fallback is 0, that HasCheckings returns True,
'   and that the IChecking log contains at least one entry referencing "missing label".
'   This confirms graceful degradation and traceability for unrecognised labels.
Public Sub TestDesignValueMissingLabelFallsBackAndLogs()
    CustomTestSetTitles Assert, "LLFormat", "TestDesignValueMissingLabelFallsBackAndLogs"
    On Error GoTo TestFail

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

    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestDesignValueMissingLabelFallsBackAndLogs", Err.Number, Err.Description
End Sub

'===============================================================================
'@section ApplyFormat Scope Tests
'===============================================================================

'@TestMethod("LLFormat")
'@description Applying the analysis section scope should honour design-driven styling.
' @sub-title TestApplyFormatAnalysisSectionUsesDesignSettings
' @details Tests the AnalysisSection formatting scope, which styles section headers
'   in analysis sheets. Arranges by writing "Section title" into cell G10, acts by
'   calling ApplyFormat with AnalysisSection, then asserts across a 7-column merged
'   region that the font colour matches "table sections font color", bold is True,
'   font size equals the base size plus the +5 section boost, and WrapText is enabled.
'   Covers the compound styling logic that combines design colour, size arithmetic,
'   and cell layout properties.
Public Sub TestApplyFormatAnalysisSectionUsesDesignSettings()
    CustomTestSetTitles Assert, "LLFormat", "TestApplyFormatAnalysisSectionUsesDesignSettings"
    On Error GoTo TestFail

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

    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestApplyFormatAnalysisSectionUsesDesignSettings", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description Applying the analysis one-cell scope should apply missing-value formatting.
' @sub-title TestApplyFormatAnalysisOneCellAppliesMissingColours
' @details Tests the AnalysisOneCell formatting scope, which styles individual cells
'   that represent missing values. Arranges by writing "Missing value" into cell H15,
'   acts by calling ApplyFormat with AnalysisOneCell, then asserts that the font colour
'   matches "missing font color", the interior colour matches "missing interior color",
'   and bold is True. Ensures the missing-value visual cues are faithfully applied from
'   the design configuration.
Public Sub TestApplyFormatAnalysisOneCellAppliesMissingColours()
    CustomTestSetTitles Assert, "LLFormat", "TestApplyFormatAnalysisOneCellAppliesMissingColours"
    On Error GoTo TestFail

    Dim target As Range

    Set target = FormatSheet.Range("H15")
    target.Value = "Missing value"

    FormatUnderTest.ApplyFormat target, AnalysisOneCell

    Assert.AreEqual ExpectedDesignColour(LABEL_MISSING_FONT_COLOR), CLng(target.Font.Color), _
                     "Missing cell font colour should come from the design"
    Assert.AreEqual ExpectedDesignColour("missing interior color"), CLng(target.Interior.Color), _
                     "Missing cell interior colour should come from the design"
    Assert.IsTrue target.Font.Bold, "Missing cell should be bold"

    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestApplyFormatAnalysisOneCellAppliesMissingColours", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description Applying the all-analysis scope should set worksheet font and dimensions.
' @sub-title TestApplyFormatAllAnalysisSheetUsesDesignDimensions
' @details Tests the AllAnalysisSheet formatting scope, which configures sheet-wide
'   defaults. Arranges by creating a temporary worksheet, acts by calling ApplyFormat
'   with AllAnalysisSheet, then asserts that the worksheet-level font size matches the
'   design base size, the first column width matches "default analysis column width",
'   and row 2 height is fixed at 25. Includes defensive defaults (9 for font size,
'   25 for column width) if the fixture returns zero, ensuring the test does not
'   produce false failures on incomplete fixture data. The temporary sheet is deleted
'   in both the success and failure paths.
Public Sub TestApplyFormatAllAnalysisSheetUsesDesignDimensions()
    CustomTestSetTitles Assert, "LLFormat", "TestApplyFormatAllAnalysisSheetUsesDesignDimensions"
    On Error GoTo TestFail

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

    Exit Sub

TestFail:
    On Error Resume Next
    TestHelpers.DeleteWorksheet "LLFormat_AllAnalysis_Test"
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestApplyFormatAllAnalysisSheetUsesDesignDimensions", Err.Number, Err.Description
End Sub

'===============================================================================
'@section Fixture Preparation Tests
'===============================================================================

'@TestMethod("LLFormat")
'@description Preparing a fixture should define the DESIGNTYPE named range with the default.
' @sub-title TestPrepareFixtureDefinesDesignTypeRange
' @details Validates the LLFormatTestFixture.PrepareLLFormatFixture helper itself.
'   Arranges by cleaning up any leftover fixture sheet, acts by calling
'   PrepareLLFormatFixture with a dedicated sheet name, then reads the DESIGNTYPE
'   named range and asserts its value matches the default design name. The fixture
'   sheet is deleted after the assertion. This meta-test ensures that other tests
'   in this module receive a correctly seeded fixture.
Public Sub TestPrepareFixtureDefinesDesignTypeRange()
    CustomTestSetTitles Assert, "LLFormat", "TestPrepareFixtureDefinesDesignTypeRange"
    On Error GoTo TestFail

    Dim sheetName As String
    Dim fixtureSheet As Worksheet
    Dim designValue As String

    sheetName = "LLFormatFixture_DesignRange"
    On Error Resume Next
        LLFormatTestFixture.DeleteLLFormatFixture sheetName, FormatWorkbook
    On Error GoTo TestFail

    Set fixtureSheet = LLFormatTestFixture.PrepareLLFormatFixture(sheetName, FormatWorkbook)
    designValue = CStr(fixtureSheet.Range("DESIGNTYPE").Value)

    Assert.AreEqual FixtureDefaultDesign(), designValue, _
                     "Prepared fixture should seed the design type named range"

    LLFormatTestFixture.DeleteLLFormatFixture sheetName, FormatWorkbook

    Exit Sub

TestFail:
    On Error Resume Next
    LLFormatTestFixture.DeleteLLFormatFixture "LLFormatFixture_DesignRange", FormatWorkbook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestPrepareFixtureDefinesDesignTypeRange", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description Percent scope formatting should enforce a two-decimal percent number format.
' @sub-title TestApplyFormatPercentSetsNumberFormat
' @details Tests the AnalysisPercent formatting scope. Arranges by clearing cell H1
'   and writing 0.25, acts by calling ApplyFormat with AnalysisPercent, then asserts
'   that the NumberFormat property is exactly "0.00%". This confirms that numeric
'   cells intended for percentages receive a consistent, locale-independent format
'   string with two decimal places.
Public Sub TestApplyFormatPercentSetsNumberFormat()
    CustomTestSetTitles Assert, "LLFormat", "TestApplyFormatPercentSetsNumberFormat"
    On Error GoTo TestFail

    Dim target As Range

    Set target = FormatSheet.Range("H1")
    target.Clear
    target.Value = 0.25

    FormatUnderTest.ApplyFormat target, AnalysisPercent

    Assert.AreEqual "0.00%", target.NumberFormat, _
                     "Percent scope should enforce 2 decimal percent format"

    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestApplyFormatPercentSetsNumberFormat", Err.Number, Err.Description
End Sub

'===============================================================================
'@section Import Tests
'===============================================================================

'@TestMethod("LLFormat")
'@description Importing from another sheet should copy font and interior colours for designs.
' @sub-title TestImportCopiesDesignColours
' @details Tests ILLFormat.Import by verifying that design colours are transferred
'   from an import sheet into the active format instance. Arranges by preparing a
'   secondary fixture sheet with green interior and blue font for the "missing font
'   color" label under the secondary design, then sets the DESIGNTYPE to that design.
'   Acts by calling FormatUnderTest.Import with the import sheet. Asserts that the
'   DesignValue colour now returns green (the imported interior colour), the font
'   colour on the format sheet cell is blue, and the DESIGNTYPE named range has
'   switched to the secondary design name.
Public Sub TestImportCopiesDesignColours()
    CustomTestSetTitles Assert, "LLFormat", "TestImportCopiesDesignColours"
    On Error GoTo TestFail

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

    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestImportCopiesDesignColours", Err.Number, Err.Description
End Sub

'===============================================================================
'@section Export Tests
'===============================================================================

'@TestMethod("LLFormat")
'@description Export should create a new worksheet in the target workbook when it does not exist.
' @sub-title TestExportCreatesNewSheetInTargetWorkbook
' @details Tests that Export creates a new worksheet in a fresh target workbook.
'   Arranges by creating an empty workbook via TestHelpers.NewWorkbook, acts by
'   calling FormatUnderTest.Export, then asserts that a worksheet matching the
'   source format sheet name exists in the target and is the last sheet in the
'   workbook. The target workbook is deleted in both success and failure paths.
Public Sub TestExportCreatesNewSheetInTargetWorkbook()
    CustomTestSetTitles Assert, "LLFormat", "TestExportCreatesNewSheetInTargetWorkbook"
    On Error GoTo TestFail

    Dim targetWkb As Workbook
    Dim sourceSheetName As String

    Set targetWkb = TestHelpers.NewWorkbook()
    sourceSheetName = FormatSheet.Name

    FormatUnderTest.Export targetWkb

    Assert.IsTrue TestHelpers.WorksheetExists(sourceSheetName, targetWkb), _
                 "Export should create worksheet in target workbook"
    Assert.AreEqual sourceSheetName, targetWkb.Worksheets(targetWkb.Worksheets.Count).Name, _
                     "Export should add worksheet at the end of the workbook"

    TestHelpers.DeleteWorkbook targetWkb

    Exit Sub

TestFail:
    On Error Resume Next
    TestHelpers.DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportCreatesNewSheetInTargetWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description Export should copy all table data to the target workbook.
' @sub-title TestExportCopiesTableDataToNewWorkbook
' @details Tests that Export faithfully transfers the format table structure. Arranges
'   by referencing the source ListObject and creating a new target workbook, acts by
'   calling Export, then delegates to VerifyTableStructureMatches to assert that
'   column count and row count are identical between source and target. This ensures
'   that no rows or columns are lost during the export process.
Public Sub TestExportCopiesTableDataToNewWorkbook()
    CustomTestSetTitles Assert, "LLFormat", "TestExportCopiesTableDataToNewWorkbook"
    On Error GoTo TestFail

    Dim targetWkb As Workbook
    Dim sourceTable As ListObject
    Dim targetSheet As Worksheet
    Dim targetTable As ListObject

    Set targetWkb = TestHelpers.NewWorkbook()
    Set sourceTable = FormatSheet.ListObjects(1)

    FormatUnderTest.Export targetWkb

    Set targetSheet = targetWkb.Worksheets(FormatSheet.Name)
    Set targetTable = targetSheet.ListObjects(1)

    Call VerifyTableStructureMatches(sourceTable, targetTable, "Export")

    TestHelpers.DeleteWorkbook targetWkb

    Exit Sub

TestFail:
    On Error Resume Next
    TestHelpers.DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportCopiesTableDataToNewWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description Export should copy cell formatting including font and interior colors.
' @sub-title TestExportCopiesFormatTableStyles
' @details Tests that Export preserves cell-level formatting, not just data. Arranges
'   by setting the "missing font color" cell in the default design to red font and
'   green interior, acts by exporting to a new workbook, then locates the same cell
'   in the target and asserts font and interior colours match. This verifies that
'   the export mechanism copies PasteSpecial formatting rather than only values.
Public Sub TestExportCopiesFormatTableStyles()
    CustomTestSetTitles Assert, "LLFormat", "TestExportCopiesFormatTableStyles"
    On Error GoTo TestFail

    Dim targetWkb As Workbook
    Dim sourceCell As Range
    Dim targetSheet As Worksheet
    Dim targetCell As Range
    Dim defaultDesign As String

    defaultDesign = FixtureDefaultDesign()

    Set sourceCell = LLFormatTestFixture.FixtureCell(FormatSheet, LABEL_MISSING_FONT_COLOR, defaultDesign)
    sourceCell.Font.Color = RGB(255, 0, 0)
    sourceCell.Interior.Color = RGB(0, 255, 0)

    Set targetWkb = TestHelpers.NewWorkbook()

    FormatUnderTest.Export targetWkb

    Set targetSheet = targetWkb.Worksheets(FormatSheet.Name)
    Set targetCell = LLFormatTestFixture.FixtureCell(targetSheet, LABEL_MISSING_FONT_COLOR, defaultDesign)

    Assert.AreEqual CLng(RGB(255, 0, 0)), CLng(targetCell.Font.Color), _
                     "Font color should be copied to target"
    Assert.AreEqual CLng(RGB(0, 255, 0)), CLng(targetCell.Interior.Color), _
                     "Interior color should be copied to target"

    TestHelpers.DeleteWorkbook targetWkb

    Exit Sub

TestFail:
    On Error Resume Next
    TestHelpers.DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportCopiesFormatTableStyles", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description Export should create a ListObject in the target worksheet.
' @sub-title TestExportCreatesListObjectInTarget
' @details Tests that Export creates a proper ListObject (structured table) in the
'   target worksheet, not just raw cell data. Arranges by creating a new workbook,
'   acts by calling Export, then asserts that the target sheet has at least one
'   ListObject and that the first ListObject is a valid non-Nothing reference.
'   This is essential because downstream code relies on ListObject methods for
'   column and row access.
Public Sub TestExportCreatesListObjectInTarget()
    CustomTestSetTitles Assert, "LLFormat", "TestExportCreatesListObjectInTarget"
    On Error GoTo TestFail

    Dim targetWkb As Workbook
    Dim targetSheet As Worksheet
    Dim targetTable As ListObject

    Set targetWkb = TestHelpers.NewWorkbook()

    FormatUnderTest.Export targetWkb

    Set targetSheet = targetWkb.Worksheets(FormatSheet.Name)

    Assert.IsTrue targetSheet.ListObjects.Count > 0, _
                 "Export should create at least one ListObject in target"

    Set targetTable = targetSheet.ListObjects(1)
    Assert.ObjectExists targetTable, "ListObject", "Export should create a valid ListObject"

    TestHelpers.DeleteWorkbook targetWkb

    Exit Sub

TestFail:
    On Error Resume Next
    TestHelpers.DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportCreatesListObjectInTarget", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description Export should create a DESIGNTYPE named range in the target worksheet.
' @sub-title TestExportCreatesDesignTypeNamedRange
' @details Tests that Export creates the DESIGNTYPE named range in the target sheet.
'   Arranges by creating a new workbook, acts by calling Export, then attempts to
'   resolve the DESIGNTYPE range on the target sheet. Asserts that the range object
'   is not Nothing. The DESIGNTYPE range is critical for Import to determine which
'   design column to read on subsequent round-trips.
Public Sub TestExportCreatesDesignTypeNamedRange()
    CustomTestSetTitles Assert, "LLFormat", "TestExportCreatesDesignTypeNamedRange"
    On Error GoTo TestFail

    Dim targetWkb As Workbook
    Dim targetSheet As Worksheet
    Dim designRange As Range

    Set targetWkb = TestHelpers.NewWorkbook()

    FormatUnderTest.Export targetWkb

    Set targetSheet = targetWkb.Worksheets(FormatSheet.Name)

    On Error Resume Next
    Set designRange = targetSheet.Range("DESIGNTYPE")
    On Error GoTo TestFail

    Assert.ObjectExists designRange, "Range", _
                        "Export should create DESIGNTYPE named range in target"

    TestHelpers.DeleteWorkbook targetWkb

    Exit Sub

TestFail:
    On Error Resume Next
    TestHelpers.DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportCreatesDesignTypeNamedRange", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description Export should preserve the design type value in the target worksheet.
' @sub-title TestExportPreservesDesignTypeValue
' @details Tests that the DESIGNTYPE value in the target worksheet matches the
'   source. Arranges by reading the current DESIGNTYPE from the source format sheet
'   and creating a new workbook, acts by calling Export, then reads DESIGNTYPE from
'   the target sheet and asserts string equality. This complements
'   TestExportCreatesDesignTypeNamedRange by verifying the value, not just the
'   existence, of the range.
Public Sub TestExportPreservesDesignTypeValue()
    CustomTestSetTitles Assert, "LLFormat", "TestExportPreservesDesignTypeValue"
    On Error GoTo TestFail

    Dim targetWkb As Workbook
    Dim targetSheet As Worksheet
    Dim sourceDesign As String
    Dim targetDesign As String

    sourceDesign = CStr(FormatSheet.Range("DESIGNTYPE").Value)
    Set targetWkb = TestHelpers.NewWorkbook()

    FormatUnderTest.Export targetWkb

    Set targetSheet = targetWkb.Worksheets(FormatSheet.Name)
    targetDesign = CStr(targetSheet.Range("DESIGNTYPE").Value)

    Assert.AreEqual sourceDesign, targetDesign, _
                     "Export should preserve DESIGNTYPE value in target"

    TestHelpers.DeleteWorkbook targetWkb

    Exit Sub

TestFail:
    On Error Resume Next
    TestHelpers.DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportPreservesDesignTypeValue", Err.Number, Err.Description
End Sub

'@TestMethod("LLFormat")
'@description Export should call Import when target worksheet already exists.
' @sub-title TestExportWithExistingSheetCallsImport
' @details Tests the export fallback path where the target workbook already contains
'   a sheet with the same name. Arranges by creating a target workbook, preparing a
'   fixture sheet in it with the same name as the source, and setting a distinct grey
'   interior colour on the "missing font color" cell. Acts by calling Export, which
'   should detect the existing sheet and delegate to Import. Asserts that the
'   FormatUnderTest now reflects the grey colour from the target, proving that Import
'   was invoked instead of a destructive overwrite.
Public Sub TestExportWithExistingSheetCallsImport()
    CustomTestSetTitles Assert, "LLFormat", "TestExportWithExistingSheetCallsImport"
    On Error GoTo TestFail

    Dim targetWkb As Workbook
    Dim targetSheet As Worksheet
    Dim defaultDesign As String
    Dim colorValue As Long

    defaultDesign = FixtureDefaultDesign()
    Set targetWkb = TestHelpers.NewWorkbook()
    Set targetSheet = LLFormatTestFixture.PrepareLLFormatFixture(FormatSheet.Name, targetWkb)

    With LLFormatTestFixture.FixtureCell(targetSheet, LABEL_MISSING_FONT_COLOR, defaultDesign)
        .Interior.Color = RGB(100, 100, 100)
    End With

    FormatUnderTest.Export targetWkb

    colorValue = RequireNumericLong(FormatUnderTest.DesignValue(LABEL_MISSING_FONT_COLOR), _
                                    "Imported colour value from existing target sheet")
    Assert.AreEqual RGB(100, 100, 100), colorValue, _
                     "Export with existing sheet should Import instead of overwriting"

    TestHelpers.DeleteWorkbook targetWkb

    Exit Sub

TestFail:
    On Error Resume Next
    TestHelpers.DeleteWorkbook targetWkb
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestExportWithExistingSheetCallsImport", Err.Number, Err.Description
End Sub

'===============================================================================
'@section Error Handling Tests
'===============================================================================

'@TestMethod("LLFormat")
'@description Export should throw InvalidArgument error when workbook is Nothing.
' @sub-title TestExportThrowsInvalidArgumentWhenWorkbookIsNothing
' @details Tests the Nothing guard on Export. Arranges using the standard
'   FormatUnderTest, acts by calling Export with Nothing under On Error Resume Next,
'   then captures Err.Number and Err.Description. Asserts that the error number
'   equals ProjectError.InvalidArgument and the description contains the word
'   "workbook". This ensures callers receive a clear, typed error rather than a
'   generic runtime 91 (Object variable not set).
Public Sub TestExportThrowsInvalidArgumentWhenWorkbookIsNothing()
    CustomTestSetTitles Assert, "LLFormat", "TestExportThrowsInvalidArgumentWhenWorkbookIsNothing"
    On Error GoTo TestFail

    On Error Resume Next
    FormatUnderTest.Export Nothing

    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNum, _
                     "Export should throw InvalidArgument when workbook is Nothing"
    Assert.IsTrue InStr(1, errDesc, "workbook", vbTextCompare) > 0, _
                 "Error description should mention workbook"

    Exit Sub

TestFail:
    CustomTestLogFailure Assert, "TestExportThrowsInvalidArgumentWhenWorkbookIsNothing", Err.Number, Err.Description
End Sub
