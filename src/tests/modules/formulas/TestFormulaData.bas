Attribute VB_Name = "TestFormulaData"
Attribute VB_Description = "Unit tests ensuring FormulaData caches metadata and guards missing fixtures"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests ensuring FormulaData caches metadata and guards missing fixtures")

'@description
'Tests the FormulaData class, which loads and caches Excel formula function
'names, special separator characters, and grouped-function metadata from a
'worksheet containing T_XlsFonctions and T_ascii tables. Coverage includes
'successful cache initialisation (case-insensitive lookups for both functions
'and characters), factory guard clauses (Nothing worksheet, missing table),
'cache durability after worksheet mutation, and grouped-function metadata
'(aggregator mapping, native-function flags). Each test builds a minimal
'fixture sheet with two ListObjects seeded in PrepareFixtureSheet.
'@depends FormulaData, IFormulaData, CustomTest, ICustomTest, TestHelpers

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Const FORMULA_SHEET As String = "FormulaDataFixture"
Private Const FORMULAS_TABLE_NAME As String = "T_XlsFonctions"
Private Const CHARACTERS_TABLE_NAME As String = "T_ascii"

Private Assert As ICustomTest
Private FixtureSheet As Worksheet

'@section Helpers
'===============================================================================

'@sub-title Build the fixture worksheet with formula functions and characters tables
'@details
'Creates (or clears) the fixture worksheet, then populates two ListObjects:
'T_XlsFonctions with three sample function names (SUM, AVERAGE, IF) under
'an "ENG" header, and T_ascii with three ASCII separator entries (+, -, /)
'under "ASCII" and "TEXT" headers. The tables are placed side by side
'starting at columns A and C respectively.
Private Sub PrepareFixtureSheet()
    Dim functionRows As Variant
    Dim functionMatrix As Variant
    Dim characterRows As Variant
    Dim characterMatrix As Variant
    Dim formulaRange As Range
    Dim characterRange As Range
    Dim formulaTable As ListObject
    Dim characterTable As ListObject

    Set FixtureSheet = EnsureWorksheet(FORMULA_SHEET)
    ClearWorksheet FixtureSheet

    functionRows = Array(Array("ENG"), Array("SUM"), Array("AVERAGE"), Array("IF"))
    functionMatrix = RowsToMatrix(functionRows)
    WriteMatrix FixtureSheet.Range("A1"), functionMatrix
    Set formulaRange = FixtureSheet.Range("A1").Resize(UBound(functionMatrix, 1), UBound(functionMatrix, 2))
    Set formulaTable = FixtureSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                   Source:=formulaRange, XlListObjectHasHeaders:=xlYes)
    formulaTable.Name = FORMULAS_TABLE_NAME

    characterRows = Array( _
        Array("ASCII", "TEXT"), _
        Array(43, "+"), _
        Array(45, "-"), _
        Array(47, "/") _
    )
    characterMatrix = RowsToMatrix(characterRows)
    WriteMatrix FixtureSheet.Range("C1"), characterMatrix
    Set characterRange = FixtureSheet.Range("C1").Resize(UBound(characterMatrix, 1), UBound(characterMatrix, 2))
    Set characterTable = FixtureSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                     Source:=characterRange, XlListObjectHasHeaders:=xlYes)
    characterTable.Name = CHARACTERS_TABLE_NAME
End Sub

'@sub-title Create a FormulaData instance from the current fixture sheet
Private Function BuildFormulaData() As IFormulaData
    Set BuildFormulaData = FormulaData.Create(FixtureSheet)
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Initialise the test harness and suppress screen updates
'@details
'Calls BusyApp to suppress screen updates for performance, creates the
'test output sheet, and sets up the CustomTest assertion object.
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestFormulaData"
End Sub

'@sub-title Print results, clean up the fixture sheet, and restore application state
'@details
'Flushes remaining assertion output to the test sheet, deletes the formula
'fixture worksheet, restores Excel application settings via RestoreApp,
'and releases all object references.
'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet FORMULA_SHEET
    RestoreApp
    Set Assert = Nothing
    Set FixtureSheet = Nothing
End Sub

'@sub-title Rebuild the fixture sheet before each test
'@details
'Suppresses screen updates and recreates the fixture worksheet with fresh
'T_XlsFonctions and T_ascii tables, ensuring each test starts from a
'clean and predictable state.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    PrepareFixtureSheet
End Sub

'@sub-title Flush assertions and release the fixture sheet reference after each test
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set FixtureSheet = Nothing
End Sub

'@section Tests
'===============================================================================

'@sub-title Verify Create caches formula functions and special characters for lookup
'@details
'Creates a FormulaData instance from the fixture sheet and verifies
'ExcelFormulasIncludes returns True for "SUM" (exact case), True for
'"average" (case-insensitive), and False for "UNKNOWN_FUNC". Also verifies
'SpecialCharacterIncludes returns True for "+" and False for "#", confirming
'both caches are populated at creation time.
'@TestMethod("FormulaData")
Public Sub TestCreateCachesLookups()
    CustomTestSetTitles Assert, "FormulaData", "TestCreateCachesLookups"
    Dim formData As IFormulaData

    On Error GoTo Fail

    Set formData = BuildFormulaData()

    Assert.IsTrue formData.ExcelFormulasIncludes("SUM"), "Known function should be recognised"
    Assert.IsTrue formData.ExcelFormulasIncludes("average"), "Lookup should be case insensitive"
    Assert.IsFalse formData.ExcelFormulasIncludes("UNKNOWN_FUNC"), "Unexpected function should be rejected"

    Assert.IsTrue formData.SpecialCharacterIncludes("+"), "Known character should be recognised"
    Assert.IsFalse formData.SpecialCharacterIncludes("#"), "Unexpected character should be rejected"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateCachesLookups", Err.Number, Err.Description
End Sub

'@sub-title Verify Create raises when given a Nothing worksheet
'@details
'Calls FormulaData.Create with Nothing and asserts that an
'ObjectNotInitialized error is raised, confirming the factory guard clause
'rejects a missing worksheet reference.
'@TestMethod("FormulaData")
Public Sub TestCreateRequiresWorksheet()
    CustomTestSetTitles Assert, "FormulaData", "TestCreateRequiresWorksheet"
    Dim raisedError As Boolean
    Dim formData As IFormulaData

    On Error Resume Next
        '@Ignore AssignmentNotUsed
        Set formData = FormulaData.Create(Nothing)
        raisedError = (Err.Number = ProjectError.ObjectNotInitialized)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Create should raise when worksheet is missing"
    Set formData = Nothing
End Sub

'@sub-title Verify Create raises when the formulas table is absent from the worksheet
'@details
'Deletes the T_XlsFonctions ListObject from the fixture sheet before
'calling FormulaData.Create. Asserts that an ElementNotFound error is
'raised, confirming the factory detects missing required tables.
'@TestMethod("FormulaData")
Public Sub TestMissingTableRaises()
    CustomTestSetTitles Assert, "FormulaData", "TestMissingTableRaises"
    Dim raisedError As Boolean
    Dim formData As IFormulaData

    On Error GoTo Fail

    FixtureSheet.ListObjects(FORMULAS_TABLE_NAME).Delete

    On Error Resume Next
        '@Ignore AssignmentNotUsed
        Set formData = FormulaData.Create(FixtureSheet)
        raisedError = (Err.Number = ProjectError.ElementNotFound)
        Err.Clear
    On Error GoTo Fail

    Assert.IsTrue raisedError, "Create should raise when formulas table is missing"
    Set formData = Nothing
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMissingTableRaises", Err.Number, Err.Description
End Sub

'@sub-title Verify cached lookups survive after the source worksheet data is cleared
'@details
'Creates a FormulaData instance, then clears the data body of the
'T_XlsFonctions table on the worksheet. Asserts that ExcelFormulasIncludes
'still returns True for "SUM", confirming the class relies on its
'in-memory cache rather than re-reading the worksheet on each call.
'@TestMethod("FormulaData")
Public Sub TestCachesSurviveWorksheetChanges()
    CustomTestSetTitles Assert, "FormulaData", "TestCachesSurviveWorksheetChanges"
    Dim formData As IFormulaData
    Dim formulaTable As ListObject

    On Error GoTo Fail

    Set formData = BuildFormulaData()
    Set formulaTable = FixtureSheet.ListObjects(FORMULAS_TABLE_NAME)

    formulaTable.DataBodyRange.ClearContents

    Assert.IsTrue formData.ExcelFormulasIncludes("SUM"), "Lookup should rely on cached values"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCachesSurviveWorksheetChanges", Err.Number, Err.Description
End Sub

'@sub-title Verify grouped-function metadata: aggregator mapping, native flags, and lookup integration
'@details
'Creates a FormulaData instance and exercises the grouped-function API.
'Asserts IsGroupFunction returns True for SUMIFS, MEANIFS (case-insensitive),
'NIFS, and MINIFS. Verifies GroupAggregator maps SUMIFS to "SUMIFS",
'MEANIFS to "AVERAGE", NIFS to "COUNTIFS", and MINIFS to "MIN". Checks
'GroupUsesNativeFunction returns True for SUMIFS, COUNTIFS, and NIFS (which
'emit native *IFS calls) and False for MEANIFS and MINIFS (which require
'IF wrappers). Finally confirms unknown tokens return False, vbNullString,
'and False respectively, and that grouped tokens and their aggregators are
'registered in the ExcelFormulasIncludes lookup.
'@TestMethod("FormulaData")
Public Sub TestGroupedFunctionMetadata()
    CustomTestSetTitles Assert, "FormulaData", "TestGroupedFunctionMetadata"
    Dim formData As IFormulaData

    On Error GoTo Fail

    Set formData = BuildFormulaData()

    Assert.IsTrue formData.IsGroupFunction("SUMIFS"), "SUMIFS should be registered as grouped function"
    Assert.IsTrue formData.IsGroupFunction("meanifs"), "MEANIFS lookup should be case insensitive"
    Assert.IsTrue formData.IsGroupFunction("nifs"), "NIFS alias should be recognised"
    Assert.IsTrue formData.IsGroupFunction("minifs"), "MINIFS should be registered"
    Assert.AreEqual "SUMIFS", formData.GroupAggregator("SUMIFS"), "SUMIFS should map to SUMIFS"
    Assert.AreEqual "AVERAGE", formData.GroupAggregator("meanifs"), "MEANIFS should map to AVERAGE"
    Assert.AreEqual "COUNTIFS", formData.GroupAggregator("NIFS"), "NIFS should map to COUNTIFS"
    Assert.AreEqual "MIN", formData.GroupAggregator("MINIFS"), "MINIFS should map to MIN aggregator"

    Assert.IsTrue formData.GroupUsesNativeFunction("SUMIFS"), "SUMIFS should emit native SUMIFS"
    Assert.IsTrue formData.GroupUsesNativeFunction("COUNTIFS"), "COUNTIFS should emit native COUNTIFS"
    Assert.IsTrue formData.GroupUsesNativeFunction("NIFS"), "NIFS should emit native COUNTIFS"
    Assert.IsFalse formData.GroupUsesNativeFunction("MEANIFS"), "MEANIFS should require IF wrapper"
    Assert.IsFalse formData.GroupUsesNativeFunction("MINIFS"), "MINIFS should require IF wrapper"

    Assert.IsFalse formData.IsGroupFunction("UNKNOWN_GROUP"), "Unknown tokens should not be registered"
    Assert.AreEqual vbNullString, formData.GroupAggregator("UNKNOWN_GROUP"), "Unknown tokens should return empty aggregator"
    Assert.IsFalse formData.GroupUsesNativeFunction("UNKNOWN_GROUP"), "Unknown tokens should not use native aggregator"

    Assert.IsTrue formData.ExcelFormulasIncludes("SUMIFS"), "Grouped token should be available in Excel lookup"
    Assert.IsTrue formData.ExcelFormulasIncludes("COUNTIFS"), "Aggregator tokens should be ensured in Excel lookup"
    Assert.IsTrue formData.ExcelFormulasIncludes("MIN"), "Non-native aggregator tokens should be ensured in Excel lookup"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestGroupedFunctionMetadata", Err.Number, Err.Description
End Sub
