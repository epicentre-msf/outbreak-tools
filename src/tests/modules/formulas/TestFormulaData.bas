Attribute VB_Name = "TestFormulaData"
Attribute VB_Description = "Unit tests ensuring FormulaData caches metadata and guards missing fixtures"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests ensuring FormulaData caches metadata and guards missing fixtures")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Const FORMULA_SHEET As String = "FormulaDataFixture"
Private Const FORMULAS_TABLE_NAME As String = "T_XlsFonctions"
Private Const CHARACTERS_TABLE_NAME As String = "T_ascii"

Private Assert As ICustomTest
Private FixtureSheet As Worksheet

'@section Helpers
'===============================================================================

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

Private Function BuildFormulaData() As IFormulaData
    Set BuildFormulaData = FormulaData.Create(FixtureSheet)
End Function

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestFormulaData"
End Sub

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

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    PrepareFixtureSheet
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set FixtureSheet = Nothing
End Sub

'@section Tests
'===============================================================================

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
