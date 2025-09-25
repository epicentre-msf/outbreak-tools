Attribute VB_Name = "TestFormulaData"

Option Explicit
Option Private Module

'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")

Private Const FORMULA_SHEET As String = "FormulaDataFixture"
Private Const FORMULAS_TABLE_NAME As String = "T_XlsFonctions"
Private Const CHARACTERS_TABLE_NAME As String = "T_ascii"

Private Assert As Object
Private Fakes As Object
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
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet FORMULA_SHEET
    RestoreApp
    Set Assert = Nothing
    Set Fakes = Nothing
    Set FixtureSheet = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    PrepareFixtureSheet
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set FixtureSheet = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("FormulaData")
Private Sub TestCreateCachesLookups()
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
    FailUnexpectedError Assert, "TestCreateCachesLookups"
End Sub

'@TestMethod("FormulaData")
Private Sub TestCreateRequiresWorksheet()
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
Private Sub TestMissingTableRaises()
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
    FailUnexpectedError Assert, "TestMissingTableRaises"
End Sub

'@TestMethod("FormulaData")
Private Sub TestCachesSurviveWorksheetChanges()
    Dim formData As IFormulaData
    Dim formulaTable As ListObject

    On Error GoTo Fail

    Set formData = BuildFormulaData()
    Set formulaTable = FixtureSheet.ListObjects(FORMULAS_TABLE_NAME)

    formulaTable.DataBodyRange.ClearContents

    Assert.IsTrue formData.ExcelFormulasIncludes("SUM"), "Lookup should rely on cached values"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestCachesSurviveWorksheetChanges"
End Sub

