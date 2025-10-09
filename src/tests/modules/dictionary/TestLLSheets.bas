Attribute VB_Name = "TestLLSheets"
Attribute VB_Description = "Tests for the LLSheets class"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@Folder("CustomTests")
'@Folder("Tests")
'@ModuleDescription("Tests for the LLSheets class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const DICT_SHEET As String = "LLSheetsDict"
Private Const SHEET_VERTICAL As String = "vlist1D-sheet1"
Private Const SHEET_HORIZONTAL As String = "hlist2D-sheet1"
Private Const KNOWN_VARIABLE As String = "choi_v1"

Private Assert As ICustomTest
Private Dictionary As ILLdictionary
Private Sheets As ILLSheets

'@section Fixture lifecycle
'===============================================================================

Private Sub ResetDictionarySheet()
    PrepareDictionaryFixture DICT_SHEET
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLSheets"
    ResetDictionarySheet
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Sheets = Nothing
    Set Dictionary = Nothing
    Set Assert = Nothing
    DeleteWorksheet DICT_SHEET
End Sub

'@TestInitialize
Private Sub TestInitialize()
    ResetDictionarySheet
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    Set Sheets = LLSheets.Create(Dictionary)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set Sheets = Nothing
    Set Dictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("LLSheets")
Public Sub TestCreateRejectsNullDictionary()
    CustomTestSetTitles Assert, "LLSheets", "TestCreateRejectsNullDictionary"
    On Error GoTo ExpectError

    Dim invalid As ILLSheets
    '@Ignore AssignmentNotUsed
    Set invalid = LLSheets.Create(Nothing)
    Assert.LogFailure "Create should raise when dictionary is Nothing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "Create should flag missing dictionary as ObjectNotInitialized"
    Err.Clear
End Sub

'@TestMethod("LLSheets")
Public Sub TestContainsRecognisesFixtureSheets()
    CustomTestSetTitles Assert, "LLSheets", "TestContainsRecognisesFixtureSheets"
    On Error GoTo Fail

    Assert.IsTrue Sheets.Contains(SHEET_VERTICAL), "Expected fixture sheet to be present"
    Assert.IsTrue Sheets.Contains(SHEET_HORIZONTAL), "Expected horizontal fixture sheet to be present"
    Assert.IsFalse Sheets.Contains("missing-sheet"), "Contains should return False for unknown sheet"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestContainsRecognisesFixtureSheets", Err.Number, Err.Description
End Sub

'@TestMethod("LLSheets")
Public Sub TestRowIndexReturnsWorksheetRow()
    CustomTestSetTitles Assert, "LLSheets", "TestRowIndexReturnsWorksheetRow"
    On Error GoTo Fail

    Dim idx As Long
    idx = Sheets.RowIndex(SHEET_VERTICAL)
    Assert.IsTrue (idx > 0), "RowIndex should return a positive worksheet row"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRowIndexReturnsWorksheetRow", Err.Number, Err.Description
End Sub

'@TestMethod("LLSheets")
Public Sub TestDataBoundsRejectsUnknownSelector()
    CustomTestSetTitles Assert, "LLSheets", "TestDataBoundsRejectsUnknownSelector"
    On Error GoTo ExpectError

    Dim unused As Long
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.DataBounds(SHEET_VERTICAL, 99)
    Assert.LogFailure "DataBounds should raise for unsupported selectors"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Invalid selectors should return InvalidArgument - Description " & Err.Description
    Err.Clear
End Sub

'@TestMethod("LLSheets")
Public Sub TestSheetInfoRaisesWhenTableColumnMissing()
    CustomTestSetTitles Assert, "LLSheets", "TestSheetInfoRaisesWhenTableColumnMissing"
    On Error GoTo ExpectError

    Dim unused As String
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.SheetInfo(SHEET_VERTICAL, SheetInfoType.SheetInfoSheetTable)
    Assert.LogFailure "SheetInfo should raise when table name column is missing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing table column should raise ElementNotFound"
    Err.Clear
End Sub

'@TestMethod("LLSheets")
Public Sub TestContainsControlDetectsFormulaControls()
    CustomTestSetTitles Assert, "LLSheets", "TestContainsControlDetectsFormulaControls"
    On Error GoTo Fail

    Assert.IsTrue Sheets.ContainsControl(SHEET_VERTICAL, "formula", colName:="Control"), _
                  "Expected the fixture sheet to include formula controls"
    Assert.IsFalse Sheets.ContainsControl(SHEET_VERTICAL, "__missing__"), _
                   "Non-existent control types should return False"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestContainsControlDetectsFormulaControls", Err.Number, Err.Description
End Sub

'@TestMethod("LLSheets")
Public Sub TestNumberOfVarsRaisesWhenSheetMissing()
    CustomTestSetTitles Assert, "LLSheets", "TestNumberOfVarsRaisesWhenSheetMissing"
    On Error GoTo ExpectError

    Dim unused As Long
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.NumberOfVars("unknown-sheet")
    Assert.LogFailure "NumberOfVars should raise when the sheet is absent"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing sheets should raise ElementNotFound"
    Err.Clear
End Sub

'@TestMethod("LLSheets")
Public Sub TestVariableAddressRequiresPreparedDictionary()
    CustomTestSetTitles Assert, "LLSheets", "TestVariableAddressRequiresPreparedDictionary"
    On Error GoTo ExpectError

    Dim unused As String
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.VariableAddress(KNOWN_VARIABLE)
    Assert.LogFailure "VariableAddress should require a prepared dictionary"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "VariableAddress should signal missing preparation"
    Err.Clear
End Sub
