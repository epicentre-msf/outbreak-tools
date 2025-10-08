Attribute VB_Name = "TestLLSheets"
Attribute VB_Description = "Tests for the LLSheets class"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests for the LLSheets class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const DICT_SHEET As String = "LLSheetsDict"
Private Const SHEET_VERTICAL As String = "vlist1D-sheet1"
Private Const SHEET_HORIZONTAL As String = "hlist2D-sheet1"
Private Const KNOWN_VARIABLE As String = "choi_v1"

Private Assert As Object
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
    Set Assert = CreateObject("Rubberduck.AssertClass")
    ResetDictionarySheet
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
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
    Set Sheets = Nothing
    Set Dictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("LLSheets")
Private Sub TestCreateRejectsNullDictionary()
    On Error GoTo ExpectError

    Dim invalid As ILLSheets
    '@Ignore AssignmentNotUsed
    Set invalid = LLSheets.Create(Nothing)
    Assert.Fail "Create should raise when dictionary is Nothing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "Create should flag missing dictionary as ObjectNotInitialized"
    Err.Clear
End Sub

'@TestMethod("LLSheets")
Private Sub TestContainsRecognisesFixtureSheets()
    On Error GoTo Fail

    Assert.IsTrue Sheets.Contains(SHEET_VERTICAL), "Expected fixture sheet to be present"
    Assert.IsTrue Sheets.Contains(SHEET_HORIZONTAL), "Expected horizontal fixture sheet to be present"
    Assert.IsFalse Sheets.Contains("missing-sheet"), "Contains should return False for unknown sheet"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestContainsRecognisesFixtureSheets"
End Sub

'@TestMethod("LLSheets")
Private Sub TestRowIndexReturnsWorksheetRow()
    On Error GoTo Fail

    Dim idx As Long
    idx = Sheets.RowIndex(SHEET_VERTICAL)
    Assert.IsTrue (idx > 0), "RowIndex should return a positive worksheet row"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestRowIndexReturnsWorksheetRow"
End Sub

'@TestMethod("LLSheets")
Private Sub TestDataBoundsRejectsUnknownSelector()
    On Error GoTo ExpectError

    Dim unused As Long
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.DataBounds(SHEET_VERTICAL, 99)
    Assert.Fail "DataBounds should raise for unsupported selectors"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Invalid selectors should return InvalidArgument"
    Err.Clear
End Sub

'@TestMethod("LLSheets")
Private Sub TestSheetInfoRaisesWhenTableColumnMissing()
    On Error GoTo ExpectError

    Dim unused As String
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.SheetInfo(SHEET_VERTICAL, SheetInfoType.SheetInfoSheetTable)
    Assert.Fail "SheetInfo should raise when table name column is missing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing table column should raise ElementNotFound"
    Err.Clear
End Sub

'@TestMethod("LLSheets")
Private Sub TestContainsControlDetectsFormulaControls()
    On Error GoTo Fail

    Assert.IsTrue Sheets.ContainsControl(SHEET_VERTICAL, "formula"), _
                  "Expected the fixture sheet to include formula controls"
    Assert.IsFalse Sheets.ContainsControl(SHEET_VERTICAL, "__missing__"), _
                   "Non-existent control types should return False"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestContainsControlDetectsFormulaControls"
End Sub

'@TestMethod("LLSheets")
Private Sub TestNumberOfVarsRaisesWhenSheetMissing()
    On Error GoTo ExpectError

    Dim unused As Long
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.NumberOfVars("unknown-sheet")
    Assert.Fail "NumberOfVars should raise when the sheet is absent"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing sheets should raise ElementNotFound"
    Err.Clear
End Sub

'@TestMethod("LLSheets")
Private Sub TestVariableAddressRequiresPreparedDictionary()
    On Error GoTo ExpectError

    Dim unused As String
    '@Ignore VariableNotUsed, AssignmentNotUsed
    unused = Sheets.VariableAddress(KNOWN_VARIABLE)
    Assert.Fail "VariableAddress should require a prepared dictionary"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "VariableAddress should signal missing preparation"
    Err.Clear
End Sub

