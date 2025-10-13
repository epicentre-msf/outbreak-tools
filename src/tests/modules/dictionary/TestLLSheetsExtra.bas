Attribute VB_Name = "TestLLSheetsExtra"
Attribute VB_Description = "Additional tests for the LLSheets class"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'@Folder("CustomTests")

'@ModuleDescription("Additional tests for the LLSheets class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const DICT_SHEET As String = "LLSheetsExtraDict"

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
    Assert.SetModuleName "TestLLSheetsExtra"
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

'@TestMethod("LLSheetsExtra")
Public Sub TestContainsRejectsHeaderName()
    CustomTestSetTitles Assert, "LLSheets", "TestContainsRejectsHeaderName"
    On Error GoTo Fail

    Assert.IsFalse Sheets.Contains("Sheet Name"), "Contains should return False when passed the header name"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestContainsRejectsHeaderName", Err.Number, Err.Description
End Sub

'@TestMethod("LLSheetsExtra")
Public Sub TestDataBoundsForBothLayouts()
    CustomTestSetTitles Assert, "LLSheets", "TestDataBoundsForBothLayouts"
    On Error GoTo Fail

    Dim vTop As Long, vBottom As Long, vLeft As Long, vRight As Long
    Dim hTop As Long, hBottom As Long, hLeft As Long, hRight As Long
    Dim vCount As Long, hCount As Long

    vTop = Sheets.DataBounds("vlist1D-sheet1", SheetBound.RowSart)
    vBottom = Sheets.DataBounds("vlist1D-sheet1", SheetBound.RowEnd)
    vLeft = Sheets.DataBounds("vlist1D-sheet1", SheetBound.ColStart)
    vRight = Sheets.DataBounds("vlist1D-sheet1", SheetBound.ColEnd)

    hTop = Sheets.DataBounds("hlist2D-sheet1", SheetBound.RowSart)
    hBottom = Sheets.DataBounds("hlist2D-sheet1", SheetBound.RowEnd)
    hLeft = Sheets.DataBounds("hlist2D-sheet1", SheetBound.ColStart)
    hRight = Sheets.DataBounds("hlist2D-sheet1", SheetBound.ColEnd)

    vCount = Sheets.NumberOfVars("vlist1D-sheet1")
    hCount = Sheets.NumberOfVars("hlist2D-sheet1")

    Assert.AreEqual 4, vTop, "Vertical layout top row should be 4"
    Assert.AreEqual 5, vLeft, "Vertical layout left column should be 5"
    Assert.AreEqual 5, vRight, "Vertical layout right column equals left column"
    Assert.AreEqual vTop + IIf(vCount > 0, vCount - 1, 0), vBottom, _
                     "Vertical bottom should match top + count - 1"

    Assert.AreEqual 8, hTop, "Horizontal layout top row should be 8"
    Assert.AreEqual 1, hLeft, "Horizontal layout left column should be 1"
    Assert.AreEqual 8 + 201, hBottom, "Horizontal layout bottom should be top + 201 rows"
    Assert.AreEqual hLeft + IIf(hCount > 0, hCount - 1, 0), hRight, _
                     "Horizontal right should match left + count - 1"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDataBoundsForBothLayouts", Err.Number, Err.Description
End Sub

'@TestMethod("LLSheetsExtra")
Public Sub TestContainsControlRaisesWhenControlColumnMissing()
    CustomTestSetTitles Assert, "LLSheets", "TestContainsControlRaisesWhenControlColumnMissing"
    On Error GoTo ExpectError

    Dictionary.RemoveColumn "control"
    Dim hasControl As Boolean
    '@Ignore VariableNotUsed, AssignmentNotUsed
    hasControl = Sheets.ContainsControl("vlist1D-sheet1", "formula")
    Assert.LogFailure "ContainsControl should raise when control column is missing"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing control column should raise ElementNotFound"
    Err.Clear
End Sub



'@TestMethod("LLSheetsExtra")
Public Sub TestVariableAddressHorizontalAndVertical()
    CustomTestSetTitles Assert, "LLSheets", "TestVariableAddressHorizontalAndVertical"
    
    On Error GoTo Fail

    'Prepare the dictionary minimally so Prepared() is True
    Dictionary.Prepare

    'Seed column index values for two representative variables
    Dim vars As ILLVariables
    Set vars = LLVariables.Create(Dictionary)

    vars.SetValue "num_valid_h2", "column index", 3 'Horizontal sheet
    vars.SetValue  "choi_v1", "column index", 10     'Vertical sheet

    'When on the same sheet, horizontal address should be relative and omit prefix
    Dim addrH As String
    addrH = Sheets.VariableAddress("num_valid_h2", onSheet:="hlist2D-sheet1")
    Assert.AreEqual "C9", addrH, "Expected horizontal variable address to be B9 given index=3 and top=8"

    'Vertical address should include sheet prefix and be absolute in A1 style
    Dim addrV As String
    addrV = Sheets.VariableAddress("choi_v1")
    Assert.AreEqual "'vlist1D-sheet1'!$E$10", addrV, _
                     "Expected vertical variable address to target column E and supplied row index"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestVariableAddressHorizontalAndVertical", Err.Number, Err.Description
End Sub

