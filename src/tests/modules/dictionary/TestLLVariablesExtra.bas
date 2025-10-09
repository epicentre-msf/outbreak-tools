Attribute VB_Name = "TestLLVariablesExtra"
Attribute VB_Description = "Additional tests for the LLVariables class"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@Folder("Tests")
'@ModuleDescription("Additional tests for the LLVariables class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const DICT_SHEET As String = "LLVarExtraDict"

Private Assert As ICustomTest
Private Dictionary As ILLdictionary
Private Variables As ILLVariables

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLVariablesExtra"
    PrepareDictionaryFixture DICT_SHEET
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet DICT_SHEET
    Set Variables = Nothing
    Set Dictionary = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    PrepareDictionaryFixture DICT_SHEET
    Set Dictionary = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    Set Variables = LLVariables.Create(Dictionary)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set Variables = Nothing
    Set Dictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("LLVariablesExtra")
Public Sub TestContainsReturnsFalseForEmptyName()
    CustomTestSetTitles Assert, "LLVariables", "TestContainsReturnsFalseForEmptyName"
    On Error GoTo Fail

    Assert.IsFalse Variables.Contains(vbNullString), "Contains should return False for empty names"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestContainsReturnsFalseForEmptyName", Err.Number, Err.Description
End Sub

'@TestMethod("LLVariablesExtra")
Public Sub TestValueCaseInsensitiveColumnLookup()
    CustomTestSetTitles Assert, "LLVariables", "TestValueCaseInsensitiveColumnLookup"
    On Error GoTo Fail

    Dim val As String
    val = Variables.Value("main label", "choi_v1")
    Assert.AreEqual "Choices on vlist1D", val, _
                     "Value should resolve headers ignoring case differences"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValueCaseInsensitiveColumnLookup", Err.Number, Err.Description
End Sub

'@TestMethod("LLVariablesExtra")
Public Sub TestCellRangeValidAndInvalid()
    CustomTestSetTitles Assert, "LLVariables", "TestCellRangeValidAndInvalid"
    On Error GoTo Fail

    Dim rng As Range
    Set rng = Variables.CellRange("Dev Comments", "choi_v1")
    Assert.IsTrue (Not rng Is Nothing), "CellRange should return a usable Range for existing values"

    Set rng = Variables.CellRange("Dev Comments", "__unknown__")
    Assert.IsTrue (rng Is Nothing), "CellRange should return Nothing for unknown variables"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCellRangeValidAndInvalid", Err.Number, Err.Description
End Sub

'@TestMethod("LLVariablesExtra")
Public Sub TestSetValueRaisesForUnknownVariable()
    CustomTestSetTitles Assert, "LLVariables", "TestSetValueRaisesForUnknownVariable"
    On Error GoTo ExpectError

    Variables.SetValue "__missing__", "Dev Comments", "value"
    Assert.LogFailure "SetValue should raise when variable is absent"
    Exit Sub

ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing variable should raise ElementNotFound when setting values"
    Err.Clear
End Sub

'@TestMethod("LLVariablesExtra")
Public Sub TestHasCheckingsAfterSkippedSetValue()
    CustomTestSetTitles Assert, "LLVariables", "TestHasCheckingsAfterSkippedSetValue"
    On Error GoTo Fail

    Dim devComments As Range
    Set devComments = Dictionary.DataRange("Dev Comments")
    devComments.Cells(2, 1).Value = "existing"

    Variables.SetValue "choi_v1", "Dev Comments", "new text", onEmpty:=True
    Assert.IsTrue Variables.HasCheckings, "Skipping SetValue should log a warning and create checkings"
    Assert.IsTrue (Not Variables.CheckingValues Is Nothing), "CheckingValues should expose the internal checking object"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestHasCheckingsAfterSkippedSetValue", Err.Number, Err.Description
End Sub

'@TestMethod("LLVariablesExtra")
Public Sub TestResolveColumnIndexCacheInvalidation()
    CustomTestSetTitles Assert, "LLVariables", "TestResolveColumnIndexCacheInvalidation"
    On Error GoTo Fail

    Dim first As String
    Dim colIdx As Long
    Dim sh As Worksheet

    'Warm the cache for Dev Comments
    first = Variables.Value("Dev Comments", "choi_v1")

    'Rename the header to invalidate the cached index
    colIdx = Dictionary.Data.ColumnIndex("Dev Comments", shouldExist:=True, matchCase:=False)
    Set sh = Dictionary.Data.Wksh
    sh.Cells(1, colIdx).Value = "Dev Comments 2"

    'Request the old header again; should not error and should return empty
    Assert.AreEqual vbNullString, Variables.Value("Dev Comments", "choi_v1"), _
                     "Cached column indexes should be validated against current headers"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestResolveColumnIndexCacheInvalidation", Err.Number, Err.Description
End Sub

