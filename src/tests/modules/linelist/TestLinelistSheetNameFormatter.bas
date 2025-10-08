Attribute VB_Name = "TestLinelistSheetNameFormatter"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Formatter As ILinelistSheetNameFormatter
Private FixtureWorkbook As Workbook


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelistSheetNameFormatter"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
    RestoreApp
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set Formatter = LinelistSheetNameFormatter.Create
    Set FixtureWorkbook = TestHelpers.NewWorkbook
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    If Not FixtureWorkbook Is Nothing Then
        BusyApp
        TestHelpers.DeleteWorkbook FixtureWorkbook
        Set FixtureWorkbook = Nothing
    End If

    Set Formatter = Nothing
End Sub


'@section Tests
'===============================================================================
'@TestMethod("LinelistSheetNameFormatter")
Public Sub TestFormatSheetNameAppliesPrefixAndTruncation()
    CustomTestSetTitles Assert, "LinelistSheetNameFormatter", "FormatAppliesPrefixAndTruncation"

    Dim baseName As String
    Dim formatted As String

    baseName = String(50, "A")
    formatted = Formatter.FormatSheetName(baseName, sheetScopePrint)

    Assert.AreEqual Formatter.PrintPrefix, Left$(formatted, Len(Formatter.PrintPrefix)), _
                     "Print scope should prepend the print prefix"
    Assert.IsTrue Len(formatted) <= Formatter.SheetNameMaxLength, _
                   "Formatted sheet name should respect the maximum length"
End Sub

'@TestMethod("LinelistSheetNameFormatter")
Public Sub TestSheetExistsRecognisesScopedSheet()
    CustomTestSetTitles Assert, "LinelistSheetNameFormatter", "SheetExistsRecognisesScopedSheet"

    Dim scopedName As String
    scopedName = Formatter.FormatSheetName("Analysis Long Sheet Name", sheetScopeCrf)

    TestHelpers.EnsureWorksheet scopedName, FixtureWorkbook

    Assert.IsTrue Formatter.SheetExists(FixtureWorkbook, "Analysis Long Sheet Name", sheetScopeCrf), _
                   "SheetExists should match against the scoped sheet name"
End Sub

'@TestMethod("LinelistSheetNameFormatter")
Public Sub TestResolveWorksheetRaisesWhenMissing()
    CustomTestSetTitles Assert, "LinelistSheetNameFormatter", "ResolveWorksheetRaisesWhenMissing"

    On Error GoTo ExpectError
        Formatter.ResolveWorksheet FixtureWorkbook, "Missing", sheetScopeStandard
        Assert.Fail "ResolveWorksheet should raise when the sheet is absent"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "ResolveWorksheet should raise ElementNotFound when the sheet is missing"
    Err.Clear
End Sub

'@TestMethod("LinelistSheetNameFormatter")
Public Sub TestFormatSheetNameRejectsEmpty()
    CustomTestSetTitles Assert, "LinelistSheetNameFormatter", "FormatRejectsEmpty"

    On Error GoTo ExpectError
        Formatter.FormatSheetName vbNullString
        Assert.Fail "FormatSheetName should reject empty base names"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "FormatSheetName should raise InvalidArgument for empty names"
    Err.Clear
End Sub

'@TestMethod("LinelistSheetNameFormatter")
Public Sub TestInvalidScopeRaises()
    CustomTestSetTitles Assert, "LinelistSheetNameFormatter", "InvalidScopeRaises"

    On Error GoTo ExpectError
        Formatter.FormatSheetName "Sheet", CLng(99)
        Assert.Fail "Passing an invalid scope should raise an error"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Invalid scope should raise InvalidArgument"
    Err.Clear
End Sub
