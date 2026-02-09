Attribute VB_Name = "TestAnalysisOutput"
Attribute VB_Description = "Tests for AnalysisOutput class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for AnalysisOutput class")

' AnalysisOutput tests focus on factory validation.
' Full integration tests require a complete linelist workbook with all
' analysis setup ListObjects, output worksheets, translation tables,
' formula data, and dictionary — making them unsuitable for unit tests.

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Assert As ICustomTest

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisOutput"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory validation tests
'===============================================================================

'@TestMethod("AnalysisOutput")
Public Sub TestCreateRejectsNothingSpecSheet()
    CustomTestSetTitles Assert, "AnalysisOutput", "TestCreateRejectsNothingSpecSheet"
    On Error GoTo TestFail

    On Error Resume Next
    Dim ao As IAnalysisOutput
    Set ao = AnalysisOutput.Create(Nothing, Nothing)
    On Error GoTo 0

    Assert.IsTrue (ao Is Nothing), _
                  "Create with Nothing specs sheet should fail"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingSpecSheet", Err.Number, Err.Description
End Sub

'@TestMethod("AnalysisOutput")
Public Sub TestCreateRejectsNothingLinelist()
    CustomTestSetTitles Assert, "AnalysisOutput", "TestCreateRejectsNothingLinelist"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = EnsureWorksheet("AOTestSheet", clearSheet:=True, visibility:=xlSheetHidden)

    On Error Resume Next
    Dim ao As IAnalysisOutput
    Set ao = AnalysisOutput.Create(sh, Nothing)
    On Error GoTo 0

    Assert.IsTrue (ao Is Nothing), _
                  "Create with Nothing linelist should fail"

    DeleteWorksheet "AOTestSheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingLinelist", Err.Number, Err.Description
End Sub
