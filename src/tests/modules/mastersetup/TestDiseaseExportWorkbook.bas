Attribute VB_Name = "TestDiseaseExportWorkbook"
Attribute VB_Description = "Tests verifying DiseaseExportWorkbook manages workbook lifecycle safely"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests verifying DiseaseExportWorkbook manages workbook lifecycle safely")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const EXPORT_TEST_FILE As String = "disease_export_test.xlsx"

Private Assert As ICustomTest
Private Manager As IDiseaseExportWorkbook

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseExportWorkbook"
    Set Manager = New DiseaseExportWorkbook
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        If Not Manager Is Nothing Then
            Manager.Release
        End If
        DeleteTestFile
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
    Set Manager = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    DeleteTestFile
    If Not Manager Is Nothing Then
        Manager.Release
    End If
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Manager Is Nothing Then
        Manager.Release
    End If
    DeleteTestFile
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseExportWorkbook")
Public Sub TestEnsureSaveAndCloseReleasesWorkbook()
    CustomTestSetTitles Assert, "DiseaseExportWorkbook", "TestEnsureSaveAndCloseReleasesWorkbook"

    Dim wb As Workbook
    Dim exportPath As String

    On Error GoTo Fail

    exportPath = BuildTestFilePath()

    Set wb = Manager.EnsureWorkbook(Application)
    Assert.IsFalse wb Is Nothing, "EnsureWorkbook should return a workbook instance"
    Assert.IsTrue Manager.HasWorkbook, "Manager should report active workbook"

    Manager.SaveAs exportPath
    Assert.IsTrue FileExists(exportPath), "SaveAs should create the export file"

    Manager.Close saveChanges:=False
    Assert.IsFalse Manager.HasWorkbook, "Manager should release workbook after Close"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEnsureSaveAndCloseReleasesWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseExportWorkbook")
Public Sub TestSaveWithoutWorkbookRaises()
    CustomTestSetTitles Assert, "DiseaseExportWorkbook", "TestSaveWithoutWorkbookRaises"

    Dim raisedError As Boolean

    On Error Resume Next
        Manager.Release
        Manager.SaveAs BuildTestFilePath()
        raisedError = (Err.Number = ProjectError.ObjectNotInitialized)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "SaveAs should raise when workbook is missing"
End Sub

'@TestMethod("DiseaseExportWorkbook")
Public Sub TestEnsureAfterCloseCreatesFreshWorkbook()
    CustomTestSetTitles Assert, "DiseaseExportWorkbook", "TestEnsureAfterCloseCreatesFreshWorkbook"

    Dim firstWorkbook As Workbook
    Dim secondWorkbook As Workbook

    On Error GoTo Fail

    Set firstWorkbook = Manager.EnsureWorkbook(Application)
    Manager.Close saveChanges:=False
    Set secondWorkbook = Manager.EnsureWorkbook(Application)

    Assert.IsFalse firstWorkbook Is secondWorkbook, "EnsureWorkbook should create a fresh workbook after Close"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEnsureAfterCloseCreatesFreshWorkbook", Err.Number, Err.Description
End Sub

'@section Helpers
'===============================================================================

Private Function BuildTestFilePath() As String
    BuildTestFilePath = ThisWorkbook.Path & Application.PathSeparator & "temp" & _
                        Application.PathSeparator & EXPORT_TEST_FILE
End Function

Private Sub DeleteTestFile()
    Dim path As String
    path = BuildTestFilePath()

    On Error Resume Next
        If FileExists(path) Then Kill path
    On Error GoTo 0
End Sub

Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (LenB(Dir(filePath)) > 0)
End Function
