Attribute VB_Name = "TestSetupImportService"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Unit tests covering the improved setup import service")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private Service As ISetupImportService
Private ProgressStub As ProgressDisplayStub
Private PasswordsHandler As IPasswords

Private Const PASSWORD_SHEET As String = "TST_SetupImport_Passwords"
Private Const CLEAN_TARGET_SHEET As String = "TST_SetupImport_Clean"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Private Sub TestInitialize()
    Set ProgressStub = New ProgressDisplayStub
    Set Service = New SetupImportService
    Service.Path = ThisWorkbook.FullName
    Set Service.ProgressObject = ProgressStub
    EnsurePasswordsFixture
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Service = Nothing
    Set ProgressStub = Nothing
    Set PasswordsHandler = Nothing
    TestHelpers.DeleteWorksheet CLEAN_TARGET_SHEET
    TestHelpers.DeleteWorksheet PASSWORD_SHEET
End Sub


'@section Tests
'===============================================================================
'@TestMethod("SetupImportService")
Private Sub TestCheckRaisesWhenNoSelection()
    On Error GoTo ExpectInvalid

    Service.Check False, False, False, False, False
    Assert.Fail "Check should raise when no import option is selected."
    Exit Sub

ExpectInvalid:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, "Unexpected error code."
    Assert.AreEqual "Please select at least one import option (Dictionary, Choices, Exports, Analysis or Translations).", _
                    ProgressStub.Caption, "Expected message to be surfaced through the progress display."
    Err.Clear
End Sub

'@TestMethod("SetupImportService")
Private Sub TestCheckRaisesWhenFileMissing()
    Dim missingPath As String

    missingPath = BuildMissingSetupPath()
    Service.Path = missingPath

    On Error GoTo ExpectMissing
        Service.Check True, False, False, False, False
        Assert.Fail "Check should raise when the source workbook cannot be located."
        Exit Sub

ExpectMissing:
    Assert.AreEqual CLng(ProjectError.ElementNotFound), Err.Number, "Unexpected error code when file is missing."
    Assert.IsTrue InStr(1, ProgressStub.Caption, missingPath, vbTextCompare) > 0, _
                   "Progress display should include the missing path."
    Err.Clear
End Sub

'@TestMethod("SetupImportService")
Private Sub TestCleanRemovesWorksheetComments()
    Dim targetSheet As Worksheet
    Dim sheetsList As BetterArray

    Set targetSheet = TestHelpers.EnsureWorksheet(CLEAN_TARGET_SHEET)
    PrepareComment targetSheet

    Set sheetsList = SheetsListOf(CLEAN_TARGET_SHEET)
    Service.Clean PasswordsHandler, sheetsList

    Assert.IsTrue targetSheet.Cells(1, 1).Comment Is Nothing, "Clean should remove classic comments."
End Sub

'@TestMethod("SetupImportService")
Private Sub TestImportClosesWorkbookAfterRun()
    Dim tempBook As Workbook
    Dim exportFolder As String
    Dim workbookPath As String
    Dim sheetsList As BetterArray
    Dim workbookName As String

    Set tempBook = TestHelpers.NewWorkbook
    tempBook.Worksheets(1).Name = "TempData"

    exportFolder = TestHelpers.BuildTempFolder(ThisWorkbook, "SetupImportTests")
    workbookPath = TestHelpers.BuildWorkbookPath(exportFolder, "setup_import_source", ".xlsx")
    tempBook.SaveAs Filename:=workbookPath, FileFormat:=xlOpenXMLWorkbook
    tempBook.Close SaveChanges:=False

    workbookName = FileNameFromPath(workbookPath)
    Service.Path = workbookPath
    Set sheetsList = SheetsListOf("MissingSheet")

    Service.Import PasswordsHandler, sheetsList
    Assert.IsFalse IsWorkbookOpen(workbookName), "Import should close the source workbook on completion."

    'Calling Import again should reopen and close the workbook without errors.
    Service.Import PasswordsHandler, sheetsList
    Assert.IsFalse IsWorkbookOpen(workbookName), "Import should leave no lingering workbook reference."

    DeleteFileIfExists workbookPath
End Sub


'@section Helpers
'===============================================================================
Private Sub EnsurePasswordsFixture()
    Dim passwordSheet As Worksheet

    PasswordsTestFixture.PreparePasswordsFixture PASSWORD_SHEET, ThisWorkbook
    Set passwordSheet = ThisWorkbook.Worksheets(PASSWORD_SHEET)
    Set PasswordsHandler = Passwords.Create(passwordSheet)
End Sub

Private Sub PrepareComment(ByVal targetSheet As Worksheet)
    On Error Resume Next
        targetSheet.Cells(1, 1).ClearComments
        targetSheet.Cells(1, 1).ClearCommentsThreaded
    On Error GoTo 0

    targetSheet.Cells(1, 1).Value = "Sample"
    targetSheet.Cells(1, 1).AddComment "Temporary note"
End Sub

Private Function SheetsListOf(ParamArray sheetNames() As Variant) As BetterArray
    Dim list As BetterArray
    Dim idx As Long

    Set list = New BetterArray
    list.LowerBound = 1

    For idx = LBound(sheetNames) To UBound(sheetNames)
        list.Push CStr(sheetNames(idx))
    Next idx

    Set SheetsListOf = list
End Function

Private Function BuildMissingSetupPath() As String
    Dim baseFolder As String

    baseFolder = ThisWorkbook.Path
    If LenB(baseFolder) = 0 Then baseFolder = CurDir$

    BuildMissingSetupPath = baseFolder & Application.PathSeparator & "missing_setup_source.xlsx"
End Function

Private Function IsWorkbookOpen(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Function FileNameFromPath(ByVal filePath As String) As String
    Dim separatorPos As Long

    separatorPos = InStrRev(filePath, Application.PathSeparator)
    If separatorPos = 0 Then
        FileNameFromPath = filePath
    Else
        FileNameFromPath = Mid$(filePath, separatorPos + 1)
    End If
End Function

Private Sub DeleteFileIfExists(ByVal filePath As String)
    If LenB(Dir$(filePath)) = 0 Then Exit Sub

    On Error Resume Next
        Kill filePath
    On Error GoTo 0
End Sub
