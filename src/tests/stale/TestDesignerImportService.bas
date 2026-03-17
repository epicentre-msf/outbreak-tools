Attribute VB_Name = "TestDesignerImportService"
Attribute VB_Description = "Unit tests for DesignerImportService class"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates DesignerImportService factory, configuration, input validation, and exported property defaults.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    TestHelpers.BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDesignerImportService"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    TestHelpers.RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    TestHelpers.BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        TestHelpers.DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set FixtureWorkbook = Nothing
    TestHelpers.RestoreApp
End Sub


'@section Factory and Configuration Tests
'===============================================================================

'@TestMethod("DesignerImportService")
Public Sub TestCreateDefaultsToThisWorkbook()
    CustomTestSetTitles Assert, "DesignerImportService", "TestCreateDefaultsToThisWorkbook"
    On Error GoTo Fail

    Dim subject As IDesignerImportService
    Set subject = DesignerImportService.Create()

    Assert.AreEqual ThisWorkbook.Name, subject.HostWorkbook.Name, _
                    "Default host should be ThisWorkbook."
    Exit Sub

Fail:
    ReportTestFailure "TestCreateDefaultsToThisWorkbook"
End Sub

'@TestMethod("DesignerImportService")
Public Sub TestCreateWithExplicitHostWorkbook()
    CustomTestSetTitles Assert, "DesignerImportService", "TestCreateWithExplicitHostWorkbook"
    On Error GoTo Fail

    Dim subject As IDesignerImportService
    Set subject = DesignerImportService.Create(FixtureWorkbook)

    Assert.AreEqual FixtureWorkbook.Name, subject.HostWorkbook.Name, _
                    "Host should be the explicitly provided workbook."
    Exit Sub

Fail:
    ReportTestFailure "TestCreateWithExplicitHostWorkbook"
End Sub

'@TestMethod("DesignerImportService")
Public Sub TestConfigureRejectsNothing()
    CustomTestSetTitles Assert, "DesignerImportService", "TestConfigureRejectsNothing"
    On Error GoTo Fail

    Dim subject As DesignerImportService
    Set subject = New DesignerImportService

    On Error GoTo Expected
    subject.Configure Nothing
    Assert.Fail "Should have raised an error for Nothing workbook."
    Exit Sub

Expected:
    Assert.AreEqual CLng(ObjectNotInitialized), CLng(Err.Number), _
                    "Should raise ObjectNotInitialized."
    Err.Clear
    Exit Sub

Fail:
    ReportTestFailure "TestConfigureRejectsNothing"
End Sub


'@section Input Validation Tests
'===============================================================================

'@TestMethod("DesignerImportService")
Public Sub TestImportFromSetupRejectsEmptyPath()
    CustomTestSetTitles Assert, "DesignerImportService", "TestImportFromSetupRejectsEmptyPath"
    On Error GoTo Fail

    Dim subject As IDesignerImportService
    Set subject = DesignerImportService.Create(FixtureWorkbook)

    On Error GoTo Expected
    subject.ImportFromSetup vbNullString
    Assert.Fail "Should have raised an error for empty path."
    Exit Sub

Expected:
    Assert.AreEqual CLng(ObjectNotInitialized), CLng(Err.Number), _
                    "Should raise ObjectNotInitialized for empty path."
    Err.Clear
    Exit Sub

Fail:
    ReportTestFailure "TestImportFromSetupRejectsEmptyPath"
End Sub

'@TestMethod("DesignerImportService")
Public Sub TestImportFromSetupRejectsNonExistentFile()
    CustomTestSetTitles Assert, "DesignerImportService", "TestImportFromSetupRejectsNonExistentFile"
    On Error GoTo Fail

    Dim subject As IDesignerImportService
    Set subject = DesignerImportService.Create(FixtureWorkbook)

    On Error GoTo Expected
    subject.ImportFromSetup "C:\nonexistent\path\setup.xlsb"
    Assert.Fail "Should have raised an error for missing file."
    Exit Sub

Expected:
    Assert.AreEqual CLng(ElementNotFound), CLng(Err.Number), _
                    "Should raise ElementNotFound for missing file."
    Err.Clear
    Exit Sub

Fail:
    ReportTestFailure "TestImportFromSetupRejectsNonExistentFile"
End Sub

'@TestMethod("DesignerImportService")
Public Sub TestExportToLinelistRejectsNothing()
    CustomTestSetTitles Assert, "DesignerImportService", "TestExportToLinelistRejectsNothing"
    On Error GoTo Fail

    Dim subject As IDesignerImportService
    Set subject = DesignerImportService.Create(FixtureWorkbook)

    On Error GoTo Expected
    subject.ExportToLinelist Nothing
    Assert.Fail "Should have raised an error for Nothing target."
    Exit Sub

Expected:
    Assert.AreEqual CLng(ObjectNotInitialized), CLng(Err.Number), _
                    "Should raise ObjectNotInitialized for Nothing target."
    Err.Clear
    Exit Sub

Fail:
    ReportTestFailure "TestExportToLinelistRejectsNothing"
End Sub


'@section Exported Properties Tests
'===============================================================================

'@TestMethod("DesignerImportService")
Public Sub TestExportedPropertiesAreNothingBeforeExport()
    CustomTestSetTitles Assert, "DesignerImportService", "TestExportedPropertiesAreNothingBeforeExport"
    On Error GoTo Fail

    Dim subject As IDesignerImportService
    Set subject = DesignerImportService.Create(FixtureWorkbook)

    Assert.IsTrue subject.ExportedDictionary Is Nothing, "Dictionary should be Nothing before export."
    Assert.IsTrue subject.ExportedChoices Is Nothing, "Choices should be Nothing before export."
    Assert.IsTrue subject.ExportedAnalysis Is Nothing, "Analysis should be Nothing before export."
    Assert.IsTrue subject.ExportedExport Is Nothing, "Export should be Nothing before export."
    Assert.IsTrue subject.ExportedGeo Is Nothing, "Geo should be Nothing before export."
    Assert.IsTrue subject.ExportedPasswords Is Nothing, "Passwords should be Nothing before export."
    Exit Sub

Fail:
    ReportTestFailure "TestExportedPropertiesAreNothingBeforeExport"
End Sub


'@section Internal helpers
'===============================================================================

Private Sub ReportTestFailure(ByVal context As String)
    Dim message As String

    If Assert Is Nothing Then Exit Sub

    message = context & " failed with error " & Err.Number & " (" & Err.Source & "): " & Err.Description
    Assert.LogFailure message
    Err.Clear
End Sub
