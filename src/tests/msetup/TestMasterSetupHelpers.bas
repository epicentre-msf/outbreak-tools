Attribute VB_Name = "TestMasterSetupHelpers"
Attribute VB_Description = "Tests covering row management and protection of the master setup sheets"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests covering row management and protection of the master setup sheets")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "HelpersDiseaseFixture"
Private Const PASSWORD_SHEET As String = "__pass"

Private Assert As CustomTest

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestMasterSetupHelpers"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        CleanupEnvironment
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    CleanupEnvironment
End Sub

'@TestCleanup
Private Sub TestCleanup()
    CleanupEnvironment
End Sub

'@section Tests
'===============================================================================

'@TestMethod("MasterSetupHelpers")
Public Sub TestManageRowsAddsDiseaseRows()
    CustomTestSetTitles Assert, "MasterSetupHelpers", "TestManageRowsAddsDiseaseRows"

    Dim fixtureSheet As Worksheet
    Dim table As ListObject

    On Error GoTo Fail

    Set fixtureSheet = PrepareDiseaseFixture(1)
    Set table = fixtureSheet.ListObjects(1)

    MasterSetupHelpers.ManageRows fixtureSheet, True

    Assert.AreEqual 3&, CLng(table.ListRows.Count), "A disease sheet should gain two rows per add"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestManageRowsAddsDiseaseRows", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupHelpers")
Public Sub TestManageRowsTrimsDiseaseRows()
    CustomTestSetTitles Assert, "MasterSetupHelpers", "TestManageRowsTrimsDiseaseRows"

    Dim fixtureSheet As Worksheet
    Dim table As ListObject

    On Error GoTo Fail

    Set fixtureSheet = PrepareDiseaseFixture(15)
    Set table = fixtureSheet.ListObjects(1)
    table.DataBodyRange.Cells(1, 2).Value = "var_kept"

    MasterSetupHelpers.ManageRows fixtureSheet, False

    Assert.AreEqual 10&, CLng(table.ListRows.Count), "A disease sheet resize should keep ten rows"
    Assert.AreEqual "var_kept", table.DataBodyRange.Cells(1, 2).Value, "Filled rows should survive the trim"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestManageRowsTrimsDiseaseRows", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupHelpers")
Public Sub TestDiseaseProtectionForbidsColumnDeletes()
    CustomTestSetTitles Assert, "MasterSetupHelpers", "TestDiseaseProtectionForbidsColumnDeletes"

    Dim fixtureSheet As Worksheet

    On Error GoTo Fail

    PreparePasswordsFixture PASSWORD_SHEET
    Set fixtureSheet = PrepareDiseaseFixture(1)

    MasterSetupHelpers.ProtectMasterSetupSheet fixtureSheet, "disease"

    Assert.IsTrue fixtureSheet.ProtectContents, "The disease sheet should be protected"
    Assert.IsFalse fixtureSheet.Protection.AllowDeletingColumns, "Columns of a disease sheet stay"
    Assert.IsFalse fixtureSheet.Protection.AllowDeletingRows, "Row deletes on a disease sheet go through the ribbon"

    MasterSetupHelpers.UnProtectMasterSetupSheet fixtureSheet, "disease"
    Assert.IsFalse fixtureSheet.ProtectContents, "UnProtect should release the sheet"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDiseaseProtectionForbidsColumnDeletes", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupHelpers")
Public Sub TestVariablesProtectionAllowsRowDeletes()
    CustomTestSetTitles Assert, "MasterSetupHelpers", "TestVariablesProtectionAllowsRowDeletes"

    Dim fixtureSheet As Worksheet

    On Error GoTo Fail

    PreparePasswordsFixture PASSWORD_SHEET
    Set fixtureSheet = PrepareDiseaseFixture(1)

    MasterSetupHelpers.ProtectMasterSetupSheet fixtureSheet, "variables"

    Assert.IsTrue fixtureSheet.ProtectContents, "The variables sheet should be protected"
    Assert.IsTrue fixtureSheet.Protection.AllowDeletingRows, "Rows of the variables sheet may be deleted"
    Assert.IsFalse fixtureSheet.Protection.AllowDeletingColumns, "Columns of the variables sheet stay"

    MasterSetupHelpers.UnProtectMasterSetupSheet fixtureSheet, "variables"
    Assert.IsFalse fixtureSheet.ProtectContents, "UnProtect should release the sheet"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestVariablesProtectionAllowsRowDeletes", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

'@description Build a disease-tagged fixture sheet carrying one seven-column table.
Private Function PrepareDiseaseFixture(ByVal rowCount As Long) As Worksheet
    Dim fixtureSheet As Worksheet
    Dim store As HiddenNames
    Dim tableRange As Range
    Dim header As Variant

    Set fixtureSheet = EnsureWorksheet(FIXTURE_SHEET)
    ClearWorksheet fixtureSheet

    header = RowsToMatrix(Array(Array("Variable Order", "Variable Name", "Variable Section", "Main Label", "Choice", "Choice Values", "Status")))
    WriteMatrix fixtureSheet.Range("B4"), header

    Set tableRange = fixtureSheet.Range("B4").Resize(rowCount + 1, 7)
    fixtureSheet.ListObjects.Add SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes

    Set store = HiddenNames.Create(fixtureSheet)
    store.EnsureName "sheetTag", "disease", HiddenNameTypeString

    Set PrepareDiseaseFixture = fixtureSheet
End Function

Private Sub CleanupEnvironment()
    Dim fixtureSheet As Worksheet

    On Error Resume Next
        Set fixtureSheet = ThisWorkbook.Worksheets(FIXTURE_SHEET)
        If Not fixtureSheet Is Nothing Then fixtureSheet.Unprotect "1234"
        fixtureSheet.Delete
        ThisWorkbook.Worksheets(PASSWORD_SHEET).Delete
        ThisWorkbook.Names("RNG_PublicKey").Delete
        ThisWorkbook.Names("RNG_PrivateKey").Delete
        ThisWorkbook.Names("RNG_DebuggingPassword").Delete
        ThisWorkbook.Names("RNG_DebugMode").Delete
        ThisWorkbook.Names("RNG_Version").Delete
        ThisWorkbook.Names("RNG_LabPublicKey").Delete
        ThisWorkbook.Names("RNG_LabPrivateKey").Delete
        ThisWorkbook.Names("Passwords_ProtectedSheets").Delete
    On Error GoTo 0
End Sub
