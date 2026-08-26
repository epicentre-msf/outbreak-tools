Attribute VB_Name = "TestMasterSetupHelpers"
Attribute VB_Description = "Tests covering row management and protection of the master setup sheets"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests covering row management and protection of the master setup sheets")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIXTURE_SHEET As String = "HelpersDiseaseFixture"
Private Const UNTAGGED_SHEET As String = "HelpersPlainFixture"
Private Const CHOICES_SHEET As String = "Choices"
Private Const PASSWORD_SHEET As String = "__pass"
Private Const FOREIGN_PASSWORD As String = "not-the-debug-password"

Private Assert As CustomTest
'Names of the sheets PrepareNamedSheet added, so the cleanup deletes only
'what this module made and leaves a Choices sheet another suite owns alone.
Private createdSheets As Collection

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
    'The assertions of a test reach the results sheet only once flushed.
    Assert.Flush
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

    Assert.AreEqual 6&, CLng(table.ListRows.Count), "A sheet should gain five rows per add"

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
    table.DataBodyRange.Cells(3, 2).Value = "var_also_kept"

    MasterSetupHelpers.ManageRows fixtureSheet, False

    Assert.AreEqual 2&, CLng(table.ListRows.Count), "A resize should keep the filled rows and drop the empty ones"
    Assert.AreEqual "var_kept", table.DataBodyRange.Cells(1, 2).Value, "The first filled row should survive the trim"
    Assert.AreEqual "var_also_kept", table.DataBodyRange.Cells(2, 2).Value, "A filled row below empty ones should move up"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestManageRowsTrimsDiseaseRows", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupHelpers")
Public Sub TestManageRowsTrimsLinesCarryingOnlyFormulas()
    CustomTestSetTitles Assert, "MasterSetupHelpers", "TestManageRowsTrimsLinesCarryingOnlyFormulas"

    Dim fixtureSheet As Worksheet
    Dim table As ListObject

    On Error GoTo Fail

    'A disease line carries two formulas, the label and the choice values,
    'from the day it is built. A line with nothing else on it is empty.
    Set fixtureSheet = PrepareDiseaseFixture(5)
    Set table = fixtureSheet.ListObjects(1)
    table.ListColumns(4).DataBodyRange.Formula = "=""label"""
    table.ListColumns(6).DataBodyRange.Formula = "=""values"""
    table.DataBodyRange.Cells(2, 2).Value = "var_kept"

    MasterSetupHelpers.ManageRows fixtureSheet, False

    Assert.AreEqual 2&, CLng(table.ListRows.Count), "Lines carrying only their formulas should go"
    Assert.AreEqual "var_kept", table.DataBodyRange.Cells(2, 2).Value, "The line with a name should stay"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestManageRowsTrimsLinesCarryingOnlyFormulas", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupHelpers")
Public Sub TestManageRowsRaisesWhenTheSheetStaysProtected()
    CustomTestSetTitles Assert, "MasterSetupHelpers", "TestManageRowsRaisesWhenTheSheetStaysProtected"

    Dim fixtureSheet As Worksheet
    Dim raisedNumber As Long

    On Error GoTo Fail

    'A password the passwords sheet does not hold: UnProtect cannot lift it,
    'and the add meets a protected table. The ribbon reads the raise.
    Set fixtureSheet = PrepareDiseaseFixture(1)
    fixtureSheet.Protect Password:=FOREIGN_PASSWORD

    On Error Resume Next
        MasterSetupHelpers.ManageRows fixtureSheet, True
        raisedNumber = Err.Number
        Err.Clear
    On Error GoTo Fail

    fixtureSheet.Unprotect FOREIGN_PASSWORD

    Assert.IsTrue raisedNumber <> 0, "A refused add should reach the caller as a raise"
    Assert.AreEqual 1&, CLng(fixtureSheet.ListObjects(1).ListRows.Count), "The table should be left as it was"

    Exit Sub

Fail:
    On Error Resume Next
        fixtureSheet.Unprotect FOREIGN_PASSWORD
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestManageRowsRaisesWhenTheSheetStaysProtected", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupHelpers")
Public Sub TestSheetKindReadsTheTagThenTheName()
    CustomTestSetTitles Assert, "MasterSetupHelpers", "TestSheetKindReadsTheTagThenTheName"

    Dim fixtureSheet As Worksheet
    Dim namedSheet As Worksheet

    On Error GoTo Fail

    Set fixtureSheet = PrepareDiseaseFixture(1)
    Assert.AreEqual "disease", MasterSetupHelpers.ResolveMasterSheetKind(fixtureSheet), "The hidden tag names a disease sheet"

    'A Choices sheet the preparation has never tagged is known by its name.
    Set namedSheet = PrepareNamedSheet(CHOICES_SHEET)
    Assert.AreEqual "choices", MasterSetupHelpers.ResolveMasterSheetKind(namedSheet), "An untagged Choices sheet is known by its name"

    Set namedSheet = PrepareNamedSheet(UNTAGGED_SHEET)
    Assert.AreEqual vbNullString, MasterSetupHelpers.ResolveMasterSheetKind(namedSheet), "A sheet with no tag and another name has no kind"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSheetKindReadsTheTagThenTheName", Err.Number, Err.Description
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

    MasterSetupHelpers.UnProtectMasterSetupSheet fixtureSheet
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

    MasterSetupHelpers.UnProtectMasterSetupSheet fixtureSheet
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

'@description Answer a sheet of the given name, adding a blank one when the workbook has none.
Private Function PrepareNamedSheet(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    For Each sh In ThisWorkbook.Worksheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            Set PrepareNamedSheet = sh
            Exit Function
        End If
    Next sh

    Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    sh.Name = sheetName

    If createdSheets Is Nothing Then Set createdSheets = New Collection
    createdSheets.Add sheetName

    Set PrepareNamedSheet = sh
End Function

Private Sub DeleteCreatedSheets()
    Dim idx As Long

    If createdSheets Is Nothing Then Exit Sub

    On Error Resume Next
    For idx = 1 To createdSheets.Count
        DeleteWorksheet CStr(createdSheets.Item(idx))
    Next idx
    On Error GoTo 0

    Set createdSheets = Nothing
End Sub

Private Sub CleanupEnvironment()
    Dim fixtureSheet As Worksheet

    DeleteCreatedSheets

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
