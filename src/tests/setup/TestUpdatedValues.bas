Attribute VB_Name = "TestUpdatedValues"
Attribute VB_Description = "Unit tests for the UpdatedValues watcher service"

Option Explicit

'@Folder("CustomTests.Setup")
'@ModuleDescription("Exercises the UpdatedValues class responsible for tracking watched setup columns")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private UpdatedSheet As Worksheet
Private SourceSheet As Worksheet
Private SourceTable As ListObject
Private Subject As IUpdatedValues

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const UPDATED_SHEET_NAME As String = "__updated"
Private Const SOURCE_SHEET_NAME As String = "Dictionary"
Private Const SOURCE_TABLE_NAME As String = "Tab_Source"
Private Const WATCH_IDENTIFIER As String = "dict"

Private Const TAG_WATCH_UPDATE As String = "watch for update"
Private Const TAG_TRANSLATE_TEXT As String = "translate as text"
Private Const RANGE_NAME_FIELD As String = "RNG_name_dict"
Private Const RANGE_NAME_LABEL As String = "RNG_label_dict"
Private Const STATUS_DEFAULT As String = "no"
Private Const STATUS_UPDATED As String = "yes"

'@ModuleInitialize
Private Sub ModuleInitialize()
    TestHelpers.BusyApp
    AssertSheetSetup
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestUpdatedValues"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    TestHelpers.RestoreApp
End Sub

'@TestInitialize
Private Sub TestInitialize()
    TestHelpers.BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    Set UpdatedSheet = TestHelpers.EnsureWorksheet(UPDATED_SHEET_NAME, FixtureWorkbook)
    Set SourceSheet = TestHelpers.EnsureWorksheet(SOURCE_SHEET_NAME, FixtureWorkbook)
    Set SourceTable = BuildSourceTable(SourceSheet)
    Set Subject = UpdatedValues.Create(UpdatedSheet, WATCH_IDENTIFIER)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        TestHelpers.DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set Subject = Nothing
    Set SourceTable = Nothing
    Set SourceSheet = Nothing
    Set UpdatedSheet = Nothing
    Set FixtureWorkbook = Nothing
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestAddColumnsRegistersTaggedColumns()
    CustomTestSetTitles Assert, "UpdatedValues", "TestAddColumnsRegistersTaggedColumns"
    Subject.AddColumns SourceTable

    Dim registry As ListObject
    Set registry = RegistryTable()
    Assert.IsFalse registry Is Nothing, "Registry table should be created when tagged columns exist"
    Assert.AreEqual CLng(2), registry.ListRows.Count, "Two tagged columns should be registered"
    Assert.IsTrue WorkbookHasName(RANGE_NAME_FIELD), "Name column range must be defined"
    Assert.IsTrue WorkbookHasName(RANGE_NAME_LABEL), "Label column range must be defined"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCheckUpdateMarksMatchingRange()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCheckUpdateMarksMatchingRange"
    Subject.AddColumns SourceTable

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SourceSheet, SourceTable.DataBodyRange.Cells(1, 1)

    Assert.AreEqual STATUS_UPDATED, RegistryStatusValue(RANGE_NAME_FIELD), "Matching range status should change to yes"
    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RANGE_NAME_LABEL), "Non intersecting ranges should remain unchanged"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestClearUpResetsStatuses()
    CustomTestSetTitles Assert, "UpdatedValues", "TestClearUpResetsStatuses"
    Subject.AddColumns SourceTable

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SourceSheet, SourceTable.DataBodyRange.Cells(1, 1)
    Subject.ClearUp

    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RANGE_NAME_FIELD), "ClearUp should restore the default status"
    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RANGE_NAME_LABEL), "ClearUp should reset every registered column"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestDeleteUpCleansRegistryAndNames()
    CustomTestSetTitles Assert, "UpdatedValues", "TestDeleteUpCleansRegistryAndNames"
    Subject.AddColumns SourceTable
    Subject.DeleteUp

    Assert.IsTrue RegistryTable() Is Nothing, "Registry table should be removed when DeleteUp is invoked"
    Assert.IsFalse WorkbookHasName(RANGE_NAME_FIELD), "Named range should be removed with the registry"
    Assert.IsFalse WorkbookHasName(RANGE_NAME_LABEL), "Named range should be removed with the registry"
    Assert.AreEqual vbNullString, UpdatedSheet.Cells(1, 1).Value, "Registry headers should be cleared after deletion"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestAddColumnsPrunesObsoleteEntries()
    CustomTestSetTitles Assert, "UpdatedValues", "TestAddColumnsPrunesObsoleteEntries"
    Subject.AddColumns SourceTable

    'Simulate removing the watch tag for the first column.
    SourceSheet.Cells(1, 1).Value = "skip"
    Subject.AddColumns SourceTable

    Dim registry As ListObject
    Set registry = RegistryTable()
    Assert.AreEqual CLng(1), registry.ListRows.Count, "Obsolete entries should be removed from the registry"
    Assert.IsFalse WorkbookHasName(RANGE_NAME_FIELD), "Removed watchers must delete their named ranges"
    Assert.AreEqual RANGE_NAME_LABEL, registry.ListRows(1).Range.Cells(1, 2).Value, "Remaining entry should correspond to label watcher"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCreateRejectsEmptyIdentifier()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCreateRejectsEmptyIdentifier"
    On Error GoTo ExpectError
        Dim invalid As IUpdatedValues
        Set invalid = UpdatedValues.Create(UpdatedSheet, vbNullString)
        Assert.LogFailure "Create should reject empty identifiers"
        Exit Sub
ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number, "Create should raise InvalidArgument for empty identifiers"
    Err.Clear
End Sub

'@section Helpers
'===============================================================================

Private Sub AssertSheetSetup()
    TestHelpers.EnsureWorksheet TEST_OUTPUT_SHEET, ThisWorkbook, False
End Sub

Private Function BuildSourceTable(ByVal targetSheet As Worksheet) As ListObject
    Dim matrix As Variant
    Dim tableRange As Range
    Dim table As ListObject

    matrix = TestHelpers.RowsToMatrix(Array( _
        Array(TAG_WATCH_UPDATE, TAG_TRANSLATE_TEXT, "ignore"), _
        Array("Name", "Label", "Meta"), _
        Array("Value 1", "Value 2", "Value 3")))

    TestHelpers.WriteMatrix targetSheet.Cells(1, 1), matrix
    Set tableRange = targetSheet.Range("A2:C3")
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = SOURCE_TABLE_NAME

    Set BuildSourceTable = table
End Function

Private Function RegistryTable() As ListObject
    Dim registry As ListObject
    On Error Resume Next
        Set registry = UpdatedSheet.ListObjects("UpLo_" & WATCH_IDENTIFIER)
    On Error GoTo 0
    Set RegistryTable = registry
End Function

Private Function RegistryStatusValue(ByVal rangeName As String) As String
    Dim registry As ListObject
    Dim row As ListRow

    Set registry = RegistryTable()
    If registry Is Nothing Then Exit Function

    For Each row In registry.ListRows
        If StrComp(CStr(row.Range.Cells(1, 2).Value), rangeName, vbTextCompare) = 0 Then
            RegistryStatusValue = CStr(row.Range.Cells(1, 3).Value)
            Exit Function
        End If
    Next row
End Function

Private Function WorkbookHasName(ByVal nameText As String) As Boolean
    Dim definedName As Name
    On Error Resume Next
        Set definedName = FixtureWorkbook.Names(nameText)
    On Error GoTo 0
    WorkbookHasName = Not (definedName Is Nothing)
End Function
