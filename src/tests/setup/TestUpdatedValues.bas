Attribute VB_Name = "TestUpdatedValues"
Attribute VB_Description = "Unit tests for the UpdatedValues watcher service"

Option Explicit

'@Folder("CustomTests.Setup")
'@ModuleDescription("Exercises the UpdatedValues class responsible for tracking watched setup columns")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private UpdatedSheet As Worksheet
Private SourceSheet As Worksheet
Private SourceTable As ListObject
Private Subject As UpdatedValues

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const UPDATED_SHEET_NAME As String = "__updated"
Private Const SOURCE_SHEET_NAME As String = "Dictionary"
Private Const NOTES_SHEET_NAME As String = "Notes"
Private Const SOURCE_TABLE_NAME As String = "Tab_Source"
Private Const SECOND_TABLE_NAME As String = "Tab_Secondary"
Private Const NOTED_TABLE_NAME As String = "Tab_Noted"
Private Const UNTAGGED_TABLE_NAME As String = "Tab_Untagged"
Private Const LEGACY_TABLE_NAME As String = "UpLo_Tab_Legacy___updated"
Private Const LEGACY_INDEX_NAME As String = "__UpLo__Names__"
Private Const LEGACY_RANGE_NAME As String = "RNG_tab_legacy_old_updated"
Private Const TAG_WATCH_UPDATE As String = "watch for update"
Private Const TAG_TRANSLATE_TEXT As String = "translate as text"
Private Const RANGE_PREFIX As String = "RNG_"
Private Const LISTOBJECT_PREFIX As String = "UpLo_"
Private Const STATUS_DEFAULT As String = "no"
Private Const STATUS_UPDATED As String = "yes"

Private RangeNameField As String
Private RangeNameLabel As String
Private RangeNameControl As String

'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    AssertSheetSetup
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestUpdatedValues"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    RestoreApp
End Sub

'@TestInitialize
Public Sub TestInitialize()
    BusyApp
    Set FixtureWorkbook = NewWorkbook
    Set UpdatedSheet = EnsureWorksheet(UPDATED_SHEET_NAME, FixtureWorkbook)
    Set SourceSheet = EnsureWorksheet(SOURCE_SHEET_NAME, FixtureWorkbook)
    Set SourceTable = BuildSourceTable(SourceSheet)
    Set Subject = UpdatedValues.Create(UpdatedSheet)
    RangeNameField = ExpectedRangeName(SOURCE_TABLE_NAME, "Name")
    RangeNameLabel = ExpectedRangeName(SOURCE_TABLE_NAME, "Label")
    RangeNameControl = ExpectedRangeName(SOURCE_TABLE_NAME, "Control Details")
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        DeleteWorkbook FixtureWorkbook
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
    On Error GoTo Fail

    Dim registry As ListObject

    Subject.AddColumns SourceTable

    Set registry = RegistryTable()
    Assert.IsFalse registry Is Nothing, "Registry table should be created when tagged columns exist"
    Assert.AreEqual CLng(3), registry.ListRows.Count, "Three tagged columns should be registered"
    Assert.IsTrue WorkbookHasName(RangeNameField), "Name column range must be defined"
    Assert.IsTrue WorkbookHasName(RangeNameLabel), "Label column range must be defined"
    Assert.IsTrue WorkbookHasName(RangeNameControl), "Control Details column range must be defined"
    Exit Sub

Fail:
    ReportTestFailure "TestAddColumnsRegistersTaggedColumns"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestAddColumnsSkipsUntaggedTable()
    CustomTestSetTitles Assert, "UpdatedValues", "TestAddColumnsSkipsUntaggedTable"
    On Error GoTo Fail

    Dim untagged As ListObject
    Dim notesSheet As Worksheet

    Set notesSheet = EnsureWorksheet(NOTES_SHEET_NAME, FixtureWorkbook)
    Set untagged = BuildUntaggedTable(notesSheet)

    Subject.AddColumns untagged

    Assert.AreEqual CLng(0), CLng(UpdatedSheet.ListObjects.Count), _
        "A table carrying no tag should leave the registry sheet empty"
    Exit Sub

Fail:
    ReportTestFailure "TestAddColumnsSkipsUntaggedTable"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestColumnTagIgnoresNoteAboveHeader()
    CustomTestSetTitles Assert, "UpdatedValues", "TestColumnTagIgnoresNoteAboveHeader"
    On Error GoTo Fail

    Dim noted As ListObject
    Dim notesSheet As Worksheet
    Dim notedRangeName As String

    Set notesSheet = EnsureWorksheet(NOTES_SHEET_NAME, FixtureWorkbook)
    Set noted = BuildNotedTable(notesSheet)
    notedRangeName = ExpectedRangeName(NOTED_TABLE_NAME, "Ident")

    Subject.AddColumns noted

    Assert.IsTrue RegistryHasRange(notedRangeName), _
        "A section title between the tag row and the header must not hide the tag"
    Assert.IsTrue WorkbookHasName(notedRangeName), _
        "The column found under the title must get its defined name"
    Exit Sub

Fail:
    ReportTestFailure "TestColumnTagIgnoresNoteAboveHeader"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCheckUpdateMarksMatchingRange()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCheckUpdateMarksMatchingRange"
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SOURCE_TABLE_NAME, SourceTable.DataBodyRange.Cells(1, 1)

    Assert.AreEqual STATUS_UPDATED, RegistryStatusValue(RangeNameField), "Matching range status should change to yes"
    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RangeNameLabel), "Non intersecting ranges should remain unchanged"
    Exit Sub

Fail:
    ReportTestFailure "TestCheckUpdateMarksMatchingRange"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCheckUpdateFlagsEveryTouchedColumn()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCheckUpdateFlagsEveryTouchedColumn"
    On Error GoTo Fail

    Dim changed As Range

    Subject.AddColumns SourceTable

    Set changed = SourceTable.DataBodyRange.Cells(1, 1).Resize(1, 2)
    changed.Value = "Changed"
    Subject.CheckUpdate SourceSheet, changed

    Assert.AreEqual STATUS_UPDATED, RegistryStatusValue(RangeNameField), _
        "A change covering two watched columns should flag the first"
    Assert.AreEqual STATUS_UPDATED, RegistryStatusValue(RangeNameLabel), _
        "A change covering two watched columns should flag the second as well"
    Exit Sub

Fail:
    ReportTestFailure "TestCheckUpdateFlagsEveryTouchedColumn"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCheckUpdateKeepsRowWhenNameIsMissing()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCheckUpdateKeepsRowWhenNameIsMissing"
    On Error GoTo Fail

    Dim rowsBefore As Long

    Subject.AddColumns SourceTable
    rowsBefore = RegistryTable().ListRows.Count

    On Error Resume Next
        FixtureWorkbook.Names(RangeNameField).Delete
    On Error GoTo Fail

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SourceSheet, SourceTable.DataBodyRange.Cells(1, 1)

    Assert.AreEqual rowsBefore, CLng(RegistryTable().ListRows.Count), _
        "A missing defined name must not cost the registry a row"
    Assert.IsTrue RegistryHasRange(RangeNameField), _
        "The row whose defined name is gone must still be registered"
    Exit Sub

Fail:
    ReportTestFailure "TestCheckUpdateKeepsRowWhenNameIsMissing"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestIsUpdatedMatchesSpacedHeader()
    CustomTestSetTitles Assert, "UpdatedValues", "TestIsUpdatedMatchesSpacedHeader"
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    Assert.IsFalse Subject.IsUpdated("control details"), "A spaced header should read as not updated before any change"

    SourceTable.DataBodyRange.Cells(1, 4).Value = "Changed"
    Subject.CheckUpdate SourceSheet, SourceTable.DataBodyRange.Cells(1, 4)

    Assert.IsTrue Subject.IsUpdated("control details"), "The header name as the sheet carries it should answer True"
    Assert.IsTrue Subject.IsUpdated("control_details"), "The underscore spelling should answer the same"
    Exit Sub

Fail:
    ReportTestFailure "TestIsUpdatedMatchesSpacedHeader"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestAcknowledgeUpdatesSilencesSecondRead()
    CustomTestSetTitles Assert, "UpdatedValues", "TestAcknowledgeUpdatesSilencesSecondRead"
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SourceSheet, SourceTable.DataBodyRange.Cells(1, 1)
    Assert.IsTrue Subject.IsUpdated("Name"), "The first read after a change should answer True"

    Subject.AcknowledgeUpdates
    Assert.IsFalse Subject.IsUpdated("Name"), "A read after AcknowledgeUpdates should answer False"

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed again"
    Subject.CheckUpdate SourceSheet, SourceTable.DataBodyRange.Cells(1, 1)
    Assert.IsTrue Subject.IsUpdated("Name"), "A later edit should be seen again"
    Exit Sub

Fail:
    ReportTestFailure "TestAcknowledgeUpdatesSilencesSecondRead"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestClearUpResetsStatuses()
    CustomTestSetTitles Assert, "UpdatedValues", "TestClearUpResetsStatuses"
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SOURCE_TABLE_NAME, SourceTable.DataBodyRange.Cells(1, 1)
    Subject.ClearUp

    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RangeNameField), "ClearUp should restore the default status"
    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RangeNameLabel), "ClearUp should reset every registered column"
    Exit Sub

Fail:
    ReportTestFailure "TestClearUpResetsStatuses"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestSwitchTagsReachesUntrackedRegistry()
    CustomTestSetTitles Assert, "UpdatedValues", "TestSwitchTagsReachesUntrackedRegistry"
    On Error GoTo Fail

    Dim legacy As ListObject

    Set legacy = BuildLegacyRegistry(UpdatedSheet)
    Assert.AreEqual STATUS_UPDATED, CStr(legacy.ListRows(1).Range.Cells(1, 3).Value), _
        "The hand-built registry should start flagged"

    Subject.SwitchTagsToNo

    Assert.AreEqual STATUS_DEFAULT, CStr(legacy.ListRows(1).Range.Cells(1, 3).Value), _
        "A registry this class never built should still be reset"
    Exit Sub

Fail:
    ReportTestFailure "TestSwitchTagsReachesUntrackedRegistry"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestDeleteUpCleansRegistryAndNames()
    CustomTestSetTitles Assert, "UpdatedValues", "TestDeleteUpCleansRegistryAndNames"
    On Error GoTo Fail

    Subject.AddColumns SourceTable
    Subject.DeleteUp

    Assert.IsTrue RegistryTable() Is Nothing, "Registry table should be removed when DeleteUp is invoked"
    Assert.IsFalse WorkbookHasName(RangeNameField), "Named range should be removed with the registry"
    Assert.IsFalse WorkbookHasName(RangeNameLabel), "Named range should be removed with the registry"
    Assert.AreEqual vbNullString, CStr(UpdatedSheet.Cells(1, 1).Value), "Registry headers should be cleared after deletion"
    Exit Sub

Fail:
    ReportTestFailure "TestDeleteUpCleansRegistryAndNames"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestDeleteUpSweepsLegacyRegistries()
    CustomTestSetTitles Assert, "UpdatedValues", "TestDeleteUpSweepsLegacyRegistries"
    On Error GoTo Fail

    BuildLegacyRegistry UpdatedSheet
    BuildLegacyNameIndex UpdatedSheet
    Subject.AddColumns SourceTable

    Subject.DeleteUp

    Assert.AreEqual CLng(0), CLng(UpdatedSheet.ListObjects.Count), _
        "Every table on the registry sheet should go, including the ones this class never built"
    Assert.IsFalse WorkbookHasName(LEGACY_RANGE_NAME), _
        "A defined name left by the older layout should go with it"
    Assert.AreEqual vbNullString, CStr(UpdatedSheet.Cells(1, 1).Value), _
        "The registry sheet should be clear after the sweep"
    Exit Sub

Fail:
    ReportTestFailure "TestDeleteUpSweepsLegacyRegistries"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestAddColumnsPrunesObsoleteEntries()
    CustomTestSetTitles Assert, "UpdatedValues", "TestAddColumnsPrunesObsoleteEntries"
    On Error GoTo Fail

    Dim registry As ListObject

    Subject.AddColumns SourceTable

    'Take the watch tag off the first column.
    SourceSheet.Cells(1, 1).Value = "skip"
    Subject.AddColumns SourceTable

    Set registry = RegistryTable()
    Assert.AreEqual CLng(2), CLng(registry.ListRows.Count), "Obsolete entries should be removed from the registry"
    Assert.IsFalse WorkbookHasName(RangeNameField), "Removed watchers must delete their named ranges"
    Assert.IsTrue RegistryHasRange(RangeNameLabel), "Remaining entries should be left in place"
    Exit Sub

Fail:
    ReportTestFailure "TestAddColumnsPrunesObsoleteEntries"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestAddColumnsKeepsAPendingFlag()
    CustomTestSetTitles Assert, "UpdatedValues", "TestAddColumnsKeepsAPendingFlag"
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SourceSheet, SourceTable.DataBodyRange.Cells(1, 1)

    Subject.AddColumns SourceTable

    Assert.AreEqual STATUS_UPDATED, RegistryStatusValue(RangeNameField), _
        "A rebuild should carry a pending flag over rather than lose the change"
    Exit Sub

Fail:
    ReportTestFailure "TestAddColumnsKeepsAPendingFlag"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCreateWritesNothingToTheSheet()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCreateWritesNothingToTheSheet"
    On Error GoTo Fail

    Dim watcher As UpdatedValues

    Set watcher = UpdatedValues.Create(UpdatedSheet)

    Assert.AreEqual CLng(0), CLng(UpdatedSheet.ListObjects.Count), "Create should add no table to the registry sheet"
    Assert.AreEqual vbNullString, CStr(UpdatedSheet.Cells(1, 1).Value), "Create should write no cell on the registry sheet"

    watcher.AddColumns SourceTable
    Assert.IsTrue WorkbookHasName(RangeNameField), "The watcher should still build its names when asked to"

    Set watcher = Nothing
    Exit Sub

Fail:
    ReportTestFailure "TestCreateWritesNothingToTheSheet"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestAddSheetRegistersOneTableForTheSheet()
    CustomTestSetTitles Assert, "UpdatedValues", "TestAddSheetRegistersOneTableForTheSheet"
    On Error GoTo Fail

    Dim secondary As ListObject
    Dim secondaryRangeName As String
    Dim registry As ListObject

    Set secondary = BuildSecondaryTable(SourceSheet)
    secondaryRangeName = ExpectedRangeName(SECOND_TABLE_NAME, "Code")

    Subject.AddSheet SourceSheet

    Assert.AreEqual CLng(1), CLng(UpdatedSheet.ListObjects.Count), _
        "The registry sheet should carry one table whatever the source sheet holds"

    Set registry = RegistryTable()
    Assert.IsFalse registry Is Nothing, "The registry should exist after AddSheet"
    Assert.AreEqual CLng(4), CLng(registry.ListRows.Count), _
        "Every tagged column of both source tables should be registered"
    Assert.IsTrue WorkbookHasName(RangeNameField), "Primary table watcher should exist after AddSheet"
    Assert.IsTrue WorkbookHasName(secondaryRangeName), "Secondary table watcher should exist after AddSheet"
    Exit Sub

Fail:
    ReportTestFailure "TestAddSheetRegistersOneTableForTheSheet"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestRemoveLoRemovesTargetedEntries()
    CustomTestSetTitles Assert, "UpdatedValues", "TestRemoveLoRemovesTargetedEntries"
    On Error GoTo Fail

    Dim secondary As ListObject
    Dim secondaryRangeName As String
    Dim registry As ListObject

    Set secondary = BuildSecondaryTable(SourceSheet)
    secondaryRangeName = ExpectedRangeName(SECOND_TABLE_NAME, "Code")

    Subject.AddSheet SourceSheet
    Subject.RemoveLo secondary

    Set registry = RegistryTable()
    Assert.IsFalse registry Is Nothing, "The registry should persist after removing one source table"
    Assert.AreEqual CLng(3), CLng(registry.ListRows.Count), "RemoveLo should leave the other watchers intact"
    Assert.IsFalse RegistryHasRange(secondaryRangeName), "The removed table should keep no registry row"
    Assert.IsTrue WorkbookHasName(RangeNameLabel), "Remaining watchers should be left intact"
    Assert.IsFalse WorkbookHasName(secondaryRangeName), "RemoveLo should delete the secondary table named ranges"
    Exit Sub

Fail:
    ReportTestFailure "TestRemoveLoRemovesTargetedEntries"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCheckUpdateWithWorksheetMarksMatchingRange()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCheckUpdateWithWorksheetMarksMatchingRange"
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    'Pass the Worksheet the way EventSetup does
    Subject.CheckUpdate SourceSheet, SourceTable.DataBodyRange.Cells(1, 1)

    Assert.AreEqual STATUS_UPDATED, RegistryStatusValue(RangeNameField), "Worksheet-scoped CheckUpdate should mark the matching range"
    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RangeNameLabel), "Non-intersecting ranges should remain unchanged"
    Exit Sub

Fail:
    ReportTestFailure "TestCheckUpdateWithWorksheetMarksMatchingRange"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCheckUpdateWithWorksheetReportsIsUpdated()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCheckUpdateWithWorksheetReportsIsUpdated"
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    Assert.IsFalse Subject.IsUpdated("Name"), "Column should not be updated before any change"

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SourceSheet, SourceTable.DataBodyRange.Cells(1, 1)

    Assert.IsTrue Subject.IsUpdated("Name"), "IsUpdated should return True after worksheet-scoped CheckUpdate"
    Assert.IsFalse Subject.IsUpdated("Label"), "Untouched column should remain not updated"
    Exit Sub

Fail:
    ReportTestFailure "TestCheckUpdateWithWorksheetReportsIsUpdated"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCheckUpdateWithWorksheetMultiTableScoping()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCheckUpdateWithWorksheetMultiTableScoping"
    On Error GoTo Fail

    Dim secondary As ListObject
    Dim secondaryRangeName As String

    Set secondary = BuildSecondaryTable(SourceSheet)
    secondaryRangeName = ExpectedRangeName(SECOND_TABLE_NAME, "Code")

    Subject.AddSheet SourceSheet

    'Edit a cell in the primary table
    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SourceSheet, SourceTable.DataBodyRange.Cells(1, 1)

    Assert.AreEqual STATUS_UPDATED, RegistryStatusValue(RangeNameField), "Primary table range should be marked as updated"
    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(secondaryRangeName), "Secondary table range should remain unchanged when primary is edited"

    'Reset and edit a cell in the secondary table
    Subject.ClearUp
    secondary.DataBodyRange.Cells(1, 1).Value = "S1-Changed"
    Subject.CheckUpdate SourceSheet, secondary.DataBodyRange.Cells(1, 1)

    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RangeNameField), "Primary table range should remain unchanged when secondary is edited"
    Assert.AreEqual STATUS_UPDATED, RegistryStatusValue(secondaryRangeName), "Secondary table range should be marked as updated"
    Exit Sub

Fail:
    ReportTestFailure "TestCheckUpdateWithWorksheetMultiTableScoping"
End Sub

'@section Helpers
'===============================================================================

Private Sub AssertSheetSetup()
    EnsureWorksheet TEST_OUTPUT_SHEET, ThisWorkbook, False
End Sub

'Tag row 1, header row 2, one data row. Three of the four columns are tagged,
'and "Control Details" carries a space so the header normalising is exercised.
Private Function BuildSourceTable(ByVal targetSheet As Worksheet) As ListObject
    Dim matrix As Variant
    Dim tableRange As Range
    Dim table As ListObject

    matrix = RowsToMatrix(Array( _
        Array(TAG_WATCH_UPDATE, TAG_TRANSLATE_TEXT, "ignore", TAG_WATCH_UPDATE), _
        Array("Name", "Label", "Meta", "Control Details"), _
        Array("Value 1", "Value 2", "Value 3", "Value 4")))

    WriteMatrix targetSheet.Cells(1, 1), matrix
    Set tableRange = targetSheet.Range("A2:D3")
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = SOURCE_TABLE_NAME

    Set BuildSourceTable = table
End Function

Private Function BuildSecondaryTable(ByVal targetSheet As Worksheet) As ListObject
    Dim matrix As Variant
    Dim tableRange As Range
    Dim table As ListObject

    matrix = RowsToMatrix(Array( _
        Array(TAG_WATCH_UPDATE, "skip"), _
        Array("Code", "Description"), _
        Array("S1", "S2")))

    WriteMatrix targetSheet.Cells(1, 6), matrix
    Set tableRange = targetSheet.Range("F2:G3")
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = SECOND_TABLE_NAME

    Set BuildSecondaryTable = table
End Function

'Tag row 1, a section title on row 2, a blank row 3 and the header on row 4.
'This is the Dictionary sheet shape that hid the tag from the old walk.
Private Function BuildNotedTable(ByVal targetSheet As Worksheet) As ListObject
    Dim matrix As Variant
    Dim tableRange As Range
    Dim table As ListObject

    matrix = RowsToMatrix(Array( _
        Array(TAG_WATCH_UPDATE, "ignore"), _
        Array("Section title", vbNullString), _
        Array(vbNullString, vbNullString), _
        Array("Ident", "Other"), _
        Array("N1", "N2")))

    WriteMatrix targetSheet.Cells(1, 1), matrix
    Set tableRange = targetSheet.Range("A4:B5")
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = NOTED_TABLE_NAME

    Set BuildNotedTable = table
End Function

Private Function BuildUntaggedTable(ByVal targetSheet As Worksheet) As ListObject
    Dim matrix As Variant
    Dim tableRange As Range
    Dim table As ListObject

    matrix = RowsToMatrix(Array( _
        Array("note", "note"), _
        Array("Alpha", "Beta"), _
        Array("A", "B")))

    WriteMatrix targetSheet.Cells(1, 4), matrix
    Set tableRange = targetSheet.Range("D2:E3")
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = UNTAGGED_TABLE_NAME

    Set BuildUntaggedTable = table
End Function

'A registry in the shape the older code left behind: one table per source table,
'placed away from A1, with no entry anywhere telling this class about it.
Private Function BuildLegacyRegistry(ByVal targetSheet As Worksheet) As ListObject
    Dim matrix As Variant
    Dim tableRange As Range
    Dim table As ListObject

    matrix = RowsToMatrix(Array( _
        Array("colname", "rngname", "updated", "headername"), _
        Array("Tab_Legacy-Old", LEGACY_RANGE_NAME, STATUS_UPDATED, TAG_WATCH_UPDATE)))

    WriteMatrix targetSheet.Cells(1, 8), matrix
    Set tableRange = targetSheet.Range("H1:K2")
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = LEGACY_TABLE_NAME

    FixtureWorkbook.Names.Add Name:=LEGACY_RANGE_NAME, RefersTo:="=" & SOURCE_SHEET_NAME & "!$A$1"

    Set BuildLegacyRegistry = table
End Function

Private Function BuildLegacyNameIndex(ByVal targetSheet As Worksheet) As ListObject
    Dim matrix As Variant
    Dim tableRange As Range
    Dim table As ListObject

    matrix = RowsToMatrix(Array( _
        Array("sheet", "listobject", "registry"), _
        Array(UPDATED_SHEET_NAME, "Tab_Legacy", LEGACY_TABLE_NAME)))

    WriteMatrix targetSheet.Cells(1, 13), matrix
    Set tableRange = targetSheet.Range("M1:O2")
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = LEGACY_INDEX_NAME

    Set BuildLegacyNameIndex = table
End Function

Private Function RegistryTable() As ListObject
    Dim registry As ListObject

    If UpdatedSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set registry = UpdatedSheet.ListObjects(ExpectedRegistryName())
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

Private Function RegistryHasRange(ByVal rangeName As String) As Boolean
    Dim registry As ListObject
    Dim row As ListRow

    Set registry = RegistryTable()
    If registry Is Nothing Then Exit Function

    For Each row In registry.ListRows
        If StrComp(CStr(row.Range.Cells(1, 2).Value), rangeName, vbTextCompare) = 0 Then
            RegistryHasRange = True
            Exit Function
        End If
    Next row
End Function

Private Function ExpectedRangeName(ByVal tableName As String, _
                                   ByVal columnName As String) As String
    ExpectedRangeName = RANGE_PREFIX & NormalizeKey(tableName) & "_" & _
                        NormalizeKey(columnName) & "_" & NormalizeKey(UpdatedSheet.Name)
End Function

Private Function ExpectedRegistryName() As String
    ExpectedRegistryName = LISTOBJECT_PREFIX & NormalizeKey(UpdatedSheet.Name)
End Function

Private Function WorkbookHasName(ByVal nameText As String) As Boolean
    Dim definedName As Name

    On Error Resume Next
        Set definedName = FixtureWorkbook.Names(nameText)
    On Error GoTo 0

    WorkbookHasName = Not (definedName Is Nothing)
End Function

Private Function NormalizeKey(ByVal valueText As String) As String
    Dim idx As Long
    Dim ch As String
    Dim buffer As String

    valueText = Replace(valueText, Chr$(160), " ")
    valueText = Trim$(valueText)

    For idx = 1 To Len(valueText)
        ch = Mid$(valueText, idx, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z", "0" To "9"
                buffer = buffer & LCase$(ch)
            Case Else
                buffer = buffer & "_"
        End Select
    Next idx

    buffer = ReplaceRepeatedUnderscores(buffer)
    buffer = TrimUnderscores(buffer)

    If LenB(buffer) = 0 Then buffer = "field"

    NormalizeKey = buffer
End Function

Private Function ReplaceRepeatedUnderscores(ByVal valueText As String) As String
    Do While InStr(valueText, "__") > 0
        valueText = Replace(valueText, "__", "_")
    Loop
    ReplaceRepeatedUnderscores = valueText
End Function

Private Function TrimUnderscores(ByVal valueText As String) As String
    Do While Len(valueText) > 0 And Left$(valueText, 1) = "_"
        valueText = Mid$(valueText, 2)
    Loop

    Do While Len(valueText) > 0 And Right$(valueText, 1) = "_"
        valueText = Left$(valueText, Len(valueText) - 1)
    Loop

    TrimUnderscores = valueText
End Function

Private Sub ReportTestFailure(ByVal context As String)
    Dim message As String

    If Assert Is Nothing Then Exit Sub

    message = context & " failed with error " & Err.Number & " (" & Err.Source & "): " & Err.Description
    Assert.LogFailure message
    Err.Clear
End Sub
