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
Private Const TAG_WATCH_UPDATE As String = "watch for update"
Private Const TAG_TRANSLATE_TEXT As String = "translate as text"
Private Const RANGE_PREFIX As String = "RNG_"
Private Const STATUS_DEFAULT As String = "no"
Private Const STATUS_UPDATED As String = "yes"
Private Const SECOND_TABLE_NAME As String = "Tab_Secondary"
Private Const NAMES_TABLE_NAME As String = "__UpLo__Names__"

Private RangeNameField As String
Private RangeNameLabel As String
Private PrimaryRegistryName As String
Private SecondaryRegistryName As String

'@ModuleInitialize
Public Sub ModuleInitialize()
    TestHelpers.BusyApp
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
    TestHelpers.RestoreApp
End Sub

'@TestInitialize
Public Sub TestInitialize()
    TestHelpers.BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    Set UpdatedSheet = TestHelpers.EnsureWorksheet(UPDATED_SHEET_NAME, FixtureWorkbook)
    Set SourceSheet = TestHelpers.EnsureWorksheet(SOURCE_SHEET_NAME, FixtureWorkbook)
    Set SourceTable = BuildSourceTable(SourceSheet)
    Set Subject = UpdatedValues.Create(UpdatedSheet)
    PrimaryRegistryName = ExpectedRegistryName(SOURCE_TABLE_NAME, UpdatedSheet)
    SecondaryRegistryName = ExpectedRegistryName(SECOND_TABLE_NAME, UpdatedSheet)
    RangeNameField = ExpectedRangeName(SOURCE_TABLE_NAME, "Name", UpdatedSheet)
    RangeNameLabel = ExpectedRangeName(SOURCE_TABLE_NAME, "Label", UpdatedSheet)
End Sub

'@TestCleanup
Public Sub TestCleanup()
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
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    Dim registry As ListObject
    Set registry = RegistryTable(PrimaryRegistryName)
    Assert.IsFalse registry Is Nothing, "Registry table should be created when tagged columns exist"
    Assert.AreEqual CLng(2), registry.ListRows.Count, "Two tagged columns should be registered"
    Assert.IsTrue WorkbookHasName(RangeNameField), "Name column range must be defined"
    Assert.IsTrue WorkbookHasName(RangeNameLabel), "Label column range must be defined"
    Assert.AreEqual PrimaryRegistryName, RegistryNameFromIndex(SOURCE_TABLE_NAME), "Name index should track the primary registry"
    Exit Sub

Fail:
    ReportTestFailure "TestAddColumnsRegistersTaggedColumns"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCheckUpdateMarksMatchingRange()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCheckUpdateMarksMatchingRange"
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SOURCE_TABLE_NAME, SourceTable.DataBodyRange.Cells(1, 1)

    Assert.AreEqual STATUS_UPDATED, RegistryStatusValue(RangeNameField, PrimaryRegistryName), "Matching range status should change to yes"
    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RangeNameLabel, PrimaryRegistryName), "Non intersecting ranges should remain unchanged"
    Exit Sub

Fail:
    ReportTestFailure "TestCheckUpdateMarksMatchingRange"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestClearUpResetsStatuses()
    CustomTestSetTitles Assert, "UpdatedValues", "TestClearUpResetsStatuses"
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    SourceTable.DataBodyRange.Cells(1, 1).Value = "Changed"
    Subject.CheckUpdate SOURCE_TABLE_NAME, SourceTable.DataBodyRange.Cells(1, 1)
    Subject.ClearUp

    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RangeNameField, PrimaryRegistryName), "ClearUp should restore the default status"
    Assert.AreEqual STATUS_DEFAULT, RegistryStatusValue(RangeNameLabel, PrimaryRegistryName), "ClearUp should reset every registered column"
    Exit Sub

Fail:
    ReportTestFailure "TestClearUpResetsStatuses"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestDeleteUpCleansRegistryAndNames()
    CustomTestSetTitles Assert, "UpdatedValues", "TestDeleteUpCleansRegistryAndNames"
    On Error GoTo Fail

    Subject.AddColumns SourceTable
    Subject.DeleteUp

    Assert.IsTrue RegistryTable(PrimaryRegistryName) Is Nothing, "Registry table should be removed when DeleteUp is invoked"
    Assert.IsFalse WorkbookHasName(RangeNameField), "Named range should be removed with the registry"
    Assert.IsFalse WorkbookHasName(RangeNameLabel), "Named range should be removed with the registry"
    Assert.AreEqual vbNullString, UpdatedSheet.Cells(1, 1).Value, "Registry headers should be cleared after deletion"
    Assert.AreEqual CLng(0), NamesIndexCount(), "Name index should be cleared after deletion"
    Exit Sub

Fail:
    ReportTestFailure "TestDeleteUpCleansRegistryAndNames"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestAddColumnsPrunesObsoleteEntries()
    CustomTestSetTitles Assert, "UpdatedValues", "TestAddColumnsPrunesObsoleteEntries"
    On Error GoTo Fail

    Subject.AddColumns SourceTable

    'Simulate removing the watch tag for the first column.
    SourceSheet.Cells(1, 1).Value = "skip"
    Subject.AddColumns SourceTable

    Dim registry As ListObject
    Set registry = RegistryTable(PrimaryRegistryName)
    Assert.AreEqual CLng(1), registry.ListRows.Count, "Obsolete entries should be removed from the registry"
    Assert.IsFalse WorkbookHasName(RangeNameField), "Removed watchers must delete their named ranges"
    Assert.AreEqual RangeNameLabel, registry.ListRows(1).Range.Cells(1, 2).Value, "Remaining entry should correspond to label watcher"
    Exit Sub

Fail:
    ReportTestFailure "TestAddColumnsPrunesObsoleteEntries"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestCreateRegistersWithoutIdentifier()
    CustomTestSetTitles Assert, "UpdatedValues", "TestCreateRegistersWithoutIdentifier"
    On Error GoTo Fail

    Dim watcher As IUpdatedValues
    Dim defaultRange As String

    defaultRange = ExpectedRangeName(SOURCE_TABLE_NAME, "Name", UpdatedSheet)
    Set watcher = UpdatedValues.Create(UpdatedSheet)

    watcher.AddColumns SourceTable

    Assert.IsTrue WorkbookHasName(defaultRange), "Default identifier should build expected named ranges"
    Assert.AreEqual PrimaryRegistryName, RegistryNameFromIndex(SOURCE_TABLE_NAME), "Name index should capture the primary registry name"

    watcher.DeleteUp
    Set watcher = Nothing
    Exit Sub

Fail:
    ReportTestFailure "TestCreateRegistersWithoutIdentifier"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestAddSheetRegistersAllTables()
    CustomTestSetTitles Assert, "UpdatedValues", "TestAddSheetRegistersAllTables"
    On Error GoTo Fail

    Dim secondary As ListObject
    Dim secondaryRangeName As String
    Dim registry As ListObject

    Set secondary = BuildSecondaryTable(SourceSheet)
    secondaryRangeName = ExpectedRangeName(SECOND_TABLE_NAME, "Code", UpdatedSheet)

    Subject.AddSheet SourceSheet

    Assert.AreEqual SourceSheet.ListObjects.Count, RegistryTableCount(), "AddSheet should create a registry for every source table"

    Set registry = RegistryTable(PrimaryRegistryName)
    Assert.IsFalse registry Is Nothing, "Primary registry should exist after AddSheet"
    Assert.AreEqual CLng(2), registry.ListRows.Count, "Primary registry should capture both tagged columns"

    Set registry = RegistryTable(SecondaryRegistryName)
    Assert.IsFalse registry Is Nothing, "Secondary registry should exist after AddSheet"
    Assert.AreEqual CLng(1), registry.ListRows.Count, "Secondary registry should capture its tagged column"

    Assert.IsTrue WorkbookHasName(RangeNameField), "Primary table watcher should exist after AddSheet"
    Assert.IsTrue WorkbookHasName(secondaryRangeName), "Secondary table watcher should exist after AddSheet"
    Assert.AreEqual CLng(2), NamesIndexCount(), "Name index should contain both registries"
    Assert.AreEqual PrimaryRegistryName, RegistryNameFromIndex(SOURCE_TABLE_NAME), "Name index should track the primary registry"
    Assert.AreEqual SecondaryRegistryName, RegistryNameFromIndex(SECOND_TABLE_NAME), "Name index should track the secondary registry"
    Exit Sub

Fail:
    ReportTestFailure "TestAddSheetRegistersAllTables"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestRemoveLoRemovesTargetedTable()
    CustomTestSetTitles Assert, "UpdatedValues", "TestRemoveLoRemovesTargetedTable"
    On Error GoTo Fail

    Dim secondary As ListObject
    Dim secondaryRangeName As String
    Dim registry As ListObject

    Set secondary = BuildSecondaryTable(SourceSheet)
    secondaryRangeName = ExpectedRangeName(SECOND_TABLE_NAME, "Code", UpdatedSheet)

    Subject.AddSheet SourceSheet
    Subject.RemoveLo secondary

    Assert.AreEqual CLng(1), RegistryTableCount(), "Only the primary registry should remain after removing the secondary table"

    Set registry = RegistryTable(PrimaryRegistryName)
    Assert.IsFalse registry Is Nothing, "Primary registry should persist after removing a single ListObject"
    Assert.AreEqual CLng(2), registry.ListRows.Count, "RemoveLo should leave primary watchers intact"
    Assert.IsTrue RegistryTable(SecondaryRegistryName) Is Nothing, "Secondary registry should be removed"

    Assert.IsTrue WorkbookHasName(RangeNameLabel), "Remaining watchers should be left intact"
    Assert.IsFalse WorkbookHasName(secondaryRangeName), "RemoveLo should delete secondary table named ranges"
    Assert.AreEqual CLng(1), NamesIndexCount(), "Name index should retain only the primary registry"
    Assert.AreEqual PrimaryRegistryName, RegistryNameFromIndex(SOURCE_TABLE_NAME), "Primary registry should remain indexed"
    Assert.AreEqual vbNullString, RegistryNameFromIndex(SECOND_TABLE_NAME), "Secondary registry should be removed from the name index"
    Exit Sub

Fail:
    ReportTestFailure "TestRemoveLoRemovesTargetedTable"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestRegistryPlacementSkipsOccupiedColumns()
    CustomTestSetTitles Assert, "UpdatedValues", "TestRegistryPlacementSkipsOccupiedColumns"
    On Error GoTo Fail

    Dim blocking As ListObject
    Dim registry As ListObject
    Dim expectedStart As Long

    Set blocking = BuildBlockingRegistryTable(UpdatedSheet)

    Subject.AddColumns SourceTable
    Set registry = RegistryTable(PrimaryRegistryName)
    Assert.IsFalse registry Is Nothing, "Registry table should exist after AddColumns"

    expectedStart = blocking.Range.Column + blocking.Range.Columns.Count + 5
    Assert.AreEqual expectedStart, registry.Range.Column, "Registry should start after existing ListObject block"

    Subject.DeleteUp
    Subject.AddColumns SourceTable

    Set registry = RegistryTable(PrimaryRegistryName)
    Assert.IsFalse registry Is Nothing, "Registry table should be recreated after DeleteUp"
    Assert.AreEqual expectedStart, registry.Range.Column, "Registry rebuild should continue after existing ListObject block"
    Exit Sub

Fail:
    ReportTestFailure "TestRegistryPlacementSkipsOccupiedColumns"
End Sub

'@TestMethod("UpdatedValues")
Public Sub TestDeleteUpResetsPlacementAfterClearingSheet()
    CustomTestSetTitles Assert, "UpdatedValues", "TestDeleteUpResetsPlacementAfterClearingSheet"
    On Error GoTo Fail

    Dim blocking As ListObject
    Dim registry As ListObject

    Set blocking = BuildBlockingRegistryTable(UpdatedSheet)
    Subject.AddColumns SourceTable

    Subject.DeleteUp
    blocking.Delete

    Subject.AddColumns SourceTable
    Set registry = RegistryTable(PrimaryRegistryName)
    Assert.IsFalse registry Is Nothing, "Registry should be recreated after deletion"
    Exit Sub

Fail:
    ReportTestFailure "TestDeleteUpResetsPlacementAfterClearingSheet"
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

Private Function BuildSecondaryTable(ByVal targetSheet As Worksheet) As ListObject
    Dim matrix As Variant
    Dim tableRange As Range
    Dim table As ListObject

    matrix = TestHelpers.RowsToMatrix(Array( _
        Array(TAG_WATCH_UPDATE, "skip"), _
        Array("Code", "Description"), _
        Array("S1", "S2")))

    TestHelpers.WriteMatrix targetSheet.Cells(1, 5), matrix
    Set tableRange = targetSheet.Range("E2:F3")
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = SECOND_TABLE_NAME

    Set BuildSecondaryTable = table
End Function

Private Function BuildBlockingRegistryTable(ByVal targetSheet As Worksheet) As ListObject
    Dim matrix As Variant
    Dim tableRange As Range
    Dim table As ListObject

    matrix = TestHelpers.RowsToMatrix(Array( _
        Array("BlockCol1", "BlockCol2", "BlockCol3", "BlockCol4"), _
        Array("B1", "B2", "B3", "B4")))

    TestHelpers.WriteMatrix targetSheet.Cells(1, 1), matrix
    Set tableRange = targetSheet.Range("A1:D2")
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = "ExistingRegistryBlock"

    Set BuildBlockingRegistryTable = table
End Function

Private Function RegistryTable(Optional ByVal registryName As String = vbNullString) As ListObject
    Dim registry As ListObject
    Dim targetName As String

    targetName = registryName
    If LenB(targetName) = 0 Then
        targetName = LastRegistryName()
    End If

    If LenB(targetName) = 0 Then Exit Function

    On Error Resume Next
        Set registry = UpdatedSheet.ListObjects(targetName)
    On Error GoTo 0
    Set RegistryTable = registry
End Function

Private Function NamesIndexTable() As ListObject
    Dim nameList As ListObject

    If UpdatedSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set nameList = UpdatedSheet.ListObjects(NAMES_TABLE_NAME)
    On Error GoTo 0

    Set NamesIndexTable = nameList
End Function

Private Function NamesIndexCount() As Long
    Dim nameList As ListObject
    Dim listRow As ListRow

    Set nameList = NamesIndexTable()
    If nameList Is Nothing Then Exit Function
    If UpdatedSheet Is Nothing Then Exit Function

    For Each listRow In nameList.ListRows
        If StrComp(CStr(listRow.Range.Cells(1, 1).Value), UpdatedSheet.Name, vbTextCompare) = 0 Then
            NamesIndexCount = NamesIndexCount + 1
        End If
    Next listRow
End Function

Private Function RegistryNameFromIndex(ByVal tableName As String) As String
    Dim nameList As ListObject
    Dim listRow As ListRow

    Set nameList = NamesIndexTable()
    If nameList Is Nothing Then Exit Function
    If UpdatedSheet Is Nothing Then Exit Function

    For Each listRow In nameList.ListRows
        If StrComp(CStr(listRow.Range.Cells(1, 1).Value), UpdatedSheet.Name, vbTextCompare) = 0 _
           And StrComp(CStr(listRow.Range.Cells(1, 2).Value), tableName, vbTextCompare) = 0 Then
            RegistryNameFromIndex = CStr(listRow.Range.Cells(1, 3).Value)
            Exit Function
        End If
    Next listRow
End Function

Private Function RegistryTableCount() As Long
    Dim lo As ListObject

    For Each lo In UpdatedSheet.ListObjects
        If IsRegistryListObjectName(lo.Name) Then
            RegistryTableCount = RegistryTableCount + 1
        End If
    Next lo
End Function

Private Function LastRegistryName() As String
    Dim lo As ListObject

    For Each lo In UpdatedSheet.ListObjects
        If IsRegistryListObjectName(lo.Name) Then
            LastRegistryName = lo.Name
        End If
    Next lo
End Function

Private Function IsRegistryListObjectName(ByVal nameText As String) As Boolean
    If LenB(nameText) = 0 Then Exit Function
    IsRegistryListObjectName = (Left$(nameText, Len("UpLo_")) = "UpLo_")
End Function

Private Function RegistryStatusValue(ByVal rangeName As String, _
                                     Optional ByVal registryName As String = vbNullString) As String
    Dim registry As ListObject
    Dim row As ListRow

    Set registry = RegistryTable(registryName)
    If registry Is Nothing Then Exit Function

    For Each row In registry.ListRows
        If StrComp(CStr(row.Range.Cells(1, 2).Value), rangeName, vbTextCompare) = 0 Then
            RegistryStatusValue = CStr(row.Range.Cells(1, 3).Value)
            Exit Function
        End If
    Next row
End Function

Private Function ExpectedRangeName(ByVal tableName As String, _
                                   ByVal columnName As String, _
                                   ByVal registrySheet As Worksheet) As String
    ExpectedRangeName = RANGE_PREFIX & NormalizeKey(tableName) & "_" & NormalizeKey(columnName) & "_" & NormalizeKey(registrySheet.Name)
End Function

Private Function ExpectedRegistryName(ByVal tableName As String, _
                                      ByVal registrySheet As Worksheet) As String
    ExpectedRegistryName = "UpLo_" & NormalizeKey(tableName) & "_" & NormalizeKey(registrySheet.Name)
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
            Case "_"
                buffer = buffer & "_"
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
