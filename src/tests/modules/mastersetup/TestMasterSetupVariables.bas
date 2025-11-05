Attribute VB_Name = "TestMasterSetupVariables"
Option Explicit

'@Folder("CustomTests")
'@Folder("Tests")
'@ModuleDescription("Unit tests covering the MasterSetupVariables class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Manager As IMasterSetupVariables
Private FixtureSheet As Worksheet

Private Const VARIABLES_SHEET As String = "TST_MasterVariables"
Private Const CHOICES_SHEET As String = "TST_MasterChoices"
Private Const DROPDOWNS_SHEET As String = "TST_MasterDropdowns"
Private Const SOURCE_SHEET As String = "TST_MasterSource"
Private Const TARGET_SHEET As String = "TST_MasterTarget"
Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const VARIABLE_TABLE_NAME As String = "TST_MasterVariablesTable"
Private Const STATUS_DROPDOWN_NAME As String = "ms_variables_default_status"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestMasterSetupVariables"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    Dim lo As ListObject

    TestHelpers.DeleteWorksheets VARIABLES_SHEET, CHOICES_SHEET, DROPDOWNS_SHEET, SOURCE_SHEET, TARGET_SHEET
    Set FixtureSheet = TestHelpers.EnsureWorksheet(VARIABLES_SHEET)

    With FixtureSheet
        .Cells.Clear
        .Range("A1").Value = "Variable Name"
        .Range("B1").Value = "Variable Label"
        .Range("A2:B2").Value = vbNullString
        Set lo = .ListObjects.Add(xlSrcRange, .Range("A1:B2"), , xlYes)
    End With

    lo.Name = VARIABLE_TABLE_NAME
    lo.HeaderRowRange.Cells(1, 1).Value = "Variable Name"
    lo.HeaderRowRange.Cells(1, 2).Value = "Variable Label"
    lo.DataBodyRange.Cells(1, 1).Value = "patient_status"
    lo.DataBodyRange.Cells(1, 2).Value = "Patient status"

    Set Manager = MasterSetupVariables.Create(lo)
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    Set Manager = Nothing
    Set FixtureSheet = Nothing
    TestHelpers.DeleteWorksheets VARIABLES_SHEET, CHOICES_SHEET, DROPDOWNS_SHEET, SOURCE_SHEET, TARGET_SHEET
End Sub


'@section Tests
'===============================================================================
'@TestMethod("MasterSetupVariables")
Public Sub TestCreateEnsuresAllColumns()
    Dim lo As ListObject

    CustomTestSetTitles Assert, "MasterSetupVariables", "TestCreateEnsuresAllColumns"

    Set lo = FixtureSheet.ListObjects(VARIABLE_TABLE_NAME)

    Assert.AreEqual 8&, lo.ListColumns.Count, "Expected eight columns created by the manager."
    Assert.AreEqual "Variable Order", lo.ListColumns(1).Name
    Assert.AreEqual "Variable Section", lo.ListColumns(2).Name
    Assert.AreEqual "Variable Name", lo.ListColumns(3).Name
    Assert.AreEqual "Variable Label", lo.ListColumns(4).Name
    Assert.AreEqual "Default Choice", lo.ListColumns(5).Name
    Assert.AreEqual "Choices Values", lo.ListColumns(6).Name
    Assert.AreEqual "Default Status", lo.ListColumns(7).Name
    Assert.AreEqual "Comments", lo.ListColumns(8).Name
End Sub

'@TestMethod("MasterSetupVariables")
Public Sub TestRefreshChoicesPopulatesConcatenatedValues()
    Dim choices As ILLChoices
    Dim result As String
    Dim choicesColumn As Range

    CustomTestSetTitles Assert, "MasterSetupVariables", "TestRefreshChoicesPopulatesConcatenatedValues"

    Set choices = BuildChoicesStub()

    Manager.RefreshChoices "patient_status", choices

    Set choicesColumn = FixtureSheet.ListObjects(VARIABLE_TABLE_NAME).ListColumns("Choices Values").DataBodyRange
    result = CStr(choicesColumn.Cells(1, 1).Value)

    Assert.AreEqual "Active | Inactive", result, "Expected concatenated choices to match stub values."
End Sub

'@TestMethod("MasterSetupVariables")
Public Sub TestInitialisePersistsMetadataAndValidation()
    Dim dropdowns As IDropdownLists
    Dim statusRange As Range
    Dim hidden As IHiddenNames

    CustomTestSetTitles Assert, "MasterSetupVariables", "TestInitialisePersistsMetadataAndValidation"

    Set dropdowns = BuildDropdownsStub()
    Manager.Initialise dropdowns

    Set statusRange = FixtureSheet.ListObjects(VARIABLE_TABLE_NAME).ListColumns("Default Status").DataBodyRange
    Assert.IsTrue Not statusRange Is Nothing, "Default Status column should expose a data range."
    Assert.AreEqual xlValidateList, statusRange.Validation.Type, "Expected list validation applied to Default Status."
    Assert.IsTrue InStr(1, statusRange.Validation.Formula1, STATUS_DROPDOWN_NAME, vbTextCompare) > 0, _
                 "Validation should point to the master setup status dropdown."

    Set hidden = HiddenNames.Create(FixtureSheet)
    Assert.IsTrue hidden.HasName("__MSV__ColName"), "Expected hidden name for Variable Name column."
    Assert.IsTrue Manager.Initialised, "Initialise should set the internal flag to True."
End Sub

'@TestMethod("MasterSetupVariables")
Public Sub TestCloneToWorkbookCopiesStructureAndMetadata()
    Dim dropdowns As IDropdownLists
    Dim targetBook As Workbook
    Dim clone As IMasterSetupVariables
    Dim targetSheet As Worksheet
    Dim targetHidden As IHiddenNames

    CustomTestSetTitles Assert, "MasterSetupVariables", "TestCloneToWorkbookCopiesStructureAndMetadata"

    Set dropdowns = BuildDropdownsStub()
    Manager.Initialise dropdowns
    FixtureSheet.ListObjects(VARIABLE_TABLE_NAME).ListColumns("Variable Section").DataBodyRange.Cells(1, 1).Value = "Core"
    FixtureSheet.ListObjects(VARIABLE_TABLE_NAME).ListColumns("Default Status").DataBodyRange.Cells(1, 1).Value = "active"

    Set targetBook = TestHelpers.NewWorkbook
    On Error GoTo Cleanup

    With targetBook.Worksheets(1)
        .Name = FixtureSheet.Name
        .Range("Z10").Value = "stale"
    End With

    Set clone = Manager.CloneToWorkbook(targetBook)
    Assert.IsTrue Not clone Is Nothing, "Clone should return a new interface instance."
    Assert.AreEqual FixtureSheet.Name, clone.Table.Parent.Name, "Clone should reuse the source sheet name by default."
    Assert.AreEqual FixtureSheet.ListObjects(VARIABLE_TABLE_NAME).ListColumns.Count, _
                    clone.Table.ListColumns.Count, _
                    "Clone should reproduce all columns."
    Assert.AreEqual FixtureSheet.ListObjects(VARIABLE_TABLE_NAME).ListRows.Count, _
                    clone.Table.ListRows.Count, _
                    "Clone should reproduce all rows."

    Set targetSheet = targetBook.Worksheets(FixtureSheet.Name)
    Set targetHidden = HiddenNames.Create(targetSheet)
    Assert.IsTrue targetHidden.HasName("__MSV__ColDefaultStatus"), _
                  "Expected hidden metadata copied to the target sheet."
    Assert.AreEqual 1&, targetSheet.ListObjects.Count, "Target sheet should contain a single table."
    Assert.AreEqual vbNullString, CStr(targetSheet.Range("Z10").Value), "Sheet content should be cleared before exporting."
    AssertColumnOrder targetSheet.ListObjects(1)

Cleanup:
    TestHelpers.DeleteWorkbook targetBook
End Sub

'@TestMethod("MasterSetupVariables")
Public Sub TestImportFromWorksheetHandlesOffsetHeaders()
    Dim sourceSheet As Worksheet
    Dim headers As Variant
    Dim values As Variant
    Dim idx As Long
    Dim table As ListObject

    CustomTestSetTitles Assert, "MasterSetupVariables", "TestImportFromWorksheetHandlesOffsetHeaders"

    headers = Array("Variable Order", "Variable Section", "Variable Name", "Variable Label", _
                    "Default Choice", "Choices Values", "Default Status", "Comments")
    values = Array(5, "Vitals", "blood_pressure", "Blood pressure", "normal", _
                   "normal | high", "mandatory", "Autogenerated")

    Set sourceSheet = TestHelpers.EnsureWorksheet(SOURCE_SHEET)
    sourceSheet.Cells.Clear

    For idx = LBound(headers) To UBound(headers)
        sourceSheet.Cells(5, 2 + idx).Value = headers(idx)
        sourceSheet.Cells(6, 2 + idx).Value = values(idx)
    Next idx

    Manager.ImportFromWorksheet sourceSheet

    Set table = FixtureSheet.ListObjects(VARIABLE_TABLE_NAME)
    Assert.AreEqual 1&, table.DataBodyRange.Rows.Count, "Import should overwrite the table with a single row."
    Assert.AreEqual values(0), table.ListColumns("Variable Order").DataBodyRange.Cells(1, 1).Value
    Assert.AreEqual values(1), table.ListColumns("Variable Section").DataBodyRange.Cells(1, 1).Value
    Assert.AreEqual values(2), table.ListColumns("Variable Name").DataBodyRange.Cells(1, 1).Value
    Assert.AreEqual values(4), table.ListColumns("Default Choice").DataBodyRange.Cells(1, 1).Value
    Assert.AreEqual values(5), table.ListColumns("Choices Values").DataBodyRange.Cells(1, 1).Value
    Assert.AreEqual values(6), table.ListColumns("Default Status").DataBodyRange.Cells(1, 1).Value
    Assert.AreEqual values(7), table.ListColumns("Comments").DataBodyRange.Cells(1, 1).Value
End Sub

'@TestMethod("MasterSetupVariables")
Public Sub TestCopyToListObjectAppliesNameStyleAndOrder()
    Dim dropdowns As IDropdownLists
    Dim targetSheet As Worksheet
    Dim targetTable As ListObject
    Dim expectedStyle As String

    CustomTestSetTitles Assert, "MasterSetupVariables", "TestCopyToListObjectAppliesNameStyleAndOrder"

    expectedStyle = "TableStyleMedium2"
    FixtureSheet.ListObjects(VARIABLE_TABLE_NAME).TableStyle = expectedStyle

    Set dropdowns = BuildDropdownsStub()
    Manager.Initialise dropdowns

    Set targetSheet = TestHelpers.EnsureWorksheet(TARGET_SHEET)
    targetSheet.Cells.Clear
    targetSheet.Range("A1").Value = "Foo"
    targetSheet.Range("B1").Value = "Bar"
    targetSheet.Range("A2:B2").Value = vbNullString
    Set targetTable = targetSheet.ListObjects.Add(xlSrcRange:=targetSheet.Range("A1:B2"), XlListObjectHasHeaders:=xlYes)
    targetTable.Name = "TempTable"

    Manager.CopyToListObject targetTable

    Assert.AreEqual FixtureSheet.ListObjects(VARIABLE_TABLE_NAME).Name, targetTable.Name, _
                     "Copy should apply the source table name to the target."
    Assert.AreEqual expectedStyle, targetTable.TableStyle, "Copy should apply the source table style."
    AssertColumnOrder targetTable
    Assert.AreEqual FixtureSheet.ListObjects(VARIABLE_TABLE_NAME).ListRows.Count, targetTable.ListRows.Count, _
                     "Target table should contain the same number of rows."
End Sub


'@section Helpers
'===============================================================================
Private Function BuildChoicesStub() As ILLChoices
    Dim sh As Worksheet

    Set sh = TestHelpers.EnsureWorksheet(CHOICES_SHEET)
    sh.Cells.Clear

    sh.Cells(4, 1).Value = "list name"
    sh.Cells(4, 2).Value = "label"
    sh.Cells(4, 3).Value = "short label"
    sh.Cells(4, 4).Value = "ordering list"

    sh.Cells(5, 1).Value = "patient_status"
    sh.Cells(5, 2).Value = "Active"
    sh.Cells(6, 1).Value = "patient_status"
    sh.Cells(6, 2).Value = "Inactive"

    Set BuildChoicesStub = LLChoices.Create(sh, 4, 1)
End Function

Private Function BuildDropdownsStub() As IDropdownLists
    Dim sh As Worksheet

    Set sh = TestHelpers.EnsureWorksheet(DROPDOWNS_SHEET)
    sh.Cells.Clear

    Set BuildDropdownsStub = DropdownLists.Create(sh)
End Function

Private Sub AssertColumnOrder(ByVal lo As ListObject)
    Dim expected As Variant
    Dim idx As Long

    expected = Array("Variable Order", "Variable Section", "Variable Name", "Variable Label", _
                     "Default Choice", "Choices Values", "Default Status", "Comments")

    For idx = LBound(expected) To UBound(expected)
        Assert.AreEqual expected(idx), lo.ListColumns(idx + 1).Name, "Unexpected column order at position " & CStr(idx + 1)
    Next idx
End Sub
