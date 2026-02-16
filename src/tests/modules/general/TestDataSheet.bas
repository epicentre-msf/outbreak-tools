Attribute VB_Name = "TestDataSheet"
Option Explicit

'@ModuleDescription("Unit tests for the DataSheet class (IDataSheet interface). Validates range-based " & _
'                    "data access, column lookups, single and multi-condition filtering, export to " & _
'                    "workbook (with and without hidden names), and import with case-insensitive " & _
'                    "column matching. All tests run against a dictionary fixture sheet that " & _
'                    "simulates a linelist dictionary layout.")
'
'@description        Test module for DataSheet / IDataSheet. Covers factory creation and property
'                    initialization, DataRange retrieval for single columns and all columns,
'                    ColumnExists with exact / case-insensitive / partial matching, AddFormatsColumns,
'                    FilterData (single condition), FiltersData (multiple conditions with edge cases),
'                    Export (data + formatting + hidden names), and Import (case-insensitive header
'                    matching plus format import). Uses the CustomTest harness with the standard
'                    CustomTestSetTitles / CustomTestLogFailure pattern.
'
'@depends            DataSheet, IDataSheet, HiddenNames, IHiddenNames, BetterArray, CustomTest,
'                    TestHelpers, DictionaryTestFixture

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulNames
'@Folder("CustomTests")

Private Const DICTIONARYFIXTURESHEET As String = "LLDictTest"
Private Const DICTOUTPUTSHEET As String = "DataOut"

Private fixtureRowCount As Long
Private fixtureColumnCount As Long

Private Assert As ICustomTest
Private dataObject As IDataSheet
Private dataWorksheet As Worksheet

'@section Helpers
'===============================================================================

'@sub-title ResetDataSheet
'@description Rebuilds the dictionary fixture sheet and re-creates the DataSheet object so that
'             every test starts from a known, clean state.
Private Sub ResetDataSheet()
    PrepareDictionaryFixture DICTIONARYFIXTURESHEET
    Set dataWorksheet = ThisWorkbook.Worksheets(DICTIONARYFIXTURESHEET)
    Set dataObject = DataSheet.Create(dataWorksheet, 1, 1)
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
'@sub-title ModuleInitialize
'@description Sets up the shared test infrastructure once before all tests in the module run.
'             Creates the test-output sheet, initializes the CustomTest assert object, prepares
'             the dictionary fixture, and caches the expected row and column counts for later
'             assertions.
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDataSheet"
    ResetDataSheet
    EnsureWorksheet DICTOUTPUTSHEET
    fixtureRowCount = DictionaryFixtureRowCount()
    fixtureColumnCount = DictionaryFixtureColumnCount()
End Sub

'@ModuleCleanup
'@sub-title ModuleCleanup
'@description Prints accumulated test results to the output sheet and tears down all fixture
'             worksheets (DICTOUTPUTSHEET and DICTIONARYFIXTURESHEET) that were created during
'             module initialization.
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
    DeleteWorksheet DICTOUTPUTSHEET
    DeleteWorksheet DICTIONARYFIXTURESHEET
End Sub

'@TestInitialize
'@sub-title TestInitialize
'@description Runs before every individual test method. Resets the DataSheet object to a clean
'             fixture state and pre-adds formatting columns (formatting condition, formatting
'             values, lock cells) so that tests which depend on those columns do not need to add
'             them individually. Errors during AddFormatsColumns are silently ignored because not
'             all tests require format columns.
Public Sub TestInitialize()
    BusyApp
    ResetDataSheet
    On Error Resume Next
        dataObject.AddFormatsColumns False, False, "formatting condition", "formatting values", "lock cells"
    On Error GoTo 0
End Sub

'@TestCleanUp
'@sub-title TestCleanUp
'@description Runs after every individual test method. Flushes the CustomTest assert buffer so
'             that each test's results are recorded independently.
Public Sub TestCleanUp()
    Assert.Flush
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Datasheet")
'@sub-title TestObjectInit
'@details Verifies that the DataSheet factory method (DataSheet.Create) correctly initializes all
'         read-only properties. Asserts that DataStartColumn, DataStartRow, Wksh.Name, HeaderRange
'         address, DataEndRow, and DataEndColumn all match the expected values derived from the
'         dictionary fixture dimensions. No arrange step beyond TestInitialize is needed; the act
'         is the factory call that already happened in ResetDataSheet.
Public Sub TestObjectInit()
    CustomTestSetTitles Assert, "Datasheet", "TestObjectInit"
    Assert.IsTrue (dataObject.DataStartColumn = 1), "Start column changed"
    Assert.IsTrue (dataObject.DataStartRow = 1), "Start line changed"
    Assert.IsTrue (dataObject.Wksh.Name = DICTIONARYFIXTURESHEET), "Dictionary name changed"
    Assert.IsTrue (dataObject.HeaderRange.Address = dataWorksheet.Range(dataWorksheet.Cells(1, 1), dataWorksheet.Cells(1, fixtureColumnCount)).Address), "Header Range address not correct"
    Assert.IsTrue (dataObject.DataEndRow = fixtureRowCount + 1), "End row not correct"
    Assert.IsTrue (dataObject.DataEndColumn = fixtureColumnCount), "End column not correct"
End Sub

'@TestMethod("Datasheet")
'@sub-title TestDataRange
'@details Exercises the DataRange method across several scenarios. First, retrieves a single
'         column ("Variable Name") and verifies the length matches the fixture row count. Then
'         retrieves the same column with includeHeaders:=True and confirms the extra header row.
'         Next, uses the "__all__" sentinel to fetch all columns and checks that the result is a
'         multidimensional BetterArray with the correct column count. Also validates a chunked
'         column name ("Control"). Finally, asserts that requesting a non-existent column
'         ("Formula") raises ProjectError.ElementNotFound.
Public Sub TestDataRange()
    CustomTestSetTitles Assert, "Datasheet", "TestDataRange"
    On Error GoTo Fail

    Dim values As BetterArray
    Dim firstRow As BetterArray
    Dim rng As Range

    Set values = New BetterArray

    values.FromExcelRange dataObject.DataRange("Variable Name")
    Assert.IsTrue (values.Length = fixtureRowCount), "Variable Name length is not equal to dictionary length"

    values.FromExcelRange dataObject.DataRange("Variable Name", includeHeaders:=True)
    Assert.IsTrue (values.Length = fixtureRowCount + 1), "Variable name length with headers included is not equal to dictionary length"

    values.FromExcelRange dataObject.DataRange("__all__", includeHeaders:=True)
    Assert.IsTrue (values.Length = fixtureRowCount + 1), "All the data length is not equal to the dictionary length"
    Assert.IsTrue (values.ArrayType = BA_MULTIDIMENSION), "All the dictionary data is not in multidimensional array"

    Set firstRow = New BetterArray
    firstRow.Items = values.Item(1)
    Assert.IsTrue (firstRow.Length = fixtureColumnCount), "Number of columns of the dictionary: " & firstRow.Length & " - Expected number of columns: " & fixtureColumnCount

    values.FromExcelRange dataObject.DataRange("Control")
    Assert.IsTrue (values.Length = fixtureRowCount), "Control and other chunked variable names are not completely extracted"

    On Error Resume Next
        Err.Clear
        '@Ignore AssignmentNotUsed
        Set rng = dataObject.DataRange("Formula")
        Assert.IsTrue (Err.Number = ProjectError.ElementNotFound), "Failed to raise error on unfound columns"
    On Error GoTo Fail

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDataRange", Err.Number, Err.Description
End Sub

'@TestMethod("DataSheet")
'@sub-title TestColumnExist
'@details Tests the ColumnExists method under five conditions: a garbage string that should not
'         match any header, an empty string, an exact case-sensitive match ("Variable Name"), a
'         case-insensitive match ("variable name" with matchCase:=False), and a partial
'         case-insensitive match ("variable" with strictSearch:=False). Each assertion confirms
'         the expected Boolean return value.
Public Sub TestColumnExist()
    CustomTestSetTitles Assert, "DataSheet", "TestColumnExist"
    Assert.IsFalse dataObject.ColumnExists("&222!\"), "Weird column Name found"
    Assert.IsFalse dataObject.ColumnExists(""), "Empty column name found"
    Assert.IsTrue dataObject.ColumnExists("Variable Name"), "Variable Name not found"
    Assert.IsTrue dataObject.ColumnExists("variable name", matchCase:=False), "Variable name not found when searching without case"
    Assert.IsTrue dataObject.ColumnExists("variable", matchCase:=False, strictSearch:=False), "Variable name not found when searching partially"
End Sub

'@TestMethod("DataSheet")
'@sub-title TestAddFormat
'@details Verifies that AddFormatsColumns completes without raising an error when called with
'         valid formatting column names ("formatting condition" and "formatting values"). The
'         test uses an On Error GoTo Fail pattern: if the call succeeds the test exits normally;
'         if it throws, the failure is logged with CustomTestLogFailure.
Public Sub TestAddFormat()
    CustomTestSetTitles Assert, "DataSheet", "TestAddFormat"
    On Error GoTo Fail

    dataObject.AddFormatsColumns False, False, "formatting condition", "formatting values"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddFormat", Err.Number, Err.Description
End Sub

'@TestMethod("DataSheet")
'@sub-title TestSimpleFilter
'@details Validates single-condition filtering via FilterData. Arranges three scenarios: (1)
'         filtering "Sheet Type" for "hlist2D" and returning "Variable Name" should yield a
'         non-empty result; (2) the same filter returning "__all__" should produce a
'         multidimensional BetterArray; (3) filtering on a nonsense value ("&&&&&") should return
'         an empty array. Additionally verifies the error path: requesting columns that do not
'         exist ("Sheet", "OO") raises ProjectError.ElementNotFound.
Public Sub TestSimpleFilter()
    CustomTestSetTitles Assert, "DataSheet", "TestSimpleFilter"
    On Error GoTo Fail

    Dim values As BetterArray

    Set values = dataObject.FilterData("Sheet Type", "hlist2D", "Variable Name")
    Assert.IsTrue (values.Length > 0), "Filtering on 2D worksheets result in error"

    Set values = dataObject.FilterData("Sheet Type", "hlist2D", "__all__")
    Assert.IsTrue (values.ArrayType = BA_MULTIDIMENSION), "unable to filter all the data on one condition"

    Set values = dataObject.FilterData("Sheet Name", "&&&&&", "Variable Name")
    Assert.IsTrue (values.Length = 0), "Unable to filter on unfound values"

    On Error Resume Next
        Err.Clear
        '@Ignore AssignmentNotUsed
        Set values = dataObject.FilterData("Sheet", "Test", "OO")
        Assert.IsTrue (Err.Number = ProjectError.ElementNotFound), "Failed to raise error on unfound columns"
    On Error GoTo Fail

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSimpleFilter", Err.Number, Err.Description
End Sub

'@TestMethod("DataSheet")
'@sub-title TestMultipleFilters
'@details Exercises the multi-condition FiltersData method across four scenarios. First, applies
'         two valid conditions (Sheet Name = "hlist2D-sheet1" AND Main Section = "Validation")
'         and asserts a non-empty result. Second, uses nonsense condition values to confirm an
'         empty result. Third, pops one element from variableData to create a length mismatch
'         between variables and conditions, verifying that FiltersData handles this gracefully by
'         returning an empty array. Fourth, passes completely unknown column names and asserts
'         that ProjectError.ElementNotFound is raised.
Public Sub TestMultipleFilters()
    CustomTestSetTitles Assert, "DataSheet", "TestMultipleFilters"
    On Error GoTo Fail

    Dim returnedValues As BetterArray
    Dim variableData As BetterArray
    Dim conditionData As BetterArray
    Dim returnData As BetterArray

    Set variableData = BetterArrayFromList("Sheet Name", "Main Section")
    Set conditionData = BetterArrayFromList("hlist2D-sheet1", "Validation")
    Set returnData = BetterArrayFromList("Variable Name", "Sheet Type")

    Set returnedValues = dataObject.FiltersData(variableData, conditionData, returnData)
    Assert.IsTrue (returnedValues.Length > 0), "unable to multiple filter on known returnedValues"
    Set returnedValues = Nothing

    Set conditionData = BetterArrayFromList("&&&&", "AAAA")
    Set returnedValues = dataObject.FiltersData(variableData, conditionData, returnData)
    Assert.IsTrue (returnedValues.Length = 0), "Unable to multiple filter on unknown returnedValues"
    Set returnedValues = Nothing

    variableData.Pop
    Set returnedValues = dataObject.FiltersData(variableData, conditionData, returnData)
    Assert.IsTrue (returnedValues.Length = 0), "Filters should handle mismatched variable and condition counts"
    Set returnedValues = Nothing

    On Error Resume Next
    '@Ignore AssignmentNotUsed
    Set returnedValues = dataObject.FiltersData(BetterArrayFromList("Unknown"), BetterArrayFromList("Unknown"), returnData)
    Assert.IsTrue (Err.Number = ProjectError.ElementNotFound), "FiltersData should raise ElementNotFound for unknown columns"
    On Error GoTo Fail

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestMultipleFilters", Err.Number, Err.Description
End Sub

'@TestMethod("DataSheet")
'@sub-title TestExport
'@details Tests that Export copies the DataSheet contents and formatting into a new workbook.
'         Arranges by creating a fresh workbook, then acts by calling dataObject.Export. Asserts
'         that (1) the exported workbook contains a worksheet with the same name as the source,
'         (2) the workbook has at least one worksheet, and (3) cell-level formatting (interior
'         color) in the last data cell matches between source and export. The temporary workbook
'         is deleted in both the success and failure paths.
Public Sub TestExport()
    CustomTestSetTitles Assert, "DataSheet", "TestExport"
    On Error GoTo Fail

    Dim exportBook As Workbook
    Dim exportedSheet As Worksheet

    Set exportBook = NewWorkbook()
    dataObject.Export exportBook

    On Error Resume Next
        Set exportedSheet = exportBook.Worksheets(dataObject.Wksh.Name)
    On Error GoTo Fail

    Assert.IsFalse (exportedSheet Is Nothing), "Dictionary not exported"
    Assert.IsTrue (exportBook.Worksheets.Count >= 1), "Export should create at least one worksheet"
    Assert.AreEqual dataObject.Wksh.Cells(fixtureRowCount + 1, fixtureColumnCount).Interior.Color, _
                   exportedSheet.Cells(fixtureRowCount + 1, fixtureColumnCount).Interior.Color, _
                   "Formatting not exported"

    DeleteWorkbook exportBook
    Exit Sub

Fail:
    If Not exportBook Is Nothing Then DeleteWorkbook exportBook
    CustomTestLogFailure Assert, "TestExport", Err.Number, Err.Description
End Sub

'@TestMethod("DataSheet")
'@sub-title TestExportIncludesHiddenNamesWhenRequested
'@details Verifies that Export with includeNames:=True replicates worksheet-level hidden names to
'         the target workbook. Arranges by creating a HiddenNames store on the source sheet and
'         setting a Long value (42) under the key "__DataSheetHidden__". Acts by calling Export
'         with includeNames:=True into a new workbook. Asserts that a HiddenNames store on the
'         exported sheet returns the same Long value for the same key. Cleans up the temporary
'         workbook in both success and failure paths.
Public Sub TestExportIncludesHiddenNamesWhenRequested()
    CustomTestSetTitles Assert, "DataSheet", "TestExportIncludesHiddenNamesWhenRequested"
    On Error GoTo Fail

    Dim exportBook As Workbook
    Dim sourceStore As IHiddenNames
    Dim exportedStore As IHiddenNames
    Const NAME_ID As String = "__DataSheetHidden__"

    Set sourceStore = HiddenNames.Create(dataObject.Wksh)
    sourceStore.EnsureName NAME_ID, 42, HiddenNameTypeLong
    sourceStore.SetValue NAME_ID, 42

    Set exportBook = NewWorkbook()
    dataObject.Export exportBook, includeNames:=True

    Set exportedStore = HiddenNames.Create(exportBook.Worksheets(dataObject.Name))
    Assert.AreEqual 42, exportedStore.ValueAsLong(NAME_ID, -1), _
                     "Export should replicate hidden names when includeNames is True."

    DeleteWorkbook exportBook
    Exit Sub

Fail:
    If Not exportBook Is Nothing Then DeleteWorkbook exportBook
    CustomTestLogFailure Assert, "TestExportIncludesHiddenNamesWhenRequested", Err.Number, Err.Description
End Sub

'@TestMethod("DataSheet")
'@sub-title TestImport
'@details Tests data import with case-insensitive column matching. Arranges by creating a target
'         sheet with lower-cased copies of the source headers to simulate a case mismatch. Acts
'         by calling importData.Import with strictColumnSearch:=False, which forces the DataSheet
'         to match columns in a case-insensitive manner. Asserts that the imported sheet has the
'         correct HeaderRange address, DataEndRow, and DataEndColumn matching the original fixture
'         dimensions. Also exercises ImportFormat by importing the "Formatting Values" column to
'         confirm no errors are raised during format transfer.
Public Sub TestImport()
    CustomTestSetTitles Assert, "DataSheet", "TestImport"
    On Error GoTo Fail

    Dim outputSheet As Worksheet
    Dim headerArray As BetterArray
    Dim importData As IDataSheet
    Dim columnIndex As Long

    Set outputSheet = EnsureWorksheet(DICTOUTPUTSHEET)
    ClearWorksheet outputSheet

    Set headerArray = New BetterArray
    headerArray.FromExcelRange dataObject.HeaderRange
    headerArray.ToExcelRange outputSheet.Cells(1, 1), TransposeValues:=True

    For columnIndex = 1 To headerArray.Length
        outputSheet.Cells(1, columnIndex).Value = LCase$(CStr(outputSheet.Cells(1, columnIndex).Value))
    Next columnIndex

    Set importData = DataSheet.Create(outputSheet, 1, 1)
    importData.Import dataObject, strictColumnSearch:=False

    Assert.AreEqual importData.HeaderRange.Address,  outputSheet.Range(outputSheet.Cells(1, 1), outputSheet.Cells(1, fixtureColumnCount)).Address, "Header Range address not correct"
    Assert.IsTrue (importData.DataEndRow = fixtureRowCount + 1), "End row not correct"
    Assert.IsTrue (importData.DataEndColumn = fixtureColumnCount), "End column not correct"

    importData.ImportFormat dataObject.DataRange("Formatting Values", includeHeaders:=True)
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImport", Err.Number, Err.Description
End Sub
