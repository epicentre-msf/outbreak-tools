Attribute VB_Name = "TestDataSheet"
Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulNames
'@TestModule
'@Folder("Tests")

Private Const DICTIONARYFIXTURESHEET As String = "LLDictTest"
Private Const DICTOUTPUTSHEET As String = "DataOut"

Private fixtureRowCount As Long
Private fixtureColumnCount As Long

Private Assert As Object
Private Fakes As Object
Private dataObject As IDataSheet
Private dataWorksheet As Worksheet

'@section Helpers
'===============================================================================

Private Sub ResetDataSheet()
    PrepareDictionaryFixture DICTIONARYFIXTURESHEET
    Set dataWorksheet = ThisWorkbook.Worksheets(DICTIONARYFIXTURESHEET)
    Set dataObject = DataSheet.Create(dataWorksheet, 1, 1)
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ResetDataSheet
    EnsureWorksheet DICTOUTPUTSHEET
    fixtureRowCount = DictionaryFixtureRowCount()
    fixtureColumnCount = DictionaryFixtureColumnCount()
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    DeleteWorksheet DICTOUTPUTSHEET
    DeleteWorksheet DICTIONARYFIXTURESHEET
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    ResetDataSheet
    On Error Resume Next
        dataObject.AddFormatsColumns False, "formatting condition", "formatting values"
    On Error GoTo 0
End Sub

'@section Tests
'===============================================================================

'@TestMethod("Datasheet")
Private Sub TestObjectInit()
    Assert.IsTrue (dataObject.DataStartColumn = 1), "Start column changed"
    Assert.IsTrue (dataObject.DataStartRow = 1), "Start line changed"
    Assert.IsTrue (dataObject.Wksh.Name = DICTIONARYFIXTURESHEET), "Dictionary name changed"
    Assert.IsTrue (dataObject.HeaderRange.Address = dataWorksheet.Range(dataWorksheet.Cells(1, 1), dataWorksheet.Cells(1, fixtureColumnCount)).Address), "Header Range address not correct"
    Assert.IsTrue (dataObject.DataEndRow = fixtureRowCount + 1), "End row not correct"
    Assert.IsTrue (dataObject.DataEndColumn = fixtureColumnCount), "End column not correct"
End Sub

'@TestMethod("Datasheet")
Private Sub TestDataRange()
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
    FailUnexpectedError Assert, "TestDataRange"
End Sub

'@TestMethod("DataSheet")
Private Sub TestColumnExist()
    Assert.IsFalse dataObject.ColumnExists("&222!\"), "Weird column Name found"
    Assert.IsFalse dataObject.ColumnExists(""), "Empty column name found"
    Assert.IsTrue dataObject.ColumnExists("Variable Name"), "Variable Name not found"
    Assert.IsTrue dataObject.ColumnExists("variable name", matchCase:=False), "Variable name not found when searching without case"
    Assert.IsTrue dataObject.ColumnExists("variable", matchCase:=False, strictSearch:=False), "Variable name not found when searching partially"
End Sub

'@TestMethod("DataSheet")
Private Sub TestAddFormat()
    On Error GoTo Fail

    dataObject.AddFormatsColumns False, "formatting condition", "formatting values"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestAddFormat"
End Sub

'@TestMethod("DataSheet")
Private Sub TestSimpleFilter()
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
    FailUnexpectedError Assert, "TestSimpleFilter"
End Sub

'@TestMethod("DataSheet")
Private Sub TestMultipleFilters()
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
    FailUnexpectedError Assert, "TestMultipleFilters"
End Sub

'@TestMethod("DataSheet")
Private Sub TestExport()
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
    FailUnexpectedError Assert, "TestExport"
End Sub

'@TestMethod("DataSheet")
Private Sub TestImport()
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
    FailUnexpectedError Assert, "TestImport"
End Sub
