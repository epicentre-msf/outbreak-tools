Attribute VB_Name = "TestDataSheet"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private dataObject As IDataSheet
Private dataWksh As Worksheet

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    Set dataWksh = ThisWorkbook.Worksheets("TestDictionary")
    Set dataObject = DataSheet.Create(dataWksh, 1, 1)
End Sub

'@TestMethod
Private Sub TestObjectInit()
    Assert.IsTrue (dataObject.StartColumn = 1), "Start column changed"
    Assert.IsTrue (dataObject.StartRow = 1), "Start line changed"
    Assert.IsTrue (dataObject.Wksh.Name = "TestDictionary"), "Dictionary name changed"
End Sub

'@TestMethod
Private Sub TestColumnName()

    Dim var As BetterArray
    Dim var1 As BetterArray
    Dim rng As Range
    Dim nbRows As Long

    On Error GoTo ColumnFail

    nbRows = 50

    'variable names only
    Set var = New BetterArray
    var.FromExcelRange dataObject.DataRange("variable name")
    Assert.IsTrue (var.Length = nbRows), "Variable Name length is not equal to dictionary length"

    'variable names with the header
    var.FromExcelRange dataObject.DataRange("variable name", includeHeaders:=True)
    Assert.IsTrue (var.Length = nbRows + 1), "Variable name length with headers included is not equal to dictionary length"

    'all the dictionary data
    var.FromExcelRange dataObject.DataRange("__all__", includeHeaders:=True)
    Assert.IsTrue (var.Length = nbRows + 1), "All the data length is not equal to the dictionary Length"
    Assert.IsTrue (var.ArrayType = BA_MULTIDIMENSION), "All the dictionary data is not in multidimensional array"

    Set var1 = New BetterArray
    var1.Items = var.Item(1)
    Assert.IsTrue (var1.Length = 24 Or var1.Length = 23 Or var1.Length = 25), "Number of columns of the dictionary: " & var1.Length & " - " & "Expected number of columns: " & "23, 24 or 25"

    'chuncked variables like control (variables with whole within values)
    var.FromExcelRange dataObject.DataRange("control")
    Assert.IsTrue (var.Length = nbRows), "Control and other chunked variable names are not complety extracted"


    'unfound variable
    On Error GoTo UnFoundFail
    var.FromExcelRange dataObject.DataRange("Formula")

    Exit Sub

UnFoundFail:
    Assert.IsTrue (Err.Description = "Column Formula does not exists in worksheet TestDictionary"), "Failed to raise error on unfound columns"
    Exit Sub

ColumnFail:
    Assert.Fail "Test raised an error: #" & Err.Number & "-" & Err.Description
End Sub

'@TestMethod
Private Sub TestColumnExist()
    Assert.IsFalse dataObject.ColumnExists("&222!\"), "Weird column Name found"
    Assert.IsFalse dataObject.ColumnExists(""), "Empty column name found"
    Assert.IsTrue dataObject.ColumnExists("variable name"), "Variable Name not found"
End Sub

'@TestMethod
Private Sub TestSimpleFilter()

    On Error GoTo SimpleFilterFail

    Dim var As BetterArray
    Set var = New BetterArray

    'Filter on found variables
    Set var = dataObject.FilterData("sheet type", "hlist2D", "variable name")
    Assert.IsTrue (var.Length > 0), "no 2D linelist found"
    'Filter on all the data
    Set var = dataObject.FilterData("sheet type", "hlist2D", "__all__")
    Assert.IsTrue (var.ArrayType = BA_MULTIDIMENSION), "unable to filter all the data on one condition"

    'Filter on unfound values
    Set var = dataObject.FilterData("sheet name", "&&&&&", "variable name")
    Assert.AreEqual CInt(var.Length), 0, "Unable to filter on unfound values"

    'Filter on unfound columns
    On Error GoTo UnFoundFail
    Set var = dataObject.FilterData("sheet", "Test", "OO")
    Exit Sub

UnFoundFail:
    Assert.IsTrue (Err.Description = "Column OO does not exists in worksheet TestDictionary"), "Failed to raise error on unfound columns"
    Exit Sub

SimpleFilterFail:
    Assert.Fail "Simple filter raised an error: #" & Err.Number & " : " & Err.Description
End Sub

'@TestMethod
Private Sub TestMultipleFilters()

    On Error GoTo MultipleFiltersFail

    Dim var As BetterArray
    Dim varData As BetterArray
    Dim condData As BetterArray
    Dim retrData As BetterArray

    Set var = New BetterArray
    Set varData = New BetterArray
    Set condData = New BetterArray
    Set retrData = New BetterArray

    'Found values
    varData.Push "sheet name", "sub section"
    condData.Push "A, B, C", "Sub section 1"
    retrData.Push "variable name", "sheet type"
    Set var = dataObject.FiltersData(varData, condData, retrData)
    Assert.IsTrue (var.Length > 0), "unable to filter on found values"

    'Unfound values
    condData.Clear
    condData.Push "&&&&", "AAAA"
    Set var = dataObject.FiltersData(varData, condData, retrData)
    Assert.IsTrue (var.Length = 0), "Unable to filter on Unfound values"

    'Number of conditions not equal number of variables
    varData.Pop
    Set var = dataObject.FiltersData(varData, condData, retrData)
    Assert.IsTrue (var.Length = 0), "Unable to filter when number of conditions <> number of variables"

    On Error GoTo UnFoundFail
    'Unfound variables
    varData.Clear
    varData.Push "AAAA", "BBBB"
    condData.Clear
    condData.Push "A, B, C", "Sub section 1"
    Set var = dataObject.FiltersData(varData, condData, retrData)
    Assert.IsTrue (var.Length = 0), "Unable to filter on Unfound variables"

    Exit Sub

UnFoundFail:
    Assert.IsTrue (Err.Description = "Column AAAA does not exists in worksheet TestDictionary"), "Failed to raise error on unfound columns"
    Exit Sub

MultipleFiltersFail:
    Assert.Fail "Multiple filters raised an error: #" & Err.Number & " : " & Err.Description
End Sub

