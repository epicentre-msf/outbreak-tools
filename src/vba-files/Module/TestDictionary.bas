Attribute VB_Name = "TestDictionary"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private dictObject As ILLdictionary
Private dictWksh As Worksheet

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
    Set dictWksh = ThisWorkbook.Worksheets("Dictionary")
    Set dictObject = LLdictionary.Create(dictWksh, 1, 1)
End Sub

'@TestMethod
Private Sub TestObjectInit()
    Assert.AreEqual CInt(dictObject.StartColumn), 1, "Start column changed"
    Assert.AreEqual CInt(dictObject.StartLine), 1, "Start line changed"
End Sub

'@TestMethod
Private Sub TestColumnName()
    
    Dim var As BetterArray
    Dim rng As Range
    
    On Error GoTo ColumnFail
    
    Set var = New BetterArray
    Set var = dictObject.Column("Variable Name")
    
    Assert.AreEqual CInt(var.Length), 47, "Variable Name length is not equal to dictionary length"
    Set var = dictObject.Column("Formula")
    Assert.AreEqual CInt(var.Length), 0, "Unfound variable does not result to empty BA"
    
    Set var = dictObject.Column("Control")
    Assert.AreEqual CInt(var.Length), 47, "Control and other chunked variable names are not complety extracted"
    
    Exit Sub
    
ColumnFail:
    Assert.Fail "Test raised an error: #" & Err.Number & "-" & Err.Description
End Sub


'@TestMethod
Private Sub TestColumnExist()
    Assert.isFalse dictObject.ColumnExists("&222!\"), "Weird column Name found"
    Assert.isFalse dictObject.ColumnExists(""), "Empty column name found"
    Assert.IsTrue dictObject.ColumnExists("Variable Name"), "Variable Name not found"
End Sub

'@TestMethod
Private Sub TestSimpleFilter()

    On Error GoTo SimpleFilterFail
    
    Dim var As BetterArray
    Set var = New BetterArray
    
    Set var = dictObject.FilterData("Sheet Type", "hlist2D", "Variable Name")
    Assert.IsTrue (var.Length > 0), "no 2D linelist found"
    Set var = dictObject.FilterData("Sheet Name", "&&&&&", "Variable Name")
    Assert.AreEqual CInt(var.Length), 0, "Unable to filter on unfound values"
    Set var = dictObject.FilterData("Sheet", "Test", "OO")
    Assert.AreEqual CInt(var.Length), 0, "Unable to filter on unfound columns"
    
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
    varData.Push "Sheet Name", "Sub Section"
    condData.Push "A, B, C", "Sub section 1"
    retrData.Push "Variable Name", "Sheet Type"
    Set var = dictObject.FiltersData(varData, condData, retrData)
    Assert.IsTrue (var.Length > 0), "unable to filter on found values"
    
    'Unfound values
    condData.Clear
    condData.Push "&&&&", "AAAA"
    Set var = dictObject.FiltersData(varData, condData, retrData)
    Assert.IsTrue (var.Length = 0), "Unable to filter on Unfound values"
    
    'Unfound variables
    varData.Clear
    varData.Push "AAAA", "BBBB"
    condData.Clear
    condData.Push "A, B, C", "Sub section 1"
    Set var = dictObject.FiltersData(varData, condData, retrData)
    Assert.IsTrue (var.Length = 0), "Unable to filter on Unfound variables"
    
Exit Sub

MultipleFiltersFail:
    Assert.Fail "Multiple filters raised an error: #" & Err.Number & " : " & Err.Description
End Sub


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set dictWksh = Nothing
    Set dictObject = Nothing
End Sub
