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
    Assert.AreEqual CInt(dictObject.StartRow), 1, "Start line changed"
End Sub

'@TestMethod
Private Sub TestColumnName()

    Dim var As BetterArray
    Dim var1 As BetterArray
    Dim rng As Range
    Dim nbRows As Long

    On Error GoTo ColumnFail

    nbRows = IIf(dictObject.Prepared, 50, 47)

    'variable names only
    Set var = New BetterArray
    Set var = dictObject.Column("variable name")
    Assert.IsTrue (var.Length = nbRows), "Variable Name length is not equal to dictionary length"

    'variable names with the header
    Set var = dictObject.Column("variable name", includeHeaders:=True)
    Assert.IsTrue (var.Length = nbRows + 1), "Variable name length with headers included is not equal to dictionary length"

    'all the dictionary data
    Set var = dictObject.Column("__all__", includeHeaders:=True)
    Assert.IsTrue (var.Length = nbRows + 1), "All the data length is not equal to the dictionary Length"
    Assert.IsTrue (var.ArrayType = BA_MULTIDIMENSION), "All the dictionary data is not in multidimensional array"

    Set var1 = New BetterArray
    var1.Items = var.Item(1)
    Assert.IsTrue (var1.Length = 24 Or var1.Length = 23 Or var1.Length = 25), "Number of columns of the dictionary: " & var1.Length & " - " & "Expected number of columns: " & "23, 24 or 25"

    'unfound variable
    Set var = dictObject.Column("Formula")
    Assert.IsTrue (var.Length = 0), "Unfound variable does not result to empty BA"

    Set var = dictObject.Column("control")
    Assert.IsTrue (var.Length = nbRows), "Control and other chunked variable names are not complety extracted"

    Exit Sub

ColumnFail:
    Assert.Fail "Test raised an error: #" & Err.Number & "-" & Err.Description
End Sub


'@TestMethod
Private Sub TestColumnExist()
    Assert.IsFalse dictObject.ColumnExists("&222!\"), "Weird column Name found"
    Assert.IsFalse dictObject.ColumnExists(""), "Empty column name found"
    Assert.IsTrue dictObject.ColumnExists("variable name"), "Variable Name not found"
End Sub

'@TestMethod
Private Sub TestSimpleFilter()

    On Error GoTo SimpleFilterFail

    Dim var As BetterArray
    Set var = New BetterArray

    Set var = dictObject.FilterData("sheet type", "hlist2D", "variable name")
    Assert.IsTrue (var.Length > 0), "no 2D linelist found"
    Set var = dictObject.FilterData("sheet type", "hlist2D", "__all__")
    Assert.IsTrue (var.ArrayType = BA_MULTIDIMENSION), "unable to filter all the data on one condition"
    Set var = dictObject.FilterData("sheet name", "&&&&&", "variable name")
    Assert.AreEqual CInt(var.Length), 0, "Unable to filter on unfound values"
    Set var = dictObject.FilterData("sheet", "Test", "OO")
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
    varData.Push "sheet name", "sub section"
    condData.Push "A, B, C", "Sub section 1"
    retrData.Push "variable name", "sheet type"
    Set var = dictObject.FiltersData(varData, condData, retrData)
    Assert.IsTrue (var.Length > 0), "unable to filter on found values"

    'Unfound values
    condData.Clear
    condData.Push "&&&&", "AAAA"
    Set var = dictObject.FiltersData(varData, condData, retrData)
    Assert.IsTrue (var.Length = 0), "Unable to filter on Unfound values"

    'Number of conditions not equal number of variables
    varData.Pop
    Set var = dictObject.FiltersData(varData, condData, retrData)
    Assert.IsTrue (var.Length = 0), "Unable to filter when number of conditions <> number of variables"

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


'@TestMethod
Private Sub TestPreparation()

    On Error GoTo PreparationFailed
    Dim dictWksh As Worksheet
    Dim dictRng As Range
    Dim randRng As Range
    Dim endCol As Long

    Set dictWksh = dictObject.Wksh

    If Not dictObject.Prepared Then
        With dictWksh
            endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column + 1
            If Not dictObject.ColumnExists("randnumber") Then
                .Cells(1, endCol) = "randnumber"
                .Cells(2, endCol).Formula = "= RAND()"
                Set randRng = dictObject.DataRange("randnumber")
                .Cells(2, endCol).AutoFill randRng, Type:=xlFillValues
            End If
            Set dictRng = dictObject.DataRange
            Set randRng = dictObject.DataRange("randnumber")
            dictRng.Sort key1:=randRng
            dictObject.Prepare
        End With
    End If

    Assert.IsTrue dictObject.Prepared, "dictionary not prepared for buildlist"
    Exit Sub

PreparationFailed:
    Assert.Fail "Prepared Failed: #" & Err.Number & " : " & Err.Description
End Sub

