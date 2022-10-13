Attribute VB_Name = "TestDictionary"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")


Private Assert As Object
Private Fakes As Object
Private Dictionary As ILLdictionary

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
    Set Dictionary = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Dim dataWksh As Worksheet
    'This method runs before every test in the module..
    Set dataWksh = ThisWorkbook.Worksheets("TestDictionary")
    Set Dictionary = LLdictionary.Create(dataWksh, 1, 1)
End Sub

'@TestMethod
Private Sub TestObjectInit()
    Assert.IsTrue (Dictionary.Data.StartColumn = 1), "Start column changed"
    Assert.IsTrue (Dictionary.Data.StartRow = 1), "Start line changed"
    Assert.IsTrue (Dictionary.Data.Wksh.Name = "TestDictionary"), "Dictionary name changed"
End Sub



'@TestMethod
Private Sub TestUniqueValues()

    On Error GoTo UniqueValuesFailed

    Dim sheetsList As BetterArray
    Set sheetsList = Dictionary.UniqueValues("sheet name")

    Assert.IsTrue (sheetsList.Length = 3), "Unable to find all the unique values of sheet names"
    Assert.IsTrue (sheetsList.Includes("A, B, C")), "Unable to find A, B, C sheet"
    Assert.IsTrue (sheetsList.Includes("C, B, A")), "Unable to find C, B, A sheet"
    Assert.IsTrue (sheetsList.Includes("B-H2D")), "Unable to find B-H2D sheet"

    Exit Sub

UniqueValuesFailed:
    Assert.Fail "Unique values Failed: #" & Err.Number & " : " & Err.Description
End Sub

'@TestMethod
Private Sub TestColumnExist()
    On Error GoTo ColumnExistFailed

    Assert.IsTrue Dictionary.ColumnExists("variable name"), "variable name not found"
    Assert.IsTrue (Not Dictionary.ColumnExists("random column for testing")), "random column found"
    Assert.IsFalse Dictionary.ColumnExists("column indexes", checkValidity:=True), "column indexes which is not present, is found"


    Exit Sub

ColumnExistFailed:
    Assert.Fail "ColumnExist Failed: #" & Err.Number & " : " & Err.Description

End Sub



'@TestMethod
Private Sub TestPreparation()

    On Error GoTo PreparationFailed
    Dim dictWksh As Worksheet
    Dim dictRng As Range
    Dim randRng As Range
    Dim endCol As Long

    Set dictWksh = Dictionary.Data.Wksh

    If Not Dictionary.Prepared Then
        With dictWksh
            endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column + 1
            If Not Dictionary.ColumnExists("randnumber") Then
                .Cells(1, endCol) = "randnumber"
                .Cells(2, endCol).Formula = "= RAND()"
                Set randRng = Dictionary.DataRange("randnumber")
                .Cells(2, endCol).AutoFill randRng, Type:=xlFillValues
            End If
            Set dictRng = Dictionary.DataRange
            Set randRng = Dictionary.DataRange("randnumber")
            dictRng.Sort key1:=randRng
            Dictionary.Prepare
        End With
    End If

    Assert.IsTrue Dictionary.Prepared, "dictionary not prepared for buildlist"
    Exit Sub

PreparationFailed:
    Assert.Fail "Prepared Failed: #" & Err.Number & " : " & Err.Description
End Sub
