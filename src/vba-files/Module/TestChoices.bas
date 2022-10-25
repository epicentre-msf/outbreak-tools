Attribute VB_Name = "TestChoices"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private formCond As IFormulaCondition
Private choice As ILLchoice


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
    Dim choicewksh As Worksheet
    Dim choiceWorksheet As Worksheet

    Set choiceWorksheet = ThisWorkbook.Worksheets("TestChoices")
    Set choice = LLchoice.Create(choiceWorksheet, 1, 1)
End Sub


'@TestMethod
Private Sub TestInit()

    On Error GoTo InitFailed

    Assert.IsTrue (choice.StartRow = 1), "Bad choice startRow"
    Assert.IsTrue (choice.StartColumn = 1), "Bad choice startcolumn"
    Assert.IsTrue (choice.Wksh.Name = "TestChoices"), "Bad choice worksheet"

    Exit Sub
InitFailed:
    Assert.Fail "Init Failed: #" & Err.Number & " : " & Err.Description
End Sub

'@TestMethod
Private Sub TestSort()
    On Error GoTo SortFailed
    choice.Sort
Exit Sub

SortFailed:
    Assert.Fail "Sort Failed: #" & Err.Number & " : " & Err.Description
End Sub


'@TestMethod
Private Sub TestAddChoice()
    Dim cat As BetterArray
    Set cat = New BetterArray

    cat.Push "simple", "test"
    On Error GoTo AddFailed

    choice.AddChoice "list_test", cat
Exit Sub

AddFailed:
    Assert.Fail "Sort Failed: #" & Err.Number & " : " & Err.Description
End Sub


'@TestMethod
Private Sub TestDataRange()
    On Error GoTo DataRangeFailed

    Assert.IsTrue (choice.DataRange("list name").Column = 1), "Bad list name column returned"
    Assert.IsTrue (choice.DataRange("label").Column = 3), "Bad label column returned"
Exit Sub

DataRangeFailed:
    Assert.Fail "Sort Failed: #" & Err.Number & " : " & Err.Description

End Sub
