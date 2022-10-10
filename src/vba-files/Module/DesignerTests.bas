Attribute VB_Name = "DesignerTests"
Option Explicit
Option Private Module

Sub ShowWindows()
    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    EndWork xlsapp:=Application
End Sub


Sub testform()
    Dim choicewksh As Worksheet
    Dim choiceDict As ILLdictionary
    Dim choiceWorksheet As Worksheet
    Dim choice As ILLchoice
    Dim Cat As BetterArray
    Dim rng As Range
    
    Set choiceWorksheet = ThisWorkbook.Worksheets("Choices")

    Set choiceDict = LLdictionary.Create(choiceWorksheet, 1, 1)
    Set choice = LLchoice.Create(choiceDict)
    Set rng = choice.ChoiceDictionary.DataRange
    
    Set Cat = New BetterArray
    Cat.Push "simple", "test"
    
    Set Cat = choice.Categories("list_a1")
    
    Debug.Print choice.StartRow
    Debug.Print choice.StartColumn
    Debug.Print choice.Wksh.Name
    Debug.Print choice.DataRange().Address
End Sub
