Attribute VB_Name = "DesignerTests"
Option Explicit
Option Private Module

Sub ShowWindows()
    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    EndWork xlsapp:=Application
End Sub


Sub test()
    Dim trad As ITranslation
    Dim dict As ILLdictionary
    Dim choi As ILLchoice
    Dim headRng As Range
    Dim rowRng As Range
    Dim Lo As ListObject
    Dim lData As ILinelistSpecs
    Dim table As ICrossTable
    Dim specs As ITablesSpecs
    

    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("TestDictionary"), 1, 1)
    Set choi = LLchoice.Create(ThisWorkbook.Worksheets("TestChoices"), 1, 1)
    Set lData = LinelistSpecs.Create(dict, choi)
    Set Lo = ThisWorkbook.Worksheets("Analysis").ListObjects(3)
    Set headRng = Lo.HeaderRowRange
    Set rowRng = Lo.ListRows(1).Range
    Set trad = Translation.Create(ThisWorkbook.Worksheets("LinelistTranslation").ListObjects("T_TradLLMsg"), "FRA")
    
    Set specs = TablesSpecs.Create(headRng, rowRng, lData)

    Set table = CrossTable.Create(specs, ThisWorkbook.Worksheets("TestAnalysis"), trad)
    
    table.AddHeader
    table.AddRows
    table.AddColumns
    table.NameRanges
    
    
   
    'Debug.Print table.RowsCategoriesRange.Address
    
End Sub
