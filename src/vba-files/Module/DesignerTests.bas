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
    Dim Table As ICrossTable
    Dim specs As ITablesSpecs
    Dim counter As Long
    Dim formData As IFormulaData
    Dim TableFormula As ICrossTableFormula
    Dim designerFormat As ILLFormat
    

    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("TestDictionary"), 1, 1)
    dict.Prepare
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("ControleFormule")
    Set formData = FormulaData.Create(sh, "T_XlsFonctions", "T_ascii")
    Set choi = LLchoice.Create(ThisWorkbook.Worksheets("TestChoices"), 1, 1)
    Set lData = LinelistSpecs.Create(dict, choi)
    Set Lo = ThisWorkbook.Worksheets("Analysis").ListObjects(2)
    Set headRng = Lo.HeaderRowRange
    Set rowRng = Lo.ListRows(1).Range
    Set trad = Translation.Create(ThisWorkbook.Worksheets("LinelistTranslation").ListObjects("T_TradLLMsg"), "FRA")
    
    Set specs = TablesSpecs.Create(headRng, rowRng, lData)

    Set Table = CrossTable.Create(specs, ThisWorkbook.Worksheets("TestAnalysis"), trad)
    Set TableFormula = CrossTableFormula.Create(Table, formData)
    Set sh = ThisWorkbook.Worksheets("LinelistStyle")
    Set designerFormat = LLFormat.Create(sh, "design 1")
    
    Table.AddHeader
    Table.AddRows
    Table.AddColumns
    Table.NameRanges
    Table.Format designerFormat
    TableFormula.AddFormulas
    
    
   
    'Debug.Print table.RowsCategoriesRange.Address
    
End Sub


Sub TestGraph()
    Dim co As ChartObject
    Dim rng As Range
    Set rng = SheetTest.Range("N12")
    Dim cw As Long
    Dim rh As Long
    
    rh = SheetTest.Range("A1").Height
    cw = SheetTest.Range("A1").Width
    Debug.Print rng.Left
    Debug.Print rng.Top
    Set co = SheetTest.ChartObjects.Add(rng.Left, rng.Top, cw * 8, rh * 20)
    'Debug.Print co.Left
    'Debug.Print co.Top
    With co.Chart.PlotArea
        .Interior.color = RGB(235, 235, 235)
    End With
    
    
End Sub
