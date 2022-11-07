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
    Dim headRng As Range
    Dim rowRng As Range
    Dim Lo As ListObject
    Dim lData As ILinelistSpecs
    Dim Table As ICrossTable
    Dim specs As ITablesSpecs
    Dim TableFormula As ICrossTableFormula
    Dim Wkb As Workbook
    Dim grSpecs As IGraphSpecs
    Dim gr As IGraphs
    Dim sh As Worksheet
    Dim counter As Long
    Dim nbSeries As Long
    
    Set Wkb = ThisWorkbook
    Set lData = LinelistSpecs.Create(Wkb)
    
    Set sh = ThisWorkbook.Worksheets("TestAnalysis")
    Set Lo = ThisWorkbook.Worksheets("Analysis").ListObjects(4)
    Set headRng = Lo.HeaderRowRange
    Set rowRng = Lo.ListRows(2).Range
    Set trad = lData.TransObject()               'Form translation object
    lData.Prepare
    Set specs = TablesSpecs.Create(headRng, rowRng, lData)

    Set Table = CrossTable.Create(specs, ThisWorkbook.Worksheets("TestAnalysis"), trad)
    Set TableFormula = CrossTableFormula.Create(Table, lData.FormDataObject)
    
    Table.AddHeader
    Table.AddRows
    Table.AddColumns
    'Those three should be done is this order
    Table.NameRanges
    TableFormula.AddFormulas
    Table.Format lData.DesignFormat

    'Set grSpecs = GraphSpecs.Create(Table)
    'grSpecs.CreateSeries
    'Set gr = Graphs.Create(sh, sh.Range("Q90"))
    'gr.Add
    'nbSeries = grSpecs.NumberOfSeries
    
    'For counter = 1 To nbSeries
    '    gr.AddSeries grSpecs.SeriesName(counter), grSpecs.SeriesType(counter), grSpecs.SeriesPos(counter)
    'Next
    
    'gr.Format
    
    
End Sub

Sub TestGraphs()
    Dim Wkb As Workbook
    Dim ana As ILLAnalysis
    Dim lData As ILinelistSpecs
    Dim sh As Worksheet
    Dim test As BetterArray
    
    Set Wkb = ThisWorkbook
    Set lData = LinelistSpecs.Create(Wkb)
    lData.Prepare
    
    Set ana = LLAnalysis.Create(Wkb.Worksheets("Analysis"), lData)
    
    Set sh = ThisWorkbook.Worksheets("TestAnalysis")
    lData.Prepare
    ana.Build sh
 
End Sub

Sub testgr()
    Dim co As ChartObject
    Dim rng As Range
    Set rng = SheetTest.Range("N12")
    Dim cw As Long
    Dim rh As Long
    
    rh = SheetTest.Range("A1").Height
    cw = SheetTest.Range("A1").Width
    Set co = SheetTest.ChartObjects.Add(rng.Left, rng.Top, cw * 8, rh * 20)
    'Debug.Print co.Top
    With co.Chart.PlotArea
        .Interior.color = RGB(235, 235, 235)
    End With
    
    
End Sub


Sub testformula()

  Dim Wksh As Worksheet
  Dim formData As IFormulaData
  Dim setupform As String
  Dim dict As ILLdictionary
  Dim formObject As Formulas
    
    Set Wksh = ThisWorkbook.Worksheets("Dictionary")
    Set dict = LLdictionary.Create(Wksh, 1, 1)
    Set Wksh = ThisWorkbook.Worksheets("ControleFormule")
    Set formData = FormulaData.Create(Wksh, "T_XlsFonctions", "T_ascii")
    setupform = "COUNTIF(outcome," & Chr(34) & "Decede" & Chr(34) & ") - COUNTIF(outcome," & Chr(34) & "Gueri" & Chr(34) & ")"
    
    Set formObject = Formulas.Create(dict, formData, setupform)
    
    Debug.Print formObject.valid
End Sub


Sub testcasewhen()
    Dim Wkb As Workbook
    Dim ana As ILLAnalysis
    Dim lData As ILinelistSpecs
    Dim sh As Worksheet
    Dim test As BetterArray
    
    Set Wkb = ThisWorkbook
    Set lData = LinelistSpecs.Create(Wkb)
    lData.Prepare
    
    Set test = lData.Categories("age_group")
End Sub

Sub TestHListVars()
    
    Dim Wkb As Workbook
    Dim ana As ILLAnalysis
    Dim lData As ILinelistSpecs
    Dim sh As Worksheet
    Dim test As BetterArray
    Dim dict As ILLdictionary
    Dim preserved As BetterArray
    Dim outsh As Worksheet
    Dim hvar As IHListVars
    
    Set preserved = New BetterArray
    
    Set Wkb = ThisWorkbook
    Set sh = Wkb.Worksheets("Dictionary")
    Set outsh = Wkb.Worksheets("Test HList")
    outsh.Cells.Clear

    Set dict = LLdictionary.Create(sh, 1, 1)
    Set lData = LinelistSpecs.Create(Wkb)
    
    lData.Prepare
    Set hvar = HListVars.Create("varb19", outsh, lData)
    hvar.WriteInfo
    
End Sub

