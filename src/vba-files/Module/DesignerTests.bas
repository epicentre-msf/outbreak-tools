Attribute VB_Name = "DesignerTests"
Option Explicit
Option Private Module

Sub ShowWindows()
    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    Application.EnableEvents = True
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
    Set TableFormula = CrossTableFormula.Create(Table, lData.FormulaDataObject)
    
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
    Dim shUA As Worksheet
    Dim shTS As Worksheet
    Dim test As BetterArray
    
    Set Wkb = ThisWorkbook
    Set lData = LinelistSpecs.Create(Wkb)
    lData.Prepare
    
    Set ana = LLAnalysis.Create(Wkb.Worksheets("Analysis"), lData)
    
    Set shUA = ThisWorkbook.Worksheets("TestAnalysisUA")
    Set shTS = ThisWorkbook.Worksheets("TestAnalysisTS")
    lData.Prepare
    ana.Build shUA, shTS
 
End Sub

Sub testgr()
    Dim co As ChartObject
    Dim rng As Range
    Set rng = SheetTest.Range("N12")
    Dim cw As Long
    Dim rh As Long
    
    rh = SheetTest.Range("A1").height
    cw = SheetTest.Range("A1").width
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
    Set formData = FormulaData.Create(Wksh)
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
    Dim outdropsh As Worksheet
    Dim hvar As IHListVars
    Dim listTest As BetterArray
    Dim drop As IDropdownLists
    
    
    Set preserved = New BetterArray
    
    Set Wkb = ThisWorkbook
    Set outdropsh = Wkb.Worksheets("Test Dropdown")
    Set outsh = Wkb.Worksheets("Test HList")
    Set lData = LinelistSpecs.Create(Wkb)
    
    lData.Prepare
    
    Set drop = DropdownLists.Create(outdropsh)
    outsh.Cells.Clear
    outdropsh.Cells.Clear
    Set hvar = HListVars.Create("varb14", outsh, lData, drop)

    hvar.WriteInfo
    
End Sub

Sub TestOs()
    Dim io As IOSFiles

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    
    Debug.Print io.HasValidFolder()
    Debug.Print io.HasValidFile()
    Debug.Print io.File
    io.LoadFolder
    Debug.Print io.Folder
End Sub

Sub TestGeo()
    Dim geoObject As ILLGeo
    Dim admname As String
    Dim wb As Workbook
    Dim admList As String
    Dim admNames As BetterArray
    
    
    Set geoObject = LLGeo.Create(SheetGeo)
    'admname = geoObject.GeoNames("adm2_name")
    'geoObject.Translate rawNames:=True
    'geoObject.Clear
    'geoObject.ClearHistoric
    'Set wb = Workbooks.Open("D:\Projects\outbreak-tools\input\geobase\OUTBREAK-TOOLS-GEOBASE-SSD-2022-09-07.xlsx")
    'geoObject.Import wb
    'Get the list of admin 1
    Set admNames = New BetterArray
    admNames.LowerBound = 1
    admNames.Push "Abyei", "Abyei Region", "Alel"
    'admNames = "Unity"
     admList = geoObject.Population(LevelAdmin3, admNames)
    
    
End Sub


Sub TestSections()

    Dim ll As ILinelist
    Dim lData As ILinelistSpecs
    Dim horizSheet As IHlist
    
    Set lData = LinelistSpecs.Create(ThisWorkbook)
    lData.Prepare
    Set ll = Linelist.Create(lData)
    ll.Prepare
    Set horizSheet = Hlist.Create("B-H2D", ll)
    horizSheet.Build

End Sub

