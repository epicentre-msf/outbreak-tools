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

    With co.Chart.PlotArea
        .Interior.color = RGB(235, 235, 235)
    End With
    
    
End Sub

Sub testRange()
    Dim sh As Worksheet
    Dim rng As Range
    Dim testRng As Range
    
    Set sh = SheetTest
    
    With sh
        Set rng = .Range(.Cells(1, 1), .Cells(2, 1))
        Set testRng = .Range(rng.Cells(2, 1), rng.Cells(2, 1))
    End With
    
    Debug.Print testRng.Address
    
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
    setupform = "COUNTIF(outcome," & chr(34) & "Decede" & chr(34) & ") - COUNTIF(outcome," & chr(34) & "Gueri" & chr(34) & ")"
    
    Set formObject = Formulas.Create(dict, formData, setupform)
    
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
  
    io.LoadFolder

End Sub

Sub TestGeo()
    Dim geoObject As ILLGeo
    Dim admname As String
    Dim wb As Workbook
    Dim admList As BetterArray
    Dim pcodeValue As String
    Dim rng As Range
    Dim val As String
    
    Dim admNames As BetterArray
    
    Set geoObject = LLGeo.Create(SheetGeo)
    Set rng = SheetGeo.Range("N15")
    
    val = GEOPCODE(rng, 3)
    
    
    
End Sub

Sub TestSections()

    Dim ll As ILinelist
    Dim lData As ILinelistSpecs
    Dim buildingSheet As Object
    Dim currSheetName As String
    Dim dict As ILLdictionary
    Dim llshs As ILLSheets
    Dim llana As ILLAnalysis
    Dim mainobj As IMain
    Dim rng As Range

    Set lData = LinelistSpecs.Create(ThisWorkbook)
    lData.Prepare
    Set ll = Linelist.Create(lData)
    ll.Prepare
    Set dict = lData.Dictionary()
    Set llshs = LLSheets.Create(dict)
    Set mainobj = lData.MainObject()
    
    Set rng = dict.DataRange("sheet name")

    mainobj.UpdateStatus (10)

    currSheetName = dict.DataRange("sheet name").Cells(1, 1).Value
    
    If llshs.SheetInfo(currSheetName) = "vlist1D" Then

        Set buildingSheet = Vlist.Create(currSheetName, ll)
    
    ElseIf llshs.SheetInfo(currSheetName) = "hlist2D" Then

        Set buildingSheet = Hlist.Create(currSheetName, ll)
    
    End If

    If buildingSheet Is Nothing Then Exit Sub
    
    mainobj.UpdateStatus (15)
     
    'Build the first sheet
    buildingSheet.Build

    'Loop through the other sheets and build them also
    Do While (buildingSheet.NextSheet() <> vbNullString)
        
        currSheetName = buildingSheet.NextSheet()

        If llshs.SheetInfo(currSheetName) = "vlist1D" Then
            Set buildingSheet = Vlist.Create(currSheetName, ll)
        ElseIf llshs.SheetInfo(currSheetName) = "hlist2D" Then
            Set buildingSheet = Hlist.Create(currSheetName, ll)
        End If
        
        'If you still remain on the same sheet exit (something weird happened)
        If currSheetName = buildingSheet.NextSheet() Then Exit Do
        buildingSheet.Build
    Loop

    Set llana = lData.Analysis()

End Sub

