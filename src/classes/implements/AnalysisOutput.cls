VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Build And format the analysis"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Build And format the analysis")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements IAnalysisOutput

'Name of differents
Private Const UPDATEFILTERCOMMAND As String = "UpdateFilterTables"
Private Const SHPFILTERNAME As String = "SHP_Filter"
Private Const GOTOSECCODE As String = "go_to_section" 'Goto section code
Private Const UPDATESPCOMMAND As String = "UpdateSpTables" 'Spatial tables update

'Names of listobjects in the analysis worksheet
Private Const GLOBALSUMMARY As String = "Tab_global_summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"
Private Const SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"


Private Type TAnaOutput
  anaLL As ILinelist
  sh As Worksheet
  sectionsList As BetterArray
  headersList As BetterArray
  graphsList As BetterArray
  goToSection As String
  goToHeader As String
  goToGraph As String
End Type

Private this As TAnaOutput

'specSh is the worksheet where the specifications of the analyses are
'll is the linelist Object
Public Function Create(ByVal specSh As Worksheet, ByVal ll As ILinelist) As IAnalysisOutput
  With New AnalysisOutput
    Set .OutputLL = ll
    Set .Wksh = specSh
    Set Create = .Self()
  End With

End Function

Public Function Self() As IAnalysisOutput
  Set Self = Me
End Function

Public Property Get OutputLL() As ILinelist
  Set OutputLL = this.anaLL
End Property

Public Property Set OutputLL(ByVal ll As ILinelist)
  Set this.anaLL = ll
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal specSh As Worksheet)
  Set this.sh = specSh
End Property

'Get the linelist and linelistData attributes used when building the linelist

'Linelist Specs
Private Property Get LinelistData() As ILinelistSpecs
  Dim ll As ILinelist
  Set ll = OutputLL()
  Set LinelistData = ll.LinelistData()
End Property

'Design
Private Property Get DesignFormat() As ILLFormat
  Dim lData As ILinelistSpecs
  Set lData = LinelistData()
  Set DesignFormat = lData.DesignFormat()
End Property

'Translation of shapes and messages
Public Property Get TransObject(Optional ByVal scope As _
                                TradScopes = TranslationOfMessages) As ITranslation
  Dim lData As ILinelistSpecs
  Set lData = LinelistData()
  Set TransObject = lData.TransObject(scope)
End Property

'Output Worksheets (where to write the analysis)
Private Property Get OutputSh(Optional ByVal scope _
                             As AnalysisScope = AnalysisScopeNormal) As Worksheet
  Dim trads As ITranslation
  Dim ll As ILinelist

  Set ll = OutputLL()
  Set trads = TransObject(TranslationOfMessages)

  Select Case scope
  Case AnalysisScopeTimeSeries
    Set OutputSh = ll.Wksh(trads.TranslatedValue("LLSHEET_TemporalAnalysis"))
  Case AnalysisScopeSpatial
    Set OutputSh = ll.Wksh(trads.TranslatedValue("LLSHEET_SpatialAnalysis"))
  Case AnalysisScopeNormal
    Set OutputSh = ll.Wksh(trads.TranslatedValue("LLSHEET_Analysis"))
  Case AnalysisScopeSpatioTemporal
    Set OutputSh = ll.Wksh(trads.TranslatedValue("LLSHEET_SpatioTemporalAnalysis"))
  End Select

End Property

'Speed Up the application
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

Private Sub TransferCodeWksh(Optional ByVal scope As AnalysisScope = AnalysisScopeNormal)

    Const CHANGEMODULENAME As String = "LinelistAnalysisChange"

    Dim sh As Worksheet                          'Actual worksheet
    Dim inpWkb As Workbook                       'Input workbook
    Dim lData As ILinelistSpecs

    Dim codeContent As String                    'a string to contain code to add
    Dim vbProj As Object                         'component, project and modules
    Dim vbComp As Object
    Dim codeMod As Object
    Dim sheetName As String

    Set lData = LinelistData()
    Set inpWkb = lData.specswkb()
    Set sh = OutputSh(scope)
    sheetName = sh.Name


    'save the code module in the string sNouvCode
    With inpWkb.VBProject.VBComponents(CHANGEMODULENAME).CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With

    With sh.Parent
        Set vbProj = .VBProject
        Set vbComp = vbProj.VBComponents(.sheets(sheetName).codeName)
        Set codeMod = vbComp.CodeModule
    End With

    'Adding the code
    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub

'Split the output sheet
Private Sub SplitOutputSheet(Optional ByVal scope As AnalysisScope = AnalysisScopeNormal)
  Dim outsh As Worksheet
  Dim rowSplit As Integer

  'OUtput worksheet
  Set outsh = OutputSh(scope)
  rowSplit = IIf(scope = AnalysisScopeTimeSeries, 3, 2)
  On Error Resume Next
  'Activating the worksheet restores all the application
  'status to true, remember to recall the busy app just
  'after and before
  BusyApp
  outsh.Activate
  BusyApp
  With Application.ActiveWindow
    .SplitRow = rowSplit
    'Split on the third column for Time series and spatio-temporal only
    If scope = AnalysisScopeTimeSeries Or _ 
       scope = AnalysisScopeSpatioTemporal Then _ 
      .SplitColumn = 3
    'For spatial or time series, reduce the zoom for the worksheeet
    If scope <> AnalysisScopeNormal Then .Zoom = 70
    .FreezePanes = True
  End With
  On Error GoTo 0
End Sub

'Add the command buttons to a worksheet
Private Sub AddCommands(Optional ByVal scope _
                       As AnalysisScope = AnalysisScopeNormal)
  Dim cellRng As Range
  Dim commandName As String
  Dim shapeLabel As String
  Dim design As ILLFormat
  Dim but As IButtons
  Dim outsh As Worksheet 'output sheet
  Dim trads As ITranslation
  Dim mainobj As IMain
  Dim lData As ILinelistSpecs

  Set lData = LinelistData()
  Set mainobj = lData.MainObject()

  If mainobj.HasTemplate() Then Exit Sub

  Set design = DesignFormat()
  Set outsh = OutputSh(scope)
  Set trads = TransObject(TranslationOfShapes)
  Set cellRng = outsh.Cells(2, 2)

  'Update shapeLabel and commandName depending on the scope of the butto
  commandName = IIf(scope = AnalysisScopeSpatial, UPDATESPCOMMAND, UPDATEFILTERCOMMAND)
  shapeLabel = trads.TranslatedValue(SHPFILTERNAME)
  Set but = Buttons.Create(outrng:=cellRng, _
                           codeName:=SHPFILTERNAME, _
                           scope:=ButtonScopeSmall)
  but.Add actionCommand:=commandName, shapeLabel:=shapeLabel
  but.Format design
End Sub

'Initialize the output worksheet (format, values)
Private Sub PrepareSheet(Optional ByVal scope As AnalysisScope _
                         = AnalysisScopeNormal)

  Dim outsh As Worksheet
  Dim cellRng As Range
  Dim sheetInfo As String
  Dim rngName As String
  Dim design As ILLFormat

  Select Case scope
  Case AnalysisScopeTimeSeries
    sheetInfo = "TS-Analysis"
    rngName = "RNG_HasTable_TS"
  Case AnalysisScopeSpatial
    sheetInfo = "SP-Analysis"
    rngName = "RNG_HasTable_SP"
  Case AnalysisScopeNormal
    sheetInfo = "Uni-Bi-Analysis"
    rngName = "RNG_HasTable_UNI"
  Case AnalysisScopeSpatioTemporal
    sheetInfo = "SPT-Analysis"
    rngName = "RNG_HasTable_SPT"
  End Select

  Set outsh = OutputSh(scope)
  Set design = DesignFormat()

  'Create a range called RGN_HasTable whith value Yes if there are tables
  Set cellRng = outsh.Cells(1, 10)
  cellRng.Name = rngName
  design.ApplyFormat cellRng, scope:=LinelistHiddenCell

  'Add the whole format of the worksheet
  design.ApplyFormat outsh, scope:=AllAnalysisSheet

  'Add the info of a worksheet
  Set cellRng = outsh.Cells(1, 3)
  cellRng.Value = sheetInfo
  design.ApplyFormat cellRng, scope:=LinelistHiddenCell

  'Add warning if the cells are filtered
  Set cellRng = outsh.Cells(2, 5)
  cellRng.formula = "= ComputedOnFiltered()"
  design.ApplyFormat cellRng, scope:=LinelistWarning

  'Add the command
  AddCommands scope:=scope

  'Split the output sheet
  SplitOutputSheet scope:=scope
End Sub

'Push section, graphs or headers to the tables
Private Sub Push(pushedValue As String, _
                 Optional ByVal scope As String = "section")
  'Initialize arrays of sections, headers and graphs

  Select Case scope
  Case "section"
    this.sectionsList.Push pushedValue
  Case "header"
    this.headersList.Push pushedValue
  Case "graph"
    this.graphsList.Push pushedValue
  End Select

End Sub

'Add a GoTo (could be section or Header, or Graphs)
Private Sub InitializeArrays()
    Dim trads As ITranslation

    Set trads = TransObject()
    this.goToSection = trads.TranslatedValue("MSG_GoToSec")
    this.goToHeader = trads.TranslatedValue("MSG_GoToHead")
    this.goToGraph = trads.TranslatedValue("MSG_GoToGraph")

    'initialize section list, graphList and headersList
    If this.sectionsList Is Nothing Then Set this.sectionsList = New BetterArray
    If this.headersList Is Nothing Then Set this.headersList = New BetterArray
    If this.graphsList Is Nothing Then Set this.graphsList = New BetterArray
End Sub

'Get the goToSection, header, graph value
Private Function GoToLabel(Optional ByVal scope As String = "section") As String

  GoToLabel = Switch(scope = "section", this.goToSection, _
                     scope = "graph", this.goToGraph, _
                     scope = "header", this.goToHeader, _
                     True, this.goToSection)
End Function

Private Sub AddGoTo(ByVal cellRng As Range, _
                    Optional ByVal goToPrefix As String = "ua_", _
                    Optional ByVal scope As String = "section", _
                    Optional ByVal goToSuffix As String = vbNullString)

  Dim drop As IDropdownLists
  Dim ll As ILinelist
  Dim dropArray As BetterArray
  Dim dropName As String
  Dim design As ILLFormat

  Set dropArray = New BetterArray
  Select Case scope
  Case "section"
    Set dropArray = this.sectionsList.Clone()
  Case "header"
    Set dropArray = this.headersList.Clone()
  Case "graph"
    Set dropArray = this.graphsList.Clone()
  End Select

  'Add the goTo Section to the range
  Set ll = OutputLL()
  Set drop = ll.Dropdown()
  Set design = DesignFormat()

  'dropdown name
  dropName = goToPrefix & "gotosection" & goToSuffix
  drop.Add dropArray, dropName
  drop.SetValidation cellRng, dropName

  'Apply the dropdown format to the cell
  design.ApplyFormat cellRng, scope:=LinelistSelectionDropdown

  'Add a Name and value to the cell Range
  cellRng.Value = GoToLabel(scope)
  cellRng.Name = goToPrefix & GOTOSECCODE & goToSuffix

  'Clear the table for another usage
  Select Case scope
  Case "section"
    this.sectionsList.Clear
  Case "header"
    this.headersList.Clear
  Case "graph"
    this.graphsList.Clear
  End Select
End Sub

'Write tables of one section: Global summary, univariate or bivariate or spatial
Private Sub WriteTable(ByVal loName As String, _
                       Optional ByVal scope As AnalysisScope = AnalysisScopeNormal)
  Dim sh As Worksheet
  Dim outsh As Worksheet
  Dim tabRng As Range
  Dim hRng As Range
  Dim graphCounter As Long 'counter for graphs
  Dim tabl As ICrossTable
  Dim tabform As ICrossTableFormula
  Dim gr As IGraphs
  Dim Lo As listObject
  Dim specs As ITablesSpecs
  Dim lData As LinelistSpecs
  Dim goToSection As String
  Dim trads As ITranslation
  Dim formData As IFormulaData
  Dim cellRng As Range
  Dim grSpecs As IGraphSpecs
  Dim tabId As String
  Dim tabType As AnalysisTablesType
  Dim design As ILLFormat
  Dim anchorRng As Range 'the last Range of the listobject
  Dim containsTimeSeries As Boolean
  Dim spatRng As Range 'Spatio-temporal plage for one table
  Dim sheetInfo As String 'Tag on sheet scope

  'Variables for Time Series Tables
  Dim tablename As String
  Dim rowVar As String
  Dim minFormula As String
  Dim maxFormula As String
  Dim prevSpecs As ITablesSpecs 'Previous specs (could be Nothing)
  Dim nextSpecs As ITablesSpecs 'Next specs (could be a non valid table)
  Dim prevSecId As String 'Previous table Id
  Dim dict As ILLdictionary
  Dim vars As ILLVariables
  Dim goToHeader As String
  Dim grHeight As Long
  Dim isNewSection As Boolean 'Test if a specification is new section
  Dim addMinMaxFormula As Boolean 'Test if there is a need to update min/max formula
  Dim addSpatioTemporalName As Boolean 'Test if we should add spatio temporal name range to update formulas
  Dim upRngName As String

  'Prefix for add Goto on time series vs spatial analysis
  Dim goToPrefix As String

  'Initialize the arrays
  InitializeArrays

  'Worksheet of the analyses
  Set sh = Wksh()
  'Output Worksheet
  Set outsh = OutputSh(scope)
  Set Lo = sh.ListObjects(loName)
  Set lData = LinelistData()
  Set trads = TransObject()
  Set formData = lData.FormulaDataObject()
  Set dict = lData.Dictionary()
  Set vars = LLVariables.Create(dict)
  Set design = DesignFormat()
  sheetInfo = outsh.Cells(1, 3).Value
  upRngName = "RNG_HasTable_" & UCase(Split(sheetInfo, "-")(0))


  goToSection = GoToLabel("section")
  goToHeader = GoToLabel("header")

  If Lo.DataBodyRange Is Nothing Then Exit Sub
  
  'This is the last range of the listObject for anchor purposes
  Set hRng = Lo.HeaderRowRange()
  Set anchorRng = hRng.Offset(Lo.ListRows.Count)
  Set tabRng = Lo.ListRows(1).Range

  Do
    
    Set specs = TablesSpecs.Create(hRng, tabRng, lData)
    Set nextSpecs = Nothing
    addMinMaxFormula = False
    addSpatioTemporalName = False

    'For non valid table move to next one
    If Not specs.ValidTable() Then GoTo continueLoop

    tabType = specs.TableType()
    isNewSection = specs.isNewSection()
    containsTimeSeries = (tabType = TypeTimeSeries Or _ 
                          tabType = TypeSpatioTemporal)
    tabId = specs.TableId()  
    

    'Update goto section list on new section
    If isNewSection Then Push goToSection & ": " & specs.Value("section"), _ 
       scope:="section"

    Set tabl = CrossTable.Create(specs, outsh, trads)
    tabl.Build
    'Add Formulas
    Set tabform = CrossTableFormula.Create(tabl, formData)
    tabform.AddFormulas
    'Format the created table
    tabl.Format design

    'Will be used to test if a sheet has a table
    If IsEmpty(outsh.Range(upRngName)) Then outsh.Range(upRngName).Value = "yes"

    'Add Graphs
    If specs.HasGraph() And (Not containsTimeSeries) Then
    
      'The graph height depends on the number of elements in Geo,
      'but could result in error if n geo is not found
      grHeight = 1
      On Error Resume Next
      grHeight = CLng(specs.Value("n geo"))
      On Error GoTo 0
      
      'cellRng is where to write the graph on the worksheet
      Set cellRng = outsh.Cells(tabl.StartRow - 2, tabl.EndColumn + 2)
      'For spatial table move three columns to keep room for deviding by 
      'population
      If tabType = TypeSpatial Then Set cellRng = cellRng.Offset(, 2)

      'Create the specs of the graph and add it
      Set grSpecs = GraphSpecs.Create(tabl)
      
      'Create series uses informations on the table to create columns of
      'series of ranges to plot.
      grSpecs.CreateSeries
      Set gr = Graphs.Create(outsh, cellRng)
      
      gr.Add
      'Add the graph
      For graphCounter = 1 To grSpecs.NumberOfSeries()
        'Add the series
        gr.AddSeries grSpecs.SeriesName(graphCounter), _
        grSpecs.SeriesType(graphCounter), grSpecs.SeriesPos(graphCounter)
       'Add the series label
        gr.AddLabels grSpecs.SeriesLabel(graphCounter), _
        grSpecs.SeriesColumnLabel(graphCounter), hardCodeLabels:=False
      Next
      
      'Format the graph
      gr.Format catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
                scope:=scope, heightFactor:=grHeight
    End If

    'For TimeSeries and Spatio-Temporal, add formula for minimum and Maximum
    If Not containsTimeSeries Then GoTo continueLoop

    'Get the previous specs for time series
    On Error Resume Next
    Set prevSpecs = specs.Previous()
    Set nextSpecs = specs.NextSpecs(anchorRng)
    On Error GoTo 0

    rowVar = specs.Value("row")
    tablename = vars.Value(colName:="table name", varName:=rowVar)
        
    'Initialize the minimum and Maximum formulas
    goToPrefix = IIF(tabType = TypeTimeSeries, "ts_", "spt_")

    'If previous section is Nothing, initialize minimum and maximum
    'formulas

    If (prevSpecs Is Nothing) Then  
      'First New section, minimum and maximum formulas
      minFormula = "MIN(" & tablename & "[" & rowVar & "]" & ")"
      maxFormula = "MAX(" & tablename & "[" & rowVar & "]" & ")"
      'add goTo Headers
      Push goToHeader & ": " & tabl.TimeSeriesHeader(), scope:="header"
      addMinMaxFormula = False
    End If

    'We are on a new section, and at the end of previous section
    'New section after the first New Section
    If isNewSection And (Not prevSpecs Is Nothing)  Then
      prevSecId = prevSpecs.TableSectionId()
      addMinMaxFormula = True

      'Add Name for range of the whole table for spatio-temporal analysis
      'This is used to update formulas
      addSpatioTemporalName = (tabType = TypeSpatioTemporal)
    End If

    If (Not (nextSpecs Is Nothing)) Then
      
      If (nextSpecs.IsNewSection() And isNewSection) Then
        'A time series table with only one section, and with next sections
        prevSecId = specs.TableSectionId()
        addMinMaxFormula = True
        addSpatioTemporalName = (tabType = TypeSpatioTemporal)
      End If
    Else
      'last section of a time series table
      prevSecId = specs.TableSectionId()
      addMinMaxFormula = True

      'Add an anchor to where to start building graphs for time series (START_GRAPHS_TS)
      Set cellRng = outsh.Range("ENDTABLE_" & prevSecId).Offset(5, 2)

      If tabType = TypeTimeSeries Then cellRng.Name = "START_GRAPHS_TS"
      If tabType = TypeSpatioTemporal Then cellRng.Name = "START_GRAPHS_SPT"
    End If

    If (Not isNewSection) And (Not (nextSpecs Is Nothing)) Then
      addMinMaxFormula = False
      'No new section, update minimum and maximum formula
      minFormula = minFormula & ", " & "MIN(" & tablename & "[" & rowVar & "]" & ")"
      maxFormula = maxFormula & ", " & "MAX(" & tablename & "[" & rowVar & "]" & ")"
      Push goToHeader & ": " & tabl.TimeSeriesHeader(), scope:="header"
    End If

    'Add minimum and maximum formulas
    If addMinMaxFormula Then

      outsh.Range("MAX_MAX_DATE_" & prevSecId).formula = "= MAX(" & maxFormula & ")"
      outsh.Range("MIN_MIN_DATE_" & prevSecId).formula = "= MIN(" & minFormula & ")"

      'On new section update min and max formula
      If isNewSection Then
        minFormula = "MIN(" & tablename & "[" & rowVar & "]" & ")"
        maxFormula = "MAX(" & tablename & "[" & rowVar & "]" & ")"
      End If

      'Add goTo Header to the previous section
      Set cellRng = outsh.Range("SECTION_" & prevSecId).Offset(1, -2)
      AddGoTo cellRng:=cellRng, goToPrefix:=goToPrefix, _ 
              scope:="header",  goToSuffix:=prevSecId
    End If

    'Add name for spatio-temporal table
    If addSpatioTemporalName Then
      With outsh
        Set cellRng = .Range("COLUMN_CATEGORIES_" & tabId)
        Set spatRng = .Range( _ 
                        .Range("COLUMN_CATEGORIES_" & prevSecId).Cells(1, 1), _ 
                        cellRng.Cells(1, cellRng.Columns.Count) _ 
                      )
      End With
      spatRng.Name = "SPT_FORMULA_COLUMN_" & prevSecId
    End If

  continueLoop:
  Set tabRng = tabRng.Offset(1)
  Loop While (tabRng.Row <= anchorRng.Row)
End Sub

'Write table for time series (oe table of type long table)
'tabLoName: listobject name of the table
'graphLoName: listobject name of the graph
'graphTitleLoName : lisobject name of graph titles
Private Sub WriteTimeSeriesGraphs(ByVal graphLoName As String, _
                                  ByVal graphTitleLoName As String, _
                                  ByVal tabLoName As String)
  Dim sh As Worksheet
  Dim outsh As Worksheet
  Dim graphCounter As Long
  Dim cellRng As Range
  Dim lData As ILinelistSpecs
  Dim seriesCounter As Long

  'ListObject for graphs, table and title
  Dim grLo As listObject
  Dim tabLo As listObject
  Dim grTitleLo As listObject
  Dim grSpecs As IGraphSpecs
  Dim gr As IGraphs
  Dim trads As ITranslation
  Dim design As ILLFormat

  'Different BetterArrays for stocking informations on graphs
  Dim grBAinput As BetterArray 'input elements for graph specs
  Dim grBAoutput As BetterArray

  'temporary tables for all the graphs
  Dim tabGraphInfos As BetterArray 'is a BetterArray of 6 elements (following)

  Dim tabSeries As BetterArray
  Dim tabTypes As BetterArray
  Dim tabPos As BetterArray
  Dim tabLabel As BetterArray
  Dim tabColLabel As BetterArray
  Dim tabPrefLabel As BetterArray
  Dim plotTitle As String

  'Worksheets
  Set sh = Wksh()
  Set outsh = OutputSh(AnalysisScopeTimeSeries)

  If IsEmpty(outsh.Range("RNG_HasTable_TS")) Then Exit Sub

  'Initialize the arrays
  Set grBAinput = New BetterArray '
  Set grBAoutput = New BetterArray
  Set tabGraphInfos = New BetterArray
  Set tabSeries = New BetterArray
  Set tabTypes = New BetterArray
  Set tabPos = New BetterArray
  Set tabLabel = New BetterArray
  Set tabColLabel = New BetterArray
  Set tabPrefLabel = New BetterArray


  'ListObjects and linelist data
  Set lData = LinelistData()
  'Where to start the graphs on time series
  Set cellRng = outsh.Range("START_GRAPHS_TS")
  Set grLo = sh.ListObjects(graphLoName)
  Set tabLo = sh.ListObjects(tabLoName)
  Set grTitleLo = sh.ListObjects(graphTitleLoName)
  Set trads = TransObject()
  Set design = DesignFormat()

  If (grLo.DataBodyRange Is Nothing) Or _
  (grTitleLo.DataBodyRange Is Nothing) Then _
  Exit Sub

  grBAinput.Push grLo, tabLo, grTitleLo

  Set grSpecs = GraphSpecs.CreateRangeSpecs(grBAinput, outsh, lData)

  'Build the graphs if they are valid
  If grSpecs.Valid() Then
    Set grBAoutput = grSpecs.SpecsLists()
    For graphCounter = 1 To grSpecs.NumberOfGraphs()
      Set tabGraphInfos = grBAoutput.Item(graphCounter)

      'temptabl is a table of five tables
      Set tabSeries = tabGraphInfos.Item(1) 'List of series for one graph
      Set tabTypes = tabGraphInfos.Item(2) 'List of types for one graph (line, bar, point)
      Set tabPos = tabGraphInfos.Item(3) 'List of the positions for one graph

      'Labels and column labels
      Set tabLabel = tabGraphInfos.Item(4)
      Set tabColLabel = tabGraphInfos.Item(5)
      Set tabPrefLabel = tabGraphInfos.Item(6)

      Set gr = Graphs.Create(outsh, cellRng)
      gr.Add

      For seriesCounter = 1 To tabSeries.Length
          gr.AddSeries rngName:=tabSeries.Item(seriesCounter), _
                       chrtType:=tabTypes.Item(seriesCounter), _
                       axisPos:=tabPos.Item(seriesCounter)

          gr.AddLabels catName:=tabLabel.Item(seriesCounter), _
                       lblName:=tabColLabel.Item(seriesCounter), _
                       prefix:=tabPrefLabel.Item(seriesCounter), prefixOnly:=True
      Next

      'Format the graphs
      plotTitle = tabGraphInfos.Item(7)
      gr.Format catTitle:=trads.TranslatedValue("MSG_Period"), _
                valuesTitle:=trads.TranslatedValue("MSG_Values"), _
                plotTitle:=plotTitle, scope:=GraphScopeTimeSeries

      'Add the graph title to graph lists
      cellRng.Value = plotTitle

      'Format the plot title cell
      design.ApplyFormat cellRng, scope:=LinelistHiddenCell
      Push GoToLabel("graph") & ": " & plotTitle, scope:="graph"
      Set cellRng = cellRng.Offset(50)
    Next
  End If

  'Add the GoTo graph dropdown
  AddGoTo outsh.Cells(3, 1), scope:="graph", _
          goToPrefix:="ts_", goToSuffix:="_graph"
End Sub

'Write spatio temporal graphs
Private Sub WriteSpatioTemporalGraph(ByVal loName As String)


  Dim sh As Worksheet 'Worksheet with informations on the spatio-temporal setups
  Dim outsh As Worksheet
  Dim cellRng As Range
  Dim specs As ITablesSpecs
  Dim lData As LinelistSpecs
  Dim design As ILLFormat
  Dim hRng As Range
  Dim tabRng As Range
  Dim anchorRng As Range 'the last Range of the listobject
  Dim gr As IGraphs
  Dim grSpecs As IGraphSpecs
  Dim plotTitle As String
  Dim tabId As String
  Dim graphCounter As Long
  Dim trads As ITranslation
  Dim Lo As ListObject
  Dim tabl As ICrossTable

  'Worksheet of the analyses
  Set sh = Wksh()
  'Output Worksheet
  Set outsh = OutputSh(AnalysisScopeSpatioTemporal)
  Set Lo = sh.ListObjects(loName)
  Set lData = LinelistData()
  Set trads = TransObject()
  Set design = DesignFormat()


  If Lo.DataBodyRange Is Nothing Then Exit Sub

  Set cellRng = outsh.Range("START_GRAPHS_SPT")
  Set hRng = Lo.HeaderRowRange()
  Set anchorRng = hRng.Offset(Lo.ListRows.Count)
  Set tabRng = Lo.ListRows(1).Range

  Do
    
    Set specs = TablesSpecs.Create(hRng, tabRng, lData)

    If Not specs.ValidTable() Then GoTo continueLoop
    If Not specs.HasGraph() Then GoTo continueLoop

    tabId = specs.TableId()
    Set tabl = CrossTable.Create(specs, outsh, trads)
    
    'Create the specs of the graph and add it
    Set grSpecs = GraphSpecs.Create(tabl)
    plotTitle = trads.TranslatedValue("MSG_Graph") & _ 
                outsh.Range("STARTCOL_" & tabId).Value
      
    'Create series uses informations on the table to create columns of
    'series of ranges to plot.
    grSpecs.CreateSeries
    Set gr = Graphs.Create(outsh, cellRng)
    gr.Add
    
    Debug.Print grSpecs.NumberOfSeries()

    'Add the graph
    For graphCounter = 1 To grSpecs.NumberOfSeries()
      'Add the series
      gr.AddSeries grSpecs.SeriesName(graphCounter), _
      grSpecs.SeriesType(graphCounter), grSpecs.SeriesPos(graphCounter)
     'Add the series label
      gr.AddLabels grSpecs.SeriesLabel(graphCounter), _
      grSpecs.SeriesColumnLabel(graphCounter), hardCodeLabels:=False
    Next
      
    'Format the graph
    gr.Format catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
              valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
              scope:=GraphScopeTimeSeries, plotTitle:=plotTitle

    cellRng.Value = plotTitle
    design.ApplyFormat cellRng, scope:=LinelistHiddenCell
    Push GoToLabel("graph") & ": " & plotTitle, scope:="graph"
    Set cellRng = cellRng.Offset(50)

  continueLoop:
  Set tabRng = tabRng.Offset(1)
  Loop While (tabRng.Row <= anchorRng.Row)

  'Add the GoTo graph dropdown
  AddGoTo outsh.Cells(3, 1), scope:="graph", goToPrefix:="spt_", _ 
          goToSuffix:="_graph"
End Sub

'Write the table
Private Sub WriteAnalysis()
  Dim outsh As Worksheet
  Dim cellRng As Range
  Dim scope As AnalysisScope
  Dim pass As ILLPasswords
  Dim lData As ILinelistSpecs

  Set lData = LinelistData()
  Set pass = lData.Password()

  'Univariate and Bivariate Analysis Worksheet =================================

  scope = AnalysisScopeNormal
  Set outsh = OutputSh(scope)
  ' --- Prepare the worksheet

  PrepareSheet scope:=scope

  'Add an Error Management for the tables
   On Error GoTo ErrHand
  ' --- Write Global summary
   WriteTable loName:=GLOBALSUMMARY, scope:=scope
   Push GoToLabel("section") & ""
   ' ---- Write Univariate Analysis
   WriteTable loName:=UNIVARIATE, scope:=scope
   ' ---- Write the Bivariate Analysis
   WriteTable loName:=BIVARIATE, scope:=scope

   On Error GoTo 0

   'Add The goTo for sections
   Set cellRng = outsh.Cells(2, 1)
   AddGoTo cellRng:=cellRng, goToPrefix:="ua_", scope:="section"
   'Add the command
   AddCommands scope:=scope

   'Hide the output sheet if required
   pass.Protect outsh, True
   If IsEmpty(outsh.Range("RNG_HasTable_UNI")) Then _ 
    outsh.Visible = xlSheetHidden

   'Transfer code to the worksheet
   TransferCodeWksh scope:=scope

   'Time Series Analysis Worksheet =============================================
  scope = AnalysisScopeTimeSeries
  Set outsh = OutputSh(scope)
  PrepareSheet scope:=scope
  'Write the tables of time series
  On Error GoTo ErrHand

  WriteTable loName:=TIMESERIES, scope:=scope
  'Write the graphs
  WriteTimeSeriesGraphs graphLoName:=GRAPHTIMESERIES, _
                        graphTitleLoName:=TITLEGRAPHTIMESERIES, _
                        tabLoName:=TIMESERIES
  On Error GoTo 0

  Set cellRng = outsh.Cells(2, 1)
  AddGoTo cellRng:=cellRng, scope:="section", goToPrefix:="ts_"
  'Add the commands
  AddCommands scope:=scope

  pass.Protect outsh, True
  If IsEmpty(outsh.Range("RNG_HasTable_TS")) Then _ 
   outsh.Visible = xlSheetHidden

  'Transfert code to worksheet
   TransferCodeWksh scope:=scope

  'Spatial Analysis Worksheet ==================================================
  scope = AnalysisScopeSpatial
  Set outsh = OutputSh(scope)

  PrepareSheet scope:=scope

  On Error GoTo ErrHand
    WriteTable loName:=SPATIAL, scope:=scope
  On Error GoTo 0

  Set cellRng = outsh.Cells(2, 1)

  AddGoTo cellRng:=cellRng, scope:="section", goToPrefix:="sp_"
  AddCommands scope:=scope

  pass.Protect outsh, True
  If IsEmpty(outsh.Range("RNG_HasTable_SP")) Then _ 
   outsh.Visible = xlSheetHidden

  'Transfer code to spatial analysis sheet
  TransferCodeWksh scope:=scope

  'Spatio-Temporal Analysis worksheet ==========================================
  scope = AnalysisScopeSpatioTemporal
  Set outsh = OutputSh(scope)
  PrepareSheet scope:=scope

  On Error GoTo ErrHand
    WriteTable loName:=SPATIOTEMPORAL, scope:=scope
    WriteSpatioTemporalGraph loName:=SPATIOTEMPORAL
  On Error GoTo 0

  Set cellRng = outsh.Cells(2, 1)
  AddGoTo cellRng:=cellRng, scope:="section", goToPrefix:="spt_"
  AddCommands scope:=scope
  
  'pass.Protect outsh, True
  If IsEmpty(outsh.Range("RNG_HasTable_SPT")) Then _ 
   outsh.Visible = xlSheetHidden

  'Transfer code to spatio-temporal analysis sheet
  TransferCodeWksh scope:=scope

  Exit Sub

ErrHand:
    Debug.Print "Error: " & Err.Number & " " & Err.Description
    Resume Next
End Sub

Private Sub IAnalysisOutput_WriteAnalysis()
  'write analysis
  WriteAnalysis
End Sub
