VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "AnalysisOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Module to put all together for the analysis"

'@Folder("Analysis")
'@ModuleDescription("Build And format the analysis")
'@IgnoreModule

Option Explicit
Implements IAnalysisOutput

'Name of differents 
Private Const UPDATEFILTERCOMMAND As String = "UpdateFilterTables"
Private Const SHPFILTERNAME As String = "SHP_Filter"
Private Const GOTOSECCODE As String = "go_to_section" 'Goto section code
Private Const UPDATESPCOMMAND As String = "UpdateSpTables" 'Spatial tables update
Private Const CLASSNAME As String = "AnalysisOutput"

'Names of listobjects in the analysis worksheet
Private Const GLOBALSUMMARY As String = "Tab_global_summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"


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
Public Function Create(ByVal specSh As Worksheet, ByVal ll As ILinelist) As IAnalysisOutpout
  With New AnalysisOutput
    Set .OutputLL = ll
    Set .Wksh = specSh
    Set Create = .Self()
  End With

End Function

Public Function Self() As IAnalysisOutpout
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

'Translation of messages / other translations
Private Property Get Trads(Optional ByVal scope As TradScopes = TranslationOfMessages) As ITranslation
  Dim ll As ILinelist
  Dim lData As LinelistSpecs
  Set ll = OutputLL()
  Set lData = ll.LinelistData()
  Set Trads = lData.TransObject(scope)
End Property

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
Public Property Get TransObject(ByVal scope As  _ 
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
    Set OutputSh = trads.TranslatedValue("LLSHEET_TemporalAnalysis")
  Case AnalysisScopeSpatial
    Set OutputSh = trads.TranslatedValue("LLSHEET_SpatialAnalysis")
  Case AnalysisScopeNormal
    Set OutputSh = trads.TranslatedValue("LLSHEET_Analysis")
  Case Else
    Set OutputSh = trads.TranslatedValue("LLSHEET_Analysis")
  End Select
End Property

'Speed Up the application
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'Return previous state
Private Sub ReturnState()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableAnimations = True
End Sub

'Split the output sheet

Private Sub SplitOutputSheet(Optional ByVal scope As AnalysisScope = AnalysisScopeNormal)
  Dim outsh As Worksheet
  Dim rowSplit As Long
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
    'Split on the third column for Time series only
    If scope = AnalysisScopeTimeSeries Then .SplitColumn = 3
    'For spatial or time series, reduce the zoom for the worksheeet
    If scope <> AnalysisScopeNormal Then .Zoom = 65
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

  Set design = DesignFormat()
  Set outsh = OutputSh(scope)
  Set trads = TransObject(TranslationOfShapes)
  Set cellRng = outsh.Cells(2, 2)

  'Update shapeLabel and commandName depending on the scope of the butto
  commandName = IIf(scope = AnalysisScopeSpatial, UPDATESPCOMMAND, UPDATEFILTERCOMMAND)
  shapeLabel = trads.TranslatedValue(SHPFILTERNAME)
  Set but = Buttons.Create(outRng:=cellRng, codeName:=SHPFILTERNAME, _ 
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
  Dim design As ILLFormat

  Select Case scope
  Case AnalysisScopeTimeSeries
    sheetInfo = "TS-Analysis"
  Case AnalysisScopeSpatial
    sheetInfo = "SP-Analysis"
  Case AnalysisScopeNormal
    sheetInfo = "Uni-Bi-Analysis"
  End Select

  Set outsh = OutputSh(scope)
  Set design = DesignFormat()
  Set cellRng = outsh.Cells(1, 3)

  'Add the whole format of the worksheet
  design.ApplyFormat outsh, scope:=AllAnalysisSheet

  'Add the info of a worksheet
  cellRng.Value = sheetInfo
  design.ApplyFormat cellRng, scope:=LinelistHiddenCell

  'Add warning if the cells are filtered
  Set cellRng = outshCells(2, 5)
  cellRng.Formula = "= ComputedOnFiltered()"
  design.ApplyFormat cellRng, scope:=LinelistWarning


  'Add the command
  AddCommands scope:=scope

  'Split the output sheet
  SplitOutputSheet scope:=scope
End Sub


'Push section, graphs or headers to the tables
Private Sub Push(pushedValue As String, Optional ByVal scope As String = "section")
  'Initialize arrays of sections, headers and graphs
  Select case scope
  Case "section"
    this.sectiionsList.Push pushedValue
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
Private Function GoToLabel(optional ByVal scope As String = "section") As String

  GoToLabel = Switch(scope = "section", this.goToSection, _ 
                     scope = "graph", this.goToGraph, _ 
                     scope = "header", this.goToHeader, _ 
                     True, this.goToSection)
End Function

Private Sub AddGoTo(cellRng As Range, _ 
                    Optional ByVal goToPrefix As String = "ua_", _ 
                    Optional ByVal scope As String = "section", _ 
                    Optional ByVal goToSuffix As String = vbNullString)
  
  Dim drop As IDropdownLists
  Dim ll As ILinelist
  Dim dropArray As BetterArray
  Dim trads As ITranslation
  Dim dropName As String
  Dim design As ILLFormat

  Set dropArray = New BetterArray
  Select Case scope
  Case "section"
    Set dropArray = this.sectionsList.Clone()
  Case "graph"
    Set dropArray = this.headersList.Clone()
  Case "header"
    Set dropArray = this.graphsList.Clone()
  End Select

  'Add the goTo Section to the range
  Set ll = OutputLL()
  Set drop = ll.Dropdown()
  Set design = DesignFormat()
  Set trads = TransObject()

  'dropdown name
  dropName = goToPrefix & "gotosection"
  drop.Add dropArray, dropName
  drop.SetValidation cellrng, dropName

  'Apply the dropdown format to the cell
  design.ApplyFormat cellRng, scope:=LinelistSelectionDropdown
  
  'Add a Name and value to the cell Range
  cellRng.Value = GoToLabel(scope)
  cellRng.Name = goToPrefix & GOTOSECCODE & goToSuffix

  'Clear the table for another usage
  Select Case scope
  Case "section"
    this.sectionsList.Clear()
  Case "graph"
    this.headersList.Clear()
  Case "header"
    this.graphsList.Clear()
  End Select
End Sub

'Write tables of one section: Global summary, univariate or bivariate or spatial
Private Sub WriteTable(ByVal LoName As String,  _ 
                       Optional ByVal scope As AnalysisScope = AnalysisScopeNormal)
  Dim sh As Worksheet
  Dim outsh As Worksheet
  Dim tabRng As Range
  Dim hRng As Range
  Dim tabCounter As Long 'counter for tables
  Dim graphCounter As Long 'counter for graphs
  Dim tabl As ICrossTable
  Dim tabform As ICrossTableFormula
  Dim gr As IGraphs
  Dim Lo As ListObjects
  Dim specs As ITablesSpecs
  Dim lData As LinelistSpecs
  Dim goToSection As String
  Dim trads As ITranslation
  Dim formData As IFormulaData
  Dim cellRng As Range
  Dim grSpecs As IGraphSpecs
  Dim tabId As String
  Dim tabType As AnalysisTablesType
  Dim addGraph As String
  
  'Variables for Time Series Tables
  Dim tableName As String
  Dim rowVar As String
  Dim minFormula As String
  Dim maxFormula As String
  Dim lastValid As ITablesSpecs
  Dim prevSpecs As ITableSpecs 'Previous specs (could be Nothing)
  Dim prevSecId As String 'Previous table Id
  Dim dict As ILLDictionary
  Dim vars As ILLVariables
  Dim goToHeader As STring

  'Initialize the arrays
  InitializeArrays

  'Worksheet of the analyses
  Set sh = Wksh()
  'Output Worksheet
  Set outsh = OutputSh(scope)
  Set Lo = sh.ListObjects(LoName)
  Set lData = LinelistData()
  Set trads = TransObject()
  Set formData = lData.FormulaDataObject()
  Set dict = lData.Dictionary()
  Set vars = LLVariables.Create(dict)

  
  goToSection = GoToLabel("section")
  goToHeader = GoToLabel("header")

  tabType = specs.TableType()

  If Lo.DataBodyRange Is Nothing Then Exit SuB
  Set hRng = Lo.HeaderRowRange

  For tabCounter = 1 To Lo.ListRows.Count
    Set tabRng = Lo.ListRows(counter).Range
    Set specs = TablesSpecs.Create(hRng, tabRng, lData)
    
    If specs.ValidTable() Then

      If specs.isNewSection Then 
        Push goToSection & ": " & specs.Value("section")
        'The last valid specs is the specs of the last valid new section for TS
        Set lastValid = specs
      End If
      
      'Get the previous specs 
      On Error Resume Next
      Set prevSpecs = specs.Previous()
      On Error GoTo 0

      Set tabl = CrossTable.Create(specs, outsh, trads)
      tabl.Build
      'Add Formulas
      Set tabForm = CrossTableFormula.Create(tabl, formData)
      tabForm.AddFormulas
      'Format the created table
      tabl.Format design

      'Minimum and maximum formula for Time series; GoTo Headers for time series
      If (tabType = TypeTimeSeries) Then

        rowVar = specs.Value("row")
        tableName = vars.Value(colName:="table name", varName:=rowVar)
        'Initialize the minimum and Maximum formulas
        If (prevSpecs Is Nothing) Then
        'First New section, minimum and maximum formulas
          minFormula = "MIN(" & tableName & "[" & rowVar & "]" & ")"
          maxFormula = "MAX(" & tableName & "[" & rowVar & "]" & ")"

          'add goTo Headers
          Push goToHeader & ": " & tabl.TimeSeriesHeader(), scope:="header"
        
        'It is Possible to have new section and previous section is Nothing
        ElseIf specs.isNewSection() And (Not prevSpecs Is Nothing) Then

          'New section after the first New Section
          prevSecId = prevSpecs.TableSectionId()
          
          outsh.Range("MAX_MAX_DATE_" & prevSecId).formula = "= MAX(" & maxFormula & ")"
          outsh.Range("MIN_MIN_DATE_" & prevSecId).formula = "= MIN(" & minFormula & ")"

          minFormula = "MIN(" & tableName & "[" & rowVar & "]" & ")"
          maxFormula = "MAX(" & tableName & "[" & rowVar & "]" & ")"

          'Add goTo Header to the previous section
          Set cellRng = outsh.Range("SECTION_" & prevSecId).Offset(1, -2)
          AddGoTo cellRng:=cellRng, goToPrefix:="ts_", scope:="header", _ 
                   goToSuffix:=prevSecId
        
        ElseIf (Not specs.isNewSection()) Then  
          'No new section, update minimum and maximum formula
          minFormula = minFormula & ", " & "MIN(" & tableName & "[" & rowVar & "]" & ")"
          maxFormula = maxFormula & ", " & "MAX(" & tableName & "[" & rowVar & "]" & ")"

          Push goToHeader & ": " & tabl.TimeSeriesHeader(), scope:="header"
        End If
      End If
      
      'Add graphs if needed for bivariate, univariate and spatial
      addGraph = specs.Value("graph")

      If (addGraph = "yes" And tabType = TypeUnivariate) Or _ 
        ((addGraph = "values" Or addGraph = "percentages") _ 
         And ((tabType = TypeBivariate) Or (tabType = TypeSpatial))) Then

        Set cellRng = outsh.Cells(tabl.StartRow - 2, tabl.EndColumn + 2)
        tabId = specs.TableId()

        'For spatial table move three columns to keep room for deviding by population
        If tabType = TypeSpatial Then Set cellRng = cellRng.Offset(, 3)
        
        'Create the specs of the graph and add it
        Set grSpecs = GraphsSpecs.Create(tabl)
        grSpecs.CreateSeries
        Set gr = Graphs.Create(outsh, cellRng)
        gr.Add
        
        'Add 
        For graphCounter = 1 To grSpecs.NumberOfSeries()
          'Add the series
          gr.AddSeries grSpecs.SeriesName(graphCounter), _ 
          grSpecs.SeriesType(graphCounter), grSpecs.SeriesPos(graphCounter)
          
          'Add the series label
          gr.AddLabels grSpecs.SeriesLabel(graphCounter), _ 
          grSpecs.SeriesColumnLabel(graphCounter), hardCodeLabels:=False
        Next
        
        'format the graph
        gr.Format catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _ 
        valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
        scope:=scope
      End If
    End If
  Next

  'Add minimum and maximum formula for last section if on time series
  If (Not lastValid is Nothing) And (tabType = TypeTimeSeries) Then
    tabId = lastValid.TableSectionId
    outsh.Range("MAX_MAX_DATE_" & tabId).formula = "= MAX(" & maxFormula & ")"
    outsh.Range("MIN_MIN_DATE_" & tabId).formula = "= MIN(" & minFormula & ")"
    Set cellRng = outsh.Range("SECTION_" & tabId).Offset(1, -2)
    'Add dropdown
    AddGoTo cellRng:=cellRng, goToPrefix:="ts_", scope:="header", _ 
                   goToSuffix:=tabId
    'Add an anchor to where to start building graphs for time series
    Set cellRng = outsh.Range("ENDTABLE_" & lastValid.TableId).Offset(5, 2)
    cellRng.Name = "START_GRAPHS_TS"
  End If

End Sub

'Write table for time series (oe table of type long table)
'tabLoName: listobject name of the table
'graphLoName: listobject name of the graph
'graphTitleLoName : lisobject name of graph titles
Private Sub WirteTimeSeriesGraphs(ByVal graphLoName As String, graphTitleLoName As String, _ 
                                  ByVal tabLoName As String)
  Dim sh As Worksheet
  Dim outsh As Worksheet
  Dim graphCounter As Long
  Dim cellRng As Range
  Dim lData As ILinelistSpecs
  Dim seriesCcounter As Long

  'ListObject for graphs, table and title
  Dim grLo As ListObject
  Dim tabLo As ListObject
  Dim grTitleLo As LisObject
  Dim grSpecs As IGraphSpecs
  Dim grTitle As String
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
  
  
  'ListObjects
  Set sh = Wksh()
  Set lData = LinelistData()
  Set outsh = OutputSh(AnalysisScopeTimeSeries)
  Set cellRng = outsh.Range("START_GRAPHS_TS")
  Set grLo = sh.ListObjects(graphLoName)
  Set tabLo = sh.ListObjects(tabLoName)
  Set grTitleLo = sh.ListObjects(graphTitleLoName)
  Set trads = TransObject()
  Set design = DesignFormat()

  grBAinput.Push grLo, tabLo, grTitleLo

  Set grSpecs = GraphSpecs.CreateRangeSpecs(grBAinput, outsh, lData)

  'Build the graphs if they are valid
  If grSpecs.Valid() Then
    Set grBAOutput = grSpecs.SpecsLists()

    For graphCounter = 1 To grSpecs.NumberOfGraphs()
      Set tabGraphInfos = grBAoutput.Item(graphCounter)

      'temptabl is a table of five tables
      Set tabSeries = tabGraphInfos.Item(1) 'List of series for one graph
      Set tabType = tabGraphInfos.Item(2) 'List of types for one graph (line, bar, point)
      Set tabPos = tabGraphInfos.Item(3) 'List of the positions for one graph

      'Labels and column labels
      Set tabLabel = tabGraphInfos.Item(4)
      Set tabColLabel = tabGraphInfos.Item(5)
      Set tabPrefLabel = tabGraphInfos.Item(6)

      Set gr = Graphs.Create(outsh, cellRng)
      gr.Add

      For seriesCounter = 1 To tabSeries.Length
          gr.AddSeries tabSeries.Item(counter), tabType.Item(counter), _ 
                       tabPos.Item(counter)
          gr.AddLabels tabLabel.Item(counter), tabColLabel.Item(counter), _ 
                       tabPrefLabel.Item(counter), prefixOnly:=True
      Next

      'Format the graphs
      plotTitle = tabGraphInfos.Item(7)
      gr.Format catTitle:=trads.TranslatedValue("MSG_Period"), _ 
                valuesTitle:=trads.TranslatedValue("MSG_Values"), _ 
                plotTitle:=plotTitle, scope:=GraphScopeTimeSeries
      
      'Add the graph title to graph lists
      cellRng.Value = plotTitle

      'Format the plot title cell
      design.ApplyFormat cellRng, scope:=LInelistHiddenCell
      Push GoToLabel("graph") & ": " & plotTitle, scope:="graph"
      Set cellRng = cellRng.Offset(50)
    Next
  End If

  'Add the GoTo graph dropdown
  AddGoTo outsh.Cells(3, 1), scope:="graph", _
          goToPrefix:="ts_", goTosuffix:="_graph"
End Sub


'Write the table

Private Sub WriteAnalysis()
  
  Dim outsh As Worksheet
  Dim cellRng As Range
  Dim scope As AnalysisScope

  'Univariate and Bivariate Analysis worksheet ==================
  scope = AnalysisScopeNormal
  Set outsh = OutputSh(scope)
  ' --- Prepare the worksheet
   PrepareSheet scope:=scope
  ' --- Write Global summary
   WriteTable LoName:=GLOBALSUMMARY, scope:=scope
   ' ---- Write Univariate Analysis
   WriteTable LoName:=UNIVARIATE, scope:=scope
   ' ---- Write the Bivariate Analysis
   WriteTable LoName:=BIVARIATE, scope:=scope
   'Add The goTo for sections
   Set cellRng = outsh.Cells(2, 1)
   AddGoTo cellRng:=cellRng, scope:=scope
   'Add the command
   AddCommands scope:=scope

   'Time series analysis
  scope = AnalysisScopeTimeSeries
  Set outsh = OutputSh(scope)
  PrepareSheet scope:=scope
  'Write the tables of time series
  WriteTable LoName:=TIMESERIES, scope:=scope
  'Write the graphs
  WriteTimeSeriesGraphs graphLoName:=GRAPHTIMESERIES, _ 
                         graphTitleLoName:=TITLEGRAPHTIMESERIES, _ 
                         tabLoName:=TIMESERIES
  
  Set cellRng = outsh.Cells(2, 1)
  AddGoTo cellRng:=cellRng, scope:=scope
  'Add the commands
  AddCommands scope:=scope

  'Spatial Analysis
  scope = AnalysisScopeSpatial
  Set outsh = OutputSh(scope)
  PrepareSheet scope:=scope
  WriteTable LoName:=SPATIAL, scope:=scope
  Set cellRng = outsh.Cells(2, 1)
  AddGoTo cellRng:=cellRng, scope:=scope
  AddCommands scope:=scope
End Sub


'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, className, "Error : " & errorMessage
End Sub

Public Sub IAnalysisOutput_WriteAnalysis()
  'write analysis 
  WriteAnalysis
End Sub
