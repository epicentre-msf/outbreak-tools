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
    Set Outputsh = ll.Wksh(trads.TranslatedValue("LLSHEET_SpatioTemporalAnalysis"))
  End Select

End Property

'Speed Up the application
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

Private Sub TransferCodeWksh(Optional ByVal scope _
                             As AnalysisScope = AnalysisScopeNormal)

    Const CHANGEMODULENAME = "LinelistAnalysisChange"

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
    If scope = AnalysisScopeTimeSeries Or scope = AnalysisScopeSpatioTemporal Then .SplitColumn = 3
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
  Set but = Buttons.Create(outRng:=cellRng, _
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
  Dim design As ILLFormat

  Select Case scope
  Case AnalysisScopeTimeSeries
    sheetInfo = "TS-Analysis"
  Case AnalysisScopeSpatial
    sheetInfo = "SP-Analysis"
  Case AnalysisScopeNormal
    sheetInfo = "Uni-Bi-Analysis"
  Case AnalysisScopeSpatioTemporal
    sheetInfo = "SPT-Analysis"
  End Select

  Set outsh = OutputSh(scope)
  Set design = DesignFormat()

  'Create a range called RGN_HasTable whith value Yes if there are tables
  Set cellRng = outsh.Cells(1, 10)
  cellRng.Name = "RNG_HasTable"
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
  Dim trads As ITranslation
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
  Set trads = TransObject()

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
Private Sub WriteTable(ByVal LoName As String, _
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
  Dim Lo As ListObject
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
  Dim design As ILLFormat

  'Variables for Time Series Tables
  Dim tableName As String
  Dim rowVar As String
  Dim minFormula As String
  Dim maxFormula As String
  Dim lastValid As ITablesSpecs
  Dim prevSpecs As ITablesSpecs 'Previous specs (could be Nothing)
  Dim prevSecId As String 'Previous table Id
  Dim dict As ILLdictionary
  Dim vars As ILLVariables
  Dim goToHeader As String
  Dim grHeight As Long

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
  Set design = DesignFormat()

  goToSection = GoToLabel("section")
  goToHeader = GoToLabel("header")

  If Lo.DataBodyRange Is Nothing Then Exit Sub
  Set hRng = Lo.HeaderRowRange()

  For tabCounter = 1 To Lo.ListRows.Count
    Set tabRng = Lo.ListRows(tabCounter).Range
    Set specs = TablesSpecs.Create(hRng, tabRng, lData)
    tabType = specs.TableType()

    If specs.ValidTable() Then

      If specs.isNewSection() Then
        Push goToSection & ": " & specs.Value("section"), scope:="section"
        'The last valid specs is the specs of the last valid new section for TS
        Set lastValid = specs
      End If

      Set tabl = CrossTable.Create(specs, outsh, trads)
      tabl.Build
      'Add Formulas
      Set tabform = CrossTableFormula.Create(tabl, formData)
      tabform.AddFormulas
      'Format the created table
      tabl.Format design

      'Will be used to test if a sheet has a table
      If IsEmpty(outsh.Range("RNG_HasTable")) Then  _
      outsh.Range("RNG_HasTable").Value = "yes"

      'Minimum and maximum formula for Time series; GoTo Headers for time series
      If (tabType = TypeTimeSeries) Then
        'Get the previous specs for time series
        On Error Resume Next
        Set prevSpecs = specs.Previous()
        On Error GoTo 0

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

         'The graph height depends on the number of elements in Geo,
         'but could result in error if n geo is not found
         grHeight = 1
         On Error Resume Next
         grHeight = CLng(specs.Value("n geo"))
         On Error GoTo 0

        'cellRng is where to write the graph on the worksheet
        Set cellRng = outsh.Cells(tabl.StartRow - 2, tabl.EndColumn + 2)
        tabId = specs.TableId()

        'For spatial table move three columns to keep room for deviding by population
        If tabType = TypeSpatial Then Set cellRng = cellRng.Offset(, 2)

        'Create the specs of the graph and add it
        Set grSpecs = GraphSpecs.Create(tabl)
        'Create series uses informations on the table to create columns of
        'series of ranges to plot.
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
        scope:=scope, heightFactor:=grHeight
      End If
    End If
  Next

  'Add minimum and maximum formula for last section if on time series
  If (Not lastValid Is Nothing) And (tabType = TypeTimeSeries) Then
    tabId = lastValid.TableSectionId
    outsh.Range("MAX_MAX_DATE_" & tabId).formula = "= MAX(" & maxFormula & ")"
    outsh.Range("MIN_MIN_DATE_" & tabId).formula = "= MIN(" & minFormula & ")"
    Set cellRng = outsh.Range("SECTION_" & tabId).Offset(1, -2)
    'Add dropdown
    AddGoTo cellRng:=cellRng, goToPrefix:="ts_", scope:="header", _
                   goToSuffix:=tabId
    'Add an anchor to where to start building graphs for time series (START_GRAPHS_TS)
    Set cellRng = outsh.Range("ENDTABLE_" & lastValid.TableId).Offset(5, 2)
    cellRng.Name = "START_GRAPHS_TS"
  End If

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
  Dim grLo As ListObject
  Dim tabLo As ListObject
  Dim grTitleLo As ListObject
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

  'Worksheets
  Set sh = Wksh()
  Set outsh = OutputSh(AnalysisScopeTimeSeries)

  If IsEmpty(outsh.Range("RNG_HasTable")) Then Exit Sub

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
   WriteTable LoName:=GLOBALSUMMARY, scope:=scope
   ' ---- Write Univariate Analysis
   WriteTable LoName:=UNIVARIATE, scope:=scope
   ' ---- Write the Bivariate Analysis
   WriteTable LoName:=BIVARIATE, scope:=scope

   On Error GoTo 0

   'Add The goTo for sections
   Set cellRng = outsh.Cells(2, 1)
   AddGoTo cellRng:=cellRng, goToprefix:="ua_", scope:="section"
   'Add the command
   AddCommands scope:=scope

   'Hide the output sheet if required
   pass.Protect outsh, True
   If IsEmpty(outsh.Range("RNG_HasTable")) Then outsh.Visible = xlSheetHidden

   'Transfer code to the worksheet
   TransferCodeWksh scope:=scope

   'Time Series Analysis Worksheet =============================================
  scope = AnalysisScopeTimeSeries
  Set outsh = OutputSh(scope)
  PrepareSheet scope:=scope
  'Write the tables of time series
  On Error GoTo ErrHand

  WriteTable LoName:=TIMESERIES, scope:=scope
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
  If IsEmpty(outsh.Range("RNG_HasTable")) Then outsh.Visible = xlSheetHidden

  'Transfert code to worksheet
   TransferCodeWksh scope:=scope

  'Spatial Analysis Worksheet ==================================================
  scope = AnalysisScopeSpatial
  Set outsh = OutputSh(scope)

  PrepareSheet scope:=scope

  On Error GoTo ErrHand
  WriteTable LoName:=SPATIAL, scope:=scope
  On Error GoTo 0

  Set cellRng = outsh.Cells(2, 1)

  AddGoTo cellRng:=cellRng, scope:="section", goToPrefix:="sp_"
  AddCommands scope:=scope

  pass.Protect outsh, True
  If IsEmpty(outsh.Range("RNG_HasTable")) Then outsh.Visible = xlSheetHidden

  'Transfer code to spatial analysis sheet
  TransferCodeWksh scope:=scope

  'Spatio-Temporal Analysis worksheet ==========================================
  scope = AnalysisScopeSpatioTemporal
  Set outsh = Outputsh(scope)
  PrepareSheet scope:=scope

  On Error GoTo ErrHand
  WriteTable LoName:=SPATIOTEMPORAL, scope:=scope
  On Error GoTo 0

  Set cellRng = outsh.Cells(2, 1)
  AddGoTo cellRng:=cellRng, scope:="section", goToPrefix:="spt_"
  AddCommands scope:=scope
  
  pass.Protect outsh, True
  If IsEmpty(outsh.Range("RNG_HasTable")) Then outsh.Visible = xlSheetHidden
  Exit Sub

ErrHand:
    Debug.Print "Error: " & Err.Number & " " &  Err.Description
    Resume Next
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

Public Sub IAnalysisOutput_WriteAnalysis()
  'write analysis
  WriteAnalysis
End Sub