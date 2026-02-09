VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Build and format all analysis worksheets"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Build and format all analysis worksheets")
'@IgnoreModule UnrecognizedAnnotation, IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IAnalysisOutput

Private Const UPDATEFILTERCOMMAND As String = "UpdateFilterTables"
Private Const SHPFILTERNAME As String = "SHP_Filter"
Private Const GOTOSECCODE As String = "go_to_section"
Private Const UPDATESPCOMMAND As String = "UpdateSpTables"
Private Const CLASS_NAME As String = "AnalysisOutput"

' ListObject names in the setup worksheet
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
    ' Bug #2 fix: GoTo arrays are now local to each scope, not shared instance state.
    ' They are passed as parameters to WriteTable and managed per-scope in WriteAnalysis.
End Type

Private this As TAnaOutput

'@section Factory
'===============================================================================

Public Function Create(ByVal specSh As Worksheet, _
                       ByVal ll As ILinelist) As IAnalysisOutput

    If specSh Is Nothing Then ThrowError InvalidArgument, "Specifications worksheet not defined"
    If ll Is Nothing Then ThrowError InvalidArgument, "Linelist not defined"

    With New AnalysisOutput
        Set .OutputLL = ll
        Set .Wksh = specSh
        Set Create = .Self
    End With
End Function

'@section Internal properties (factory use)
'===============================================================================

Public Property Get Self() As IAnalysisOutput
    Set Self = Me
End Property

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

'@section Dependency properties
'===============================================================================

Private Property Get LinelistData() As ILinelistSpecs
    Set LinelistData = OutputLL().LinelistData()
End Property

Private Property Get DesignFormat() As ILLFormat
    Set DesignFormat = LinelistData().DesignFormat()
End Property

Public Property Get TransObject(Optional ByVal scope As Byte = TranslationOfMessages) As ITranslationObject
    Set TransObject = LinelistData().TransObject(scope)
End Property

Private Property Get OutputSh(Optional ByVal scope As Byte = AnalysisScopeNormal) As Worksheet
    Dim trads As ITranslationObject
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

Private Property Get NamesObject() As IAnaTabIds
    Dim lData As ILinelistSpecs
    Dim ll As ILinelist
    Dim sh As Worksheet

    Set ll = OutputLL()
    Set lData = ll.LinelistData()
    Set sh = ll.Wksh(lData.TemporarySheetName(TemporarySheetAnalysis))
    Set NamesObject = AnaTabIds.Create(sh, False)
End Property

'@section Application helpers
'===============================================================================

Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

Private Sub TransferCodeWksh(Optional ByVal scope As Byte = AnalysisScopeNormal)
    Const CHANGEMODULENAME As String = "LinelistAnalysisChange"

    Dim sh As Worksheet
    Dim inpWkb As Workbook
    Dim lData As ILinelistSpecs
    Dim codeContent As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object
    Dim sheetName As String

    Set lData = LinelistData()
    Set inpWkb = lData.SpecsWkb()
    Set sh = OutputSh(scope)
    sheetName = sh.Name

    With inpWkb.VBProject.VBComponents(CHANGEMODULENAME).CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With

    With sh.Parent
        Set vbProj = .VBProject
        Set vbComp = vbProj.VBComponents(.Sheets(sheetName).CodeName)
        Set codeMod = vbComp.CodeModule
    End With

    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub

Private Sub SplitOutputSheet(Optional ByVal scope As Byte = AnalysisScopeNormal)
    Dim outsh As Worksheet
    Dim rowSplit As Long

    Set outsh = OutputSh(scope)
    rowSplit = IIf(scope = AnalysisScopeTimeSeries Or scope = AnalysisScopeSpatioTemporal, 3, 2)

    On Error Resume Next
    BusyApp
    outsh.Activate
    BusyApp
    With Application.ActiveWindow
        .SplitRow = rowSplit
        .SplitColumn = 3
        .Zoom = 80
        .FreezePanes = True
    End With
    On Error GoTo 0
End Sub

'@section Sheet preparation
'===============================================================================

Private Sub AddCommands(Optional ByVal scope As Byte = AnalysisScopeNormal)
    Dim cellRng As Range
    Dim commandName As String
    Dim shapeLabel As String
    Dim design As ILLFormat
    Dim but As IButtons
    Dim outsh As Worksheet
    Dim trads As ITranslationObject
    Dim mainobj As IMain
    Dim lData As ILinelistSpecs

    Set lData = LinelistData()
    Set mainobj = lData.MainObject()
    If mainobj.HasTemplate() Then Exit Sub

    Set design = DesignFormat()
    Set outsh = OutputSh(scope)
    Set trads = TransObject(TranslationOfShapes)
    Set cellRng = outsh.Cells(2, 2)

    commandName = IIf(scope = AnalysisScopeSpatial, UPDATESPCOMMAND, UPDATEFILTERCOMMAND)
    shapeLabel = trads.TranslatedValue(SHPFILTERNAME)
    Set but = Buttons.Create(outrng:=cellRng, codeName:=SHPFILTERNAME, scope:=ButtonScopeSmall)
    but.Add actionCommand:=commandName, shapeLabel:=shapeLabel
    but.Format design
End Sub

Private Sub PrepareSheet(Optional ByVal scope As Byte = AnalysisScopeNormal)
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

    Set cellRng = outsh.Cells(1, 10)
    cellRng.Name = rngName
    design.ApplyFormat cellRng, scope:=LinelistHiddenCell

    design.ApplyFormat outsh, scope:=AllAnalysisSheet

    Set cellRng = outsh.Cells(1, 3)
    cellRng.Value = sheetInfo
    design.ApplyFormat cellRng, scope:=LinelistHiddenCell

    Set cellRng = outsh.Cells(2, 5)
    cellRng.Formula = "= ComputedOnFiltered()"
    design.ApplyFormat cellRng, scope:=LinelistWarning

    AddCommands scope:=scope
    SplitOutputSheet scope:=scope
End Sub

'@section GoTo dropdown helpers
'===============================================================================

Private Sub AddGoTo(ByVal cellRng As Range, _
                    ByVal dropArray As BetterArray, _
                    ByVal goToLabel As String, _
                    Optional ByVal goToPrefix As String = "ua_", _
                    Optional ByVal goToSuffix As String = vbNullString)

    Dim drop As IDropdownLists
    Dim ll As ILinelist
    Dim dropName As String
    Dim design As ILLFormat

    If dropArray.Length = 0 Then Exit Sub

    Set ll = OutputLL()
    Set drop = ll.Dropdown()
    Set design = DesignFormat()

    dropName = goToPrefix & "gotosection" & goToSuffix

    drop.Add dropArray, dropName
    drop.SetValidation cellRng, dropName

    design.ApplyFormat cellRng, scope:=LinelistSelectionDropdown

    cellRng.Value = goToLabel
    cellRng.Name = goToPrefix & GOTOSECCODE & goToSuffix
End Sub

'@section Min/Max formula for temporal sections
'===============================================================================

Private Sub MinMaxFormula(ByVal outsh As Worksheet, ByVal secId As String, _
                          ByVal minFormula As String, ByVal maxFormula As String, _
                          ByVal goToPrefix As String, _
                          ByVal headersList As BetterArray, _
                          ByVal goToHeaderLabel As String)
    Dim cellRng As Range

    outsh.Range("MAX_MAX_DATE_" & secId).Formula = "= MAX(" & maxFormula & ")"
    outsh.Range("MIN_MIN_DATE_" & secId).Formula = "= MIN(" & minFormula & ")"

    ' Add goTo Header dropdown to this section
    Set cellRng = outsh.Range("SECTION_" & secId).Offset(2)
    AddGoTo cellRng:=cellRng, dropArray:=headersList, goToLabel:=goToHeaderLabel, _
            goToPrefix:=goToPrefix, goToSuffix:=secId
End Sub

'@section WriteTable — main table creation loop
'===============================================================================

Private Sub WriteTable(ByVal loName As String, _
                       ByVal sectionsList As BetterArray, _
                       ByVal headersList As BetterArray, _
                       Optional ByVal scope As Byte = AnalysisScopeNormal)
    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim tabRng As Range
    Dim hRng As Range
    Dim graphCounter As Long
    Dim tabl As ICrossTable
    Dim tabform As ICrossTableFormula
    Dim gr As IGraphs
    Dim lo As ListObject
    Dim specs As ITableSpecs
    Dim lData As ILinelistSpecs
    Dim trads As ITranslationObject
    Dim formData As IFormulaData
    Dim cellRng As Range
    Dim grSpecs As IGraphSpecs
    Dim tabId As String
    Dim tabType As Byte
    Dim design As ILLFormat
    Dim anchorRng As Range
    Dim containsTimeSeries As Boolean
    Dim spatRng As Range
    Dim sheetInfo As String
    Dim namObj As IAnaTabIds
    Dim tabRangesNames As BetterArray
    Dim goToSectionLabel As String
    Dim goToHeaderLabel As String

    ' Variables for time series
    Dim tableName As String
    Dim rowVar As String
    Dim minFormula As String
    Dim maxFormula As String
    Dim prevSpecs As ITableSpecs
    Dim nextSpecs As ITableSpecs
    Dim prevSecId As String
    Dim secId As String
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim grHeight As Long
    Dim isNewSection As Boolean
    Dim addPrevSecMinMaxFormula As Boolean
    Dim addActSecMinMaxFormula As Boolean
    Dim addSpatioTemporalName As Boolean
    Dim upRngName As String
    Dim goToPrefix As String

    Set sh = Wksh()
    Set outsh = OutputSh(scope)
    Set lo = sh.ListObjects(loName)
    Set lData = LinelistData()
    Set trads = TransObject()
    Set formData = lData.FormulaDataObject()
    Set dict = lData.Dictionary()
    Set vars = LLVariables.Create(dict)
    Set design = DesignFormat()
    Set namObj = NamesObject()

    sheetInfo = outsh.Cells(1, 3).Value
    upRngName = "RNG_HasTable_" & UCase(Split(sheetInfo, "-")(0))
    goToSectionLabel = trads.TranslatedValue("MSG_GoToSec")
    goToHeaderLabel = trads.TranslatedValue("MSG_GoToHead")

    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set hRng = lo.HeaderRowRange()
    Set anchorRng = hRng.Offset(lo.ListRows.Count)
    Set tabRng = lo.ListRows(1).Range

    Do
        Set specs = TableSpecs.Create(hRng, tabRng, lData)
        Set nextSpecs = Nothing
        addPrevSecMinMaxFormula = False
        addActSecMinMaxFormula = False
        addSpatioTemporalName = False

        If Not specs.ValidTable() Then GoTo ContinueLoop

        tabType = specs.TableType()
        isNewSection = specs.IsNewSection()
        containsTimeSeries = (tabType = TypeTimeSeries Or tabType = TypeSpatioTemporal)
        tabId = specs.TableId()

        ' Bug #2 fix: push to LOCAL sectionsList, not shared instance array
        If isNewSection Then sectionsList.Push goToSectionLabel & ": " & specs.Value("section")

        ' Bug #3 fix: per-table error handling — wrap table creation in error handler
        On Error GoTo TableError

        Set tabl = CrossTable.Create(specs, outsh, trads)
        tabl.Build
        Set tabform = CrossTableFormula.Create(tabl, formData)
        tabform.AddFormulas
        tabl.Format design

        On Error GoTo 0

        Set tabRangesNames = tabl.NamedRangesList()
        namObj.AddTableInfos scope, tabId, tabRangesNames

        If IsEmpty(outsh.Range(upRngName)) Then outsh.Range(upRngName).Value = "yes"

        ' Non-temporal graphs
        If specs.HasGraph() And (Not containsTimeSeries) Then
            grHeight = 1
            On Error Resume Next
            grHeight = CLng(specs.Value("n geo"))
            On Error GoTo 0

            Set cellRng = outsh.Cells(tabl.StartRow - 2, tabl.EndColumn + 2)
            If tabType = TypeSpatial Then Set cellRng = cellRng.Offset(, 2)

            Set grSpecs = GraphSpecs.Create(tabl)
            grSpecs.CreateSeries
            Set gr = Graphs.Create(outsh, cellRng)
            gr.Add

            For graphCounter = 1 To grSpecs.NumberOfSeries()
                gr.AddSeries grSpecs.SeriesName(graphCounter), _
                             grSpecs.SeriesType(graphCounter), _
                             grSpecs.SeriesPos(graphCounter)
                gr.AddLabels grSpecs.SeriesLabel(graphCounter), _
                             grSpecs.SeriesColumnLabel(graphCounter), _
                             hardCodeLabels:=False

                namObj.AddGraphInfo scope:=scope, tabId:=tabId, graphId:=tabId, _
                                    seriesName:=grSpecs.SeriesName(graphCounter), _
                                    seriesType:=grSpecs.SeriesType(graphCounter), _
                                    seriesPos:=grSpecs.SeriesPos(graphCounter), _
                                    seriesLabel:=grSpecs.SeriesLabel(graphCounter), _
                                    seriesColumnLabel:=grSpecs.SeriesColumnLabel(graphCounter), _
                                    hardCodeLabels:=False, _
                                    outRangeAddress:=cellRng.Address
            Next

            gr.Format catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                      valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
                      scope:=scope, heightFactor:=grHeight

            namObj.AddGraphFormat scope:=scope, tabId:=tabId, graphId:=tabId, _
                                  catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                                  valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
                                  heightFactor:=grHeight
        End If

        ' Time series / spatio-temporal min/max and GoTo logic
        If Not containsTimeSeries Then GoTo ContinueLoop

        On Error Resume Next
        Set prevSpecs = specs.Previous()
        Set nextSpecs = specs.NextSpecs(anchorRng)
        prevSecId = prevSpecs.TableSectionId()
        secId = specs.TableSectionId()
        On Error GoTo 0

        rowVar = specs.Value("row")
        tableName = vars.Value(colName:="table name", varName:=rowVar)
        goToPrefix = IIf(tabType = TypeTimeSeries, "ts_", "spt_")

        If prevSpecs Is Nothing Then
            minFormula = "MIN(" & tableName & "[" & rowVar & "])"
            maxFormula = "MAX(" & tableName & "[" & rowVar & "])"
            headersList.Push goToHeaderLabel & ": " & tabl.TimeSeriesHeader()
            addPrevSecMinMaxFormula = False
        End If

        If isNewSection And (Not prevSpecs Is Nothing) Then
            addPrevSecMinMaxFormula = True
            addSpatioTemporalName = (tabType = TypeSpatioTemporal)
        End If

        If Not nextSpecs Is Nothing Then
            If nextSpecs.IsNewSection() And isNewSection Then
                addActSecMinMaxFormula = True
                addSpatioTemporalName = (tabType = TypeSpatioTemporal)
            End If
        Else
            addActSecMinMaxFormula = True
            addSpatioTemporalName = (tabType = TypeSpatioTemporal)
            Set cellRng = outsh.Range("ENDTABLE_" & secId).Offset(5, 2)
            If tabType = TypeTimeSeries Then cellRng.Name = "START_GRAPHS_TS"
            If tabType = TypeSpatioTemporal Then cellRng.Name = "START_GRAPHS_SPT"
            headersList.Push goToHeaderLabel & ": " & tabl.TimeSeriesHeader()
        End If

        If (Not isNewSection) And (Not nextSpecs Is Nothing) Then
            addPrevSecMinMaxFormula = False
            addActSecMinMaxFormula = False
            minFormula = minFormula & ", MIN(" & tableName & "[" & rowVar & "])"
            maxFormula = maxFormula & ", MAX(" & tableName & "[" & rowVar & "])"
            headersList.Push goToHeaderLabel & ": " & tabl.TimeSeriesHeader()
        End If

        If addPrevSecMinMaxFormula Then
            MinMaxFormula outsh:=outsh, secId:=prevSecId, _
                          minFormula:=minFormula, maxFormula:=maxFormula, _
                          goToPrefix:=goToPrefix, headersList:=headersList, _
                          goToHeaderLabel:=goToHeaderLabel
            ' Bug #2 fix: clear headers for the new section after writing previous section's dropdown
            headersList.Clear
        End If

        If addActSecMinMaxFormula Then
            MinMaxFormula outsh:=outsh, secId:=secId, _
                          minFormula:=minFormula, maxFormula:=maxFormula, _
                          goToPrefix:=goToPrefix, headersList:=headersList, _
                          goToHeaderLabel:=goToHeaderLabel
            headersList.Clear
        End If

        If isNewSection Then
            minFormula = "MIN(" & tableName & "[" & rowVar & "])"
            maxFormula = "MAX(" & tableName & "[" & rowVar & "])"
        End If

        If addSpatioTemporalName Then
            With outsh
                Set cellRng = .Range("COLUMN_CATEGORIES_" & tabId)
                Set spatRng = .Range( _
                    .Range("COLUMN_CATEGORIES_" & secId).Cells(1, 1), _
                    cellRng.Cells(1, cellRng.Columns.Count))
            End With
            spatRng.Name = "SPT_FORMULA_COLUMN_" & secId
        End If

        GoTo ContinueLoop

TableError:
        ' Bug #3 fix: per-table error handling — log error and continue to next table
        Debug.Print "Error on table " & tabId & ": " & Err.Number & " " & Err.Description
        On Error GoTo 0
        Resume ContinueLoop

ContinueLoop:
        Set tabRng = tabRng.Offset(1)
    Loop While tabRng.Row <= anchorRng.Row
End Sub

'@section WriteTimeSeriesGraphs
'===============================================================================

Private Sub WriteTimeSeriesGraphs(ByVal graphLoName As String, _
                                   ByVal graphTitleLoName As String, _
                                   ByVal tabLoName As String, _
                                   ByVal graphsList As BetterArray)
    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim graphCounter As Long
    Dim cellRng As Range
    Dim lData As ILinelistSpecs
    Dim seriesCounter As Long

    Dim grLo As ListObject
    Dim tabLo As ListObject
    Dim grTitleLo As ListObject
    Dim grSpecs As IGraphSpecs
    Dim gr As IGraphs
    Dim trads As ITranslationObject
    Dim design As ILLFormat
    Dim namObj As IAnaTabIds
    Dim goToGraphLabel As String

    Dim grBAinput As BetterArray
    Dim grBAoutput As BetterArray
    Dim tabGraphInfos As BetterArray
    Dim tabSeries As BetterArray
    Dim tabTypes As BetterArray
    Dim tabPos As BetterArray
    Dim tabLabel As BetterArray
    Dim tabColLabel As BetterArray
    Dim tabPrefLabel As BetterArray
    Dim plotTitle As String

    Set sh = Wksh()
    Set outsh = OutputSh(AnalysisScopeTimeSeries)

    If IsEmpty(outsh.Range("RNG_HasTable_TS")) Then Exit Sub

    Set grBAinput = New BetterArray
    Set namObj = NamesObject()
    Set lData = LinelistData()
    Set cellRng = outsh.Range("START_GRAPHS_TS")
    Set grLo = sh.ListObjects(graphLoName)
    Set tabLo = sh.ListObjects(tabLoName)
    Set grTitleLo = sh.ListObjects(graphTitleLoName)
    Set trads = TransObject()
    Set design = DesignFormat()
    goToGraphLabel = trads.TranslatedValue("MSG_GoToGraph")

    If (grLo.DataBodyRange Is Nothing) Or (grTitleLo.DataBodyRange Is Nothing) Then Exit Sub

    grBAinput.Push grLo, tabLo, grTitleLo
    Set grSpecs = GraphSpecs.CreateRangeSpecs(grBAinput, outsh, lData)

    If grSpecs.Valid() Then
        Set grBAoutput = grSpecs.SpecsLists()
        For graphCounter = 1 To grSpecs.NumberOfGraphs()
            Set tabGraphInfos = grBAoutput.Item(graphCounter)
            Set tabSeries = tabGraphInfos.Item(1)
            Set tabTypes = tabGraphInfos.Item(2)
            Set tabPos = tabGraphInfos.Item(3)
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

                namObj.AddGraphInfo scope:=AnalysisScopeTimeSeries, _
                                    tabId:=graphCounter, graphId:=graphCounter, _
                                    seriesName:=tabSeries.Item(seriesCounter), _
                                    seriesType:=tabTypes.Item(seriesCounter), _
                                    seriesPos:=tabPos.Item(seriesCounter), _
                                    seriesLabel:=tabLabel.Item(seriesCounter), _
                                    seriesColumnLabel:=vbNullString, _
                                    outRangeAddress:=cellRng.Address, _
                                    prefix:=tabPrefLabel.Item(seriesCounter), _
                                    prefixOnly:=True
            Next

            plotTitle = tabGraphInfos.Item(7)
            gr.Format catTitle:=trads.TranslatedValue("MSG_Period"), _
                      valuesTitle:=trads.TranslatedValue("MSG_Values"), _
                      plotTitle:=plotTitle, scope:=GraphScopeTimeSeries

            cellRng.Value = plotTitle

            namObj.AddGraphFormat scope:=AnalysisScopeTimeSeries, _
                                  tabId:=graphCounter, graphId:=graphCounter, _
                                  catTitle:=trads.TranslatedValue("MSG_Period"), _
                                  valuesTitle:=trads.TranslatedValue("MSG_Values"), _
                                  plotTitle:=plotTitle

            design.ApplyFormat cellRng, scope:=LinelistHiddenCell
            graphsList.Push goToGraphLabel & ": " & plotTitle
            Set cellRng = cellRng.Offset(50)
        Next
    End If

    ' Add GoTo graph dropdown
    AddGoTo outsh.Cells(3, 3), dropArray:=graphsList, _
            goToLabel:=goToGraphLabel, goToPrefix:="ts_", goToSuffix:="_graph"
End Sub

'@section WriteSpatioTemporalGraph
'===============================================================================

Private Sub WriteSpatioTemporalGraph(ByVal loName As String, _
                                      ByVal graphsList As BetterArray)
    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim cellRng As Range
    Dim specs As ITableSpecs
    Dim lData As ILinelistSpecs
    Dim design As ILLFormat
    Dim hRng As Range
    Dim tabRng As Range
    Dim anchorRng As Range
    Dim gr As IGraphs
    Dim grSpecs As IGraphSpecs
    Dim plotTitle As String
    Dim tabId As String
    Dim graphCounter As Long
    Dim trads As ITranslationObject
    Dim lo As ListObject
    Dim tabl As ICrossTable
    Dim namObj As IAnaTabIds
    Dim goToGraphLabel As String

    Set sh = Wksh()
    Set outsh = OutputSh(AnalysisScopeSpatioTemporal)
    Set lo = sh.ListObjects(loName)
    Set lData = LinelistData()
    Set trads = TransObject()
    Set design = DesignFormat()
    Set namObj = NamesObject()
    goToGraphLabel = trads.TranslatedValue("MSG_GoToGraph")

    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set cellRng = outsh.Range("START_GRAPHS_SPT")
    Set hRng = lo.HeaderRowRange()
    Set anchorRng = hRng.Offset(lo.ListRows.Count)
    Set tabRng = lo.ListRows(1).Range

    Do
        Set specs = TableSpecs.Create(hRng, tabRng, lData)

        If Not specs.ValidTable() Then GoTo ContinueLoop
        If Not specs.HasGraph() Then GoTo ContinueLoop

        tabId = specs.TableId()
        Set tabl = CrossTable.Create(specs, outsh, trads)

        Set grSpecs = GraphSpecs.Create(tabl)
        plotTitle = trads.TranslatedValue("MSG_Graph") & _
                    outsh.Range("STARTCOL_" & tabId).Value

        grSpecs.CreateSeries
        Set gr = Graphs.Create(outsh, cellRng)
        gr.Add

        For graphCounter = 1 To grSpecs.NumberOfSeries()
            gr.AddSeries grSpecs.SeriesName(graphCounter), _
                         grSpecs.SeriesType(graphCounter), _
                         grSpecs.SeriesPos(graphCounter)
            gr.AddLabels grSpecs.SeriesLabel(graphCounter), _
                         grSpecs.SeriesColumnLabel(graphCounter), _
                         hardCodeLabels:=False

            namObj.AddGraphInfo scope:=AnalysisScopeSpatioTemporal, tabId:=tabId, _
                                graphId:=tabId, _
                                seriesName:=grSpecs.SeriesName(graphCounter), _
                                seriesType:=grSpecs.SeriesType(graphCounter), _
                                seriesPos:=grSpecs.SeriesPos(graphCounter), _
                                seriesLabel:=grSpecs.SeriesLabel(graphCounter), _
                                seriesColumnLabel:=grSpecs.SeriesColumnLabel(graphCounter), _
                                hardCodeLabels:=False, _
                                outRangeAddress:=cellRng.Address
        Next

        gr.Format catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                  valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
                  scope:=GraphScopeTimeSeries, plotTitle:=plotTitle

        namObj.AddGraphFormat scope:=AnalysisScopeSpatioTemporal, tabId:=tabId, _
                              graphId:=tabId, _
                              catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                              valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
                              heightFactor:=1, plotTitle:=plotTitle

        cellRng.Value = plotTitle
        design.ApplyFormat cellRng, scope:=LinelistHiddenCell
        graphsList.Push goToGraphLabel & ": " & plotTitle
        Set cellRng = cellRng.Offset(50)

ContinueLoop:
        Set tabRng = tabRng.Offset(1)
    Loop While tabRng.Row <= anchorRng.Row

    AddGoTo outsh.Cells(3, 3), dropArray:=graphsList, _
            goToLabel:=goToGraphLabel, goToPrefix:="spt_", goToSuffix:="_graph"
End Sub

'@section WriteAnalysis — main orchestration
'===============================================================================

Private Sub WriteAnalysis()
    Dim outsh As Worksheet
    Dim cellRng As Range
    Dim scope As Byte
    Dim pass As IPasswords
    Dim lData As ILinelistSpecs
    Dim trads As ITranslationObject
    Dim goToSectionLabel As String

    ' Bug #2 fix: GoTo arrays are LOCAL to each scope
    Dim sectionsList As BetterArray
    Dim headersList As BetterArray
    Dim graphsList As BetterArray

    Set lData = LinelistData()
    Set pass = lData.Password()
    Set trads = TransObject()
    goToSectionLabel = trads.TranslatedValue("MSG_GoToSec")

    ' Univariate and Bivariate Analysis ========================================
    scope = AnalysisScopeNormal
    Set outsh = OutputSh(scope)
    PrepareSheet scope:=scope

    ' Bug #2 fix: fresh arrays per scope
    Set sectionsList = New BetterArray
    Set headersList = New BetterArray

    WriteTable loName:=GLOBALSUMMARY, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope
    sectionsList.Push goToSectionLabel & ""
    WriteTable loName:=UNIVARIATE, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope
    WriteTable loName:=BIVARIATE, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope

    Set cellRng = outsh.Cells(2, 3)
    AddGoTo cellRng, dropArray:=sectionsList, goToLabel:=goToSectionLabel, goToPrefix:="ua_"
    AddCommands scope:=scope

    pass.Protect outsh, True
    If IsEmpty(outsh.Range("RNG_HasTable_UNI")) Then outsh.Visible = xlSheetHidden
    TransferCodeWksh scope:=scope

    ' Time Series Analysis =====================================================
    scope = AnalysisScopeTimeSeries
    Set outsh = OutputSh(scope)
    PrepareSheet scope:=scope

    Set sectionsList = New BetterArray
    Set headersList = New BetterArray
    Set graphsList = New BetterArray

    WriteTable loName:=TIMESERIES, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope

    WriteTimeSeriesGraphs graphLoName:=GRAPHTIMESERIES, _
                          graphTitleLoName:=TITLEGRAPHTIMESERIES, _
                          tabLoName:=TIMESERIES, graphsList:=graphsList

    Set cellRng = outsh.Cells(2, 3)
    AddGoTo cellRng, dropArray:=sectionsList, goToLabel:=goToSectionLabel, goToPrefix:="ts_"
    AddCommands scope:=scope

    pass.Protect outsh, True
    If IsEmpty(outsh.Range("RNG_HasTable_TS")) Then outsh.Visible = xlSheetHidden
    TransferCodeWksh scope:=scope

    ' Spatial Analysis =========================================================
    scope = AnalysisScopeSpatial
    Set outsh = OutputSh(scope)
    PrepareSheet scope:=scope

    Set sectionsList = New BetterArray
    Set headersList = New BetterArray

    WriteTable loName:=SPATIAL, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope

    Set cellRng = outsh.Cells(2, 3)
    AddGoTo cellRng, dropArray:=sectionsList, goToLabel:=goToSectionLabel, goToPrefix:="sp_"
    AddCommands scope:=scope

    pass.Protect outsh, True
    If IsEmpty(outsh.Range("RNG_HasTable_SP")) Then outsh.Visible = xlSheetHidden
    TransferCodeWksh scope:=scope

    ' Spatio-Temporal Analysis =================================================
    scope = AnalysisScopeSpatioTemporal
    Set outsh = OutputSh(scope)
    PrepareSheet scope:=scope

    Set sectionsList = New BetterArray
    Set headersList = New BetterArray
    Set graphsList = New BetterArray

    WriteTable loName:=SPATIOTEMPORAL, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope

    WriteSpatioTemporalGraph loName:=SPATIOTEMPORAL, graphsList:=graphsList

    Set cellRng = outsh.Cells(2, 3)
    AddGoTo cellRng, dropArray:=sectionsList, goToLabel:=goToSectionLabel, goToPrefix:="spt_"
    AddCommands scope:=scope

    pass.Protect outsh, True
    If IsEmpty(outsh.Range("RNG_HasTable_SPT")) Then outsh.Visible = xlSheetHidden
    TransferCodeWksh scope:=scope
End Sub

'@section Error handling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface implementation
'===============================================================================

Private Sub IAnalysisOutput_WriteAnalysis()
    WriteAnalysis
End Sub
