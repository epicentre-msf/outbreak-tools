VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Graph series specifications for analysis charts"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Graph series specifications for analysis charts")
'@IgnoreModule IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IGraphSpecs

' Table title constants used to validate listobject order
Private Const TAB_ON_TIME_SERIES As String = "time series analysis"
Private Const TAB_ON_GRAPH_TIME_SERIES As String = "graph on time series"
Private Const TAB_ON_TITLE_GRAPHS_TIME_SERIES As String = "labels for time series graphs"

Private Type TGraphSpecs
    specsType As GraphSpecsType
    ' Simple mode: one cross-table
    tabl As ICrossTable
    ' Complex mode: 3 listobjects + output sheet + linelist specs
    loList As BetterArray
    outSh As Worksheet
    lData As ILinelistSpecs
    ' Series data (simple mode)
    seriesNames As BetterArray
    seriesTypes As BetterArray
    seriesPositions As BetterArray
    seriesLabels As BetterArray
    seriesColumnLabels As BetterArray
    seriesLabelPrefixes As BetterArray
    ' Graph collection (complex mode)
    graphsChartSpecs As BetterArray
    graphIdsList As BetterArray
    ' Validity cache
    validityTested As Boolean
    validityResult As Boolean
End Type

Private this As TGraphSpecs
Private Const CLASS_NAME As String = "GraphSpecs"

'@section Factory
'===============================================================================

' @description Create graph specs from a non-time-series cross-table.
'              Works for univariate, bivariate, spatial, and spatio-temporal.
' @param tabl The built cross-table to create graph specs for
Public Function Create(ByVal tabl As ICrossTable) As IGraphSpecs
    If tabl Is Nothing Then
        ThrowError InvalidArgument, "Cross-table is not defined"
    End If

    If tabl.Specifications.TableType = TypeTimeSeries Then
        ThrowError InvalidArgument, _
                   "Simple mode is for non-time-series tables. Use CreateRangeSpecs for time series."
    End If

    If tabl.Specifications.TableType = TypeGlobalSummary Then
        ThrowError InvalidArgument, "Global summary tables do not support graphs"
    End If

    With New GraphSpecs
        Set .Table = tabl
        .GraphType = GraphTypeSimple
        Set Create = .Self
    End With
End Function

' @description Create graph specs for time series from setup listobjects.
' @param loTable BetterArray of 3 ListObjects: [graphLo, timeSeriesLo, titleLo]
' @param outputSh Worksheet where time series graphs will be drawn
' @param lData Linelist specifications
Public Function CreateRangeSpecs(ByVal loTable As BetterArray, _
                                 ByVal outputSh As Worksheet, _
                                 ByVal lData As ILinelistSpecs) As IGraphSpecs

    ValidateListObjects loTable
    If lData Is Nothing Then ThrowError InvalidArgument, "Linelist specifications not defined"
    If outputSh Is Nothing Then ThrowError InvalidArgument, "Output worksheet not defined"

    With New GraphSpecs
        Set .LoList = loTable
        Set .OutputSheet = outputSh
        .GraphType = GraphTypeComplex
        Set .LinelistSpecifications = lData
        Set CreateRangeSpecs = .Self()
    End With
End Function

'@section Internal setters (factory use only)
'===============================================================================

Public Property Get Self() As IGraphSpecs
    Set Self = Me
End Property

Public Property Get Table() As ICrossTable
    Set Table = this.tabl
End Property

Public Property Set Table(ByVal tabl As ICrossTable)
    Set this.tabl = tabl
End Property

Public Property Get GraphType() As GraphSpecsType
    GraphType = this.specsType
End Property

Public Property Let GraphType(ByVal specType As GraphSpecsType)
    this.specsType = specType
End Property

Public Property Get LoList() As BetterArray
    Set LoList = this.loList
End Property

Public Property Set LoList(ByVal loTable As BetterArray)
    Set this.loList = loTable.Clone()
End Property

Public Property Get OutputSheet() As Worksheet
    Set OutputSheet = this.outSh
End Property

Public Property Set OutputSheet(ByVal sh As Worksheet)
    Set this.outSh = sh
End Property

Public Property Get LinelistSpecifications() As ILinelistSpecs
    Set LinelistSpecifications = this.lData
End Property

Public Property Set LinelistSpecifications(ByVal lData As ILinelistSpecs)
    Set this.lData = lData
End Property

'@section Core operations
'===============================================================================

' @description Build all series from the table or listobject specifications.
'              For simple mode, builds series directly from cross-table named ranges.
'              For complex mode, iterates graph IDs and builds nested BetterArray structure.
Public Sub CreateSeries()
    Select Case GraphType

    Case GraphTypeSimple
        BuildSimpleSeries

    Case GraphTypeComplex
        BuildComplexSeries

    End Select
End Sub

'@section Simple mode: series building from cross-table
'===============================================================================

Private Sub BuildSimpleSeries()
    Dim specs As ITableSpecs
    Dim tabId As String
    Dim tabType As Byte
    Dim nbCols As Long
    Dim plotType As String
    Dim flipCoords As String
    Dim tabSecId As String
    Dim counter As Long

    Set specs = Table.Specifications
    tabId = specs.TableId
    tabType = specs.TableType
    nbCols = Table.NumberOfColumns

    ' For spatio-temporal, column count comes from "n geo" setting
    If tabType = TypeSpatioTemporal Then
        nbCols = 5
        On Error Resume Next
        nbCols = CLng(specs.Value("n geo"))
        On Error GoTo 0
        tabSecId = specs.TableSectionId
    End If

    ' Determine chart direction (horizontal or vertical bars)
    flipCoords = specs.Value("flip")
    plotType = IIf(flipCoords = "yes", "hbar", "bar")

    InitSeriesArrays

    Select Case tabType

    Case TypeUnivariate
        BuildUnivariateSeries tabId, plotType, specs.HasPercentage

    Case TypeBivariate, TypeSpatial, TypeSpatioTemporal
        BuildMultiColumnSeries specs, tabId, tabSecId, tabType, nbCols, plotType

    End Select
End Sub

Private Sub BuildUnivariateSeries(ByVal tabId As String, _
                                   ByVal plotType As String, _
                                   ByVal hasPerc As Boolean)
    ' Values series (bars)
    PushSeries "VALUES_COL_1_" & tabId, plotType, "left"
    PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_1_" & tabId

    ' Percentage overlay (points on right axis) — only for vertical bars
    If hasPerc And (plotType = "bar") Then
        PushSeries "PERC_COL_1_" & tabId, "point", "right"
        PushLabels "ROW_CATEGORIES_" & tabId, "PERC_LABEL_COL_" & tabId
    End If
End Sub

Private Sub BuildMultiColumnSeries(ByVal specs As ITableSpecs, _
                                    ByVal tabId As String, _
                                    ByVal tabSecId As String, _
                                    ByVal tabType As Byte, _
                                    ByVal nbCols As Long, _
                                    ByVal plotType As String)
    Dim graphValue As String
    Dim counter As Long
    Dim rowCatName As String

    graphValue = specs.Value("graph")

    Select Case graphValue

    Case "values", "yes"
        For counter = 1 To nbCols
            PushSeries "VALUES_COL_" & counter & "_" & tabId, plotType, "left"
            ' Spatio-temporal uses section-level row categories
            If tabType = TypeSpatioTemporal Then
                rowCatName = "ROW_CATEGORIES_" & tabSecId
            Else
                rowCatName = "ROW_CATEGORIES_" & tabId
            End If
            PushLabels rowCatName, "LABEL_COL_" & counter & "_" & tabId
        Next

    Case "percentages"
        If specs.HasPercentage Then
            For counter = 1 To nbCols
                PushSeries "PERC_COL_" & counter & "_" & tabId, plotType, "left"
                PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_" & counter & "_" & tabId
            Next
        End If

    Case "both"
        For counter = 1 To nbCols
            PushSeries "VALUES_COL_" & counter & "_" & tabId, plotType, "left"
            PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_" & counter & "_" & tabId
            If specs.HasPercentage Then
                PushSeries "PERC_COL_" & counter & "_" & tabId, "point", "right"
                PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_" & counter & "_" & tabId
            End If
        Next

    End Select
End Sub

'@section Complex mode: time series graph building from listobjects
'===============================================================================

Private Sub BuildComplexSeries()
    Dim mGraphIds As BetterArray
    Dim graphId As String
    Dim counter As Long

    Set mGraphIds = BuildGraphIdsList()
    Set this.graphIdsList = mGraphIds.Clone()

    For counter = mGraphIds.LowerBound To mGraphIds.UpperBound
        graphId = mGraphIds.Item(counter)
        DefineGraphSpecs graphId
        PushGraph graphId
    Next
End Sub

' @description Define series for one complex graph ID by reading from setup listobjects
Private Sub DefineGraphSpecs(ByVal graphId As String)
    Dim tsLo As ListObject
    Dim hRng As Range
    Dim idsList As BetterArray
    Dim posList As BetterArray
    Dim percList As BetterArray
    Dim typeList As BetterArray
    Dim choiceList As BetterArray
    Dim labelList As BetterArray
    Dim counter As Long
    Dim val As String
    Dim cellRng As Range
    Dim tabRng As Range
    Dim specs As ITableSpecs
    Dim tabId As String
    Dim secId As String
    Dim choiValue As String
    Dim labName As String
    Dim colName As String

    Set tsLo = TimeSeriesListObject()
    Set hRng = tsLo.HeaderRowRange

    ' Read all column values for this graph ID from the graph listobject
    Set idsList = GraphValues(colName:="series id", graphId:=graphId)
    Set posList = GraphValues(colName:="axis", graphId:=graphId)
    Set percList = GraphValues(colName:="percentages", graphId:=graphId)
    Set typeList = GraphValues(colName:="type", graphId:=graphId)
    Set choiceList = GraphValues(colName:="choices", graphId:=graphId)
    Set labelList = GraphValues(colName:="label", graphId:=graphId)

    ClearSeries

    ' Convert series IDs to table-based named ranges
    For counter = idsList.LowerBound To idsList.UpperBound
        val = Application.WorksheetFunction.Trim(idsList.Item(counter))

        ' Find the series row in the time series listobject
        Set cellRng = Nothing
        Set cellRng = tsLo.ListColumns(1).Range.Find(What:=val, _
                                                      LookAt:=xlWhole, _
                                                      MatchCase:=False)
        ' Skip if not found or in header
        If cellRng Is Nothing Then GoTo ContinueLoop
        If cellRng.Row <= hRng.Row Then GoTo ContinueLoop

        Set tabRng = tsLo.ListRows(cellRng.Row - hRng.Row).Range
        Set specs = TableSpecs.Create(hRng, tabRng, LinelistSpecifications)
        tabId = specs.TableId
        secId = specs.TableSectionId
        choiValue = choiceList.Item(counter)

        ' Resolve the column name from the choice value
        labName = TimeSeriesColumnName(choiValue, tabId)
        colName = ResolveSeriesRangeName(choiValue, tabId, labName)

        ' Apply percentage override if requested
        If percList.Item(counter) = "percentages" Then
            If choiValue = "Total" Then
                colName = Replace(labName, "COL", "PERC")
            Else
                colName = Replace(labName, "LABEL", "PERC")
            End If
        End If

        ' Push the series if we resolved a valid range name
        If colName <> vbNullString Then
            PushSeries colName, typeList.Item(counter), posList.Item(counter)
            PushLabels "ROW_CATEGORIES_" & secId, labName, labelList.Item(counter)
        End If

ContinueLoop:
    Next
End Sub

' @description Resolve the data range name from a choice value and table ID
Private Function ResolveSeriesRangeName(ByVal choiValue As String, _
                                         ByVal tabId As String, _
                                         ByVal labName As String) As String
    If choiValue = "Total" Then
        ResolveSeriesRangeName = "TOTAL_COL_VALUES_" & tabId
    ElseIf InStr(1, labName, "COLUMN_CATEGORIES_") > 0 Then
        ResolveSeriesRangeName = "INTERIOR_VALUES_" & tabId
    Else
        ResolveSeriesRangeName = Replace(labName, "LABEL", "VALUES")
    End If
End Function

' @description Find the named range for a choice value in a time series table
Private Function TimeSeriesColumnName(ByVal choiValue As String, _
                                       ByVal tabId As String) As String
    Dim rng As Range
    Dim cellRng As Range
    Dim colName As String
    Dim sh As Worksheet

    Set sh = Wksh()

    If choiValue = vbNullString Then Exit Function

    If choiValue = "Total" Then
        TimeSeriesColumnName = "TOTAL_LABEL_COL_" & tabId
        Exit Function
    End If

    Set rng = sh.Range("COLUMN_CATEGORIES_" & tabId)
    Set cellRng = rng.Find(What:=choiValue, LookAt:=xlWhole, MatchCase:=True)

    On Error Resume Next
    If Not cellRng Is Nothing Then colName = cellRng.Name.Name
    On Error GoTo 0

    TimeSeriesColumnName = colName
End Function

'@section Series data management
'===============================================================================

Private Sub InitSeriesArrays()
    Set this.seriesNames = New BetterArray
    this.seriesNames.LowerBound = 1
    Set this.seriesTypes = New BetterArray
    this.seriesTypes.LowerBound = 1
    Set this.seriesPositions = New BetterArray
    this.seriesPositions.LowerBound = 1
    Set this.seriesLabels = New BetterArray
    this.seriesLabels.LowerBound = 1
    Set this.seriesColumnLabels = New BetterArray
    this.seriesColumnLabels.LowerBound = 1
    Set this.seriesLabelPrefixes = New BetterArray
    this.seriesLabelPrefixes.LowerBound = 1
End Sub

Private Sub PushSeries(ByVal chrtName As String, _
                        ByVal chrtType As String, _
                        ByVal chrtPos As String)
    ' Lazy init for complex mode (called per graph ID after ClearSeries)
    If this.seriesNames Is Nothing Then InitSeriesArrays

    this.seriesNames.Push chrtName
    this.seriesTypes.Push chrtType
    this.seriesPositions.Push chrtPos
End Sub

Private Sub PushLabels(ByVal chrtRowCategory As String, _
                        ByVal chrtColumnLab As String, _
                        Optional ByVal chrtLabPref As String = vbNullString)
    If this.seriesLabels Is Nothing Then InitSeriesArrays

    this.seriesLabels.Push chrtRowCategory
    this.seriesColumnLabels.Push chrtColumnLab
    this.seriesLabelPrefixes.Push chrtLabPref
End Sub

Private Sub ClearSeries()
    On Error Resume Next
    If Not this.seriesNames Is Nothing Then this.seriesNames.Clear
    If Not this.seriesTypes Is Nothing Then this.seriesTypes.Clear
    If Not this.seriesPositions Is Nothing Then this.seriesPositions.Clear
    If Not this.seriesLabels Is Nothing Then this.seriesLabels.Clear
    If Not this.seriesColumnLabels Is Nothing Then this.seriesColumnLabels.Clear
    If Not this.seriesLabelPrefixes Is Nothing Then this.seriesLabelPrefixes.Clear
    On Error GoTo 0
End Sub

' @description Package current series into a graph spec entry and push to collection.
'              Each graph entry is a BetterArray of 7 cloned BetterArrays:
'              [seriesNames, seriesTypes, seriesPositions, labels, catLabels, prefixes, title]
Private Sub PushGraph(ByVal graphId As String)
    Dim graphSpecsTable As BetterArray
    Dim tabSeries As BetterArray
    Dim tabType As BetterArray
    Dim tabPos As BetterArray
    Dim tabLabel As BetterArray
    Dim tabCatLabel As BetterArray
    Dim tabPrefLabel As BetterArray
    Dim title As String

    Set graphSpecsTable = New BetterArray
    graphSpecsTable.LowerBound = 1

    Set tabSeries = this.seriesNames.Clone()
    Set tabType = this.seriesTypes.Clone()
    Set tabPos = this.seriesPositions.Clone()
    Set tabLabel = this.seriesLabels.Clone()
    Set tabCatLabel = this.seriesColumnLabels.Clone()
    Set tabPrefLabel = this.seriesLabelPrefixes.Clone()

    ' Safety check: BetterArray.Clone() can sometimes add an extra element
    If tabSeries.Length = this.seriesNames.Length + 1 Then
        tabSeries.Pop
        tabType.Pop
        tabPos.Pop
        tabLabel.Pop
        tabCatLabel.Pop
        tabPrefLabel.Pop
    End If

    title = GraphTitle(graphId)

    graphSpecsTable.Push tabSeries, tabType, tabPos, _
                         tabLabel, tabCatLabel, tabPrefLabel, title

    If this.graphsChartSpecs Is Nothing Then
        Set this.graphsChartSpecs = New BetterArray
        this.graphsChartSpecs.LowerBound = 1
    End If

    this.graphsChartSpecs.Push graphSpecsTable
End Sub

'@section ListObject helpers (complex mode)
'===============================================================================

Private Property Get GraphListObject() As ListObject
    Set GraphListObject = LoList.Item(LoList.LowerBound)
End Property

Private Property Get TimeSeriesListObject() As ListObject
    Set TimeSeriesListObject = LoList.Item(LoList.LowerBound + 1)
End Property

Private Property Get TitlesListObject() As ListObject
    Set TitlesListObject = LoList.Item(LoList.LowerBound + 2)
End Property

' @description Look up graph title from the titles listobject
Private Function GraphTitle(ByVal graphId As String) As String
    Dim titleLo As ListObject
    Dim idRng As Range
    Dim cellRng As Range

    Set titleLo = TitlesListObject()
    Set idRng = titleLo.ListColumns(3).DataBodyRange

    Set cellRng = idRng.Find(What:=graphId, LookAt:=xlWhole, MatchCase:=True)
    If Not cellRng Is Nothing Then
        GraphTitle = cellRng.Offset(, -2).Value
    End If
End Function

' @description Build the list of unique graph IDs from the graph listobject
Private Function BuildGraphIdsList() As BetterArray
    Dim idsList As BetterArray
    Dim graphLo As ListObject
    Dim idIndex As Long
    Dim rng As Range
    Dim counter As Long

    Set idsList = New BetterArray
    idsList.LowerBound = 1
    Set graphLo = GraphListObject()
    idIndex = LoColumnIndex(graphLo, "graph id")
    Set rng = graphLo.ListColumns(idIndex).Range

    ' Start at 2 to skip header row
    For counter = 2 To rng.Rows.Count
        If Not IsEmpty(rng.Cells(counter, 1)) Then
            If Not idsList.Includes(rng.Cells(counter, 1).Value) Then
                idsList.Push rng.Cells(counter, 1).Value
            End If
        End If
    Next

    Set BuildGraphIdsList = idsList.Clone()
End Function

' @description Get all values for a column filtered by graph ID
Private Function GraphValues(ByVal colName As String, _
                              ByVal graphId As String) As BetterArray
    Dim grLo As ListObject
    Dim valuesList As BetterArray
    Dim idIndex As Long
    Dim colIndex As Long
    Dim rng As Range
    Dim colRng As Range
    Dim rowIndex As Long
    Dim cellVal As String
    Dim sh As Worksheet

    Set grLo = GraphListObject()
    Set valuesList = New BetterArray
    valuesList.LowerBound = 1

    idIndex = LoColumnIndex(grLo, "graph id")
    colIndex = LoColumnIndex(grLo, colName)

    If idIndex < 0 Then ThrowError ElementNotFound, "graph id column not found"
    If colIndex < 0 Then ThrowError ElementNotFound, colName & " column not found in graph listobject"

    Set rng = grLo.ListColumns(idIndex).DataBodyRange
    Set colRng = grLo.ListColumns(colIndex).DataBodyRange
    Set sh = rng.Worksheet

    If rng.Find(What:=graphId, LookAt:=xlWhole, MatchCase:=True) Is Nothing Then
        Set GraphValues = valuesList
        Exit Function
    End If

    ' Walk through all matching rows for this graph ID
    With sh
        Do
            Set rng = rng
            Dim foundCell As Range
            Set foundCell = rng.Find(What:=graphId, LookAt:=xlWhole, _
                                     MatchCase:=True, _
                                     After:=rng.Cells(rng.Rows.Count, 1))
            If foundCell Is Nothing Then Exit Do

            rowIndex = foundCell.Row - rng.Row + 1
            cellVal = colRng.Cells(rowIndex, 1).Value

            If cellVal = vbNullString Then
                valuesList.Push "&"
            Else
                valuesList.Push cellVal
            End If

            ' Single-row range: done
            If rng.Rows.Count = 1 Then Exit Do
            If rowIndex + 1 > rng.Rows.Count Then Exit Do

            ' Narrow the search range to avoid re-finding the same row
            Set rng = .Range(rng.Cells(rowIndex + 1, 1), rng.Cells(rng.Rows.Count, 1))
            Set colRng = .Range(colRng.Cells(rowIndex + 1, 1), colRng.Cells(colRng.Rows.Count, 1))
        Loop
    End With

    Set GraphValues = valuesList.Clone()
End Function

' @description Column index lookup for a listobject header (1-based, relative)
Private Function LoColumnIndex(ByVal lo As ListObject, _
                                ByVal colName As String) As Long
    Dim hRng As Range
    Dim foundCell As Range

    Set hRng = lo.HeaderRowRange
    Set foundCell = hRng.Find(What:=colName, LookAt:=xlPart, MatchCase:=False)

    If foundCell Is Nothing Then
        LoColumnIndex = -1
    Else
        LoColumnIndex = foundCell.Column - hRng.Column + 1
    End If
End Function

'@section Validation
'===============================================================================

' @description Validate the 3 listobjects for complex graph specs
Private Sub ValidateListObjects(ByVal loTable As BetterArray)
    Dim counter As Long
    Dim lo As ListObject
    Dim testObj As Variant
    Dim rng As Range
    Dim titleVal As String

    If loTable Is Nothing Then
        ThrowError InvalidArgument, "ListObject collection is empty"
    End If
    If loTable.Length <> 3 Then
        ThrowError InvalidArgument, _
                   "Expected 3 ListObjects (graph, time series, titles), got " & loTable.Length
    End If

    ' Validate each entry is a non-empty ListObject
    For counter = loTable.LowerBound To loTable.UpperBound
        Set testObj = loTable.Item(counter)
        If Not (TypeName(testObj) = "ListObject") Then
            ThrowError InvalidArgument, "Item " & counter & " is not a ListObject"
        End If
        Set lo = loTable.Item(counter)
        If lo Is Nothing Then ThrowError InvalidArgument, "ListObject at index " & counter & " is Nothing"
        If lo.DataBodyRange Is Nothing Then
            ThrowError InvalidArgument, "ListObject " & lo.Name & " has no data"
        End If
    Next

    ' Validate listobject order by title labels
    Set rng = loTable.Item(loTable.LowerBound).HeaderRowRange
    titleVal = LCase(Application.WorksheetFunction.Trim(rng.Cells(-1, 1).Value))
    If titleVal <> TAB_ON_GRAPH_TIME_SERIES Then
        ThrowError ErrorUnexpectedState, _
                   "First ListObject must be graph table. Found: " & titleVal
    End If

    Set rng = loTable.Item(loTable.LowerBound + 1).HeaderRowRange
    titleVal = LCase(Application.WorksheetFunction.Trim(rng.Cells(-1, 1).Value))
    If titleVal <> TAB_ON_TIME_SERIES Then
        ThrowError ErrorUnexpectedState, _
                   "Second ListObject must be time series table. Found: " & titleVal
    End If

    Set rng = loTable.Item(loTable.LowerBound + 2).HeaderRowRange
    titleVal = LCase(Application.WorksheetFunction.Trim(rng.Cells(-1, 1).Value))
    If titleVal <> TAB_ON_TITLE_GRAPHS_TIME_SERIES Then
        ThrowError ErrorUnexpectedState, _
                   "Third ListObject must be title table. Found: " & titleVal
    End If
End Sub

'@section Properties
'===============================================================================

Public Property Get Wksh() As Worksheet
    Select Case GraphType
    Case GraphTypeSimple
        Set Wksh = Table.Wksh
    Case GraphTypeComplex
        Set Wksh = OutputSheet
    End Select
End Property

Public Property Get NumberOfSeries() As Long
    If (this.seriesNames Is Nothing) Or (GraphType = GraphTypeComplex) Then
        NumberOfSeries = 0
        Exit Property
    End If
    NumberOfSeries = this.seriesNames.Length
End Property

Public Property Get NumberOfGraphs() As Long
    If (this.graphsChartSpecs Is Nothing) Or (GraphType <> GraphTypeComplex) Then
        NumberOfGraphs = 0
        Exit Property
    End If
    NumberOfGraphs = this.graphsChartSpecs.Length
End Property

Public Property Get SeriesName(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesName = this.seriesNames.Item(index)
End Property

Public Property Get SeriesType(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesType = this.seriesTypes.Item(index)
End Property

Public Property Get SeriesPos(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesPos = this.seriesPositions.Item(index)
End Property

Public Property Get SeriesLabel(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesLabel = this.seriesLabels.Item(index)
End Property

Public Property Get SeriesColumnLabel(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesColumnLabel = this.seriesColumnLabels.Item(index)
End Property

Public Property Get SeriesLabelPrefix(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesLabelPrefix = this.seriesLabelPrefixes.Item(index)
End Property

Public Property Get Valid() As Boolean
    If this.validityTested Then
        Valid = this.validityResult
        Exit Property
    End If

    CreateSeries

    If GraphType = GraphTypeSimple Then
        this.validityResult = True
    Else
        this.validityResult = Not (this.graphsChartSpecs Is Nothing)
    End If

    this.validityTested = True
    Valid = this.validityResult
End Property

Public Property Get SpecsLists() As BetterArray
    If this.graphsChartSpecs Is Nothing Then
        Set SpecsLists = New BetterArray
    Else
        Set SpecsLists = this.graphsChartSpecs.Clone()
    End If
End Property

Public Property Get GraphIDs() As BetterArray
    If this.graphIdsList Is Nothing Then
        Set this.graphIdsList = BuildGraphIdsList()
    End If
    Set GraphIDs = this.graphIdsList.Clone()
End Property

'@section Error handling
'===============================================================================

Private Sub ValidateSeriesIndex(ByVal index As Long)
    If (index > NumberOfSeries) Or (index < 1) Then
        ThrowError InvalidArgument, _
                   "Series index " & index & " out of bounds (1 to " & NumberOfSeries & ")"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface implementation
'===============================================================================

Private Property Get IGraphSpecs_NumberOfSeries() As Long
    IGraphSpecs_NumberOfSeries = NumberOfSeries()
End Property

Private Property Get IGraphSpecs_NumberOfGraphs() As Long
    IGraphSpecs_NumberOfGraphs = NumberOfGraphs()
End Property

Private Property Get IGraphSpecs_SeriesName(ByVal index As Long) As String
    IGraphSpecs_SeriesName = SeriesName(index)
End Property

Private Property Get IGraphSpecs_SeriesType(ByVal index As Long) As String
    IGraphSpecs_SeriesType = SeriesType(index)
End Property

Private Property Get IGraphSpecs_SeriesPos(ByVal index As Long) As String
    IGraphSpecs_SeriesPos = SeriesPos(index)
End Property

Private Property Get IGraphSpecs_SeriesLabel(ByVal index As Long) As String
    IGraphSpecs_SeriesLabel = SeriesLabel(index)
End Property

Private Property Get IGraphSpecs_SeriesColumnLabel(ByVal index As Long) As String
    IGraphSpecs_SeriesColumnLabel = SeriesColumnLabel(index)
End Property

Private Property Get IGraphSpecs_SeriesLabelPrefix(ByVal index As Long) As String
    IGraphSpecs_SeriesLabelPrefix = SeriesLabelPrefix(index)
End Property

Private Sub IGraphSpecs_CreateSeries()
    CreateSeries
End Sub

Private Property Get IGraphSpecs_SpecsLists() As BetterArray
    Set IGraphSpecs_SpecsLists = SpecsLists()
End Property

Private Property Get IGraphSpecs_Valid() As Boolean
    IGraphSpecs_Valid = Valid()
End Property

Private Property Get IGraphSpecs_Wksh() As Worksheet
    Set IGraphSpecs_Wksh = Wksh()
End Property

Private Property Get IGraphSpecs_GraphIDs() As BetterArray
    Set IGraphSpecs_GraphIDs = GraphIDs()
End Property
