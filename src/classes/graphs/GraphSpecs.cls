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

'@section Enumerations
'===============================================================================
'@description
'Graph specification modes. Re-homed from the removed IGraphSpecs interface;
'depending on GraphSpecs resolves these everywhere.
Public Enum GraphSpecsType
    GraphTypeSimple = 1                          'Non-time-series (univariate, bivariate, spatial, spatio-temporal)
    GraphTypeComplex = 2                         'Time series (multiple graphs from setup listobjects)
End Enum


'@class GraphSpecs
'@description
'Builds and stores chart series specifications for both simple (non-time-series)
'and complex (time series) analysis graphs. In simple mode, the class reads chart
'configuration directly from an CrossTable and produces flat parallel arrays of
'series metadata (names, types, positions, labels). In complex mode, it reads from
'three setup ListObjects to produce a nested BetterArray collection of graph entries,
'each containing its own set of parallel series arrays plus a display title. Consumers
'create an instance via Create (simple) or CreateRangeSpecs (complex), call
'CreateSeries (or let Valid trigger it lazily), then read series data via the
'public accessors.
'
'EVERY SKIP IS REPORTED
'-------------------------------------------------------------------------------
'The class decides to build fewer series than the setup asked for in nine
'places: a mode it has no builder for, a graph setting it does not recognise, a
'percentage chart on a table with no percentage column, a series id absent from
'the time series table, a choice value with no column, a column with no named
'range, a graph id with no title row, a graph that ends with no series, and a
'setup table with a column missing. Each one files an entry through LogInfo, and
'AnalysisOutput reads them into the generation report. A user whose chart came
'out with two series instead of five can now see why.
'
'THE THREE SETUP TABLES ARE IDENTIFIED BY NAME
'-------------------------------------------------------------------------------
'The order check used to read a caption two rows above each header row. The gap
'between the caption and the header was measured in six workbooks on 2026-07-31:
'it is four rows in .mock/setup_mock.xlsb, src/bin/setup/setup.xlsb and
'setup_dev.xlsb, and two rows in every file under releases/. A row offset is a
'contract with a file nobody versions, which is the lesson TableSpecs paid for
'when it moved the analysis scope onto the ListObject name. The three names are
'byte-identical in all six files, so this class compares names too.
'@depends CrossTable, TableSpecs, LinelistSpecs, Checking, BetterArray
'@version 2026-07-31



' The three setup tables this class reads, in the order CreateRangeSpecs
' expects them. The names are the ones the setup workbook gives the tables and
' they are identical across the development files and every release.
Private Const TABLE_GRAPH_TIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TABLE_TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const TABLE_LABEL_TSGRAPH As String = "Tab_Label_TSGraph"
Private Const EXPECTED_LISTOBJECT_COUNT As Long = 3

' Column search terms. Every one is matched against the header row without
' regard to case: an exact header wins, and a header that merely contains the
' term is the fallback. "axis" finds "Y-Axis", "type" finds "Chart type" and
' "percentages" finds "Plot values or percentages".
Private Const COLUMN_GRAPH_ID As String = "graph id"
Private Const COLUMN_SERIES_ID As String = "series id"
Private Const COLUMN_AXIS As String = "axis"
Private Const COLUMN_PERCENTAGES As String = "percentages"
Private Const COLUMN_TYPE As String = "type"
Private Const COLUMN_CHOICES As String = "choices"
Private Const COLUMN_LABEL As String = "label"
Private Const COLUMN_GRAPH_TITLE As String = "graph title"

' The graph setting values the simple builder knows, and the choice value that
' asks for the total column rather than a category column.
Private Const MODE_VALUES As String = "values"
Private Const MODE_YES As String = "yes"
Private Const MODE_PERCENTAGES As String = "percentages"
Private Const MODE_BOTH As String = "both"
Private Const CHOICE_TOTAL As String = "total"

Private Type TGraphSpecs
    specsType As GraphSpecsType
    ' Simple mode: one cross-table
    tabl As CrossTable
    ' Complex mode: 3 listobjects + output sheet + linelist specs
    loList As BetterArray
    graphLo As ListObject
    timeSeriesLo As ListObject
    titlesLo As ListObject
    outSh As Worksheet
    lData As LinelistSpecs
    ' Series data (simple mode)
    seriesNames As BetterArray
    seriesTypes As BetterArray
    seriesPositions As BetterArray
    seriesLabels As BetterArray
    seriesColumnLabels As BetterArray
    seriesLabelPrefixes As BetterArray
    ' Graph collection (complex mode)
    graphsChartSpecs As BetterArray
    ' Validity cache
    validityTested As Boolean
    validityResult As Boolean
    hasChecks As Boolean
    sealed As Boolean
End Type

Private this As TGraphSpecs
Private checkCounter As Long
Private internalChecks As Checking
Private Const CLASS_NAME As String = "GraphSpecs"


'@section Instantiation
'===============================================================================
'@description Factory methods and Self accessor for the predeclared instance pattern.

'@label:create
'@sub-title Create a GraphSpecs instance in simple mode from a cross-table
'@details
'Factory method that creates a GraphSpecs object in simple mode for non-time-series
'cross-tables. Simple mode is used for univariate, bivariate, spatial, and
'spatio-temporal analysis tables. This method explicitly rejects TimeSeries tables
'(which require CreateRangeSpecs) and GlobalSummary tables (which do not support
'graphing). The returned GraphSpecs instance stores a reference to the cross-table
'and can later produce chart series definitions via CreateSeries.
'@param tabl CrossTable. The fully-built cross-table whose named ranges contain
'the data to graph. Must not be Nothing, ScopeTimeSeries, or ScopeGlobalSummary.
'@return GraphSpecs. An instance configured in simple mode (GraphTypeSimple)
'ready for series creation via CreateSeries.
'@throws InvalidArgument When tabl is Nothing, ScopeTimeSeries, or ScopeGlobalSummary.
'@remarks After calling Create, you must call CreateSeries before accessing any
'series properties (SeriesName, SeriesType, etc.). The cross-table must already
'have its named ranges written to the worksheet before series data can be consumed.
'@export
Public Function Create(ByVal tabl As CrossTable) As GraphSpecs
    Dim tableScope As Byte
    Dim newSpecs As GraphSpecs

    If tabl Is Nothing Then
        ThrowError InvalidArgument, "Cross-table is not defined"
    End If

    ' One read of the specification chain. It used to be walked twice.
    tableScope = tabl.Specifications.TableScope

    If tableScope = ScopeTimeSeries Then
        ThrowError InvalidArgument, _
                   "Simple mode is for non-time-series tables. Use CreateRangeSpecs for time series."
    End If

    If tableScope = ScopeGlobalSummary Then
        ThrowError InvalidArgument, "Global summary tables do not support graphs"
    End If

    Set newSpecs = New GraphSpecs
    Set newSpecs.Table = tabl
    newSpecs.GraphType = GraphTypeSimple
    newSpecs.Seal
    Set Create = newSpecs
End Function

'@label:create-range-specs
'@sub-title Create a GraphSpecs instance in complex mode for time series graphs
'@details
'Factory method that creates a GraphSpecs object in complex mode for time series
'graphs. Complex mode requires exactly 3 ListObjects passed as a BetterArray:
'(1) the graph setup table defining series per graph ID with types, axis positions,
'and choice values; (2) the time series setup table mapping series IDs to table
'specifications; (3) the title labels table providing display titles per graph ID.
'All three ListObjects are validated via ValidateListObjects before construction.
'Complex mode produces a nested BetterArray collection of graph spec entries (one
'per unique graph ID) when CreateSeries is called later.
'@param loTable BetterArray. Exactly 3 ListObject entries in order:
'[Tab_Graph_TimeSeries, Tab_TimeSeries_Analysis, Tab_Label_TSGraph]. Each must
'carry at least one data row.
'@param outputSh Worksheet. The worksheet where time series graphs will be drawn.
'@param lData LinelistSpecs. Linelist-level metadata (section names, sheet info)
'needed to resolve table specifications during series building.
'@return GraphSpecs. An instance configured in complex mode (GraphTypeComplex)
'ready for series creation via CreateSeries.
'@throws InvalidArgument When loTable fails validation, lData is Nothing, or
'outputSh is Nothing.
'@throws ErrorUnexpectedState When the ListObjects are in the wrong order.
'@remarks The 3 ListObjects must appear in the order ValidateListObjects checks:
'the graph table first, the time series table second, the title labels table
'third. Each one is recognised by its name.
'@export
Public Function CreateRangeSpecs(ByVal loTable As BetterArray, _
                                 ByVal outputSh As Worksheet, _
                                 ByVal lData As LinelistSpecs) As GraphSpecs

    ValidateListObjects loTable
    If lData Is Nothing Then ThrowError InvalidArgument, "Linelist specifications not defined"
    If outputSh Is Nothing Then ThrowError InvalidArgument, "Output worksheet not defined"

    Dim newSpecs As GraphSpecs
    Set newSpecs = New GraphSpecs
    Set newSpecs.LoList = loTable
    Set newSpecs.OutputSheet = outputSh
    newSpecs.GraphType = GraphTypeComplex
    Set newSpecs.LinelistSpecifications = lData
    newSpecs.Seal
    Set CreateRangeSpecs = newSpecs
End Function

'@label:seal
'@sub-title Seal the instance against further setup writes
'@details
'Called by the factory after the last setup write. Once sealed, the
'creation-only setters raise through ThrowError.
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against post-creation writes
'@param propName String. The property name reported in the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub


'@section InternalSetters
'===============================================================================
'@description Property accessors used only during factory construction.
'These are Public to support the With New pattern but are not part of the
'caller contract.

'@label:table
'@prop-title Cross-table reference for simple mode
'@return CrossTable. The stored cross-table instance.
Public Property Get Table() As CrossTable
    Set Table = this.tabl
End Property

'@label:table-set
'@prop-title Assign the cross-table reference
'@param tabl CrossTable. The cross-table to store.
Public Property Set Table(ByVal tabl As CrossTable)
    GuardNotSealed "Table"
    Set this.tabl = tabl
End Property

'@label:graph-type
'@prop-title Current graph specifications mode
'@return GraphSpecsType. GraphTypeSimple or GraphTypeComplex.
Public Property Get GraphType() As GraphSpecsType
    GraphType = this.specsType
End Property

'@label:graph-type-set
'@prop-title Assign the graph specifications mode
'@param specType GraphSpecsType. The mode to set.
Public Property Let GraphType(ByVal specType As GraphSpecsType)
    GuardNotSealed "GraphType"
    this.specsType = specType
End Property

'@label:lo-list
'@prop-title ListObject collection for complex mode
'@return BetterArray. The stored ListObject array.
Public Property Get LoList() As BetterArray
    Set LoList = this.loList
End Property

'@label:lo-list-set
'@prop-title Assign the ListObject collection
'@details
'Stores the array and resolves the three tables into fields of their own. The
'three accessors are read inside loops, and each read used to go through the
'array again.
'@param loTable BetterArray. The 3-element ListObject array to clone and store.
Public Property Set LoList(ByVal loTable As BetterArray)
    GuardNotSealed "LoList"
    Set this.loList = loTable.Clone()
    Set this.graphLo = loTable.Item(loTable.LowerBound)
    Set this.timeSeriesLo = loTable.Item(loTable.LowerBound + 1)
    Set this.titlesLo = loTable.Item(loTable.LowerBound + 2)
End Property

'@label:output-sheet
'@prop-title Output worksheet for complex mode
'@return Worksheet. The stored output worksheet reference.
Public Property Get OutputSheet() As Worksheet
    Set OutputSheet = this.outSh
End Property

'@label:output-sheet-set
'@prop-title Assign the output worksheet
'@param sh Worksheet. The worksheet to store.
Public Property Set OutputSheet(ByVal sh As Worksheet)
    GuardNotSealed "OutputSheet"
    Set this.outSh = sh
End Property

'@label:linelist-specifications
'@prop-title Linelist specifications for complex mode
'@return LinelistSpecs. The stored linelist specifications instance.
Public Property Get LinelistSpecifications() As LinelistSpecs
    Set LinelistSpecifications = this.lData
End Property

'@label:linelist-specifications-set
'@prop-title Assign the linelist specifications
'@param lData LinelistSpecs. The linelist specifications to store.
Public Property Set LinelistSpecifications(ByVal lData As LinelistSpecs)
    GuardNotSealed "LinelistSpecifications"
    Set this.lData = lData
End Property


'@section CoreOperations
'===============================================================================
'@description Entry point for building all chart series definitions.

'@label:create-series
'@sub-title Build all series from the stored specifications
'@details
'Dispatches series building to the appropriate handler based on the current
'GraphType. For simple mode (GraphTypeSimple), delegates to BuildSimpleSeries
'which reads chart configuration directly from the cross-table's named ranges.
'For complex mode (GraphTypeComplex), delegates to BuildComplexSeries which
'iterates over unique graph IDs from the setup ListObjects and builds a nested
'BetterArray collection of graph spec entries. Must be called explicitly before
'accessing any series data properties or the SpecsLists collection.
'@remarks Also called internally by the Valid property if the validity cache has
'not yet been populated. Calling it multiple times is safe but will rebuild all
'series data from scratch each time.
'@export
Public Sub CreateSeries()
    Select Case GraphType

    Case GraphTypeSimple
        BuildSimpleSeries

    Case GraphTypeComplex
        BuildComplexSeries

    Case Else
        LogInfo "The graph specifications carry no mode, so no chart series were built.", _
                checkingError

    End Select
End Sub


'@section SimpleMode
'===============================================================================
'@description Series building logic for non-time-series cross-tables.

'@label:build-simple-series
'@sub-title Build chart series for a non-time-series cross-table
'@details
'Reads the table scope and configuration from the cross-table's TableSpecs to
'determine what kind of series to generate. FlipCoordinates controls the plot
'direction: horizontal bars ("hbar") when it answers True, vertical bars ("bar")
'otherwise. For spatio-temporal tables the column count is the table's
'geographic unit count, and the section-level table ID (tabSecId) is captured
'for use as row category ranges. After initializing fresh series arrays,
'dispatches to BuildUnivariateSeries for ScopeUnivariate tables, or
'BuildMultiColumnSeries for ScopeBivariate, ScopeSpatial, and
'ScopeSpatioTemporal tables.
'@remarks TableSpecs.GeoCount carries the default of 5 and clamps the value to
'1..20, so an empty or unreadable cell needs no guard here.
Private Sub BuildSimpleSeries()
    Dim specs As TableSpecs
    Dim tabId As String
    Dim tabType As Byte
    Dim nbCols As Long
    Dim plotType As String
    Dim tabSecId As String

    Set specs = Table.Specifications
    tabId = specs.TableId
    tabType = specs.TableScope
    nbCols = Table.NumberOfColumns

    ' For spatio-temporal, the column count is the geographic unit count.
    ' GeoCount carries the default of 5 and clamps to 1..20, so an empty cell
    ' no longer raises out of CLng.
    If tabType = ScopeSpatioTemporal Then
        nbCols = specs.GeoCount()
        tabSecId = specs.TableSectionId
    End If

    ' The flag is read through TableSpecs, which trims it and folds its case.
    ' Comparing the raw cell here made a cell holding " Yes " open the gate in
    ' AnalysisOutput and match nothing in this class.
    plotType = IIf(specs.FlipCoordinates(), "hbar", "bar")

    InitSeriesArrays

    Select Case tabType

    Case ScopeUnivariate
        BuildUnivariateSeries tabId, plotType, specs.HasPercentage

    Case ScopeBivariate, ScopeSpatial, ScopeSpatioTemporal
        BuildMultiColumnSeries specs, tabId, tabSecId, tabType, nbCols, plotType

    Case Else
        LogInfo "Table " & tabId & " has the scope " & tabType & ", which has no " & _
                "series builder. No chart series were built for it.", checkingError

    End Select

    If BufferCount() = 0 Then
        LogInfo "Table " & tabId & " asks for a chart and no series could be built " & _
                "for it, so the chart is empty.", checkingWarning
    End If
End Sub

'@label:build-univariate-series
'@sub-title Build chart series for a single-variable cross-table
'@details
'Pushes one bar series referencing the VALUES_COL_1_ named range for the primary
'data, paired with the ROW_CATEGORIES_ range for axis labels and LABEL_COL_1_
'for the legend entry. If the table has a percentage column and the plot direction
'is vertical bars (plotType = "bar"), an additional point series is pushed
'referencing PERC_COL_1_ on the right axis, allowing a percentage overlay on top
'of the value bars. Horizontal bars ("hbar") do not get the percentage overlay
'because point series on a horizontal chart do not render well.
'@param tabId String. The unique table identifier used as a suffix for all named
'range lookups.
'@param plotType String. Chart series type: "bar" for vertical or "hbar" for
'horizontal bars.
'@param hasPerc Boolean. Whether the table includes a percentage column. When
'True and plotType is "bar", an additional percentage point series is pushed.
'@remarks The percentage overlay is intentionally skipped for horizontal bars
'because Excel point series render poorly on horizontally-oriented charts.
Private Sub BuildUnivariateSeries(ByVal tabId As String, _
                                   ByVal plotType As String, _
                                   ByVal hasPerc As Boolean)
    ' Values series (bars)
    PushSeries "VALUES_COL_1_" & tabId, plotType, "left"
    PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_1_" & tabId

    ' Percentage overlay (points on right axis) for vertical bars
    If hasPerc And (plotType = "bar") Then
        PushSeries "PERC_COL_1_" & tabId, "point", "right"
        PushLabels "ROW_CATEGORIES_" & tabId, "PERC_LABEL_COL_" & tabId
    End If
End Sub

'@label:build-multi-column-series
'@sub-title Build chart series for multi-column cross-tables
'@details
'Builds series for bivariate, spatial, or spatio-temporal tables. The graph mode
'comes from TableSpecs.GraphMode: (1) "values" or "yes" pushes one bar series
'per column referencing VALUES_COL_<n>_ named ranges on the left axis;
'(2) "percentages" pushes one bar series per column referencing PERC_COL_<n>_
'ranges, and reports when the table carries no percentage column; (3) "both"
'pushes value bars on the left axis AND percentage points on the right axis for
'each column, creating a dual-axis overlay chart. For spatio-temporal tables the
'row categories use the section-level range (ROW_CATEGORIES_ + tabSecId), because
'spatio-temporal tables share row categories across all tables in the same
'section.
'@param specs TableSpecs. The table specifications providing the graph mode and
'HasPercentage.
'@param tabId String. The unique table identifier for named range construction.
'@param tabSecId String. The section-level table identifier, used only for
'spatio-temporal tables. Empty for other types.
'@param tabType Byte. The numeric table scope (ScopeBivariate, ScopeSpatial, or
'ScopeSpatioTemporal).
'@param nbCols Long. The number of data columns to iterate over.
'@param plotType String. Chart series type: "bar" or "hbar".
'@remarks A graph mode outside the three the setup dropdowns offer files an
'error entry. It used to fall through the Select Case with no arm, which left
'the chart with no series and raised 1004 in Graphs.Format two frames later.
Private Sub BuildMultiColumnSeries(ByVal specs As TableSpecs, _
                                    ByVal tabId As String, _
                                    ByVal tabSecId As String, _
                                    ByVal tabType As Byte, _
                                    ByVal nbCols As Long, _
                                    ByVal plotType As String)
    Dim graphValue As String
    Dim counter As Long
    Dim rowCatName As String

    graphValue = specs.GraphMode()

    ' Spatio-temporal tables share their row categories across the section.
    If tabType = ScopeSpatioTemporal Then
        rowCatName = "ROW_CATEGORIES_" & tabSecId
    Else
        rowCatName = "ROW_CATEGORIES_" & tabId
    End If

    Select Case graphValue

    Case MODE_VALUES, MODE_YES
        For counter = 1 To nbCols
            PushSeries "VALUES_COL_" & counter & "_" & tabId, plotType, "left"
            PushLabels rowCatName, "LABEL_COL_" & counter & "_" & tabId
        Next

    Case MODE_PERCENTAGES
        If specs.HasPercentage Then
            For counter = 1 To nbCols
                PushSeries "PERC_COL_" & counter & "_" & tabId, plotType, "left"
                PushLabels rowCatName, "LABEL_COL_" & counter & "_" & tabId
            Next
        Else
            LogInfo "Table " & tabId & " asks for a chart of percentages and carries " & _
                    "no percentage column, so no series were built for it.", checkingWarning
        End If

    Case MODE_BOTH
        For counter = 1 To nbCols
            PushSeries "VALUES_COL_" & counter & "_" & tabId, plotType, "left"
            PushLabels rowCatName, "LABEL_COL_" & counter & "_" & tabId
            If specs.HasPercentage Then
                PushSeries "PERC_COL_" & counter & "_" & tabId, "point", "right"
                PushLabels rowCatName, "LABEL_COL_" & counter & "_" & tabId
            End If
        Next

    Case Else
        LogInfo "Table " & tabId & " asks for a chart and its graph setting """ & _
                graphValue & """ is none of values, percentages or both. " & _
                "No series were built for it.", checkingError

    End Select
End Sub


'@section ComplexMode
'===============================================================================
'@description Time series graph building from setup ListObjects.

'@label:build-complex-series
'@sub-title Build the complete set of graph specifications for complex mode
'@details
'Iterates over all unique graph IDs extracted by BuildGraphIdsList. For each
'graph ID, calls DefineGraphSpecs to populate the series arrays with that
'graph's series definitions, then PushGraph to package those arrays into a
'7-element BetterArray entry appended to the graphsChartSpecs collection.
'The resulting collection can be retrieved via SpecsLists after completion.
'@remarks Each iteration clears the previous graph's series data via
'ClearSeries (the first statement of DefineGraphSpecs), so the series arrays
'are temporary buffers that get packaged and reset for each graph. A graph that
'ends with no series is reported and left out, because a chart with no series
'raises 1004 out of Graphs.Format.
Private Sub BuildComplexSeries()
    Dim mGraphIds As BetterArray
    Dim graphId As String
    Dim counter As Long

    Set mGraphIds = BuildGraphIdsList()

    If mGraphIds.Length = 0 Then
        LogInfo "The graph setup table holds no graph identifier, so no time series " & _
                "chart was built.", checkingWarning
        Exit Sub
    End If

    For counter = mGraphIds.LowerBound To mGraphIds.UpperBound
        graphId = mGraphIds.Item(counter)
        DefineGraphSpecs graphId

        If BufferCount() = 0 Then
            LogInfo "Graph " & graphId & " ended with no series, so no chart was " & _
                    "drawn for it.", checkingWarning
        Else
            PushGraph graphId
        End If
    Next
End Sub

'@label:define-graph-specs
'@sub-title Define chart series for a single graph ID in complex mode
'@details
'Reads all column values for the given graphId from the graph ListObject using
'GraphValues: series IDs, axis positions, percentage overrides, chart types,
'choice values, and display labels. For each series ID, looks up the
'corresponding row in the time series ListObject to obtain a TableSpecs
'instance, which provides the table ID and section ID. The choice value is
'resolved to a named range name via TimeSeriesColumnName, then to the actual
'data range name via ResolveSeriesRangeName. If the percentages column asks for
'them, the resolved range name is replaced with its percentage twin. Each
'resolved series is pushed via PushSeries and PushLabels, using the
'section-level ROW_CATEGORIES_ range for row labels.
'@param graphId String. The unique graph identifier used to filter rows from
'the graph ListObject.
'@remarks The series buffer is cleared first, so a graph the routine gives up
'on leaves nothing of the previous graph behind. Every reason for skipping a
'series is reported.
'@depends TableSpecs
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
    Dim seriesId As String
    Dim cellRng As Range
    Dim tabRng As Range
    Dim specs As TableSpecs
    Dim tabId As String
    Dim secId As String
    Dim choiValue As String
    Dim labName As String
    Dim seriesRangeName As String

    ' Cleared before anything can go wrong, so an early exit cannot package the
    ' previous graph's series a second time.
    ClearSeries

    Set tsLo = TimeSeriesListObject()
    Set hRng = tsLo.HeaderRowRange

    ' Read all column values for this graph ID from the graph listobject
    Set idsList = GraphValues(colName:=COLUMN_SERIES_ID, graphId:=graphId)
    Set posList = GraphValues(colName:=COLUMN_AXIS, graphId:=graphId)
    Set percList = GraphValues(colName:=COLUMN_PERCENTAGES, graphId:=graphId)
    Set typeList = GraphValues(colName:=COLUMN_TYPE, graphId:=graphId)
    Set choiceList = GraphValues(colName:=COLUMN_CHOICES, graphId:=graphId)
    Set labelList = GraphValues(colName:=COLUMN_LABEL, graphId:=graphId)

    If idsList.Length = 0 Then
        LogInfo "Graph " & graphId & " has no series row in the graph setup table.", _
                checkingWarning
        Exit Sub
    End If

    ' The six lists are read row by row over the same table, so they answer the
    ' same count unless one of the columns is missing. A short list would make
    ' the reads below fall off the end of an array.
    If (posList.Length <> idsList.Length) Or (percList.Length <> idsList.Length) Or _
       (typeList.Length <> idsList.Length) Or (choiceList.Length <> idsList.Length) Or _
       (labelList.Length <> idsList.Length) Then
        LogInfo "Graph " & graphId & " could not be read: the columns of the graph " & _
                "setup table do not all answer the same number of rows.", checkingError
        Exit Sub
    End If

    ' Convert series IDs to table-based named ranges
    For counter = idsList.LowerBound To idsList.UpperBound
        ' WorksheetFunction.Trim also collapses runs of spaces inside the text,
        ' which is what a series id typed by hand needs.
        seriesId = Application.WorksheetFunction.Trim(idsList.Item(counter))

        ' Range.Find with an empty search string matches the first cell it
        ' meets, so an empty series id would attach the row to whatever sits at
        ' the top of the column.
        If LenB(seriesId) = 0 Then
            LogInfo "Graph " & graphId & " has a series row naming no series, so it " & _
                    "was skipped.", checkingWarning
            GoTo ContinueLoop
        End If

        ' Find the series row in the time series listobject
        Set cellRng = tsLo.ListColumns(1).Range.Find(What:=seriesId, _
                                                      LookIn:=xlValues, _
                                                      LookAt:=xlWhole, _
                                                      SearchOrder:=xlByRows, _
                                                      SearchDirection:=xlNext, _
                                                      MatchCase:=False)

        If cellRng Is Nothing Then
            LogInfo "Graph " & graphId & ": the series """ & seriesId & """ has no row " & _
                    "in the time series table, so it was skipped.", checkingWarning
            GoTo ContinueLoop
        End If

        If cellRng.Row <= hRng.Row Then
            LogInfo "Graph " & graphId & ": the series """ & seriesId & """ matched the " & _
                    "header of the time series table, so it was skipped.", checkingWarning
            GoTo ContinueLoop
        End If

        Set tabRng = tsLo.ListRows(cellRng.Row - hRng.Row).Range
        Set specs = TableSpecs.Create(hRng, tabRng, LinelistSpecifications.Dictionary())
        tabId = specs.TableId
        secId = specs.TableSectionId
        choiValue = Trim$(CStr(choiceList.Item(counter)))

        ' Resolve the column name from the choice value. An empty answer is
        ' already reported by TimeSeriesColumnName.
        labName = TimeSeriesColumnName(choiValue, tabId, graphId)
        If LenB(labName) = 0 Then GoTo ContinueLoop

        seriesRangeName = ResolveSeriesRangeName(choiValue, tabId, labName)

        ' Apply the percentage override when the row asks for one. For a total,
        ' the percentage twin is TOTAL_PERC_VALUES_. Replacing COL with PERC in
        ' the label name gave TOTAL_LABEL_PERC_, which nothing ever creates, so
        ' every percentage total in a time series chart plotted a name that did
        ' not resolve.
        If IsPercentageChoice(CStr(percList.Item(counter))) Then
            If IsTotalChoice(choiValue) Then
                seriesRangeName = "TOTAL_PERC_VALUES_" & tabId
            Else
                seriesRangeName = Replace(labName, "LABEL", "PERC")
            End If
        End If

        If LenB(seriesRangeName) = 0 Then
            LogInfo "Graph " & graphId & ": the series """ & seriesId & """ of table " & _
                    tabId & " resolved to no data range, so it was skipped.", checkingWarning
            GoTo ContinueLoop
        End If

        PushSeries seriesRangeName, CStr(typeList.Item(counter)), CStr(posList.Item(counter))
        PushLabels "ROW_CATEGORIES_" & secId, labName, CStr(labelList.Item(counter))

ContinueLoop:
    Next
End Sub

'@label:resolve-series-range-name
'@sub-title Resolve a choice value and label name into the data range name
'@details
'Maps a choice value to the correct data range name for chart series values
'using three rules: (1) If the choice value is the total, returns
'TOTAL_COL_VALUES_ + tabId. (2) If the label name contains the prefix
'"COLUMN_CATEGORIES_", returns INTERIOR_VALUES_ + tabId. (3) For all other
'cases, replaces the LABEL prefix in the label name with VALUES (e.g.,
'"LABEL_COL_2_xyz" becomes "VALUES_COL_2_xyz").
'@param choiValue String. The choice value from the graph setup table. The
'total triggers special handling and is matched without regard to case.
'@param tabId String. The unique table identifier for range name construction.
'@param labName String. The label range name previously resolved by
'TimeSeriesColumnName.
'@return String. The fully-qualified named range name for the data values.
Private Function ResolveSeriesRangeName(ByVal choiValue As String, _
                                         ByVal tabId As String, _
                                         ByVal labName As String) As String
    If IsTotalChoice(choiValue) Then
        ResolveSeriesRangeName = "TOTAL_COL_VALUES_" & tabId
    ElseIf InStr(1, labName, "COLUMN_CATEGORIES_") > 0 Then
        ResolveSeriesRangeName = "INTERIOR_VALUES_" & tabId
    Else
        ResolveSeriesRangeName = Replace(labName, "LABEL", "VALUES")
    End If
End Function

'@label:time-series-column-name
'@sub-title Find the named range for a choice value within a time series table
'@details
'Looks up a category label (e.g., a geographic area or demographic group) in
'the COLUMN_CATEGORIES_ range of the table on the output worksheet. When the
'choice value is the total, returns the TOTAL_LABEL_COL_ + tabId range name
'directly. For all other values, searches the COLUMN_CATEGORIES_ range for an
'exact match and returns the .Name.Name of the found cell. Answers an empty
'string when the choice value is empty, when the table has no column categories
'on the sheet, when the value is not among them, or when the matching cell
'carries no named range. Each of those files a report entry.
'@param choiValue String. The category value to look up (e.g., "Male",
'"Region A", or the total). An empty string files an entry and returns.
'@param tabId String. The unique table identifier for range name construction.
'@param graphId String. The graph being built, used in the report entries.
'@return String. The full named range name string for the matching column, or
'vbNullString.
'@remarks The COLUMN_CATEGORIES_ read is guarded. The table row it belongs to
'may have been skipped during the table pass, and then the name was never
'created; reading it raised 1004 and took every remaining time series chart of
'the sheet with it, because the loop above has no handler.
Private Function TimeSeriesColumnName(ByVal choiValue As String, _
                                       ByVal tabId As String, _
                                       ByVal graphId As String) As String
    Dim rng As Range
    Dim cellRng As Range
    Dim colName As String
    Dim sh As Worksheet

    If LenB(choiValue) = 0 Then
        LogInfo "Graph " & graphId & ": a series of table " & tabId & " names no " & _
                "choice, so it was skipped.", checkingWarning
        Exit Function
    End If

    If IsTotalChoice(choiValue) Then
        TimeSeriesColumnName = "TOTAL_LABEL_COL_" & tabId
        Exit Function
    End If

    Set sh = Wksh()

    On Error Resume Next
    Set rng = sh.Range("COLUMN_CATEGORIES_" & tabId)
    Err.Clear
    On Error GoTo 0

    If rng Is Nothing Then
        LogInfo "Table " & tabId & " has no column categories on worksheet " & sh.Name & _
                ", so the series """ & choiValue & """ of graph " & graphId & _
                " was skipped.", checkingWarning
        Exit Function
    End If

    Set cellRng = rng.Find(What:=choiValue, _
                           LookIn:=xlValues, _
                           LookAt:=xlWhole, _
                           SearchOrder:=xlByRows, _
                           SearchDirection:=xlNext, _
                           MatchCase:=True)

    If cellRng Is Nothing Then
        LogInfo "Table " & tabId & " has no column headed """ & choiValue & """, so " & _
                "that series of graph " & graphId & " was skipped.", checkingWarning
        Exit Function
    End If

    ' A cell of COLUMN_CATEGORIES_ carries a name only when the cross-table
    ' gave it one, and Range.Name raises when it has none.
    On Error Resume Next
    colName = cellRng.Name.Name
    Err.Clear
    On Error GoTo 0

    If LenB(colName) = 0 Then
        LogInfo "The column """ & choiValue & """ of table " & tabId & " carries no " & _
                "named range, so that series of graph " & graphId & " was skipped.", _
                checkingWarning
        Exit Function
    End If

    TimeSeriesColumnName = colName
End Function


'@section SeriesDataManagement
'===============================================================================
'@description Parallel array management for series names, types, positions, and labels.

'@label:init-series-arrays
'@sub-title Create fresh BetterArray instances for all six series arrays
'@details
'Creates six fresh BetterArray instances (all with LowerBound=1) to store the
'parallel series data arrays: seriesNames, seriesTypes, seriesPositions,
'seriesLabels, seriesColumnLabels, and seriesLabelPrefixes. These six arrays
'form matched parallel arrays where index N across all six describes one complete
'series entry. LowerBound is set to 1 to align with the 1-based indexing used
'throughout GraphSpecs and consumer code.
'@remarks Called at the start of BuildSimpleSeries, and by PushSeries and
'PushLabels on the first push of a complex build, when the arrays are still
'Nothing. Unconditionally replaces any existing arrays.
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

'@label:push-series
'@sub-title Append one series entry to the series-related parallel arrays
'@details
'Appends one entry to seriesNames, seriesTypes, and seriesPositions. The series
'name is typically an Excel named range reference (e.g., "VALUES_COL_1_table1")
'that the chart builder uses as the data source. If seriesNames is Nothing
'(the first push of a complex build), this method initializes all six arrays
'via InitSeriesArrays before pushing.
'@param chrtName String. The named range name for the data source.
'@param chrtType String. Chart rendering type: "bar", "hbar", "point", or "line".
'@param chrtPos String. Axis position: "left" (primary) or "right" (secondary).
'@remarks Always call PushLabels immediately after PushSeries to keep the
'parallel arrays synchronized.
Private Sub PushSeries(ByVal chrtName As String, _
                        ByVal chrtType As String, _
                        ByVal chrtPos As String)
    If this.seriesNames Is Nothing Then InitSeriesArrays

    this.seriesNames.Push chrtName
    this.seriesTypes.Push chrtType
    this.seriesPositions.Push chrtPos
End Sub

'@label:push-labels
'@sub-title Append one label entry to the label-related parallel arrays
'@details
'Appends one entry to seriesLabels, seriesColumnLabels, and seriesLabelPrefixes.
'The row category is the named range containing category axis labels. The column
'label is the named range for the series legend entry. The optional prefix is a
'display string prepended to the series label in the chart legend, used in complex
'mode where a user-defined label augments the default column label. If seriesLabels
'is Nothing, initializes all six arrays via InitSeriesArrays.
'@param chrtRowCategory String. Named range for the row (category axis) labels.
'@param chrtColumnLab String. Named range for the column (legend) label.
'@param chrtLabPref Optional String. Display prefix for the series label.
'Defaults to vbNullString.
Private Sub PushLabels(ByVal chrtRowCategory As String, _
                        ByVal chrtColumnLab As String, _
                        Optional ByVal chrtLabPref As String = vbNullString)
    If this.seriesLabels Is Nothing Then InitSeriesArrays

    this.seriesLabels.Push chrtRowCategory
    this.seriesColumnLabels.Push chrtColumnLab
    this.seriesLabelPrefixes.Push chrtLabPref
End Sub

'@label:clear-series
'@sub-title Clear all six parallel series and label arrays
'@details
'Calls .Clear on each BetterArray instance to reset them between complex mode
'graph IDs inside DefineGraphSpecs. After clearing, the arrays remain as
'instantiated (non-Nothing) objects with zero length.
'@remarks Each line tests for Nothing itself, so the routine carries no error
'handler. A real error out of .Clear is worth seeing.
Private Sub ClearSeries()
    If Not this.seriesNames Is Nothing Then this.seriesNames.Clear
    If Not this.seriesTypes Is Nothing Then this.seriesTypes.Clear
    If Not this.seriesPositions Is Nothing Then this.seriesPositions.Clear
    If Not this.seriesLabels Is Nothing Then this.seriesLabels.Clear
    If Not this.seriesColumnLabels Is Nothing Then this.seriesColumnLabels.Clear
    If Not this.seriesLabelPrefixes Is Nothing Then this.seriesLabelPrefixes.Clear
End Sub

'@label:buffer-count
'@sub-title How many series the buffer holds right now
'@details
'Both builders write into the same six arrays, so this answers for either mode.
'The public NumberOfSeries answers 0 in complex mode by contract, which is why
'the builders cannot use it.
'@return Long. The number of series currently buffered.
Private Function BufferCount() As Long
    If this.seriesNames Is Nothing Then Exit Function
    BufferCount = this.seriesNames.Length
End Function

'@label:push-graph
'@sub-title Package buffered series data into a graph specification entry
'@details
'Packages the currently-buffered series data into a single graph specification
'entry and appends it to the graphsChartSpecs collection. Each entry is a
'BetterArray containing exactly 7 elements: (1) seriesNames copy, (2) seriesTypes
'copy, (3) seriesPositions copy, (4) seriesLabels copy, (5) seriesColumnLabels
'copy, (6) seriesLabelPrefixes copy, (7) the graph title string from GraphTitle.
'All six are copied through SafeClone so the entry is independent of the mutable
'series buffers. The graphsChartSpecs collection is created on first call.
'@param graphId String. The unique graph identifier for title lookup via GraphTitle.
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

    Set tabSeries = SafeClone(this.seriesNames)
    Set tabType = SafeClone(this.seriesTypes)
    Set tabPos = SafeClone(this.seriesPositions)
    Set tabLabel = SafeClone(this.seriesLabels)
    Set tabCatLabel = SafeClone(this.seriesColumnLabels)
    Set tabPrefLabel = SafeClone(this.seriesLabelPrefixes)

    title = GraphTitle(graphId)

    graphSpecsTable.Push tabSeries, tabType, tabPos, _
                         tabLabel, tabCatLabel, tabPrefLabel, title

    If this.graphsChartSpecs Is Nothing Then
        Set this.graphsChartSpecs = New BetterArray
        this.graphsChartSpecs.LowerBound = 1
    End If

    this.graphsChartSpecs.Push graphSpecsTable
End Sub

'@label:safe-clone
'@sub-title Copy a BetterArray, answering an empty one for an empty source
'@details
'BetterArray hands out a one-element array holding Empty when it is asked for
'the items of an array of length 0, so Clone turns a length of 0 into a length
'of 1. Every consumer of this class walks the result by index, so one phantom
'item becomes one phantom series, one chart with nothing on it, and a 1004 out
'of Graphs.Format. This answers a fresh empty array in that case.
'@param source BetterArray. The array to copy.
'@return BetterArray. A copy holding the same items, and the same count.
Private Function SafeClone(ByVal source As BetterArray) As BetterArray
    Dim copy As BetterArray

    If Not source Is Nothing Then
        If source.Length > 0 Then
            Set SafeClone = source.Clone()
            Exit Function
        End If
    End If

    Set copy = New BetterArray
    copy.LowerBound = 1
    Set SafeClone = copy
End Function


'@section ListObjectHelpers
'===============================================================================
'@description ListObject accessors and lookup utilities for complex mode.

'@label:graph-list-object
'@prop-title Graph setup ListObject from the LoList collection
'@details
'Returns the graph setup table, Tab_Graph_TimeSeries. This ListObject defines
'which series belong to each graph ID, including their chart types, axis
'positions, choice values, percentage overrides, and display labels.
'@return ListObject. The graph setup table.
Private Property Get GraphListObject() As ListObject
    Set GraphListObject = this.graphLo
End Property

'@label:time-series-list-object
'@prop-title Time series setup ListObject from the LoList collection
'@details
'Returns the time series setup table, Tab_TimeSeries_Analysis. This ListObject
'maps each series ID to its full table specification row (variable name,
'section, time unit), enabling DefineGraphSpecs to construct a TableSpecs
'instance for each series.
'@return ListObject. The time series definition table.
Private Property Get TimeSeriesListObject() As ListObject
    Set TimeSeriesListObject = this.timeSeriesLo
End Property

'@label:titles-list-object
'@prop-title Title labels ListObject from the LoList collection
'@details
'Returns the title labels table, Tab_Label_TSGraph, which provides a display
'title per graph. Used by GraphTitle to look up human-readable plot titles.
'@return ListObject. The title labels table.
Private Property Get TitlesListObject() As ListObject
    Set TitlesListObject = this.titlesLo
End Property

'@label:graph-title
'@sub-title Look up the display title for a graph ID
'@details
'Walks the graph identifier column of the titles ListObject and answers the
'title cell of the matching row. Both columns are found by header name, and the
'identifier is matched without regard to case and outer spaces, because the two
'tables are typed by two hands and a case difference used to give a chart with
'an empty title and a navigation entry reading "Go to graph: ".
'@param graphId String. The unique graph identifier to look up.
'@return String. The display title, or vbNullString when the graph has no row.
Private Function GraphTitle(ByVal graphId As String) As String
    Dim titleLo As ListObject
    Dim bodyRng As Range
    Dim idIndex As Long
    Dim titleIndex As Long
    Dim counter As Long

    Set titleLo = TitlesListObject()
    Set bodyRng = titleLo.DataBodyRange
    If bodyRng Is Nothing Then Exit Function

    idIndex = LoColumnIndex(titleLo, COLUMN_GRAPH_ID)
    titleIndex = LoColumnIndex(titleLo, COLUMN_GRAPH_TITLE)

    If (idIndex < 0) Or (titleIndex < 0) Then
        LogInfo "The graph titles table " & titleLo.Name & " has no """ & _
                COLUMN_GRAPH_ID & """ or """ & COLUMN_GRAPH_TITLE & """ column, so " & _
                "graph " & graphId & " is drawn with no title.", checkingWarning
        Exit Function
    End If

    For counter = 1 To bodyRng.Rows.Count
        If StrComp(Trim$(CellText(bodyRng.Cells(counter, idIndex).Value)), _
                   Trim$(graphId), vbTextCompare) = 0 Then
            GraphTitle = CellText(bodyRng.Cells(counter, titleIndex).Value)
            Exit Function
        End If
    Next

    LogInfo "Graph " & graphId & " has no row in the graph titles table, so it is " & _
            "drawn with no title.", checkingWarning
End Function

'@label:build-graph-ids-list
'@sub-title Extract unique graph IDs from the graph setup ListObject
'@details
'Walks the graph identifier column of the graph setup table, keeps every
'non-empty value that is not already in the list, and preserves the order of
'first appearance.
'@return BetterArray. Unique graph ID strings with LowerBound = 1. Empty when
'the table has no data row or no identifier column.
'@remarks The duplicate check uses BetterArray.Includes, which performs a linear
'scan. The graph setup table typically has fewer than 100 rows.
Private Function BuildGraphIdsList() As BetterArray
    Dim idsList As BetterArray
    Dim graphLo As ListObject
    Dim idIndex As Long
    Dim bodyRng As Range
    Dim counter As Long
    Dim idValue As String

    Set idsList = New BetterArray
    idsList.LowerBound = 1
    Set BuildGraphIdsList = idsList

    Set graphLo = GraphListObject()
    Set bodyRng = graphLo.DataBodyRange
    If bodyRng Is Nothing Then Exit Function

    idIndex = LoColumnIndex(graphLo, COLUMN_GRAPH_ID)
    If idIndex < 0 Then
        LogInfo "The graph setup table " & graphLo.Name & " has no """ & _
                COLUMN_GRAPH_ID & """ column, so no time series chart was built.", _
                checkingError
        Exit Function
    End If

    For counter = 1 To bodyRng.Rows.Count
        idValue = CellText(bodyRng.Cells(counter, idIndex).Value)
        If LenB(idValue) > 0 Then
            If Not idsList.Includes(idValue) Then idsList.Push idValue
        End If
    Next
End Function

'@label:graph-values
'@sub-title Extract filtered column values from the graph setup ListObject
'@details
'Answers the values of one column, for the rows whose graph identifier matches
'the given graphId, in row order. The walk reads the identifier column and the
'wanted column of the same row, so the six lists DefineGraphSpecs reads always
'line up by index. An empty cell answers an empty string.
'@param colName String. Header name of the column to extract (e.g. "series id",
'"axis", "type", "choices", "label", "percentages").
'@param graphId String. The graph ID value to filter rows by.
'@return BetterArray. The extracted cell values, LowerBound = 1. Empty when the
'table has no data row, no identifier column, or no column of that name.
'@remarks This used to be driven by Range.Find with a search range narrowed on
'every iteration. Find wraps around by default, and none of the calls named
'LookIn, so the search inherited whatever the last search in the application
'used, the user's own Ctrl+F dialog included.
Private Function GraphValues(ByVal colName As String, _
                              ByVal graphId As String) As BetterArray
    Dim grLo As ListObject
    Dim valuesList As BetterArray
    Dim idIndex As Long
    Dim colIndex As Long
    Dim bodyRng As Range
    Dim counter As Long

    Set valuesList = New BetterArray
    valuesList.LowerBound = 1
    Set GraphValues = valuesList

    Set grLo = GraphListObject()
    Set bodyRng = grLo.DataBodyRange
    If bodyRng Is Nothing Then Exit Function

    idIndex = LoColumnIndex(grLo, COLUMN_GRAPH_ID)
    If idIndex < 0 Then
        LogInfo "The graph setup table " & grLo.Name & " has no """ & COLUMN_GRAPH_ID & _
                """ column, so graph " & graphId & " was skipped.", checkingError
        Exit Function
    End If

    colIndex = LoColumnIndex(grLo, colName)
    If colIndex < 0 Then
        LogInfo "The graph setup table " & grLo.Name & " has no """ & colName & _
                """ column, so graph " & graphId & " was skipped.", checkingError
        Exit Function
    End If

    For counter = 1 To bodyRng.Rows.Count
        If CellText(bodyRng.Cells(counter, idIndex).Value) = graphId Then
            valuesList.Push CellText(bodyRng.Cells(counter, colIndex).Value)
        End If
    Next
End Function

'@label:lo-column-index
'@sub-title Look up a column index in a ListObject header by name
'@details
'Walks the header row once. A header equal to the search term wins outright; a
'header that contains it is the fallback, which is what lets "axis" find
'"Y-Axis" and "type" find "Chart type". When more than one header contains the
'term the first is used and the ambiguity is reported, because a renamed or an
'added column would otherwise change what the chart plots with nothing said.
'@param lo ListObject. The ListObject whose header row is searched.
'@param colName String. The column header text to search for. Matched without
'regard to case or outer spaces.
'@return Long. The 1-based index relative to the ListObject, or -1 when no
'header carries the term.
Private Function LoColumnIndex(ByVal lo As ListObject, _
                                ByVal colName As String) As Long
    Dim hRng As Range
    Dim counter As Long
    Dim headerText As String
    Dim token As String
    Dim partialIndex As Long
    Dim partialCount As Long

    Set hRng = lo.HeaderRowRange
    token = LCase$(Trim$(colName))

    For counter = 1 To hRng.Columns.Count
        headerText = LCase$(Trim$(CellText(hRng.Cells(1, counter).Value)))

        If headerText = token Then
            LoColumnIndex = counter
            Exit Function
        End If

        If InStr(1, headerText, token, vbBinaryCompare) > 0 Then
            partialCount = partialCount + 1
            If partialIndex = 0 Then partialIndex = counter
        End If
    Next

    If partialCount > 1 Then
        LogInfo "The column """ & colName & """ matches " & partialCount & " headers " & _
                "of table " & lo.Name & ". The first one was used.", checkingWarning
    End If

    If partialIndex = 0 Then
        LoColumnIndex = -1
    Else
        LoColumnIndex = partialIndex
    End If
End Function


'@section Validation
'===============================================================================
'@description ListObject validation for complex mode construction.

'@label:validate-list-objects
'@sub-title Validate the BetterArray of ListObjects required by CreateRangeSpecs
'@details
'Three levels of validation. The array must hold exactly three entries. Each
'entry must be a ListObject carrying at least one data row. And each one must
'be the table its position expects, which is settled by comparing its name
'against Tab_Graph_TimeSeries, Tab_TimeSeries_Analysis and Tab_Label_TSGraph.
'@param loTable BetterArray. The ListObject collection to validate. Exactly
'three entries in order: graph setup, time series, title labels.
'@throws InvalidArgument When loTable is Nothing, holds the wrong number of
'entries, holds something that is not a ListObject, or holds an empty one.
'@throws ErrorUnexpectedState When the ListObjects are in the wrong order.
'@remarks The order used to be read off a caption two rows above the header
'row. That gap is two rows in a release file and four rows in every development
'file, so the check passed for one family and refused the other.
Private Sub ValidateListObjects(ByVal loTable As BetterArray)
    Dim counter As Long
    Dim position As Long
    Dim lo As ListObject
    Dim testObj As Object
    Dim expectedNames As Variant

    If loTable Is Nothing Then
        ThrowError InvalidArgument, "ListObject collection is empty"
    End If

    If loTable.Length <> EXPECTED_LISTOBJECT_COUNT Then
        ThrowError InvalidArgument, _
                   "Expected 3 ListObjects (graph, time series, titles), got " & loTable.Length
    End If

    expectedNames = Array(TABLE_GRAPH_TIMESERIES, TABLE_TIMESERIES, TABLE_LABEL_TSGRAPH)

    For counter = loTable.LowerBound To loTable.UpperBound
        position = counter - loTable.LowerBound + 1

        ' Set on a value raises 424 before TypeName can answer, so the kind of
        ' the item is asked first.
        If Not IsObject(loTable.Item(counter)) Then
            ThrowError InvalidArgument, "Item " & position & " is not a ListObject"
        End If

        Set testObj = loTable.Item(counter)
        If testObj Is Nothing Then
            ThrowError InvalidArgument, "ListObject at index " & position & " is Nothing"
        End If
        If TypeName(testObj) <> "ListObject" Then
            ThrowError InvalidArgument, "Item " & position & " is not a ListObject"
        End If

        Set lo = testObj
        If lo.DataBodyRange Is Nothing Then
            ThrowError InvalidArgument, "ListObject " & lo.Name & " has no data"
        End If

        If StrComp(lo.Name, CStr(expectedNames(position - 1)), vbTextCompare) <> 0 Then
            ThrowError ErrorUnexpectedState, _
                       "ListObject " & position & " must be " & _
                       expectedNames(position - 1) & ". Found: " & lo.Name
        End If
    Next
End Sub


'@section Properties
'===============================================================================
'@description Public accessors for series data, validity, and worksheet reference.

'@label:wksh
'@prop-title Output worksheet where charts will be drawn
'@details
'Dispatches based on the current graph mode. In simple mode, returns
'Table.Wksh -- the worksheet containing the cross-table's named ranges. In
'complex mode, returns OutputSheet -- a dedicated output worksheet provided
'during factory construction. This abstraction allows consumers to call Wksh
'without knowing the mode.
'@return Worksheet. The target worksheet for chart creation.
Public Property Get Wksh() As Worksheet
    Select Case GraphType
    Case GraphTypeSimple
        Set Wksh = Table.Wksh
    Case GraphTypeComplex
        Set Wksh = OutputSheet
    End Select
End Property

'@label:number-of-series
'@prop-title Number of chart series in simple mode
'@details
'Returns the length of the seriesNames BetterArray after CreateSeries has been
'called. Returns 0 in complex mode (GraphTypeComplex), because complex mode
'stores series data within the nested graphsChartSpecs collection. Also returns
'0 if seriesNames is Nothing (CreateSeries not yet called).
'@return Long. Count of series entries in simple mode, or 0.
'@export
Public Property Get NumberOfSeries() As Long
    If GraphType = GraphTypeComplex Then Exit Property
    NumberOfSeries = BufferCount()
End Property

'@label:number-of-graphs
'@prop-title Number of distinct graphs in complex mode
'@details
'Each graph corresponds to one unique graph ID from the graph setup ListObject
'and has its own entry in the graphsChartSpecs collection. Returns 0 in simple
'mode (GraphTypeSimple) or if graphsChartSpecs is Nothing (CreateSeries not yet
'called or produced no valid graphs).
'@return Long. Count of graph entries in complex mode, or 0.
'@export
Public Property Get NumberOfGraphs() As Long
    If (this.graphsChartSpecs Is Nothing) Or (GraphType <> GraphTypeComplex) Then
        NumberOfGraphs = 0
        Exit Property
    End If
    NumberOfGraphs = this.graphsChartSpecs.Length
End Property

'@label:series-name
'@prop-title Named range reference for series data at a given index
'@details
'Returns the Excel named range string (e.g., "VALUES_COL_1_table1") that a chart
'builder uses as the data source for this series. Only valid in simple mode after
'CreateSeries has been called.
'@param index Long. 1-based index of the series (1 to NumberOfSeries).
'@return String. Named range reference for the series data.
'@throws InvalidArgument When index is out of bounds.
'@export
Public Property Get SeriesName(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesName = this.seriesNames.Item(index)
End Property

'@label:series-type
'@prop-title Chart rendering type for a series at a given index
'@details
'Returns the chart type string controlling how the series is rendered. Common
'values include "bar" (vertical bars), "hbar" (horizontal bars), "point"
'(scatter for percentage overlays), and "line". Only valid in simple mode.
'@param index Long. 1-based series index.
'@return String. Chart rendering type.
'@throws InvalidArgument When index is out of bounds.
'@export
Public Property Get SeriesType(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesType = this.seriesTypes.Item(index)
End Property

'@label:series-pos
'@prop-title Axis position for a series at a given index
'@details
'Returns "left" for the primary value axis or "right" for the secondary axis,
'typically used for percentage overlays. Only valid in simple mode.
'@param index Long. 1-based series index.
'@return String. Axis placement ("left" or "right").
'@throws InvalidArgument When index is out of bounds.
'@export
Public Property Get SeriesPos(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesPos = this.seriesPositions.Item(index)
End Property

'@label:series-label
'@prop-title Row category named range for a series at a given index
'@details
'Returns the named range providing category axis labels (e.g.,
'"ROW_CATEGORIES_table1") displayed along the chart axis. Only valid in simple
'mode after CreateSeries has been called.
'@param index Long. 1-based series index.
'@return String. Named range for row category labels.
'@throws InvalidArgument When index is out of bounds.
'@export
Public Property Get SeriesLabel(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesLabel = this.seriesLabels.Item(index)
End Property

'@label:series-column-label
'@prop-title Column label named range for a series at a given index
'@details
'Returns the named range providing the legend entry text (e.g.,
'"LABEL_COL_2_table1") that identifies this series in the chart legend. Only
'valid in simple mode after CreateSeries has been called.
'@param index Long. 1-based series index.
'@return String. Named range for the legend entry.
'@throws InvalidArgument When index is out of bounds.
'@export
Public Property Get SeriesColumnLabel(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesColumnLabel = this.seriesColumnLabels.Item(index)
End Property

'@label:series-label-prefix
'@prop-title Label prefix for a series at a given index
'@details
'Returns the optional display prefix prepended to the series legend entry. In
'simple mode this is always vbNullString. In complex mode (when accessed via
'SpecsLists), this contains the user-defined "label" value from the graph setup
'ListObject. Only valid in simple mode after CreateSeries has been called.
'@param index Long. 1-based series index.
'@return String. Prefix string or vbNullString when none is set.
'@throws InvalidArgument When index is out of bounds.
'@export
Public Property Get SeriesLabelPrefix(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesLabelPrefix = this.seriesLabelPrefixes.Item(index)
End Property

'@label:valid
'@prop-title Whether the graph specifications were successfully built
'@details
'The result is cached after the first evaluation using the validityTested and
'validityResult flags. On the first call, invokes CreateSeries to build all
'series data, then determines validity: in simple mode when at least one series
'was built; in complex mode when at least one graph ID produced series.
'@return Boolean. True if series data is available for consumption.
'@remarks Because this property calls CreateSeries on first access, it can be
'used as a validity gate without separately calling CreateSeries. The caching
'means the answer is settled for the life of the instance.
'@export
Public Property Get Valid() As Boolean
    If this.validityTested Then
        Valid = this.validityResult
        Exit Property
    End If

    CreateSeries

    If GraphType = GraphTypeSimple Then
        this.validityResult = (BufferCount() > 0)
    Else
        this.validityResult = (NumberOfGraphs() > 0)
    End If

    this.validityTested = True
    Valid = this.validityResult
End Property

'@label:specs-lists
'@prop-title Nested collection of all graph specifications
'@details
'Returns a copy of the graphsChartSpecs collection. Each element is a 7-element
'BetterArray representing one graph: [seriesNames, seriesTypes, seriesPositions,
'seriesLabels, seriesColumnLabels, seriesLabelPrefixes, titleString]. Returns a
'fresh empty BetterArray when nothing was built.
'@return BetterArray. Copy of the collection of graph spec entries, or empty.
'@remarks The copy is one level deep: each entry is the same BetterArray this
'class holds, and so are the six lists inside it. No consumer writes into them.
'Only meaningful in complex mode; simple mode data is read through the indexed
'accessors.
'@export
Public Property Get SpecsLists() As BetterArray
    Set SpecsLists = SafeClone(this.graphsChartSpecs)
End Property


'@section Checking
'===============================================================================
'@description The report of everything this class decided to leave out.

'@label:specs-id
'@sub-title A name for this instance that no other instance shares
'@details
'AnalysisOutput merges the entries of every collaborator into one report, and
'Checking.Add raises on a duplicate key, so a key made of a bare counter
'collides as soon as two instances file an entry. In simple mode one instance
'serves one table, so the table id names it. In complex mode one instance
'serves one output worksheet.
'@return String. The identifier every key of this instance carries.
Private Function SpecsId() As String
    Select Case GraphType

    Case GraphTypeSimple
        If Not this.tabl Is Nothing Then SpecsId = this.tabl.Specifications.TableId

    Case GraphTypeComplex
        If Not this.outSh Is Nothing Then SpecsId = this.outSh.Name

    End Select

    If LenB(SpecsId) = 0 Then SpecsId = "unplaced"
End Function

'@label:log-info
'@sub-title Record a diagnostic entry in the internal checking store
'@details
'Creates the Checking store on first use and appends the message under a key
'naming this class, this instance and a counter. The store is read through
'HasCheckings and CheckingValues.
'@param label String. The message to record.
'@param scope Optional Byte. Severity from CheckingLogType. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)
    Dim specsKey As String

    specsKey = SpecsId()

    If internalChecks Is Nothing Then _
        Set internalChecks = Checking.Create("Graph specifications " & specsKey)

    internalChecks.Add CLASS_NAME & "-" & specsKey & "-" & CStr(checkCounter), _
                       label, scope
    checkCounter = checkCounter + 1

    this.hasChecks = True
End Sub

'@label:has-checkings
'@prop-title Whether anything was reported while the series were built
'@return Boolean. True when at least one entry was filed.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title The report of this build
'@return Checking. The store, or Nothing when nothing was filed.
'@export
Public Property Get CheckingValues() As Checking
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = internalChecks
End Property


'@section ErrorHandling
'===============================================================================
'@description Index validation, value reading and error raising helpers.

'@label:cell-text
'@sub-title Read one cell value as text
'@details
'CStr on an Error variant raises runtime error 13, and the analysis setup sheet
'carries formula columns, so an ordinary edit can produce one. An error value, a
'Null and an empty cell all answer an empty string here.
'@param cellValue Variant. One value read out of a worksheet range.
'@return String. The text of the cell, or vbNullString.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@label:is-total-choice
'@sub-title Whether a choice value asks for the total column
'@details
'The value comes out of a setup cell, so it is matched trimmed and without
'regard to case. It used to be compared against the literal "Total", which a
'cell holding " total " failed.
'@param choiValue String. The choice value read from the graph setup table.
'@return Boolean. True when the value names the total.
Private Function IsTotalChoice(ByVal choiValue As String) As Boolean
    IsTotalChoice = (LCase$(Trim$(choiValue)) = CHOICE_TOTAL)
End Function

'@label:is-percentage-choice
'@sub-title Whether a graph setup row asks for percentages
'@param percValue String. The "plot values or percentages" cell of the row.
'@return Boolean. True when the row asks for percentages.
Private Function IsPercentageChoice(ByVal percValue As String) As Boolean
    IsPercentageChoice = (LCase$(Trim$(percValue)) = MODE_PERCENTAGES)
End Function

'@label:validate-series-index
'@sub-title Validate that a series index is within bounds
'@param index Long. The 1-based series index to validate.
'@throws InvalidArgument When index is less than 1 or greater than NumberOfSeries.
Private Sub ValidateSeriesIndex(ByVal index As Long)
    If (index > NumberOfSeries) Or (index < 1) Then
        ThrowError InvalidArgument, _
                   "Series index " & index & " out of bounds (1 to " & NumberOfSeries & ")"
    End If
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME, providing
'a consistent stack trace across all methods in this class.
'@param errNumber Long. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
