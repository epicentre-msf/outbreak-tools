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

'@class GraphSpecs
'@description
'Builds and stores chart series specifications for both simple (non-time-series)
'and complex (time series) analysis graphs. In simple mode, the class reads chart
'configuration directly from an ICrossTable and produces flat parallel arrays of
'series metadata (names, types, positions, labels). In complex mode, it reads from
'three setup ListObjects to produce a nested BetterArray collection of graph entries,
'each containing its own set of parallel series arrays plus a display title. Consumers
'create an instance via Create (simple) or CreateRangeSpecs (complex), call
'CreateSeries (or let Valid trigger it lazily), then read series data via the
'IGraphSpecs interface.
'@depends ICrossTable, ITableSpecs, TableSpecs, ILinelistSpecs, BetterArray
'@version 1.0 (2026-02-09)



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
'graphing). The returned IGraphSpecs instance stores a reference to the cross-table
'and can later produce chart series definitions via CreateSeries.
'@param tabl ICrossTable. The fully-built cross-table whose named ranges contain
'the data to graph. Must not be Nothing, ScopeTimeSeries, or ScopeGlobalSummary.
'@return IGraphSpecs. An instance configured in simple mode (GraphTypeSimple)
'ready for series creation via CreateSeries.
'@throws InvalidArgument When tabl is Nothing, ScopeTimeSeries, or ScopeGlobalSummary.
'@remarks After calling Create, you must call CreateSeries before accessing any
'series properties (SeriesName, SeriesType, etc.). The cross-table must already
'have its named ranges written to the worksheet before series data can be consumed.
'@export
Public Function Create(ByVal tabl As ICrossTable) As IGraphSpecs
    If tabl Is Nothing Then
        ThrowError InvalidArgument, "Cross-table is not defined"
    End If

    If tabl.Specifications.TableScope = ScopeTimeSeries Then
        ThrowError InvalidArgument, _
                   "Simple mode is for non-time-series tables. Use CreateRangeSpecs for time series."
    End If

    If tabl.Specifications.TableScope = ScopeGlobalSummary Then
        ThrowError InvalidArgument, "Global summary tables do not support graphs"
    End If

    With New GraphSpecs
        Set .Table = tabl
        .GraphType = GraphTypeSimple
        Set Create = .Self
    End With
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
'[graphLo, timeSeriesLo, titleLo]. Each must be a non-empty ListObject with the
'correct title label in its header area.
'@param outputSh Worksheet. The worksheet where time series graphs will be drawn.
'@param lData ILinelistSpecs. Linelist-level metadata (section names, sheet info)
'needed to resolve table specifications during series building.
'@return IGraphSpecs. An instance configured in complex mode (GraphTypeComplex)
'ready for series creation via CreateSeries.
'@throws InvalidArgument When loTable fails validation, lData is Nothing, or
'outputSh is Nothing.
'@throws ErrorUnexpectedState When ListObjects are in the wrong order.
'@remarks The 3 ListObjects must appear in the exact order validated by
'ValidateListObjects: graph table first, time series table second, title labels
'table third.
'@export
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

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor so consuming code can fluently retrieve the interface
'reference from the predeclared Create or CreateRangeSpecs method.
'@return IGraphSpecs. The current instance cast to the interface.
Public Property Get Self() As IGraphSpecs
    Set Self = Me
End Property


'@section InternalSetters
'===============================================================================
'@description Property accessors used only during factory construction.
'These are Public to support the With New pattern but are not part of the
'caller contract.

'@label:table
'@prop-title Cross-table reference for simple mode
'@return ICrossTable. The stored cross-table instance.
Public Property Get Table() As ICrossTable
    Set Table = this.tabl
End Property

'@label:table-set
'@prop-title Assign the cross-table reference
'@param tabl ICrossTable. The cross-table to store.
Public Property Set Table(ByVal tabl As ICrossTable)
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
'@param loTable BetterArray. The 3-element ListObject array to clone and store.
Public Property Set LoList(ByVal loTable As BetterArray)
    Set this.loList = loTable.Clone()
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
    Set this.outSh = sh
End Property

'@label:linelist-specifications
'@prop-title Linelist specifications for complex mode
'@return ILinelistSpecs. The stored linelist specifications instance.
Public Property Get LinelistSpecifications() As ILinelistSpecs
    Set LinelistSpecifications = this.lData
End Property

'@label:linelist-specifications-set
'@prop-title Assign the linelist specifications
'@param lData ILinelistSpecs. The linelist specifications to store.
Public Property Set LinelistSpecifications(ByVal lData As ILinelistSpecs)
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

    End Select
End Sub


'@section SimpleMode
'===============================================================================
'@description Series building logic for non-time-series cross-tables.

'@label:build-simple-series
'@sub-title Build chart series for a non-time-series cross-table
'@details
'Reads the table scope and configuration from the cross-table's ITableSpecs to
'determine what kind of series to generate. The "flip" setting controls plot
'direction: when flip is "yes", horizontal bars ("hbar") are used; otherwise
'vertical bars ("bar") are the default. For spatio-temporal tables, the number
'of columns is overridden by the "n geo" setting (defaulting to 5 if not
'specified), and the section-level table ID (tabSecId) is captured for use as
'row category ranges. After initializing fresh series arrays, dispatches to
'BuildUnivariateSeries for ScopeUnivariate tables, or BuildMultiColumnSeries
'for ScopeBivariate, ScopeSpatial, and ScopeSpatioTemporal tables.
'@remarks The "n geo" override uses On Error Resume Next to gracefully handle
'missing or non-numeric values, falling back to 5 columns.
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
    tabType = specs.TableScope
    nbCols = Table.NumberOfColumns

    ' For spatio-temporal, column count comes from "n geo" setting
    If tabType = ScopeSpatioTemporal Then
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

    Case ScopeUnivariate
        BuildUnivariateSeries tabId, plotType, specs.HasPercentage

    Case ScopeBivariate, ScopeSpatial, ScopeSpatioTemporal
        BuildMultiColumnSeries specs, tabId, tabSecId, tabType, nbCols, plotType

    End Select
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

    ' Percentage overlay (points on right axis) — only for vertical bars
    If hasPerc And (plotType = "bar") Then
        PushSeries "PERC_COL_1_" & tabId, "point", "right"
        PushLabels "ROW_CATEGORIES_" & tabId, "PERC_LABEL_COL_" & tabId
    End If
End Sub

'@label:build-multi-column-series
'@sub-title Build chart series for multi-column cross-tables
'@details
'Builds series for bivariate, spatial, or spatio-temporal tables. The graph mode
'is determined by the "graph" setting in the table specs: (1) "values" or "yes"
'pushes one bar series per column referencing VALUES_COL_<n>_ named ranges on
'the left axis; (2) "percentages" pushes one bar series per column referencing
'PERC_COL_<n>_ ranges (only if HasPercentage is True); (3) "both" pushes value
'bars on the left axis AND percentage points on the right axis for each column,
'creating a dual-axis overlay chart. For spatio-temporal tables, the row
'categories use the section-level range (ROW_CATEGORIES_ + tabSecId) instead
'of the table-level range, because spatio-temporal tables share row categories
'across all tables in the same section.
'@param specs ITableSpecs. The table specifications providing access to the
'"graph" and "flip" settings as well as HasPercentage.
'@param tabId String. The unique table identifier for named range construction.
'@param tabSecId String. The section-level table identifier, used only for
'spatio-temporal tables. Empty for other types.
'@param tabType Byte. The numeric table scope (ScopeBivariate, ScopeSpatial, or
'ScopeSpatioTemporal).
'@param nbCols Long. The number of data columns to iterate over.
'@param plotType String. Chart series type: "bar" or "hbar".
'@remarks When graphValue is "both", percentage point series are only added if
'specs.HasPercentage is True. The "percentages" mode silently produces no series
'if HasPercentage is False.
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
            If tabType = ScopeSpatioTemporal Then
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
'ClearSeries (called inside DefineGraphSpecs), so the series arrays are
'temporary buffers that get packaged and reset for each graph.
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

'@label:define-graph-specs
'@sub-title Define chart series for a single graph ID in complex mode
'@details
'Reads all column values for the given graphId from the graph ListObject using
'GraphValues: series IDs, axis positions, percentage overrides, chart types,
'choice values, and display labels. Clears any previously buffered series data
'via ClearSeries. For each series ID, looks up the corresponding row in the
'time series ListObject to obtain a TableSpecs instance, which provides the
'table ID and section ID. The choice value is resolved to a named range name
'via TimeSeriesColumnName, then to the actual data range name via
'ResolveSeriesRangeName. If the "percentages" column indicates a percentage
'override, the resolved range name is replaced with its PERC_ equivalent.
'Each valid series is pushed via PushSeries and PushLabels, using the
'section-level ROW_CATEGORIES_ range for row labels.
'@param graphId String. The unique graph identifier used to filter rows from
'the graph ListObject.
'@remarks Series IDs that cannot be found in the time series ListObject (or
'that resolve to the header row) are silently skipped via GoTo ContinueLoop.
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
        Set specs = TableSpecs.Create(hRng, tabRng, LinelistSpecifications.Dictionary())
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

'@label:resolve-series-range-name
'@sub-title Resolve a choice value and label name into the data range name
'@details
'Maps a choice value to the correct data range name for chart series values
'using three rules: (1) If the choice value is "Total", returns
'TOTAL_COL_VALUES_ + tabId. (2) If the label name contains the prefix
'"COLUMN_CATEGORIES_", returns INTERIOR_VALUES_ + tabId. (3) For all other
'cases, replaces the LABEL prefix in the label name with VALUES (e.g.,
'"LABEL_COL_2_xyz" becomes "VALUES_COL_2_xyz").
'@param choiValue String. The choice value from the graph setup table. "Total"
'triggers special handling.
'@param tabId String. The unique table identifier for range name construction.
'@param labName String. The label range name previously resolved by
'TimeSeriesColumnName.
'@return String. The fully-qualified named range name for the data values.
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

'@label:time-series-column-name
'@sub-title Find the named range for a choice value within a time series table
'@details
'Looks up a category label (e.g., a geographic area or demographic group) in
'the COLUMN_CATEGORIES_ range of the table on the output worksheet. If the
'choice value is "Total", returns the predefined TOTAL_LABEL_COL_ + tabId range
'name directly. For all other values, searches the COLUMN_CATEGORIES_ range for
'an exact match and returns the .Name.Name of the found cell. Returns an empty
'string if the choice value is empty or not found.
'@param choiValue String. The category value to look up (e.g., "Male", "Region A",
'or "Total"). An empty string causes an early exit returning vbNullString.
'@param tabId String. The unique table identifier for range name construction.
'@return String. The full named range name string for the matching column, or
'vbNullString if no match is found.
'@remarks Uses On Error Resume Next when reading cellRng.Name.Name because some
'cells in the COLUMN_CATEGORIES_ range may not have a named range assigned.
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
'@remarks Called at the start of BuildSimpleSeries and lazily by PushSeries and
'PushLabels when the arrays are Nothing (after ClearSeries in complex mode).
'Unconditionally replaces any existing arrays.
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
'(which occurs in complex mode after ClearSeries resets the arrays), this method
'lazy-initializes all six arrays via InitSeriesArrays before pushing.
'@param chrtName String. The named range name for the data source.
'@param chrtType String. Chart rendering type: "bar", "hbar", "point", or "line".
'@param chrtPos String. Axis position: "left" (primary) or "right" (secondary).
'@remarks Always call PushLabels immediately after PushSeries to keep the
'parallel arrays synchronized.
Private Sub PushSeries(ByVal chrtName As String, _
                        ByVal chrtType As String, _
                        ByVal chrtPos As String)
    ' Lazy init for complex mode (called per graph ID after ClearSeries)
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
'is Nothing, lazy-initializes all six arrays via InitSeriesArrays.
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
'instantiated (non-Nothing) objects with zero length. Uses On Error Resume Next
'because during the first call the arrays may still be Nothing.
'@remarks The On Error Resume Next wrapper is intentional and safe: the only
'possible error is calling .Clear on a Nothing reference.
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

'@label:push-graph
'@sub-title Package buffered series data into a graph specification entry
'@details
'Packages the currently-buffered series data into a single graph specification
'entry and appends it to the graphsChartSpecs collection. Each entry is a
'BetterArray containing exactly 7 elements: (1) seriesNames clone, (2) seriesTypes
'clone, (3) seriesPositions clone, (4) seriesLabels clone, (5) seriesColumnLabels
'clone, (6) seriesLabelPrefixes clone, (7) the graph title string from GraphTitle.
'All BetterArray elements are cloned to ensure independence from the mutable
'series buffers. The graphsChartSpecs collection is lazy-initialized on first call.
'@param graphId String. The unique graph identifier for title lookup via GraphTitle.
'@remarks BetterArray.Clone() has a known quirk where it can produce an array with
'one extra element. A safety check compares the cloned length to the original and
'calls .Pop on all six clones if they are one element too long.
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


'@section ListObjectHelpers
'===============================================================================
'@description ListObject accessors and lookup utilities for complex mode.

'@label:graph-list-object
'@prop-title Graph setup ListObject from the LoList collection
'@details
'Returns the first ListObject in the LoList array, which is the graph setup
'table. This ListObject defines which series belong to each graph ID, including
'their chart types, axis positions, choice values, percentage overrides, and
'display labels.
'@return ListObject. The graph setup table ("graph on time series").
Private Property Get GraphListObject() As ListObject
    Set GraphListObject = LoList.Item(LoList.LowerBound)
End Property

'@label:time-series-list-object
'@prop-title Time series setup ListObject from the LoList collection
'@details
'Returns the second ListObject in the LoList array, which is the time series
'setup table. This ListObject maps each series ID to its full table specification
'row (variable name, section, time unit, etc.), enabling DefineGraphSpecs to
'construct a TableSpecs instance for each series.
'@return ListObject. The time series definition table ("time series analysis").
Private Property Get TimeSeriesListObject() As ListObject
    Set TimeSeriesListObject = LoList.Item(LoList.LowerBound + 1)
End Property

'@label:titles-list-object
'@prop-title Title labels ListObject from the LoList collection
'@details
'Returns the third ListObject in the LoList array, which provides display
'titles for each graph. Keyed by graph ID in column 3, with the title text in
'column 1. Used by GraphTitle to look up human-readable plot titles.
'@return ListObject. The title labels table ("labels for time series graphs").
Private Property Get TitlesListObject() As ListObject
    Set TitlesListObject = LoList.Item(LoList.LowerBound + 2)
End Property

'@label:graph-title
'@sub-title Look up the display title for a graph ID
'@details
'Searches the titles ListObject's third column for an exact, case-sensitive
'match of the graphId string. If found, returns the title text from the same
'row's first column (offset -2). If not found, returns an empty string.
'@param graphId String. The unique graph identifier to look up.
'@return String. The display title, or vbNullString if not found.
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

'@label:build-graph-ids-list
'@sub-title Extract unique graph IDs from the graph setup ListObject
'@details
'Iterates through every data row in the graph ID column (skipping the header),
'checks whether each cell is non-empty and not already present in the accumulator
'array (using BetterArray.Includes for duplicate detection), and pushes new
'unique values. The resulting list preserves the order of first appearance.
'@return BetterArray. A cloned array of unique graph ID strings with LowerBound=1.
'@remarks The duplicate check uses BetterArray.Includes which performs a linear
'scan (O(n*m)), acceptable because the graph setup table typically has fewer than
'100 rows.
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

'@label:graph-values
'@sub-title Extract filtered column values from the graph setup ListObject
'@details
'Returns all values from a specified column, filtered to rows whose "graph id"
'matches the given graphId. Walks through the "graph id" column progressively:
'after finding each matching row, narrows the search range to start after the
'found row, preventing re-matches. Empty cell values are replaced with "&" as
'a placeholder to maintain array index alignment. Returns an empty BetterArray
'if the graphId is not found at all.
'@param colName String. Header name of the column to extract (e.g., "series id",
'"axis", "type", "choices", "label", "percentages").
'@param graphId String. The graph ID value to filter rows by.
'@return BetterArray. A cloned array (LowerBound=1) of extracted cell values.
'Empty cells are represented as "&".
'@throws ElementNotFound When the "graph id" or colName column is not found in
'the ListObject header.
'@remarks The progressive range narrowing technique is necessary because Excel's
'Range.Find wraps around by default, which would cause infinite loops.
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

'@label:lo-column-index
'@sub-title Look up a column index in a ListObject header by partial match
'@details
'Searches the HeaderRowRange for a partial, case-insensitive match of the given
'column name using Range.Find with LookAt:=xlPart. The returned index is
'relative to the ListObject (not the worksheet), calculated as the found cell's
'absolute column minus the header range's starting column plus 1. Returns -1
'if the column name is not found in the header.
'@param lo ListObject. The ListObject whose header row will be searched.
'@param colName String. The column header text to search for. Uses partial
'matching and is case-insensitive.
'@return Long. The 1-based relative column index, or -1 if not found.
'@remarks Uses xlPart matching intentionally to allow flexible header naming.
'Partial matching could return unexpected results if multiple headers share a
'common substring.
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
'@description ListObject validation for complex mode construction.

'@label:validate-list-objects
'@sub-title Validate the BetterArray of ListObjects required by CreateRangeSpecs
'@details
'Performs three levels of validation: (1) Count validation -- the array must
'contain exactly 3 elements. (2) Type and content validation -- each element must
'be a non-Nothing ListObject with a non-Nothing DataBodyRange. (3) Order
'validation -- the ListObjects must appear in the correct order, verified by
'reading the title label from 2 rows above each ListObject's header and comparing
'against the module-level constants TAB_ON_GRAPH_TIME_SERIES,
'TAB_ON_TIME_SERIES, and TAB_ON_TITLE_GRAPHS_TIME_SERIES.
'@param loTable BetterArray. The ListObject collection to validate. Must contain
'exactly 3 entries in order: graph setup, time series, title labels.
'@throws InvalidArgument When loTable is Nothing, wrong length, contains
'non-ListObject entries, or has empty ListObjects.
'@throws ErrorUnexpectedState When ListObjects are in the wrong order.
'@remarks The order validation reads from Cells(-1, 1) relative to the
'HeaderRowRange, which is 2 rows above the header. This assumes a specific
'worksheet layout where the table title is placed in that position.
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
Public Property Get NumberOfSeries() As Long
    If (this.seriesNames Is Nothing) Or (GraphType = GraphTypeComplex) Then
        NumberOfSeries = 0
        Exit Property
    End If
    NumberOfSeries = this.seriesNames.Length
End Property

'@label:number-of-graphs
'@prop-title Number of distinct graphs in complex mode
'@details
'Each graph corresponds to one unique graph ID from the graph setup ListObject
'and has its own entry in the graphsChartSpecs collection. Returns 0 in simple
'mode (GraphTypeSimple) or if graphsChartSpecs is Nothing (CreateSeries not yet
'called or produced no valid graphs).
'@return Long. Count of graph entries in complex mode, or 0.
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
Public Property Get SeriesLabelPrefix(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesLabelPrefix = this.seriesLabelPrefixes.Item(index)
End Property

'@label:valid
'@prop-title Whether the graph specifications were successfully built
'@details
'The result is cached after the first evaluation using the validityTested and
'validityResult flags. On the first call, invokes CreateSeries to build all
'series data, then determines validity: in simple mode always True (because
'BuildSimpleSeries always succeeds if the factory accepted the table); in complex
'mode True only if graphsChartSpecs was successfully created (at least one graph
'ID produced valid series definitions).
'@return Boolean. True if series data is available for consumption.
'@remarks Because this property calls CreateSeries on first access, it can be
'used as a validity gate without separately calling CreateSeries. The caching
'means errors during CreateSeries will permanently mark the instance as invalid
'(or valid) until discarded.
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

'@label:specs-lists
'@prop-title Nested collection of all graph specifications
'@details
'Returns a cloned copy of the graphsChartSpecs collection. Each element is a
'7-element BetterArray representing one graph: [seriesNames, seriesTypes,
'seriesPositions, seriesLabels, seriesColumnLabels, seriesLabelPrefixes,
'titleString]. The clone ensures consumers cannot mutate internal state. Returns
'a fresh empty BetterArray if graphsChartSpecs is Nothing.
'@return BetterArray. Cloned collection of graph spec entries, or empty.
'@remarks Only meaningful in complex mode. In simple mode, the collection is
'never populated. Use indexed SeriesName/SeriesType/etc. accessors for simple
'mode data.
Public Property Get SpecsLists() As BetterArray
    If this.graphsChartSpecs Is Nothing Then
        Set SpecsLists = New BetterArray
    Else
        Set SpecsLists = this.graphsChartSpecs.Clone()
    End If
End Property

'@label:graph-ids
'@prop-title Unique graph identifiers for complex mode
'@details
'Returns a cloned list of all unique graph IDs, lazily built on first access by
'calling BuildGraphIdsList if the internal graphIdsList is Nothing. If
'CreateSeries has already been called, the cached list is reused. Each element
'is a string graph ID in the order of first appearance in the graph setup
'ListObject.
'@return BetterArray. Cloned list of unique graph ID strings.
'@remarks Primarily useful for complex mode consumers that need to iterate over
'graph IDs independently of SpecsLists. In simple mode, the graph setup
'ListObject is not available, so calling this property will raise an error.
Public Property Get GraphIDs() As BetterArray
    If this.graphIdsList Is Nothing Then
        Set this.graphIdsList = BuildGraphIdsList()
    End If
    Set GraphIDs = this.graphIdsList.Clone()
End Property


'@section ErrorHandling
'===============================================================================
'@description Index validation and error raising helpers.

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


'@section InterfaceImplementation
'===============================================================================
'@description Delegated members satisfying the IGraphSpecs contract.
'See the corresponding Public members above for full documentation.

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
