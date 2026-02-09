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

' @description Factory method that creates a GraphSpecs object in "simple mode" from a
'              non-time-series cross-table. Simple mode is used for univariate, bivariate,
'              spatial, and spatio-temporal analysis tables -- any table type except
'              TimeSeries and GlobalSummary. This method explicitly rejects TimeSeries tables
'              (which require the more complex CreateRangeSpecs factory) and GlobalSummary
'              tables (which do not support graphing at all). The returned IGraphSpecs
'              instance stores a reference to the cross-table and can later produce chart
'              series definitions via CreateSeries.
' @param tabl The fully-built ICrossTable whose named ranges contain the data to graph.
'             Must not be Nothing, must not be TypeTimeSeries, and must not be TypeGlobalSummary.
' @return An IGraphSpecs instance configured in simple mode (GraphTypeSimple) ready for
'         series creation via CreateSeries.
' @remark After calling Create, you must call CreateSeries before accessing any series
'         properties (SeriesName, SeriesType, etc.). The cross-table must already have its
'         named ranges written to the worksheet before series data can be consumed by a chart.
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

' @description Factory method that creates a GraphSpecs object in "complex mode" for time
'              series graphs. Complex mode requires exactly 3 ListObjects passed as a
'              BetterArray: (1) the graph setup table that defines which series belong to
'              each graph ID along with their types, axis positions, and choice values;
'              (2) the time series setup table that maps series IDs to their table
'              specifications (variable name, section, etc.); and (3) the title labels
'              table that provides the display title for each graph ID. The method validates
'              all three ListObjects via ValidateListObjects before constructing the instance.
'              Complex mode produces a nested BetterArray collection of graph spec entries
'              (one per unique graph ID) when CreateSeries is called later.
' @param loTable A BetterArray containing exactly 3 ListObject entries in order:
'                [graphLo, timeSeriesLo, titleLo]. Each must be a non-empty ListObject
'                with the correct title label in its header area.
' @param outputSh The Worksheet where time series graphs will be drawn. This is stored
'                 and returned by the Wksh property in complex mode.
' @param lData The ILinelistSpecs instance providing linelist-level metadata (section
'              names, sheet info, etc.) needed to resolve table specifications during
'              series building.
' @return An IGraphSpecs instance configured in complex mode (GraphTypeComplex) ready
'         for series creation via CreateSeries.
' @remark The 3 ListObjects must appear in the exact order validated by ValidateListObjects:
'         graph table first, time series table second, title labels table third. Passing
'         them in the wrong order will raise an ErrorUnexpectedState error.
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

' @description Dispatches series building to the appropriate handler based on the current
'              GraphType. For simple mode (GraphTypeSimple), delegates to BuildSimpleSeries
'              which reads chart configuration directly from the cross-table's named ranges.
'              For complex mode (GraphTypeComplex), delegates to BuildComplexSeries which
'              iterates over unique graph IDs from the setup ListObjects and builds a nested
'              BetterArray collection of graph spec entries. This method must be called
'              explicitly before accessing any series data properties (SeriesName, SeriesType,
'              SeriesPos, etc.) or the SpecsLists collection.
' @remark This method is also called internally by the Valid property if the validity
'         cache has not yet been populated. Calling it multiple times is safe but will
'         rebuild all series data from scratch each time.
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

' @description Builds chart series for a non-time-series cross-table in simple mode. Reads
'              the table type and configuration from the cross-table's ITableSpecs to determine
'              what kind of series to generate. The "flip" setting controls plot direction:
'              when flip is "yes", horizontal bars ("hbar") are used; otherwise vertical bars
'              ("bar") are the default. For spatio-temporal tables, the number of columns is
'              overridden by the "n geo" setting (defaulting to 5 if not specified), and the
'              section-level table ID (tabSecId) is captured for use as row category ranges.
'              After initializing fresh series arrays, dispatches to BuildUnivariateSeries for
'              TypeUnivariate tables, or BuildMultiColumnSeries for TypeBivariate, TypeSpatial,
'              and TypeSpatioTemporal tables.
' @remark The "n geo" override uses On Error Resume Next to gracefully handle missing or
'         non-numeric values, falling back to 5 columns. This is the only place where
'         spatio-temporal column count is determined.
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

' @description Builds chart series for a univariate (single-variable) cross-table. Pushes
'              one bar series referencing the VALUES_COL_1_ named range for the primary data,
'              paired with the ROW_CATEGORIES_ range for axis labels and LABEL_COL_1_ for
'              the legend entry. If the table has a percentage column and the plot direction
'              is vertical bars (plotType = "bar"), an additional point series is pushed
'              referencing PERC_COL_1_ on the right axis, allowing a percentage overlay on
'              top of the value bars. Horizontal bars ("hbar") do not get the percentage
'              overlay because point series on a horizontal chart do not render well.
' @param tabId The unique table identifier used as a suffix for all named range lookups
'              (e.g., "VALUES_COL_1_" & tabId resolves to the actual Excel named range).
' @param plotType The chart series type string: "bar" for vertical bars or "hbar" for
'                 horizontal bars. Determines both the primary series type and whether
'                 percentage points are added.
' @param hasPerc Whether the table specification includes a percentage column. When True
'                and plotType is "bar", an additional PERC_COL_1_ point series is pushed
'                on the right axis.
' @remark The percentage overlay is intentionally skipped for horizontal bars because
'         Excel point series render poorly on horizontally-oriented charts.
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

' @description Builds chart series for multi-column cross-tables (bivariate, spatial, or
'              spatio-temporal). The graph mode is determined by the "graph" setting in the
'              table specs, which can be one of three values: (1) "values" or "yes" -- pushes
'              one bar series per column referencing VALUES_COL_<n>_ named ranges on the left
'              axis; (2) "percentages" -- pushes one bar series per column referencing
'              PERC_COL_<n>_ named ranges (only if HasPercentage is True); (3) "both" --
'              pushes value bars on the left axis AND percentage points on the right axis for
'              each column, creating a dual-axis overlay chart. For spatio-temporal tables,
'              the row categories use the section-level range (ROW_CATEGORIES_ + tabSecId)
'              instead of the table-level range, because spatio-temporal tables share row
'              categories across all tables in the same section.
' @param specs The ITableSpecs for the cross-table, providing access to the "graph" and
'              "flip" settings as well as HasPercentage.
' @param tabId The unique table identifier used as a suffix for named range construction.
' @param tabSecId The section-level table identifier, used only for spatio-temporal tables
'                 to build section-wide ROW_CATEGORIES_ range names. Empty for other types.
' @param tabType The numeric table type (TypeBivariate, TypeSpatial, or TypeSpatioTemporal)
'                used to determine whether section-level row categories should be applied.
' @param nbCols The number of data columns to iterate over. For spatio-temporal tables this
'               comes from the "n geo" setting; for others it comes from Table.NumberOfColumns.
' @param plotType The chart series type string: "bar" for vertical bars or "hbar" for
'                 horizontal bars, derived from the "flip" setting.
' @remark When graphValue is "both", percentage point series are only added if
'         specs.HasPercentage is True. The "percentages" mode silently produces no series
'         if HasPercentage is False, which may be surprising.
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

' @description Builds the complete set of graph specifications for complex (time series)
'              mode by iterating over all unique graph IDs. First calls BuildGraphIdsList to
'              extract the deduplicated list of graph IDs from the graph setup ListObject,
'              then stores a clone of that list in the instance state for later retrieval via
'              the GraphIDs property. For each graph ID, it calls DefineGraphSpecs to populate
'              the series arrays with that graph's series definitions, and then PushGraph to
'              package those arrays into a 7-element BetterArray entry that gets appended to
'              the graphsChartSpecs collection. The resulting collection can be retrieved via
'              the SpecsLists property after this method completes.
' @remark Each iteration through a graph ID starts by clearing the previous graph's series
'         data via ClearSeries (called inside DefineGraphSpecs), so the series arrays are
'         effectively temporary buffers that get packaged and reset for each graph.
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

' @description Defines the chart series for a single graph ID in complex mode by reading
'              from the setup ListObjects. First, it reads all column values for the given
'              graphId from the graph ListObject using GraphValues: series IDs, axis positions,
'              percentage overrides, chart types, choice values, and display labels. Then it
'              clears any previously buffered series data via ClearSeries. For each series ID,
'              it looks up the corresponding row in the time series ListObject to obtain a
'              TableSpecs instance, which provides the table ID and section ID. The choice
'              value is resolved to a named range name via TimeSeriesColumnName, and then to
'              the actual data range name via ResolveSeriesRangeName. If the "percentages"
'              column indicates a percentage override, the resolved range name is replaced
'              with its PERC_ equivalent. Finally, each valid series is pushed to the series
'              arrays via PushSeries and PushLabels, using the section-level ROW_CATEGORIES_
'              range for row labels.
' @param graphId The unique graph identifier string used to filter rows from the graph
'                ListObject. All series belonging to this graph will be extracted and
'                buffered in the series arrays.
' @remark Series IDs that cannot be found in the time series ListObject (or that resolve
'         to the header row) are silently skipped via GoTo ContinueLoop. This allows
'         partial graph definitions where some series may reference deleted or renamed
'         time series entries.
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

' @description Resolves a choice value and label name into the correct data range name for
'              chart series values. The mapping follows three rules: (1) If the choice value
'              is "Total", the range name is TOTAL_COL_VALUES_ + tabId, pointing to the
'              aggregated total column for the table. (2) If the label name contains the
'              prefix "COLUMN_CATEGORIES_", the range name is INTERIOR_VALUES_ + tabId,
'              pointing to the interior data area of the cross-table. (3) For all other
'              cases, the LABEL prefix in the label name is replaced with VALUES to produce
'              the corresponding data range name (e.g., "LABEL_COL_2_xyz" becomes
'              "VALUES_COL_2_xyz"). This three-way dispatch ensures that each series type
'              (total, interior, or individual column) maps to the correct named range.
' @param choiValue The choice value string from the graph setup table. "Total" triggers
'                  special handling; other values use the label-based resolution logic.
' @param tabId The unique table identifier used as a suffix for the TOTAL_COL_VALUES_ and
'              INTERIOR_VALUES_ range name patterns.
' @param labName The label range name previously resolved by TimeSeriesColumnName. Used to
'                detect the COLUMN_CATEGORIES_ prefix and as the base for LABEL-to-VALUES
'                string replacement.
' @return The fully-qualified named range name string pointing to the data values for
'         this series, ready to be used as a chart series source.
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

' @description Finds the Excel named range that corresponds to a given choice value within
'              a time series table. The choice value represents a category label (e.g., a
'              specific geographic area or demographic group) that appears in the
'              COLUMN_CATEGORIES_ range of the table on the output worksheet. If the choice
'              value is "Total", it returns the predefined TOTAL_LABEL_COL_ + tabId range
'              name directly without searching. For all other values, it searches the
'              COLUMN_CATEGORIES_ range for an exact match and returns the .Name.Name of the
'              found cell, which is the workbook-level named range assigned to that column
'              header. Returns an empty string if the choice value is empty or not found.
' @param choiValue The category value to look up (e.g., "Male", "Region A", or "Total").
'                  An empty string causes an early exit returning vbNullString.
' @param tabId The unique table identifier used to construct the COLUMN_CATEGORIES_ and
'              TOTAL_LABEL_COL_ range name patterns.
' @return The full named range name string (e.g., "LABEL_COL_3_ts_table_1") for the
'         matching column, or vbNullString if no match is found.
' @remark Uses On Error Resume Next when reading cellRng.Name.Name because some cells in
'         the COLUMN_CATEGORIES_ range may not have a named range assigned. In that case,
'         colName remains empty and the function returns vbNullString.
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

' @description Creates six fresh BetterArray instances (all with LowerBound=1) to store
'              the parallel series data arrays: seriesNames, seriesTypes, seriesPositions,
'              seriesLabels, seriesColumnLabels, and seriesLabelPrefixes. These six arrays
'              form matched parallel arrays where index N across all six describes one
'              complete series entry. LowerBound is set to 1 (not the default 0) to align
'              with the 1-based indexing used throughout the GraphSpecs class and the
'              consumer code that iterates from 1 to NumberOfSeries.
' @remark This method is called at the start of BuildSimpleSeries and lazily by PushSeries
'         and PushLabels when the arrays are Nothing (which happens after ClearSeries in
'         complex mode). It unconditionally replaces any existing arrays, so calling it
'         when arrays already contain data will discard that data.
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

' @description Appends one series entry to the three series-related parallel arrays
'              (seriesNames, seriesTypes, seriesPositions). The series name is typically an
'              Excel named range reference (e.g., "VALUES_COL_1_table1") that the chart
'              builder will use as the data source. The type is the chart series rendering
'              style (e.g., "bar", "hbar", "point", "line"). The position is the axis
'              placement ("left" for primary axis, "right" for secondary axis). If the
'              seriesNames array is Nothing (which occurs in complex mode after ClearSeries
'              resets the arrays between graph IDs), this method lazy-initializes all six
'              arrays via InitSeriesArrays before pushing.
' @param chrtName The named range name string identifying the data source for this series
'                 (e.g., "VALUES_COL_1_ts_table_2" or "PERC_COL_3_bivar_table_1").
' @param chrtType The chart type string that controls how the series is rendered. Common
'                 values: "bar" (vertical bars), "hbar" (horizontal bars), "point" (scatter
'                 points, typically for percentage overlay), "line" (line series).
' @param chrtPos The axis position string: "left" places the series on the primary (left)
'                value axis, "right" places it on the secondary (right) value axis.
' @remark Always call PushLabels immediately after PushSeries to keep the parallel arrays
'         synchronized. The series arrays and label arrays must maintain a 1:1 index mapping.
Private Sub PushSeries(ByVal chrtName As String, _
                        ByVal chrtType As String, _
                        ByVal chrtPos As String)
    ' Lazy init for complex mode (called per graph ID after ClearSeries)
    If this.seriesNames Is Nothing Then InitSeriesArrays

    this.seriesNames.Push chrtName
    this.seriesTypes.Push chrtType
    this.seriesPositions.Push chrtPos
End Sub

' @description Appends one label entry to the three label-related parallel arrays
'              (seriesLabels, seriesColumnLabels, seriesLabelPrefixes). The row category is
'              the named range containing the category axis labels (e.g., "ROW_CATEGORIES_sec1").
'              The column label is the named range for the series legend entry (e.g.,
'              "LABEL_COL_2_table1"). The optional prefix is a display string prepended to
'              the series label in the chart legend, used in complex mode where a user-defined
'              label from the graph setup table overrides or augments the default column label.
'              If the seriesLabels array is Nothing, this method lazy-initializes all six
'              arrays via InitSeriesArrays before pushing.
' @param chrtRowCategory The named range name for the row (category axis) labels. In simple
'                        mode this is "ROW_CATEGORIES_" + tabId; in complex mode it uses the
'                        section ID: "ROW_CATEGORIES_" + secId.
' @param chrtColumnLab The named range name for the column (legend) label of this series.
'                      For example, "LABEL_COL_1_table1" or "TOTAL_LABEL_COL_ts_table_2".
' @param chrtLabPref Optional display prefix for the series label. Defaults to vbNullString.
'                    Used in complex mode to pass the user-defined "label" value from the
'                    graph setup ListObject for display in the chart legend.
Private Sub PushLabels(ByVal chrtRowCategory As String, _
                        ByVal chrtColumnLab As String, _
                        Optional ByVal chrtLabPref As String = vbNullString)
    If this.seriesLabels Is Nothing Then InitSeriesArrays

    this.seriesLabels.Push chrtRowCategory
    this.seriesColumnLabels.Push chrtColumnLab
    this.seriesLabelPrefixes.Push chrtLabPref
End Sub

' @description Clears all six parallel series/label arrays by calling .Clear on each
'              BetterArray instance. This is called between complex mode graph IDs inside
'              DefineGraphSpecs to reset the series buffers before populating them with the
'              next graph's series definitions. After clearing, the arrays remain as
'              instantiated (non-Nothing) BetterArray objects with zero length, but
'              PushSeries and PushLabels check for Nothing and will lazy-reinitialize if
'              needed. Uses On Error Resume Next because during the first call the arrays
'              may still be Nothing (before any series have been pushed), and calling .Clear
'              on Nothing would raise an error.
' @remark The On Error Resume Next / On Error GoTo 0 wrapper is intentional and safe here:
'         the only possible error is calling .Clear on a Nothing reference, which is a
'         benign no-op that should be silently ignored.
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

' @description Packages the currently-buffered series data into a single graph specification
'              entry and appends it to the graphsChartSpecs collection. Each entry is a
'              BetterArray containing exactly 7 elements: (1) seriesNames clone -- the named
'              range references for each series data source; (2) seriesTypes clone -- the
'              chart type for each series (bar, line, point, etc.); (3) seriesPositions clone
'              -- the axis position for each series (left or right); (4) seriesLabels clone --
'              the row category named ranges; (5) seriesColumnLabels clone -- the legend label
'              named ranges; (6) seriesLabelPrefixes clone -- user-defined label prefixes;
'              (7) the graph title string retrieved from GraphTitle. All BetterArray elements
'              are cloned to ensure the packaged data is independent of the mutable series
'              buffers that get cleared for the next graph. The graphsChartSpecs collection
'              is lazy-initialized on the first call (LowerBound=1).
' @param graphId The unique graph identifier string, used to look up the graph title from
'                the titles ListObject via GraphTitle.
' @remark BetterArray.Clone() has a known quirk where it can sometimes produce an array
'         with one extra element. The safety check compares the cloned length to the
'         original and calls .Pop on all six clones if they are one element too long.
'         This ensures the packaged data exactly matches the original series buffers.
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

' @description Returns the first ListObject in the LoList array, which is the graph setup
'              table. This ListObject defines which series belong to each graph ID, including
'              their chart types, axis positions, choice values, percentage overrides, and
'              display labels. Its rows are keyed by "graph id" and "series id" columns.
' @return The ListObject at index LoList.LowerBound (position 0 in the BetterArray),
'         representing the graph setup table ("graph on time series").
Private Property Get GraphListObject() As ListObject
    Set GraphListObject = LoList.Item(LoList.LowerBound)
End Property

' @description Returns the second ListObject in the LoList array, which is the time series
'              setup table. This ListObject maps each series ID to its full table
'              specification row (variable name, section, time unit, etc.), enabling
'              DefineGraphSpecs to construct a TableSpecs instance for each series. Its
'              first column contains the series IDs that match the "series id" values in the
'              graph ListObject.
' @return The ListObject at index LoList.LowerBound + 1 (position 1 in the BetterArray),
'         representing the time series definition table ("time series analysis").
Private Property Get TimeSeriesListObject() As ListObject
    Set TimeSeriesListObject = LoList.Item(LoList.LowerBound + 1)
End Property

' @description Returns the third ListObject in the LoList array, which is the title labels
'              table. This ListObject provides display titles for each graph, keyed by graph
'              ID in its third column and the title text in its first column. It is used by
'              GraphTitle to look up the human-readable plot title for each graph entry.
' @return The ListObject at index LoList.LowerBound + 2 (position 2 in the BetterArray),
'         representing the title labels table ("labels for time series graphs").
Private Property Get TitlesListObject() As ListObject
    Set TitlesListObject = LoList.Item(LoList.LowerBound + 2)
End Property

' @description Looks up the human-readable plot title for a given graph ID from the titles
'              ListObject. The titles table has the graph ID in column 3 and the display
'              title in column 1 (offset -2 from the found cell). Searches the third column's
'              DataBodyRange for an exact, case-sensitive match of the graphId string. If a
'              match is found, returns the title text from the same row's first column. If no
'              match is found, returns an empty string, which means the graph will have no
'              title when rendered.
' @param graphId The unique graph identifier string to look up in the titles ListObject's
'                third column.
' @return The display title string for the graph, or an empty string if the graphId is not
'         found in the titles table.
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

' @description Extracts all unique graph IDs from the "graph id" column of the graph setup
'              ListObject and returns them as a BetterArray with LowerBound=1. Iterates
'              through every data row in the graph ID column (skipping the header at row 1),
'              checks whether each cell is non-empty and not already present in the
'              accumulator array (using BetterArray.Includes for duplicate detection), and
'              pushes new unique values. The resulting list preserves the order of first
'              appearance, which determines the order in which graphs are built and displayed
'              in the output.
' @return A cloned BetterArray of unique graph ID strings, with LowerBound=1. The clone
'         ensures the caller gets an independent copy that cannot be mutated from outside.
' @remark The duplicate check uses BetterArray.Includes which performs a linear scan, so
'         performance is O(n*m) where n is the total row count and m is the number of
'         unique IDs. This is acceptable because the graph setup table typically has fewer
'         than 100 rows.
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

' @description Returns all values from a specified column in the graph setup ListObject,
'              filtered to only include rows whose "graph id" matches the given graphId.
'              This is the primary data extraction method for DefineGraphSpecs, called once
'              per column (series id, axis, percentages, type, choices, label) to gather all
'              settings for a single graph. The method walks through the "graph id" column
'              progressively: after finding each matching row, it narrows the search range
'              to start after the found row, preventing the same row from being matched
'              twice. Empty cell values are replaced with "&" as a placeholder to maintain
'              array index alignment with other parallel column extractions. If the graphId
'              is not found at all, an empty BetterArray is returned immediately.
' @param colName The header name of the column to extract values from (e.g., "series id",
'                "axis", "type", "choices", "label", "percentages"). Looked up via
'                LoColumnIndex using partial, case-insensitive matching.
' @param graphId The graph ID value to filter rows by. Only rows where the "graph id"
'                column exactly matches this value (case-sensitive, whole cell) are included.
' @return A cloned BetterArray (LowerBound=1) containing the extracted cell values as
'         strings, in row order. Empty cells are represented as "&".
' @remark The progressive range narrowing technique (shrinking rng and colRng after each
'         find) is necessary because Excel's Range.Find wraps around by default, which
'         would cause infinite loops if the same range were searched repeatedly. The method
'         exits the loop when the remaining range has only one row or when Find returns
'         Nothing.
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

' @description Performs a 1-based, relative column index lookup in a ListObject's header row.
'              Searches the HeaderRowRange for a partial, case-insensitive match of the given
'              column name using Range.Find with LookAt:=xlPart. The returned index is
'              relative to the ListObject (not the worksheet), calculated as the found cell's
'              absolute column minus the header range's starting column plus 1. This relative
'              index can be used directly with lo.ListColumns(index) to access the column.
'              Returns -1 if the column name is not found in the header.
' @param lo The ListObject whose header row will be searched.
' @param colName The column header text to search for. Uses partial matching (xlPart) and
'                is case-insensitive, so "graph id" will match a header like "Graph ID".
' @return The 1-based relative column index within the ListObject, or -1 if not found.
' @remark Uses xlPart matching intentionally to allow flexible header naming (e.g.,
'         searching for "series id" will match "Series ID" or "series id (unique)"). Be
'         aware that partial matching could return unexpected results if multiple headers
'         share a common substring.
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

' @description Validates the BetterArray of ListObjects required by CreateRangeSpecs,
'              performing three levels of validation: (1) Count validation -- the array must
'              contain exactly 3 elements, corresponding to the graph setup, time series
'              setup, and title labels tables. (2) Type and content validation -- each element
'              must be a non-Nothing ListObject with a non-Nothing DataBodyRange (i.e., it
'              must contain at least one data row). (3) Order validation -- the ListObjects
'              must appear in the correct order, verified by reading the title label from 2
'              rows above each ListObject's header (at Cells(-1, 1)) and comparing against
'              the module-level constants TAB_ON_GRAPH_TIME_SERIES, TAB_ON_TIME_SERIES, and
'              TAB_ON_TITLE_GRAPHS_TIME_SERIES. If any validation fails, an appropriate error
'              is raised with a descriptive message.
' @param loTable The BetterArray of ListObjects to validate. Must contain exactly 3 entries
'                in the order: graph setup table, time series table, title labels table.
' @remark The order validation reads from Cells(-1, 1) relative to the HeaderRowRange,
'         which is 2 rows above the header. This assumes a specific worksheet layout where
'         the table title is placed in that position. If the worksheet layout changes, this
'         validation will break with an ErrorUnexpectedState error.
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

' @description Returns the Worksheet where charts will be drawn, dispatching based on the
'              current graph mode. In simple mode (GraphTypeSimple), returns Table.Wksh --
'              the worksheet that contains the cross-table's named ranges, which is where
'              the chart will be embedded alongside its source data. In complex mode
'              (GraphTypeComplex), returns OutputSheet -- a dedicated output worksheet
'              provided during factory construction, separate from the setup ListObjects'
'              worksheet. This abstraction allows consumers to call Wksh without needing
'              to know which mode the GraphSpecs instance was created in.
' @return The Worksheet reference where charts should be created for this GraphSpecs
'         instance.
Public Property Get Wksh() As Worksheet
    Select Case GraphType
    Case GraphTypeSimple
        Set Wksh = Table.Wksh
    Case GraphTypeComplex
        Set Wksh = OutputSheet
    End Select
End Property

' @description Returns the number of chart series defined in simple mode. In simple mode,
'              this is the length of the seriesNames BetterArray after CreateSeries has been
'              called. In complex mode (GraphTypeComplex), this always returns 0 because
'              complex mode stores series data within the nested graphsChartSpecs collection
'              rather than in the flat series arrays. Also returns 0 if the seriesNames array
'              is Nothing (i.e., CreateSeries has not yet been called in simple mode).
' @return The count of series entries in simple mode (1-based), or 0 for complex mode
'         or before CreateSeries has been invoked.
Public Property Get NumberOfSeries() As Long
    If (this.seriesNames Is Nothing) Or (GraphType = GraphTypeComplex) Then
        NumberOfSeries = 0
        Exit Property
    End If
    NumberOfSeries = this.seriesNames.Length
End Property

' @description Returns the number of distinct graphs defined in complex mode. Each graph
'              corresponds to one unique graph ID from the graph setup ListObject, and each
'              has its own entry in the graphsChartSpecs collection containing all of its
'              series definitions. Returns 0 in simple mode (GraphTypeSimple) because simple
'              mode produces a single implicit graph with no collection. Also returns 0 if
'              graphsChartSpecs is Nothing (i.e., CreateSeries has not been called or
'              produced no valid graphs).
' @return The count of graph entries in the graphsChartSpecs collection for complex mode,
'         or 0 for simple mode or before series building.
Public Property Get NumberOfGraphs() As Long
    If (this.graphsChartSpecs Is Nothing) Or (GraphType <> GraphTypeComplex) Then
        NumberOfGraphs = 0
        Exit Property
    End If
    NumberOfGraphs = this.graphsChartSpecs.Length
End Property

' @description Returns the named range reference string for the series at the given 1-based
'              index. This is the Excel named range (e.g., "VALUES_COL_1_table1") that the
'              chart builder uses as the data source for this series. Only valid in simple
'              mode after CreateSeries has been called.
' @param index The 1-based index of the series to retrieve (1 to NumberOfSeries).
' @return The named range name string for the series data source.
Public Property Get SeriesName(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesName = this.seriesNames.Item(index)
End Property

' @description Returns the chart type string for the series at the given 1-based index.
'              Common values include "bar" (vertical bars), "hbar" (horizontal bars), "point"
'              (scatter points for percentage overlays), and "line". Only valid in simple mode
'              after CreateSeries has been called.
' @param index The 1-based index of the series to retrieve (1 to NumberOfSeries).
' @return The chart type string controlling how this series is rendered.
Public Property Get SeriesType(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesType = this.seriesTypes.Item(index)
End Property

' @description Returns the axis position string for the series at the given 1-based index.
'              "left" places the series on the primary (left) value axis; "right" places it
'              on the secondary (right) value axis, typically used for percentage overlays.
'              Only valid in simple mode after CreateSeries has been called.
' @param index The 1-based index of the series to retrieve (1 to NumberOfSeries).
' @return The axis position string ("left" or "right") for this series.
Public Property Get SeriesPos(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesPos = this.seriesPositions.Item(index)
End Property

' @description Returns the row category (axis label) named range for the series at the
'              given 1-based index. This named range (e.g., "ROW_CATEGORIES_table1") provides
'              the category axis labels displayed along the chart's horizontal (or vertical
'              for hbar) axis. Only valid in simple mode after CreateSeries has been called.
' @param index The 1-based index of the series to retrieve (1 to NumberOfSeries).
' @return The named range name string for the row category labels of this series.
Public Property Get SeriesLabel(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesLabel = this.seriesLabels.Item(index)
End Property

' @description Returns the column label (legend entry) named range for the series at the
'              given 1-based index. This named range (e.g., "LABEL_COL_2_table1") provides
'              the text displayed in the chart legend to identify this series. Only valid
'              in simple mode after CreateSeries has been called.
' @param index The 1-based index of the series to retrieve (1 to NumberOfSeries).
' @return The named range name string for the column label (legend entry) of this series.
Public Property Get SeriesColumnLabel(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesColumnLabel = this.seriesColumnLabels.Item(index)
End Property

' @description Returns the optional label prefix string for the series at the given 1-based
'              index. In simple mode this is always vbNullString. In complex mode (when
'              accessed via SpecsLists), this contains the user-defined "label" value from
'              the graph setup ListObject, which is prepended to the series name in the chart
'              legend. Only valid in simple mode after CreateSeries has been called.
' @param index The 1-based index of the series to retrieve (1 to NumberOfSeries).
' @return The label prefix string for this series, or vbNullString if none was set.
Public Property Get SeriesLabelPrefix(ByVal index As Long) As String
    ValidateSeriesIndex index
    SeriesLabelPrefix = this.seriesLabelPrefixes.Item(index)
End Property

' @description Returns whether this GraphSpecs instance has valid, successfully-built series
'              data. The result is cached after the first evaluation using the validityTested
'              and validityResult flags, so subsequent calls return instantly without
'              rebuilding. On the first call, it invokes CreateSeries to build all series
'              data, then determines validity: in simple mode (GraphTypeSimple), the result
'              is always True because BuildSimpleSeries always succeeds if the factory
'              constructor accepted the table. In complex mode (GraphTypeComplex), validity
'              is True only if the graphsChartSpecs collection was successfully created
'              (i.e., at least one graph ID produced valid series definitions). A Nothing
'              graphsChartSpecs means no valid graphs were found, resulting in False.
' @return True if series data was successfully built and is available for consumption;
'         False if complex mode produced no valid graph entries.
' @remark Because this property calls CreateSeries on first access, it can be used as a
'         validity gate without separately calling CreateSeries. However, be aware that
'         the caching means errors during CreateSeries will permanently mark the instance
'         as invalid (or valid) until it is discarded and a new one is created.
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

' @description Returns a cloned copy of the graphsChartSpecs collection containing all
'              graph specification entries built during complex mode. Each element in the
'              returned BetterArray is itself a 7-element BetterArray representing one
'              complete graph: [seriesNames, seriesTypes, seriesPositions, seriesLabels,
'              seriesColumnLabels, seriesLabelPrefixes, titleString]. The clone ensures that
'              the consumer cannot accidentally mutate the internal state. If graphsChartSpecs
'              is Nothing (either because CreateSeries has not been called or because no valid
'              graphs were produced), returns a fresh empty BetterArray.
' @return A cloned BetterArray of graph spec entries for complex mode, or an empty
'         BetterArray if no specs are available.
' @remark This property is only meaningful in complex mode (GraphTypeComplex). In simple
'         mode, the graphsChartSpecs collection is never populated, so this will always
'         return an empty BetterArray. Use the indexed SeriesName/SeriesType/etc. accessors
'         for simple mode data instead.
Public Property Get SpecsLists() As BetterArray
    If this.graphsChartSpecs Is Nothing Then
        Set SpecsLists = New BetterArray
    Else
        Set SpecsLists = this.graphsChartSpecs.Clone()
    End If
End Property

' @description Returns a cloned list of all unique graph IDs for complex mode. The list is
'              lazily built on first access by calling BuildGraphIdsList if the internal
'              graphIdsList is Nothing. If CreateSeries has already been called, the cached
'              list from that build is reused instead of rebuilding. The returned BetterArray
'              is always a clone, ensuring that callers cannot modify the internal list. Each
'              element is a string graph ID in the order of first appearance in the graph
'              setup ListObject.
' @return A cloned BetterArray of unique graph ID strings with LowerBound=1.
' @remark This property is primarily useful for complex mode consumers that need to iterate
'         over graph IDs independently of SpecsLists (e.g., for UI display or filtering).
'         In simple mode, the graph setup ListObject is not available, so calling this
'         property will raise an error when it attempts to access GraphListObject.
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
