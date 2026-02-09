VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnaTabIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Export tracking for analysis tables and graphs"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Export tracking for analysis tables and graphs")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IAnaTabIds

Private Type TAnaIds
    sh As Worksheet
End Type

Private this As TAnaIds
Private Const CLASS_NAME As String = "AnaTabIds"

' Named ranges for sheet names
Private Const RNGUASHEET As String = "RNG_SheetUAName"
Private Const RNGTSSHEET As String = "RNG_SheetTSName"
Private Const RNGSPSHEET As String = "RNG_SheetSPName"
Private Const RNGSPTEMPSHEET As String = "RNG_SheetSPTempName"

'@section Factory
'===============================================================================

' @description Create a new AnaTabIds instance bound to a tracking worksheet. AnaTabIds
'   is responsible for recording named ranges and graph series metadata into
'   ListObjects on a hidden tracking sheet, so that the full analysis structure
'   (tables and charts) can be faithfully recreated on an export worksheet later.
'   When the check parameter is True (the default), the factory validates that all
'   12 tracking ListObjects (3 prefixes x 4 scopes) and 4 sheet-name named ranges
'   exist on the worksheet before creating the instance, raising an error if any
'   are missing.
' @param idsh The worksheet that contains the tracking ListObjects and sheet-name
'   named ranges. This sheet acts as the persistent store for all analysis metadata.
' @param check When True, calls CheckRequirements to validate that every expected
'   ListObject and named range is present on idsh. Pass False to skip validation
'   (e.g., when the tracking sheet is being built incrementally).
' @return A new IAnaTabIds interface reference backed by this AnaTabIds instance.
Public Function Create(ByVal idsh As Worksheet, _
                       Optional ByVal check As Boolean = True) As IAnaTabIds

    If idsh Is Nothing Then ThrowError InvalidArgument, "Worksheet not defined"
    If check Then CheckRequirements idsh

    With New AnaTabIds
        Set .Wksh = idsh
        Set Create = .Self
    End With
End Function

'@section Internal properties (factory use)
'===============================================================================

Public Property Get Self() As IAnaTabIds
    Set Self = Me
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@section ListObject access
'===============================================================================

' @description Resolve and return the tracking ListObject for a given analysis scope
'   and prefix. The scope is mapped to a suffix string: AnalysisIdsScopeTimeSeries
'   maps to "ts", AnalysisIdsScopeSpatial to "sp", AnalysisIdsScopeNormal to "uba",
'   and AnalysisIdsScopeSpatioTemporal to "sptemp". The final ListObject name is
'   the prefix concatenated with this suffix (e.g., "tab_ids_ts", "graph_ids_sp",
'   "graph_formats_uba"). This property is the central routing mechanism that all
'   tracking methods use to locate the correct ListObject on the tracking worksheet.
' @param scope A Byte value from the AnalysisIdsScope enum indicating which analysis
'   type to target: TimeSeries, Spatial, Normal, or SpatioTemporal.
' @param prefix The ListObject name prefix. Defaults to "tab_ids_" for table tracking.
'   Pass "graph_ids_" for graph series tracking or "graph_formats_" for graph format
'   tracking.
' @return The ListObject on the tracking worksheet matching the resolved name.
' @remark If the scope value does not match any known case, actName remains empty and
'   the ListObject lookup will fail with a runtime error.
Private Property Get ActualLo(ByVal scope As Byte, _
                              Optional ByVal prefix As String = "tab_ids_") As ListObject
    Dim actName As String

    Select Case scope
    Case AnalysisIdsScopeTimeSeries
        actName = "ts"
    Case AnalysisIdsScopeSpatial
        actName = "sp"
    Case AnalysisIdsScopeNormal
        actName = "uba"
    Case AnalysisIdsScopeSpatioTemporal
        actName = "sptemp"
    End Select

    Set ActualLo = Wksh().ListObjects(prefix & actName)
End Property

' @description Resolve the analysis output worksheet for a given scope by reading the
'   corresponding sheet-name named range from the tracking worksheet. Each scope maps
'   to a different named range: AnalysisIdsScopeNormal reads RNG_SheetUAName,
'   AnalysisIdsScopeTimeSeries reads RNG_SheetTSName, AnalysisIdsScopeSpatial reads
'   RNG_SheetSPName, and AnalysisIdsScopeSpatioTemporal reads RNG_SheetSPTempName.
'   The value stored in that named range is the actual worksheet name within the same
'   workbook. This indirection allows the analysis sheet names to be configured
'   dynamically rather than hard-coded.
' @param scope A Byte value from the AnalysisIdsScope enum indicating which analysis
'   output sheet to resolve.
' @return The Worksheet object from the tracking sheet's parent workbook whose name
'   matches the value stored in the scope-specific named range.
' @remark If the named range value does not correspond to an existing worksheet, a
'   runtime error will be raised by the Worksheets collection access.
Private Property Get ActualAnalysisSheet(ByVal scope As Byte) As Worksheet
    Dim sheetName As String
    Dim sh As Worksheet

    Set sh = Wksh()

    Select Case scope
    Case AnalysisIdsScopeNormal
        sheetName = sh.Range(RNGUASHEET).Value
    Case AnalysisIdsScopeTimeSeries
        sheetName = sh.Range(RNGTSSHEET).Value
    Case AnalysisIdsScopeSpatial
        sheetName = sh.Range(RNGSPSHEET).Value
    Case AnalysisIdsScopeSpatioTemporal
        sheetName = sh.Range(RNGSPTEMPSHEET).Value
    End Select

    Set ActualAnalysisSheet = sh.Parent.Worksheets(sheetName)
End Property

'@section Table tracking
'===============================================================================

' @description Write table named-range entries into the scope-specific tracking
'   ListObject (prefixed "tab_ids_") so they can be transferred to an export
'   worksheet later. Each entry in the ListObject is a row with three columns:
'   the table identifier (tabId) in column 1, the named range name in column 2,
'   and the export flag ("yes"/"no") in column 3. The named range names are
'   provided as a BetterArray and are bulk-written via ToExcelRange, while
'   tabId and the export flag are written row-by-row in a loop. The method
'   handles ListObject resizing: if the table body is empty (first data row is
'   blank), it writes starting from row 1; otherwise it appends after the last
'   existing row. After writing, the ListObject is resized to encompass all rows
'   including the new entries.
' @param scope A Byte value from the AnalysisIdsScope enum selecting the correct
'   tracking ListObject (e.g., tab_ids_ts for TimeSeries).
' @param tabId A string identifier for the analysis table. All named ranges in
'   tabRangesNames will be tagged with this id in column 1.
' @param tabRangesNames A BetterArray of named range name strings to record. If
'   empty (Length = 0), the method exits immediately without modifying the ListObject.
' @param exportTable A flag string written to column 3 of each row. Defaults to
'   "yes", indicating the range should be included during export. Pass "no" to
'   record the range but skip it during export.
' @remark The BetterArray.ToExcelRange call writes all range names in a single
'   operation to column 2, which is more efficient than cell-by-cell assignment.
'   The ListObject is resized at the end to include all newly appended rows.
Private Sub AddTableInfos(ByVal scope As Byte, _
                           ByVal tabId As String, _
                           ByVal tabRangesNames As BetterArray, _
                           Optional ByVal exportTable As String = "yes")
    Dim lo As ListObject
    Dim loRng As Range
    Dim cellRng As Range
    Dim counter As Long
    Dim sh As Worksheet

    If tabRangesNames.Length = 0 Then Exit Sub

    Set lo = ActualLo(scope)
    Set loRng = lo.Range
    Set sh = Wksh()

    If loRng.Cells(2, 1).Value = vbNullString Then
        Set cellRng = loRng.Cells(1, 1)
    Else
        Set cellRng = loRng.Cells(loRng.Rows.Count, 1)
    End If

    tabRangesNames.ToExcelRange cellRng.Offset(1, 1)

    For counter = 1 To tabRangesNames.Length
        cellRng.Offset(counter).Value = tabId
        cellRng.Offset(counter, 2).Value = exportTable
    Next

    Set loRng = sh.Range(loRng.Cells(1, 1), cellRng.Offset(counter - 1, 2))
    lo.Resize loRng
End Sub

'@section Graph tracking
'===============================================================================

' @description Write one graph series row into the scope-specific graph tracking
'   ListObject (prefixed "graph_ids_"). Each row captures all the metadata needed
'   to recreate a single chart series during export. The 11 columns are laid out
'   as follows: column 1 = tabId (groups series by parent table), column 2 =
'   graphId (groups series into a single chart), column 3 = seriesName (the named
'   range for the series data), column 4 = seriesType (chart type string, e.g.,
'   "xlLine"), column 5 = seriesPos (primary/secondary axis), column 6 =
'   seriesLabel (category label range name), column 7 = seriesColumnLabel (column
'   header label range name), column 8 = hardCodeLabels (Boolean as Long, 0/1),
'   column 9 = outRangeAddress (the cell address where the graph should be placed
'   on the output sheet), column 10 = prefix (optional label prefix string),
'   column 11 = prefixOnly (Boolean as Long, whether to use only the prefix as
'   the label). The method appends to the end of the ListObject and resizes it.
' @param scope A Byte value from the AnalysisIdsScope enum selecting the correct
'   graph tracking ListObject (e.g., graph_ids_sp for Spatial).
' @param tabId The parent table identifier that owns this graph.
' @param graphId An identifier for the graph. Multiple series rows sharing the
'   same graphId will be grouped into a single chart during export.
' @param seriesName The named range name pointing to the data for this series.
' @param seriesType A string representing the Excel chart type (e.g., "xlLine",
'   "xlColumnClustered").
' @param seriesPos Axis position indicator for the series (e.g., primary or
'   secondary axis).
' @param seriesLabel The named range name for the category (X-axis) labels.
' @param seriesColumnLabel The named range name for the column header labels.
' @param hardCodeLabels When True, labels are written as literal values rather
'   than range references. Stored as CLng (0 or 1) in the ListObject.
' @param outRangeAddress The cell address string (e.g., "A10") where the graph
'   should be placed on the output worksheet during export.
' @param prefix An optional string prefix to prepend to series labels.
' @param prefixOnly When True, only the prefix is used as the label, ignoring
'   the actual series label value. Stored as CLng (0 or 1) in the ListObject.
' @remark Boolean parameters are stored as Long (0/1) via CLng() because Excel
'   ListObject cells handle numeric values more reliably than VBA Boolean for
'   cross-session persistence.
Private Sub AddGraphInfo(ByVal scope As Byte, _
                          ByVal tabId As String, _
                          ByVal graphId As String, _
                          ByVal seriesName As String, _
                          ByVal seriesType As String, _
                          ByVal seriesPos As String, _
                          ByVal seriesLabel As String, _
                          ByVal seriesColumnLabel As String, _
                          ByVal hardCodeLabels As Boolean, _
                          ByVal outRangeAddress As String, _
                          Optional ByVal prefix As String = vbNullString, _
                          Optional ByVal prefixOnly As Boolean = False)
    Dim lo As ListObject
    Dim loRng As Range
    Dim cellRng As Range
    Dim sh As Worksheet

    Set lo = ActualLo(scope, prefix:="graph_ids_")
    Set loRng = lo.Range
    Set sh = Wksh()

    If loRng.Cells(2, 1).Value = vbNullString Then
        Set cellRng = loRng.Cells(2, 1)
    Else
        Set cellRng = loRng.Cells(loRng.Rows.Count + 1, 1)
    End If

    cellRng.Value = tabId
    cellRng.Cells(1, 2).Value = graphId
    cellRng.Cells(1, 3).Value = seriesName
    cellRng.Cells(1, 4).Value = seriesType
    cellRng.Cells(1, 5).Value = seriesPos
    cellRng.Cells(1, 6).Value = seriesLabel
    cellRng.Cells(1, 7).Value = seriesColumnLabel
    cellRng.Cells(1, 8).Value = CLng(hardCodeLabels)
    cellRng.Cells(1, 9).Value = outRangeAddress
    cellRng.Cells(1, 10).Value = prefix
    cellRng.Cells(1, 11).Value = CLng(prefixOnly)

    Set loRng = sh.Range(loRng.Cells(1, 1), cellRng.Cells(1, 11))
    lo.Resize loRng
End Sub

' @description Write one graph formatting metadata row into the scope-specific format
'   tracking ListObject (prefixed "graph_formats_"). This stores the visual
'   configuration for a graph that is separate from its series data. The 7 columns
'   are: column 1 = tabId (parent table identifier), column 2 = graphId (matches
'   the graphId in the graph_ids_ table to associate format with series),
'   column 3 = catTitle (category/X-axis title string), column 4 = valuesTitle
'   (values/Y-axis title string), column 5 = hardCodeLabels (Boolean as Long),
'   column 6 = heightFactor (multiplier for graph height, defaults to 1),
'   column 7 = plotTitle (the chart title string, empty by default). The method
'   appends after the last populated row and resizes the ListObject to include
'   the new entry.
' @param scope A Byte value from the AnalysisIdsScope enum selecting the correct
'   format tracking ListObject (e.g., graph_formats_ts for TimeSeries).
' @param tabId The parent table identifier.
' @param graphId The graph identifier that links this format entry to its
'   corresponding series entries in the graph_ids_ ListObject.
' @param catTitle The title string for the category (X) axis of the chart.
' @param valuesTitle The title string for the values (Y) axis of the chart.
' @param hardCodeLabels When True, axis labels are hard-coded as literal values
'   rather than referencing ranges. Stored as CLng (0 or 1).
' @param heightFactor A multiplier applied to the default chart height. Defaults
'   to 1 (standard height). Pass 2 for double-height charts, etc.
' @param plotTitle The main title displayed above the chart. Defaults to an empty
'   string (no title).
Private Sub AddGraphFormat(ByVal scope As Byte, _
                            ByVal tabId As String, _
                            ByVal graphId As String, _
                            ByVal catTitle As String, _
                            ByVal valuesTitle As String, _
                            ByVal hardCodeLabels As Boolean, _
                            Optional ByVal heightFactor As Long = 1, _
                            Optional ByVal plotTitle As String = vbNullString)
    Dim lo As ListObject
    Dim loRng As Range
    Dim cellRng As Range
    Dim sh As Worksheet

    Set lo = ActualLo(scope, prefix:="graph_formats_")
    Set loRng = lo.Range
    Set sh = Wksh()

    If loRng.Cells(2, 1).Value = vbNullString Then
        Set cellRng = loRng.Cells(2, 1)
    Else
        Set cellRng = loRng.Cells(loRng.Rows.Count + 1, 1)
    End If

    cellRng.Value = tabId
    cellRng.Cells(1, 2).Value = graphId
    cellRng.Cells(1, 3).Value = catTitle
    cellRng.Cells(1, 4).Value = valuesTitle
    cellRng.Cells(1, 5).Value = CLng(hardCodeLabels)
    cellRng.Cells(1, 6).Value = heightFactor
    cellRng.Cells(1, 7).Value = plotTitle

    Set loRng = sh.Range(loRng.Cells(1, 1), cellRng.Cells(1, 7))
    lo.Resize loRng
End Sub

'@section Range transfer and graph writing
'===============================================================================

' @description Check whether a named range exists on the given worksheet by
'   attempting to resolve it via sh.Range(rngName). Uses On Error Resume Next
'   to suppress the runtime error that occurs when the name does not exist,
'   then checks whether the resulting Range object is Nothing. This is the
'   standard VBA pattern for safe named-range existence checking since there
'   is no built-in "NameExists" method on Worksheet objects.
' @param sh The worksheet to check for the named range.
' @param rngName The name of the range to look for (e.g., "RNG_SheetTSName").
' @return True if the named range resolves to a valid Range on the worksheet,
'   False otherwise.
' @remark This checks worksheet-scoped names resolved through sh.Range, which
'   also picks up workbook-level names. It does not distinguish between the two.
Private Function RangeExists(ByVal sh As Worksheet, ByVal rngName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExists = (Not rng Is Nothing)
End Function

' @description Copy a named range definition from the input (source analysis) sheet to
'   the output (export) sheet, preserving the same cell address. This is the core
'   mechanism used during export to recreate the analysis table structure: the named
'   range on the output sheet points to the same cell address as on the input sheet,
'   allowing graphs and formulas that reference these names to work correctly on the
'   new sheet. If the named range does not exist on the input sheet, the method
'   silently exits without error.
' @param inpsh The source worksheet that currently owns the named range.
' @param outsh The destination worksheet where the named range will be recreated.
' @param rngName The name of the range to transfer (e.g., a table data range name).
' @remark The address is transferred literally via inpRng.Address, which returns a
'   local cell reference (e.g., "$A$1:$C$10"). The new name is assigned to the
'   equivalent range on outsh, so both sheets end up with identically-addressed
'   named ranges.
Private Sub TransferRange(ByVal inpsh As Worksheet, ByVal outsh As Worksheet, _
                           ByVal rngName As String)
    Dim inpRng As Range
    Dim outRng As Range

    If Not RangeExists(inpsh, rngName) Then Exit Sub

    Set inpRng = inpsh.Range(rngName)
    Set outRng = outsh.Range(inpRng.Address)
    outRng.Name = rngName
End Sub

' @description Iterate all entries in the scope-specific table tracking ListObject
'   (tab_ids_*) and transfer each named range to the output worksheet. This
'   recreates the full analysis table structure on the export sheet by walking
'   down column 1 of the ListObject starting from the first data row (row 2)
'   and calling TransferRange for each range name found in column 2. The loop
'   continues until it hits an empty cell in column 1, indicating the end of
'   entries. The source worksheet is resolved via ActualAnalysisSheet for the
'   given scope.
' @param outsh The destination worksheet where all named ranges will be recreated.
' @param scope A Byte value from the AnalysisIdsScope enum selecting which
'   tracking ListObject and source analysis sheet to use.
' @remark This method reads range names from column 2 (offset 1 from the cellRng
'   in column 1). Column 3 (the export flag) is not checked here; all entries
'   are transferred regardless of their export flag value.
Private Sub WriteTableNames(ByVal outsh As Worksheet, ByVal scope As Byte)
    Dim inpsh As Worksheet
    Dim tabLo As ListObject
    Dim cellRng As Range
    Dim rngName As String

    Set inpsh = ActualAnalysisSheet(scope)
    Set tabLo = ActualLo(scope)
    Set cellRng = tabLo.Range.Cells(2, 1)

    Do While Not IsEmpty(cellRng)
        rngName = cellRng.Offset(, 1).Value
        TransferRange inpsh, outsh, rngName
        Set cellRng = cellRng.Offset(1)
    Loop
End Sub

' @description Read one series entry from a row in the graph tracking ListObject and
'   add it to the given graph object by calling gr.AddSeries and gr.AddLabels.
'   The cellRng parameter points to the graphId cell (column 2 of the ListObject),
'   and all other values are read as offsets from it: offset 1 = seriesName (data
'   range name), offset 2 = seriesType (chart type), offset 3 = seriesPos (axis
'   position), offset 4 = seriesLabel (category label range name), offset 5 =
'   seriesColumnLabel (column header label range name), offset 6 = hardCodeLabels
'   (Boolean), offset 8 = prefix (label prefix string), offset 9 = prefixOnly
'   (Boolean). Note that offset 7 corresponds to the outRangeAddress column which
'   is not needed here (it was already used by WriteGraphs to position the chart).
'   Boolean values are read with On Error Resume Next because they may be stored
'   as numeric 0/1 and CBool conversion can fail on unexpected values.
' @param gr The IGraphs object representing the chart being built. AddSeries and
'   AddLabels are called on this object to configure one data series.
' @param cellRng A Range pointing to the graphId cell in the current row of the
'   graph tracking ListObject. All column values are accessed as offsets from
'   this cell.
' @remark The offset numbering skips offset 7 (outRangeAddress) and jumps to
'   offset 8 (prefix) and offset 9 (prefixOnly). This matches the 11-column
'   layout stored by AddGraphInfo.
Private Sub AddSeries(ByVal gr As IGraphs, ByVal cellRng As Range)
    Dim seriesName As String
    Dim seriesType As String
    Dim seriesPos As String
    Dim seriesLabel As String
    Dim seriesColumnLabel As String
    Dim prefix As String
    Dim prefixOnly As Boolean
    Dim hardCodeLabels As Boolean

    seriesName = cellRng.Offset(, 1).Value
    seriesType = cellRng.Offset(, 2).Value
    seriesPos = cellRng.Offset(, 3).Value
    seriesLabel = cellRng.Offset(, 4).Value
    seriesColumnLabel = cellRng.Offset(, 5).Value
    prefix = cellRng.Offset(, 8).Value

    On Error Resume Next
    hardCodeLabels = CBool(cellRng.Offset(, 6).Value)
    prefixOnly = CBool(cellRng.Offset(, 9).Value)
    On Error GoTo 0

    gr.AddSeries rngName:=seriesName, chrtType:=seriesType, axisPos:=seriesPos
    gr.AddLabels catName:=seriesLabel, lblName:=seriesColumnLabel, _
                 hardCodeLabels:=hardCodeLabels, prefix:=prefix, _
                 prefixOnly:=prefixOnly
End Sub

' @description Apply formatting settings to a fully-built graph by reading values
'   from the format tracking ICustomTable (backed by the graph_formats_ ListObject)
'   and calling gr.Format. The method looks up each format property by graphId
'   using the ICustomTable.Value accessor with column names as keys: "catTitle",
'   "valuesTitle", "plotTitle", "heightFactor", and "hardCodeLabels". It also
'   maps the AnalysisIdsScope to the corresponding GraphScope enum value:
'   AnalysisIdsScopeTimeSeries and AnalysisIdsScopeSpatioTemporal both map to
'   GraphScopeTimeSeries, AnalysisIdsScopeSpatial maps to GraphScopeSpatial,
'   and all other scopes default to GraphScopeNormal. This scope mapping ensures
'   that time-series and spatio-temporal charts share the same graph formatting
'   behavior.
' @param gr The IGraphs object to format. gr.Format is called with the resolved
'   titles, scope, height factor, and hard-code flag.
' @param grTab An ICustomTable wrapping the graph_formats_ ListObject, keyed by
'   the "graphId" column. Used to look up formatting values by graphId.
' @param scope A Byte value from the AnalysisIdsScope enum. Mapped internally to
'   the GraphScope enum for the gr.Format call.
' @param graphId The graph identifier used to look up the format row in grTab.
' @remark The heightFactor lookup uses On Error Resume Next because the value
'   may be missing or non-numeric; it defaults to 1 if the conversion fails.
'   Similarly, hardCodeLabels defaults to False if CBool conversion fails.
Private Sub AddFormat(ByVal gr As IGraphs, ByVal grTab As ICustomTable, _
                       ByVal scope As Byte, ByVal graphId As String)
    Dim valuesTitle As String
    Dim catTitle As String
    Dim plotTitle As String
    Dim heightFactor As Long
    Dim grscope As Byte
    Dim hardCodeLabels As Boolean

    heightFactor = 1
    On Error Resume Next
    heightFactor = CLng(grTab.Value(colName:="heightFactor", keyName:=graphId))
    On Error GoTo 0

    grscope = Switch( _
        (scope = AnalysisIdsScopeTimeSeries Or scope = AnalysisIdsScopeSpatioTemporal), _
            GraphScopeTimeSeries, _
        scope = AnalysisIdsScopeSpatial, GraphScopeSpatial, _
        True, GraphScopeNormal)

    On Error Resume Next
    hardCodeLabels = CBool(grTab.Value(colName:="hardCodeLabels", keyName:=graphId))
    On Error GoTo 0

    valuesTitle = grTab.Value(colName:="valuesTitle", keyName:=graphId)
    catTitle = grTab.Value(colName:="catTitle", keyName:=graphId)
    plotTitle = grTab.Value(colName:="plotTitle", keyName:=graphId)

    gr.Format valuesTitle:=valuesTitle, catTitle:=catTitle, _
              plotTitle:=plotTitle, scope:=grscope, _
              heightFactor:=heightFactor, hardCodeLabels:=hardCodeLabels
End Sub

' @description The main export method: recreates all tracked tables and graphs on the
'   output worksheet for a given scope. First, it calls WriteTableNames to transfer
'   all named ranges from the source analysis sheet to the output sheet, establishing
'   the data foundation. Then it iterates through the graph_ids_ ListObject starting
'   from the first data row, grouping consecutive series rows into charts. A new
'   IGraphs object is created whenever a graph boundary is detected. Graph boundaries
'   are identified by two conditions: (1) firstGraph is True when the previous row
'   is the header row (value = "graphId"), indicating the very first data row; (2)
'   newId is True when the current graphId differs from the previous row's graphId
'   AND it is not the first graph. When a boundary is detected (newId or firstGraph),
'   the previous graph (if any) is finalized with AddFormat, and a new Graphs object
'   is created at the outRangeAddress position. Each series row is processed by
'   AddSeries. When the loop reaches an empty cell (end of data), the final graph
'   is formatted via AddFormat. The format metadata is read from the graph_formats_
'   ListObject wrapped in an ICustomTable keyed by "graphId".
' @param outsh The destination worksheet where tables and graphs will be recreated.
' @param scope A Byte value from the AnalysisIdsScope enum selecting which set of
'   tracking ListObjects (tab_ids_, graph_ids_, graph_formats_) and which source
'   analysis sheet to use.
' @remark The loop walks column 2 (graphId column) of the graph_ids_ ListObject.
'   The outRangeAddress is read from column 9 (offset 7 from column 2). The
'   Application.GoTo call scrolls the view to the chart position, which is
'   required for correct chart placement in some Excel versions.
Private Sub WriteGraphs(ByVal outsh As Worksheet, ByVal scope As Byte)
    Dim gr As IGraphs
    Dim cellRng As Range
    Dim outCellRng As Range
    Dim grLo As ListObject
    Dim grFormatTab As ICustomTable
    Dim grFormLo As ListObject
    Dim newId As Boolean
    Dim firstGraph As Boolean
    Dim graphId As String
    Dim outRangeAddress As String

    WriteTableNames outsh, scope

    Set grLo = ActualLo(scope, prefix:="graph_ids_")
    Set grFormLo = ActualLo(scope, prefix:="graph_formats_")
    Set grFormatTab = CustomTable.Create(grFormLo, idCol:="graphId")

    Set cellRng = grLo.Range.Cells(2, 2)

    Do While Not IsEmpty(cellRng)
        firstGraph = (cellRng.Offset(-1).Value = "graphId")
        newId = (cellRng.Offset(-1).Value <> cellRng.Value) And (Not firstGraph)

        If newId And (Not gr Is Nothing) Then AddFormat gr, grFormatTab, scope, graphId

        graphId = cellRng.Value
        outRangeAddress = cellRng.Offset(, 7).Value

        If newId Or firstGraph Then
            Set outCellRng = outsh.Range(outRangeAddress)
            Set gr = Graphs.Create(outsh, outCellRng)
            Application.GoTo outCellRng
            gr.Add
        End If

        AddSeries gr, cellRng

        If IsEmpty(cellRng.Offset(1)) Then AddFormat gr, grFormatTab, scope, graphId

        Set cellRng = cellRng.Offset(1)
    Loop
End Sub

'@section Validation
'===============================================================================

' @description Check whether a ListObject (Excel table) with the given name exists
'   on the specified worksheet. Uses the standard VBA error-trapping pattern:
'   attempts to access sh.ListObjects(loName) with On Error Resume Next, then
'   checks if the result is Nothing. This is used by CheckRequirements to
'   validate that all 12 tracking ListObjects are present before the factory
'   creates an instance.
' @param sh The worksheet to search for the ListObject.
' @param loName The name of the ListObject to look for (e.g., "tab_ids_ts",
'   "graph_ids_sp", "graph_formats_uba").
' @return True if a ListObject with the given name exists on the worksheet,
'   False otherwise.
Private Function LoExists(ByVal sh As Worksheet, ByVal loName As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = sh.ListObjects(loName)
    On Error GoTo 0
    LoExists = Not (lo Is Nothing)
End Function

' @description Check whether a worksheet with the given name exists in the specified
'   workbook. Uses the standard VBA error-trapping pattern: attempts to access
'   wb.Worksheets(sheetName) with On Error Resume Next, then checks if the
'   result is Nothing. This is used by CheckRequirements to verify that the
'   analysis output sheets referenced by the sheet-name named ranges actually
'   exist in the workbook.
' @param wb The workbook to search for the worksheet.
' @param sheetName The name of the worksheet to look for.
' @return True if a worksheet with the given name exists in the workbook,
'   False otherwise.
Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (sh Is Nothing)
End Function

' @description Validate that the tracking worksheet has all the required infrastructure
'   before allowing an AnaTabIds instance to be created. This method checks for the
'   existence of all 12 tracking ListObjects (3 prefixes x 4 scopes): tab_ids_uba,
'   tab_ids_sp, tab_ids_ts, tab_ids_sptemp, graph_ids_uba, graph_ids_sp,
'   graph_ids_ts, graph_ids_sptemp, graph_formats_uba, graph_formats_sp,
'   graph_formats_ts, graph_formats_sptemp. It also validates the 4 sheet-name
'   named ranges (RNG_SheetUAName, RNG_SheetTSName, RNG_SheetSPName,
'   RNG_SheetSPTempName) and verifies that each named range's value corresponds
'   to an actual worksheet in the workbook. If any ListObject, named range, or
'   referenced worksheet is missing, the method raises an InvalidArgument error
'   via ThrowError with a descriptive message identifying which item is missing.
'   This is called by the Create factory method when check is True (the default).
' @param sh The tracking worksheet to validate. Must contain all 12 ListObjects
'   and 4 named ranges, and the workbook (sh.Parent) must contain the worksheets
'   referenced by those named ranges.
' @remark The validation is intentionally strict: every single ListObject and
'   range must be present. This catches configuration errors early rather than
'   allowing them to surface as cryptic runtime errors during export operations.
Private Sub CheckRequirements(ByVal sh As Worksheet)
    Dim loArray As BetterArray
    Dim counter As Long
    Dim wb As Workbook
    Dim actName As String
    Dim rngArray As BetterArray

    Set loArray = New BetterArray
    loArray.Push "tab_ids_uba", "tab_ids_sp", "tab_ids_ts", "tab_ids_sptemp", _
                 "graph_ids_uba", "graph_ids_sp", "graph_ids_ts", "graph_ids_sptemp", _
                 "graph_formats_uba", "graph_formats_sp", "graph_formats_ts", "graph_formats_sptemp"

    For counter = loArray.LowerBound To loArray.UpperBound
        actName = loArray.Item(counter)
        If Not LoExists(sh, actName) Then _
            ThrowError InvalidArgument, "Table " & actName & " is missing"
    Next

    Set rngArray = New BetterArray
    rngArray.Push RNGUASHEET, RNGTSSHEET, RNGSPSHEET, RNGSPTEMPSHEET
    Set wb = sh.Parent

    For counter = rngArray.LowerBound To rngArray.UpperBound
        actName = rngArray.Item(counter)
        If Not RangeExists(sh, actName) Then _
            ThrowError InvalidArgument, "Range " & actName & " is missing"
        If Not SheetExists(wb, sh.Range(actName).Value) Then _
            ThrowError InvalidArgument, "Sheet for " & actName & " is missing"
    Next
End Sub

'@section Error handling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface implementation
'===============================================================================

Private Sub IAnaTabIds_AddGraphInfo(ByVal scope As Byte, _
    ByVal tabId As String, ByVal graphId As String, _
    ByVal seriesName As String, ByVal seriesType As String, _
    ByVal seriesPos As String, ByVal seriesLabel As String, _
    ByVal seriesColumnLabel As String, ByVal outRangeAddress As String, _
    Optional ByVal hardCodeLabels As Boolean = False, _
    Optional ByVal prefix As String = vbNullString, _
    Optional ByVal prefixOnly As Boolean = False)

    AddGraphInfo scope, tabId, graphId, seriesName, seriesType, _
                 seriesPos, seriesLabel, seriesColumnLabel, _
                 hardCodeLabels, outRangeAddress, prefix, prefixOnly
End Sub

Private Sub IAnaTabIds_AddGraphFormat(ByVal scope As Byte, _
    ByVal tabId As String, ByVal graphId As String, _
    ByVal catTitle As String, ByVal valuesTitle As String, _
    Optional ByVal hardCodeLabels As Boolean = False, _
    Optional ByVal heightFactor As Long = 1, _
    Optional ByVal plotTitle As String = vbNullString)

    AddGraphFormat scope, tabId, graphId, catTitle, valuesTitle, _
                   hardCodeLabels, heightFactor, plotTitle
End Sub

Private Sub IAnaTabIds_AddTableInfos(ByVal scope As Byte, _
    ByVal tabId As String, ByVal tabRangesNames As BetterArray, _
    Optional ByVal exportTable As String = "yes")

    AddTableInfos scope, tabId, tabRangesNames, exportTable
End Sub

Private Sub IAnaTabIds_WriteGraphs(ByVal outsh As Worksheet, ByVal scope As Byte)
    WriteGraphs outsh, scope
End Sub
