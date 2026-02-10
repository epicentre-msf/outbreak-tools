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

'@class AnaTabIds
'@description
'Records analysis table named ranges and graph series metadata into
'ListObjects on a hidden tracking worksheet. The tracking data is consumed
'during export to recreate the full analysis structure (tables and charts)
'on output worksheets. Manages three ListObject families (tab_ids_*,
'graph_ids_*, graph_formats_*) across four analysis scopes (normal, time
'series, spatial, spatio-temporal).
'@depends CustomTable, Graphs, IGraphs, BetterArray, IFormulas
'@version 2025-01-15
'@author EpiData

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

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a new AnaTabIds instance bound to a tracking worksheet
'@details
'Factory method on the predeclared instance. Validates the worksheet reference
'and optionally checks that all 12 tracking ListObjects (3 prefixes x 4 scopes)
'and 4 sheet-name named ranges exist. Pass check as False to skip validation
'when the tracking sheet is being built incrementally.
'@param idsh Worksheet. The worksheet containing the tracking ListObjects and
'sheet-name named ranges.
'@param check Optional Boolean. When True, calls CheckRequirements to validate
'all expected infrastructure. Defaults to True.
'@return IAnaTabIds. A fully initialised tracking instance.
'@throws ProjectError.InvalidArgument When idsh is Nothing.
'@throws ProjectError.InvalidArgument When check is True and any required
'ListObject or named range is missing.
'@depends CustomTable, BetterArray
'@export
Public Function Create(ByVal idsh As Worksheet, _
                       Optional ByVal check As Boolean = True) As IAnaTabIds

    If idsh Is Nothing Then ThrowError InvalidArgument, "Worksheet not defined"
    If check Then CheckRequirements idsh

    With New AnaTabIds
        Set .Wksh = idsh
        Set Create = .Self
    End With
End Function


'@section Internal Properties
'===============================================================================
'@description Factory-support properties for construction and worksheet access.

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor so consuming code can fluently retrieve the interface
'reference from the predeclared Create method.
'@return IAnaTabIds. The current instance cast to the interface.
Public Property Get Self() As IAnaTabIds
    Set Self = Me
End Property

'@label:wksh-set
'@prop-title Assign the tracking worksheet
'@param sh Worksheet. The worksheet to bind.
Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

'@label:wksh
'@prop-title Tracking worksheet reference
'@details
'Returns the worksheet that hosts the tracking ListObjects and sheet-name
'named ranges.
'@return Worksheet. The bound tracking worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@section ListObject Access
'===============================================================================
'@description Routing helpers that resolve scope-specific ListObjects and worksheets.

'@label:actual-lo
'@prop-title Resolve the tracking ListObject for a scope and prefix
'@details
'Maps the AnalysisIdsScope enum value to a suffix string (ts, sp, uba,
'sptemp) and concatenates it with the prefix to form the ListObject name.
'This is the central routing mechanism all tracking methods use.
'@param scope Byte. AnalysisIdsScope value selecting the analysis type.
'@param prefix Optional String. ListObject name prefix. Defaults to
'"tab_ids_". Pass "graph_ids_" or "graph_formats_" for other families.
'@return ListObject. The resolved tracking ListObject.
'@remarks If the scope does not match any known case, actName remains
'empty and the ListObject lookup will raise a runtime error.
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

'@label:actual-analysis-sheet
'@prop-title Resolve the analysis output worksheet for a scope
'@details
'Reads the sheet-name named range corresponding to the scope from the
'tracking worksheet (RNG_SheetUAName, RNG_SheetTSName, RNG_SheetSPName,
'or RNG_SheetSPTempName) and returns the matching worksheet from the
'parent workbook.
'@param scope Byte. AnalysisIdsScope value selecting the output sheet.
'@return Worksheet. The analysis output worksheet for the given scope.
'@remarks If the named range value does not correspond to an existing
'worksheet, the Worksheets collection access raises a runtime error.
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

'@section Table Tracking
'===============================================================================
'@description Methods for registering table named ranges for export.

'@label:add-table-infos
'@sub-title Register table named ranges in the tracking ListObject
'@details
'Writes entries into the scope-specific tab_ids_ ListObject. Each entry
'consists of the table identifier in column 1, the named range name in
'column 2, and the export flag in column 3. The range names are bulk-
'written via BetterArray.ToExcelRange. The ListObject is resized after
'appending. Exits immediately when tabRangesNames is empty.
'@param scope Byte. AnalysisIdsScope value selecting the tracking ListObject.
'@param tabId String. Identifier of the analysis table.
'@param tabRangesNames BetterArray. Named range names to register.
'@param exportTable Optional String. Export flag ("yes" or "no").
'Defaults to "yes".
'@export
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

'@section Graph Tracking
'===============================================================================
'@description Methods for registering graph series and formatting metadata.

'@label:add-graph-info
'@sub-title Register a graph series row in the tracking ListObject
'@details
'Writes one series entry into the scope-specific graph_ids_ ListObject.
'The 11 columns store: tabId, graphId, seriesName, seriesType, seriesPos,
'seriesLabel, seriesColumnLabel, hardCodeLabels (as CLng 0/1),
'outRangeAddress, prefix, and prefixOnly (as CLng 0/1). The method
'appends after the last populated row and resizes the ListObject.
'@param scope Byte. AnalysisIdsScope value selecting the tracking ListObject.
'@param tabId String. Identifier of the parent analysis table.
'@param graphId String. Identifier of the graph receiving this series.
'@param seriesName String. Named range name for the series data.
'@param seriesType String. Chart series type (e.g. "xlLine").
'@param seriesPos String. Axis position indicator for the series.
'@param seriesLabel String. Named range name for category labels.
'@param seriesColumnLabel String. Named range name for column header labels.
'@param hardCodeLabels Boolean. When True, labels are written as literals.
'@param outRangeAddress String. Cell address for graph placement on output.
'@param prefix Optional String. Admin-level prefix for spatial graphs.
'Defaults to vbNullString.
'@param prefixOnly Optional Boolean. When True, only the prefix is used as
'the label. Defaults to False.
'@remarks Boolean parameters are stored as Long (0/1) via CLng for reliable
'cross-session persistence in Excel ListObject cells.
'@export
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

'@label:add-graph-format
'@sub-title Register graph formatting metadata in the tracking ListObject
'@details
'Writes one format entry into the scope-specific graph_formats_ ListObject.
'The 7 columns store: tabId, graphId, catTitle, valuesTitle, hardCodeLabels
'(as CLng 0/1), heightFactor, and plotTitle. The method appends after the
'last populated row and resizes the ListObject. The graphId links this
'format entry to its corresponding series entries in graph_ids_.
'@param scope Byte. AnalysisIdsScope value selecting the tracking ListObject.
'@param tabId String. Identifier of the parent analysis table.
'@param graphId String. Identifier of the graph to format.
'@param catTitle String. Category (X) axis title.
'@param valuesTitle String. Value (Y) axis title.
'@param hardCodeLabels Boolean. When True, axis labels are written as literals.
'@param heightFactor Optional Long. Multiplier for graph height. Defaults to 1.
'@param plotTitle Optional String. Chart title. Defaults to vbNullString.
'@export
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

'@section Range Transfer and Graph Writing
'===============================================================================
'@description Helpers for transferring named ranges and rendering charts.

'@label:range-exists
'@sub-title Check whether a named range exists on a worksheet
'@details
'Attempts to resolve the name via sh.Range(rngName) with error trapping.
'Returns True when the name resolves to a valid Range. Also picks up
'workbook-level names resolved through the worksheet.
'@param sh Worksheet. The worksheet to check.
'@param rngName String. The named range to look for.
'@return Boolean. True when the named range exists.
Private Function RangeExists(ByVal sh As Worksheet, ByVal rngName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExists = (Not rng Is Nothing)
End Function

'@label:transfer-range
'@sub-title Copy a named range definition from source to output worksheet
'@details
'Recreates a named range on the output worksheet using the same cell address
'as the source. This is the core mechanism for export: graphs and formulas
'that reference these names work correctly on the new sheet. Silently exits
'when the source range does not exist.
'@param inpsh Worksheet. The source worksheet owning the named range.
'@param outsh Worksheet. The destination worksheet for the new name.
'@param rngName String. The named range to transfer.
Private Sub TransferRange(ByVal inpsh As Worksheet, ByVal outsh As Worksheet, _
                           ByVal rngName As String)
    Dim inpRng As Range
    Dim outRng As Range

    If Not RangeExists(inpsh, rngName) Then Exit Sub

    Set inpRng = inpsh.Range(rngName)
    Set outRng = outsh.Range(inpRng.Address)
    outRng.Name = rngName
End Sub

'@label:write-table-names
'@sub-title Transfer all tracked named ranges to the output worksheet
'@details
'Iterates the tab_ids_ ListObject for the scope and calls TransferRange
'for each named range name found in column 2. The source analysis sheet
'is resolved via ActualAnalysisSheet. The loop continues until an empty
'cell is hit in column 1.
'@param outsh Worksheet. The destination worksheet for the named ranges.
'@param scope Byte. AnalysisIdsScope value selecting the tracking table.
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

'@label:add-series
'@sub-title Read one series entry and add it to a graph object
'@details
'Reads series metadata from a row in the graph_ids_ ListObject and calls
'gr.AddSeries and gr.AddLabels. The cellRng points to the graphId cell
'(column 2); all other values are read as offsets. Boolean values are
'read with error trapping because they are stored as numeric 0/1.
'@param gr IGraphs. The chart object receiving the series.
'@param cellRng Range. The graphId cell in the current tracking row.
'@remarks Offset numbering skips column 9 (outRangeAddress, offset 7)
'and jumps to offset 8 (prefix) and offset 9 (prefixOnly), matching
'the 11-column layout stored by AddGraphInfo.
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

'@label:add-format
'@sub-title Apply formatting settings to a built graph
'@details
'Reads format metadata from the graph_formats_ ICustomTable by graphId
'and calls gr.Format. Maps AnalysisIdsScope to GraphScope: TimeSeries
'and SpatioTemporal map to GraphScopeTimeSeries, Spatial maps to
'GraphScopeSpatial, and all others default to GraphScopeNormal.
'@param gr IGraphs. The chart object to format.
'@param grTab ICustomTable. Wrapper around graph_formats_ ListObject,
'keyed by "graphId".
'@param scope Byte. AnalysisIdsScope value for scope-to-GraphScope mapping.
'@param graphId String. The graph identifier for the format lookup.
'@remarks heightFactor defaults to 1 and hardCodeLabels defaults to False
'when the stored values cannot be converted.
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

'@label:write-graphs
'@sub-title Write all tracked graphs to the output worksheet
'@details
'First calls WriteTableNames to transfer all named ranges, then iterates
'the graph_ids_ ListObject to build charts. A new IGraphs object is
'created at each graph boundary (detected when graphId changes). Each
'series row is processed by AddSeries. When the loop ends, the final
'graph is formatted via AddFormat. Format metadata is read from the
'graph_formats_ ListObject wrapped in an ICustomTable keyed by "graphId".
'@param outsh Worksheet. The output worksheet receiving the chart objects.
'@param scope Byte. AnalysisIdsScope value selecting the tracking tables.
'@remarks The loop walks column 2 (graphId) of graph_ids_. The
'outRangeAddress is read from column 9 (offset 7 from column 2).
'Application.GoTo scrolls to the chart position, which is required for
'correct chart placement in some Excel versions.
'@export
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
'@description Existence checks and factory validation.

'@label:lo-exists
'@sub-title Check whether a ListObject exists on a worksheet
'@details
'Uses the standard VBA error-trapping pattern to test for ListObject
'existence. Used by CheckRequirements to validate the tracking sheet.
'@param sh Worksheet. The worksheet to search.
'@param loName String. The ListObject name to look for.
'@return Boolean. True when the ListObject exists.
Private Function LoExists(ByVal sh As Worksheet, ByVal loName As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = sh.ListObjects(loName)
    On Error GoTo 0
    LoExists = Not (lo Is Nothing)
End Function

'@label:sheet-exists
'@sub-title Check whether a worksheet exists in a workbook
'@details
'Uses the standard VBA error-trapping pattern to test for worksheet
'existence. Used by CheckRequirements to verify that the sheets
'referenced by the named ranges actually exist.
'@param wb Workbook. The workbook to search.
'@param sheetName String. The worksheet name to look for.
'@return Boolean. True when the worksheet exists.
Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (sh Is Nothing)
End Function

'@label:check-requirements
'@sub-title Validate the tracking worksheet infrastructure
'@details
'Checks for all 12 tracking ListObjects (tab_ids_uba/sp/ts/sptemp,
'graph_ids_uba/sp/ts/sptemp, graph_formats_uba/sp/ts/sptemp) and the
'4 sheet-name named ranges (RNG_SheetUAName, RNG_SheetTSName,
'RNG_SheetSPName, RNG_SheetSPTempName). Also verifies that each named
'range value corresponds to an actual worksheet in the workbook.
'@param sh Worksheet. The tracking worksheet to validate.
'@throws ProjectError.InvalidArgument When any ListObject, named range,
'or referenced worksheet is missing.
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

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME.
'@param errNumber Long. The ProjectError code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError. Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the IAnaTabIds contract.
'See the corresponding Private members above for full documentation.

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
