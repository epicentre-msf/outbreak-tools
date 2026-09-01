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

'@class AnaTabIds
'@description
'Records the charts of the analysis sheets so an export can draw them again on
'the sheets it creates, and copies the named ranges those charts read.
'PrepareSheet builds the tracking worksheet, AddGraphInfo records one chart
'series and AddGraphFormat the look of one chart, TransferNames copies the
'named ranges and their values onto an output worksheet, and WriteGraphs draws
'the charts there.
'
'ONE REGISTRY TABLE
'-------------------------------------------------------------------------------
'The tracking worksheet holds a single ListObject. Every row carries the scope
'it belongs to and a kind: a series row describes one chart series, a format
'row describes the look of one chart. The rows of one chart stay together
'because they are appended in call order, and the reader starts a new chart
'when the graph identifier changes.
'
'THE TABLE NAMES COME FROM THE WORKBOOK
'-------------------------------------------------------------------------------
'Every named range of an analysis table was created by CrossTable through
'Range.Name, so the workbook already holds it. TransferNames reads the workbook
'name collection, keeps the names that resolve to a range on the analysis
'sheet, and creates each one at the same address on the output sheet. The
'collection is walked once per export and sorted by owning worksheet, because
'the export asks for four sheets in turn.
'@depends Graphs
'@version 2026-07-30


Private Type TAnaIds
    sh As Worksheet
    sealed As Boolean
    registry As ListObject
    namesBySheet As Collection
End Type

Private this As TAnaIds
Private Const CLASS_NAME As String = "AnaTabIds"

Private Const REGISTRY_TABLE As String = "ana_registry"
Private Const REGISTRY_COLUMN_COUNT As Long = 16
Private Const KIND_SERIES As String = "series"
Private Const KIND_FORMAT As String = "format"

'@description
'Analysis scopes. Each one is a sheet of the generated linelist. This class is
'the leaf that reads the value, so it owns the enum: AnalysisOutput and
'LLExporter both name it, and the generated linelist carries this class
'without carrying AnalysisOutput.
'
'A fifth member used to sit here, standing for the tables-only pass. It was a
'build stage wearing the type of a scope, and four Select Case blocks that
'route on a scope had no branch for it. AnalysisOutput owns that enum now,
'under the name AnalysisBuildStage.
Public Enum AnalysisScope
    AnalysisScopeTimeSeries = 1
    AnalysisScopeSpatial = 2
    AnalysisScopeNormal = 3
    AnalysisScopeSpatioTemporal = 4
End Enum

'@description
'Column positions of the registry table. The headers are written from this
'enum by PrepareSheet, so one place decides both the layout and every read.
Private Enum RegistryColumnIndex
    RegistryScope = 1
    RegistryKind = 2
    RegistryGraphId = 3
    RegistrySeriesName = 4
    RegistrySeriesType = 5
    RegistrySeriesPos = 6
    RegistrySeriesLabel = 7
    RegistrySeriesColumnLabel = 8
    RegistryHardCodeLabels = 9
    RegistryOutRangeAddress = 10
    RegistryPrefix = 11
    RegistryPrefixOnly = 12
    RegistryCatTitle = 13
    RegistryValuesTitle = 14
    RegistryHeightFactor = 15
    RegistryPlotTitle = 16
End Enum

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a new AnaTabIds instance bound to a tracking worksheet
'@details
'Factory method on the predeclared instance. Validates the worksheet reference
'and, when asked, that the registry table is on the sheet. Pass check as False
'while the linelist is still being built.
'@param idsh Worksheet. The worksheet hosting the registry table.
'@param check Optional Boolean. When True, calls CheckRequirements. Defaults
'to True.
'@return AnaTabIds. A fully initialised tracking instance.
'@throws ProjectError.InvalidArgument When idsh is Nothing.
'@throws ProjectError.InvalidArgument When check is True and the registry
'table is missing.
'@export
Public Function Create(ByVal idsh As Worksheet, _
                       Optional ByVal check As Boolean = True) As AnaTabIds

    If idsh Is Nothing Then ThrowError ProjectError.InvalidArgument, "Worksheet not defined"
    If check Then CheckRequirements idsh

    Dim instance As AnaTabIds
    Set instance = New AnaTabIds
    Set instance.Wksh = idsh
    instance.Seal
    Set Create = instance
End Function

'@label:prepare-sheet
'@sub-title Build the registry table on a tracking worksheet
'@details
'Called on the predeclared instance while the linelist is being built. Writes
'the header row from RegistryColumnIndex and lists it as REGISTRY_TABLE. A
'sheet that already carries the table is left as it is, so a second build over
'the same workbook changes nothing.
'@param idsh Worksheet. The worksheet to prepare.
'@throws ProjectError.InvalidArgument When idsh is Nothing.
'@export
Public Sub PrepareSheet(ByVal idsh As Worksheet)
    Dim headers As Variant
    Dim headerStart As Range
    Dim tableRange As Range
    Dim counter As Long

    If idsh Is Nothing Then ThrowError ProjectError.InvalidArgument, "Worksheet not defined"
    If LoExists(idsh, REGISTRY_TABLE) Then Exit Sub

    headers = RegistryHeaders()
    Set headerStart = idsh.Cells(1, 1)

    For counter = LBound(headers) To UBound(headers)
        headerStart.Cells(1, counter + 1).Value = headers(counter)
    Next

    Set tableRange = idsh.Range(headerStart, headerStart.Cells(2, REGISTRY_COLUMN_COUNT))
    idsh.ListObjects.Add(xlSrcRange, tableRange, , xlYes).Name = REGISTRY_TABLE
End Sub


'@section Internal Properties
'===============================================================================
'@description Factory-support properties for construction and worksheet access.

'@label:seal
'@sub-title Seal the instance against further setup writes.
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against post-creation writes.
'@param propName String. The property name for the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:wksh-set
'@prop-title Assign the tracking worksheet
'@param sh Worksheet. The worksheet to bind.
Public Property Set Wksh(ByVal sh As Worksheet)
    GuardNotSealed "Wksh"
    Set this.sh = sh
End Property

'@label:wksh
'@prop-title Tracking worksheet reference
'@return Worksheet. The bound tracking worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@section Registry Access
'===============================================================================
'@description The one table every read and every write goes through.

'@label:registry-headers
'@sub-title The header row of the registry table, in column order
'@return Variant. A zero-based array of REGISTRY_COLUMN_COUNT header names.
Private Function RegistryHeaders() As Variant
    RegistryHeaders = Array("scope", "kind", "graphId", _
                            "seriesName", "seriesType", "seriesPos", _
                            "seriesLabel", "seriesColumnLabel", "hardCodeLabels", _
                            "outRangeAddress", "prefix", "prefixOnly", _
                            "catTitle", "valuesTitle", "heightFactor", "plotTitle")
End Function

'@label:registry
'@prop-title Resolve the registry table, held after the first read
'@return ListObject. The registry table of the tracking worksheet.
'@throws ProjectError.InvalidArgument When the table is missing.
Private Property Get Registry() As ListObject
    If this.registry Is Nothing Then
        If Not LoExists(Wksh(), REGISTRY_TABLE) Then _
            ThrowError ProjectError.InvalidArgument, _
                       "Table " & REGISTRY_TABLE & " is missing"
        Set this.registry = Wksh().ListObjects(REGISTRY_TABLE)
    End If
    Set Registry = this.registry
End Property

'@label:registry-values
'@sub-title Read the whole registry table into memory
'@details
'One read serves the whole walk of one scope. A table with no rows answers
'Empty, which the callers test with IsEmpty.
'@return Variant. A two-dimensional array of the data rows, or Empty.
Private Function RegistryValues() As Variant
    Dim lo As ListObject

    Set lo = Registry()
    If lo.DataBodyRange Is Nothing Then Exit Function
    RegistryValues = lo.DataBodyRange.Value2
End Function

'@label:append-row
'@sub-title Append one row to the registry table in a single write
'@details
'The tracking worksheet holds this one table and nothing else, so growing the
'table over the cells below it is safe. A freshly built table carries one blank
'row and the first entry goes into it.
'@param fields Variant. A one-dimensional array of REGISTRY_COLUMN_COUNT
'values, in RegistryColumnIndex order.
Private Sub AppendRow(ByRef fields As Variant)
    Dim lo As ListObject
    Dim rowCount As Long

    Set lo = Registry()

    If lo.DataBodyRange Is Nothing Then
        lo.Resize lo.Range.Resize(2)
        lo.DataBodyRange.Rows(1).Value2 = fields
        Exit Sub
    End If

    rowCount = lo.DataBodyRange.Rows.Count

    If rowCount = 1 And IsEmpty(lo.DataBodyRange.Cells(1, 1)) Then
        lo.DataBodyRange.Rows(1).Value2 = fields
        Exit Sub
    End If

    lo.Resize lo.Range.Resize(rowCount + 2)
    lo.DataBodyRange.Rows(rowCount + 1).Value2 = fields
End Sub

'@label:validate-scope
'@sub-title Refuse a scope this class cannot route
'@param scope Byte. AnalysisScope value to test.
'@throws ProjectError.InvalidArgument When the scope is outside the four
'storage scopes.
Private Sub ValidateScope(ByVal scope As Byte)
    Select Case scope
    Case AnalysisScopeTimeSeries, AnalysisScopeSpatial, _
         AnalysisScopeNormal, AnalysisScopeSpatioTemporal
        Exit Sub
    Case Else
        ThrowError ProjectError.InvalidArgument, _
                   "Unknown analysis scope: " & scope
    End Select
End Sub

'@section Graph Tracking
'===============================================================================
'@description Methods for registering graph series and formatting metadata.

'@label:add-graph-info
'@sub-title Register a graph series row in the registry
'@details
'Writes one series row. The rows of one chart have to stay together, and they
'do because every writer registers the series of a chart one after the other.
'The two flags are stored as 0 or 1 so they survive a save and reopen.
'@param scope Byte. AnalysisScope value the chart belongs to.
'@param graphId String. Identifier of the chart receiving this series.
'@param seriesName String. Named range name for the series data.
'@param seriesType String. Chart series type (e.g. "line").
'@param seriesPos String. Axis position indicator for the series.
'@param seriesLabel String. Named range name for category labels.
'@param seriesColumnLabel String. Named range name for column header labels.
'@param hardCodeLabels Boolean. When True, labels are written as literals.
'@param outRangeAddress String. Cell address for chart placement on output.
'@param prefix Optional String. Admin-level prefix for spatial graphs.
'Defaults to vbNullString.
'@param prefixOnly Optional Boolean. When True, only the prefix is used as
'the label. Defaults to False.
'@throws ProjectError.InvalidArgument When the scope is unknown.
'@export
Public Sub AddGraphInfo(ByVal scope As Byte, _
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
    Dim fields As Variant

    ValidateScope scope

    fields = EmptyRow()
    fields(RegistryScope) = CLng(scope)
    fields(RegistryKind) = KIND_SERIES
    fields(RegistryGraphId) = graphId
    fields(RegistrySeriesName) = seriesName
    fields(RegistrySeriesType) = seriesType
    fields(RegistrySeriesPos) = seriesPos
    fields(RegistrySeriesLabel) = seriesLabel
    fields(RegistrySeriesColumnLabel) = seriesColumnLabel
    fields(RegistryHardCodeLabels) = CLng(hardCodeLabels)
    fields(RegistryOutRangeAddress) = outRangeAddress
    fields(RegistryPrefix) = prefix
    fields(RegistryPrefixOnly) = CLng(prefixOnly)

    AppendRow fields
End Sub

'@label:add-graph-format
'@sub-title Register the look of one chart in the registry
'@details
'Writes one format row, found again by the graph identifier when the chart has
'all its series.
'@param scope Byte. AnalysisScope value the chart belongs to.
'@param graphId String. Identifier of the chart to format.
'@param catTitle String. Category (X) axis title.
'@param valuesTitle String. Value (Y) axis title.
'@param hardCodeLabels Boolean. When True, axis labels are written as literals.
'@param heightFactor Optional Long. Multiplier for chart height. Defaults to 1.
'@param plotTitle Optional String. Chart title. Defaults to vbNullString.
'@throws ProjectError.InvalidArgument When the scope is unknown.
'@export
Public Sub AddGraphFormat(ByVal scope As Byte, _
                          ByVal graphId As String, _
                          ByVal catTitle As String, _
                          ByVal valuesTitle As String, _
                          ByVal hardCodeLabels As Boolean, _
                          Optional ByVal heightFactor As Long = 1, _
                          Optional ByVal plotTitle As String = vbNullString)
    Dim fields As Variant

    ValidateScope scope

    fields = EmptyRow()
    fields(RegistryScope) = CLng(scope)
    fields(RegistryKind) = KIND_FORMAT
    fields(RegistryGraphId) = graphId
    fields(RegistryHardCodeLabels) = CLng(hardCodeLabels)
    fields(RegistryCatTitle) = catTitle
    fields(RegistryValuesTitle) = valuesTitle
    fields(RegistryHeightFactor) = heightFactor
    fields(RegistryPlotTitle) = plotTitle

    AppendRow fields
End Sub

'@label:empty-row
'@sub-title A registry row with every column blank
'@return Variant. A one-based array of REGISTRY_COLUMN_COUNT empty strings.
Private Function EmptyRow() As Variant
    Dim fields As Variant
    Dim counter As Long

    ReDim fields(1 To REGISTRY_COLUMN_COUNT)
    For counter = 1 To REGISTRY_COLUMN_COUNT
        fields(counter) = vbNullString
    Next

    EmptyRow = fields
End Function

'@section Name Transfer
'===============================================================================
'@description Copying the named ranges of an analysis sheet to the output sheet.

'@label:transfer-names
'@sub-title Recreate every named range of an analysis sheet on the output sheet, values and all
'@details
'A chart reads its values and its labels through named ranges, so the output
'sheet needs the same names at the same addresses. The names are read from the
'workbook collection: every one of them was created on the analysis sheet by
'the table builder, which is why nothing has to be stored to find them again.
'A name holding a value has no range behind it, so it is skipped, and so is a
'name that no longer resolves.
'
'THE VALUES TRAVEL WITH THE NAMES
'-------------------------------------------------------------------------------
'The export used to carry the values of an analysis sheet through a whole-grid
'clipboard copy, every cell of the sheet, before this routine ran. That pass is
'gone: each name here is one contiguous block, and its values ride onto the same
'address of the output sheet through BetterArray. Every cell a table writes sits
'inside a name (TestCrossTable and TestCrossTableFormula pin that, scope by
'scope), so nothing is left behind. Two names that overlap write the same cells
'twice with the same values, which is harmless.
'@param inpsh Worksheet. The analysis sheet the names live on.
'@param outsh Worksheet. The sheet the names are recreated on.
'@throws ProjectError.InvalidArgument When either worksheet is Nothing.
'@export
Public Sub TransferNames(ByVal inpsh As Worksheet, ByVal outsh As Worksheet)
    Dim bucket As Collection
    Dim entry As Variant

    If inpsh Is Nothing Then _
        ThrowError ProjectError.InvalidArgument, "Input worksheet not defined"
    If outsh Is Nothing Then _
        ThrowError ProjectError.InvalidArgument, "Output worksheet not defined"

    Set bucket = SheetNames(inpsh)

    For Each entry In bucket
        outsh.Range(entry(1)).Name = entry(0)
        MoveValues inpsh.Range(entry(1)), outsh.Range(entry(1))
    Next entry
End Sub

'@label:move-values
'@sub-title Copy the values of one named block onto the same address of the output sheet
'@details
'BetterArray.FromExcelRange flattens a single row and a single column to one
'dimension, and ToExcelRange writes one dimension as a COLUMN. So a one-row
'block goes back transposed, or its values would run down the sheet instead of
'across it. A single cell never enters the array: its Value is a scalar, and
'one read and one write is the whole move. A name over several areas is moved
'an area at a time, each one a contiguous block.
'@param source Range. The block on the analysis sheet.
'@param target Range. The same address on the output sheet.
Private Sub MoveValues(ByVal source As Range, ByVal target As Range)
    Dim block As BetterArray
    Dim area As Range

    If source.Areas.Count > 1 Then
        For Each area In source.Areas
            MoveValues area, target.Worksheet.Range(area.Address)
        Next area
        Exit Sub
    End If

    If source.Cells.Count = 1 Then
        target.Value = source.Value
        Exit Sub
    End If

    Set block = New BetterArray
    block.FromExcelRange source
    block.ToExcelRange target.Cells(1, 1), TransposeValues:=(source.Rows.Count = 1)
End Sub

'@label:sheet-names
'@sub-title The tracked names of one worksheet, name and address
'@details
'Answers from the bucket built on the first call. The export asks for four
'sheets in turn and the workbook name collection holds thousands of entries,
'so it is walked once.
'@param inpsh Worksheet. The worksheet to answer for.
'@return Collection. Array(simple name, address) entries, empty when the sheet
'carries no name.
Private Function SheetNames(ByVal inpsh As Worksheet) As Collection
    Dim bucket As Collection

    If this.namesBySheet Is Nothing Then BuildNameBuckets inpsh.Parent

    Set bucket = Nothing
    'A sheet with no tracked name has no bucket, and Collection answers a
    'missing key by raising. This one call is what the trap covers.
    On Error Resume Next
    Set bucket = this.namesBySheet(inpsh.Name)
    On Error GoTo 0

    If bucket Is Nothing Then Set bucket = New Collection
    Set SheetNames = bucket
End Function

'@label:build-name-buckets
'@sub-title Walk the workbook names once and bucket them by owning worksheet
'@details
'The owning sheet is read off RefersToRange. An analysis sheet name is
'translated and can carry spaces and apostrophes, which makes the RefersTo text
'a poor thing to match on. A name that holds a value has no range behind it and
'raises, which is how the two are told apart.
'@param wkb Workbook. The workbook whose names are read.
Private Sub BuildNameBuckets(ByVal wkb As Workbook)
    Dim definition As Name
    Dim target As Range
    Dim bucket As Collection
    Dim sheetKey As String

    Set this.namesBySheet = New Collection

    For Each definition In wkb.Names
        Set target = Nothing
        'A constant name, a name over several sheets and a broken name all
        'raise here. The trap covers this one call and nothing else.
        On Error Resume Next
        Set target = definition.RefersToRange
        On Error GoTo 0

        If Not target Is Nothing Then
            sheetKey = target.Worksheet.Name

            Set bucket = Nothing
            On Error Resume Next
            Set bucket = this.namesBySheet(sheetKey)
            On Error GoTo 0

            If bucket Is Nothing Then
                Set bucket = New Collection
                this.namesBySheet.Add bucket, sheetKey
            End If

            bucket.Add Array(SimpleName(definition.Name), target.Address)
        End If
    Next definition
End Sub

'@label:simple-name
'@sub-title The name without its worksheet qualifier
'@details
'A worksheet-scoped name reads as "Sheet name!LOCAL_NAME". The output sheet
'gets the part after the separator.
'@param fullName String. The name as the Names collection reports it.
'@return String. The name with no qualifier.
Private Function SimpleName(ByVal fullName As String) As String
    Dim position As Long

    position = InStrRev(fullName, "!")
    If position = 0 Then
        SimpleName = fullName
    Else
        SimpleName = Mid$(fullName, position + 1)
    End If
End Function

'@section Graph Writing
'===============================================================================
'@description Drawing the tracked charts on an output worksheet.

'@label:write-graphs
'@sub-title Recreate the named ranges and the charts of one scope
'@details
'Copies the named ranges of the analysis sheet first, because a chart cannot
'read a series before its name exists. Then walks the series rows of the scope
'in registry order and starts a new chart every time the graph identifier
'changes. Each chart is formatted once it holds all its series.
'@param outsh Worksheet. The output worksheet receiving the charts.
'@param scope Byte. AnalysisScope value selecting the registry rows.
'@param inpsh Worksheet. The analysis sheet the names are copied from. The
'caller holds it already, which is why it is passed in.
'@throws ProjectError.InvalidArgument When a worksheet is Nothing or the
'scope is unknown.
'@remarks Application.GoTo scrolls to the chart position, which some Excel
'versions need for correct chart placement. It also means the output sheet
'has to be selectable.
'@export
Public Sub WriteGraphs(ByVal outsh As Worksheet, ByVal scope As Byte, _
                       ByVal inpsh As Worksheet)
    Dim gr As Graphs
    Dim values As Variant
    Dim outCellRng As Range
    Dim graphId As String
    Dim previousId As String
    Dim rowIndex As Long

    ValidateScope scope
    TransferNames inpsh, outsh

    values = RegistryValues()
    If IsEmpty(values) Then Exit Sub

    For rowIndex = LBound(values, 1) To UBound(values, 1)
        If IsSeriesRow(values, rowIndex, scope) Then
            graphId = CStr(values(rowIndex, RegistryGraphId))

            If graphId <> previousId Then
                If Not gr Is Nothing Then ApplyFormat gr, values, scope, previousId

                Set outCellRng = outsh.Range( _
                    CStr(values(rowIndex, RegistryOutRangeAddress)))
                Set gr = Graphs.Create(outsh, outCellRng)
                Application.GoTo outCellRng
                gr.Add
                previousId = graphId
            End If

            AddSeries gr, values, rowIndex
        End If
    Next rowIndex

    If Not gr Is Nothing Then ApplyFormat gr, values, scope, previousId
End Sub

'@label:is-series-row
'@sub-title Test whether a registry row describes a series of this scope
'@param values Variant. The registry rows.
'@param rowIndex Long. The row to test.
'@param scope Byte. The scope being written.
'@return Boolean. True for a series row of that scope.
Private Function IsSeriesRow(ByRef values As Variant, ByVal rowIndex As Long, _
                             ByVal scope As Byte) As Boolean
    If ReadLong(values(rowIndex, RegistryScope), 0) <> scope Then Exit Function
    IsSeriesRow = (CStr(values(rowIndex, RegistryKind)) = KIND_SERIES)
End Function

'@label:add-series
'@sub-title Read one series row and add it to a chart
'@param gr Graphs. The chart receiving the series.
'@param values Variant. The registry rows.
'@param rowIndex Long. The series row to read.
Private Sub AddSeries(ByVal gr As Graphs, ByRef values As Variant, _
                      ByVal rowIndex As Long)

    gr.AddSeries rngName:=CStr(values(rowIndex, RegistrySeriesName)), _
                 chrtType:=CStr(values(rowIndex, RegistrySeriesType)), _
                 axisPos:=CStr(values(rowIndex, RegistrySeriesPos))

    gr.AddLabels catName:=CStr(values(rowIndex, RegistrySeriesLabel)), _
                 lblName:=CStr(values(rowIndex, RegistrySeriesColumnLabel)), _
                 hardCodeLabels:=ReadBoolean(values(rowIndex, RegistryHardCodeLabels)), _
                 prefix:=CStr(values(rowIndex, RegistryPrefix)), _
                 prefixOnly:=ReadBoolean(values(rowIndex, RegistryPrefixOnly))
End Sub

'@label:apply-format
'@sub-title Apply the format row of one chart
'@details
'Finds the format row of the graph identifier inside the scope. A chart with
'no format row keeps the chart defaults.
'@param gr Graphs. The chart to format.
'@param values Variant. The registry rows.
'@param scope Byte. The scope being written.
'@param graphId String. The chart whose format row is wanted.
Private Sub ApplyFormat(ByVal gr As Graphs, ByRef values As Variant, _
                        ByVal scope As Byte, ByVal graphId As String)
    Dim rowIndex As Long
    Dim scopeValue As Byte

    'Graphs queues its series and draws them when it formats. A chart with no
    'format row is never formatted, so it is told to draw them here.
    gr.CommitSeries

    rowIndex = FormatRowIndex(values, scope, graphId)
    If rowIndex = 0 Then Exit Sub

    scopeValue = Switch( _
        (scope = AnalysisScopeTimeSeries Or scope = AnalysisScopeSpatioTemporal), _
            GraphScopeTimeSeries, _
        scope = AnalysisScopeSpatial, GraphScopeSpatial, _
        True, GraphScopeNormal)

    gr.Format valuesTitle:=CStr(values(rowIndex, RegistryValuesTitle)), _
              catTitle:=CStr(values(rowIndex, RegistryCatTitle)), _
              plotTitle:=CStr(values(rowIndex, RegistryPlotTitle)), _
              scope:=scopeValue, _
              heightFactor:=ReadLong(values(rowIndex, RegistryHeightFactor), 1), _
              hardCodeLabels:=ReadBoolean(values(rowIndex, RegistryHardCodeLabels))
End Sub

'@label:format-row-index
'@sub-title Locate the format row of one chart
'@param values Variant. The registry rows.
'@param scope Byte. The scope being written.
'@param graphId String. The chart to look for.
'@return Long. The row index, or 0 when the chart has no format row.
Private Function FormatRowIndex(ByRef values As Variant, ByVal scope As Byte, _
                                ByVal graphId As String) As Long
    Dim rowIndex As Long

    For rowIndex = LBound(values, 1) To UBound(values, 1)
        If ReadLong(values(rowIndex, RegistryScope), 0) = scope Then
            If CStr(values(rowIndex, RegistryKind)) = KIND_FORMAT Then
                If CStr(values(rowIndex, RegistryGraphId)) = graphId Then
                    FormatRowIndex = rowIndex
                    Exit Function
                End If
            End If
        End If
    Next rowIndex
End Function

'@section Typed Reads
'===============================================================================
'@description Turning a stored cell value back into a typed value, with no
'silent trap.

'@label:read-boolean
'@sub-title Read a stored flag
'@details
'The writers store a flag as 0 or 1, so the numeric arm answers every value
'this class has ever written. The text arms cover a value typed by hand into
'the sheet, and anything else answers False.
'@param value Variant. The cell value.
'@return Boolean. The stored flag.
Private Function ReadBoolean(ByVal value As Variant) As Boolean
    If IsEmpty(value) Then Exit Function

    If IsNumeric(value) Then
        ReadBoolean = (CDbl(value) <> 0)
        Exit Function
    End If

    Select Case LCase$(CStr(value))
    Case "true", "yes", "1"
        ReadBoolean = True
    End Select
End Function

'@label:read-long
'@sub-title Read a stored whole number
'@param value Variant. The cell value.
'@param defaultValue Long. What an empty or non-numeric cell answers.
'@return Long. The stored number, or the default.
Private Function ReadLong(ByVal value As Variant, ByVal defaultValue As Long) As Long
    ReadLong = defaultValue
    If IsEmpty(value) Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    ReadLong = CLng(value)
End Function

'@section Validation
'===============================================================================
'@description Existence checks and factory validation.

'@label:lo-exists
'@sub-title Check whether a ListObject exists on a worksheet
'@param sh Worksheet. The worksheet to search.
'@param loName String. The ListObject name to look for.
'@return Boolean. True when the ListObject exists.
Private Function LoExists(ByVal sh As Worksheet, ByVal loName As String) As Boolean
    Dim lo As ListObject

    'ListObjects answers a missing name by raising. The trap covers this one
    'call and the test that follows reads the outcome.
    On Error Resume Next
    Set lo = sh.ListObjects(loName)
    On Error GoTo 0

    LoExists = Not (lo Is Nothing)
End Function

'@label:check-requirements
'@sub-title Validate the tracking worksheet infrastructure
'@details
'Checks that the registry table is on the sheet. The export passes the sheet it
'copies from, so no analysis sheet name is looked up here.
'@param sh Worksheet. The tracking worksheet to validate.
'@throws ProjectError.InvalidArgument When the registry table is missing.
Private Sub CheckRequirements(ByVal sh As Worksheet)
    If Not LoExists(sh, REGISTRY_TABLE) Then _
        ThrowError ProjectError.InvalidArgument, _
                   "Table " & REGISTRY_TABLE & " is missing"
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
