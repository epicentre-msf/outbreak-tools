VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeSeriesGraphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Chart definitions for every time series graph"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Chart definitions for every time series graph")
'@IgnoreModule IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class TimeSeriesGraphs
'@description
'The chart definitions of every time series graph, read from the three graph
'setup tables. One entry per graph identifier, each holding its series and its
'display title. Count answers how many graphs there are, Series gives the
'SeriesBuffer of one graph and Title its title. AnalysisOutput walks the
'entries, draws one chart each and puts the title in the cell above it.
'
'THE THREE SETUP TABLES ARE IDENTIFIED BY NAME
'-------------------------------------------------------------------------------
'The three names are byte-identical in the six workbooks measured on
'2026-07-31, so this class compares names. An earlier version read a caption a
'fixed number of rows above each header row. That gap is four rows in
'.mock/setup_mock.xlsb, src/bin/setup/setup.xlsb and setup_dev.xlsb, and two
'rows in every file under releases/.
'
'THE GRAPHS ARE BUILT ON THE FIRST READ
'-------------------------------------------------------------------------------
'Count builds the entries the first time it is asked and holds the answer, and
'so do Series and Title. A Count of 0 says nothing could be drawn, and the
'reasons are in the report. HasCheckings is read after Count.
'
'EVERY SKIP IS REPORTED
'-------------------------------------------------------------------------------
'This class builds fewer series than the setup asked for in eight cases: a
'setup table with a column missing, a graph with no series row, a series row
'naming no series, a series identifier absent from the time series table, a
'series matching the header row, a choice value with no column on the output
'sheet, a column carrying no named range, and a graph identifier with no title
'row. A graph that ends with no series is left out and named.
'@depends SeriesBuffer, TableSpecs, LinelistSpecs, Checking, BetterArray
'@version 1.0 (2026-07-31)



'@section State
'===============================================================================

Private Type TTimeSeriesGraphs
    graphLo As ListObject
    timeSeriesLo As ListObject
    titlesLo As ListObject
    outSh As Worksheet
    lData As LinelistSpecs
    ' One entry per graph. The two lists are appended in PushGraph and nowhere
    ' else, so index N of one always describes the same graph as index N of the
    ' other.
    graphBuffers As BetterArray
    graphTitles As BetterArray
    built As Boolean
    sealed As Boolean
    checkCounter As Long
    internalChecks As Checking
End Type

Private this As TTimeSeriesGraphs

Private Const CLASS_NAME As String = "TimeSeriesGraphs"

' The three setup tables, in the order Create expects them. The names are the
' ones the setup workbook gives the tables and they are identical across the
' development files and every release.
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

' The choice value that asks for the total column, and the setting that asks
' for percentages.
Private Const CHOICE_TOTAL As String = "total"
Private Const MODE_PERCENTAGES As String = "percentages"


'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create the time series graph builder from the three setup tables
'@details
'The three ListObjects come in the order the validation checks: the graph setup
'table, which says which series belong to each graph identifier; the time series
'table, which maps a series identifier to its table specification row; and the
'titles table, which gives a display title per graph identifier.
'@param loTable BetterArray. Exactly 3 ListObject entries in order:
'[Tab_Graph_TimeSeries, Tab_TimeSeries_Analysis, Tab_Label_TSGraph]. Each must
'carry at least one data row.
'@param outputSh Worksheet. The worksheet the time series charts are drawn on,
'which is also where the column category ranges are read from.
'@param lData LinelistSpecs. The linelist specifications, read for the
'dictionary that TableSpecs needs.
'@return TimeSeriesGraphs. A builder ready for Count.
'@throws ProjectError.InvalidArgument When loTable fails validation, lData is
'Nothing, or outputSh is Nothing.
'@throws ProjectError.ErrorUnexpectedState When the ListObjects are in the wrong
'order.
'@export
Public Function Create(ByVal loTable As BetterArray, _
                       ByVal outputSh As Worksheet, _
                       ByVal lData As LinelistSpecs) As TimeSeriesGraphs

    ValidateListObjects loTable
    If lData Is Nothing Then ThrowError InvalidArgument, "Linelist specifications not defined"
    If outputSh Is Nothing Then ThrowError InvalidArgument, "Output worksheet not defined"

    Dim newGraphs As TimeSeriesGraphs
    Set newGraphs = New TimeSeriesGraphs
    Set newGraphs.LoList = loTable
    Set newGraphs.OutputSheet = outputSh
    Set newGraphs.LinelistSpecifications = lData
    newGraphs.Seal
    Set Create = newGraphs
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

'@label:lo-list-set
'@prop-title Assign the three setup tables
'@details
'Resolves the three tables into fields of their own. They are read inside loops,
'and each read used to go through the array again.
'@param loTable BetterArray. The 3-element ListObject array.
Public Property Set LoList(ByVal loTable As BetterArray)
    GuardNotSealed "LoList"
    Set this.graphLo = loTable.Item(loTable.LowerBound)
    Set this.timeSeriesLo = loTable.Item(loTable.LowerBound + 1)
    Set this.titlesLo = loTable.Item(loTable.LowerBound + 2)
End Property

'@label:output-sheet-set
'@prop-title Assign the output worksheet
'@param sh Worksheet. The worksheet to store.
Public Property Set OutputSheet(ByVal sh As Worksheet)
    GuardNotSealed "OutputSheet"
    Set this.outSh = sh
End Property

'@label:linelist-specifications-set
'@prop-title Assign the linelist specifications
'@param lData LinelistSpecs. The linelist specifications to store.
Public Property Set LinelistSpecifications(ByVal lData As LinelistSpecs)
    GuardNotSealed "LinelistSpecifications"
    Set this.lData = lData
End Property


'@section Reading
'===============================================================================

'@label:count
'@prop-title How many charts the setup tables ask for
'@details
'Builds on the first read and keeps the answer. A graph that ended with no
'series is left out, so this counts the charts that can actually be drawn.
'@return Long. The number of graphs.
'@export
Public Property Get Count() As Long
    If Not this.built Then BuildGraphs
    Count = this.graphBuffers.Length
End Property

'@label:series
'@prop-title The series of one graph
'@param index Long. 1-based position, from 1 to Count.
'@return SeriesBuffer. One entry per series, in plot order.
'@throws ProjectError.InvalidArgument When index is outside the collection.
'@export
Public Property Get Series(ByVal index As Long) As SeriesBuffer
    ValidateIndex index
    Set Series = this.graphBuffers.Item(index)
End Property

'@label:title
'@prop-title The display title of one graph
'@param index Long. 1-based position, from 1 to Count.
'@return String. The title, or an empty string when the graph has no title row.
'@throws ProjectError.InvalidArgument When index is outside the collection.
'@export
Public Property Get Title(ByVal index As Long) As String
    ValidateIndex index
    Title = CStr(this.graphTitles.Item(index))
End Property

'@label:wksh
'@prop-title The worksheet the charts are drawn on
'@return Worksheet. The output worksheet.
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.outSh
End Property


'@section Building
'===============================================================================

'@label:build-graphs
'@sub-title Build one entry per graph identifier
'@details
'Walks the graph identifiers of the setup table in the order they first appear.
'A graph that ends with no series is reported and left out, because a chart with
'no series has no axes and Graphs.Format used to raise 1004 on one.
Private Sub BuildGraphs()
    Dim graphIds As BetterArray
    Dim graphId As String
    Dim buffer As SeriesBuffer
    Dim counter As Long

    this.built = True
    Set this.graphBuffers = New BetterArray
    this.graphBuffers.LowerBound = 1
    Set this.graphTitles = New BetterArray
    this.graphTitles.LowerBound = 1

    Set graphIds = BuildGraphIdsList()

    If graphIds.Length = 0 Then
        LogInfo "The graph setup table holds no graph identifier, so no time series " & _
                "chart was built.", checkingWarning
        Exit Sub
    End If

    For counter = graphIds.LowerBound To graphIds.UpperBound
        graphId = graphIds.Item(counter)
        Set buffer = DefineGraph(graphId)

        If buffer.Count = 0 Then
            LogInfo "Graph " & graphId & " ended with no series, so no chart was " & _
                    "drawn for it.", checkingWarning
        Else
            PushGraph buffer, GraphTitle(graphId)
        End If
    Next
End Sub

'@label:push-graph
'@sub-title Append one finished graph to the collection
'@details
'The buffer and the title are appended together, in this one routine, so the two
'lists cannot come apart.
'@param buffer SeriesBuffer. The series of the graph.
'@param graphTitle String. Its display title.
Private Sub PushGraph(ByVal buffer As SeriesBuffer, ByVal graphTitle As String)
    this.graphBuffers.Push buffer
    this.graphTitles.Push graphTitle
End Sub

'@label:define-graph
'@sub-title Build the series of one graph identifier
'@details
'Reads the six columns of the graph setup table for the rows carrying this
'identifier: the series identifiers, the axis sides, the percentage overrides,
'the chart types, the choice values and the display labels. Each series
'identifier is looked up in the time series table to get a TableSpecs, which
'gives the table identifier and its section. The choice value resolves to a
'column name through TimeSeriesColumnName, and that to a data range name through
'ResolveSeriesRangeName. A row asking for percentages takes the percentage twin
'of that range.
'@param graphId String. The graph identifier the rows are filtered by.
'@return SeriesBuffer. The series of that graph, empty when none resolved.
'@depends TableSpecs
Private Function DefineGraph(ByVal graphId As String) As SeriesBuffer
    Dim buffer As SeriesBuffer
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

    Set buffer = SeriesBuffer.Create()
    Set DefineGraph = buffer

    Set tsLo = this.timeSeriesLo
    Set hRng = tsLo.HeaderRowRange

    Set idsList = GraphValues(colName:=COLUMN_SERIES_ID, graphId:=graphId)
    Set posList = GraphValues(colName:=COLUMN_AXIS, graphId:=graphId)
    Set percList = GraphValues(colName:=COLUMN_PERCENTAGES, graphId:=graphId)
    Set typeList = GraphValues(colName:=COLUMN_TYPE, graphId:=graphId)
    Set choiceList = GraphValues(colName:=COLUMN_CHOICES, graphId:=graphId)
    Set labelList = GraphValues(colName:=COLUMN_LABEL, graphId:=graphId)

    If idsList.Length = 0 Then
        LogInfo "Graph " & graphId & " has no series row in the graph setup table.", _
                checkingWarning
        Exit Function
    End If

    ' The six lists are read row by row over the same table, so they answer the
    ' same count unless one of the columns is missing. A short list would make
    ' the reads below fall off the end of an array.
    If (posList.Length <> idsList.Length) Or (percList.Length <> idsList.Length) Or _
       (typeList.Length <> idsList.Length) Or (choiceList.Length <> idsList.Length) Or _
       (labelList.Length <> idsList.Length) Then
        LogInfo "Graph " & graphId & " could not be read: the columns of the graph " & _
                "setup table do not all answer the same number of rows.", checkingError
        Exit Function
    End If

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
        Set specs = TableSpecs.Create(hRng, tabRng, this.lData.Dictionary())
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

        buffer.Add rangeName:=seriesRangeName, _
                   chartType:=CStr(typeList.Item(counter)), _
                   axisSide:=CStr(posList.Item(counter)), _
                   categoryRange:="ROW_CATEGORIES_" & secId, _
                   legendRange:=labName, _
                   labelPrefix:=CStr(labelList.Item(counter))

ContinueLoop:
    Next
End Function

'@label:resolve-series-range-name
'@sub-title Resolve a choice value and label name into the data range name
'@details
'Three rules. The total answers TOTAL_COL_VALUES_ and the table identifier. A
'label name carrying the COLUMN_CATEGORIES_ prefix answers INTERIOR_VALUES_ and
'the table identifier. Everything else replaces LABEL with VALUES in the label
'name, so "LABEL_COL_2_xyz" becomes "VALUES_COL_2_xyz".
'@param choiValue String. The choice value from the graph setup table.
'@param tabId String. The table identifier every named range carries.
'@param labName String. The label range name TimeSeriesColumnName resolved.
'@return String. The named range holding the data of the series.
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
'@sub-title Find the named range of a choice value inside a time series table
'@details
'Looks the choice value up in the COLUMN_CATEGORIES_ range of the table on the
'output worksheet and answers the name of the cell that carries it. The total
'answers TOTAL_LABEL_COL_ and the table identifier directly. An empty answer
'means the series is skipped, and every route to one files an entry.
'@param choiValue String. The category value to look up.
'@param tabId String. The table identifier every named range carries.
'@param graphId String. The graph being built, named in the report entries.
'@return String. The named range of the matching column, or an empty string.
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


'@section SetupTables
'===============================================================================
'@description Reading the three graph setup ListObjects.

'@label:graph-title
'@sub-title Look up the display title of a graph identifier
'@details
'Walks the identifier column of the titles table and answers the title cell of
'the matching row. Both columns are found by header name, and the identifier is
'matched without regard to case and outer spaces, because the two tables are
'typed by two hands and a case difference used to give a chart with an empty
'title and a navigation entry reading "Go to graph: ".
'@param graphId String. The graph identifier to look up.
'@return String. The display title, or an empty string when there is no row.
Private Function GraphTitle(ByVal graphId As String) As String
    Dim titleLo As ListObject
    Dim bodyRng As Range
    Dim idIndex As Long
    Dim titleIndex As Long
    Dim counter As Long

    Set titleLo = this.titlesLo
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
'@sub-title Read the graph identifiers of the setup table
'@details
'Walks the identifier column, keeps every value it has not seen, and holds the
'order of first appearance.
'@return BetterArray. The identifiers, LowerBound = 1. Empty when the table has
'no data row or no identifier column.
'@remarks The duplicate test uses BetterArray.Includes, which scans. The graph
'setup table carries fewer than a hundred rows.
Private Function BuildGraphIdsList() As BetterArray
    Dim idsList As BetterArray
    Dim bodyRng As Range
    Dim idIndex As Long
    Dim counter As Long
    Dim idValue As String

    Set idsList = New BetterArray
    idsList.LowerBound = 1
    Set BuildGraphIdsList = idsList

    Set bodyRng = this.graphLo.DataBodyRange
    If bodyRng Is Nothing Then Exit Function

    idIndex = LoColumnIndex(this.graphLo, COLUMN_GRAPH_ID)
    If idIndex < 0 Then
        LogInfo "The graph setup table " & this.graphLo.Name & " has no """ & _
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
'@sub-title Read one column of the graph setup table, for one graph
'@details
'Answers the values of one column, for the rows whose graph identifier matches,
'in row order. The walk reads the identifier column and the wanted column of the
'same row, so the six lists DefineGraph reads always line up by index. An empty
'cell answers an empty string.
'@param colName String. Header of the column to read.
'@param graphId String. The graph identifier the rows are filtered by.
'@return BetterArray. The values, LowerBound = 1. Empty when the table has no
'data row, no identifier column, or no column of that name.
'@remarks This used to be driven by Range.Find with a search range narrowed on
'every iteration. Find wraps around by default, and none of the calls named
'LookIn, so the search inherited whatever the last search in the application
'used, the user's own Ctrl+F dialog included.
Private Function GraphValues(ByVal colName As String, _
                             ByVal graphId As String) As BetterArray
    Dim valuesList As BetterArray
    Dim idIndex As Long
    Dim colIndex As Long
    Dim bodyRng As Range
    Dim counter As Long

    Set valuesList = New BetterArray
    valuesList.LowerBound = 1
    Set GraphValues = valuesList

    Set bodyRng = this.graphLo.DataBodyRange
    If bodyRng Is Nothing Then Exit Function

    idIndex = LoColumnIndex(this.graphLo, COLUMN_GRAPH_ID)
    If idIndex < 0 Then
        LogInfo "The graph setup table " & this.graphLo.Name & " has no """ & _
                COLUMN_GRAPH_ID & """ column, so graph " & graphId & " was skipped.", _
                checkingError
        Exit Function
    End If

    colIndex = LoColumnIndex(this.graphLo, colName)
    If colIndex < 0 Then
        LogInfo "The graph setup table " & this.graphLo.Name & " has no """ & colName & _
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
'"Y-Axis" and "type" find "Chart type". When more than one header carries the
'term the first is used and the ambiguity is reported, because a renamed or an
'added column would otherwise change what the chart plots with nothing said.
'@param lo ListObject. The ListObject whose header row is searched.
'@param colName String. The header text to search for. Matched without regard to
'case or outer spaces.
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

'@label:validate-list-objects
'@sub-title Validate the three ListObjects the factory is given
'@details
'Three levels. The array holds exactly three entries. Each entry is a ListObject
'carrying at least one data row. And each one is the table its position expects,
'which is settled by comparing its name against Tab_Graph_TimeSeries,
'Tab_TimeSeries_Analysis and Tab_Label_TSGraph.
'@param loTable BetterArray. The ListObject collection to validate.
'@throws ProjectError.InvalidArgument When loTable is Nothing, holds the wrong
'number of entries, holds something that is not a ListObject, or holds an empty
'one.
'@throws ProjectError.ErrorUnexpectedState When the ListObjects are in the wrong
'order.
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

'@label:validate-index
'@sub-title Refuse a graph position the collection does not hold
'@param index Long. The position to test.
'@throws ProjectError.InvalidArgument When index is outside the collection.
Private Sub ValidateIndex(ByVal index As Long)
    If (index < 1) Or (index > Count()) Then _
        ThrowError ProjectError.InvalidArgument, _
                   "Graph index " & index & " out of bounds (1 to " & Count() & ")"
End Sub


'@section Checking
'===============================================================================
'@description The report of everything this builder decided to leave out.

'@label:specs-id
'@sub-title A name for this instance that no other instance shares
'@details
'AnalysisOutput merges the entries of every collaborator into one report, and
'Checking.Add raises on a duplicate key, so a key made of a bare counter
'collides as soon as two instances file an entry. One instance serves one output
'worksheet, so the sheet names it.
'@return String. The identifier every key of this instance carries.
Private Function SpecsId() As String
    If Not this.outSh Is Nothing Then SpecsId = this.outSh.Name
    If LenB(SpecsId) = 0 Then SpecsId = "unplaced"
End Function

'@label:log-info
'@sub-title Record a diagnostic entry in the internal checking store
'@details
'Creates the Checking store on first use and appends the message under a key
'naming this class, this worksheet and a counter.
'@param label String. The message to record.
'@param scope Optional Byte. Severity from CheckingLogType. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)
    Dim specsKey As String

    specsKey = SpecsId()

    If this.internalChecks Is Nothing Then _
        Set this.internalChecks = Checking.Create("Time series graphs of " & specsKey)

    this.internalChecks.Add CLASS_NAME & "-" & specsKey & "-" & CStr(this.checkCounter), _
                            label, scope
    this.checkCounter = this.checkCounter + 1
End Sub

'@label:has-checkings
'@prop-title Whether anything was reported while the graphs were built
'@return Boolean. True when at least one entry was filed.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = Not (this.internalChecks Is Nothing)
End Property

'@label:checking-values
'@prop-title The report of this build
'@return Checking. The store, or Nothing when nothing was filed.
'@export
Public Property Get CheckingValues() As Checking
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = this.internalChecks
End Property


'@section ErrorHandling
'===============================================================================

'@label:cell-text
'@sub-title Read one cell value as text
'@details
'CStr on an Error variant raises runtime error 13, and the analysis setup sheet
'carries formula columns, so an ordinary edit can produce one. An error value, a
'Null and an empty cell all answer an empty string here.
'@param cellValue Variant. One value read out of a worksheet range.
'@return String. The text of the cell, or an empty string.
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

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber Long. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
