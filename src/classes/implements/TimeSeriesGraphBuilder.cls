VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeSeriesGraphBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis.Graphs")
'@ModuleDescription("Builds graph chart specifications for complex time-series configurations")
'@depends ProjectError, GraphSeriesBuilder, GraphListObjectUtilities
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITimeSeriesGraphBuilder

Private Type TBuilderState
    cache As IGraphSpecsCache
    timeSeriesList As ListObject
    titleList As ListObject
    linelist As ILinelistSpecs
    seriesBuilder As IGraphSeriesBuilder
    specsFactory As ITimeSeriesTableSpecsFactory
    seriesRowIndex As Collection
End Type

Private this As TBuilderState
Private Const CLASS_NAME As String = "TimeSeriesGraphBuilder"

'@section Factory
'===============================================================================
Public Function Create(ByVal graphCache As IGraphSpecsCache, _
                       ByVal timeSeriesList As ListObject, _
                       ByVal titleList As ListObject, _
                       ByVal linelist As ILinelistSpecs, _
                       Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing, _
                       Optional ByVal specsFactory As ITimeSeriesTableSpecsFactory = Nothing) As ITimeSeriesGraphBuilder

    With New TimeSeriesGraphBuilder
        .Initialise graphCache, timeSeriesList, titleList, linelist, seriesBuilder, specsFactory
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ITimeSeriesGraphBuilder
    Set Self = Me
End Property

Public Sub Initialise(ByVal graphCache As IGraphSpecsCache, _
                      ByVal timeSeriesList As ListObject, _
                      ByVal titleList As ListObject, _
                      ByVal linelist As ILinelistSpecs, _
                      ByVal seriesBuilder As IGraphSeriesBuilder, _
                      ByVal specsFactory As ITimeSeriesTableSpecsFactory)

    ValidateCache graphCache, "graphCache"
    ValidateListObject timeSeriesList, "timeSeriesList"
    ValidateListObject titleList, "titleList"
    ValidateLinelist linelist, "linelist"

    Set this.cache = graphCache
    Set this.timeSeriesList = timeSeriesList
    Set this.titleList = titleList
    Set this.linelist = linelist

    If seriesBuilder Is Nothing Then
        Set this.seriesBuilder = GraphSeriesBuilder.Create
    Else
        Set this.seriesBuilder = seriesBuilder
    End If

    If specsFactory Is Nothing Then
        Set this.specsFactory = TimeSeriesTableSpecsFactory.CreateInstance
    Else
        Set this.specsFactory = specsFactory
    End If
End Sub

'@section Build operations
'===============================================================================
Public Function BuildGraph(ByVal graphId As String) As IGraphChartSpec
    Dim normalizedId As String
    Dim seriesIds As BetterArray
    Dim axisList As BetterArray
    Dim typeList As BetterArray
    Dim percentList As BetterArray
    Dim choiceList As BetterArray
    Dim labelList As BetterArray
    Dim idx As Long
    Dim seriesId As String
    Dim axisPosition As String
    Dim chartType As String
    Dim percentMode As String
    Dim choiceValue As String
    Dim labelPrefix As String
    Dim specs As ITablesSpecs
    Dim tableId As String
    Dim rowLabelName As String
    Dim labName As String
    Dim valuesRangeName As String
    Dim chartTitle As String

    normalizedId = NormalizeGraphKey(graphId)
    If normalizedId = vbNullString Then
        RaiseError CLng(ProjectError.InvalidArgument), "graphId cannot be empty"
    End If

    Set seriesIds = this.cache.ColumnValues(graphId, "series id")
    Set axisList = this.cache.ColumnValues(graphId, "axis")
    Set typeList = this.cache.ColumnValues(graphId, "type")
    Set percentList = this.cache.ColumnValues(graphId, "percentages")
    Set choiceList = this.cache.ColumnValues(graphId, "choices")
    Set labelList = this.cache.ColumnValues(graphId, "label")

    chartTitle = ResolveGraphTitle(graphId)
    this.seriesBuilder.Clear

    If seriesIds Is Nothing Or seriesIds.Length = 0 Then
        Set BuildGraph = this.seriesBuilder.BuildGraph(graphId, chartTitle)
        Exit Function
    End If

    For idx = seriesIds.LowerBound To seriesIds.UpperBound
        seriesId = CStr(seriesIds.Item(idx))
        axisPosition = SafeArrayValue(axisList, idx)
        chartType = SafeArrayValue(typeList, idx)
        percentMode = NormalizeGraphKey(SafeArrayValue(percentList, idx))
        choiceValue = SafeArrayValue(choiceList, idx)
        labelPrefix = SafeArrayValue(labelList, idx)

        Set specs = ResolveTableSpecs(seriesId)
        If specs Is Nothing Then GoTo ContinueLoop

        tableId = Trim$(specs.TableId)
        rowLabelName = ResolveRowLabelName(specs)
        labName = ResolveTimeSeriesColumnName(choiceValue, tableId)

        If labName = vbNullString Then GoTo ContinueLoop

        valuesRangeName = ResolveValuesRangeName(labName, percentMode, choiceValue, tableId)
        If valuesRangeName = vbNullString Then GoTo ContinueLoop

        If chartType = vbNullString Then chartType = "line"

        this.seriesBuilder.AddSeries valuesRangeName, chartType, axisPosition
        this.seriesBuilder.AssignLabels rowLabelName, labName, labelPrefix

ContinueLoop:
    Next idx

    Set BuildGraph = this.seriesBuilder.BuildGraph(graphId, chartTitle)
End Function

Public Function BuildAll() As BetterArray
    Dim graphIds As BetterArray
    Dim result As BetterArray
    Dim idx As Long
    Dim graphId As String
    Dim chartSpec As IGraphChartSpec

    Set graphIds = this.cache.GraphIds
    Set result = New BetterArray
    result.LowerBound = 1

    If graphIds Is Nothing Then
        Set BuildAll = result
        Exit Function
    End If

    For idx = graphIds.LowerBound To graphIds.UpperBound
        graphId = CStr(graphIds.Item(idx))
        Set chartSpec = BuildGraph(graphId)
        If Not chartSpec Is Nothing Then result.Push chartSpec
    Next idx

    Set BuildAll = result
End Function

Public Sub Invalidate()
    Set this.seriesRowIndex = Nothing
End Sub

'@section Internal helpers
'===============================================================================
Private Function ResolveGraphTitle(ByVal graphId As String) As String
    Dim titleRange As Range
    Dim idColumn As Range
    Dim hit As Range
    Dim title As String

    On Error Resume Next
        Set titleRange = this.titleList.ListColumns(3).DataBodyRange
    On Error GoTo 0

    If titleRange Is Nothing Then Exit Function

    Set idColumn = titleRange
    Set hit = idColumn.Find(What:=graphId, LookAt:=xlWhole, MatchCase:=True)

    If Not hit Is Nothing Then
        title = CStr(hit.Offset(0, -2).Value)
    End If

    ResolveGraphTitle = title
End Function

Private Function ResolveTableSpecs(ByVal seriesId As String) As ITablesSpecs
    Dim listRow As ListRow

    Set listRow = SeriesListRow(seriesId)
    If listRow Is Nothing Then Exit Function

    Set ResolveTableSpecs = this.specsFactory.Create(this.timeSeriesList.HeaderRowRange, listRow, this.linelist)
End Function

Private Function SeriesListRow(ByVal seriesId As String) As ListRow
    Dim key As String
    Dim rowIndex As Long

    key = NormalizeGraphKey(seriesId)
    If key = vbNullString Then Exit Function

    EnsureSeriesIndex

    On Error Resume Next
        rowIndex = CLng(this.seriesRowIndex.Item(key))
    On Error GoTo 0

    If rowIndex <= 0 Then Exit Function
    If rowIndex > this.timeSeriesList.ListRows.Count Then Exit Function

    Set SeriesListRow = this.timeSeriesList.ListRows(rowIndex)
End Function

Private Sub EnsureSeriesIndex()
    Dim dataRange As Range
    Dim rowIdx As Long
    Dim key As String

    If Not this.seriesRowIndex Is Nothing Then Exit Sub

    Set this.seriesRowIndex = New Collection

    On Error Resume Next
        Set dataRange = this.timeSeriesList.ListColumns(1).DataBodyRange
    On Error GoTo 0

    If dataRange Is Nothing Then Exit Sub

    For rowIdx = 1 To dataRange.Rows.Count
        key = NormalizeGraphKey(CStr(dataRange.Cells(rowIdx, 1).Value))
        If key <> vbNullString Then
            On Error Resume Next
                this.seriesRowIndex.Add CLng(rowIdx), key
            Err.Clear
            On Error GoTo 0
        End If
    Next rowIdx
End Sub

Private Function ResolveRowLabelName(ByVal specs As ITablesSpecs) As String
    Dim tableTypeCode As Byte
    Dim sectionId As String
    Dim tableId As String

    tableTypeCode = CByte(specs.TableType)
    sectionId = Trim$(specs.TableSectionId)
    tableId = Trim$(specs.TableId)

    If (tableTypeCode = TABLE_TYPE_SPATIO_TEMPORAL) And (sectionId <> vbNullString) Then
        ResolveRowLabelName = "ROW_CATEGORIES_" & sectionId
    Else
        ResolveRowLabelName = "ROW_CATEGORIES_" & tableId
    End If
End Function

Private Function ResolveTimeSeriesColumnName(ByVal choiceValue As String, ByVal tableId As String) As String
    Dim sheet As Worksheet
    Dim categoryRange As Range
    Dim hit As Range
    Dim rangeName As String

    If Trim$(choiceValue) = vbNullString Then Exit Function

    If NormalizeGraphKey(choiceValue) = "total" Then
        ResolveTimeSeriesColumnName = "TOTAL_LABEL_COL_" & tableId
        Exit Function
    End If

    Set sheet = this.timeSeriesList.Parent

    On Error Resume Next
        Set categoryRange = sheet.Range("COLUMN_CATEGORIES_" & tableId)
    On Error GoTo 0

    If categoryRange Is Nothing Then Exit Function

    Set hit = categoryRange.Find(What:=choiceValue, LookAt:=xlWhole, MatchCase:=True)
    If hit Is Nothing Then Exit Function

    On Error Resume Next
        rangeName = hit.Name.Name
    On Error GoTo 0

    ResolveTimeSeriesColumnName = rangeName
End Function

Private Function ResolveValuesRangeName(ByVal labelRangeName As String, _
                                        ByVal percentMode As String, _
                                        ByVal choiceValue As String, _
                                        ByVal tableId As String) As String

    If percentMode = "percentages" Then
        If NormalizeGraphKey(choiceValue) = "total" Then
            ResolveValuesRangeName = Replace(labelRangeName, "COL", "PERC")
        Else
            ResolveValuesRangeName = Replace(labelRangeName, "LABEL", "PERC")
        End If
        Exit Function
    End If

    If NormalizeGraphKey(choiceValue) = "total" Then
        ResolveValuesRangeName = "TOTAL_COL_VALUES_" & tableId
        Exit Function
    End If

    If InStr(1, labelRangeName, "COLUMN_CATEGORIES_", vbTextCompare) > 0 Then
        ResolveValuesRangeName = "INTERIOR_VALUES_" & tableId
    Else
        ResolveValuesRangeName = Replace(labelRangeName, "LABEL", "VALUES")
    End If
End Function

Private Function SafeArrayValue(ByVal values As BetterArray, ByVal index As Long) As String
    If values Is Nothing Then Exit Function
    If values.Length = 0 Then Exit Function
    If index < values.LowerBound Or index > values.UpperBound Then Exit Function
    SafeArrayValue = CStr(values.Item(index))
End Function

Private Sub ValidateCache(ByVal graphCache As IGraphSpecsCache, ByVal argumentName As String)
    If graphCache Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub ValidateListObject(ByVal candidate As ListObject, ByVal argumentName As String)
    If candidate Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
        Exit Sub
    End If

    If StrComp(TypeName(candidate), "ListObject", vbBinaryCompare) <> 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " must be a ListObject"
    End If
End Sub

Private Sub ValidateLinelist(ByVal linelist As ILinelistSpecs, ByVal argumentName As String)
    If linelist Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASS_NAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function ITimeSeriesGraphBuilder_Create(ByVal graphCache As IGraphSpecsCache, _
                                                ByVal timeSeriesList As ListObject, _
                                                ByVal titleList As ListObject, _
                                                ByVal linelist As ILinelistSpecs, _
                                                Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing, _
                                                Optional ByVal specsFactory As ITimeSeriesTableSpecsFactory = Nothing) As ITimeSeriesGraphBuilder
    Set ITimeSeriesGraphBuilder_Create = Create(graphCache, timeSeriesList, titleList, linelist, seriesBuilder, specsFactory)
End Function

Private Property Get ITimeSeriesGraphBuilder_Self() As ITimeSeriesGraphBuilder
    Set ITimeSeriesGraphBuilder_Self = Self
End Property

Private Function ITimeSeriesGraphBuilder_BuildGraph(ByVal graphId As String) As IGraphChartSpec
    Set ITimeSeriesGraphBuilder_BuildGraph = BuildGraph(graphId)
End Function

Private Function ITimeSeriesGraphBuilder_BuildAll() As BetterArray
    Set ITimeSeriesGraphBuilder_BuildAll = BuildAll()
End Function

Private Sub ITimeSeriesGraphBuilder_Invalidate()
    Invalidate
End Sub

