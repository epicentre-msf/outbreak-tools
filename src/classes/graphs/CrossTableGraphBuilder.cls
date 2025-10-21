VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableGraphBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis.Graphs")
'@ModuleDescription("Builds chart specifications for non time-series analysis tables")
'@depends ProjectError, GraphSeriesBuilder
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ICrossTableGraphBuilder

Private Type TBuilderState
    seriesBuilder As IGraphSeriesBuilder
End Type

Private this As TBuilderState
Private Const CLASS_NAME As String = "CrossTableGraphBuilder"

'@section Factory
'===============================================================================
Public Function Create(Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing) As ICrossTableGraphBuilder
    With New CrossTableGraphBuilder
        .Initialise seriesBuilder
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ICrossTableGraphBuilder
    Set Self = Me
End Property

Public Sub Initialise(Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing)
    If seriesBuilder Is Nothing Then
        Set this.seriesBuilder = GraphSeriesBuilder.Create
    Else
        Set this.seriesBuilder = seriesBuilder
    End If
End Sub

'@section Build operations
'===============================================================================
Public Function BuildGraph(ByVal table As ICrossTable, _
                           ByVal graphId As String, _
                           Optional ByVal graphTitle As String = vbNullString) As IGraphChartSpec
    Dim specs As ITablesSpecs
    Dim chart As IGraphChartSpec
    Dim tableTypeCode As Byte

    ValidateGraphTarget table, graphId

    Set specs = table.Specifications
    If specs Is Nothing Then
        ThrowError CLng(ProjectError.ObjectNotInitialized), "Table specifications are required"
        Exit Function
    End If

    tableTypeCode = CByte(specs.TableType)

    this.seriesBuilder.Clear

    Select Case tableTypeCode
        Case TABLE_TYPE_UNIVARIATE
            BuildUnivariateSeries table, specs

        Case TABLE_TYPE_BIVARIATE, TABLE_TYPE_SPATIAL, TABLE_TYPE_SPATIO_TEMPORAL
            BuildMultiColumnSeries table, specs

        Case Else
            ThrowError CLng(ProjectError.InvalidArgument), _
                      "Unsupported table type for cross table graph: " & CStr(tableTypeCode)
    End Select

    Set chart = this.seriesBuilder.BuildGraph(graphId, graphTitle)
    Set BuildGraph = chart
End Function

Public Property Get SeriesBuilder() As IGraphSeriesBuilder
    Set SeriesBuilder = this.seriesBuilder
End Property

'@section Series builders
'===============================================================================
Private Sub BuildUnivariateSeries(ByVal table As ICrossTable, ByVal specs As ITablesSpecs)
    Dim tableId As String
    Dim plotType As String
    Dim rowLabelName As String
    Dim columnLabelName As String

    tableId = Trim$(specs.TableId)
    plotType = DeterminePlotType(specs)
    rowLabelName = "ROW_CATEGORIES_" & tableId
    columnLabelName = "LABEL_COL_1_" & tableId

    AddSeriesWithLabels "VALUES_COL_1_" & tableId, plotType, "left", _
                        rowLabelName, columnLabelName

    If specs.HasPercentage Then
        If NormalizeGraphKey(plotType) = "bar" Then
            AddSeriesWithLabels "PERC_COL_1_" & tableId, "point", "right", _
                                rowLabelName, "PERC_LABEL_COL_" & tableId
        End If
    End If
End Sub

Private Sub BuildMultiColumnSeries(ByVal table As ICrossTable, ByVal specs As ITablesSpecs)
    Dim tableId As String
    Dim tableSectionId As String
    Dim plotType As String
    Dim nbCols As Long
    Dim columnIndex As Long
    Dim modeValue As String
    Dim isSpatioTemporal As Boolean

    tableId = Trim$(specs.TableId)
    tableSectionId = Trim$(specs.TableSectionId)
    plotType = DeterminePlotType(specs)
    nbCols = ResolveColumnCount(table, specs)
    modeValue = NormalizeGraphKey(specs.Value("graph"))
    isSpatioTemporal = (CByte(specs.TableType) = TABLE_TYPE_SPATIO_TEMPORAL)

    Select Case modeValue
        Case "values", "yes"
            For columnIndex = 1 To nbCols
                AppendValuesSeries columnIndex, tableId, tableSectionId, plotType, isSpatioTemporal
            Next columnIndex

        Case "percentages"
            If specs.HasPercentage Then
                For columnIndex = 1 To nbCols
                    AppendPercentSeries columnIndex, tableId, tableSectionId, plotType, isSpatioTemporal
                Next columnIndex
            End If

        Case "both"
            For columnIndex = 1 To nbCols
                AppendValuesSeries columnIndex, tableId, tableSectionId, plotType, isSpatioTemporal
                If specs.HasPercentage Then
                    AppendPercentSeries columnIndex, tableId, tableSectionId, "point", isSpatioTemporal
                End If
            Next columnIndex

        Case Else
            'Maintain historical behaviour by emitting value series when configuration is empty.
            For columnIndex = 1 To nbCols
                AppendValuesSeries columnIndex, tableId, tableSectionId, plotType, isSpatioTemporal
            Next columnIndex
    End Select
End Sub

Private Sub AppendValuesSeries(ByVal columnIndex As Long, _
                               ByVal tableId As String, _
                               ByVal tableSectionId As String, _
                               ByVal plotType As String, _
                               ByVal isSpatioTemporal As Boolean)
    Dim rowLabel As String

    rowLabel = ResolveRowLabelName(tableId, tableSectionId, isSpatioTemporal)

    AddSeriesWithLabels "VALUES_COL_" & columnIndex & "_" & tableId, _
                        plotType, "left", _
                        rowLabel, _
                        "LABEL_COL_" & columnIndex & "_" & tableId
End Sub

Private Sub AppendPercentSeries(ByVal columnIndex As Long, _
                                ByVal tableId As String, _
                                ByVal tableSectionId As String, _
                                ByVal plotType As String, _
                                ByVal isSpatioTemporal As Boolean)
    Dim rowLabel As String

    rowLabel = ResolveRowLabelName(tableId, tableSectionId, isSpatioTemporal)

    AddSeriesWithLabels "PERC_COL_" & columnIndex & "_" & tableId, _
                        plotType, "left", _
                        rowLabel, _
                        "LABEL_COL_" & columnIndex & "_" & tableId
End Sub

Private Function ResolveRowLabelName(ByVal tableId As String, _
                                     ByVal tableSectionId As String, _
                                     ByVal isSpatioTemporal As Boolean) As String
    If isSpatioTemporal And (Trim$(tableSectionId) <> vbNullString) Then
        ResolveRowLabelName = "ROW_CATEGORIES_" & tableSectionId
    Else
        ResolveRowLabelName = "ROW_CATEGORIES_" & tableId
    End If
End Function

Private Sub AddSeriesWithLabels(ByVal seriesName As String, _
                                ByVal seriesType As String, _
                                ByVal axisPosition As String, _
                                ByVal rowLabel As String, _
                                ByVal columnLabel As String, _
                                Optional ByVal labelPrefix As String = vbNullString)
    this.seriesBuilder.AddSeries seriesName, seriesType, axisPosition
    this.seriesBuilder.AssignLabels rowLabel, columnLabel, labelPrefix
End Sub

Private Function DeterminePlotType(ByVal specs As ITablesSpecs) As String
    Dim flipValue As String

    flipValue = NormalizeGraphKey(specs.Value("flip"))
    If flipValue = "yes" Then
        DeterminePlotType = "hbar"
    Else
        DeterminePlotType = "bar"
    End If
End Function

Private Function ResolveColumnCount(ByVal table As ICrossTable, ByVal specs As ITablesSpecs) As Long
    Dim defaultCount As Long
    Dim overrideText As String
    Dim overrideValue As Long

    defaultCount = table.NumberOfColumns

    If CByte(specs.TableType) = TABLE_TYPE_SPATIO_TEMPORAL Then
        ResolveColumnCount = 5
        overrideText = Trim$(specs.Value("n geo"))

        If overrideText <> vbNullString Then
            On Error Resume Next
                overrideValue = CLng(overrideText)
            If Err.Number = 0 Then ResolveColumnCount = overrideValue
            Err.Clear
            On Error GoTo 0
        End If
    Else
        ResolveColumnCount = defaultCount
    End If

    If ResolveColumnCount <= 0 Then ResolveColumnCount = defaultCount
End Function

Private Sub ValidateGraphTarget(ByVal table As ICrossTable, ByVal graphId As String)
    If table Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), "table cannot be Nothing"
        Exit Sub
    End If

    If NormalizeGraphKey(graphId) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), "graphId cannot be empty"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASS_NAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function ICrossTableGraphBuilder_Create(Optional ByVal seriesBuilder As IGraphSeriesBuilder = Nothing) As ICrossTableGraphBuilder
    Set ICrossTableGraphBuilder_Create = Create(seriesBuilder)
End Function

Private Property Get ICrossTableGraphBuilder_Self() As ICrossTableGraphBuilder
    Set ICrossTableGraphBuilder_Self = Self
End Property

Private Function ICrossTableGraphBuilder_BuildGraph(ByVal table As ICrossTable, _
                                                    ByVal graphId As String, _
                                                    Optional ByVal graphTitle As String = vbNullString) As IGraphChartSpec
    Set ICrossTableGraphBuilder_BuildGraph = BuildGraph(table, graphId, graphTitle)
End Function

Private Property Get ICrossTableGraphBuilder_SeriesBuilder() As IGraphSeriesBuilder
    Set ICrossTableGraphBuilder_SeriesBuilder = SeriesBuilder
End Property

