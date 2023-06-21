VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Add specifications for graphs in analysis"

'@Folder("Analysis")
'@ModuleDescription("Add specifications for graphs in analysis")

Option Explicit
Implements IGraphSpecs

'Constants for names of tables in range specs
Private Const TABONTIMESERIES = "time series analysis"
Private Const TABONGRAPHTIMESERIES = "graph on time series"
Private Const TABONTITLEGRAPHSTIMESERIES = "labels for time series graphs"

Public Enum GraphSpecsType
    GraphTypeSimple = 1
    GraphTypeComplex = 2
End Enum

'Will have two instanciation methods (one like a tablespecs, another which
'norma and reserved to most cases (bivariate, spatial))
Private Type TGraphSpecs
    testDone As Boolean
    validity As Boolean
    tabl As ICrossTable
    timelos As BetterArray
    specstype As GraphSpecsType
    outsh As Worksheet
    chartSeriesName As BetterArray
    chartSeriesType As BetterArray
    chartSeriesPosition As BetterArray
    chartSeriesLabel As BetterArray
    chartSeriesLabelPrefix As BetterArray
    chartColumnCategoriesLabel As BetterArray
    lData As ILinelistSpecs
    graphsChartsSpecs As BetterArray             'A BetterArray of BetterArray all with length 5 of same length, each one for series, Type, label and category label
    GraphIdsList As BetterArray
End Type

Private this As TGraphSpecs
Private Const CLASSNAME As String = "GraphSpecs"

'You can create graph specs in two various way: First one is to give a crosstable (for bivariate, univariate and graphs);
'Second one is to give a BetterArray of two listObjects, the output sheet and the linelist Data (For graphs for time series)

Public Function Create(ByVal tabl As ICrossTable) As IGraphSpecs

    If tabl Is Nothing Then ThrowErrorUnableInstanciate "Table does not exist"
    If tabl.Specifications.TableType = TypeTimeSeries Then
        ThrowErrorUnableInstanciate "Simple form instanciations are only for non time series table. For time series, use long one"
    End If

    If tabl.Specifications.TableType = TypeGlobalSummary Then
        ThrowErrorUnableInstanciate "I don't know how to build graphs on global summary tables"
    End If

    With New GraphSpecs
        Set .Table = tabl
        .GraphType = GraphTypeSimple
        Set Create = .Self
    End With
End Function

Public Function CreateRangeSpecs(ByVal loTable As BetterArray, ByVal OutputSh As Worksheet, ByVal lData As ILinelistSpecs) As IGraphSpecs
    'loTable is a table of two listObjects, one for the graphs, and another for the tables.
    'The first table of the list object is the graph table, the second table of the listObject is the list of tables

    CheckSpecifications loTable
    If lData Is Nothing Then ThrowErrorUnableInstanciate "No linelist data provided"
    If OutputSh Is Nothing Then ThrowErrorUnableInstanciate "No OutputSheet provided"

    With New GraphSpecs
        'Lo list is the list of listobjects
        Set .LoList = loTable
        'Output sheet is where to write the graphs
        Set .OutputSheet = OutputSh
        .GraphType = GraphTypeComplex
        Set .LinelistSpecifications = lData
        Set CreateRangeSpecs = .Self()
    End With

End Function

Private Sub CheckSpecifications(ByVal loTable As BetterArray)
    Dim counter As Long
    Dim Lo As ListObject
    Dim testLo As Variant
    Dim graphLo As ListObject
    Dim tabLo As ListObject
    Dim titleLo As ListObject
    Dim rng As Range
    Dim titleVal As String

    'Test and check required specifications for listObjects for graph specs
    If loTable Is Nothing Then ThrowErrorUnableInstanciate "The List of listobjects is Empty"
    If loTable.Length <> 3 Then ThrowErrorUnableInstanciate "Seems like the table of listobjects is too long or too short? should be length 3"

    'Now test the graphListObject and the table listObject
    For counter = loTable.LowerBound To loTable.UpperBound
        Set testLo = loTable.Item(counter)
        If Not (TypeName(testLo) = "ListObject") Then ThrowErrorUnableInstanciate "Your table does not contains listobject"
        Set Lo = loTable.Item(counter)
        If Lo Is Nothing Then ThrowErrorUnableInstanciate "The listobject found in the table is not initialized"
        If Lo.DataBodyRange Is Nothing Then ThrowErrorUnableInstanciate "The listobject" & Lo.Name & "has no data to proceed with"
    Next

    'Now test for graphs and time series, and title of graphs
    Set graphLo = loTable.Item(loTable.LowerBound)
    'Test for graph title
    Set rng = graphLo.HeaderRowRange()
    titleVal = LCase(Application.WorksheetFunction.Trim(rng.Cells(-1, 1).Value))
    If titleVal <> TABONGRAPHTIMESERIES Then _
    ThrowErrorUnableInstanciate "The tables order is Graph Table and Then Time Series table, and then title table. No graph table found the title of first table is: " & titleVal

    Set tabLo = loTable.Item(loTable.LowerBound + 1)
    Set rng = tabLo.HeaderRowRange()
    titleVal = LCase(Application.WorksheetFunction.Trim(rng.Cells(-1, 1).Value))
    If titleVal <> TABONTIMESERIES Then _
    ThrowErrorUnableInstanciate "The tables order is Graph Table and Then Time Series table and then title table. No time series table found the title second table is: " & titleVal

    Set titleLo = loTable.Item(loTable.LowerBound + 2)
    Set rng = titleLo.HeaderRowRange()

    titleVal = LCase(Application.WorksheetFunction.Trim(rng.Cells(-1, 1).Value))
    If titleVal <> TABONTITLEGRAPHSTIMESERIES Then _
    ThrowErrorUnableInstanciate "The tables order is Graph Table and Then Time Series table and then title table. No title of graph table found the title of the third table is: " & titleVal

End Sub

Public Property Get OutputSheet() As Worksheet
    Set OutputSheet = this.outsh
End Property

Public Property Set OutputSheet(ByVal sh As Worksheet)
    Set this.outsh = sh
End Property

Public Property Get LinelistSpecifications() As ILinelistSpecs
    Set LinelistSpecifications = this.lData
End Property

Public Property Set LinelistSpecifications(ByVal lData As ILinelistSpecs)
    Set this.lData = lData
End Property

Public Property Get Table() As ICrossTable
    Set Table = this.tabl
End Property

Public Property Get GraphType() As GraphSpecsType
    GraphType = this.specstype
End Property

Public Property Let GraphType(ByVal spectype As GraphSpecsType)
    this.specstype = spectype
End Property

Public Property Set Table(ByVal tabl As ICrossTable)
    Set this.tabl = tabl
End Property

Public Property Get Self() As IGraphSpecs
    Set Self = Me
End Property

Public Property Get LoList() As BetterArray
    Set LoList = this.timelos
End Property

Public Property Set LoList(ByVal los As BetterArray)
    Set this.timelos = New BetterArray
    Set this.timelos = los.Clone()
End Property

Private Property Get GraphListObject() As ListObject
    Set GraphListObject = LoList.Item(LoList.LowerBound)
End Property

Private Property Get TimeSeriesListObject() As ListObject
    Set TimeSeriesListObject = LoList.Item(LoList.LowerBound + 1)
End Property

Private Property Get TitlesListObject() As ListObject
    Set TitlesListObject = LoList.Item(LoList.LowerBound + 2)
End Property

Public Property Get Wksh() As Worksheet          'Worksheet where to draw the graph
    Dim sh As Worksheet
    Select Case GraphType

    Case GraphTypeSimple
        Set sh = Table.Wksh()
    Case GraphTypeComplex
        Set sh = OutputSheet
    End Select

    Set Wksh = sh
End Property

Private Sub PushSeries(ByVal chrtName As String, ByVal chrtType As String, ByVal chrtPos As String)

    If this.chartSeriesName Is Nothing Then
        Set this.chartSeriesName = New BetterArray
        this.chartSeriesName.LowerBound = 1
    End If

    If this.chartSeriesType Is Nothing Then
        Set this.chartSeriesType = New BetterArray
        this.chartSeriesType.LowerBound = 1
    End If

    If this.chartSeriesPosition Is Nothing Then
        Set this.chartSeriesPosition = New BetterArray
        this.chartSeriesPosition.LowerBound = 1
    End If

    this.chartSeriesName.Push chrtName
    this.chartSeriesType.Push chrtType
    this.chartSeriesPosition.Push chrtPos
End Sub

Private Sub PushLabels(ByVal chrtRowCategory As String, ByVal chrtColumnLab As String, Optional ByVal chrtLabPref As String = vbNullString)

    If this.chartSeriesLabel Is Nothing Then
        Set this.chartSeriesLabel = New BetterArray
        this.chartSeriesLabel.LowerBound = 1
    End If

    If this.chartColumnCategoriesLabel Is Nothing Then
        Set this.chartColumnCategoriesLabel = New BetterArray
        this.chartColumnCategoriesLabel.LowerBound = 1
    End If

    'Chartseriles label prefix (prefix can be used in place of label names)
    If this.chartSeriesLabelPrefix Is Nothing Then
        Set this.chartSeriesLabelPrefix = New BetterArray
        this.chartSeriesLabelPrefix.LowerBound = 1
    End If

    this.chartSeriesLabel.Push chrtRowCategory
    this.chartColumnCategoriesLabel.Push chrtColumnLab
    this.chartSeriesLabelPrefix.Push chrtLabPref
End Sub

Private Sub PushGraph(ByVal graphId As String)

    'One graph consists of five betterarrays of same length:
    'One for series, one for types, one for position and the two others for the labels
    Dim graphSpecsTable As BetterArray
    Set graphSpecsTable = New BetterArray
    graphSpecsTable.LowerBound = 1

    'You need to define the five bette arrays because of (memory management issues)
    Dim tabSeries As BetterArray
    Dim tabType As BetterArray
    Dim tabPos As BetterArray
    Dim tabLabel As BetterArray
    Dim tabCatLabel As BetterArray
    Dim tabPrefLabel As BetterArray
    Dim title As String

    Set tabSeries = this.chartSeriesName.Clone()
    Set tabType = this.chartSeriesType.Clone()
    Set tabPos = this.chartSeriesPosition.Clone()
    Set tabLabel = this.chartSeriesLabel.Clone()
    Set tabCatLabel = this.chartColumnCategoriesLabel.Clone()
    Set tabPrefLabel = this.chartSeriesLabelPrefix.Clone()

    'Check the length of the tabSeries and remove the last element if needed
    'It looks like it is a bug from BetterArray memory management
    If tabSeries.Length = this.chartSeriesName.Length + 1 Then
        tabSeries.pop
        tabType.pop
        tabPos.pop
        tabLabel.pop
        tabCatLabel.pop
        tabPrefLabel.pop
    End If

    title = GraphTitle(graphId)

    graphSpecsTable.Push tabSeries, tabType, tabPos, _
                         tabLabel, tabCatLabel, tabPrefLabel, title

    'Higher Level, for one graph
    If this.graphsChartsSpecs Is Nothing Then
        Set this.graphsChartsSpecs = New BetterArray
        this.graphsChartsSpecs.LowerBound = 1
    End If

    'The list of graphs is a BetterArray of BetterArray, where each inner BetterArray Is a graph
    this.graphsChartsSpecs.Push graphSpecsTable
End Sub

Private Property Get SpecsLists() As BetterArray
    Set SpecsLists = this.graphsChartsSpecs.Clone()
End Property

Private Sub ClearSeries()
    'Avoid errors if the tables doesn't exists
    On Error Resume Next
    this.chartSeriesName.Clear
    this.chartSeriesType.Clear
    this.chartSeriesPosition.Clear
    this.chartSeriesLabel.Clear
    this.chartColumnCategoriesLabel.Clear
    this.chartSeriesLabelPrefix.Clear
    On Error GoTo 0
End Sub

Public Sub CreateSeries()
    Dim specs As ITablesSpecs
    Dim tabl As ICrossTable
    Dim tabId As String
    Dim mGraphIds As BetterArray
    Dim graphId As String
    Dim valueName As String
    Dim labelName As String
    Dim tabType As AnalysisTablesType
    Dim counter As Long
    Dim nbCols As Long
    Dim plotType As String
    Dim flipCoords As String

    Dim series As BetterArray
    Set series = New BetterArray

    Select Case GraphType

    Case GraphTypeSimple

        Set tabl = Table
        Set specs = tabl.Specifications
        tabId = specs.TableId
        tabType = specs.TableType
        nbCols = tabl.NumberOfColumns

        'flip coordinates
        flipCoords = specs.Value("flip")
        plotType = IIf(flipCoords = "yes", "hbar", "bar")

        Select Case tabType

        Case TypeUnivariate
            'First series, values for column 1
            'Plotype can plot either horizontal or vertical bars depending of the flip coordinates
            PushSeries "VALUES_COL_1_" & tabId, plotType, "left" 'bars at the left
            PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_1_" & tabId

            If (specs.HasPercentage And plotType = "bar") Then
                PushSeries "PERC_COL_1_" & tabId, "point", "right"
                PushLabels "ROW_CATEGORIES_" & tabId, "PERC_LABEL_COL_" & tabId
            End If

        Case TypeBivariate, TypeSpatial
            'Only values
            Select Case specs.Value("graph")

            Case "values"
                For counter = 1 To nbCols
                    PushSeries "VALUES_COL_" & counter & "_" & tabId, plotType, "left"
                    PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_" & counter & "_" & tabId
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
        End Select

    Case GraphTypeComplex
        'Time series graph

        'You define the graph  and add all the graphs to the list of graphs
        Set mGraphIds = New BetterArray
        Set mGraphIds = GraphIdsList()
        For counter = mGraphIds.LowerBound To mGraphIds.UpperBound
            graphId = mGraphIds.Item(counter)
            DefineGraphSpecs graphId
            PushGraph graphId                          'push the five tables to the nested BetterArray graphList
        Next
    End Select
End Sub

Public Property Get NumberOfSeries() As Long
    If (this.chartSeriesName Is Nothing) Or (GraphType = GraphTypeComplex) Then
        NumberOfSeries = 0
        Exit Property
    End If
    NumberOfSeries = this.chartSeriesName.Length()
End Property

Public Property Get NumberOfGraphs() As Long
    If (this.graphsChartsSpecs Is Nothing) Or (GraphType <> GraphTypeComplex) Then
        NumberOfGraphs = 0
        Exit Property
    End If
    NumberOfGraphs = this.graphsChartsSpecs.Length()
End Property

'Get the title of one complex graph from an index (a Long)

Private Function GraphTitle(ByVal graphId As String) As String
    'The first column of the listObject is for ids, the second column is for the title

    Dim title As String
    Dim titleLo As ListObject
    Dim idRng As Range
    Dim cellRng As Range

    'List of all ids for one graph
    Set titleLo = TitlesListObject()

    'idRng is the first column of the listobject where there are ids for graphs
    Set idRng = titleLo.ListColumns(3).DataBodyRange

    If Not idRng.Find(What:=graphId, LookAt:=xlWhole, MatchCase:=True) Is Nothing Then
        Set cellRng = idRng.Find(What:=graphId, LookAt:=xlWhole, MatchCase:=True)
        title = cellRng.Offset(, -2).Value
    End If

    'The title can be empty, if unable to find the graphId in the title table
    GraphTitle = title
End Function



Public Property Get SeriesName(ByVal Index As Long) As String
    If (Index > NumberOfSeries) Or (Index < 0) Then ThrowError "Index of series out of bound"
    SeriesName = this.chartSeriesName.Item(Index)
End Property

Public Property Get SeriesType(ByVal Index As Long) As String
    If (Index > NumberOfSeries) Or (Index < 0) Then ThrowError "Index of series types out of bound"
    SeriesType = this.chartSeriesType.Item(Index)
End Property

Public Property Get SeriesPos(ByVal Index As Long) As String
    If (Index > NumberOfSeries) Or (Index < 0) Then ThrowError "Index of series positions out of bound"
    SeriesPos = this.chartSeriesPosition.Item(Index)
End Property

Public Property Get SeriesLabel(ByVal Index As Long) As String
    If (Index > NumberOfSeries) Or (Index < 0) Then ThrowError "Index of series positions out of bound"
    SeriesLabel = this.chartSeriesLabel.Item(Index)
End Property

Public Property Get SeriesColumnLabel(ByVal Index As Long) As String
    If (Index > NumberOfSeries) Or (Index < 0) Then ThrowError "Index of series positions out of bound"
    SeriesColumnLabel = this.chartColumnCategoriesLabel.Item(Index)
End Property

Private Property Get Valid() As Boolean

    If this.testDone Then
        Valid = this.validity
        Exit Property
    End If

    If GraphType = GraphTypeSimple Then
        CreateSeries
        Valid = True
    Else
        CreateSeries
        this.validity = Not (this.graphsChartsSpecs Is Nothing)
        this.testDone = True
    End If

    Valid = this.validity
End Property

Private Sub DefineGraphSpecs(ByVal graphId As String)
    'Defines Specs for one grpah Id of type complex
    'Complex graphs are graphs for time series (They are tables of tables of graphs)

    Dim tsLo As ListObject
    Dim counter As Long
    Dim lData As ILinelistSpecs
    Dim specs As ITablesSpecs
    Dim val As String
    Dim tabId As String
    Dim secId As String
    Dim colName As String
    Dim labName As String
    Dim posList As BetterArray
    Dim percList As BetterArray
    Dim typeList As BetterArray
    Dim choiceList As BetterArray
    Dim labelList As BetterArray
    Dim idsList As BetterArray
    Dim choiValue As String
    Dim hRng As Range
    Dim tabRng As Range


    Set posList = New BetterArray
    Set percList = New BetterArray
    Set typeList = New BetterArray
    Set choiceList = New BetterArray
    Set idsList = New BetterArray
    Set labelList = New BetterArray


    Set lData = LinelistSpecifications

    'I can't run definegraph specs on simple graphs.
    If GraphType <> GraphTypeComplex Then ThrowError "Multiple table Ids are only on time series graph"


    'I can suppose that all the graphIDs are sorted in the order they appear
    Set tsLo = TimeSeriesListObject()
    Set hRng = tsLo.HeaderRowRange
    Set idsList = GraphValues(colName:="series id", graphId:=graphId)
    Set posList = GraphValues(colName:="axis", graphId:=graphId)
    Set percList = GraphValues(colName:="percentages", graphId:=graphId)
    Set typeList = GraphValues(colName:="type", graphId:=graphId)
    Set choiceList = GraphValues(colName:="choices", graphId:=graphId)
    Set labelList = GraphValues(colName:="label", graphId:=graphId)

    'Clear al the actual series
    ClearSeries

    'Convert all times to table id
    For counter = idsList.LowerBound To idsList.UpperBound

        val = idsList.Item(counter)
        val = Application.WorksheetFunction.Trim(Replace(LCase(val), "series", vbNullString))

        Set tabRng = tsLo.ListRows(CLng(val)).Range
        Set specs = TablesSpecs.Create(hRng, tabRng, lData)
        tabId = specs.TableId
        secId = specs.TableSectionId
        choiValue = choiceList.Item(counter)

        labName = TimeSeriesColumnName(choiValue, tabId)
        colName = Switch(choiValue = "Total", "TOTAL_COL_VALUES_" & tabId, _
                        InStr(1, labName, "COLUMN_CATEGORIES_") > 0, "INTERIOR_VALUES_" & tabId, _
                        True, Replace(labName, "LABEL", "VALUES"))

        If (percList.Item(counter) = "percentages") Then
            colName = IIf(choiValue = "Total", Replace(labName, "COL", "PERC"), Replace(labName, "LABEL", "PERC"))
        End If

        'I could not find the TimeSeriesColumn, in that case we jump to the next iteration
        If colName <> vbNullString Then
            PushSeries colName, typeList.Item(counter), posList.Item(counter)
            PushLabels "ROW_CATEGORIES_" & secId, labName, labelList.Item(counter)
        End If

    Next

End Sub

'Get the column Name of the choice in the table
Private Function TimeSeriesColumnName(ByVal choiValue As String, ByVal tabId As String) As String
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

    'If Not ValueExists(rng, choiValue, True) Then ThrowError choiValue & " is not in the drawn table " & tabId
    Set cellRng = rng.Find(What:=choiValue, LookAt:=xlWhole, MatchCase:=True)

    On Error Resume Next
    If Not cellRng Is Nothing Then colName = cellRng.Name.Name
    On Error GoTo 0

    'If colName = vbNullString Then ThrowError choiValue & "Has no name in the table " & tabId
    TimeSeriesColumnName = colName
End Function

'Test if a value Exists in a Range (pay attention the .Find is a tricky function, be sure your range does not contains duplicates
'or starts after cell(1, 1))
Private Function ValueExists(ByVal rng As Range, ByVal Value As String, Optional ByVal strictSearch As Boolean = False)
    If Not strictSearch Then
        ValueExists = (Not rng.Find(What:=Value, LookAt:=xlPart, MatchCase:=False) Is Nothing)
    Else
        ValueExists = (Not rng.Find(What:=Value, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
    End If

End Function

'Column Index of a colname in a listObject
Private Function ColumnIndex(ByVal Lo As ListObject, ByVal colName As String, Optional ByVal inRange As Boolean = True) As Long
    Dim Index As Long
    Dim hRng As Range

    Set hRng = Lo.HeaderRowRange
    If ValueExists(hRng, colName) Then
        If inRange Then
            Index = hRng.Find(What:=colName, LookAt:=xlPart, MatchCase:=False).Column - hRng.Column + 1
        Else
            Index = hRng.Find(What:=colName, LookAt:=xlPart, MatchCase:=False).Column
        End If
    Else
        Index = -1
    End If

    ColumnIndex = Index
End Function

Private Function GraphIdsList() As BetterArray
    Dim idsList As BetterArray
    Dim graphLo As ListObject
    Dim idIndex As Long
    Dim rng As Range
    Dim counter As Long

    Set idsList = New BetterArray
    idsList.LowerBound = 1
    Set graphLo = GraphListObject()
    idIndex = ColumnIndex(graphLo, "graph id")
    Set rng = graphLo.ListColumns(idIndex).Range

    counter = 2
    Do While counter <= rng.Rows.Count

        If Not idsList.Includes(rng.Cells(counter, 1).Value) And Not IsEmpty(rng.Cells(counter, 1)) Then idsList.Push rng.Cells(counter, 1).Value
        counter = counter + 1

    Loop
    Set GraphIdsList = idsList.Clone()
End Function

Public Property Get GraphIDs() As BetterArray
    If this.GraphIdsList Is Nothing Then
        Set this.GraphIdsList = New BetterArray
        Set this.GraphIdsList = GraphIdsList()
        Set GraphIDs = this.GraphIdsList.Clone()
    Else
        Set GraphIDs = this.GraphIdsList.Clone()
    End If
End Property

Private Function GraphValues(ByVal colName As String, ByVal graphId As String) As BetterArray

    'Returns all the values for one column of a given graph (a kind of filter)

    Dim idIndex As Long
    Dim RowIndex As Long
    Dim colIndex As Long
    Dim grLo As ListObject
    Dim rng As Range
    Dim colRng As Range
    Dim cellVal As String
    Dim Values As BetterArray
    Dim sh As Worksheet


    Set grLo = GraphListObject()
    Set Values = New BetterArray
    Values.LowerBound = 1

    idIndex = ColumnIndex(grLo, "graph id")
    colIndex = ColumnIndex(grLo, colName)

    If idIndex < 0 Then ThrowError "graph id column not found"
    If colIndex < 0 Then ThrowError colName & " is not present in graph colnames"

    Set rng = grLo.ListColumns(idIndex).DataBodyRange
    Set colRng = grLo.ListColumns(colIndex).DataBodyRange

    Set sh = rng.Worksheet

    If ValueExists(rng, graphId, strictSearch:=True) Then
        With sh
            Do While Not (rng.Find(What:=graphId, LookAt:=xlWhole, MatchCase:=True, After:=rng.Cells(rng.Rows.Count, 1)) Is Nothing)
                RowIndex = rng.Find(What:=graphId, LookAt:=xlWhole, MatchCase:=True, After:=rng.Cells(rng.Rows.Count, 1)).Row - rng.Row + 1

                cellVal = colRng.Cells(RowIndex, 1).Value

                If cellVal = vbNullString Then

                    Values.Push "&"

                Else

                    Values.Push cellVal

                End If

                'Important to exit if the range has only one cell since reducing could
                'return the same range
                If rng.Rows.Count = 1 Then Exit Do
                If RowIndex + 1 > rng.Rows.Count Then Exit Do

                'Reduce the ranges for further searches
                Set rng = .Range(rng.Cells(RowIndex + 1, 1), rng.Cells(rng.Rows.Count, 1))
                Set colRng = .Range(colRng.Cells(RowIndex + 1, 1), colRng.Cells(colRng.Rows.Count, 1))
            Loop
        End With
    End If

    Set GraphValues = Values.Clone()
End Function

'Errors
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Execution Error : " & errorMessage
End Sub

'Exposed Properties and methods
Public Property Get IGraphSpecs_NumberOfSeries() As Long

    IGraphSpecs_NumberOfSeries = NumberOfSeries()
End Property

Public Property Get IGraphSpecs_NumberOfGraphs() As Long
    IGraphSpecs_NumberOfGraphs = NumberOfGraphs()
End Property

Public Property Get IGraphSpecs_SeriesName(ByVal Index As Long) As String
    IGraphSpecs_SeriesName = SeriesName(Index)
End Property

Public Property Get IGraphSpecs_SeriesType(ByVal Index As Long) As String
    IGraphSpecs_SeriesType = SeriesType(Index)
End Property

Public Property Get IGraphSpecs_SeriesPos(ByVal Index As Long) As String
    IGraphSpecs_SeriesPos = SeriesPos(Index)
End Property

Public Property Get IGraphSpecs_SeriesLabel(ByVal Index As Long) As String
    IGraphSpecs_SeriesLabel = SeriesLabel(Index)
End Property

Public Property Get IGraphSpecs_SeriesColumnLabel(ByVal Index As Long) As String
    IGraphSpecs_SeriesColumnLabel = SeriesColumnLabel(Index)
End Property

Public Sub IGraphSpecs_CreateSeries()
    CreateSeries
End Sub

Public Property Get IGraphSpecs_SpecsLists() As BetterArray
    Set IGraphSpecs_SpecsLists = SpecsLists()
End Property

Public Property Get IGraphSpecs_Valid() As Boolean
    IGraphSpecs_Valid = Valid()
End Property
