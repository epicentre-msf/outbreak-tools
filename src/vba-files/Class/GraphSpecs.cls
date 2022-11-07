VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("")
'@IgnoreModule

Option Explicit
Implements IGraphSpecs

Public Enum GraphSpecsType
    GraphTypeSimple = 1
    GraphTypeComplex = 2
End Enum

'Will have two instanciation methods (one like a tablespecs, another which
'norma and reserved to most cases (bivariate, spatial))
Private Type TGraphSpecs
    tabl As ICrossTable
    timelos As BetterArray
    specstype As GraphSpecsType
    outsh As Worksheet
    chartSeriesName As BetterArray
    chartSeriesType As BetterArray
    chartSeriesPosition As BetterArray
    chartSeriesLabel As BetterArray
    chartColumnCategoriesLabel As BetterArray
    lData As ILinelistSpecs
    graphsChartsSpecs As BetterArray             'A BetterArray of BetterArray all with length 55 of same length, each one for series, Type, label and category label
    GraphIdsList As BetterArray
End Type

Private this As TGraphSpecs
Private Const CLASSNAME As String = "GraphSpecs"

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

Public Function CreateRangeSpecs(ByVal loTable As BetterArray, ByVal outputsh As Worksheet, ByVal lData As ILinelistSpecs) As IGraphSpecs
    'loTable is a table of two listObjects, one for the graphs, and another for the tables.
    CheckSpecifications loTable
    If lData Is Nothing Then ThrowErrorUnableInstanciate "No linelist data provided"
    If outputsh Is Nothing Then ThrowErrorUnableInstanciate "No OutputSheet provided"

    With New GraphSpecs
        Set .LoList = loTable
        Set .OutputSheet = outputsh
        .GraphType = GraphTypeComplex
        .LinelistSpecifications = lData
        Set Create = .Self
    End With
End Function

Private Sub CheckSpecifications(ByVal loTable As BetterArray)
    Dim counter As Long
    Dim Lo As ListObject
    Dim testLo As Variant
    Dim graphLo As ListObject
    Dim tabLo As ListObject
    Dim rng As Range

    'Test and check required specifications for listObjects for graph specs
    If loTable Is Nothing Then ThrowErrorUnableInstanciate "The List of listobjects is Empty"
    If Length(loTable) <> 2 Then ThrowErrorUnableInstanciate "Seems like the table of listobjects is too long or too short? should be length 2"

    'Now test the graphListObject and the table listObject
    For counter = loTable.LowerBound To loTable.UpperBound
        Set testLo = loTable.Item(counter)
        If Not (TypeName(testLo) = "ListObject") Then ThrowErrorUnableInstanciate "Your table does not contains listobject"
        Set Lo = loTable.Item(counter)
        If Lo Is Nothing Then ThrowErrorUnableInstanciate "The listobject found in the table is not initialized"
        If Lo.DataBodyRange Is Nothing Then ThrowErrorUnableInstanciate "The listobject" & Lo.Name & "has no data to proceed with"
    Next

    'Now test for graphs and time series
    Set graphLo = loTable.Item(loTable.LowerBound)
    'Test for graph title
    Set rng = graphLo.HeaderRowRange()
    titleVal = LCase(Application.WorksheetFunction.Trim(rng.Cells(-1, 1).Value))
    If titleVal <> "graph on time series" Then ThrowErrorUnableInstanciate "The tables order is Graph Table and Then Time Series table. No graph table found the title is: " & titleVal
  
    Set tabLo = loTable.Item(loTable.LowerBound + 1)
    Set rng = graphLo.HeaderRowRange()
    titleVal = LCase(Application.WorksheetFunction.Trim(rng.Cells(-1, 1).Value))
    If titleVal <> "time series analysis" Then ThrowErrorUnableInstanciate "The tables order is Graph Table and Then Time Series table. No time series table found the title is: " & titleVal
End Sub

Public Property Get OutputSheet() As Worksheet
    Set OutputSheet = this.outsh
End Property

Public Property Set OutputShet(ByVal sh As Worksheet)
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
    this.timelos.LowerBound = 1
End Property

Private Property Get GraphListObject() As ListObject
    Set GraphListObject = LoList.Item(1)
End Property

Private Property Get TimeSeriesListObject() As ListObject
    Set TimeSeriesListObject = LoList.Item(2)
End Property

Public Property Get Wksh() As Worksheet          'Worksheet where to draw the graph
    Dim sh As Worksheet
    Select Case GraphType

    Case GraphTypeSimple
        Set sh = Table.Wksh()
    Case GrahTypeComplex
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

Private Sub PushLabels(ByVal chrtRowCategory As String, ByVal chrtColumnLab As String)

    If this.chartSeriesLabel Is Nothing Then
        Set this.chartSeriesLabel = New BetterArray
        this.chartSeriesLabel.LowerBound = 1
    End If

    If this.chartColumnCategoriesLabel Is Nothing Then
        Set this.chartColumnCategoriesLabel = New BetterArray
        this.chartColumnCategoriesLabel.LowerBound = 1
    End If

    this.chartSeriesLabel.Push chrtRowCategory
    this.chartColumnCategoriesLabel.Push chrtColumnLab
End Sub

Private Sub PushGraph()
    Dim graphSpecsTable As BetterArray
    Set graphSpecsTable = New BetterArray
    graphSpecsTable.LowerBound = 1

    graphSpecsTable.Push this.chartSeriesName, this.chartSeriesType, this.chartSeriesPosition, _
                         this.chartSeriesLabel, this.chartColumnCategoriesLabel
  
    'Higher Level, for one graph
    If this.graphsChartsSpecs Is Nothing Then
        Set this.graphsChartsSpecs = New BetterArray
        this.graphsChatsSpecs.LowerBound = 1
    End If

    this.graphsChartsSpecs.Push graphSpecsTable
End Sub

Private Sub ClearSeries()
    'Avoid errors if the tables don't exists
    On Error Resume Next
    this.chartSeriesName.Clear
    this.chartSeriesType.Clear
    this.chartSeriesPosition.Clear
    this.chartSeriesLabel.Clear
    this.chartColumnCategoriesLabel.Clear
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

    Dim series As BetterArray
    Set series = New BetterArray
  
    Select Case GraphType
    
    Case GraphTypeSimple
        Set tabl = Table
        Set specs = tabl.Specifications
        tabId = specs.TableId
        tabType = specs.TableType
        nbCols = tabl.NumberOfColumns

        Select Case tabType

        Case TypeUnivariate
            'First series, values for column 1
            PushSeries "VALUES_COL_1_" & tabId, "bar", "left" 'bars at the left
            PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_1_" & tabId
            If specs.HasPercentage Then
                PushSeries "PERC_COL_1_" & tabId, "point", "right"
                PushLabels "ROW_CATEGORIES_" & tabId, "PERC_LABEL_COL_" & tabId
            End If

        Case TypeBivariate, TypeSpatial
            'Only values
            Select Case specs.Value("graph")
      
            Case "values"
                For counter = 1 To nbCols
                    PushSeries "VALUES_COL_" & counter & "_" & tabId, "bar", "left"
                    PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_" & counter & "_" & tabId
                Next
      
            Case "percentage"
                If specs.HasPercentage Then
                    For counter = 1 To nbCols
                        PushSeries "PERC_COL_" & counter & "_" & tabId, "bar", "left"
                        PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_" & counter & "_" & tabId
                    Next
                End If

            Case "both"
                For counter = 1 To nbCols
                    PushSeries "VALUES_COL_" & counter & "_" & tabId, "bar", "left"
                    PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_" & counter & "_" & tabId
                    If specs.HasPercentage Then
                        PushSeries "PERC_COL_" & counter & "_" & tabId, "point", "right"
                        PushLabels "ROW_CATEGORIES_" & tabId, "LABEL_COL_" & counter & "_" & tabId
                    End If
                Next
            End Select
        End Select

    Case GraphTypeComplex
        'Time series
        Set mGraphIds = New BetterArray
        Set mGraphIds = GraphIdsList()
        For counter = mGraphIds.LowerBound To mGraphIds.UpperBound
            ClearSeries
            graphId = mGraphIds.Item(counter)
            DefineGraphSpecs graphId
            PushGraph                            'push the five tables to the nested BetterArray graphList
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

Public Property Get SeriesName(ByVal index As Long) As String
    If (index > NumberOfSeries) Or (index < 0) Then ThrowError "Index of series out of bound"
    SeriesName = this.chartSeriesName.Item(index)
End Property

Public Property Get SeriesType(ByVal index As Long) As String
    If (index > NumberOfSeries) Or (index < 0) Then ThrowError "Index of series types out of bound"
    SeriesType = this.chartSeriesType.Item(index)
End Property

Public Property Get SeriesPos(ByVal index As Long) As String
    If (index > NumberOfSeries) Or (index < 0) Then ThrowError "Index of series positions out of bound"
    SeriesPos = this.chartSeriesPosition.Item(index)
End Property

Public Property Get SeriesLabel(ByVal index As Long) As String
    If (index > NumberOfSeries) Or (index < 0) Then ThrowError "Index of series positions out of bound"
    SeriesLabel = this.chartSeriesLabel.Item(index)
End Property

Public Property Get SeriesColumnLabel(ByVal index As Long) As String
    If (index > NumberOfSeries) Or (index < 0) Then ThrowError "Index of series positions out of bound"
    SeriesColumnLabel = this.chartColumnCategoriesLabel.Item(index)
End Property

Private Sub DefineGraphSpecs(ByVal graphId As String)
    'Defines Specs for one grpah Id

    Dim tsLo As ListObject
    Dim counter As Long
    Dim lData As ILinelistSpecs
    Dim specs As ITablesSpecs
    Dim Val As String
    Dim tabId As String
    Dim secId As String
    Dim colName As String
    Dim labName As String
    Dim posList As BetterArray
    Dim percList As BetterArray
    Dim typeList As BetterArray
    Dim choiList As BetterArray
    Dim idsList As BetterArray
    Dim choiValue As String


    Set posList = New BetterArray
    Set perList = New BetterArray
    Set typeList = New BetterArray
    Set choiList = New BetterArray
    Set idsList = New BetterArray


    Set lData = LinelistSpecifications
    If GraphType <> GraphTypeComplex Then ThrowError "Multiple table Ids are only on time series graph"


    Set tsLo = TimeSeriesListObject()
    Set hRng = tsLo.HeaderRowRange
    Set idsList = GraphValues(colName:="series id", graphId:=graphId)
    Set posList = GraphValues(colName:="axis", graphId:=graphId)
    Set percList = GraphValues(colName:="percentages", graphId:=graphId)
    Set typeList = GraphValues(colName:="type", graphId:=graphId)
    Set choiceList = GraphValues(colName:="choices", graphId:=graphId)

    'Convert all itmes to table id
    For counter = idsList.LowerBound To idsList.UpperBound
        
        Val = idsList.Item(counter)
        Val = Application.WorksheetFunction.Trim(Replace(LCase(Val), "series", vbNullString))

        Set tabRng = tsLo.ListRows(CLng(Val))
        Set specs = TablesSpecs.Create(hRng, tabRng, lData)
        tabId = specs.TableId
        secId = specs.TableSectionId
        choiValue = choiceList.Item(counter)
    
        labName = TimeSeriesColumnName(choiValue, tabId)
        colName = Replace(labName, "LABEL", "VALUES")

        If (percList.Item(counter) = "percentages") Then
            colName = IIf(choiValue = "Total", Replace(labName, "COL", "PERC"), Replace(labName, "LABEL", "PERC"))
        End If


        PushSeries colName, typeList.Item(counter), posList.Item(counter)
        PushLabels "ROW_CATEGORIES_" & secId, labName
        
    Next

End Sub

'Get the column Name of the choice in the table
Private Function TimeSeriesColumnName(ByVal choiValue As String, ByVal tabId As String) As String
    Dim rng As Range
    Dim cellRng As Range
    Dim colName As String
    Dim sh As Worksheet

    Set sh = Wksh()

    If choiValue = "Total" Then
        TimeSeriesColumnName = "TOTAL_COL_VALUES_" & tabId
        Exit Function
    End If
  

    Set rng = sh.Range("COLUMN_CATEGORIES_" & tabId)
    'If Not ValueExists(rng, choiValue, True) Then ThrowError choiValue & " is not in the drawn table " & tabId
    Set cellRng = rng.Find(What:=choiValue, LookAt:=xlWhole, MatchCase:=True)

    On Error Resume Next
    colName = cellRng.Name.Name
    On Error GoTo 0
  
    'If colName = vbNullString Then ThrowError choiValue & "Has no name in the table " & tabId
    TimeSeriesColumnName = colName
End Function

'Test if a value Exists in a Range
Private Function ValueExists(ByVal rng As Range, ByVal Value As String, Optional ByVal strictSearch As Boolean = False)
    If Not strictSearch Then
        ValueExists = (Not rng.Find(What:=Value, LookAt:=xlPart, MatchCase:=Fase) Is Nothing)
    Else
        ValueExists = (Not rng.Find(What:=Value, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
    End If

End Function

'Column Index
Private Function ColumnIndex(ByVal Lo As ListObject, ByVal colName As String, Optional ByVal inRange As Boolean = True) As Long
    Dim index As Long
    Dim hRng As Range

    Set hRng = Lo.HeaderRowRange
    If ValueExists(hRng, colName) Then
        If inRange Then
            index = hRng.Find(What:=colName, LookAt:=xlPart, MatchCase:=Fase).Column - hRng.Column + 1
        Else
            index = hRng.Find(What:=colName, LookAt:=xlPart, MatchCase:=Fase).Column
        End If
    Else
        index = -1
    End If

    ColumnIndex = index
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

    Do While counter <= rng.Rows.Count
        If Not idsList.Includes(rng.Cells(counter, 1).Value) Then
            idsList.Push rng.Cells(counter, 1).Value
        End If
    Loop
    GraphIdsList = idsList.Clone()
End Function

Public Property Get GraphIDs() As BetterArray
    If this.GraphIdsList Is Nothing Then
        Set this.GraphIdsList = New BetterArray
        Set this.GraphIdsList = GraphIldsList()
        Set GraphIDs = this.GraphIdsList.Clone()
    Else
        Set GraphIDs = this.GraphIdsList.Clone()
    End If
End Property

Private Function GraphValues(ByVal colName As String, ByVal graphId As String) As BetterArray
    Dim idIndex As String
    Dim RowIndex As Long
    Dim colIndex As Long
    Dim grLo As ListObject
    Dim rng As Range
    Dim colRng As Range
    Dim Values As BetterArray
    Dim sh As Worksheet

  
    Set grLo = GraphListObject()
    Set Values = New BetterArray
    Values.LowerBound = 1

    idIndex = ColumnIndex(grLo, "graph id")
    colIndex = ColumnIndex(grLo, colName)

    If idIndex < 0 Then ThrowError "graph id column not found"
    If colIndex < 0 Then ThrowError colName & " is not present in graph colnames"

    Set rng = grLo.ListColumns(idIndex)
    Set colRng = grLo.ListColumns(colIndex)
    Set sh = rng.Worksheet

    If ValueExists(rng, graphId, strictSearch:=True) Then
        With sh
            Do While Not (rng.Find(What:=graphId, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
                RowIndex = rng.Find(What:=graphId, LookAt:=xlWhole, MatchCase:=True).row
                Values.Push colRng.Cells(RowIndex, 1).Value
                Set rng = .Range(rng.Cells(RowIndex + 1, 1), rng.Cells(rng.Rows.Count, 1))
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

Public Property Get IGraphSpecs_SeriesName(ByVal index As Long) As String
    IGraphSpecs_SeriesName = SeriesName(index)
End Property

Public Property Get IGraphSpecs_SeriesType(ByVal index As Long) As String
    IGraphSpecs_SeriesType = SeriesType(index)
End Property

Public Property Get IGraphSpecs_SeriesPos(ByVal index As Long) As String
    IGraphSpecs_SeriesPos = SeriesPos(index)
End Property

Public Property Get IGraphSpecs_SeriesLabel(ByVal index As Long) As String
    IGraphSpecs_SeriesLabel = SeriesLabel(index)
End Property

Public Property Get IGraphSpecs_SeriesColumnLabel(ByVal index As Long) As String
    IGraphSpecs_SeriesColumnLabel = SeriesColumnLabel(index)
End Property

Public Sub IGraphSpecs_CreateSeries()
    CreateSeries
End Sub


