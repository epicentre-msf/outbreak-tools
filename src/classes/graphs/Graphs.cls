VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Graphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manage Excel graphs"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Manage Excel graphs")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  UseMeaningfulName, HungarianNotation, ExcelMemberMayReturnNothing

Option Explicit
Implements IGraphs

Private Type TGraph
    sh As Worksheet
    strtRng As Range
    co As ChartObject
    Index As Long
    grName As String
    hasChecks As Boolean
End Type

Private this As TGraph

'Checking objects
Private checkCounter As Long
Private internalChecks As IChecking

'@section Instantiation
'===============================================================================

'@label:Create
'@prop-title Create a Graph object
'@details
'Graph objects simplify the manipulation of Excel graphs. They are based on a basic
'graphic grammar with the ability to stack graphs on top of each other.
'At its core, a graph object is created with a worksheet and a range where
'to write the graph. Values to be plotted should be in a named range on the
'worksheet used to instantiate the class.
'
'@param sh A Worksheet that contains the values to be plotted.
'@param posRng A Range that determines where to position the graph.
'@param grName A String for the graph name.
'@depends Checking
'@export
Public Function Create(sh As Worksheet, posRng As Range, Optional ByVal grName As String = vbNullString) As IGraphs
    
    If (sh Is Nothing) Then ThrowError ProjectError.ObjectNotInitialized, "Worksheet not initialized, unable to create graph"
    If (posRng Is Nothing) Then ThrowError ProjectError.ObjectNotInitialized, "Range not initialized, unable to create graph"
    
    With New Graphs
        Set .Wksh = sh
        Set .StartRange = posRng
        .Name = grName
        Set Create = .Self
    End With
End Function

'@prop-title Current object instance
'@return An IGraphs instance (Self)
Public Property Get Self() As IGraphs
    Set Self = Me
End Property

'@prop-title Host worksheet for the graph
'@return A Worksheet
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@prop-title Set the host worksheet
'@param sh A Worksheet where graph data resides
Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

'@prop-title Starting cell/range for the graph
'@return A Range
Public Property Get StartRange() As Range
    Set StartRange = this.strtRng
End Property

'@prop-title Set the starting cell/range
'@param strtRng A Range used to position the graph
Public Property Set StartRange(ByVal strtRng As Range)
    Set this.strtRng = strtRng
End Property


'@prop-title Set the graph name
'@param grName A String, the graph name
Public Property Let Name(ByVal grName As String)
    this.grName = grName
End Property

'@prop-title Get the graph name
'@return A String
Public Property Get Name() As String
    Name = this.grName
End Property

'@label:elements
'@section Graph Elements
'===============================================================================

'@prop-title Graph chart object
'@return A ChartObject
Private Property Get graphChart() As ChartObject
    Set graphChart = this.co
End Property

'@prop-title Latest series index
'@return A Long
Private Property Get seriesIndex() As Long
    seriesIndex = this.Index
End Property

'@prop-title Update latest series index
'@param ind A Long, new index value
Private Property Let seriesIndex(ByVal ind As Long)
    this.Index = ind
End Property

'@label:operate
'@section Operate with the outside world
'===============================================================================

'@label:add
'@sub-title Add a graph to the worksheet
'@details
'Creates an empty graph on a worksheet, ready for data to be
'attached to the graph. The graph is empty at creation. The width and height
'of the graph object may change depending on the scope in the future.
'@export
Private Sub Add()
    'Upon creation, the width and height of graphs are constant
    
    Const CW As Long = 488 'graph width
    Const RH As Long = 214  'graph height
    
    Dim co As ChartObject
    Dim sh As Worksheet
    Dim rng As Range

    Set rng = StartRange()
    Set sh = rng.Parent

    Set co = sh.ChartObjects.Add(rng.Left, rng.Top, RH, CW)
    
    co.Left = rng.Left
    co.Top = rng.Top

    With co.Chart
        .PlotArea.Interior.color = RGB(235, 235, 235)
        .Legend.Position = xlLegendPositionTop
    End With

    Set this.co = co
End Sub

'@label:addseries
'@sub-title Add series to the graph
'@details
'One graph can have multiple series added. Series are referenced using
'named ranges on the graph object worksheet.
'@param rngName A String, the named range for the values
'@param chrtType A String, chart type to add. Can be "line", "point", "bar", or "hbar". Defaults to "bar".
'@param axisPos A String, axis position of the series. Can be "right" or "left". If "right", a secondary axis is added to the graph.
'@export
Private Sub AddSeries(ByVal rngName As String, ByVal chrtType As String, _
                    Optional axisPos As String = vbNullString)

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim ind As Long
    Dim customType As XlChartType
    Dim axGrp As XlAxisGroup
    
    'The index is initialized at 0 by default
    Set co = graphChart()
    Set sh = Wksh()
    
    'If the provided range does not exist, exit the sub
    If Not RangeExist(rngName) Then
        LogInfo "Attempt to add series " & rngName & " which does not exists in worksheet " & sh.Name, checkingWarning
        Exit Sub
    End If
    
    'If no graph exists, add one first
    If co Is Nothing Then
        Add
        Set co = graphChart()
    End If

    ind = seriesIndex() + 1

    Select Case chrtType
        Case "line": customType = xlLineMarkers
        Case "point": customType = xlXYScatter
        Case "bar": customType = xlColumnClustered
        Case "hbar": customType = xlBarClustered
        Case Else: customType = xlColumnClustered
    End Select

    axGrp = IIf(axisPos = "right", xlSecondary, xlPrimary)

    With co.Chart
        .SeriesCollection.Add _
        Source:=sh.Range(rngName), RowCol:=xlColumns, _
        CategoryLabels:=False, Replace:=False
        .SeriesCollection(ind).Values = sh.Range(rngName)

        'Add a right axis if required
        If axGrp = xlSecondary Then
            .HasAxis(xlValue, xlSecondary) = True
            .SeriesCollection(ind).AxisGroup = axGrp
        End If

        .SeriesCollection(ind).chartType = customType
    End With

    seriesIndex = ind
End Sub


'@label:addlabels
'@sub-title Add text labels to a graph
'@param catName String, legend entry for the series values (axis labels).
'@param lblName String, legend entry for the series (often the data header).
'@param prefix Optional String. Adds a prefix to value labels.
'@param prefixOnly Optional Boolean. Adds only the prefix to the legend entry. Defaults to False.
'@param hardCodeLabels Optional Boolean. 
Private Sub AddLabels(ByVal catName As String, _
                     ByVal lblName As String, _
                     Optional ByVal prefix As String = vbNullString, _
                     Optional ByVal prefixOnly As Boolean = False, _
                     Optional ByVal hardCodeLabels As Boolean = True)

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim labelValue As String
    Dim ind As Long
    Dim currPrefix As String

    Set co = graphChart
    Set sh = Wksh()

    'Labels are added to the current index
    ind = seriesIndex()

    If ind <= 0 Then
        LogInfo "Attempt to add labels, but there a no series on this graph", checkingWarning
        Exit Sub
    End If

    'Add a prefix to the label name (useful for time series graphs)
    currPrefix = prefix
    If currPrefix <> vbNullString And (Not prefixOnly) Then currPrefix = currPrefix & " - "

    If prefixOnly Then
        labelValue = currPrefix
    ElseIf RangeExist(lblName) Then
        labelValue = currPrefix & CStr(sh.Range(lblName).Value)
    Else
        labelValue = currPrefix
    End If

    'Throw an error when co is Nothing
    With co.Chart
        If RangeExist(catName) Then .SeriesCollection(ind).XValues = sh.Range(catName)
        'Labels can be hardcoded or dynamic; by default, they are hardcoded
        If (Not hardCodeLabels) And RangeExist(lblName) Then
            'The user can modify by changing values in the range, but
            'You have to assume the range exists and is not empty
            .SeriesCollection(ind).Name = "= " & RngAddress(lblName)
        Else
            .SeriesCollection(ind).Name = labelValue
        End If

        .SeriesCollection(ind).ApplyDataLabels xlDataLabelsShowValue, False
        .SeriesCollection(ind).DataLabels.Font.Size = 6
        .SeriesCollection(ind).DataLabels.Font.color = RGB(2, 6, 91)
        .SeriesCollection(ind).DataLabels.Font.Name = "+mj-lt"
    End With
End Sub

'@label:format
'@sub-title Format the current graph
'@details
' Formats are defined based on the scope of the graph. For time series, graphs are larger;
' for spatio-temporal analyses, they are wider.
'@param valuesTitle A String, title of the values axis (usually the Y-axis)
'@param catTitle A String, title of the categories axis (usually the X-axis)
'@param plotTitle A String, title of the plot
'@param scope A GraphScope value indicating the scope of the graph
'@param heightFactor A Long, multiplicative factor applied to the default height and width
'@param hardCodeLabels Boolean, whether to use literal values for labels. Default is True.
Private Sub Format(Optional ByVal valuesTitle As String = vbNullString, _
                  Optional ByVal catTitle As String = vbNullString, _
                  Optional ByVal plotTitle As String = vbNullString, _
                  Optional ByVal secondAxisTitle As String = vbNullString, _
                  Optional ByVal scope As Byte = GraphScopeNormal, _
                  Optional ByVal heightFactor As Long = 1, _
                  Optional ByVal hardCodeLabels As Boolean = True)
    
    Const CW As Long = 488
    Const RH As Long = 214

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim coef As Long
    Dim heightCoef As Long
    
    Set co = graphChart
    Set sh = Wksh()

    coef = 1
    If (scope = GraphScopeTimeSeries) Then coef = 1.75
    heightCoef = IIf(heightFactor <= 1, 1, (heightFactor + 1) * 0.08)

    'Reapply the width and height
    co.width = CW * coef
    co.height = RH * coef * heightCoef

    With sh.Shapes(co.Name)
        .Shadow.Type = msoShadow21
        .line.Visible = msoTrue
        .line.ForeColor.RGB = RGB(173, 216, 230)
        'Avoid changing graph width when adjusting row height
        .Placement = xlMove
    End With

    With co.Chart
        
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .HasMajorGridlines = True
            .MajorGridlines.Border.color = RGB(255, 255, 255)
            'If not hardcoded, titles may be range names. Test their existence
            'and use references instead of literal values
            If (Not hardCodeLabels) And RangeExist(catTitle) Then
                .AxisTitle.Caption = "= " & RngAddress(catTitle)
            Else
                .AxisTitle.Caption = catTitle
            End If

            .TickLabels.Font.Size = 10
            
            'Reverse plot order for spatial analysis
            If (scope = GraphScopeSpatial) Then .ReversePlotOrder = True
        End With
        
        'Formatting on primary axis
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .HasMajorGridlines = True
            .MajorGridlines.Border.color = RGB(255, 255, 255)
            .TickLabels.Font.Size = 10
            .TickLabels.Font.Name = "+mj-lt"

            If (Not hardCodeLabels) And RangeExist(valuesTitle) Then
                .AxisTitle.Caption = "= " & RngAddress(valuesTitle)
            Else
                .AxisTitle.Caption = valuesTitle
            End If
        End With

        'Add a title if needed
        .HasTitle = False
        If plotTitle <> vbNullString Then
            .HasTitle = True
            With .ChartTitle
                'You can add either address or value of plot title
                If (Not hardCodeLabels) And RangeExist(plotTitle) Then
                    .Caption = "= " & RngAddress(plotTitle)
                Else
                    .Caption = plotTitle
                End If
                .Font.Size = 14
                .Font.Bold = True
                .Font.Name = "+mj-lt"
            End With
        End If
        
        'Formatting on the secondary axis
        If .HasAxis(xlValue, xlSecondary) Then
            With .Axes(xlValue, xlSecondary)
                .HasTitle = True
                .HasMajorGridlines = False
                .MajorGridlines.Border.color = RGB(255, 255, 255)
                .TickLabels.Font.Size = 10
                .TickLabels.Font.Name = "+mj-lt"
                .AxisTitle.Text = "%"
                If secondAxisTitle <> vbNullString Then .AxisTitle.Text = secondAxisTitle
                .MaximumScale = 1
                .MajorUnit = 0.1
            End With
        End If
        
        With .PlotArea
            
            If scope = GraphScopeTimeSeries Then
                .Left = 25 * coef
            Else
                .Left = 16.618 * coef
            End If
            
            .Top = 16.618 * coef
            .width = 425 * coef
            .height = 178 * coef * heightCoef
        End With
        
        'Legend position
        If (scope = GraphScopeSpatial) Then
            .Legend.Position = xlLegendPositionBottom
        End If
    End With
    
End Sub


'@prop-title Test if a range exists in the worksheet
Private Function RangeExist(ByVal rngName As String) As Boolean
    Dim sh As Worksheet
    Dim rng As Range
    Set sh = Wksh()
    On Error Resume Next
        Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExist = Not (rng Is Nothing)
End Function

'@prop-title Return the full address of a range relative to the worksheet
Private Function RngAddress(ByVal rngName As String) As String
    Dim sh As Worksheet
    Dim rng As Range

    If Not RangeExist(rngName) Then Exit Function 'address can be empty
    Set sh = Wksh()
    Set rng = sh.Range(rngName)

    RngAddress = "'" & sh.Name & "'" & "!" & rng.Address
End Function

'@label:Checkings
'@section Checkings
'===============================================================================

'@sub-title Add entries to the checking object for future notes
'@param label String, the message to add
'@param scope Byte: the scope of the checking object
'@export
Private Sub LogInfo(ByVal label As String, _ 
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "CustomTable " & Name() & " object trace"

    'First, instantiate the internalChecks object
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)
    
    'Add new information to the checking object
    internalChecks.Add Cstr(checkCounter), label, scope
    checkCounter = checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:HasChecking
'@prop-title Check if the object has checks
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:CheckingValues
'@prop-title Get the checking values
'@export
Private Property Get CheckingValues() As Object

    'If there are no checks, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
End Property


'@sub-title Print debug information
'@details Outputs basic debug information to the Immediate window
Private Sub ShowDebug()
    Debug.Print "Check the informations of choice object " & Name()
End Sub


'@sub-title Raise a project-scoped error
'@param errNumb Long, the error number to raise
'@param errorMessage String, descriptive error message
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "DataSheet", "Error : " & Chr(34) & errorMessage
End Sub


'@label:interface
'@section Interface
'===============================================================================

'@jump:operate
'Operate with the outside world ------------------------------------------------

'@jump:add
'@sub-title Implements IGraphs.Add
Private Sub IGraphs_Add()
    Add
End Sub

'@jump:addseries
'@sub-title Implements IGraphs.AddSeries
Private Sub IGraphs_AddSeries(ByVal rngName As String, ByVal chrtType As String, Optional axisPos As String = vbNullString)
    AddSeries rngName, chrtType, axisPos
End Sub

'@jump:addlabels
'@sub-title Implements IGraphs.AddLabels
Private Sub IGraphs_AddLabels(ByVal catName As String, ByVal lblName As String, Optional ByVal prefix As String = vbNullString, _
                            Optional ByVal prefixOnly As Boolean = False, Optional ByVal hardCodeLabels As Boolean = True)
    AddLabels catName, lblName, prefix, prefixOnly, hardCodeLabels
End Sub

'@jump:format
'@sub-title Implements IGraphs.Format
Private Sub IGraphs_Format(Optional ByVal valuesTitle As String = vbNullString, _
                          Optional ByVal catTitle As String = vbNullString, _
                          Optional ByVal plotTitle As String = vbNullString, _
                          Optional ByVal scope As Byte = GraphScopeNormal, _
                          Optional ByVal heightFactor As Long = 1, _
                          Optional ByVal hardCodeLabels As Boolean = True)
    Format valuesTitle:=valuesTitle, catTitle:=catTitle, plotTitle:=plotTitle, scope:=scope, _
           heightFactor:=heightFactor, hardCodeLabels:=hardCodeLabels
End Sub
