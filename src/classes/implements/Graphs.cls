VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Graphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Analysis")
'@ModuleDescription("Build Graphs for the analysis")
'@IgnoreModule

Option Explicit
Implements IGraphs

Private Type TGraph
    sh As Worksheet
    strtRng As Range
    co As ChartObject
    Index As Long
End Type

Private this As TGraph


Public Function Create(sh As Worksheet, posRng As Range) As IGraphs
    With New Graphs
        Set .Wksh = sh
        Set .StartRange = posRng
        Set Create = .Self
    End With
End Function

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Get Self() As IGraphs
    Set Self = Me
End Property

Public Property Get StartRange() As Range
    Set StartRange = this.strtRng
End Property

Public Property Set StartRange(ByVal strtRng As Range)
    Set this.strtRng = strtRng
End Property

Public Property Get graphChart() As ChartObject
    Set graphChart = this.co
End Property

Private Property Get seriesIndex() As Long
    seriesIndex = this.Index
End Property

Private Property Let seriesIndex(ByVal ind As Long)
    this.Index = ind
End Property

'Exposed methods
Public Sub Add()
    
    'upon creation width and height of graphs are constants
    
    Const cw As Long = 488 'graph width
    Const rh As Long = 214  'graph height
    
    Dim co As ChartObject
    Dim sh As Worksheet
    Dim rng As Range

    
    
    Set sh = Wksh()
    Set rng = StartRange()


    Set co = sh.ChartObjects.Add(rng.Left, rng.Top, rh, cw)
    co.Left = rng.Left
    co.Top = rng.Top

    With co.Chart
        .PlotArea.Interior.color = RGB(235, 235, 235)
        .Legend.Position = xlLegendPositionTop
    End With

    Set this.co = co
End Sub

Public Sub AddSeries(ByVal rngName As String, ByVal chrtType As String, _ 
                    Optional axisPos As String = vbNullString)

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim ind As Long
    Dim customType As XlChartType
    Dim axGrp As XlAxisGroup

    ind = seriesIndex + 1

    Set co = graphChart
    Set sh = Wksh()

    If co Is Nothing Then
        Add
        Set co = graphChart
    End If

    Select Case chrtType
    Case "line"
        customType = xlLineMarkers

    Case "point"
        customType = xlXYScatter

    Case "bar"
        customType = xlColumnClustered

    Case "hbar"

        customType = xlBarClustered

    Case Else
        customType = xlColumnClustered
    End Select

    axGrp = IIf(axisPos = "right", xlSecondary, xlPrimary)

    With co.Chart
        .SeriesCollection.Add _
        Source:=sh.Range(rngName), RowCol:=xlColumns, _
        CategoryLabels:=False, Replace:=False
        .SeriesCollection(ind).Values = sh.Range(rngName)

        'Add axis at the right if required
        If axGrp = xlSecondary Then
            .HasAxis(xlValue, xlSecondary) = True
            .SeriesCollection(ind).AxisGroup = axGrp
        End If


        .SeriesCollection(ind).chartType = customType
    End With


    seriesIndex = ind
End Sub

'Test if a range exists in the worksheet
Private Function RangeExist(ByVal rngName As String) As Boolean
    Dim sh As Worksheet
    Dim rng As Range
    Set sh = Wksh()
    On Error Resume Next
        Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExist = Not (rng Is Nothing)
End Function

'Get the True adress of a rangename to use in captions
Private Function RngAddress(ByVal rngName As String) As String
    Dim sh As Worksheet
    Dim rng As Range

    If Not RangeExist(rngName) Then Exit Function 'adress can be empty
    Set sh = Wksh()
    Set rng = sh.Range(rngName)

    RngAddress = "'" & sh.Name & "'" & "!" & rng.Address
End Function

Public Sub AddLabels(ByVal catName As String, _
                     ByVal lblName As String, _
                     Optional ByVal prefix As String = vbNullString, _
                     Optional ByVal prefixOnly As Boolean = False, _
                     Optional ByVal hardCodeLabels As Boolean = True)

    Dim co As ChartObject
    Dim sh As Worksheet
    Set co = graphChart
    Set sh = Wksh()
    Dim labelValue As String
    Dim ind As Long

    ind = seriesIndex

    'Adding prefix to the label name of graphs (useful for time series graphs)
    If prefix <> vbNullString And (Not prefixOnly) Then prefix = prefix & " - "

    If prefixOnly Then
        labelValue = prefix
    ElseIf RangeExist(lblName) Then
        labelValue = prefix & sh.Range(lblName).Value
    End If

    'Throw error when co is Nothing
    With co.Chart
        .SeriesCollection(ind).XValues = sh.Range(catName)

        'Labels can be hardcoded or Not, by default, they are hardcoded
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

Public Sub Format(Optional ByVal valuesTitle As String = vbNullString, _
                  Optional ByVal catTitle As String = vbNullString, _
                  Optional ByVal plotTitle As String = vbNullString, _
                  Optional ByVal secondAxisTitle As String = vbNullString, _
                  Optional ByVal scope As GraphScope = GraphScopeNormal, _
                  Optional ByVal heightFactor As Long = 1, _
                  Optional ByVal hardCodeLabels As Boolean = True)
    
    Const cw As Long = 488
    Const rh As Long = 214

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim coef As Long
    Dim heightCoef As Long
    
    Set co = graphChart
    Set sh = Wksh()

    coef = 1
    If (scope = GraphScopeTimeSeries) Then coef = 3
    heightCoef = IIf(heightFactor <= 1, 1, (heightFactor + 1) * 0.08)

    'Coerce again the width and the height
    co.width = cw * coef
    co.height = rh * coef * heightCoef

    With sh.Shapes(co.Name)
        .Shadow.Type = msoShadow21
        .line.Visible = msoTrue
        .line.ForeColor.RGB = RGB(173, 216, 230)
        'Avoid changing graph width
        .Placement = xlMove
    End With

    With co.Chart
        
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .HasMajorGridlines = True
            .MajorGridlines.Border.color = RGB(255, 255, 255)
            'If hardcoded labels, the title are range names , test if they exist
            'and use reference instead of hard values
            If (Not hardCodeLabels) And RangeExist(catTitle) Then
                .AxisTitle.Caption = "= " & RngAddress(catTitle)
            Else
                .AxisTitle.Caption = catTitle
            End If

            .TickLabels.Font.Size = 8
            
            'Reverse plot order for spatial analysis
            If (scope = GraphScopeSpatial) Then .ReversePlotOrder = True
        End With
        
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .HasMajorGridlines = True
            .MajorGridlines.Border.color = RGB(255, 255, 255)
            .TickLabels.Font.Size = 8
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
        
        'formatting on the secondary axe
        If .HasAxis(xlValue, xlSecondary) Then
            With .Axes(xlValue, xlSecondary)
                .HasTitle = True
                .HasMajorGridlines = False
                .MajorGridlines.Border.color = RGB(255, 255, 255)
                .TickLabels.Font.Size = 8
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
        
        'Legend position of the chart
        If (scope = GraphScopeSpatial) Then
            .Legend.Position = xlLegendPositionBottom
        End If
    End With
    
End Sub

Public Sub IGraphs_Add()
    Add
End Sub

Public Sub IGraphs_AddSeries(ByVal rngName As String, ByVal chrtType As String, Optional axisPos As String = vbNullString)
    AddSeries rngName, chrtType, axisPos
End Sub

Public Sub IGraphs_AddLabels(ByVal catName As String, ByVal lblName As String, Optional ByVal prefix As String = vbNullString, _
                            Optional ByVal prefixOnly As Boolean = False, Optional ByVal hardCodeLabels As Boolean = True)
    AddLabels catName, lblName, prefix, prefixOnly, hardCodeLabels
End Sub

Public Sub IGraphs_Format(Optional ByVal valuesTitle As String = vbNullString, _
                          Optional ByVal catTitle As String = vbNullString, _
                          Optional ByVal plotTitle As String = vbNullString, _
                          Optional ByVal scope As GraphScope = GraphScopeNormal, _
                          Optional ByVal heightFactor As Long = 1, _
                          Optional ByVal hardCodeLabels As Boolean = True)
    Format valuesTitle:=valuesTitle, catTitle:=catTitle, plotTitle:=plotTitle, scope:=scope, _
           heightFactor:=heightFactor, hardCodeLabels:=hardCodeLabels
End Sub
