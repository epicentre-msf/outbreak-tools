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
    index As Long
End Type

Private this As TGraph

Private Const GRAPHHEIGHT As Long = 14.3
Private Const GRAPHWIDTH As Long = 4.1

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
    seriesIndex = this.index
End Property

Private Property Let seriesIndex(ByVal ind As Long)
    this.index = ind
End Property

'Exposed methods
Public Sub Add()
    Dim co As ChartObject
    Dim sh As Worksheet
    Dim rng As Range
    Dim cw As Long
    Dim rh As Long

    Set sh = Wksh()
    Set rng = StartRange()
    cw = sh.Range("A1").width
    rh = sh.Range("A1").height

    Set co = sh.ChartObjects.Add(rng.Left, rng.Top, GRAPHWIDTH * cw, GRAPHHEIGHT * rh)
    co.Left = rng.Left
    co.Top = rng.Top

    With co.Chart
        .PlotArea.Interior.color = RGB(235, 235, 235)
        .Legend.Position = xlLegendPositionTop
    End With

    Set this.co = co
End Sub

Public Sub AddSeries(ByVal rngName As String, ByVal chrtType As String, Optional axisPos As String = vbNullString)

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

        customType =  xlBarClustered

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

Public Sub AddLabels(ByVal catName As String, ByVal lblName As String, Optional ByVal prefix As String = vbNullString, _ 
                     Optional Byval prefixOnly As Boolean = False)

    Dim co As ChartObject
    Dim sh As Worksheet
    Set co = graphChart
    Set sh = Wksh()
    Dim label As String
    Dim ind As Long

    ind = seriesIndex

    'Adding prefix to the label name of graphs (useful for time series graphs)
    If prefix <> vbNullString And (Not prefixOnly) Then prefix = prefix & " - "

    'Do not evaluate lblName on range if not necessary (avoid IIf())
    If prefixOnly Then 
        label = prefix
    Else
        label = prefix & sh.Range(lblName).Value
    End If

    'Throw error when co is Nothing
    With co.Chart
        .SeriesCollection(ind).XValues = sh.Range(catName)
        .SeriesCollection(ind).Name = label 'labels are hardcoded, the user can not modify by changing values in ranges
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
                  Optional ByVal asTimeSeries As Boolean = False)
    
    Const cw As Long = 119
    Const rh As Long = 15

    Dim co As ChartObject
    Dim sh As Worksheet
    Dim coef As Long
    Set co = graphChart
    Set sh = Wksh()

    
    coef = 1
    If asTimeSeries Then coef = 3

    'Coerce again the width and the height
    co.width = GRAPHWIDTH * cw * coef
    co.height = GRAPHHEIGHT * rh * coef
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
            .AxisTitle.text = catTitle
            .TickLabels.Font.Size = 8
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .HasMajorGridlines = True
            .MajorGridlines.Border.color = RGB(255, 255, 255)
            .TickLabels.Font.Size = 8
            .TickLabels.Font.Name = "+mj-lt"
            .AxisTitle.text = valuesTitle
        End With
        .HasTitle = False
        
        'formatting on the secondary axe
        If .HasAxis(xlValue, xlSecondary) Then
            With .Axes(xlValue, xlSecondary)
                .HasTitle = True
                .HasMajorGridlines = False
                .MajorGridlines.Border.color = RGB(255, 255, 255)
                .TickLabels.Font.Size = 8
                .TickLabels.Font.Name = "+mj-lt"
                .AxisTitle.text = "%"
                .MaximumScale = 1
                .MajorUnit = 0.1
            End With
        End If
        
        With .PlotArea
            .Left = 16.618 * coef
            .Top = 20 * coef
            .width = 400 * coef
            .height = 150 * coef
        End With
    End With
    
End Sub

Public Sub IGraphs_Add()
    Add
End Sub

Public Sub IGraphs_AddSeries(ByVal rngName As String, ByVal chrtType As String, Optional axisPos As String = vbNullString)
    AddSeries rngName, chrtType, axisPos
End Sub

Public Sub IGraphs_AddLabels(ByVal catName As String, ByVal lblName As String, Optional ByVal prefix As String = vbNullString, _ 
                            Optional ByVal prefixOnly As Boolean = False)
    AddLabels catName, lblName, prefix, prefixOnly
End Sub

Public Sub IGraphs_Format(Optional ByVal valuesTitle As String = vbNullString, _
                          Optional ByVal catTitle As String = vbNullString, _
                          Optional ByVal plotTitle As String = vbNullString, _
                          Optional ByVal asTimeSeries As Boolean = False)
    Format valuesTitle:=valuesTitle, catTitle:=catTitle, plotTitle:=plotTitle, asTimeSeries:=asTimeSeries
End Sub


