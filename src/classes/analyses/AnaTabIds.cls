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

'@section Factory
'===============================================================================

' @description Create an AnaTabIds instance
' @param idsh Worksheet containing the tracking tables
' @param check Whether to validate that all required ListObjects exist
Public Function Create(ByVal idsh As Worksheet, _
                       Optional ByVal check As Boolean = True) As IAnaTabIds

    If idsh Is Nothing Then ThrowError InvalidArgument, "Worksheet not defined"
    If check Then CheckRequirements idsh

    With New AnaTabIds
        Set .Wksh = idsh
        Set Create = .Self
    End With
End Function

'@section Internal properties (factory use)
'===============================================================================

Public Property Get Self() As IAnaTabIds
    Set Self = Me
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@section ListObject access
'===============================================================================

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

'@section Table tracking
'===============================================================================

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

'@section Graph tracking
'===============================================================================

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

'@section Range transfer and graph writing
'===============================================================================

Private Function RangeExists(ByVal sh As Worksheet, ByVal rngName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExists = (Not rng Is Nothing)
End Function

Private Sub TransferRange(ByVal inpsh As Worksheet, ByVal outsh As Worksheet, _
                           ByVal rngName As String)
    Dim inpRng As Range
    Dim outRng As Range

    If Not RangeExists(inpsh, rngName) Then Exit Sub

    Set inpRng = inpsh.Range(rngName)
    Set outRng = outsh.Range(inpRng.Address)
    outRng.Name = rngName
End Sub

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

Private Function LoExists(ByVal sh As Worksheet, ByVal loName As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = sh.ListObjects(loName)
    On Error GoTo 0
    LoExists = Not (lo Is Nothing)
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (sh Is Nothing)
End Function

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

'@section Error handling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface implementation
'===============================================================================

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
