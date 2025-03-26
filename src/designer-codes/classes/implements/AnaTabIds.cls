VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "AnaTabIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Export the analyses"

'@IgnoreModule UseMeaningfulName, HungarianNotation
'@Folder("Linelist")
'@PredeclaredId
'@ModuleDescription("Export the analyses")

Option Explicit
Implements IAnaTabIds

Private Type TAnaIds
  sh As Worksheet
End Type

Private this As TAnaIds
Private Const CLASSNAME As String = "AnaTabIds"
'Ranges That should be present
Private Const RNGUASHEET As String = "RNG_SheetUAName"
Private Const RNGTSSHEET As String = "RNG_SheetTSName"
Private Const RNGSPSHEET As String = "RNG_SheetSPName"
Private Const RNGSPTEMPSHEET As String = "RNG_SheetSPTempName"

'Exposed methods

Public Function Create(ByVal idsh As Worksheet, Optional ByVal check As Boolean = True) As IAnaTabIds

  'Check requirements (Or Not)
  If check Then CheckRequirements idsh

  With New AnaTabIds
    Set .Wksh = idsh
    Set Create = .Self()
  End With
End Function

Public Property Get Self() As IAnaTabIds
  Set Self = Me
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  Set this.sh = sh
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Private Property Get ActualLo(ByVal scope As Byte, _ 
                            Optional ByVal prefix As String = "tab_ids_") As ListObject

  Dim sh As Worksheet
  Dim loName As String
  Dim actName As String

  Select Case scope
  Case AnalysisIdsScopeTimeSeries  'A time series analysis
    actName = "ts"
  Case AnalysisIdsScopeSpatial  'A spatial analysis
    actName = "sp"
  Case AnalysisIdsScopeNormal  'A standard univariate / bivariate analysis
    actName = "uba"
  Case AnalysisIdsScopeSpatioTemporal  'spatio-Temporal analysis
    actName = "sptemp"
  End Select

  loName = prefix & actName
  Set sh = Wksh()
  Set ActualLo = sh.ListObjects(loName)
End Property

'Get univariate, time series or spatial sheets or spatio temporal sheets
'analysisName can take up to 4 values: ua for univariate, ts for time series
'sp for spatial and sptemp for spatio-temporal.
Private Property Get ActualAnalysisSheet(ByVal scope As Byte) As Worksheet

  Dim sheetName As String
  Dim currwb As Workbook
  Dim sh As Worksheet

  Set sh = Wksh()
  Set currwb = sh.Parent

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
  
  Set ActualAnalysisSheet = currwb.Worksheets(sheetName)
End Property

'Add a table names to export to a Worksheet
Private Sub AddTableInfos( _ 
  ByVal scope As Byte, _
  ByVal tabId As String, _  
  ByVal tabRangesNames As BetterArray, _ 
  Optional ByVal exportTable As String = "yes")

  Dim Lo As ListObject
  Dim LoRng As Range
  Dim cellRng As Range
  Dim counter As Long
  Dim sh As Worksheet
  
  If tabRangesNames.Length = 0 Then Exit Sub

  Set Lo = ActualLo(scope)
  Set LoRng = Lo.Range
  Set sh = Wksh()

  If (LoRng.Cells(2, 1).Value = vbNullString) Then
   Set cellRng = LoRng.Cells(1, 1)
  Else
    Set cellRng = LoRng.Cells(LoRng.Rows.Count, 1)
  End If

  'Add the betterArray elements
  tabRangesNames.ToExcelRange cellRng.Offset(1, 1)
  
  For counter = 1 To tabRangesNames.Length
    cellRng.Offset(counter).Value = tabId
    'Add an information on wether to expor the table or not 
    '(by default all the tables are exported)
    cellRng.Offset(counter, 2).Value = exportTable
  Next

  'Resize the listObject of table
  Set LoRng = sh.Range(LoRng.Cells(1, 1), cellRng.Offset(counter - 1, 2))
  Lo.Resize LoRng
End Sub

'Add graph informations
Private Sub AddGraphInfo( _
  ByVal scope As Byte, _ 
  ByVal tabId As String, _
  ByVal graphId As String, _
  ByVal SeriesName As String, _ 
  ByVal SeriesType As String, _ 
  ByVal SeriesPos As String, _ 
  ByVal SeriesLabel As String, _ 
  ByVal SeriesColumnLabel As String, _ 
  ByVal hardCodeLabels As Boolean, _
  ByVal outRangeAddress As String, _ 
  Optional ByVal prefix As String = vbNullString, _ 
  Optional ByVal prefixOnly As Boolean = False)


  Dim Lo As ListObject
  Dim LoRng As Range
  Dim cellRng As Range
  Dim sh As Worksheet

  Set Lo = ActualLo(scope, prefix:="graph_ids_")
  Set LoRng = Lo.Range
  Set sh = Wksh()

  If LoRng.Cells(2, 1).Value = vbNullString Then
    Set cellRng = LoRng.Cells(2, 1)
  Else
    Set cellRng = LoRng.Cells(LoRng.Rows.Count + 1, 1)
  End If

  cellRng.Value = tabId
  cellRng.Cells(1, 2).Value = graphId
  cellRng.Cells(1, 3).Value = SeriesName
  cellRng.Cells(1, 4).Value = SeriesType
  cellRng.Cells(1, 5).Value = SeriesPos
  cellRng.Cells(1, 6).Value = SeriesLabel
  cellRng.Cells(1, 7).Value = SeriesColumnLabel
  cellRng.Cells(1, 8).Value = CLng(hardCodeLabels)
  cellRng.Cells(1, 9).Value = outRangeAddress
  cellRng.Cells(1, 10).Value = prefix
  cellRng.Cells(1, 11).Value = CLng(prefixOnly)

  'Resize the table
  Set LoRng = sh.Range(LoRng.Cells(1, 1), cellRng.Cells(1, 11))
  Lo.Resize LoRng
End Sub


'Add graph label informations
Private Sub AddGraphFormat( _
  ByVal scope As Byte, _ 
  ByVal tabId As String, _ 
  ByVal graphId As String, _
  ByVal catTitle As String, _ 
  ByVal valuesTitle As String, _ 
  ByVal hardCodeLabels As Boolean, _ 
  Optional ByVal heightFactor As Long = 1, _ 
  Optional ByVal plotTitle As String = vbNullString)

  Dim Lo As ListObject
  Dim LoRng As Range
  Dim cellRng As Range
  Dim sh As Worksheet

  Set Lo = ActualLo(scope, prefix:="graph_formats_")
  Set LoRng = Lo.Range
  Set sh = Wksh()

  If LoRng.Cells(2, 1).Value = vbNullString Then
    Set cellRng = LoRng.Cells(2, 1)
  Else
    Set cellRng = LoRng.Cells(LoRng.Rows.Count + 1, 1)
  End If

  cellRng.Value = tabId
  cellRng.Cells(1, 2).Value = graphId
  cellRng.Cells(1, 3).Value = catTitle
  cellRng.Cells(1, 4).Value = valuesTitle
  cellRng.Cells(1, 5).Value = CLng(hardCodeLabels)
  cellRng.Cells(1, 6).Value = heightFactor
  cellRng.Cells(1, 7).Value = plotTitle

  'Resize the table
  Set LoRng = sh.Range(LoRng.Cells(1, 1), cellRng.Cells(1, 7))
  Lo.Resize LoRng
End Sub

Private Sub TransferRange(ByVal inpsh As Worksheet, ByVal outsh As Worksheet, ByVal rngName As String)

  Dim inpRng As Range
  Dim outRng As Range
 
  If Not RangeExists(inpsh, rngName) Then Exit Sub

  Set inpRng = inpsh.Range(rngName)
  Set outRng = outsh.Range(inpRng.Address)

  'Name of the output range
  outRng.Name = rngName
End Sub


Private Sub WriteTableNames(ByVal outsh As Worksheet, ByVal scope As Byte)

  Dim inpsh As Worksheet 'Input worksheet
  Dim tabLo As ListObject
  Dim cellRng As Range
  Dim rngName As String

  Set inpsh = ActualAnalysisSheet(scope)

  'Export the tables ranges for the current scope
  Set tabLo = ActualLo(scope)
  Set cellRng = tabLo.Range.Cells(2, 1)

  Do While (Not IsEmpty(cellRng))
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

  'Add series and labels
  gr.AddSeries rngName:=seriesName, chrtType:=seriesType, axisPos:=seriesPos
  gr.AddLabels catName:=seriesLabel, lblName:=seriesColumnLabel, _ 
                  hardCodeLabels:=hardCodeLabels, prefix:=prefix, _ 
                  prefixOnly:=prefixOnly
End Sub


Private Sub AddFormat(ByVal gr As IGraphs, ByVal grTab As ICustomTable, ByVal scope As Byte, _ 
                      ByVal graphId As String)

  Dim valuesTitle As String
  Dim catTitle As String
  Dim plotTitle As String
  Dim heightFactor As Long
  Dim grscope As Byte
  Dim hardCodeLabels As Boolean

  heightFactor = 1
  On Error Resume Next
  heightFactor = CLng(grTab.Value(colName:="heightFactor", keyName:=graphId))
  Debug.Print grTab.Value(colName:="heightFactor", keyName:=graphId)
  On Error GoTo 0
   
  grscope = Switch( _
     (scope = AnalysisIdsScopeTimeSeries Or scope = AnalysisIdsScopeSpatioTemporal), GraphScopeTimeSeries, _ 
     scope = AnalysisIdsScopeSpatial, GraphScopeSpatial, _ 
     True, GraphScopeNormal)
  
  On Error Resume Next
   hardCodeLabels = CBool(grTab.Value(colName:="hardCodeLabels", keyName:=graphId))
  On Error GoTo 0
  
  valuesTitle = grTab.Value(colName:="valuesTitle", keyName:=graphId)
  catTitle = grTab.Value(colName:="catTitle", keyName:=graphId)
  plotTitle = grTab.Value(colName:="plotTitle", keyName:=graphId)
  
  'Format the graph
  gr.Format valuesTitle:=valuesTitle, _ 
              catTitle:=catTitle, _ 
              plotTitle:=plotTitle, _ 
              scope:=grscope, _ 
              heightFactor:=heightFactor, _ 
              hardCodeLabels:=hardCodeLabels
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

  'First, write table names
  WriteTableNames outsh, scope

  Set grLo = ActualLo(scope, prefix:="graph_ids_")
  Set grFormLo = ActualLo(scope, prefix:="graph_formats_")
  Set grFormatTab = CustomTable.Create(grFormLo, idCol:="graphId")

  'Build the graphs one by one. The second column is the column of graph Id
  Set cellRng = grLo.Range.Cells(2, 2)

  Do While (Not IsEmpty(cellRng))

    firstGraph = (cellRng.Offset(-1).Value = "graphId")
    newId = (cellRng.Offset(-1).Value <> cellRng.Value) And (Not firstGraph)

    'On new graphId, check if there have been a previous graph. In that case,
    'add formatting for previous graphs (the graphId is unchanged)
    If newId And  (Not gr Is Nothing)  Then AddFormat gr, grFormatTab, scope, graphId

    graphId = cellRng.Value
    'the output Range is after offset 7
    outRangeAddress = cellRng.Offset(, 7).Value
    
    If newId Or firstGraph Then
      'Create a new graph on new Id
      Set outCellRng = outsh.Range(outRangeAddress)
      Set gr = Graphs.Create(outsh, outCellRng)
      Application.GoTo outCellRng
      gr.Add
    End If

    AddSeries gr, cellRng
    
    If IsEmpty(cellRng.Offset(1)) Then AddFormat gr, grFormatTab, scope, graphId

    'Loop until the graph Id range is empty
    Set cellRng = cellRng.Offset(1)
  Loop
End Sub

'Check Requirements and Helpers functions

'Test if a listobject exists
Private Function LoExists(ByVal sh As Worksheet, ByVal loName As String) As Boolean
  Dim Lo As ListObject

  On Error Resume Next
    Set Lo = sh.ListObjects(loName)
  On Error GoTo 0
  LoExists = Not (Lo Is Nothing)
End Function

'Test if a range exits in a worksheet
Private Function RangeExists(ByVal sh As Worksheet, ByVal rngName As String) As Boolean
  Dim rng As Range
  On Error Resume Next
    Set rng = sh.Range(rngName)
  On Error GoTo 0
  RangeExists = Not (rng Is Nothing)
End Function

Private Function sheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
  Dim anash As Worksheet

  On Error Resume Next
    Set anash = wb.Worksheets(sheetName)
  On Error GoTo 0
  sheetExists = Not (anash Is Nothing)
End Function

'Check Requirements of the class
Private Sub CheckRequirements(ByVal sh As Worksheet)
  Dim loArray As BetterArray
  Dim counter As Long
  Dim wb As Workbook
  Dim actName As String

  
  If sh Is Nothing Then ThrowError "Unable to instanciate, Please provide a worksheet"
  Set loArray = New BetterArray
  loArray.Push "tab_ids_uba", "tab_ids_sp", "tab_ids_ts", "tab_ids_sptemp", _ 
              "graph_ids_uba", "graph_ids_sp", "graph_ids_ts", "graph_ids_sptemp", _
              "graph_formats_uba", "graph_formats_sp", "graph_formats_ts", "graph_formats_sptemp" 

  'Test if LisObjects are presents
  For counter = loArray.LowerBound To loArray.UpperBound
    actName = loArray.Item(counter)
    If Not LoExists(sh, actName) Then _ 
      ThrowError "Unable to instanciate AnaTabIds class. Table " & actName & " is missing."
  Next

  'Test if Ranges for sheetnames (ua, ts, sp, and sptemp) are presents
  loArray.Clear
  loArray.Push RNGUASHEET, RNGTSSHEET, RNGSPSHEET, RNGSPTEMPSHEET
  Set wb = sh.Parent

  For counter = loArray.LowerBound To loArray.UpperBound
    'actual Range name
    actName = loArray.Item(counter)
    If Not RangeExists(sh, actName) Then ThrowError "Unable to instanciate AnaTabIds class. Range " & actName & " is missing."
    If Not sheetExists(wb, sh.Range(actName).Value) Then ThrowError "Unable to instanciate AnaTabIds class. sheetName on " & actName & " is missing."
  Next
End Sub

'Throw Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub


Private Sub IAnaTabIds_AddGraphInfo( _
  ByVal scope As Byte, _ 
  ByVal tabId As String, _
  ByVal graphId As String, _
  ByVal SeriesName As String, _ 
  ByVal SeriesType As String, _ 
  ByVal SeriesPos As String, _ 
  ByVal SeriesLabel As String, _ 
  ByVal SeriesColumnLabel As String, _ 
  ByVal outRangeAddress As String, _ 
  Optional ByVal hardCodeLabels As Boolean = False, _
  Optional ByVal prefix As String = vbNullString, _ 
  Optional ByVal prefixOnly As Boolean = False)

  AddGraphInfo scope, tabId, graphId, SeriesName, SeriesType, _ 
               SeriesPos, SeriesLabel, SeriesColumnLabel, _ 
               hardCodeLabels, outRangeAddress, prefix, prefixOnly
End Sub

Private Sub IAnaTabIds_AddGraphFormat( _
  ByVal scope As Byte, _ 
  ByVal tabId As String, _ 
  ByVal graphId As String, _
  ByVal catTitle As String, _ 
  ByVal valuesTitle As String, _ 
  Optional ByVal hardCodeLabels As Boolean = False, _ 
  Optional ByVal heightFactor As Long = 1, _ 
  Optional ByVal plotTitle As String = vbNullString)

  AddGraphFormat scope, tabId, graphId, catTitle, valuesTitle, hardCodeLabels, _
                 heightFactor, plotTitle
End Sub

Private Sub IAnaTabIds_AddTableInfos( _ 
  ByVal scope As Byte, _
  ByVal tabId As String, _  
  ByVal tabRangesNames As BetterArray, _ 
  Optional ByVal exportTable As String = "yes")

  AddTableInfos scope:=scope, tabId:=tabId, tabRangesNames:=tabRangesNames, _ 
                exportTable:=exportTable

End Sub

Private Sub IAnaTabIds_WriteGraphs(ByVal outsh As Worksheet, ByVal scope As Byte)
  WriteGraphs outsh, scope
End Sub