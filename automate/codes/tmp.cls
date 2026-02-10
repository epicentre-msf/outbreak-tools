Public Function Create(ByVal anash As Worksheet) As IAnalysis
  If anash Is Nothing Then ThrowError "Unable to create instance of class, Worksheet object not defined"
  CheckTablesRequirements
  With New Analysis
      Set .Wksh = anash
      Set Create = .Self
  End With
End Function


Public Property Get Wksh() As Worksheet
  Set Wksh = this.anash
End Property

Public Property Set Wksh(ByVal anash As Worksheet)
  Set this.anash = anash
End Property

Public Property Get Self() As IAnalysis
    Set Self = Me
End Property



'Add or remove Rows on worksheet analysis
Private Sub ManageRows(Optional ByVal del As Boolean = False)
  Dim sh As Worksheet
  Dim Lo As ListObject
  Dim tabl As ICustomTable
  Dim listLoNames As BetterArray

  Dim loname As String
  Dim prefName As String
  Dim idColName As String
  Dim counter As Long
  Dim selectedTable As String
  Dim totalCount As Long

  Set sh = Wksh()
  Set listLoNames = New BetterArray

  selectedTable = Replace(LCase(sh.Cells(1, 1).Value), _
                          "add or remove rows of", vbNullString)
  selectedTable = Application.WorksheetFunction.Trim(selectedTable)

  Select Case selectedTable
  Case "global summary"
    listLoNames.Push GLOBALSUMMARY
  Case "univariate analysis"
    listLoNames.Push UNIVARIATE
  Case "bivariate analysis"
    listLoNames.Push BIVARIATE
  Case "time series analysis"
    listLoNames.Push TIMESERIES
  Case "labels for time series graphs"
    listLoNames.Push TITLEGRAPHTIMESERIES
  Case "graph on time series"
    listLoNames.Push GRAPHTIMESERIES
  Case "spatial analysis"
    listLoNames.Push SPATIAL
  Case "spatio-temporal specifications"
    listLoNames.Push SPATIOTEMPSPECS
  Case "spatio-temporal analysis"
    listLoNames.Push SPATIOTEMPORAL
  Case Else
    listLoNames.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, _
                     TITLEGRAPHTIMESERIES, GRAPHTIMESERIES, SPATIAL, _
                     SPATIOTEMPSPECS, SPATIOTEMPORAL
  End Select

  For counter = listLoNames.LowerBound To listLoNames.UpperBound
    loname = listLoNames.Item(counter)

    Select Case loname

    Case TIMESERIES
      prefName = "Series"
      idColName = "series id"
      totalCount = 2
    Case TITLEGRAPHTIMESERIES
      prefName = "Graph"
      idColName = "graph id"
      totalCount = 1
    Case GRAPHTIMESERIES
      prefName = vbNullString
      idColName = vbNullString
      totalCount = 7
    Case SPATIOTEMPORAL
      prefName = vbNullString
      idColName = vbNullString
      totalCount = 3
    Case Else
      prefName = vbNullString
      idColName = vbNullString
      totalCount = 0
    End Select

    On Error Resume Next
      Set Lo = sh.ListObjects(loname)
      If Lo Is Nothing Then Exit Sub
    On Error GoTo 0
    Set tabl = CustomTable.Create(Lo, idColName, prefName)
    If del Then
        tabl.RemoveRows totalCount:=totalCount
    Else
        tabl.AddRows
    End If
  Next
End Sub


 'If a long line was like that And "Yeah, with some And in characters" AAAAAAAaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'