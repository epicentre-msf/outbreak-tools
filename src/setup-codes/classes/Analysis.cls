VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Management of events in analysis worksheet"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Management of events in analysis worksheet")


Implements IAnalysis
Option Explicit

Private Type TAnalysis
    anash As Worksheet
End Type

Private this As TAnalysis

Private Const CLASSNAME As String = "Analysis"
Private Const GLOBALSUMMARY As String = "Tab_global_summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"
Private Const SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const SPATIOTEMPSPECS As String = "Tab_SpatioTemporal_Specs"


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

'Errors
Private Sub CheckTablesRequirements()
  Dim loTable As BetterArray
  Dim counter As Long
  Dim Lo As ListObject
  Dim loname As String

  Set loTable = New BetterArray

  loTable.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, _
               GRAPHTIMESERIES, TITLEGRAPHTIMESERIES, SPATIAL, SPATIOTEMPSPECS, _ 
               SPATIOTEMPORAL

  For counter = loTable.LowerBound To loTable.UpperBound
      loname = loTable.Items(counter)
      On Error Resume Next
      Set Lo = Wksh.ListObjects(loname)
      If Lo Is Nothing Then _
       ThrowError " Listobject " & loname _
        & " does not exist in analysis worksheet "
      On Error GoTo 0
  Next
End Sub

Private Sub ThrowError(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 515, CLASSNAME, errorMessage
End Sub

'Exposed methods
Private Sub IAnalysis_AddRows()
  ManageRows
End Sub

Private Sub IAnalysis_RemoveRows()
  ManageRows del:=True
End Sub
